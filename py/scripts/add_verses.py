"""綱要附經文：在每一點綱要下面插入它引用的恢復本經文。

用法：python3 scripts/add_verses.py <綱要.docx> [-o 輸出資料夾]
產出（預設放在 output/<原檔名>/）：<原檔名>（附經文）.docx、同名 .pdf、<原檔名>_校對報告.md；原檔不動。
"""
import argparse
import re
import sys
from collections import defaultdict
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).parent))
from bible import Bible, quoted_in, quoted_segments  # noqa: E402
from export_pdf import export_pdf  # noqa: E402
from refparse import group_refs, parse, range_label, unparsed_fragments  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
START_MARKERS = ("讀經", "读经")
STOP_MARKERS = ("職事信息摘錄", "职事信息摘录", "參考與參讀資料", "参考与参读资料", "信息摘錄")

STYLE_ID = "VerseText"
STYLE_NAME = "經文"
FONT_EA = "標楷體"            # DFKai-SB；用中文名稱，Word 與 LibreOffice 都認得
FONT_LATIN = "Times New Roman"
SIZE_HALF_PT = 21             # 10.5pt（五號），比綱要 12pt（小四）小一號
LINE_TWIPS = "280"            # 固定行高 14pt（標楷體預設行高偏大，經文要排得緊密）
EMU_PER_TWIP = 635
OMIT_MIN = 6                  # 一個出處（例：二一1～13）達 6 節以上，就不列經文，寫「從畧。」
FULL_QUOTE = 0.6              # 綱要用引號引了某節的六成以上 → 算全文引用，下面不再列那節
                              # （對照簡體範本 2640 段：引了六成以上的，人工版都沒列，沒有例外）
VERSE_GAP_BEFORE = "60"       # 經文區塊上方留 3pt、下方留 6pt，跟綱要分開
VERSE_GAP_AFTER = "120"

# 印刷本寫法：引號用『』「」；經節之間的「、」用「，」（例：二十20，27，32）
QUOTE_FORMS = {"“": "『", "”": "』", "‘": "「", "’": "」"}
REF_COMMA_RE = re.compile(r"(?:(?<=[0-9])|(?<=[0-9][上中下]))、(?=[一二三四五六七八九十]*[0-9])")
ASCII_LETTER_RE = re.compile(r"[A-Za-z]")

# 各層綱要與摘錄標題的字元樣式：粗體＋顏色（在 Word 改樣式就能整份一起換色）
CHAR_STYLES = [("OutlineL1", "綱要第一層", "8A3324"),     # 壹：赭紅
               ("OutlineL2", "綱要第二層", "2E6B4F"),     # 一：墨綠
               ("OutlineL3", "綱要第三層", "2B5C8A"),     # 1：靛藍
               ("OutlineL4", "綱要第四層", "6B4C9A"),     # a：藕紫
               ("ExcerptTitle", "摘錄標題", "8A3324"),
               ("ExcerptSubtitle", "摘錄小標題", "2E6B4F")]
LEVEL_RES = [re.compile(r"^[壹貳叁參肆伍陸柒捌玖拾]+[\t\u3000 ]"), re.compile(r"^[一二三四五六七八九十]+[\t\u3000 ]"),
             re.compile(r"^[0-9]+[\t\u3000 ]"), re.compile(r"^[a-z][\t\u3000 ]")]
EXCERPT_MARKERS = ("職事信息摘錄", "职事信息摘录", "信息摘錄")
REFS_MARKERS = ("參考與參讀資料", "参考与参读资料")
EXCERPT_LINE_TWIPS = "400"   # 職事信息摘錄：固定行高 20pt（使用者 2026-10-20 的決定）
EXCERPT_TITLE_SZ = 30         # 摘錄裏置中的標題：15pt 以上是大標題，其餘是小標題
READING_FILL, READING_BORDER = "F2F2F2", "BFBFBF"   # 讀經：淺灰底、灰框
# w:pPr 裏排在 w:pBdr、w:shd 後面的元素（OOXML 規定的順序）
PPR_AFTER_BOX = ("tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct",
                 "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
                 "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
                 "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange")


# ---------- docx 輔助 ----------
def para_text(p, skip_superscript=True):
    """段落文字（含 Tab）；可略過上標（註腳號碼）。"""
    out = []
    for r in p._p.iter(qn("w:r")):
        rpr = r.find(qn("w:rPr"))
        if skip_superscript and rpr is not None:
            va = rpr.find(qn("w:vertAlign"))
            if va is not None and va.get(qn("w:val")) == "superscript":
                continue
        for child in r:
            if child.tag == qn("w:t"):
                out.append(child.text or "")
            elif child.tag == qn("w:tab"):
                out.append("\t")
    return "".join(out)


def _style_chain(style):
    while style is not None:
        yield style
        style = style.base_style


def _num_pr(p):
    """段落的自動編號 (numId, ilvl)：段落本身或它的樣式（含上層樣式）設的；沒有就回傳 None。"""
    num_id = ilvl = None
    for ppr in [p._p.pPr] + [s.element.pPr for s in _style_chain(p.style)]:
        npr = ppr.find(qn("w:numPr")) if ppr is not None else None
        if npr is None:
            continue
        nid, lvl = npr.find(qn("w:numId")), npr.find(qn("w:ilvl"))
        if num_id is None and nid is not None:
            num_id = nid.get(qn("w:val"))
        if ilvl is None and lvl is not None:
            ilvl = lvl.get(qn("w:val"))
    if num_id in (None, "0"):
        return None
    return num_id, ilvl or "0"


def _numbering_ind(p):
    """自動編號那一層設定的縮排（w:ind）；沒有編號或找不到就回傳 None。"""
    npr = _num_pr(p)
    if npr is None:
        return None
    try:
        numbering = p.part.numbering_part.element
    except (NotImplementedError, KeyError):
        return None
    num_id, ilvl = npr
    num = next((n for n in numbering.findall(qn("w:num")) if n.get(qn("w:numId")) == num_id), None)
    if num is None:
        return None
    for ov in num.findall(qn("w:lvlOverride")):           # 這個編號自己改過的層級
        if ov.get(qn("w:ilvl")) == ilvl:
            ind = ov.find(f"{qn('w:lvl')}/{qn('w:pPr')}/{qn('w:ind')}")
            if ind is not None:
                return ind
    abs_id = num.find(qn("w:abstractNumId")).get(qn("w:val"))
    for an in numbering.findall(qn("w:abstractNum")):
        if an.get(qn("w:abstractNumId")) == abs_id:
            for lvl in an.findall(qn("w:lvl")):
                if lvl.get(qn("w:ilvl")) == ilvl:
                    return lvl.find(f"{qn('w:pPr')}/{qn('w:ind')}")
    return None


def _char_twips(p):
    """一個字寬（twips）：段落樣式的字級，沒有就用 12pt。縮排的「字元」單位用它換算。"""
    for s in _style_chain(p.style):
        if s.font.size is not None:
            return int(s.font.size) // EMU_PER_TWIP
    return 240


def _ind_value(ind, names, char_twips):
    """從一個 w:ind 取值：先看「字元」單位（leftChars 等，百分之一字），再看 twips。"""
    if ind is None:
        return None
    for name in names:
        chars = ind.get(qn(f"w:{name}Chars"))
        if chars and int(chars) != 0:
            return int(chars) * char_twips // 100
    for name in names:
        v = ind.get(qn(f"w:{name}"))
        if v is not None:
            return int(v)
    return None


def text_start_twips(p):
    """綱要文字（編號後）第一個字的位置，給經文對齊用。
    縮排的優先順序跟 Word 一樣：段落本身 > 自動編號那一層 > 段落樣式（含上層樣式）。"""
    ppr = p._p.pPr
    inds = [ppr.find(qn("w:ind")) if ppr is not None else None, _numbering_ind(p)]
    inds += [s.element.pPr.find(qn("w:ind")) if s.element.pPr is not None else None for s in _style_chain(p.style)]
    cw = _char_twips(p)
    left = next((v for v in (_ind_value(i, ("left", "start"), cw) for i in inds) if v is not None), 0)
    first = 0
    for ind in inds:
        hang = _ind_value(ind, ("hanging",), cw)
        if hang is not None:
            first = -hang
            break
        fl = _ind_value(ind, ("firstLine",), cw)
        if fl is not None:
            first = fl
            break
    numbered = _num_pr(p) is not None
    txt = para_text(p, skip_superscript=False)
    if first < 0 and (numbered or "\t" in txt):
        return max(left, 0)      # 懸掛縮排：編號（自動編號或打的編號＋Tab）在前，文字從 left 開始
    return max(left + first, 0)  # 沒有編號（例：讀經）：對齊第一行起點


def ensure_style(doc):
    styles = doc.styles.element
    for s in styles.findall(qn("w:style")):
        if s.get(qn("w:styleId")) == STYLE_ID:
            return
    s = OxmlElement("w:style")
    s.set(qn("w:type"), "paragraph")
    s.set(qn("w:customStyle"), "1")
    s.set(qn("w:styleId"), STYLE_ID)
    default = next((x.get(qn("w:styleId")) for x in styles.findall(qn("w:style"))
                    if x.get(qn("w:type")) == "paragraph" and x.get(qn("w:default")) == "1"), None)
    for tag, val in (("w:name", STYLE_NAME), ("w:basedOn", default), ("w:qFormat", None)):
        if tag == "w:basedOn" and not val:
            continue
        e = OxmlElement(tag)
        if val:
            e.set(qn("w:val"), val)
        s.append(e)
    ppr = OxmlElement("w:pPr")
    for tag, attrs in (
        ("w:widowControl", {"w:val": "0"}),
        ("w:snapToGrid", {"w:val": "0"}),
        ("w:spacing", {"w:before": "0", "w:after": "0", "w:line": LINE_TWIPS, "w:lineRule": "exact"}),
        ("w:jc", {"w:val": "both"}),
    ):
        e = OxmlElement(tag)
        for k, v in attrs.items():
            e.set(qn(k), v)
        ppr.append(e)
    s.append(ppr)
    rpr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    for k, v in (("w:ascii", FONT_LATIN), ("w:hAnsi", FONT_LATIN), ("w:eastAsia", FONT_EA), ("w:cs", FONT_LATIN)):
        fonts.set(qn(k), v)
    rpr.append(fonts)
    for tag, val in (("w:b", "0"), ("w:i", "0"), ("w:color", "000000"),
                     ("w:sz", str(SIZE_HALF_PT)), ("w:szCs", str(SIZE_HALF_PT))):
        e = OxmlElement(tag)
        e.set(qn("w:val"), val)
        rpr.append(e)
    s.append(rpr)
    styles.append(s)


def verse_paragraph(label, text, left, hanging, before=None, after=None):
    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    ps = OxmlElement("w:pStyle")
    ps.set(qn("w:val"), STYLE_ID)
    ppr.append(ps)
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "left")
    tab.set(qn("w:pos"), str(left + hanging))
    tabs.append(tab)
    ppr.append(tabs)
    if before or after:
        sp = OxmlElement("w:spacing")
        if before:
            sp.set(qn("w:before"), before)
        if after:
            sp.set(qn("w:after"), after)
        ppr.append(sp)
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), str(left + hanging))
    ind.set(qn("w:hanging"), str(hanging))
    ppr.append(ind)
    p.append(ppr)

    def run(s, with_tab=False):
        r = OxmlElement("w:r")
        if with_tab:
            r.append(OxmlElement("w:tab"))
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = s
        r.append(t)
        return r

    p.append(run(label))
    p.append(run(text, with_tab=True))
    return p


def print_forms(text):
    """印刷本寫法：“”‘’ → 『』「」（英文字中間的撇號不動），經節之間的「、」→「，」。
    每個字一對一替換，長度不變。"""
    out = []
    for i, c in enumerate(text):
        if c in QUOTE_FORMS and not (c in "‘’" and 0 < i < len(text) - 1 and ASCII_LETTER_RE.match(text[i - 1])
                                     and ASCII_LETTER_RE.match(text[i + 1])):
            c = QUOTE_FORMS[c]
        out.append(c)
    return REF_COMMA_RE.sub("，", "".join(out))


def apply_print_forms(p):
    """整段（跨 run）套用 print_forms，再寫回原本的各個 w:t。"""
    ts = list(p._p.iter(qn("w:t")))
    old = [t.text or "" for t in ts]
    new = print_forms("".join(old))
    pos = 0
    for t, s in zip(ts, old):
        if new[pos:pos + len(s)] != s:
            t.text = new[pos:pos + len(s)]
        pos += len(s)


def ensure_char_styles(doc):
    styles = doc.styles.element
    have = {s.get(qn("w:styleId")) for s in styles.findall(qn("w:style"))}
    for sid, name, color in CHAR_STYLES:
        if sid in have:
            continue
        s = OxmlElement("w:style")
        s.set(qn("w:type"), "character")
        s.set(qn("w:customStyle"), "1")
        s.set(qn("w:styleId"), sid)
        n = OxmlElement("w:name")
        n.set(qn("w:val"), name)
        s.append(n)
        s.append(OxmlElement("w:qFormat"))
        rpr = OxmlElement("w:rPr")
        rpr.append(OxmlElement("w:b"))
        rpr.append(OxmlElement("w:bCs"))
        c = OxmlElement("w:color")
        c.set(qn("w:val"), color)
        rpr.append(c)
        s.append(rpr)
        styles.append(s)


def style_runs(p, style_id):
    """段落裏每個 run 套上字元樣式（已經有字元樣式的 run 不動，例：註腳號碼）。"""
    for r in p._p.iter(qn("w:r")):
        if len(r) == 0:
            continue
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            r.insert(0, rpr)
        if rpr.find(qn("w:rStyle")) is not None:
            continue
        rs = OxmlElement("w:rStyle")
        rs.set(qn("w:val"), style_id)
        rpr.insert(0, rs)


def reading_box(p):
    """讀經那一段加淺灰底和細框（已經有框或底色的不動）。"""
    ppr = p._p.get_or_add_pPr()
    if ppr.find(qn("w:pBdr")) is not None or ppr.find(qn("w:shd")) is not None:
        return
    bdr = OxmlElement("w:pBdr")
    for side in ("top", "left", "bottom", "right"):
        e = OxmlElement(f"w:{side}")
        for k, v in (("w:val", "single"), ("w:sz", "4"), ("w:space", "4"), ("w:color", READING_BORDER)):
            e.set(qn(k), v)
        bdr.append(e)
    shd = OxmlElement("w:shd")
    for k, v in (("w:val", "clear"), ("w:color", "auto"), ("w:fill", READING_FILL)):
        shd.set(qn(k), v)
    at = next((i for i, c in enumerate(ppr) if c.tag.split("}")[-1] in PPR_AFTER_BOX), len(ppr))
    ppr.insert(at, shd)
    ppr.insert(at, bdr)


def exact_line(p, twips):
    """段落改成固定行高（保留原本的段前、段後）。"""
    sp = p._p.get_or_add_pPr().get_or_add_spacing()
    sp.set(qn("w:line"), twips)
    sp.set(qn("w:lineRule"), "exact")


def max_size(p):
    return max((int(e.get(qn("w:val"))) for e in p._p.iter(qn("w:sz")) if (e.get(qn("w:val")) or "").isdigit()),
               default=0)


def is_centered(p):
    ppr = p._p.find(qn("w:pPr"))
    jc = ppr.find(qn("w:jc")) if ppr is not None else None
    return jc is not None and jc.get(qn("w:val")) == "center"


def format_document(doc):
    """印刷本寫法（『』、經節間的「，」）、各層綱要上色加粗、讀經加框、摘錄標題上色。"""
    ensure_char_styles(doc)
    paras = list(doc.paragraphs)
    region = "head" if any(para_text(p).strip().startswith(START_MARKERS) for p in paras) else "outline"
    for p in paras:
        if p.style is not None and p.style.style_id == STYLE_ID:
            continue
        text = para_text(p).strip()
        if text.startswith(START_MARKERS):
            region = "outline"
        elif text.startswith(EXCERPT_MARKERS):
            region = "excerpt"
        elif text.startswith(REFS_MARKERS):
            region = "refs"
        apply_print_forms(p)
        if region == "outline":
            if text.startswith(START_MARKERS):
                reading_box(p)
                continue
            level = next((i for i, r in enumerate(LEVEL_RES) if r.match(text)), None)
            if level is not None:
                style_runs(p, CHAR_STYLES[level][0])
        elif region == "excerpt":
            if text and not text.startswith(EXCERPT_MARKERS) and is_centered(p):
                style_runs(p, "ExcerptTitle" if max_size(p) >= EXCERPT_TITLE_SZ else "ExcerptSubtitle")
            exact_line(p, EXCERPT_LINE_TWIPS)


def label_width_twips(label):
    """估計出處標籤寬度：中文字 1 em、半形 0.5 em。"""
    em = SIZE_HALF_PT * 10  # 1 em（twips）
    return int(sum(em if ord(c) > 0x2000 else em * 0.5 for c in label))


# ---------- 主流程 ----------
def plan_verses(doc, bible, hints=None):
    """決定每一段要附哪些經文 → (plan:[(段落, [(標籤, 經文)])], report, stats)。

    hints：{段落文字: 上下文歷史}，給插入的段落用（它們的上下文來自別份文件）。"""
    report = defaultdict(list)
    plan = []          # (段落, [(label, text)])
    active = not any(para_text(p).strip().startswith(START_MARKERS) for p in doc.paragraphs)
    context = None
    stats = {"points": 0, "verses": 0}

    for p in list(doc.paragraphs):
        if p.style is not None and p.style.style_id == STYLE_ID:
            continue                                   # 已經附過的經文
        raw = para_text(p)
        text = raw.strip()
        if text.startswith(START_MARKERS):
            active, context = True, None
        elif text.startswith(STOP_MARKERS):
            active = False
        if not active or not text:
            continue

        if hints and text in hints:
            context = hints[text]
        refs, spans, context = parse(raw, bible.verse_count, context)
        where = text[:30] + ("…" if len(text) > 30 else "")
        for frag in unparsed_fragments(raw, spans):
            report["無法解析的片段（請人工確認）"].append(f"「{frag}」　← {where}")
        if not refs:
            continue

        lines, seen = [], set()
        for group in group_refs(refs):
            head = group[0]
            new = [r for r in group if r.label not in seen]
            if len(group) >= OMIT_MIN:
                # 一個出處達 6 節以上：不列經文，寫一行「創21:1-13　從畧。」
                label = range_label(group)
                if label in seen:
                    continue
                seen.add(label)
                new = group
                lines.append((label, "從畧。", None))
                report["六節以上的出處（寫「從畧。」）"].append(f"{label}（{head.src}）　← {where}")
            if not new:
                continue
            notes = list(dict.fromkeys(n for r in group for n in r.notes))
            if head.kind == "implicit":
                report["沒有書卷、由上下文推斷的出處"].append(
                    f"{range_label(new)}（原文「{head.src}」；{'；'.join(notes)}）　← {where}")
            elif notes:
                report["其他提醒"].append(f"{range_label(new)}：{'；'.join(notes)}　← {where}")
            if len(group) >= OMIT_MIN:
                continue
            for ref in new:
                seen.add(ref.label)
                add_line(ref, lines, report, where, raw, bible)
        lines = [(label, body) for label, body, _ in lines]
        if lines:
            plan.append((p, lines))
            stats["points"] += 1
            stats["verses"] += len(lines)
    return plan, report, stats


def add_line(ref, lines, report, where, raw, bible):
    """一節（或整章）的經文 → 加進 lines，並記下要校對的項目。"""
    if ref.kind == "prose" and ref.verse is None:
        # 句中只是提到某一章（例：按林前十四章），不算引經（使用者 2026-10-12 的決定）
        report["句中提到整章（不列）"].append(f"{ref.label}（{ref.src}）　← {where}")
        return
    body, note = bible.text(ref, raw)
    if body is None:
        report["查無經文"].append(f"{ref.label}（{ref.src}）　← {where}")
        return
    if ref.kind == "prose" and ref.verse is not None and quoted_in(body, bible.normalize(raw)):
        report["句中出處：綱要已全文引用，略過"].append(f"{ref.label}（{ref.src}）　← {where}")
        return
    if (ref.kind != "prose" and ref.verse is not None
            and quoted_in(body, quoted_segments(bible.normalize(raw)), FULL_QUOTE)):
        # 綱要已用引號全文引用這節（例：壹 “你將這些事…”—提前四6），下面不再列（使用者 2026-10-12 的決定）
        report["綱要已全文引用的經節（不列）"].append(f"{ref.label}（{ref.src}）　← {where}")
        return
    if ref.kind == "prose":
        report["句中出處：已附經文（請確認）"].append(f"{ref.label}（{ref.src}）　← {where}")
    if note and ref.verse is None:
        report["整章引用（寫「從畧。」）"].append(f"{ref.label}（{ref.src}）　← {where}")
    elif note:
        report["半節（請確認切分位置）"].append(f"{ref.label}　{body}　［{note}］　← {where}")
    # 連續整章（例：羅馬十二至十三章）合併成一行「羅12-13　從畧。」
    prev = lines[-1] if lines else None
    if (ref.verse is None and prev and prev[2] and prev[2][0] == ref.book
            and prev[2][2] == ref.chapter - 1 and prev[2][3] == ref.src):
        first = prev[2][1]
        lines[-1] = (f"{ref.label.rstrip('0123456789')}{first}-{ref.chapter}", body,
                     (ref.book, first, ref.chapter, ref.src))
    else:
        lines.append((ref.label, body,
                      (ref.book, ref.chapter, ref.chapter, ref.src) if ref.verse is None else None))


def default_out(src):
    """預設輸出資料夾：output/<原檔名>/，同一份綱要的 Word、PDF、報告放在一起。"""
    return ROOT / "output" / Path(src).stem


def insert_verses(doc, plan):
    """把 plan 的經文插在各段綱要下面，對齊該段文字第一個字。網頁版（site/py/web.py）也用這個。"""
    # 標籤欄寬度：取全文最長的標籤，讓整份文件對齊一致
    hanging = max([1050] + [label_width_twips(l) + 160 for _, ls in plan for l, _ in ls])
    for p, lines in plan:
        left = text_start_twips(p)
        anchor = p._p
        for k, (label, body) in enumerate(lines):
            new = verse_paragraph(label, body, left, hanging,
                                  VERSE_GAP_BEFORE if k == 0 else None,
                                  VERSE_GAP_AFTER if k == len(lines) - 1 else None)
            anchor.addnext(new)
            anchor = new


def process(src, out_dir, pdf=True, hints=None):
    bible = Bible()
    doc = Document(src)
    ensure_style(doc)
    plan, report, stats = plan_verses(doc, bible, hints)
    format_document(doc)
    insert_verses(doc, plan)

    out_dir.mkdir(parents=True, exist_ok=True)
    out_docx = out_dir / f"{src.stem}（附經文）.docx"
    doc.save(out_docx)
    out_pdf = export_pdf(out_docx) if pdf else None
    out_md = out_dir / f"{src.stem}_校對報告.md"
    out_md.write_text(render_report(src, out_docx, stats, report), encoding="utf-8")
    return out_docx, out_pdf, out_md, stats, report


ORDER = ["無法解析的片段（請人工確認）", "查無經文", "沒有書卷、由上下文推斷的出處",
         "半節（請確認切分位置）", "句中出處：已附經文（請確認）", "句中出處：綱要已全文引用，略過",
         "綱要已全文引用的經節（不列）", "句中提到整章（不列）", "整章引用（寫「從畧。」）", "六節以上的出處（寫「從畧。」）", "其他提醒"]


def render_report(src, out_docx, stats, report):
    lines = [f"# 校對報告：{src.name}", "",
             f"- 產出檔：`{out_docx.name}`",
             f"- 附上經文的綱要段落：{stats['points']} 段，共 {stats['verses']} 節", ""]
    if not any(report.values()):
        lines.append("沒有需要人工確認的項目。")
    for key in ORDER:
        items = report.get(key)
        if not items:
            continue
        lines += [f"## {key}（{len(items)}）", ""] + [f"- {x}" for x in items] + [""]
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser(description="綱要附經文")
    ap.add_argument("docx", type=Path)
    ap.add_argument("-o", "--out", type=Path, help="輸出資料夾（預設：output/<原檔名>/）")
    ap.add_argument("--no-pdf", action="store_true", help="不輸出 PDF")
    a = ap.parse_args()
    out_docx, out_pdf, out_md, stats, report = process(a.docx, a.out or default_out(a.docx), pdf=not a.no_pdf)
    print(f"完成：{stats['points']} 段綱要、{stats['verses']} 節經文")
    print(f"輸出：{out_docx}")
    if out_pdf:
        print(f"PDF ：{out_pdf}")
    print(f"報告：{out_md}")
    for key in ORDER:
        if report.get(key):
            print(f"  {key}：{len(report[key])}")


if __name__ == "__main__":
    main()
