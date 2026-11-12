"""中文綱要經文出處解析。

三種出處：
  formal   正式出處：羅五10、17，提後二15，一13~14，參徒十七11，猶20~21，詩五一
  implicit 只有節（或章節）而沒有書卷：—10~11、25節、—三四25上、（2~42）
  prose    句中出處：羅馬一章三至四節、羅馬八章十六節、希伯來一章五節下半、利未記十八至二十二章
"""
import re
from dataclasses import dataclass, field

from books import ABBR_TO_NO, BOOKS, NAME_TO_NO

CN_DIGIT = {"〇": 0, "○": 0, "零": 0, "一": 1, "二": 2, "三": 3, "四": 4,
            "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
CN = "[〇○零一二三四五六七八九十百]+"
RANGE_SEPS = "~～-－–"
LIST_SEPS = "，、,；;"
# 出處後面可以接的字（整章出處要靠它判斷，避免把「但一位」之類誤判成出處）
END_CHARS = set("，、,；;。：:）)」』”’ \t—．.")


def cn2int(s):
    """中文數字 → 整數。支援「十五」「二十」「一百一十九」與緊縮寫法「二一」「一一九」「一〇二」。"""
    if "十" in s or "百" in s:
        total, num = 0, 0
        for c in s:
            if c in CN_DIGIT:
                num = CN_DIGIT[c]
            elif c == "十":
                total += (num or 1) * 10
                num = 0
            elif c == "百":
                total += (num or 1) * 100
                num = 0
        return total + num
    n = 0
    for c in s:
        n = n * 10 + CN_DIGIT[c]
    return n


@dataclass
class Ref:
    book: int
    chapter: int
    verse: int = None        # None = 整章
    part: str = ""           # 上 / 中 / 下 / ""
    kind: str = "formal"     # formal / implicit / prose
    src: str = ""            # 原文片段
    pos: int = 0             # 在段落中的位置（排序用）
    notes: list = field(default_factory=list)
    item: int = -1           # 同一個出處項目（例：「二一1～13」）展開的各節，item 相同

    @property
    def label(self):
        b = BOOKS[self.book]["abbr"]
        if self.verse is None:
            return f"{b}{self.chapter}"
        return f"{b}{self.chapter}:{self.verse}{self.part}"


def range_label(refs):
    """一個出處項目的標籤：創21:1-13、林前1:13-2:5；只有一節就用該節的標籤。"""
    first, last = refs[0], refs[-1]
    if len(refs) == 1:
        return first.label
    b = BOOKS[first.book]["abbr"]
    end = f"{last.verse}{last.part}" if last.chapter == first.chapter else f"{last.chapter}:{last.verse}{last.part}"
    return f"{b}{first.chapter}:{first.verse}{first.part}-{end}"


def group_refs(refs):
    """把同一個出處項目展開的各節歸成一組（整章各自一組）。refs 需依原文順序。"""
    groups = []
    for r in refs:
        g = groups[-1] if groups else None
        if (g and r.verse is not None and g[0].verse is not None and r.item == g[0].item
                and r.src == g[0].src and r.kind == g[0].kind and r.pos == g[0].pos):
            g.append(r)
        else:
            groups.append([r])
    return groups


def _alt(words):
    return "|".join(re.escape(w) for w in sorted(words, key=len, reverse=True))


ABBR_RE = re.compile(_alt(ABBR_TO_NO))
VERSE_RE = re.compile(r"(\d+)(上|中|下)?")
CN_RE = re.compile(CN)
PROSE_RE = re.compile(
    rf"(?P<name>{_alt([n for n in NAME_TO_NO if len(n) >= 2])})第?(?P<ch1>{CN})(?:(?:至|到)(?P<ch2>{CN}))?(?:章|篇)"
    rf"(?:第?(?P<v1>{CN})(?P<jie1>節|节)?(?P<p1>上半|下半)?"
    rf"(?:(?:至|到)第?(?P<v2>{CN})(?:節|节)(?P<p2>上半|下半)?)?)?"
)
IMPLICIT_TRIGGER = re.compile(r"—|（|\(|參|参")
# 只有書名、沒有章節的提及（當作上下文用）；只取正式書名，避免「約翰」「雅各」等人名
MENTION_RE = re.compile("(?:" + _alt([n for n in NAME_TO_NO
                                      if len(n) >= 2 and re.search(r"(書|书|記|记|福音|篇|歌|言|錄|录|傳|传)$", n)])
                         + ")(?!生命讀經|生命读经|結晶讀經|结晶读经)")
JIE_RE = re.compile(r"[ \u00a0\u3000]?[節节]")   # 「14節」「14 節」
NOTE_RE = re.compile(r"(?:[與和及]?註\d+)+")       # 「六13與註1」「1註1」：註解號碼，跳過
NOTE_SEP_RE = re.compile(r"[與和及](?=\d+註)")      # 「1註1與8註1」的「與」當分隔


class _Scanner:
    def __init__(self, text, verse_count):
        self.t = text
        self.verse_count = verse_count  # (book, chapter) -> 節數

    def cn_verse(self, i):
        """位置 i 是否為「中文章＋阿拉伯節」，是則回傳 (章, 節的起點)。"""
        m = CN_RE.match(self.t, i)
        if m and VERSE_RE.match(self.t, m.end()):
            return cn2int(m.group()), m.end()
        return None

    def at_end(self, i):
        return i >= len(self.t) or self.t[i] in END_CHARS

    def verse_item(self, i, book, chapter, kind, src_start, pos, out):
        """解析 V、V~V、V~中文章V（跨章），後面可接「節」「節上」。回傳 (新位置, 目前章) 或 None。"""
        m = VERSE_RE.match(self.t, i)
        if not m:
            return None
        v1, p1 = int(m.group(1)), m.group(2) or ""
        j = m.end()
        end_ch, v2, p2 = chapter, None, ""
        if j < len(self.t) and self.t[j] in RANGE_SEPS:
            cv = self.cn_verse(j + 1)
            if cv:
                end_ch, k = cv
            else:
                k = j + 1
            m2 = VERSE_RE.match(self.t, k)
            if m2:
                v2, p2 = int(m2.group(1)), m2.group(2) or ""
                j = m2.end()
            else:
                end_ch = chapter
        mj = JIE_RE.match(self.t, j)
        if mj:
            j = mj.end()
            if j < len(self.t) and self.t[j] in "上中下":
                if v2 is None:
                    p1 = self.t[j]
                else:
                    p2 = self.t[j]
                j += 1
        src = self.t[src_start:j]
        mn = NOTE_RE.match(self.t, j)
        if mn:
            j = mn.end()
        if v2 is None:
            out.append(Ref(book, chapter, v1, p1, kind, src, pos, item=src_start))
        else:
            expand(book, chapter, v1, p1, end_ch, v2, p2, kind, src, pos, out, self.verse_count, src_start)
        return j, end_ch

    def verse_list(self, i, book, chapter, kind, pos, out, src_start=None):
        """從位置 i（章已知）解析節列表，並處理後面接的「，36」「、17」「，一13~14」「，二四」。
        src_start：第一項原文的起點（「十二1」的章從這裏開始）；預設為 i。"""
        r = self.verse_item(i, book, chapter, kind, i if src_start is None else src_start, pos, out)
        if not r:
            return i
        j, chapter = r
        while j < len(self.t) and (self.t[j] in LIST_SEPS or NOTE_SEP_RE.match(self.t, j)):
            k = j + 1
            if VERSE_RE.match(self.t, k):                     # 同章另一節
                r = self.verse_item(k, book, chapter, kind, k, pos, out)
            elif self.cn_verse(k):                            # 同卷另一章
                chapter, vstart = self.cn_verse(k)
                r = self.verse_item(vstart, book, chapter, kind, k, pos, out)
            else:
                r = self.whole_chapter(k, book, kind, pos, out)
            if not r:
                break
            j, chapter = r
        return j

    def whole_chapter(self, i, book, kind, pos, out):
        """整章：詩五一、創十七～十八。只在後面是標點或結尾時才算。"""
        m = CN_RE.match(self.t, i)
        if not m:
            return None
        ch1 = ch2 = cn2int(m.group())
        end = m.end()
        if end < len(self.t) and self.t[end] in RANGE_SEPS:      # 十七～十八
            m2 = CN_RE.match(self.t, end + 1)
            if not m2:
                return None
            ch2, end = cn2int(m2.group()), m2.end()
        if not self.at_end(end) or not 0 < ch1 <= ch2 <= BOOKS[book]["chapters"]:
            return None
        for ch in range(ch1, ch2 + 1):
            out.append(Ref(book, ch, None, "", kind, self.t[i:end], pos, item=i))
        return end, ch2


def expand(book, ch1, v1, p1, ch2, v2, p2, kind, src, pos, out, verse_count, item=-1):
    """把範圍展開成逐節。"""
    ch, v = ch1, v1
    first = True
    while True:
        last = (ch, v) == (ch2, v2)
        part = p1 if first else ""
        if last:
            part = (p1 or p2) if first else p2
        out.append(Ref(book, ch, v, part, kind, src, pos, item=item))
        if last:
            return
        first = False
        v += 1
        if v > verse_count(book, ch):
            ch, v = ch + 1, 1
            if ch > ch2:
                out[-1].notes.append("範圍超出章節，請檢查")
                return


def parse_formal(text, verse_count):
    sc = _Scanner(text, verse_count)
    refs, spans = [], []
    i = 0
    while i < len(text):
        m = ABBR_RE.match(text, i)
        if not m:
            i += 1
            continue
        book = ABBR_TO_NO[m.group()]
        j = m.end()
        end = None
        if BOOKS[book]["chapters"] == 1 and VERSE_RE.match(text, j):
            end = sc.verse_list(j, book, 1, "formal", i, refs)            # 猶20~21
        elif re.match(r"\d+[:：]\d", text[j:]):                          # 約3:16
            mc = re.match(r"(\d+)[:：]", text[j:])
            end = sc.verse_list(j + mc.end(), book, int(mc.group(1)), "formal", i, refs)
        elif sc.cn_verse(j):                                             # 羅五10
            ch, vstart = sc.cn_verse(j)
            if 0 < ch <= BOOKS[book]["chapters"]:
                end = sc.verse_list(vstart, book, ch, "formal", i, refs)
        else:                                                            # 詩五一
            r = sc.whole_chapter(j, book, "formal", i, refs)
            if r:
                k = r[0]
                # 整章後面也可能接「，二四1~10」「，二四」
                while k < len(text) and text[k] in LIST_SEPS:
                    if sc.cn_verse(k + 1):
                        ch, vstart = sc.cn_verse(k + 1)
                        k2 = sc.verse_list(vstart, book, ch, "formal", i, refs)
                    else:
                        r2 = sc.whole_chapter(k + 1, book, "formal", i, refs)
                        k2 = r2[0] if r2 else None
                    if not k2:
                        break
                    k = k2
                end = k
        if end and end > j:
            spans.append((i, end))
            i = end
        else:
            i += 1
    return refs, spans


def parse_prose(text, verse_count, taken):
    refs, spans = [], []
    for m in PROSE_RE.finditer(text):
        if any(a <= m.start() < b for a, b in taken):
            continue
        book = NAME_TO_NO[m.group("name")]
        ch1 = cn2int(m.group("ch1"))
        if not 0 < ch1 <= BOOKS[book]["chapters"]:
            continue
        end = m.end()
        src = m.group()
        pos = m.start()
        if m.group("ch2"):                                   # 利未記十八至二十二章
            for ch in range(ch1, cn2int(m.group("ch2")) + 1):
                refs.append(Ref(book, ch, None, "", "prose", src, pos, item=pos))
        elif m.group("v1") and (m.group("jie1") or m.group("v2")):
            v1 = cn2int(m.group("v1"))
            p1 = (m.group("p1") or "")[:1]
            if m.group("v2"):                                # 羅馬一章三至四節
                expand(book, ch1, v1, p1, ch1, cn2int(m.group("v2")),
                       (m.group("p2") or "")[:1], "prose", src, pos, refs, verse_count, pos)
            else:                                            # 羅馬八章十六節
                refs.append(Ref(book, ch1, v1, p1, "prose", src, pos, item=pos))
        else:                                                # 撒下七章
            end = m.end("ch1") + 1
            src = text[m.start():end]
            refs.append(Ref(book, ch1, None, "", "prose", src, pos, item=pos))
        spans.append((m.start(), end))
    return refs, spans


def parse_implicit(text, verse_count, taken, points, history):
    """沒有書卷的出處。書卷/章依序取自：同段前面的出處或書名（points，越近越優先），
    再來是前面段落的出處（history，越近越優先）；章數不合的書卷會被跳過
    （例：前一個出處是提前三15，「十五1~8」就往前找到約翰福音）。"""
    sc = _Scanner(text, verse_count)
    refs, spans = [], []
    points = list(points)
    for m in IMPLICIT_TRIGGER.finditer(text):
        i = m.end()
        if any(a <= m.start() < b or a <= i < b for a, b in taken + spans):
            continue
        cands = [(bk, ch) for end, bk, ch in sorted(points) if end <= m.start() + 1][::-1]
        cands += list(reversed(history))
        if not cands:
            continue
        out = []
        cv = sc.cn_verse(i)
        if cv:                                                # —三四25上
            book = next((bk for bk, _ in cands if cv[0] <= BOOKS[bk]["chapters"]), None)
            if book is None:
                continue
            j = sc.verse_list(cv[1], book, cv[0], "implicit", i, out, src_start=i)
        elif VERSE_RE.match(text, i):                         # —10~11、25節 / （2~42）
            book, chapter = _first_with_chapter(cands)
            if not chapter:
                continue
            j = sc.verse_list(i, book, chapter, "implicit", i, out)
            seg = text[i:j]
            paren = m.group() in "（("
            closed = paren and j < len(text) and (
                text[j] in "）)"
                or (text[j] in "，、；;。" and text[j + 1:j + 2] in ("）", ")"))   # （20，27，）（13下；）
                or (text[j] in "，、" and ABBR_RE.match(text, j + 1)))   # （4、19，路八11）
            if not ("節" in seg or "节" in seg or closed or (m.group() == "—" and sc.at_end(j))):
                continue
        else:
            continue
        if not out or not sc.at_end(j):
            continue
        for r in out:
            r.notes.append(f"由上下文推斷書卷為「{BOOKS[r.book]['abbr']}」")
        refs.extend(out)
        spans.append((i, j))
        points.append((j, out[-1].book, out[-1].chapter))   # 同段後面的出處沿用這一個
    return refs, spans


def _first_with_chapter(cands):
    """最近一個有章的上下文；書名提及（沒有章）則找同一卷書較早的章。"""
    for n, (bk, ch) in enumerate(cands):
        if ch:
            return bk, ch
        same = next((c for b, c in cands[n + 1:] if b == bk and c), None)
        if same:
            return bk, same
    return None, None


def parse(text, verse_count, context=None):
    """解析一段綱要文字 → (依原文順序排列的 refs, 已解析範圍 spans, 給下一段用的上下文歷史)。

    context：前面段落出處的 (書, 章) 歷史清單（越後面越近）；也接受單一 (書, 章)。"""
    if context is None:
        history = []
    elif isinstance(context, tuple):
        history = [context]
    else:
        history = list(context)
    formal, fspans = parse_formal(text, verse_count)
    prose, pspans = parse_prose(text, verse_count, fspans)
    taken = fspans + pspans
    # 上下文點：每個已解析出處的結尾，以及句中提到的書名（例：「如以斯帖記所指明」）
    points = []
    for a, b in taken:
        inside = [r for r in formal + prose if r.pos == a]
        if inside:
            points.append((b, inside[-1].book, inside[-1].chapter))
    mentions = []
    for m in MENTION_RE.finditer(text):
        if not any(a <= m.start() < b for a, b in taken):
            points.append((m.end(), NAME_TO_NO[m.group()], None))
            mentions.append((m.start(), NAME_TO_NO[m.group()]))
    implicit, ispans = parse_implicit(text, verse_count, taken, points, history)
    refs = formal + prose + implicit
    order = {id(r): n for n, r in enumerate(refs)}
    refs.sort(key=lambda r: (r.pos, order[id(r)]))
    # 上下文歷史依原文順序：出處記 (書, 章)；書名提及記 (書, None)，
    # 例：「在創世記這卷書中…：」底下各點的「—十二1」就是創世記
    events = [(r.pos, 0, r.book, r.chapter) for r in refs] + [(a, 1, bk, None) for a, bk in mentions]
    for _, _, bk, ch in sorted(events, key=lambda e: (e[0], e[1])):
        if not history or history[-1] != (bk, ch):
            history.append((bk, ch))
    return refs, sorted(taken + ispans), history[-12:]


def unparsed_fragments(text, spans):
    """含阿拉伯數字卻沒被解析的片段（給校對報告用）。"""
    chars = list(text)
    num = re.match(r"\s*[^\t]{1,4}\t", text)          # 綱要編號「1＜Tab＞」「十一＜Tab＞」
    if num:
        spans = list(spans) + [(0, num.end())]
    for a, b in spans:
        for k in range(a, b):
            chars[k] = " "
    rest = "".join(chars)
    frags = []
    for m in re.finditer(r"[^\s，、；。：:）)（(—]*\d+[^\s，、；。：:）)（(—]*", rest):
        frag = m.group()
        if re.search(r"\d{4}年", frag):
            continue
        frags.append(frag)
    return frags
