"""附經文的 docx → 版面資料（JSON），網頁再用 pdf-lib 排成 PDF（index.html 的 makePdf）。

只處理綱要用得到的格式：縮排、Tab、自動編號、字級、粗體、顏色、置中、段前段後、行高、讀經的灰底框、分頁。
表格、圖片不處理。單位都是 pt。
"""
import re

from docx.oxml.ns import qn

from add_verses import _num_pr, _style_chain, para_indent

EMU_PER_PT = 12700
DEFAULT_SIZE = 12
LINE_FACTOR = 1.3            # 單行行高 ≈ 字級 × 1.3（標楷體一類的中文字型）

CN_DIGITS = "〇一二三四五六七八九"
LEGAL = "〇壹貳參肆伍陸柒捌玖"
TIANGAN = "甲乙丙丁戊己庚辛壬癸"
DIZHI = "子丑寅卯辰巳午未申酉戌亥"


def _val(el, tag):
    e = el.find(qn(f"w:{tag}")) if el is not None else None
    return None if e is None else e.get(qn("w:val"))


def _on(e):
    return e is not None and e.get(qn("w:val"), "1") not in ("0", "false", "off")


def _cn_count(n, digits=CN_DIGITS, ten="十"):
    """1 → 一、11 → 十一、25 → 二十五（綱要編號用到 99 就夠）。"""
    if n <= 0 or n >= 100:
        return str(n)
    t, o = divmod(n, 10)
    if not t:
        return digits[o]
    return ("" if t == 1 else digits[t]) + ten + (digits[o] if o else "")


def _roman(n):
    out = ""
    for v, s in ((1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"),
                 (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")):
        while n >= v:
            out, n = out + s, n - v
    return out


def _letter(n):
    return chr(ord("a") + (n - 1) % 26) * ((n - 1) // 26 + 1)


def format_number(n, fmt):
    if fmt in ("lowerLetter", "upperLetter"):
        s = _letter(n)
        return s.upper() if fmt == "upperLetter" else s
    if fmt in ("lowerRoman", "upperRoman"):
        s = _roman(n)
        return s.upper() if fmt == "upperRoman" else s
    if fmt in ("taiwaneseCounting", "taiwaneseCountingThousand", "chineseCounting",
               "chineseCountingThousand", "ideographDigital", "taiwaneseDigital"):
        return _cn_count(n)
    if fmt in ("ideographLegalTraditional", "chineseLegalSimplified"):
        return _cn_count(n, LEGAL, "拾")
    if fmt == "ideographTraditional":
        return TIANGAN[(n - 1) % 10]
    if fmt == "ideographZodiac":
        return DIZHI[(n - 1) % 12]
    if fmt == "decimalFullWidth":
        return "".join(chr(ord(c) + 0xFEE0) for c in str(n))
    return str(n)


class Numbering:
    """自動編號的文字（壹、一、a…）；每個編號清單（numId）各自計數。"""

    def __init__(self, doc):
        try:
            self.el = doc.part.numbering_part.element
        except (NotImplementedError, KeyError):
            self.el = None
        self.counters = {}

    def _lvl(self, num_id, ilvl):
        """→ (w:lvl, 起始值)"""
        num = next((n for n in self.el.findall(qn("w:num")) if n.get(qn("w:numId")) == num_id), None)
        if num is None:
            return None, 1
        start = None
        for ov in num.findall(qn("w:lvlOverride")):
            if ov.get(qn("w:ilvl")) == str(ilvl):
                so = _val(ov, "startOverride")
                start = int(so) if so else None
                lvl = ov.find(qn("w:lvl"))
                if lvl is not None:
                    return lvl, start or int(_val(lvl, "start") or 1)
        abs_id = _val(num, "abstractNumId")
        an = next((a for a in self.el.findall(qn("w:abstractNum")) if a.get(qn("w:abstractNumId")) == abs_id), None)
        lvl = next((l for l in an.findall(qn("w:lvl")) if l.get(qn("w:ilvl")) == str(ilvl)), None) if an is not None else None
        if lvl is None:
            return None, 1
        return lvl, start or int(_val(lvl, "start") or 1)

    def label(self, p):
        """→ (編號文字, 後面接什麼：tab/space/nothing)；沒有編號回傳 None。"""
        npr = _num_pr(p)
        if npr is None or self.el is None:
            return None
        num_id, ilvl = npr[0], int(npr[1])
        lvl, start = self._lvl(num_id, ilvl)
        if lvl is None:
            return None
        c = self.counters.setdefault(num_id, {})
        c[ilvl] = c.get(ilvl, start - 1) + 1
        for k in [k for k in c if k > ilvl]:
            del c[k]

        def sub(m):
            k = int(m.group(1)) - 1
            lv, st = self._lvl(num_id, k)
            return format_number(c.get(k, st), _val(lv, "numFmt") or "decimal")
        fmt = _val(lvl, "numFmt")
        text = _val(lvl, "lvlText") or ""
        if fmt == "none":
            text = ""
        elif fmt == "bullet":
            text = text if text and ord(text[0]) < 0xF000 else "•"   # Symbol/Wingdings 字型的符號換成圓點
        else:
            text = re.sub(r"%(\d)", sub, text)
        return text, _val(lvl, "suff") or "tab"


class Layout:
    def __init__(self, doc):
        self.doc = doc
        styles = doc.styles.element
        self.style_el = {s.get(qn("w:styleId")): s for s in styles.findall(qn("w:style"))}
        dd = styles.find(qn("w:docDefaults"))
        self.default_rpr = dd.find(f"{qn('w:rPrDefault')}/{qn('w:rPr')}") if dd is not None else None
        self.default_ppr = dd.find(f"{qn('w:pPrDefault')}/{qn('w:pPr')}") if dd is not None else None
        self.numbering = Numbering(doc)

    def _char_style_rprs(self, sid):
        out, seen = [], set()
        while sid and sid not in seen and sid in self.style_el:
            seen.add(sid)
            s = self.style_el[sid]
            out.append(s.find(qn("w:rPr")))
            sid = _val(s, "basedOn")
        return out

    def _para_rprs(self, p):
        return [s.element.find(qn("w:rPr")) for s in _style_chain(p.style)] + [self.default_rpr]

    def _para_pprs(self, p):
        return [p._p.pPr] + [s.element.pPr for s in _style_chain(p.style)] + [self.default_ppr]

    @staticmethod
    def _props(rprs):
        size = bold = color = None
        sup = False
        for rpr in rprs:
            if rpr is None:
                continue
            if size is None and _val(rpr, "sz"):
                size = int(_val(rpr, "sz")) / 2
            if bold is None and rpr.find(qn("w:b")) is not None:
                bold = _on(rpr.find(qn("w:b")))
            c = _val(rpr, "color")
            if color is None and c and c != "auto":
                color = c
            sup = sup or _val(rpr, "vertAlign") == "superscript"
        return {"size": size or DEFAULT_SIZE, "bold": bool(bold), "color": color or "000000"}, sup

    def runs(self, p):
        """段落文字 → ([{text, size, bold, color}], 有沒有分頁符號)；略過上標（註腳號碼）。"""
        para = self._para_rprs(p)
        out, page_break = [], False
        for r in p._p.iter(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            props, sup = self._props([rpr] + self._char_style_rprs(_val(rpr, "rStyle")) + para)
            if sup:
                continue
            text = ""
            for child in r:
                tag = child.tag.split("}")[-1]
                if tag == "t":
                    text += child.text or ""
                elif tag == "tab":
                    text += "\t"
                elif tag == "br":
                    if child.get(qn("w:type")) == "page":
                        page_break = True
                    else:
                        text += "\n"
            if not text:
                continue
            if out and all(out[-1][k] == props[k] for k in props):
                out[-1]["text"] += text
            else:
                out.append({"text": text, **props})
        return out, page_break

    def paragraph(self, p):
        pprs = self._para_pprs(p)
        runs, page_break = self.runs(p)
        base, _ = self._props(self._para_rprs(p))
        style = runs[0] if runs else base
        left, first = para_indent(p)

        def first_attr(tag, attr):
            for ppr in pprs:
                e = ppr.find(qn(f"w:{tag}")) if ppr is not None else None
                if e is not None and e.get(qn(f"w:{attr}")) is not None:
                    return e.get(qn(f"w:{attr}"))
            return None
        before = int(first_attr("spacing", "before") or 0) / 20
        after = int(first_attr("spacing", "after") or 0) / 20
        line, rule = first_attr("spacing", "line"), first_attr("spacing", "lineRule")
        if line and rule in ("exact", "atLeast"):
            spacing = {rule: int(line) / 20}
        else:
            spacing = {"mult": int(line) / 240 if line else 1.0}
        align = next((_val(ppr, "jc") for ppr in pprs if _val(ppr, "jc")), "left")
        box = None
        for ppr in pprs[:-1]:
            if ppr is None:
                continue
            shd, bdr = ppr.find(qn("w:shd")), ppr.find(qn("w:pBdr"))
            if shd is not None or bdr is not None:
                fill = shd.get(qn("w:fill")) if shd is not None else None
                top = bdr.find(qn("w:top")) if bdr is not None else None
                box = {"fill": fill if fill and fill != "auto" else None,
                       "border": (top.get(qn("w:color")) or "000000") if top is not None else None}
                break

        head = None
        label = self.numbering.label(p)
        if label and label[0]:
            text, suff = label
            if suff == "tab":
                head = [{**style, "text": text}]
            else:
                runs = [{**style, "text": text + (" " if suff == "space" else "")}] + runs
        elif any("\t" in r["text"] for r in runs):
            head, rest, after_tab = [], [], False         # 第一個 Tab 之前是編號或經文出處
            for r in runs:
                if after_tab:
                    rest.append(r)
                elif "\t" in r["text"]:
                    a, _, b = r["text"].partition("\t")
                    if a:
                        head.append({**r, "text": a})
                    if b:
                        rest.append({**r, "text": b})
                    after_tab = True
                else:
                    head.append(r)
            runs = rest
        for r in runs:
            r["text"] = r["text"].replace("\t", "  ")
        return {"left": left / 20, "first": first / 20, "align": align, "before": before, "after": after,
                "spacing": spacing, "box": box, "head": head, "runs": runs, "size": style["size"],
                "pageBreakBefore": page_break or any(_on(ppr.find(qn("w:pageBreakBefore")))
                                                     for ppr in pprs[:1] if ppr is not None)}

    def page(self):
        s = self.doc.sections[0]

        def pt(v, default):
            return int(v) / EMU_PER_PT if v is not None else default
        return {"width": pt(s.page_width, 595.3), "height": pt(s.page_height, 841.9),
                "top": pt(s.top_margin, 72), "bottom": pt(s.bottom_margin, 72),
                "left": pt(s.left_margin, 90), "right": pt(s.right_margin, 90)}


def layout(doc):
    lay = Layout(doc)
    return {"page": lay.page(), "paras": [lay.paragraph(p) for p in doc.paragraphs]}
