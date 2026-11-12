"""PDF 綱要 → docx：讀出每段文字和縮排，重新排成 Word 檔，之後照一般綱要附經文。

文字來源有兩種，格式相同（每頁一個 list，每個文字片段 {"str", "x", "y", "w", "h"}，單位 pt，y 從頁面下緣算起）：
  - 指令列：pdfplumber（extract_items）
  - 網頁：pdf.js（index.html 的 pdfItems）
原本的字型和版面不保留；掃描的圖片 PDF 讀不出文字。
"""
import re
from io import BytesIO

# 綱要編號：壹、一、1、a、(1)…（後面隔一段空白才接文字，才算編號）
MARKER_RE = re.compile(r"^(?:[壹貳叁參肆伍陸柒捌玖拾]+|[一二三四五六七八九十]+|\d{1,2}|[a-z]|[（(]\d{1,2}[)）])[、.．]?$")
ASCII_RE = re.compile(r"[A-Za-z0-9]")
PAGE_NO_RE = re.compile(r"^[-–—\s]*\d+[-–—\s]*$")
CJK_SPACE_RE = re.compile(r"(?<![A-Za-z0-9]) | (?![A-Za-z0-9])")   # 中文旁邊的空白（排版留下的）


def extract_items(path):
    """用 pdfplumber 讀出每頁的文字片段。"""
    import pdfplumber
    pages = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(x_tolerance=1, y_tolerance=2, extra_attrs=["size"])
            pages.append([{"str": w["text"], "x": w["x0"], "y": page.height - w["bottom"],
                           "w": w["x1"] - w["x0"], "h": w["size"]} for w in words])
    return pages


def _median(values):
    s = sorted(values)
    return s[len(s) // 2] if s else 12


def _ascii(s, i):
    return bool(s) and bool(ASCII_RE.match(s[i]))


def page_lines(items):
    """一頁的文字片段 → 行：{x, right, size, chunks:[{x, right, str}], marker}。"""
    parts = [{"str": it["str"], "x": it["x"], "y": it["y"], "right": it["x"] + it["w"], "h": it["h"]}
             for it in items if it["str"].strip()]
    parts.sort(key=lambda p: (-p["y"], p["x"]))
    lines = []
    for p in parts:
        line = next((l for l in lines if abs(l["y"] - p["y"]) < max(l["h"], p["h"]) * 0.5), None)
        if line:
            line["parts"].append(p)
            line["h"] = max(line["h"], p["h"])
        else:
            lines.append({"y": p["y"], "h": p["h"], "parts": [p]})
    out = []
    for l in sorted(lines, key=lambda l: -l["y"]):
        big = max(p["h"] for p in l["parts"])
        ps = sorted((p for p in l["parts"] if p["h"] >= big * 0.7), key=lambda p: p["x"])   # 略過上標（註腳號碼）
        size = _median([p["h"] for p in ps])
        # 行首是編號、後面隔一段（Tab 留下的空白）才接文字 → 編號自成一塊
        marker = (len(ps) > 1 and bool(MARKER_RE.match(ps[0]["str"].strip()))
                  and ps[1]["x"] - ps[0]["right"] > size * 0.3)
        chunks = []
        for i, p in enumerate(ps):
            last = chunks[-1] if chunks else None
            gap = p["x"] - last["right"] if last else 0
            if last and gap < size * 0.8 and not (marker and i == 1):
                sp = " " if gap > size * 0.2 and _ascii(last["str"], -1) and _ascii(p["str"], 0) else ""
                last["str"] += sp + p["str"]
                last["right"] = max(last["right"], p["right"])
            else:
                chunks.append({"x": p["x"], "right": p["right"], "str": p["str"]})
        for c in chunks:
            c["str"] = c["str"].strip()
        out.append({"x": chunks[0]["x"], "right": max(c["right"] for c in chunks), "size": size,
                    "chunks": chunks, "marker": marker})
    return out


def paragraphs(pages):
    """每頁的文字片段 → [{text, left, first, size}]：依行首位置和上一行有沒有寫滿，把行接回段落。
    有編號的段落寫成「編號＜Tab＞文字」；left、first 是 twips（first 負值＝懸掛縮排）。"""
    lines = []
    for items in pages:
        ls = page_lines(items)
        if ls and PAGE_NO_RE.match("".join(c["str"] for c in ls[-1]["chunks"])):
            ls.pop()                                                       # 頁碼
        lines += ls
    if not lines:
        raise ValueError("這份 PDF 讀不出文字（可能是掃描的圖片），請改用 Word 檔。")
    margin = min(l["x"] for l in lines)
    right_edge = max(l["right"] for l in lines)

    def near(a, b, size):
        return b is not None and abs(a - b) < size * 0.6

    paras, cur = [], None
    for l in lines:
        first = l["chunks"][0]["str"]
        marker = l["marker"] or (len(l["chunks"]) > 1 and bool(MARKER_RE.match(first)))
        body = CJK_SPACE_RE.sub("", " ".join(c["str"] for c in (l["chunks"][1:] if marker else l["chunks"])))
        short = l["right"] < right_edge - l["size"] * 1.5
        # 接續上一段：懸掛縮排的第二行以後（行首對齊編號後的文字）一定是接續；
        # 其他情形要上一行寫滿、行首對齊上一段
        hanging = cur is not None and cur["textX"] is not None and not near(cur["textX"], cur["firstX"], l["size"])
        if (cur and not marker and abs(l["size"] - cur["size"]) < 1.5
                and ((hanging and near(l["x"], cur["textX"], l["size"]))
                     or (not cur["ended"] and any(near(l["x"], cur[k], l["size"]) for k in ("textX", "firstX", "contX"))))):
            sp = " " if _ascii(cur["text"], -1) and _ascii(body, 0) else ""
            cur["text"] += sp + body
            if cur["contX"] is None:
                cur["contX"] = l["x"]
            cur["ended"] = short
            continue
        cur = {"firstX": l["x"], "textX": l["chunks"][1]["x"] if marker else None, "contX": None,
               "size": l["size"], "ended": short, "text": f"{first}\t{body}" if marker else body}
        paras.append(cur)

    def tw(pt):
        return round((pt - margin) * 20)
    out = []
    for p in paras:
        left = next(v for v in (p["textX"], p["contX"], p["firstX"]) if v is not None)
        out.append({"text": p["text"], "left": tw(left), "first": tw(p["firstX"]) - tw(left),
                    "size": round(p["size"] * 2) / 2})
    return out


def build_docx(paras):
    """段落 → docx bytes。"""
    from docx import Document
    from docx.oxml.ns import qn
    from docx.shared import Pt, Twips
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "標楷體")
    normal.paragraph_format.space_after = Pt(0)
    for p in paras:
        para = doc.add_paragraph()
        para.paragraph_format.left_indent = Twips(p["left"])
        para.paragraph_format.first_line_indent = Twips(p["first"])
        head, tab, rest = p["text"].partition("\t")
        para.add_run(head).font.size = Pt(p["size"])
        if tab:
            r = para.add_run()
            r.add_tab()
            r.add_text(rest)
            r.font.size = Pt(p["size"])
    out = BytesIO()
    doc.save(out)
    return out.getvalue()


def pdf_to_docx(path):
    return build_docx(paragraphs(extract_items(path)))
