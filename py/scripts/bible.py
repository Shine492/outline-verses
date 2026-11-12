"""本機恢復本經文存取：取經文、字形轉換、半節（上/中/下）切分。"""
import json
import re
from difflib import SequenceMatcher
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"

PUNCT_STRENGTH = {"。": 4, "？": 4, "！": 4, "；": 3, "：": 2, "，": 1}
SOFT_END = "，；：、"
STRIP_RE = re.compile(r"[，。、；：？！「」『』“”‘’（）()〔〕《》…—\s·．]")
QUOTE_RE = re.compile(r"“([^“”]+)”|『([^『』]+)』|「([^「」]+)」")


class Bible:
    def __init__(self):
        self.data = json.loads((DATA / "rv_zht.json").read_text(encoding="utf-8"))
        v = json.loads((DATA / "variants.json").read_text(encoding="utf-8"))
        self.replace = v["replace"]          # [[網站字, 印刷字], ...]
        self.keep = v.get("keep", [])        # 不轉換的詞（例：顯著）
        p = DATA / "overrides.json"
        self.overrides = json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}

    # ---- 基本查詢 ----
    def verse_count(self, book, chapter):
        chap = self.data.get(str(book), {}).get(str(chapter), {})
        return max((int(k) for k in chap if k != "0"), default=0)

    def units(self, book, chapter, verse):
        return self.data.get(str(book), {}).get(str(chapter), {}).get(str(verse))

    # ---- 字形 ----
    def normalize(self, text):
        holders = {}
        for n, word in enumerate(self.keep):
            if word in text:
                key = f"{n}"
                holders[key] = word
                text = text.replace(word, key)
        for a, b in self.replace:
            text = text.replace(a, b)
        for key, word in holders.items():
            text = text.replace(key, word)
        return text

    # ---- 取經文 ----
    def text(self, ref, context_text=""):
        """回傳 (經文, 說明)；說明為 None 表示不需校對。整章回傳「從畧。」。"""
        if ref.verse is None:
            return "從畧。", "整章引用，依參考格式寫「從畧」"
        label = ref.label
        if label in self.overrides:
            return self.overrides[label], "使用 overrides.json 的人工設定"
        units = self.units(ref.book, ref.chapter, ref.verse)
        if not units:
            return None, "查無此節"
        units = [self.normalize(u) for u in units]
        if not ref.part:
            return "".join(units), None
        if len(units) > 1:
            idx = {"上": 0, "中": 1, "下": len(units) - 1}[ref.part]
            return _finish_upper(units[idx]) if ref.part != "下" else units[idx], "官方半節切分"
        return split_half(units[0], ref.part, self.normalize(context_text))


def _finish_upper(s):
    s = s.rstrip()
    if s and s[-1] in SOFT_END:
        s = s[:-1] + "。"
    return s


def split_half(text, part, context_text):
    """沒有官方切分時，自動判斷上/下半節的切點。"""
    body = text[:-1]
    cands = [(i + 1, PUNCT_STRENGTH[c]) for i, c in enumerate(body) if c in PUNCT_STRENGTH]
    if not cands:
        return text, "找不到可切分的標點，列出整節"
    top = max(s for _, s in cands)
    cuts = [c for c, s in cands if s == top]
    how = "依標點自動切分"
    if len(cuts) > 1:
        # 綱要文字若有引用這節的一部分，就以引用到的範圍決定切點
        hits = [k for k in range(len(text) - 2) if text[k:k + 3] in context_text]
        if hits and part == "上":
            need = max(hits) + 3
            after = [c for c in cuts if c >= need]
            cut = min(after) if after else max(cuts)
            how += "（參照綱要文字）"
        elif hits and part == "下":
            need = min(hits)
            before = [c for c in cuts if c <= need]
            cut = max(before) if before else min(cuts)
            how += "（參照綱要文字）"
        else:
            mid = len(text) / 2
            cut = min(cuts, key=lambda c: abs(c - mid))
    else:
        cut = cuts[0]
    if part == "上":
        return _finish_upper(text[:cut]), how
    if part == "下":
        return text[cut:], how
    return text, "「中」無法自動切分，列出整節"


def quoted_segments(text):
    """綱要文字裏用引號引起來的部分（“…”『…』「…」），以「｜」相連（比對時不會跨段）。"""
    return "｜".join("".join(m) for m in QUOTE_RE.findall(text))


def quoted_in(verse_text, para_text, ratio=0.6):
    """這節經文是否已在綱要文字裏被引用：最長的共同片段至少 8 字，且佔經文的 ratio 以上。"""
    a = STRIP_RE.sub("", verse_text)
    b = STRIP_RE.sub("", para_text)
    if not a:
        return False
    m = SequenceMatcher(None, a, b, autojunk=False).find_longest_match(0, len(a), 0, len(b))
    return m.size >= 8 and m.size >= ratio * len(a)
