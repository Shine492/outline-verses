"""網頁版附經文：在瀏覽器（Pyodide）裏跑和指令列版同一套程式。

網站上不放經文：用到哪一章，才由使用者的瀏覽器向恢復本網站（recoveryversion.com.tw）取那一章。
做法：先規劃一次，記下缺哪些章 → 一起下載 → 再規劃，直到不缺（範圍跨章時第二輪才知道要下一章）。
"""
import asyncio
import json
from io import BytesIO
from pathlib import PurePath

from docx import Document

from add_verses import ensure_style, format_document, insert_verses, plan_verses, render_report
from bible import DATA, Bible
from books import BOOKS

API = ("https://www.recoveryversion.com.tw/api/getVerses?VERSION=1"
       "&output[]=content&output[]=unit_code&output[]=segment_code"
       "&chapter_code={book}&section_code={chapter}&ORDER=id")
MAX_ROUNDS = 8
PARALLEL = 6


class WebBible(Bible):
    """經文從 self.data 取（格式同 rv_zht.json）；還沒下載的章記在 missing。"""

    def __init__(self):
        v = json.loads((DATA / "variants.json").read_text(encoding="utf-8"))
        self.data = {}
        self.replace = v["replace"]
        self.keep = v.get("keep", [])
        self.overrides = {}
        self.missing = set()

    def _chapter(self, book, chapter):
        book, chapter = int(book), int(chapter)
        if book not in BOOKS or not 0 < chapter <= BOOKS[book]["chapters"]:
            return {}
        chap = self.data.get(str(book), {}).get(str(chapter))
        if chap is None:
            self.missing.add((book, chapter))
            return {}
        return chap

    def verse_count(self, book, chapter):
        return max((int(k) for k in self._chapter(book, chapter) if k != "0"), default=0)

    def units(self, book, chapter, verse):
        return self._chapter(book, chapter).get(str(verse))

    def add_chapter(self, book, chapter, rows):
        """API 回傳的一章 → {節:[單位文字,...]}（同 crawl_rv.build）。"""
        chap = {}
        for row in sorted(rows, key=lambda r: (r["segment_code"], r["unit_code"])):
            chap.setdefault(str(row["segment_code"]), []).append(row["content"])
        self.data.setdefault(str(book), {})[str(chapter)] = chap


async def pyodide_fetch(book, chapter):
    from pyodide.http import pyfetch
    r = await pyfetch(API.format(book=book, chapter=chapter))
    return await r.json()


async def _fetch(bible, book, chapter, sem, fetch_rows):
    async with sem:
        for attempt in range(3):
            try:
                rows = await fetch_rows(book, chapter)
                if isinstance(rows, list) and rows:
                    bible.add_chapter(book, chapter, rows)
                    return
            except Exception:
                pass
            await asyncio.sleep(1 + attempt)
    raise RuntimeError(f"無法從恢復本網站取得 {BOOKS[book]['abbr']}{chapter}，請檢查網路後重試。")


_bible = None


async def run(data, filename, progress=None, fetch_rows=pyodide_fetch):
    """綱要 docx（bytes）→ {"docx": 附經文的 docx bytes, "name", "report": 校對報告 md, "points", "verses"}。"""
    global _bible
    if _bible is None:
        _bible = WebBible()          # 同一個分頁裏處理多份綱要時，已下載的章沿用
    bible = _bible
    data = bytes(data)
    for _ in range(MAX_ROUNDS):
        bible.missing.clear()
        doc = Document(BytesIO(data))
        ensure_style(doc)
        plan, report, stats = plan_verses(doc, bible)
        if not bible.missing:
            break
        if progress:
            progress(f"向恢復本網站取經文（{len(bible.missing)} 章）…")
        sem = asyncio.Semaphore(PARALLEL)
        await asyncio.gather(*(_fetch(bible, b, c, sem, fetch_rows) for b, c in sorted(bible.missing)))
    else:
        raise RuntimeError("經文一直取不完整，請重試。")

    format_document(doc)
    insert_verses(doc, plan)
    out = BytesIO()
    doc.save(out)
    src = PurePath(filename)
    name = f"{src.stem}（附經文）.docx"
    return {"docx": out.getvalue(), "name": name,
            "report": render_report(src, PurePath(name), stats, report),
            "points": stats["points"], "verses": stats["verses"]}
