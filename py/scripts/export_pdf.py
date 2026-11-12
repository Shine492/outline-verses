"""docx → pdf（LibreOffice）。

綱要常用的字型（方正楷体_GBK、方正书宋_GBK、DFKai-SB…）不一定有安裝，或 LibreOffice
認不得它的英文名稱，轉 PDF 時會被換成很怪的字（例如行楷）。所以轉檔前先在暫存複本裏
把字型名稱換成 LibreOffice 認得的替代字型；Word 檔本身不動。

用法：python3 scripts/export_pdf.py <檔案.docx> [輸出.pdf]
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

if sys.platform == "win32":
    # Windows 內建字型用英文名稱最穩（LibreOffice 在非中文介面不一定認得中文名稱）；
    # 繁中 Windows 預設沒有仿宋、黑體，改用楷書、正黑體
    KAI, SONG, HEI, FANGSONG = "DFKai-SB", "PMingLiU", "Microsoft JhengHei", "DFKai-SB"
    OK_FONTS = {KAI, SONG, HEI, "MingLiU", "SimSun", "Microsoft YaHei", "Times New Roman", "Arial"}
else:
    # LibreOffice（macOS）認得的字型名稱：要用中文名稱
    KAI, SONG, HEI, FANGSONG = "標楷體", "宋體-繁", "黑體-繁", "华文仿宋"
    OK_FONTS = {KAI, SONG, HEI, FANGSONG, "宋體-簡", "黑體-簡", "微軟正黑體", "教育部標準楷書",
                "華康正顏楷體W5", "Times New Roman", "Arial"}
EXPLICIT = {
    "DFKai-SB": KAI, "BiauKai": KAI, "KaiTi": KAI, "Kaiti TC": KAI, "Kaiti SC": KAI,
    "PMingLiU": SONG, "MingLiU": SONG, "SimSun": SONG, "NSimSun": SONG, "Songti TC": SONG,
    "SimHei": HEI, "Microsoft YaHei": HEI, "微软雅黑": HEI, "Heiti TC": HEI,
    "FangSong": FANGSONG, "STFangsong": FANGSONG,
}
KEYWORDS = [("仿宋", FANGSONG), ("楷", KAI), ("黑", HEI), ("宋", SONG), ("明", SONG)]
CJK = re.compile(r"[㐀-鿿]")
FONT_ATTR = re.compile(r'(w:(?:ascii|hAnsi|eastAsia|cs)=")([^"]+)(")')
THEME_ATTR = re.compile(r'(<a:(?:ea|latin) typeface=")([^"]+)(")')
PROFILE = Path(tempfile.gettempdir()) / "rv_outline_lo_profile"


def substitute(name):
    if name in OK_FONTS:
        return name
    if name in EXPLICIT:
        return EXPLICIT[name]
    if CJK.search(name):
        for key, font in KEYWORDS:
            if key in name:
                return font
    return name


def _rewrite_fonts(src, dst):
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith(".xml") and item.filename.startswith("word/"):
                xml = data.decode("utf-8")
                xml = FONT_ATTR.sub(lambda m: m.group(1) + substitute(m.group(2)) + m.group(3), xml)
                xml = THEME_ATTR.sub(lambda m: m.group(1) + substitute(m.group(2)) + m.group(3), xml)
                data = xml.encode("utf-8")
            zout.writestr(item, data)


def find_soffice():
    """找 LibreOffice：PATH 上的 soffice，或 macOS / Windows 的預設安裝位置。"""
    found = shutil.which("soffice")
    if found:
        return found
    cands = ["/Applications/LibreOffice.app/Contents/MacOS/soffice"]
    for env in ("ProgramFiles", "ProgramFiles(x86)"):
        if os.environ.get(env):
            cands.append(str(Path(os.environ[env]) / "LibreOffice" / "program" / "soffice.exe"))
    for c in cands:
        if Path(c).exists():
            return c
    raise RuntimeError("找不到 LibreOffice（soffice）。請到 https://www.libreoffice.org 下載安裝，"
                       "或先加 --no-pdf 只輸出 Word。")


def export_pdf(docx_path, pdf_path=None):
    docx_path = Path(docx_path)
    pdf_path = Path(pdf_path) if pdf_path else docx_path.with_suffix(".pdf")
    soffice = find_soffice()
    with tempfile.TemporaryDirectory() as tmp:
        tmp_docx = Path(tmp) / "outline.docx"
        _rewrite_fonts(docx_path, tmp_docx)
        # 用獨立的設定檔，避免使用者正開著 LibreOffice 時轉檔失敗
        r = subprocess.run([soffice, f"-env:UserInstallation={PROFILE.as_uri()}", "--headless",
                            "--convert-to", "pdf", "--outdir", tmp, str(tmp_docx)],
                           capture_output=True, text=True, timeout=300)
        out = Path(tmp) / "outline.pdf"
        if not out.exists():
            raise RuntimeError(f"PDF 轉檔失敗：{r.stdout}\n{r.stderr}")
        shutil.move(str(out), pdf_path)
    return pdf_path


if __name__ == "__main__":
    print(export_pdf(*sys.argv[1:3]))
