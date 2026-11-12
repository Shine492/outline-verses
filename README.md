# 綱要附經文

網址：https://shine492.github.io/outline-verses/

上傳綱要 Word 檔（.docx）或 PDF，會在每一點下面附上它引用的中文恢復本經文，並列出需要人工確認的地方（半節切分、由上下文推斷的出處等）。

- 檔案只在使用者的瀏覽器裏處理（用 [Pyodide](https://pyodide.org) 執行 Python），不會上傳到任何伺服器。
- 本網站不存放經文：用到哪一章，才由使用者的瀏覽器向[恢復本網站](https://www.recoveryversion.com.tw)取得那一章。恢復本經文有版權，只供個人查考與服事使用。
- 只輸出 Word；要 PDF 請用 Word 開啟後另存。
- PDF 用 [pdf.js](https://mozilla.github.io/pdf.js/) 讀出文字，依行首位置和編號接回段落，再建成 Word 檔處理；原本的字型和版面不保留。

## 檔案

- `index.html`：網頁
- `py/scripts/web.py`：網頁版流程（缺哪些章就下載、再規劃，直到齊全）
- `py/scripts/` 其餘檔案、`py/data/variants.json`：與 Claude Code skill「outline-verses」相同的程式，從那裏複製過來
