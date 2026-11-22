# docx-merger

把資料夾裡的多份 Word 檔（.docx）依檔名順序合併成一份，並讓頁碼與目錄頁次自動正確。

## 這個工具解決什麼問題

一份報告常拆成很多個 docx 分頭撰寫（封面、目錄、各章節），每份各自設定「從第幾頁開始」，
目錄的頁次也是手動填的。只要其中一份增減一頁，後面所有檔案的起始頁碼與目錄就全部要重算。

這支腳本合併時會：

1. 保留每份檔案的分節、頁面設定，每份都從新的一頁開始；
   第 4 份起每份第 1 頁都排在奇數頁，前一份結束在奇數頁時自動補一張空白頁（方便雙面列印）
2. 拿掉各份寫死的起始頁碼（只保留第一份的），頁碼改為一路連續
3. 把目錄檔裡手打的頁次數字，換成指向各份檔案開頭的 `PAGEREF` 欄位
4. 在 Windows 上，自動在背景開啟本機的 Word 重算目錄頁次並存檔；
   沒裝 Word 或不是 Windows 時，改設定「開檔時更新欄位」

## 需求

- Python 3.10 以上
- [uv](https://docs.astral.sh/uv/)

腳本內含 [PEP 723](https://peps.python.org/pep-0723/) 標頭，`uv run` 會自行安裝 `docxcompose`。

## 使用

```bash
uv run merge_docx.py ~/Downloads/報告                    # 輸出到 ~/Downloads/合併_20261122.docx（當天日期）
uv run merge_docx.py ~/Downloads/報告 -o 報告-完整版.docx
uv run merge_docx.py ~/Downloads/報告 --toc 02_目錄.docx  # 手動指定目錄檔
uv run merge_docx.py ~/Downloads/報告 --no-toc            # 不改目錄頁次
uv run merge_docx.py ~/Downloads/報告 --odd-from 0        # 不強制從奇數頁開始
uv run merge_docx.py ~/Downloads/報告 --odd-from 5        # 改成第 5 份起才從奇數頁開始
```

### Windows 執行檔

到 GitHub 的 Actions 頁面，打開最新一次「Build Windows exe」，下載 `docx-merger-windows`，
解壓縮後得到 `docx-merger.exe`，不需要安裝 Python。使用方式二選一：

- 把報告資料夾拖到 `docx-merger.exe` 圖示上
- 雙擊 `docx-merger.exe`，再把報告資料夾拖進跳出來的視窗，按 Enter

合併檔會輸出在報告資料夾旁邊，檔名是 `合併_` 加上執行當天日期，例如 `合併_20261122.docx`。

畫面顯示「目錄頁次已由 Word 重算並存檔」就完成了。若顯示「無法自動重算目錄頁次」，
請用 Word 開啟合併檔，按 `Ctrl+A` 全選，再按 `F9` 更新欄位。

## 資料夾與目錄的約定

- 檔名決定合併順序，建議用 `01_`、`02_` 開頭。Word 開檔產生的 `~$` 暫存檔會自動略過。
- 目錄檔預設是檔名含「目錄」的那份。
- 目錄裡「整格只有數字」的段落視為頁次，依出現順序對應到第 1、2、3⋯份檔案的開頭。
  數量對不上時會印出警告，目錄維持原樣，其他部分照常合併。

## 已知限制

- 頁首頁尾一律沿用第一份檔案的設定。若各份的頁首頁尾內容不同，需要另外處理。
- 目錄頁次要靠 Word 依實際排版計算。在 Linux／macOS 上執行時無法自動重算，需在 Word 裡按 `Ctrl+A`、`F9`；
  LibreOffice 也會自行計算，但因字型差異，頁數可能與 Word 不同。
