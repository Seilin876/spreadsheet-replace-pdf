# 批次替換 Excel／CSV 文字並另存 PDF

## Windows 64 位元

從 [GitHub Releases](https://github.com/Seilin876/spreadsheet-replace-pdf/releases/latest) 下載 `SpreadsheetReplacePDF-win-x64.exe`，雙擊執行。執行檔已包含 Python 和所需的 Python 套件。

Windows 版使用電腦上已安裝的桌面版 Microsoft Excel 匯出 PDF。首次使用前，請確認 Excel 能正常開啟及另存 PDF。Windows 版無需安裝 LibreOffice 或 Acrobat。

## 安裝（macOS）

先安裝 [LibreOffice](https://www.libreoffice.org/download/download-libreoffice/)。如果已有 Homebrew，也可以執行 `brew install --cask libreoffice`。

在終端機切換到本資料夾後執行以下指令。前兩行只需做一次。

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python replace_and_export.py
```

依序輸入資料夾路徑、要搜尋的文字（input）和替換文字（output）。路徑可從 Finder 拖入終端機。output 留空會刪除符合的文字。

- 處理指定資料夾這一層的 `.xlsx`、`.xlsm`、`.xls` 和 `.csv`。以區分大小寫的文字比對替換所有出現位置，並處理 Excel 所有工作表的文字儲存格。數值和公式保持原樣。
- 替換後以原檔名存回，並在同一個資料夾另存 PDF。例如 `報表.xlsx` 會另存為 `報表.xlsx.pdf`。同名 PDF 會覆寫。即使沒有符合的文字也會輸出 PDF。
- Excel 的 PDF 依工作簿的列印設定輸出。舊版 `.xls` 更動時會經過格式轉換，複雜的版面或功能可能有差異。含混合字型或顏色的儲存格在文字被替換後會使用儲存格的預設格式。
- CSV 支援 UTF-8、帶 BOM 的 UTF-16／UTF-32，以及依序嘗試 CP950、GB18030。替換後會保留原編碼，欄位會重新加上引號。如果新文字無法用原編碼儲存，該檔會列為失敗。
- PDF 產生失敗時，該原檔會保留，程式會繼續處理其他檔案。
