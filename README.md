# 批次替換 Excel／CSV 文字並另存 PDF

本工具在 Windows 上從 Python 原始碼執行，使用已安裝的桌面版 Microsoft Excel 內建功能匯出 PDF。程式操作介面為英文。

## 如何在新電腦安裝執行環境

### 1. 確認桌面版 Excel 可以匯出 PDF

開啟 Microsoft Excel，建立一份空白活頁簿並輸入任意文字，再用 **檔案 → 另存新檔** 將它存成 PDF。如果電腦只有網頁版 Excel，請透過公司 IT 安裝桌面版。 [Microsoft Excel PDF 匯出文件](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.exportasfixedformat)

### 2. 安裝 Python 3.12（64 位元）

優先使用公司的軟體中心或 IT 提供的安裝方式。如果公司允許使用 Python 官方安裝程式，安裝畫面有 **Add python.exe to PATH** 時可勾選。 [Python Windows 下載頁面](https://www.python.org/downloads/windows/)

安裝完成後，開啟新的 PowerShell 視窗並確認版本。

```powershell
py -3.12 --version
```

如果找不到 `py`，可以試 `python --version`。若顯示 `Python 3.12.x`，步驟 4 的第一行可改用 `python -m venv .venv`。

### 3. 解壓縮原始碼

在 GitHub 儲存庫頁面點 **Code → Download ZIP**。在檔案總管解壓縮後，開啟同時包含 `replace_and_export.py` 和 `requirements.txt` 的資料夾。在檔案總管上方的位址列輸入 `powershell` 並按 Enter，即可在該資料夾開啟 PowerShell。 [GitHub 儲存庫](https://github.com/Seilin876/spreadsheet-replace-pdf)

### 4. 建立虛擬環境並安裝套件

在 PowerShell 依序執行以下指令。這一步只需做一次。最後一行印出 `OK` 代表 `openpyxl` 與 `pywin32` 都可載入。

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -c "import openpyxl, win32com.client; print('OK')"
```

若公司網路限制套件下載，請向 IT 取得核准的 Python 套件來源。

## 執行程式

每次使用時，在程式所在資料夾的 PowerShell 執行以下指令。

```powershell
.\.venv\Scripts\python.exe .\replace_and_export.py
```

依序輸入以下資訊。

1. **Folder path** 輸入要處理的資料夾路徑。
2. **Text to find** 輸入要搜尋的文字，區分大小寫。
3. **Replacement text** 輸入替換文字，直接按 Enter 會刪除符合的文字。

第一次使用請先拿檔案副本試跑，因為程式會將替換結果存回原檔，並在同一資料夾另存 PDF。例如 `報表.xlsx` 對應 `報表.xlsx.pdf`，同名 PDF 會覆寫。

- 處理指定資料夾這一層的 `.xlsx`、`.xlsm`、`.xls` 和 `.csv`。Excel 會處理所有工作表的文字儲存格，並保留數值和公式。
- Excel 依工作簿的列印設定匯出 PDF。舊版 `.xls` 在替換時會經過格式轉換，複雜版面或功能可能有差異。含混合字型或顏色的儲存格在文字被替換後會使用儲存格的預設格式。
- CSV 支援 UTF-8、帶 BOM 的 UTF-16／UTF-32，以及依序嘗試 CP950、GB18030。替換後保留原編碼，欄位會重新加上引號。若新文字無法用原編碼儲存，該檔會列為失敗。
- 即使找不到符合的文字，也會匯出 PDF。若 PDF 匯出失敗，該原檔會保留，程式會繼續處理其他檔案。
