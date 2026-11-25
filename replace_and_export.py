import codecs
from contextlib import contextmanager
import csv
from functools import partial
import io
import os
from pathlib import Path
import shlex
import shutil
import subprocess
import sys
from tempfile import TemporaryDirectory

from openpyxl import Workbook, load_workbook
from openpyxl.cell.rich_text import CellRichText
from openpyxl.utils import get_column_letter


SUPPORTED = {".xlsx", ".xlsm", ".xls", ".csv"}


def find_office():
    executable = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    return shutil.which("soffice") or (executable if os.access(executable, os.X_OK) else None)


def convert_libreoffice(office, profile, source, target_dir, extension):
    output_filter = {
        "xlsx": "xlsx:Calc Office Open XML",
        "xls": "xls:MS Excel 97",
        "pdf": "pdf:calc_pdf_Export",
    }[extension]
    target_dir.mkdir(exist_ok=True)
    result = subprocess.run(
        [office, f"-env:UserInstallation={profile.as_uri()}", "--headless",
         "--convert-to", output_filter, "--outdir", str(target_dir), str(source)],
        capture_output=True, text=True, errors="replace", timeout=180,
    )
    target = target_dir / f"{source.stem}.{extension}"
    if result.returncode or not target.is_file() or not target.stat().st_size:
        detail = (result.stderr or result.stdout).strip()
        raise RuntimeError(f"無法轉成 {extension}。{detail}")
    return target


def convert_excel(excel, source, target_dir, extension):
    from pythoncom import Missing

    target_dir.mkdir(exist_ok=True)
    target = target_dir / f"{source.stem}.{extension}"
    # 位置參數依序為檔名、更新連結、唯讀、格式、密碼、寫入密碼及忽略唯讀建議。
    workbook = excel.Workbooks.Open(str(source), 0, True, Missing, "", "", True)
    try:
        if extension == "pdf":
            workbook.ExportAsFixedFormat(0, str(target))
        else:
            file_format = {"xlsx": 51, "xls": 56}[extension]
            workbook.SaveAs(str(target), file_format)
    finally:
        workbook.Close(False)
    if not target.is_file() or not target.stat().st_size:
        raise RuntimeError(f"Excel 未產生 {extension} 檔案")
    return target


@contextmanager
def open_converter(profile):
    if sys.platform != "win32":
        office = find_office()
        if not office:
            raise RuntimeError("找不到 LibreOffice。請從 https://www.libreoffice.org/ 安裝後重新執行。")
        yield partial(convert_libreoffice, office, profile)
        return

    from pywintypes import com_error
    from win32com.client import DispatchEx

    try:
        excel = DispatchEx("Excel.Application")
    except com_error as error:
        raise RuntimeError(
            "無法啟動桌面版 Microsoft Excel，請確認已安裝且能正常開啟。"
        ) from error
    try:
        try:
            excel.DisplayAlerts = False
            excel.EnableEvents = False
            excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            yield partial(convert_excel, excel)
        finally:
            excel.Quit()
    except com_error as error:
        raise RuntimeError(f"無法完成 Microsoft Excel 自動化作業。{error}") from error


def replace_excel(source, target, old, new):
    workbook = load_workbook(source, keep_vba=source.suffix.lower() == ".xlsm",
                             rich_text=True)
    count = 0
    try:
        for sheet in workbook.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.data_type != "s" or not isinstance(cell.value, (str, CellRichText)):
                        continue
                    value = str(cell.value)
                    occurrences = value.count(old)
                    if occurrences:
                        changed = value.replace(old, new)
                        if len(changed) > 32767:
                            raise ValueError(f"{sheet.title}!{cell.coordinate} 超過 Excel 的字數限制")
                        cell.value = changed
                        cell.data_type = "s"
                        count += occurrences
        if count:
            workbook.save(target)
        else:
            shutil.copyfile(source, target)
    finally:
        workbook.close()
    return count


def decode_csv(data):
    for marker, encoding in (
        (codecs.BOM_UTF8, "utf-8"),
        (codecs.BOM_UTF32_LE, "utf-32-le"),
        (codecs.BOM_UTF32_BE, "utf-32-be"),
        (codecs.BOM_UTF16_LE, "utf-16-le"),
        (codecs.BOM_UTF16_BE, "utf-16-be"),
    ):
        if data.startswith(marker):
            return data[len(marker):].decode(encoding), marker, encoding
    for encoding in ("utf-8", "cp950", "gb18030"):
        try:
            return data.decode(encoding), b"", encoding
        except UnicodeDecodeError:
            continue
    raise ValueError("無法辨識 CSV 編碼")


def replace_csv(source, target, pdf_source, old, new):
    data = source.read_bytes()
    content, bom, encoding = decode_csv(data)
    try:
        dialect = csv.Sniffer().sniff(content[:65536], delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "CSV"
    sheet.sheet_view.showGridLines = False
    sheet.print_options.gridLines = True
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0

    output = io.StringIO(newline="")
    newline = "\r\n" if "\r\n" in content else "\n"
    writer = csv.writer(output, dialect, lineterminator=newline, doublequote=True,
                        skipinitialspace=False, quoting=csv.QUOTE_ALL)
    widths = {}
    count = 0
    reader = csv.reader(io.StringIO(content, newline=""), dialect,
                        doublequote=True, skipinitialspace=False, strict=True)
    for row_index, row in enumerate(reader, 1):
        changed = [field.replace(old, new) for field in row]
        count += sum(field.count(old) for field in row)
        writer.writerow(changed)
        for column, value in enumerate(changed, 1):
            if len(value) > 32767:
                raise ValueError(f"CSV 第 {row_index} 列第 {column} 欄超過 PDF 匯出的字數限制")
            cell = sheet.cell(row_index, column, value)
            cell.data_type = "s"
            widths[column] = max(widths.get(column, 0), len(value))
    for column, width in widths.items():
        sheet.column_dimensions[get_column_letter(column)].width = min(255, max(12, 2 * width + 2))
    if not widths:
        sheet["A1"] = " "
    workbook.save(pdf_source)
    target.write_bytes(bom + output.getvalue().encode(encoding) if count else data)
    return count


def process(source, convert, work, old, new):
    suffix = source.suffix.lower()
    staged = work / f"updated{suffix}"
    pdf_source = staged
    if suffix == ".csv":
        pdf_source = work / "csv.xlsx"
        count = replace_csv(source, staged, pdf_source, old, new)
    elif suffix == ".xls":
        modern = convert(source, work / "import", "xlsx")
        pdf_source = work / "updated.xlsx"
        count = replace_excel(modern, pdf_source, old, new)
        if count:
            staged = convert(pdf_source, work / "legacy", "xls")
            pdf_source = staged
        else:
            pdf_source = source
    else:
        count = replace_excel(source, staged, old, new)

    pdf = convert(pdf_source, work / "pdf", "pdf")
    with pdf.open("rb") as stream:
        if stream.read(5) != b"%PDF-":
            raise ValueError("輸出檔不是有效的 PDF")
    if count:
        shutil.copymode(source, staged)
        os.replace(staged, source)
    pdf_path = source.with_name(f"{source.name}.pdf")
    os.replace(pdf, pdf_path)
    return count, pdf_path


def parse_directory(value):
    path = Path(value.strip()).expanduser()
    if path.is_dir():
        return path.resolve()
    # 支援把 Finder 資料夾拖進終端機後產生的引號或跳脫字元。
    try:
        parts = shlex.split(value)
    except ValueError:
        parts = []
    if len(parts) == 1:
        path = Path(parts[0]).expanduser()
        if path.is_dir():
            return path.resolve()
    raise ValueError("找不到此資料夾")


def main():
    raw_path = input("資料夾路徑：")
    if not raw_path.strip():
        raise ValueError("資料夾路徑不可留空")
    directory = parse_directory(raw_path)
    old = input("input（要搜尋的文字）：")
    if not old:
        raise ValueError("input 不可留空")
    new = input("output（替換文字，可留空）：")

    sources = sorted(path for path in directory.iterdir()
                     if path.suffix.lower() in SUPPORTED and path.is_file()
                     and not path.is_symlink() and not path.name.startswith("~$"))
    if not sources:
        print("資料夾內沒有支援的檔案。")
        return 0
    failures = 0
    with TemporaryDirectory(prefix="replace-pdf-", dir=directory) as temporary:
        root = Path(temporary)
        with open_converter(root / "profile") as convert:
            for index, source in enumerate(sources):
                work = root / str(index)
                work.mkdir()
                try:
                    count, pdf = process(source, convert, work, old, new)
                    print(f"完成 {source.name}，替換 {count} 處，PDF 為 {pdf.name}")
                except Exception as error:
                    failures += 1
                    print(f"失敗 {source.name}，{error}", file=sys.stderr)
    print(f"共 {len(sources)} 個檔案，成功 {len(sources) - failures} 個，失敗 {failures} 個。")
    return int(bool(failures))


if __name__ == "__main__":
    frozen_windows = sys.platform == "win32" and getattr(sys, "frozen", False)
    if frozen_windows:
        for stream in (sys.stdin, sys.stdout, sys.stderr):
            if not stream.isatty():
                stream.reconfigure(encoding="utf-8")
    try:
        exit_code = main()
    except (EOFError, KeyboardInterrupt, OSError, RuntimeError, ValueError) as error:
        print(f"執行中止。{error}", file=sys.stderr)
        exit_code = 1
    if frozen_windows and sys.stdin.isatty():
        try:
            input("按 Enter 關閉視窗。")
        except (EOFError, KeyboardInterrupt):
            pass
    sys.exit(exit_code)
