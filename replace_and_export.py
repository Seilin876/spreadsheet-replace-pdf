import codecs
import csv
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


def convert(office, profile, source, target_dir, output_filter):
    target_dir.mkdir(exist_ok=True)
    result = subprocess.run(
        [office, f"-env:UserInstallation={profile.as_uri()}", "--headless",
         "--convert-to", output_filter, "--outdir", str(target_dir), str(source)],
        capture_output=True, text=True, timeout=180,
    )
    suffix = output_filter.split(":", 1)[0]
    target = target_dir / f"{source.stem}.{suffix}"
    if result.returncode or not target.is_file() or not target.stat().st_size:
        detail = (result.stderr or result.stdout).strip()
        raise RuntimeError(f"無法轉成 {suffix}。{detail}")
    return target


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


def process(source, office, profile, work, old, new):
    suffix = source.suffix.lower()
    staged = work / f"updated{suffix}"
    pdf_source = staged
    if suffix == ".csv":
        pdf_source = work / "csv.xlsx"
        count = replace_csv(source, staged, pdf_source, old, new)
    elif suffix == ".xls":
        modern = convert(office, profile, source, work / "import", "xlsx:Calc Office Open XML")
        pdf_source = work / "updated.xlsx"
        count = replace_excel(modern, pdf_source, old, new)
        if count:
            staged = convert(office, profile, pdf_source, work / "legacy", "xls:MS Excel 97")
            pdf_source = staged
        else:
            pdf_source = source
    else:
        count = replace_excel(source, staged, old, new)

    pdf = convert(office, profile, pdf_source, work / "pdf", "pdf:calc_pdf_Export")
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
    office = shutil.which("soffice") or "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if not os.access(office, os.X_OK):
        raise RuntimeError("找不到 LibreOffice，請先安裝後再執行。")

    failures = 0
    with TemporaryDirectory(prefix="replace-pdf-", dir=directory) as temporary:
        root = Path(temporary)
        for index, source in enumerate(sources):
            work = root / str(index)
            work.mkdir()
            try:
                count, pdf = process(source, office, root / "profile", work, old, new)
                print(f"完成 {source.name}，替換 {count} 處，PDF 為 {pdf.name}")
            except Exception as error:
                failures += 1
                print(f"失敗 {source.name}，{error}", file=sys.stderr)
    print(f"共 {len(sources)} 個檔案，成功 {len(sources) - failures} 個，失敗 {failures} 個。")
    return int(bool(failures))


if __name__ == "__main__":
    try:
        sys.exit(main())
    except (EOFError, KeyboardInterrupt, OSError, RuntimeError, ValueError) as error:
        print(f"執行中止。{error}", file=sys.stderr)
        sys.exit(1)
