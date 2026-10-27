"""Unduh CSV verifikasi pendidikan per sekolah."""

import csv
import html
import json
import re
import sys
import tempfile
import time
import traceback
from datetime import datetime
from pathlib import Path

import openpyxl
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch, mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    Image as ReportLabImage,
    LongTable,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

SCRIPT_FILE = Path(__file__).resolve()
SCRIPT_DIR = SCRIPT_FILE.parent
ROOT_DIR = SCRIPT_DIR.parent
IO_DIR = SCRIPT_DIR
RESULT_DIR = ROOT_DIR / "_result_scraping" / f"result_{ROOT_DIR.name}" / SCRIPT_DIR.name
RESULT_DIR.mkdir(parents=True, exist_ok=True)
SCRIPT_NAME = SCRIPT_FILE.stem
RESULT_NAME = SCRIPT_NAME.replace("src_", "result_", 1)
STATUS_FILE = IO_DIR / f"status_{SCRIPT_NAME}.json"
COOKIE_FILE = (
    ROOT_DIR / "_helpers" / "simpan_login_siks" / "cookie_siks.json"
)
FOLDER_DATA_CSV_FASDIK = RESULT_DIR
FOLDER_DATA_FASDIK_PDF = RESULT_DIR / "data fasdik pdf"
LOGO_KEMENSOS = SCRIPT_DIR / "Logo" / "logo kemensos.jpg"
LOGO_PKH = SCRIPT_DIR / "Logo" / "logo pkh.jpg"
FOLIO_PAGE_SIZE = landscape((8.5 * inch, 13 * inch))
SIKS_URL = "https://siks.kemensos.go.id/"

if sys.platform == "win32":
    import msvcrt

    def flush_input():
        while msvcrt.kbhit():
            msvcrt.getch()
else:
    def flush_input():
        pass


def bersihkan_teks_data(value):
    if value is None:
        return ""
    value = str(value).replace("\u00a0", " ").replace("\r", " ").replace("\n", " ")
    return re.sub(r"\s+", " ", value).strip()


def kunci_header(value):
    return re.sub(r"[\s_]+", "", bersihkan_teks_data(value)).casefold()


def npsn_teks(value):
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return bersihkan_teks_data(value)


def safe_filename_part(value):
    value = bersihkan_teks_data(value)
    value = re.sub(r'[<>:"/\\|?*]', "_", value)
    value = re.sub(r"\s+", " ", value).strip(" .")
    return value or "Tidak_diketahui"


def timestamp_file():
    return datetime.now().strftime("%Y%m%d_%H%M%S_%f")


def cari_file_referensi():
    candidates = sorted(
        path
        for path in IO_DIR.iterdir()
        if path.name.casefold().startswith("ref_")
        and path.suffix.casefold() in {".xlsx", ".xlsm"}
        and path.is_file()
    )
    if not candidates:
        raise FileNotFoundError(
            f"Tidak ada file referensi berawalan 'ref_' (.xlsx/.xlsm) "
            f"di {IO_DIR}."
        )
    if len(candidates) > 1:
        raise ValueError(
            "Ditemukan lebih dari satu file referensi; sisakan satu file yang "
            "akan digunakan:\n - " + "\n - ".join(str(path) for path in candidates)
        )
    return candidates[0]


def _load_status():
    if not STATUS_FILE.is_file():
        return {"version": 1, "schools": {}}

    with STATUS_FILE.open(encoding="utf-8") as status_file:
        status = json.load(status_file)
    if (
        not isinstance(status, dict)
        or status.get("version") != 1
        or not isinstance(status.get("schools"), dict)
        or any(
            not isinstance(record, dict)
            for record in status.get("schools", {}).values()
        )
    ):
        raise ValueError(f"Format status tidak valid: {STATUS_FILE}")
    return status


def _save_status(status):
    status["updated_at"] = time.strftime("%Y-%m-%dT%H:%M:%S%z")
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            "w",
            encoding="utf-8",
            dir=IO_DIR,
            prefix=".status_",
            suffix=".tmp",
            delete=False,
        ) as temp_file:
            temp_path = Path(temp_file.name)
            json.dump(status, temp_file, ensure_ascii=False, indent=2)
            temp_file.write("\n")
        temp_path.replace(STATUS_FILE)
    finally:
        if temp_path and temp_path.exists():
            temp_path.unlink()


def _mark_status(status, npsn, request_type, result, output_path=None):
    record = status["schools"].setdefault(npsn, {})
    record[request_type] = result
    if output_path:
        record[f"{request_type}_path"] = (
            Path(output_path).resolve().relative_to(ROOT_DIR).as_posix()
        )
    _save_status(status)


def muat_daftar_sekolah_untuk_diunduh():
    reference_path = cari_file_referensi()
    workbook = openpyxl.load_workbook(
        reference_path,
        read_only=True,
        data_only=True,
    )
    if "Semua Sekolah" not in workbook.sheetnames:
        workbook.close()
        raise ValueError(
            f"Sheet 'Semua Sekolah' tidak ditemukan di {reference_path}."
        )
    worksheet = workbook["Semua Sekolah"]
    rows = worksheet.iter_rows(values_only=True)
    header_row = next(rows, None)
    if header_row is None:
        workbook.close()
        raise ValueError(f"Sheet 'Semua Sekolah' kosong: {reference_path}")
    column_names = {
        kunci_header(value): index
        for index, value in enumerate(header_row)
        if value is not None
    }
    required = (
        "npsn",
        "namasatuanpendidikan",
        "desa/kelurahan",
        "kecamatan",
        "kabupaten",
        "provinsi",
        "request",
    )
    missing = [name for name in required if name not in column_names]
    if missing:
        workbook.close()
        raise ValueError(
            f"Kolom wajib tidak ditemukan di {reference_path}: "
            f"{', '.join(missing)}."
        )
    source_rows = list(rows)

    try:
        schools = []
        seen_npsn = set()
        for row_number, row in enumerate(source_rows, start=2):
            request = bersihkan_teks_data(row[column_names["request"]]).casefold()
            if request != "download":
                continue

            npsn_value = row[column_names["npsn"]]
            name_value = row[column_names["namasatuanpendidikan"]]
            village_value = row[column_names["desa/kelurahan"]]
            npsn = npsn_teks(npsn_value)
            name = bersihkan_teks_data(name_value)
            village = bersihkan_teks_data(village_value)
            address = ", ".join(
                part
                for part in (
                    bersihkan_teks_data(row[column_names[key]])
                    for key in (
                        "desa/kelurahan",
                        "kecamatan",
                        "kabupaten",
                        "provinsi",
                    )
                )
                if part
            )
            if not npsn or not name or not village:
                raise ValueError(
                    f"Data NPSN/nama/desa kosong pada baris {row_number} "
                    f"di {reference_path}."
                )
            if npsn in seen_npsn:
                raise ValueError(
                    f"NPSN duplikat {npsn} pada daftar permintaan download."
                )
            seen_npsn.add(npsn)

            schools.append({
                "NPSN": npsn,
                "Nama Sekolah": name,
                "Desa": village,
                "Alamat": address,
                "Request": "download",
            })
    finally:
        workbook.close()

    return schools


def tandai_status_sukses(reference_path, npsn):
    workbook = openpyxl.load_workbook(reference_path)
    try:
        worksheet = workbook["Semua Sekolah"]
        headers = {
            kunci_header(worksheet.cell(row=1, column=column).value): column
            for column in range(1, worksheet.max_column + 1)
            if worksheet.cell(row=1, column=column).value is not None
        }
        npsn_column = headers.get("npsn")
        if npsn_column is None:
            raise ValueError(f"Kolom NPSN tidak ditemukan di {reference_path}.")
        status_column = headers.get(kunci_header("Status_proses"))
        if status_column is None:
            status_column = worksheet.max_column + 1
            worksheet.cell(row=1, column=status_column, value="Status_proses")

        matched = False
        for row_number in range(2, worksheet.max_row + 1):
            row_npsn = npsn_teks(
                worksheet.cell(row=row_number, column=npsn_column).value
            )
            if row_npsn == npsn:
                worksheet.cell(row=row_number, column=status_column, value="sukses")
                matched = True
        if not matched:
            raise LookupError(f"NPSN {npsn} tidak ditemukan di {reference_path}.")
        workbook.save(reference_path)
    finally:
        workbook.close()


def csv_valid_untuk_sekolah(csv_path, npsn):
    csv_path = Path(csv_path)
    if not csv_path.is_file() or csv_path.stat().st_size == 0:
        return False
    with csv_path.open("r", encoding="utf-8-sig") as source:
        contents = source.read()
    npsn_header = re.search(
        rf"(?im)^NPSN\s*;\s*:?\s*['\"]?{re.escape(npsn)}(?:['\"]?\s*(?:;|$))",
        contents,
    )
    return (
        "FORM VERIFIKASI KOMITMEN PENDIDIKAN" in contents.upper()
        and npsn_header is not None
    )


def triwulan_dari_csv(csv_path):
    nama_bulan = {
        "januari": 1, "jan": 1, "january": 1,
        "februari": 2, "feb": 2, "february": 2,
        "maret": 3, "mar": 3, "march": 3,
        "april": 4, "apr": 4,
        "mei": 5, "may": 5,
        "juni": 6, "jun": 6, "june": 6,
        "juli": 7, "jul": 7, "july": 7,
        "agustus": 8, "agu": 8, "ags": 8, "aug": 8, "august": 8,
        "september": 9, "sep": 9,
        "oktober": 10, "okt": 10, "oct": 10, "october": 10,
        "november": 11, "nov": 11,
        "desember": 12, "des": 12, "dec": 12, "december": 12,
    }

    def bulan_dari_teks(value):
        value = bersihkan_teks_data(value).casefold()
        for nama, nomor in nama_bulan.items():
            if re.search(rf"(?<![a-z]){re.escape(nama)}(?![a-z])", value):
                return nomor

        tanggal_lengkap = re.search(
            r"(?<!\d)(\d{1,2})[./-](\d{1,2})[./-]\d{2,4}(?!\d)",
            value,
        )
        if tanggal_lengkap and 1 <= int(tanggal_lengkap.group(2)) <= 12:
            return int(tanggal_lengkap.group(2))

        tahun_bulan = re.search(
            r"(?<!\d)\d{4}[./-](\d{1,2})(?:[./-]\d{1,2})?(?!\d)",
            value,
        )
        if tahun_bulan and 1 <= int(tahun_bulan.group(1)) <= 12:
            return int(tahun_bulan.group(1))

        bulan_tahun = re.search(
            r"(?<!\d)(\d{1,2})[./-]\d{4}(?!\d)",
            value,
        )
        if bulan_tahun and 1 <= int(bulan_tahun.group(1)) <= 12:
            return int(bulan_tahun.group(1))

        angka = re.fullmatch(r"(?:bulan\s*)?(\d{1,2})", value)
        if angka and 1 <= int(angka.group(1)) <= 12:
            return int(angka.group(1))
        return None

    with open(csv_path, "r", encoding="utf-8-sig", newline="") as source:
        rows = list(csv.reader(source, delimiter=";"))
    rows = [row for row in rows if any(cell.strip() for cell in row)]
    if rows and rows[0][0].strip().lower().startswith("sep="):
        rows.pop(0)
    header_index = next(
        (
            index
            for index, row in enumerate(rows)
            if len(row) > 1
            and row[0].strip().casefold() == "no"
            and row[1].strip().casefold() == "nik pengurus"
        ),
        None,
    )
    if header_index is None or header_index + 1 >= len(rows):
        raise ValueError(f"Header bulan tidak ditemukan pada CSV: {csv_path}")

    month_numbers = {
        month
        for column in (8, 13, 18)
        if column < len(rows[header_index])
        for month in [bulan_dari_teks(rows[header_index][column])]
        if month is not None
    }
    if not month_numbers:
        raise ValueError(f"Bulan tidak ditemukan pada header CSV: {csv_path}")
    quarters = {(month - 1) // 3 + 1 for month in month_numbers}
    if len(quarters) != 1:
        raise ValueError(
            f"Bulan pada header CSV berasal dari triwulan berbeda: {csv_path}"
        )
    return quarters.pop()


def cari_dan_unduh_csv_sekolah(page, school, output_dir):
    npsn = school["NPSN"]
    output_dir = Path(output_dir)
    village_dirname = safe_filename_part(school["Desa"])
    filename_stem = (
        f"{safe_filename_part(npsn)}_"
        f"{safe_filename_part(school['Nama Sekolah'])}"
    )
    existing_files = sorted(
        output_dir.glob(
            f"TRIWULAN */{village_dirname}/{filename_stem}_Triwulan *.csv"
        ),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    for existing_path in existing_files:
        if not csv_valid_untuk_sekolah(existing_path, npsn):
            continue
        try:
            quarter = triwulan_dari_csv(existing_path)
        except ValueError:
            continue
        if existing_path.name == f"{filename_stem}_Triwulan {quarter}.csv":
            return existing_path

    tombol_reset = page.get_by_role("button", name="Reset", exact=True)
    if tombol_reset.count():
        tombol_reset.click()

    input_sekolah = page.get_by_role(
        "combobox",
        name="Cari nama sekolah atau NPSN...",
    )
    input_sekolah.fill(npsn)
    tombol_cari_sekolah = page.get_by_role(
        "button",
        name="Cari Sekolah",
        exact=True,
    )
    if tombol_cari_sekolah.count() != 1:
        raise RuntimeError("Tombol 'Cari Sekolah' tidak ditemukan secara tunggal.")
    tombol_cari_sekolah.click()

    pilihan_sekolah = page.get_by_role("option").filter(
        has_text=re.compile(rf"(?<!\d){re.escape(npsn)}(?!\d)")
    )
    try:
        pilihan_sekolah.wait_for(state="visible", timeout=15000)
    except PlaywrightTimeoutError as error:
        raise LookupError(
            f"Sekolah NPSN {npsn} tidak muncul di hasil pencarian SIKS."
        ) from error
    if pilihan_sekolah.count() != 1:
        raise RuntimeError(
            f"Hasil pencarian NPSN {npsn} tidak tunggal; tidak memilih otomatis."
        )
    pilihan_sekolah.click()

    if npsn not in input_sekolah.input_value():
        raise RuntimeError(
            f"Sekolah terpilih tidak cocok dengan NPSN yang diminta: {npsn}."
        )

    tombol_cari_siswa = page.get_by_role(
        "button",
        name="Cari Data Siswa",
        exact=True,
    )
    tombol_cari_siswa.wait_for(state="visible", timeout=10000)
    tombol_cari_siswa.click()

    ringkasan_hasil = page.get_by_text(
        re.compile(r"Ditemukan\s+\d+\s+data siswa\.", re.IGNORECASE)
    )
    try:
        ringkasan_hasil.wait_for(state="visible", timeout=30000)
    except PlaywrightTimeoutError as error:
        raise RuntimeError(
            f"Hasil daftar siswa untuk NPSN {npsn} tidak selesai dimuat."
        ) from error

    tombol_download = page.get_by_role(
        "button",
        name=re.compile(r"Download CSV", re.IGNORECASE),
    )
    tombol_download.wait_for(state="visible", timeout=10000)
    if tombol_download.is_disabled():
        raise RuntimeError(f"Tombol Download CSV nonaktif untuk NPSN {npsn}.")

    with page.expect_download(timeout=120000) as download_info:
        tombol_download.click()
    download = download_info.value
    with tempfile.NamedTemporaryFile(
        dir=output_dir,
        prefix=".unduhan_",
        suffix=".tmp",
        delete=False,
    ) as temporary_file:
        temporary_path = Path(temporary_file.name)
    try:
        download.save_as(temporary_path)
        if download.failure():
            raise RuntimeError(
                f"Browser gagal mengunduh CSV NPSN {npsn}: {download.failure()}"
            )
        if not csv_valid_untuk_sekolah(temporary_path, npsn):
            raise RuntimeError(
                f"CSV yang diunduh tidak memuat identitas NPSN {npsn}."
            )
        quarter = triwulan_dari_csv(temporary_path)
        output_path = (
            output_dir
            / f"TRIWULAN {quarter}"
            / village_dirname
            / f"{filename_stem}_Triwulan {quarter}.csv"
        )
        output_path.parent.mkdir(parents=True, exist_ok=True)
        temporary_path.replace(output_path)
        return output_path
    finally:
        if temporary_path.exists():
            temporary_path.unlink()


class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total_pages = len(self._saved_page_states)
        page_width, _ = FOLIO_PAGE_SIZE
        for page_number, page_state in enumerate(self._saved_page_states, start=1):
            self.__dict__.update(page_state)
            self.saveState()
            self.setStrokeColor(colors.HexColor("#9CA3AF"))
            self.setLineWidth(0.5)
            self.line(12 * mm, 10 * mm, page_width - 12 * mm, 10 * mm)
            self.setFont("Helvetica", 8)
            self.drawString(12 * mm, 6 * mm, "FORM VERIFIKASI KOMITMEN PENDIDIKAN")
            self.drawRightString(
                page_width - 12 * mm,
                6 * mm,
                f"Halaman {page_number}/{total_pages}",
            )
            self.restoreState()
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)


def clean_value(value):
    value = (value or "").strip()
    if value.startswith("'"):
        value = value[1:]
    return value


def paragraph_text(value):
    return html.escape(clean_value(value), quote=False).replace("\n", "<br/>")


def read_form(csv_path):
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as source:
        rows = list(csv.reader(source, delimiter=";"))

    rows = [row for row in rows if any(cell.strip() for cell in row)]
    if rows and rows[0][0].strip().lower().startswith("sep="):
        rows.pop(0)
    if len(rows) < 4:
        raise ValueError(f"Format CSV tidak lengkap: {csv_path}")

    title = clean_value(rows[0][0])
    metadata = {}
    for row in rows[1:]:
        if row and row[0].strip().casefold() in {"npsn", "nama sekolah"}:
            metadata[row[0].strip().casefold()] = clean_value(
                row[1].lstrip(":") if len(row) > 1 else ""
            )

    header_index = next(
        (
            index
            for index, row in enumerate(rows)
            if len(row) > 1
            and row[0].strip().casefold() == "no"
            and row[1].strip().casefold() == "nik pengurus"
        ),
        None,
    )
    if header_index is None or header_index + 3 >= len(rows):
        raise ValueError(f"Baris header kolom tidak ditemukan: {csv_path}")

    header_rows = [row[:] for row in rows[header_index:header_index + 3]]
    column_count = max(len(row) for row in header_rows)
    if column_count < 10:
        raise ValueError(f"Jumlah kolom CSV tidak sesuai: {csv_path}")
    for row in header_rows:
        row.extend([""] * (column_count - len(row)))

    student_rows = []
    for row in rows[header_index + 3:]:
        if row and row[0].strip().isdigit():
            row = row[:]
            row.extend([""] * (column_count - len(row)))
            student_rows.append(
                [value for index, value in enumerate(row[:column_count]) if index not in {6, 23}]
            )
    if not student_rows:
        raise ValueError(f"Tidak ada baris siswa pada CSV: {csv_path}")

    header_rows = [
        [value for index, value in enumerate(row) if index not in {6, 23}]
        for row in header_rows
    ]
    header_rows[0][6] = "KELAS"
    return title, metadata, header_rows, student_rows


def make_styles():
    styles = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "FormTitle",
            parent=styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=24,
            leading=28,
            textColor=colors.HexColor("#17365D"),
            alignment=1,
            spaceAfter=2,
        ),
        "meta": ParagraphStyle(
            "FormMetadata",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=8,
            leading=10,
        ),
        "header": ParagraphStyle(
            "FormHeader",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=8,
            leading=9,
            alignment=1,
            textColor=colors.white,
        ),
        "body": ParagraphStyle(
            "FormCell",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=8,
            leading=9,
            wordWrap="CJK",
        ),
        "signature": ParagraphStyle(
            "Signature",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=8,
            leading=10,
            alignment=1,
        ),
    }


def make_table(header_rows, student_rows, available_width, styles):
    column_count = len(header_rows[0])
    month_starts = (7, 12, 17)
    for month_index in month_starts:
        header_rows[0][month_index] = "Bulan:"
    for start_column in month_starts:
        for offset, abbreviation in enumerate(("A", "I", "S", "J")):
            header_rows[2][start_column + offset] = abbreviation

    fixed_widths = [24, 86, 72, 86, 50, 90, 50]
    fixed_widths.extend([26, 26, 28, 26, 33] * 3)
    fixed_widths.append(90)
    if len(fixed_widths) < column_count:
        fixed_widths.extend([40] * (column_count - len(fixed_widths)))
    fixed_widths = fixed_widths[:column_count]
    width_scale = available_width / sum(fixed_widths)
    column_widths = [width * width_scale for width in fixed_widths]

    def fit_style(base_style, text, width, font_name):
        available_cell_width = max(1, width - 4)
        text_width = pdfmetrics.stringWidth(text, font_name, base_style.fontSize)
        font_size = min(
            base_style.fontSize,
            max(5.5, base_style.fontSize * available_cell_width / max(text_width, 1)),
        )
        return ParagraphStyle(
            f"{base_style.name}_{font_size:.2f}",
            parent=base_style,
            fontSize=font_size,
            leading=font_size * 1.15,
        )

    data = []
    for row_index, row in enumerate(header_rows):
        cells = []
        for column, value in enumerate(row):
            if not value.strip():
                cells.append("")
                continue
            if row_index == 0 and column in month_starts:
                display_text = (
                    f"{paragraph_text(value)} "
                    "<font color='#B8B8B8'>.............................</font>"
                )
                measured_text = f"{value} ............................."
            else:
                display_text = paragraph_text(value)
                measured_text = value
            cell_style = fit_style(
                styles["header"],
                measured_text,
                column_widths[column],
                "Helvetica-Bold",
            )
            cells.append(Paragraph(display_text, cell_style))
        data.append(cells)
    for row in student_rows:
        grade = clean_value(row[6])
        if grade.isdigit():
            row[6] = f"Kelas {grade}"
        cells = []
        for column, value in enumerate(row):
            if not value.strip():
                cells.append("")
                continue
            cell_style = fit_style(
                styles["body"],
                value,
                column_widths[column],
                "Helvetica",
            )
            cells.append(Paragraph(paragraph_text(value), cell_style))
        data.append(cells)

    table = LongTable(
        data,
        colWidths=column_widths,
        repeatRows=3,
        hAlign="LEFT",
        splitByRow=1,
    )
    table_styles = [
        ("BACKGROUND", (0, 0), (-1, 2), colors.HexColor("#1F4E78")),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#7F8C8D")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, 2), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("BACKGROUND", (0, 3), (-1, -1), colors.white),
    ]
    for column in list(range(7)) + [22]:
        if column < column_count:
            table_styles.append(("SPAN", (column, 0), (column, 2)))
    for first_column in month_starts:
        if first_column + 4 < column_count:
            table_styles.append(
                ("SPAN", (first_column, 0), (first_column + 4, 0))
            )
            table_styles.append(
                ("SPAN", (first_column, 1), (first_column + 4, 1))
            )
            table_styles.append(("ALIGN", (first_column, 0), (first_column + 4, 0), "LEFT"))
    for row_index in range(3, len(data)):
        if (row_index - 3) % 2:
            table_styles.append(
                ("BACKGROUND", (0, row_index), (-1, row_index),
                 colors.HexColor("#EFF5FA"))
            )
    table.setStyle(TableStyle(table_styles))
    return table


def konversi_csv_ke_pdf(csv_path, output_path, school_address, npsn):
    title, metadata, header_rows, student_rows = read_form(csv_path)
    if metadata.get("npsn", "") != npsn:
        raise ValueError(
            f"NPSN di CSV ({metadata.get('npsn') or 'kosong'}) tidak sama "
            f"dengan NPSN yang diminta ({npsn}): {csv_path}"
        )
    for logo in (LOGO_KEMENSOS, LOGO_PKH):
        if not logo.is_file():
            raise FileNotFoundError(f"File logo tidak ditemukan: {logo}")

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    styles = make_styles()
    page_width, _ = FOLIO_PAGE_SIZE
    margin = 12 * mm
    document = SimpleDocTemplate(
        str(output_path),
        pagesize=FOLIO_PAGE_SIZE,
        leftMargin=margin,
        rightMargin=margin,
        topMargin=12 * mm,
        bottomMargin=22 * mm,
        title=f"{title} - {metadata.get('nama sekolah', '')}",
        author="Scraping Data Fasdik",
    )

    available_width = page_width - 2 * margin
    address_value_width = available_width - 175 * mm
    metadata_table = Table(
        [[
            Paragraph("<b>NPSN</b>", styles["meta"]),
            Paragraph(":", styles["meta"]),
            Paragraph(paragraph_text(metadata.get("npsn", "")), styles["meta"]),
            Paragraph("<b>Nama Sekolah</b>", styles["meta"]),
            Paragraph(":", styles["meta"]),
            Paragraph(
                paragraph_text(metadata.get("nama sekolah", "")),
                styles["meta"],
            ),
            Paragraph("<b>Alamat</b>", styles["meta"]),
            Paragraph(":", styles["meta"]),
            Paragraph(paragraph_text(school_address), styles["meta"]),
        ]],
        colWidths=[
            14 * mm,
            4 * mm,
            35 * mm,
            25 * mm,
            4 * mm,
            85 * mm,
            16 * mm,
            4 * mm,
            address_value_width,
        ],
        hAlign="LEFT",
    )
    metadata_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))

    kemensos_logo = ReportLabImage(
        str(LOGO_KEMENSOS), width=30 * mm, height=22 * mm, kind="proportional"
    )
    pkh_logo = ReportLabImage(
        str(LOGO_PKH), width=30 * mm, height=22 * mm, kind="proportional"
    )
    kemensos_logo.hAlign = "CENTER"
    pkh_logo.hAlign = "CENTER"
    title_lines = title.rsplit(" ", 1)
    title_text = (
        f"{paragraph_text(title_lines[0])}<br/>{paragraph_text(title_lines[1])}"
        if len(title_lines) == 2
        else paragraph_text(title)
    )
    title_table = Table(
        [[kemensos_logo, Paragraph(title_text, styles["title"]), pkh_logo]],
        colWidths=[35 * mm, available_width - 70 * mm, 35 * mm],
    )
    title_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))

    story = [
        title_table,
        Spacer(1, 20),
        metadata_table,
        Spacer(1, 2 * mm),
        make_table(header_rows, student_rows, available_width, styles),
        Spacer(1, 8 * mm),
        Table(
            [[
                Paragraph(
                    "Mengetahui,<br/>Kepala Sekolah<br/><br/><br/><br/>",
                    styles["signature"],
                )
            ]],
            colWidths=[70 * mm],
            hAlign="RIGHT",
            style=TableStyle([
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]),
        ),
    ]
    document.build(story, canvasmaker=NumberedCanvas)
    if not output_path.is_file() or output_path.stat().st_size == 0:
        raise RuntimeError(f"PDF gagal dibuat atau kosong: {output_path}")
    return len(student_rows)


def buat_pdf_sekolah(school, csv_path):
    quarter = triwulan_dari_csv(csv_path)
    pdf_path = (
        FOLDER_DATA_FASDIK_PDF
        / f"TRIWULAN {quarter}"
        / safe_filename_part(school["Desa"])
        / f"{Path(csv_path).stem}.pdf"
    )
    if pdf_path.is_file() and pdf_path.stat().st_size > 0:
        print(f"[PDF SUDAH ADA] {pdf_path}")
        return pdf_path
    jumlah = konversi_csv_ke_pdf(
        csv_path, pdf_path, school["Alamat"], school["NPSN"]
    )
    print(f"[PDF OK] NPSN {school['NPSN']} | {jumlah} siswa | {pdf_path}")
    return pdf_path


def run():
    status = _load_status()
    _save_status(status)
    timestamp = time.strftime("%Y-%m-%dT%H:%M:%S%z")
    result = {
        "script": SCRIPT_NAME,
        "started_at": timestamp,
        "reference_file": None,
        "requested": 0,
        "successes": [],
        "failures": [],
        "status": "success",
    }
    error = None

    try:
        result.update(_run_download(status))
    except Exception as caught_error:
        error = caught_error
        result["status"] = "failed"
        result["failures"].append({"error": str(caught_error)})
        traceback.print_exc()

    result["finished_at"] = time.strftime("%Y-%m-%dT%H:%M:%S%z")
    print(
        f"\n[FINALISASI] Proses selesai. Tekan ENTER untuk menyimpan hasil "
        f"ke {RESULT_DIR}."
    )
    try:
        input("[TUNGGU] Tekan ENTER untuk menyimpan hasil...")
    except EOFError as error:
        raise RuntimeError(
            "Konfirmasi ENTER tidak diterima; file result tidak dibuat."
        ) from error

    result_path = _save_result(result)
    print(f"[HASIL XLSX] {result_path}")
    if error is not None:
        raise error
    if result["failures"]:
        raise RuntimeError(
            f"{len(result['failures'])} permintaan gagal. "
            "Periksa detail pada file result."
        )


def _run_download(status):
    schools = muat_daftar_sekolah_untuk_diunduh()
    reference_path = cari_file_referensi()
    if not schools:
        print(
            f"Tidak ada sekolah dengan request download di {reference_path}."
        )
        return {
            "reference_file": reference_path.relative_to(SCRIPT_DIR).as_posix(),
            "requested": 0,
            "successes": [],
            "failures": [],
            "status": "success",
        }

    if not COOKIE_FILE.is_file():
        raise FileNotFoundError(f"File cookie SIKS tidak ditemukan: {COOKIE_FILE}")

    output_dir = FOLDER_DATA_CSV_FASDIK
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"[INFO] {len(schools)} sekolah diminta untuk diunduh.")
    print(f"[INFO] CSV disimpan di: {output_dir}")

    failures = []
    downloaded_schools = []
    csv_paths = {}
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False)
        try:
            context = browser.new_context(
                storage_state=str(COOKIE_FILE),
                accept_downloads=True,
            )
            page = context.new_page()
            page.set_default_navigation_timeout(60000)
            try:
                response = page.goto(
                    SIKS_URL,
                    wait_until="commit",
                    timeout=60000,
                )
            except PlaywrightTimeoutError as error:
                raise RuntimeError(
                    "Browser tidak menerima respons awal dari SIKS dalam 60 detik."
                ) from error

            if response is not None and response.status >= 400:
                raise RuntimeError(
                    f"Halaman SIKS merespons HTTP {response.status}: {page.url}"
                )

            print(
                "\n[PETUNJUK] Jika diminta, login lalu buka menu "
                "'Daftar Siswa Sekolah' di SIKS NG. Setelah halaman pencarian "
                "sekolah tampil, tekan ENTER."
            )
            flush_input()
            input("[TUNGGU] Tekan ENTER untuk mulai mengunduh CSV...")

            if page.get_by_role(
                "combobox",
                name="Cari nama sekolah atau NPSN...",
            ).count() != 1:
                raise RuntimeError(
                    "Kolom pencarian sekolah tidak ditemukan. Pastikan halaman "
                    "'Daftar Siswa Sekolah' sudah terbuka di SIKS NG."
                )

            for school in schools:
                try:
                    csv_path = cari_dan_unduh_csv_sekolah(page, school, output_dir)
                except Exception as error:
                    failures.append((school["NPSN"], str(error)))
                    _mark_status(status, school["NPSN"], "csv", "gagal")
                    print(
                        f"{school['NPSN']} | {school['Nama Sekolah']} | gagal"
                    )
                    continue

                csv_paths[school["NPSN"]] = csv_path
                downloaded_schools.append(school)
                print(
                    f"{school['NPSN']} | {school['Nama Sekolah']} | sukses"
                )
                try:
                    _mark_status(
                        status,
                        school["NPSN"],
                        "csv",
                        "sukses",
                        csv_path,
                    )
                except Exception as error:
                    failures.append((
                        school["NPSN"],
                        f"CSV sukses, tetapi status lokal gagal disimpan: {error}",
                    ))
                try:
                    tandai_status_sukses(reference_path, school["NPSN"])
                except Exception as error:
                    failures.append((
                        school["NPSN"],
                        f"CSV sukses, tetapi status sukses gagal ditulis ke ref: {error}",
                    ))
        finally:
            browser.close()

    successes = [
        {
            "NPSN": school["NPSN"],
            "Nama Sekolah": school["Nama Sekolah"],
            "Desa": school["Desa"],
            "CSV": csv_paths[school["NPSN"]].resolve().relative_to(ROOT_DIR).as_posix(),
            "Request": school["Request"],
        }
        for school in downloaded_schools
    ]
    result_failures = [
        {"NPSN": npsn, "error": message}
        for npsn, message in failures
    ]
    print(
        f"\n[SELESAI] {len(successes)} berhasil; "
        f"{len(result_failures)} gagal."
    )
    return {
        "reference_file": cari_file_referensi().relative_to(SCRIPT_DIR).as_posix(),
        "requested": len(schools),
        "successes": successes,
        "failures": result_failures,
        "status": "failed" if result_failures else "success",
    }


def _save_result(result):
    path = RESULT_DIR / f"{RESULT_NAME}_{timestamp_file()}.xlsx"
    workbook = openpyxl.Workbook()
    summary = workbook.active
    summary.title = "Ringkasan"
    for key in (
        "script", "started_at", "finished_at", "reference_file",
        "requested", "status",
    ):
        summary.append([key, result.get(key)])
    for title, header in (
            ("Berhasil", ["NPSN", "Nama Sekolah", "Desa", "CSV", "Request"]),
        ("Gagal", ["NPSN", "error"]),
    ):
        sheet = workbook.create_sheet(title)
        sheet.append(header)
        for row in result["successes" if title == "Berhasil" else "failures"]:
            sheet.append([row.get(column) for column in header])
    workbook.save(path)
    return path


if __name__ == "__main__":
    try:
        run()
    except Exception:
        print("\n[ERROR] Proses selesai dengan kegagalan.")
        traceback.print_exc()
