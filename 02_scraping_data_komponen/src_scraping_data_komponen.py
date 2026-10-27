import json
import re
import sys
import tempfile
import time
import traceback
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

SCRIPT_FILE = Path(__file__).resolve()
SCRIPT_DIR = SCRIPT_FILE.parent

COOKIE_FILE = (
    SCRIPT_DIR.parent / "_helpers" / "simpan_login_siks" / "cookie_siks.json"
)

if sys.platform == "win32":
    import msvcrt
    def flush_input():
        while msvcrt.kbhit():
            msvcrt.getch()
else:
    def flush_input():
        pass

# --- KONFIGURASI ---
SCRIPT_NAME = SCRIPT_FILE.stem.removeprefix("src_")

IO_DIR = SCRIPT_DIR
RESULT_DIR = SCRIPT_DIR.parent / "_result_scraping" / f"result_{SCRIPT_DIR.parent.name}" / SCRIPT_DIR.name
STATUS_FILE = SCRIPT_DIR / f"status_{SCRIPT_FILE.stem}.json"
status_proses = None

kolom_baru = [
    'NIK_Pengurus',
    'NOK_Pengurus',
    'Nama_Pengurus',
    'Desa',
    'Alamat',
    'NIK_Anggota',
    'Nama_Anggota',
    'Komponen',          # Diubah dari Kategori_Komponen[cite: 1]
    'Nama Sekolah',
    'Kelas',
    'Status',
    
]
list_data_final = []


def tulis_json_atomik(path, data):
    path = Path(path)
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            "w",
            encoding="utf-8",
            dir=path.parent,
            prefix=".json_tmp_",
            suffix=".tmp",
            delete=False,
        ) as temp_file:
            temp_path = Path(temp_file.name)
            json.dump(data, temp_file, ensure_ascii=False, indent=2)
            temp_file.write("\n")
        temp_path.replace(path)
    finally:
        if temp_path and temp_path.exists():
            temp_path.unlink()


def waktu_sekarang():
    return datetime.now().astimezone().isoformat(timespec="seconds")


def cari_file_acuan():
    file_acuan = sorted(
        path
        for path in IO_DIR.iterdir()
        if path.name.lower().startswith("ref_")
        and path.suffix.lower() in {".json", ".xlsx", ".xlsm"}
    )
    if not file_acuan:
        raise FileNotFoundError(
            f"File acuan JSON atau Excel dengan awalan 'ref_' "
            f"tidak ditemukan di {IO_DIR}."
        )
    if len(file_acuan) > 1:
        raise ValueError(
            "Ditemukan lebih dari satu file acuan JSON/Excel di folder skrip: "
            + ", ".join(path.name for path in file_acuan)
            + ". Sisakan satu file acuan yang akan digunakan."
        )
    return file_acuan[0]


def simpan_progres(
    status_data,
    processed_nik,
    nik_gagal_dicari,
    nik_belum_diproses,
):
    status_data.update(
        {
            "updated_at": waktu_sekarang(),
            "processed_nik": processed_nik,
            "result_rows": len(list_data_final),
            "failed_nik": [
                {"nik": nik, "reason": alasan}
                for nik, alasan in nik_gagal_dicari
            ],
            "unprocessed_nik": nik_belum_diproses,
        }
    )
    tulis_json_atomik(STATUS_FILE, status_data)


def simpan_hasil(rows):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    path = RESULT_DIR / f"result_{SCRIPT_NAME}_{timestamp}.xlsx"
    headers = list(dict.fromkeys(key for row in rows for key in row))
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(headers)
    for row in rows:
        sheet.append([row.get(header) for header in headers])
    workbook.save(path)
    return str(path)

def kunci_baris_tabel(baris):
    if not baris.count():
        return None

    try:
        kunci_stabil = (
            baris.get_attribute("data-id", timeout=1000)
            or baris.get_attribute("data-rowindex", timeout=1000)
            or baris.get_attribute("aria-rowindex", timeout=1000)
        )
        if kunci_stabil:
            return kunci_stabil

        teks_sel = [
            bersihkan_teks_data(teks)
            for teks in baris.locator(
                ".MuiDataGrid-cell, td, [role='cell']"
            ).all_inner_texts()
        ]
        fingerprint = " | ".join(teks for teks in teks_sel if teks)
        if fingerprint:
            return fingerprint
    except PlaywrightTimeoutError:
        return None

    return None

def kunci_data_hasil(data_baris):
    return tuple(
        bersihkan_teks_data(str(data_baris[kolom])).casefold()
        for kolom in kolom_baru
        if kolom != "Status"
    )

def iter_baris_halaman(page, baris_ditemukan, status_scroll):
    baris_selector = ".MuiDataGrid-row, tr:has(button)"
    baris_datagrid = page.locator(".MuiDataGrid-row").first
    grid_virtual = baris_datagrid.count() > 0
    baris_selesai = status_scroll["baris_selesai"]
    jumlah_diharapkan = status_scroll["jumlah_diharapkan"]

    if grid_virtual:
        wadah_scroll = baris_datagrid.locator(
            "xpath=ancestor::*[contains("
            "concat(' ', normalize-space(@class), ' '), "
            "' MuiDataGrid-virtualScroller ')][1]"
        )
        if wadah_scroll.count() == 0:
            raise RuntimeError(
                "Baris DataGrid ditemukan tetapi wadah scroll tabelnya tidak ditemukan."
            )

    percobaan_tanpa_baris_baru = 0
    while True:
        if jumlah_diharapkan and len(baris_selesai) >= jumlah_diharapkan:
            status_scroll["mentok"] = True
            status_scroll["baris_terakhir_diproses"] = True
            print(
                f"[INFO v5.32] Semua {jumlah_diharapkan} baris pada halaman "
                "sudah ditangani (sukses, gagal, atau dilewati)."
            )
            return

        baris_tabel = page.locator(baris_selector)
        jumlah_baris = baris_tabel.count()
        if not jumlah_baris:
            status_scroll["kesalahan"] = (
                "Tidak ada baris tabel yang terlihat saat memindai halaman."
            )
            return

        jumlah_kunci_sebelum = len(baris_ditemukan)
        for indeks_baris in range(jumlah_baris):
            baris_tabel = page.locator(baris_selector)
            if indeks_baris >= baris_tabel.count():
                break
            baris = baris_tabel.nth(indeks_baris)
            row_key = kunci_baris_tabel(baris)
            if not row_key or row_key in baris_ditemukan:
                continue

            baris_ditemukan.add(row_key)
            status_scroll["baris_terakhir"] = row_key
            yield baris

        if len(baris_ditemukan) > jumlah_kunci_sebelum:
            percobaan_tanpa_baris_baru = 0

        if jumlah_diharapkan and len(baris_selesai) >= jumlah_diharapkan:
            status_scroll["mentok"] = True
            status_scroll["baris_terakhir_diproses"] = True
            return

        posisi_scroll = None
        if grid_virtual:
            posisi_scroll = wadah_scroll.evaluate(
                """el => {
                    const rows = [...el.querySelectorAll('.MuiDataGrid-row')]
                        .filter(row => row.getClientRects().length > 0);
                    const signature = rows.map(row =>
                        row.getAttribute('data-id')
                        || row.getAttribute('data-rowindex')
                        || row.getAttribute('aria-rowindex')
                        || row.innerText.trim()
                    ).join('|');
                    const sebelum = el.scrollTop;
                    const batas = el.scrollHeight - el.clientHeight;
                    const tujuan = Math.min(
                        sebelum + Math.max(el.clientHeight * 0.65, 1),
                        batas
                    );
                    el.scrollTop = tujuan;
                    return {
                        sebelum,
                        sesudah: el.scrollTop,
                        batas,
                        tujuan,
                        signature
                    };
                }"""
            )
        else:
            baris_terakhir = page.locator(baris_selector).last
            posisi_scroll = baris_terakhir.evaluate(
                """lastRow => {
                    let target = lastRow.parentElement;
                    while (target && target !== document.body) {
                        const style = getComputedStyle(target);
                        if (
                            /(auto|scroll|overlay)/.test(style.overflowY)
                            && target.scrollHeight > target.clientHeight + 2
                        ) {
                            const sebelum = target.scrollTop;
                            const batas = target.scrollHeight - target.clientHeight;
                            target.scrollTop = Math.min(
                                sebelum + Math.max(target.clientHeight * 0.65, 1),
                                batas
                            );
                            return {
                                sebelum,
                                sesudah: target.scrollTop,
                                batas,
                                signature: [...document.querySelectorAll(
                                    '.MuiDataGrid-row, tr:has(button)'
                                )].map(row => row.innerText.trim()).join('|')
                            };
                        }
                        target = target.parentElement;
                    }
                    return {
                        sebelum: window.scrollY,
                        sesudah: window.scrollY,
                        batas: window.scrollY,
                        signature: [...document.querySelectorAll(
                            '.MuiDataGrid-row, tr:has(button)'
                        )].map(row => row.innerText.trim()).join('|')
                    };
                }"""
            )

        if posisi_scroll["sesudah"] <= posisi_scroll["sebelum"] + 1:
            if posisi_scroll["sesudah"] >= posisi_scroll["batas"] - 5:
                status_scroll["mentok"] = True
                status_scroll["baris_terakhir_diproses"] = True
                print(
                    f"[INFO v5.32] Scroll mentok: {len(baris_selesai)} dari "
                    f"{jumlah_diharapkan or '?'} baris ditangani. "
                    "Tetap lanjut ke halaman berikutnya."
                )
            else:
                status_scroll["kesalahan"] = (
                    "Scroll tidak bergerak sebelum mencapai batas bawah tabel."
                )
            return

        if grid_virtual:
            try:
                page.wait_for_function(
                    """({previousTop, previousSignature}) => {
                        const scroller = [...document.querySelectorAll(
                            '.MuiDataGrid-virtualScroller'
                        )].find(element => element.getClientRects().length > 0);
                        if (!scroller) return false;
                        const signature = [...scroller.querySelectorAll(
                            '.MuiDataGrid-row'
                        )].filter(row => row.getClientRects().length > 0).map(row =>
                            row.getAttribute('data-id')
                            || row.getAttribute('data-rowindex')
                            || row.getAttribute('aria-rowindex')
                            || row.innerText.trim()
                        ).join('|');
                        return Math.abs(scroller.scrollTop - previousTop) > 1
                            && signature !== previousSignature;
                    }""",
                    arg={
                        "previousTop": posisi_scroll["sebelum"],
                        "previousSignature": posisi_scroll["signature"],
                    },
                    timeout=4000,
                    polling=100,
                )
            except PlaywrightTimeoutError:
                posisi_terkini = wadah_scroll.evaluate(
                    "el => ({top: el.scrollTop, max: el.scrollHeight - el.clientHeight})"
                )
                if posisi_terkini["top"] >= posisi_terkini["max"] - 2:
                    status_scroll["mentok"] = True
                    status_scroll["baris_terakhir_diproses"] = True
                    print(
                        f"[INFO v5.32] Scroll mentok: {len(baris_selesai)} dari "
                        f"{jumlah_diharapkan or '?'} baris ditangani. "
                        "Tetap lanjut ke halaman berikutnya."
                    )
                    return
                else:
                    percobaan_tanpa_baris_baru += 1
                    if percobaan_tanpa_baris_baru >= 3:
                        status_scroll["kesalahan"] = (
                            "Scroll bergerak tetapi baris tabel tidak berubah "
                            "setelah 3 percobaan; halaman dihentikan agar tidak "
                            "mengulang render."
                        )
                        return
        else:
            page.wait_for_timeout(250)

def iter_baris_anggota(dialog, tabel):
    row_selector = "tbody tr, [role='row'], tr"
    rows = tabel.locator(row_selector)
    scroller = dialog.locator(".MuiDataGrid-virtualScroller").first
    grid_virtual = scroller.count() > 0
    rows_seen = set()

    while True:
        jumlah_rows = rows.count()
        for row_index in range(jumlah_rows):
            row = rows.nth(row_index)
            cells = row.locator("td, [role='cell']")
            if cells.count() < 2:
                continue

            row_key = (
                row.get_attribute("data-id")
                or row.get_attribute("data-rowindex")
                or row.get_attribute("aria-rowindex")
            )
            if not row_key:
                row_key = bersihkan_teks_data(row.inner_text())
            if not row_key or row_key in rows_seen:
                continue

            rows_seen.add(row_key)
            yield row

        if not grid_virtual:
            return

        posisi_scroll = scroller.evaluate(
            """el => {
                const rows = el.querySelectorAll('.MuiDataGrid-row');
                const lastRow = rows[rows.length - 1];
                const lastKey = lastRow && (
                    lastRow.getAttribute('data-id')
                    || lastRow.getAttribute('data-rowindex')
                    || lastRow.getAttribute('aria-rowindex')
                    || lastRow.innerText
                );
                const sebelum = el.scrollTop;
                const batas = el.scrollHeight - el.clientHeight;
                el.scrollTop = Math.min(
                    sebelum + Math.max(el.clientHeight * 0.6, 100),
                    batas
                );
                return {
                    sebelum,
                    sesudah: el.scrollTop,
                    batas,
                    lastKey,
                    diBawah: el.scrollTop >= batas - 2
                };
            }"""
        )
        if posisi_scroll["sesudah"] <= posisi_scroll["sebelum"]:
            return
        try:
            dialog.page.wait_for_function(
                """previousLastKey => {
                    const rows = [...document.querySelectorAll(
                        '[role="dialog"] .MuiDataGrid-row'
                    )].filter(row => row.getClientRects().length > 0);
                    const lastRow = rows[rows.length - 1];
                    if (!lastRow) return false;
                    const lastKey = lastRow.getAttribute('data-id')
                        || lastRow.getAttribute('data-rowindex')
                        || lastRow.getAttribute('aria-rowindex')
                        || lastRow.innerText;
                    return lastKey !== previousLastKey;
                }""",
                arg=posisi_scroll["lastKey"],
                timeout=1500,
                polling=50,
            )
        except PlaywrightTimeoutError:
            dialog.page.wait_for_timeout(200)

def bersihkan_teks_data(value):
    if value is None:
        return ""

    teks = str(value)
    teks = teks.replace("\u00a0", " ")
    teks = teks.replace("\r", " ").replace("\n", " ")
    teks = re.sub(
        r"(?i)content_copy|check_circle_icon|file_upload_icon|salin nik|copy",
        " ",
        teks,
    )
    teks = re.sub(r"\s+", " ", teks).strip()
    return teks

def is_nik_utuh(value):
    return bool(re.fullmatch(r"\d{16}", bersihkan_teks_data(value)))

def nik_dari_excel(value, nama_kolom, nomor_baris, file_sumber_nik):
    if value is None or isinstance(value, bool):
        nik = ""
    elif isinstance(value, int):
        nik = f"{value:016d}" if 0 <= value < 10**16 else str(value)
    elif isinstance(value, float) and value.is_integer():
        integer_value = int(value)
        nik = (
            f"{integer_value:016d}"
            if 0 <= integer_value < 10**16
            else str(integer_value)
        )
    else:
        nik = bersihkan_teks_data(value)

    if not is_nik_utuh(nik):
        raise ValueError(
            f"{nama_kolom} pada baris {nomor_baris} di {file_sumber_nik} "
            f"tidak valid: {value!r}."
        )
    return nik

def muat_daftar_nik(file_sumber_nik):
    file_sumber_nik = Path(file_sumber_nik)
    if not file_sumber_nik.exists():
        raise FileNotFoundError(
            f"File sumber NIK tidak ditemukan: {file_sumber_nik.resolve()}"
        )

    records = []
    required_fields = ("NAMA", "NO KK", "NIK")
    if file_sumber_nik.suffix.lower() == ".json":
        with file_sumber_nik.open(encoding="utf-8") as source_file:
            json_records = json.load(source_file)
        if not isinstance(json_records, list) or any(
            not isinstance(record, dict) for record in json_records
        ):
            raise ValueError(
                f"Format JSON acuan tidak valid pada {file_sumber_nik}; "
                "gunakan array JSON berisi objek data."
            )
        for row_number, source_record in enumerate(json_records, start=1):
            normalized_record = {
                str(key).strip().upper(): value
                for key, value in source_record.items()
            }
            missing_fields = [
                field
                for field in required_fields
                if field not in normalized_record
            ]
            if missing_fields:
                raise ValueError(
                    f"Kolom {', '.join(missing_fields)} tidak ditemukan "
                    f"pada baris {row_number} di {file_sumber_nik}."
                )
            records.append(
                (row_number, normalized_record)
            )
    else:
        workbook = load_workbook(
            file_sumber_nik, read_only=True, data_only=True
        )
        try:
            worksheet = workbook.active
            headers = [
                bersihkan_teks_data(
                    worksheet.cell(row=1, column=column).value
                ).upper()
                for column in range(1, worksheet.max_column + 1)
            ]
            columns = {
                header: index + 1 for index, header in enumerate(headers)
            }
            missing_headers = [
                field for field in required_fields if field not in columns
            ]
            if missing_headers:
                raise ValueError(
                    f"Kolom {', '.join(missing_headers)} tidak ditemukan "
                    f"pada baris header {file_sumber_nik}."
                )

            for row_number in range(2, worksheet.max_row + 1):
                records.append(
                    (
                        row_number,
                        {
                            header: worksheet.cell(
                                row=row_number, column=column
                            ).value
                            for header, column in columns.items()
                        },
                    )
                )
        finally:
            workbook.close()

    data_pengurus = []
    data_by_nik = {}
    for row_number, source_record in records:
        nik_value = source_record.get("NIK")
        if nik_value is None or not bersihkan_teks_data(nik_value):
            continue

        nik = nik_dari_excel(nik_value, "NIK", row_number, file_sumber_nik)
        no_kk = nik_dari_excel(
            source_record.get("NO KK"), "NO KK", row_number, file_sumber_nik
        )
        nama = bersihkan_teks_data(source_record.get("NAMA"))
        if not nama:
            raise ValueError(
                f"NAMA kosong pada baris {row_number} di {file_sumber_nik}."
            )

        record = {
            "NIK": nik,
            "NO KK": no_kk,
            "NAMA": nama,
            "Desa": bersihkan_teks_data(source_record.get("DESA")),
            "Alamat": bersihkan_teks_data(source_record.get("ALAMAT")),
        }
        record_sebelumnya = data_by_nik.get(nik)
        if record_sebelumnya is not None:
            if record_sebelumnya != record:
                raise ValueError(
                    f"NIK duplikat {nik} memiliki data berbeda pada "
                    f"{file_sumber_nik} (baris {row_number})."
                )
            continue
        data_pengurus.append(record)
        data_by_nik[nik] = record

    if not data_pengurus:
        raise ValueError(
            f"Tidak ada NIK yang dapat diproses di {file_sumber_nik}."
        )

    print(
        f"[INFO] {len(data_pengurus)} data pengurus unik "
        f"(NIK, No. KK, dan nama) dimuat dari {file_sumber_nik}."
    )
    return data_pengurus

def kolom_pencarian_nik(page):
    fields = page.locator(
        'input[type="text"][maxlength="16"].MuiInputBase-inputAdornedEnd:visible'
    )
    nik_fields = []
    for index in range(fields.count()):
        field = fields.nth(index)
        metadata = field.evaluate(
            """element => {
                const labels = element.labels
                    ? [...element.labels].map(label => label.innerText.trim())
                    : [];
                const form = element.closest('.MuiFormControl-root');
                return {
                    labels,
                    formText: form ? form.innerText.trim() : ''
                };
            }"""
        )
        labels = [label.casefold() for label in metadata["labels"]]
        form_text = re.sub(r"\s+", " ", metadata["formText"]).casefold()
        is_nik_label = any(
            re.fullmatch(r"nik(?:\s+0/16)?", label)
            for label in labels
        )
        is_nik_field = is_nik_label or (
            re.search(r"\bnik\b", form_text)
            and not re.search(r"\bno\s*kk\b", form_text)
        )
        if is_nik_field:
            nik_fields.append(field)

    if not nik_fields:
        raise RuntimeError(
            "Kolom NIK dengan penghitung 0/16 tidak ditemukan. Pastikan halaman "
            "Filter Data Verifikasi sudah terbuka."
        )
    if len(nik_fields) > 1:
        raise RuntimeError(
            f"Ditemukan {len(nik_fields)} kolom berlabel NIK; "
            "pilihan dibatalkan agar NIK tidak dimasukkan ke kolom yang salah."
        )
    return nik_fields[0]

def cari_baris_kpm(page, input_nik, nik_target, nama_target):
    page.evaluate(
        "(nik) => navigator.clipboard.writeText(nik)",
        nik_target,
    )
    input_nik.click()
    input_nik.press("Control+A")
    input_nik.press("Control+V")
    tombol_cari = page.get_by_role("button", name="Cari Data", exact=True)
    if tombol_cari.count() != 1:
        raise RuntimeError(
            "Tombol 'Cari Data' tidak ditemukan atau tidak tunggal pada halaman."
        )
    tombol_cari.click()

    deadline = time.monotonic() + 15
    while time.monotonic() < deadline:
        rows = page.locator(
            ".MuiDataGrid-row:visible, tbody tr:visible, tr:has(button):visible"
        )
        for row_index in range(rows.count()):
            row = rows.nth(row_index)
            cells = row.locator(".MuiDataGrid-cell, td")
            if cells.count() < 4:
                continue
            teks_baris = re.sub(
                r"\s+", " ", bersihkan_teks_data(row.inner_text())
            ).casefold()
            if nama_target.casefold() in teks_baris:
                return row
        page.wait_for_timeout(400)

    raise LookupError(
        f"KPM dengan NIK {nik_target} tidak ditemukan melalui kolom pencarian."
    )

def standardisasi_kategori_komponen(teks_kotor):
    """Memastikan komponen mencakup kategori baku:
       Sekolah, Lansia, Disabilitas, Anak usia dini, Ibu Hamil"""
    if not teks_kotor:
        return "-"

    kategori = re.sub(r"\s+", " ", bersihkan_teks_data(teks_kotor)).upper()
    kategori_baku = {
        "SEKOLAH": "Sekolah",
        "LANSIA": "Lansia",
        "ANAK USIA DINI": "Anak usia dini",
        "DISABILITAS": "Disabilitas",
        "IBU HAMIL": "Ibu hamil",
    }
    return kategori_baku.get(kategori, "-")

def ambil_kategori_komponen(teks_header, teks_cells):
    idx_komponen = next(
        (
            idx for idx, header in enumerate(teks_header)
            if ("KOMPONEN" in header or "KATEGORI" in header)
            and "BULAN" not in header
        ),
        None,
    )
    if idx_komponen is not None and idx_komponen < len(teks_cells):
        kategori = standardisasi_kategori_komponen(teks_cells[idx_komponen])
        if kategori != "-":
            return kategori

    for teks_cell in teks_cells:
        kategori = standardisasi_kategori_komponen(teks_cell)
        if kategori != "-":
            return kategori

    return "-"

def ambil_data_asli_cell(page, cell_element):
    tombol_copy = None
    try:
        tombol_copy = cell_element.locator(
            "button[aria-label='Salin NIK'], "
            "button[data-testid='ContentCopyIcon'], "
            "button:has([data-testid*='ContentCopyIcon']), "
            "button:has-text('content_copy'), "
            "button.MuiIconButton-root"
        ).first
        if tombol_copy.count() == 0:
            tombol_copy = cell_element.locator(
                "[role='button'][aria-label='Salin NIK'], "
                "[role='button']:has([data-testid*='ContentCopyIcon']), "
                "[role='button']:has-text('content_copy')"
            ).first
        if tombol_copy.count() == 0:
            tombol_copy = cell_element.locator("button, [role='button']").first

        if tombol_copy.count() == 0:
            return ""

        for attr in ["data-text", "value", "title"]:
            value = bersihkan_teks_data(tombol_copy.get_attribute(attr))
            if is_nik_utuh(value):
                return value

        try:
            clipboard_sebelum = bersihkan_teks_data(
                page.evaluate("navigator.clipboard.readText()")
            )
        except Exception:
            clipboard_sebelum = ""

        try:
            page.evaluate("navigator.clipboard.writeText('')")
            clipboard_sebelum = ""
        except Exception as e:
            print(f" [WARN] Clipboard tidak dapat dikosongkan sebelum copy: {e}")

        tombol_copy.click(force=True)
        try:
            page.wait_for_function(
                """async (clipboardSebelum) => {
                    const hasil = (await navigator.clipboard.readText()).trim();
                    return /^\\d{16}$/.test(hasil)
                        && (hasil !== clipboardSebelum || !/^\\d{16}$/.test(clipboardSebelum));
                }""",
                arg=clipboard_sebelum,
                timeout=3000,
                polling=50,
            )
        except PlaywrightTimeoutError:
            return ""

        hasil = bersihkan_teks_data(
            page.evaluate("navigator.clipboard.readText()")
        )
        return hasil if is_nik_utuh(hasil) else ""
    except Exception as e:
        print(f" [WARN] Gagal menyalin nilai NIK/KK: {e}")

    return ""

def tutup_dialog_detail(page):
    dialogs = page.locator("[role='dialog']:visible, .MuiDialog-root:visible")
    if dialogs.count() == 0:
        return True

    dialog = dialogs.last
    dialog_handle = dialog.element_handle(timeout=3000)
    if dialog_handle is None:
        return False

    tombol_tutup = dialog_handle.query_selector(
        "button[aria-label*='close' i], button[title*='close' i], "
        "button:has(svg[data-testid='CloseIcon'])"
    )
    try:
        if tombol_tutup is not None:
            tombol_tutup.click(force=True, timeout=3000)
        else:
            page.keyboard.press("Escape")

        dialog_handle.wait_for_element_state("hidden", timeout=5000)
        return True
    except Exception:
        try:
            page.keyboard.press("Escape")
            dialog_handle.wait_for_element_state("hidden", timeout=3000)
            return True
        except Exception:
            return False

def tutup_peringatan_rate_limit(page, timeout=2500):
    peringatan = page.get_by_text("Too Many Requests.", exact=True).last
    try:
        peringatan.wait_for(state="visible", timeout=timeout)
    except PlaywrightTimeoutError:
        return False

    print(
        "[WARN v5.32] SIKS membatasi permintaan (Too Many Requests). "
        "Peringatan ditutup; proses akan menunggu sebelum mencoba ulang."
    )
    tombol_tutup = page.get_by_role("button", name="Tutup", exact=True).last
    if tombol_tutup.count() == 0:
        raise RuntimeError(
            "Peringatan Too Many Requests tampil, tetapi tombol 'Tutup' tidak ditemukan."
        )
    tombol_tutup.click(force=True, timeout=5000)
    peringatan.wait_for(state="hidden", timeout=5000)
    return True

def baca_info_sekolah(page, row, headers, dialog_anggota):
    nama_sekolah = ""
    kelas = ""
    rate_limit_berulang = False
    month_index = next(
        (
            index for index, header in enumerate(headers)
            if "BULAN" in header and "1" in header
        ),
        3, 
    )
    cells = row.locator("td").all()
    if not cells:
        cells = row.locator("[role='cell']").all()
    if month_index >= len(cells):
        print(" [WARN] Kolom Bulan Ke-1 tidak ditemukan pada baris komponen sekolah.")
        return nama_sekolah, kelas, rate_limit_berulang

    tombol_upload = cells[month_index].locator("button, [role='button']").first
    if tombol_upload.count() == 0:
        print(" [WARN] Tombol Bulan Ke-1 tidak ditemukan untuk data sekolah.")
        return nama_sekolah, kelas, rate_limit_berulang

    jumlah_dialog_sebelum = page.locator("[role='dialog']:visible").count()
    try:
        berhasil_membuka = False
        for percobaan in range(2):
            if percobaan:
                print(
                    "[INFO v5.32] Menunggu 30 detik sebelum mencoba ulang "
                    "permintaan data sekolah."
                )
                page.wait_for_timeout(30000)
            else:
                page.wait_for_timeout(1000)

            tombol_upload.click(force=True, timeout=5000)
            if tutup_peringatan_rate_limit(page):
                if percobaan == 0:
                    continue
                rate_limit_berulang = True
                raise RuntimeError(
                    "SIKS masih menolak permintaan data sekolah setelah satu kali percobaan ulang."
                )

            berhasil_membuka = True
            break

        if not berhasil_membuka:
            raise RuntimeError("Data sekolah tidak berhasil dibuka.")

        page.get_by_text("Data Sekolah", exact=True).wait_for(
            state="visible",
            timeout=10000,
        )
        dialog_upload = page.locator("[role='dialog']:visible").last
        teks_dialog = dialog_upload.inner_text()
        match_data = re.search(
            r"Data Sekolah\s*([\s\S]*?)(?:Status Komitmen|$)",
            teks_dialog,
            re.IGNORECASE,
        )
        if match_data:
            baris_data = [
                bersihkan_teks_data(line)
                for line in match_data.group(1).splitlines()
                if bersihkan_teks_data(line)
            ]
            if baris_data:
                nama_sekolah = baris_data[0]
            for line in baris_data[1:]:
                match_kelas = re.search(r"\bKelas\s*:?\s*(.+)", line, re.IGNORECASE)
                if match_kelas:
                    kelas = bersihkan_teks_data(match_kelas.group(1))
                    break

        if not tutup_dialog_detail(page):
            raise RuntimeError("Form detail sekolah tidak berhasil ditutup.")
        dialog_anggota.wait_for(state="visible", timeout=5000)
        return nama_sekolah, kelas, rate_limit_berulang
    except Exception as e:
        print(f" [WARN] Gagal membaca detail sekolah: {e}")
        if page.locator("[role='dialog']:visible").count() > jumlah_dialog_sebelum:
            tutup_dialog_detail(page)
            dialog_anggota.wait_for(state="visible", timeout=5000)
        return nama_sekolah, kelas, rate_limit_berulang

def run():
    global list_data_final, status_proses
    list_data_final = []
    status_proses = {
        "status": "starting",
        "started_at": waktu_sekarang(),
        "updated_at": waktu_sekarang(),
        "total_nik": None,
        "processed_nik": 0,
        "result_rows": 0,
        "failed_nik": [],
        "unprocessed_nik": [],
    }
    tulis_json_atomik(STATUS_FILE, status_proses)

    if not COOKIE_FILE.exists():
        raise FileNotFoundError(
            f"File cookie SIKS tidak ditemukan: {COOKIE_FILE}"
        )
    file_sumber_nik = cari_file_acuan()
    data_pengurus = muat_daftar_nik(file_sumber_nik)
    data_pengurus_by_nik = {
        data["NIK"]: data for data in data_pengurus
    }
    daftar_nik = list(data_pengurus_by_nik)
    status_proses.update(
        {
            "status": "running",
            "total_nik": len(daftar_nik),
            "updated_at": waktu_sekarang(),
            "reference_file": file_sumber_nik.name,
        }
    )
    tulis_json_atomik(STATUS_FILE, status_proses)
    print(
        "[INFO] Hasil akan disimpan sebagai XLSX di folder _result_scraping/result_01_app_siks "
        "setelah proses selesai dan ENTER ditekan."
    )
    baris_sukses_count = 0
    data_final_keys = set()
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(
            storage_state=str(COOKIE_FILE),
            permissions=["clipboard-read", "clipboard-write"]
        )
        page = context.new_page()
        url_siks = "https://siks.kemensos.go.id/"
        page.set_default_navigation_timeout(60000)
        try:
            response = page.goto(
                url_siks,
                wait_until="commit",
                timeout=60000,
            )
        except PlaywrightTimeoutError as e:
            raise RuntimeError(
                "Browser tidak menerima respons awal dari SIKS dalam 60 detik. "
                "Periksa koneksi, VPN/proxy, atau buka URL SIKS secara manual."
            ) from e

        if response is not None and response.status >= 400:
            raise RuntimeError(
                f"Halaman SIKS merespons HTTP {response.status}: {page.url}"
            )

        print("\n[INFO] Beranda SIKS sudah dibuka.")
        print(
            "[PETUNJUK] Silakan login jika diminta, lalu buka menu "
            "'Verifikasi Komitmen KPM PKH' secara manual di browser. "
            "Setelah menu terbuka, tekan ENTER untuk mulai."
        )
        flush_input()
        input(
            "\n[TUNGGU] Tekan ENTER untuk mulai mencari berdasarkan NIK..."
        )

        input_nik = kolom_pencarian_nik(page)
        print(
            "[INFO] Setiap NIK dari kolom 'NIK' di file sumber akan dicari "
            "satu per satu; daftar KPM tidak akan dipindai dengan scroll."
        )

        nik_gagal_dicari = []
        nik_belum_diproses = []
        while True:
            print(
                f"\n--- MEMPROSES {len(daftar_nik)} NIK DENGAN PENCARIAN LANGSUNG ---"
            )
            baris_selesai_diproses = set()
            hentikan_halaman_rate_limit = False
            for i, nik_target in enumerate(daftar_nik, start=1):
                label_baris = f"{i}/{len(daftar_nik)}"
                data_sumber = data_pengurus_by_nik[nik_target]
                try:
                    baris_aktif = cari_baris_kpm(
                        page,
                        input_nik,
                        nik_target,
                        data_sumber["NAMA"],
                    )
                except Exception as e:
                    nik_gagal_dicari.append((nik_target, str(e)))
                    print(f"[GAGAL CARI NIK] {label_baris} | {nik_target} | {e}")
                    simpan_progres(
                        status_proses,
                        i,
                        nik_gagal_dicari,
                        nik_belum_diproses,
                    )
                    continue

                kunci_baris_aktif = kunci_baris_tabel(baris_aktif)
                nik_asli = data_sumber["NIK"]
                kk_asli = data_sumber["NO KK"]
                nama_pengurus = data_sumber["NAMA"]
                desa_sumber = data_sumber["Desa"]
                alamat_sumber = data_sumber["Alamat"]
                
                try:
                    sel_kolom = baris_aktif.locator(".MuiDataGrid-cell, td").all()
                    if len(sel_kolom) < 4:
                        raise RuntimeError(
                            "Baris hasil pencarian tidak memiliki kolom data yang "
                            "diperlukan untuk membuka detail."
                        )

                    tombol_detail = baris_aktif.locator("button[aria-label='detail']").first
                    if tombol_detail.count() == 0:
                        tombol_detail = baris_aktif.locator("button:has(span.material-icons:has-text('verified_icon'))").first
                    if tombol_detail.count() == 0:
                        tombol_detail = baris_aktif.locator("button").last

                    masalah_identitas = []
                    if not is_nik_utuh(nik_asli):
                        nik_asli = ""
                        masalah_identitas.append("NIK pengurus tidak valid di Excel")
                    if not is_nik_utuh(kk_asli):
                        kk_asli = ""
                        masalah_identitas.append("No. KK pengurus tidak valid di Excel")
                    print(
                        f"[STATUS] {label_baris} | No. KK "
                        f"{kk_asli or 'TIDAK VALID'} | {nama_pengurus} | "
                        "DALAM PROSES."
                    )

                    tombol_detail.click(force=True)
                    page.wait_for_selector(
                        ".MuiDialogContent-root table tbody tr, "
                        "[role='dialog'] table tbody tr, "
                        "[role='dialog'] [role='rowgroup'] [role='row']",
                        timeout=12000,
                    )

                    dialog_detail = page.locator("[role='dialog']:visible").last
                    tabel_modal = dialog_detail.locator("table").first
                    if tabel_modal.count() == 0:
                        tabel_modal = dialog_detail.locator("[role='table']").first

                    headers_modal = tabel_modal.locator("thead th, thead [role='columnheader'], [role='rowgroup'] [role='columnheader']").all()
                    teks_header = [
                        re.sub(r"\s+", " ", bersihkan_teks_data(h.inner_text()).upper())
                        for h in headers_modal
                    ]
                    
                    wadah_baris_modal = (
                        tabel_modal
                        if tabel_modal.locator(
                            "tbody tr, [role='row'], tr"
                        ).count() > 0
                        else dialog_detail
                    )

                    kpm_sukses = not masalah_identitas
                    jumlah_anggota_diproses = 0
                    jumlah_anggota_sukses = 0
                    
                    for row_index, r in enumerate(
                        iter_baris_anggota(dialog_detail, wadah_baris_modal),
                        start=1,
                    ):
                        cells_modal = r.locator("td, [role='cell']")
                        jumlah_sel_modal = cells_modal.count()
                        if jumlah_sel_modal < 2:
                            print(
                                f" [WARN] Baris anggota {row_index} tidak lagi "
                                "ter-render; dilewati agar proses lanjut."
                            )
                            kpm_sukses = False
                            continue
                        cols_modal = [
                            cells_modal.nth(cell_index)
                            for cell_index in range(jumlah_sel_modal)
                        ]

                        nik_agt = ""
                        nama_agt = "Tidak Terbaca"

                        list_teks_cell = [
                            bersihkan_teks_data(teks)
                            for teks in cells_modal.all_inner_texts()
                        ]

                        idx_nik = next((idx for idx, h in enumerate(teks_header) if "NIK" in h), 0)
                        idx_nama = next((idx for idx, h in enumerate(teks_header) if "NAMA PENERIMA" in h or "NAMA" in h), 1)

                        tombol_nik = r.locator("button[aria-label='Salin NIK'], [role='button'][aria-label='Salin NIK'], button:has([data-testid*='ContentCopyIcon'])").first
                        if tombol_nik.count() > 0:
                            cell_nik = tombol_nik.locator("xpath=ancestor::td[1] | ancestor::div[@role='cell'][1]")
                            if cell_nik.count() > 0:
                                nik_agt = ambil_data_asli_cell(page, cell_nik)
                            else:
                                nik_agt = ambil_data_asli_cell(page, tombol_nik.locator("xpath=parent::*"))
                        elif idx_nik < len(cols_modal):
                            nik_agt = ambil_data_asli_cell(page, cols_modal[idx_nik])
                        
                        if not is_nik_utuh(nik_agt):
                            for cell in cols_modal:
                                temp_nik = ambil_data_asli_cell(page, cell)
                                if is_nik_utuh(temp_nik):
                                    nik_agt = temp_nik
                                    break

                        if idx_nama < len(list_teks_cell):
                            nama_agt = list_teks_cell[idx_nama]
                        if not nama_agt or nama_agt == "-" or is_nik_utuh(nama_agt) or any(k in nama_agt.upper() for k in ["SEKOLAH", "HAMIL", "LANSIA", "BALITA", "DISABILITAS", "ANAK USIA DINI"]):
                            for txt in list_teks_cell:
                                if len(txt) > 3 and not re.fullmatch(r"\d+", txt) and "*" not in txt and not any(k in txt.upper() for k in ["SEKOLAH", "HAMIL", "BALITA", "LANSIA", "DISABILITAS", "ANAK USIA DINI", "BULAN"]):
                                    nama_agt = txt
                                    break

                        komponen_agt = ambil_kategori_komponen(
                            teks_header,
                            list_teks_cell,
                        )

                        masalah_baris = list(masalah_identitas)
                        if not is_nik_utuh(nik_agt):
                            nik_agt = "Tidak Terbaca"
                            masalah_baris.append("NIK anggota tidak berhasil disalin")

                        nama_sekolah = ""
                        kelas = ""
                        rate_limit_berulang = False
                        if komponen_agt == "-":
                            masalah_baris.append("Kategori komponen tidak berhasil dibaca")
                        if komponen_agt == "Sekolah":
                            (
                                nama_sekolah,
                                kelas,
                                rate_limit_berulang,
                            ) = baca_info_sekolah(
                                page,
                                r,
                                teks_header,
                                dialog_detail,
                            )
                            if rate_limit_berulang:
                                masalah_baris.append(
                                    "Data sekolah tidak dibaca karena batas permintaan SIKS."
                                )

                        if masalah_baris:
                            kpm_sukses = False
                        if komponen_agt == "Sekolah" and not nama_sekolah:
                            masalah_baris.append("Nama sekolah tidak berhasil dibaca")
                            kpm_sukses = False

                        status_baris = (
                            "SUKSES" if not masalah_baris
                            else "GAGAL: " + "; ".join(masalah_baris)
                        )
                        jumlah_anggota_diproses += 1
                        if status_baris == "SUKSES":
                            jumlah_anggota_sukses += 1
                        else:
                            kpm_sukses = False

                        data_baris = {
                            'NIK_Pengurus': nik_asli,
                            'NOK_Pengurus': kk_asli,
                            'Nama_Pengurus': nama_pengurus,
                            'Desa': desa_sumber,
                            'Alamat': alamat_sumber,
                            'NIK_Anggota': nik_agt,
                            'Nama_Anggota': nama_agt,
                            'Komponen': komponen_agt,       # Diubah dari Kategori_Komponen[cite: 1]
                            'Nama Sekolah': nama_sekolah,
                            'Kelas': kelas,
                            'Status': status_baris,
                            
                        }
                        kunci_data = kunci_data_hasil(data_baris)
                        baris_baru = kunci_data not in data_final_keys
                        if not baris_baru:
                            print(
                                f"[INFO v5.32] Baris duplikat dilewati: "
                                f"{nama_pengurus} / {nama_agt} / {komponen_agt}"
                            )
                        else:
                            data_final_keys.add(kunci_data)
                            list_data_final.append(data_baris)

                        if baris_baru and status_baris == "SUKSES":
                            baris_sukses_count += 1
                            if baris_sukses_count % 25 == 0:
                                print(
                                    f" -> {baris_sukses_count} baris sukses "
                                    "terkumpul; hasil menunggu proses selesai."
                                )

                        if rate_limit_berulang:
                            hentikan_halaman_rate_limit = True
                            break

                    if not tutup_dialog_detail(page):
                        raise RuntimeError("Dialog detail anggota tidak berhasil ditutup.")

                    if (
                        jumlah_anggota_diproses > 0
                        and jumlah_anggota_sukses == jumlah_anggota_diproses
                        and kpm_sukses
                    ):
                        print(
                            f"[STATUS] {label_baris} | No. KK {kk_asli} | "
                            f"{nama_pengurus} | SELESAI - "
                            f"{jumlah_anggota_sukses} anggota sukses."
                        )
                    else:
                        print(
                            f"[STATUS] {label_baris} | "
                            f"No. KK {kk_asli or 'TIDAK TERBACA'} | "
                            f"{nama_pengurus} | BELUM SELESAI - "
                            f"{jumlah_anggota_sukses}/{jumlah_anggota_diproses} "
                            "anggota sukses."
                        )
                    
                    label_hasil = "OK" if kpm_sukses else "WARN"
                    print(
                        f"[{label_hasil} v5.32] {label_baris} | "
                        f"{nama_pengurus} | Komponen: {komponen_agt}"
                    )
                    if kunci_baris_aktif:
                        baris_selesai_diproses.add(kunci_baris_aktif)

                except Exception as e:
                    print(f" [GAGAL] {label_baris}: {e}")
                    nik_gagal_dicari.append((nik_target, str(e)))
                    tutup_dialog_detail(page)
                finally:
                    if kunci_baris_aktif:
                        baris_selesai_diproses.add(kunci_baris_aktif)
                simpan_progres(
                    status_proses,
                    i,
                    nik_gagal_dicari,
                    nik_belum_diproses,
                )
                if hentikan_halaman_rate_limit:
                    break

            if hentikan_halaman_rate_limit:
                print(
                    "[STOP v5.32] SIKS masih membatasi permintaan setelah jeda dan "
                    "satu percobaan ulang. Proses berhenti agar tidak menambah "
                    "permintaan; data parsial menunggu konfirmasi ENTER."
                )
                nik_belum_diproses = daftar_nik[i - 1:]
                simpan_progres(
                    status_proses,
                    len(daftar_nik) - len(nik_belum_diproses),
                    nik_gagal_dicari,
                    nik_belum_diproses,
                )
                break

            print(
                f"[INFO PENCARIAN NIK] {len(baris_selesai_diproses)} baris "
                "pencarian sudah ditangani; hasil menunggu proses selesai."
            )
            break
        status_proses["status"] = "awaiting_confirmation"
        simpan_progres(
            status_proses,
            len(daftar_nik) - len(nik_belum_diproses),
            nik_gagal_dicari,
            nik_belum_diproses,
        )
        if nik_gagal_dicari or nik_belum_diproses:
            print("[RINGKASAN NIK GAGAL]")
            for nik_gagal, alasan in nik_gagal_dicari:
                print(f" - {nik_gagal}: {alasan}")
            if nik_belum_diproses:
                print(
                    f" - {len(nik_belum_diproses)} NIK belum selesai karena "
                    "proses dihentikan setelah batas permintaan SIKS."
                )
        else:
            print(
                "\n[PROGRAM SELESAI v5.32] Pengambilan data selesai."
            )
        browser.close()
        print(
            "\n[TUNGGU] Tekan ENTER untuk menyimpan hasil XLSX "
            "dan menyelesaikan proses..."
        )
        flush_input()
        input()
        result_path = simpan_hasil(list_data_final)
        status_proses.update(
            {
                "status": (
                    "incomplete"
                    if nik_gagal_dicari or nik_belum_diproses
                    else "completed"
                ),
                "updated_at": waktu_sekarang(),
                "finished_at": waktu_sekarang(),
                "result_file": Path(result_path).name,
            }
        )
        tulis_json_atomik(STATUS_FILE, status_proses)
        print(
            f"[HASIL] {len(list_data_final)} baris XLSX disimpan ke {result_path}."
        )
        if nik_gagal_dicari or nik_belum_diproses:
            raise RuntimeError(
                f"Proses menyisakan {len(nik_gagal_dicari)} NIK gagal dicari dan "
                f"{len(nik_belum_diproses)} NIK belum diproses."
            )


if __name__ == "__main__":
    try:
        run()
    except Exception:
        print("\n[ERROR FATAL v5.32] Proses berhenti karena kesalahan berikut:")
        traceback.print_exc()
        if status_proses is not None and status_proses.get("status") in {
            "starting",
            "running",
            "awaiting_confirmation",
        }:
            status_proses.update(
                {
                    "status": "failed",
                    "updated_at": waktu_sekarang(),
                    "finished_at": waktu_sekarang(),
                    "error": traceback.format_exc(),
                }
            )
            try:
                tulis_json_atomik(STATUS_FILE, status_proses)
            except Exception:
                print("[ERROR] File status gagal diperbarui:")
                traceback.print_exc()
        try:
            input("\nTekan ENTER untuk menutup jendela ini...")
        except EOFError:
            pass