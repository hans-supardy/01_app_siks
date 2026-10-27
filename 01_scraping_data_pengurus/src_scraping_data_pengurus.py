import json
import tempfile
import time
from pathlib import Path

from openpyxl import Workbook
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import Locator, Page, sync_playwright


BASE_URL = "https://siks.kemensos.go.id/"
SCRIPT_FILE = Path(__file__).resolve()
SCRIPT_DIR = SCRIPT_FILE.parent
COOKIE_FILE = (
    SCRIPT_DIR.parent / "_helpers" / "simpan_login_siks" / "cookie_siks.json"
)

SCRIPT_NAME = SCRIPT_FILE.stem.removeprefix("src_")
RESULT_DIR = SCRIPT_DIR.parent / "_result_scraping" / f"result_{SCRIPT_DIR.parent.name}" / SCRIPT_DIR.name
CHECKPOINT_FILE = SCRIPT_DIR / f"status_{SCRIPT_FILE.stem}.json"
DETAIL_TABLE_CONTAINER_SELECTOR = (
    "body > div.MuiDialog-root.MuiModal-root.css-126xj0f "
    "> div.MuiDialog-container.MuiDialog-scrollPaper.css-ekeie0 "
    "> div > div > div > div.MuiGrid-root.MuiGrid-container."
    "MuiGrid-spacing-xs-2.css-1wogc6m > div"
)
TABLE_SELECTOR = (
    "table:visible, [role='grid']:visible, "
    ".MuiDataGrid-root:visible, .MuiTable-root:visible"
)
ROW_SELECTOR = (
    "tbody tr:visible, [role='row']:visible, .MuiDataGrid-row:visible, "
    ".MuiTableRow-root:visible"
)
CELL_SELECTOR = (
    "td, [role='cell'], [role='gridcell'], "
    ".MuiDataGrid-cell, .MuiDataGrid-columnHeader, .MuiTableCell-root"
)
OVERLAY_SELECTOR = (
    "[role='dialog']:visible, .modal:visible, "
    "[class*='dialog']:visible, [class*='modal']:visible, "
    ".v-overlay__content:visible"
)
WAIT_TIMEOUT_SECONDS = 30


def _data_rows(table: Locator) -> Locator:
    for selector in (
        "tbody tr:visible",
        ".MuiDataGrid-row:visible",
        "[role='row']:visible:has([role='gridcell'])",
        "[role='row']:visible:has([role='cell'])",
        ".MuiTableRow-root:visible:has(td)",
        ".MuiTableRow-root:visible:has(.MuiTableCell-body)",
    ):
        rows = table.locator(selector)
        if rows.count():
            return rows
    return table.locator("tbody tr:visible")


def _detail_tables(page: Page) -> list[Locator]:
    container = page.locator(DETAIL_TABLE_CONTAINER_SELECTOR).last
    if container.count():
        tables = container.locator(TABLE_SELECTOR)
        if tables.count():
            return [tables.nth(index) for index in range(tables.count())]
        return [container.locator("xpath=.")]

    dialog = page.locator(".MuiDialog-root:visible").last
    if dialog.count():
        data_tables = dialog.locator(TABLE_SELECTOR)
        if data_tables.count():
            return [
                data_tables.nth(index) for index in range(data_tables.count())
            ]
        return [dialog]
    return []


def _candidate_tables(page: Page) -> list[Locator]:
    candidates: list[tuple[int, Locator]] = []

    for table in _detail_tables(page):
        table_state = table.evaluate(
            """element => {
                const root = element.matches('table, [role="grid"]')
                    ? element
                    : element;
                const headerTexts = [
                    ...root.querySelectorAll(
                        'thead th, tr th, [role="columnheader"], ' +
                        '.MuiDataGrid-columnHeader'
                    )
                ].map(cell => cell.innerText.trim().toLowerCase());
                const expected = [
                    'nama', 'no kk', 'nik', 'kelurahan', 'alamat ktp'
                ];
                const matchesHeaders = expected.every(header =>
                    headerTexts.includes(header)
                );
                const rowSelectors = [
                    'tbody tr',
                    '.MuiDataGrid-row',
                    '[role="row"]:has([role="gridcell"])',
                    '[role="row"]:has([role="cell"])',
                    '.MuiTableRow-root:has(td)',
                    '.MuiTableRow-root:has(.MuiTableCell-body)'
                ];
                const rows = rowSelectors
                    .map(selector => [...root.querySelectorAll(selector)])
                    .find(matches => matches.length) || [];
                const hasWideRow = rows.some(row => {
                    const cells = row.querySelectorAll(
                        'td, [role="gridcell"], [role="cell"], ' +
                        '.MuiDataGrid-cell, .MuiTableCell-body, ' +
                        '.MuiTableCell-root'
                    );
                    return cells.length >= 5;
                });
                return {
                    matchesHeaders,
                    hasWideRow,
                    inDialog: Boolean(element.closest('.MuiDialog-root'))
                };
            }"""
        )
        matches_member_headers = table_state["matchesHeaders"]
        is_in_overlay = table_state["inDialog"]
        has_wide_row = table_state["hasWideRow"]
        if not matches_member_headers and not has_wide_row:
            continue
        score = (100 if is_in_overlay else 0) + (
            1000 if matches_member_headers else 0
        )
        score += 10 if has_wide_row else 0
        candidates.append((score, table))

    candidates.sort(key=lambda candidate: candidate[0], reverse=True)
    return [table for _, table in candidates]


def _wait_for_data_table(
    page: Page,
    timeout_seconds: int = WAIT_TIMEOUT_SECONDS,
) -> Locator | None:
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        candidates = _candidate_tables(page)
        if candidates:
            return candidates[0]
        page.wait_for_timeout(500)
    return None


def _read_table_page(table: Locator) -> tuple[list[dict[str, str]], str]:
    rendered_rows = table.evaluate(
        """element => {
            const root = element.matches('table, [role="grid"]')
                ? element
                : element;
            const rowSelectors = [
                'tbody tr',
                '.MuiDataGrid-row',
                '[role="row"]:has([role="gridcell"])',
                '[role="row"]:has([role="cell"])',
                '.MuiTableRow-root:has(td)',
                '.MuiTableRow-root:has(.MuiTableCell-body)'
            ];
            const rows = rowSelectors
                .map(selector => [...root.querySelectorAll(selector)])
                .find(matches => matches.length) || [];
            return rows.map(row => {
                const cells = [
                    ...row.querySelectorAll(
                        ':scope > td, :scope > [role="gridcell"], ' +
                        ':scope > [role="cell"], ' +
                        ':scope > .MuiDataGrid-cell, ' +
                        ':scope > .MuiTableCell-body, ' +
                        ':scope > .MuiTableCell-root'
                    )
                ].filter(cell => cell.getClientRects().length);
                return cells.map(cell => cell.innerText.trim());
            }).filter(cells => cells.length >= 5);
        }"""
    )
    page_data = [
        {
            "NAMA": values[0],
            "NO KK": values[1],
            "NIK": values[2],
            "Desa": values[3],
            "Alamat": values[4],
        }
        for values in rendered_rows
    ]
    row_text = json.dumps(rendered_rows, ensure_ascii=False)
    return page_data, row_text


def _scroll_table_to_top(table: Locator) -> None:
    table.evaluate(
        """element => {
            const root = element.closest('.MuiDialog-root') || element;
            const candidates = new Set();
            const addIfScrollable = candidate => {
                if (!candidate) return;
                const style = getComputedStyle(candidate);
                if (
                    candidate.scrollHeight > candidate.clientHeight + 2 &&
                    /(auto|scroll)/.test(style.overflowY)
                ) {
                    candidates.add(candidate);
                }
            };
            root.querySelectorAll(
                '.MuiDataGrid-virtualScroller, .MuiTableContainer-root, ' +
                '.MuiDialogContent-root'
            ).forEach(addIfScrollable);
            let ancestor = element.parentElement;
            while (ancestor && root.contains(ancestor)) {
                addIfScrollable(ancestor);
                ancestor = ancestor.parentElement;
            }
            candidates.forEach(candidate => { candidate.scrollTop = 0; });
            if (document.scrollingElement) {
                document.scrollingElement.scrollTop = 0;
            }
        }"""
    )


def _scroll_table_to_bottom_and_collect(
    page: Page,
    table: Locator,
) -> tuple[list[dict[str, str]], str]:
    collected: dict[tuple[str, ...], dict[str, str]] = {}
    max_scroll_steps = 1000

    for _ in range(max_scroll_steps):
        current_rows, _ = _read_table_page(table)
        for row in current_rows:
            key = tuple(row.values())
            collected[key] = row

        scroll_state = table.evaluate(
            """element => {
                const root = element.closest('.MuiDialog-root') || element;
                const candidates = [];
                const seen = new Set();
                const addIfScrollable = candidate => {
                    if (!candidate || seen.has(candidate)) return;
                    seen.add(candidate);
                    const style = getComputedStyle(candidate);
                    if (
                        candidate.scrollHeight > candidate.clientHeight + 2 &&
                        /(auto|scroll)/.test(style.overflowY)
                    ) {
                        candidates.push(candidate);
                    }
                };

                root.querySelectorAll(
                    '.MuiDataGrid-virtualScroller, .MuiTableContainer-root, ' +
                    '.MuiDialogContent-root'
                ).forEach(addIfScrollable);
                let ancestor = element.parentElement;
                while (ancestor && root.contains(ancestor)) {
                    addIfScrollable(ancestor);
                    ancestor = ancestor.parentElement;
                }
                const documentScroller = document.scrollingElement;
                if (
                    documentScroller &&
                    documentScroller.scrollHeight >
                        documentScroller.clientHeight + 2
                ) {
                    candidates.push(documentScroller);
                }

                const target = candidates.find(candidate =>
                    candidate.scrollTop + candidate.clientHeight <
                    candidate.scrollHeight - 2
                );
                if (!target) return {atBottom: true};

                const before = target.scrollTop;
                const firstRow = element.querySelector(
                    'tbody tr, .MuiDataGrid-row, .MuiTableRow-root, ' +
                    '[role="row"]:has([role="gridcell"])'
                );
                const rowHeight = firstRow
                    ? firstRow.getBoundingClientRect().height
                    : target.clientHeight * 0.3;
                const step = Math.max(
                    1,
                    Math.floor(Math.min(
                        target.clientHeight * 0.45,
                        rowHeight * 0.75
                    ))
                );
                target.scrollTop = Math.min(
                    before + step,
                    target.scrollHeight - target.clientHeight
                );
                return {
                    atBottom: false,
                    moved: target.scrollTop > before
                };
            }"""
        )
        if scroll_state["atBottom"]:
            break
        if not scroll_state["moved"]:
            page.wait_for_timeout(250)
        else:
            page.wait_for_timeout(300)
    else:
        raise RuntimeError(
            "Scroll tabel tidak mencapai bagian bawah setelah 1000 langkah."
        )

    final_rows, final_signature = _read_table_page(table)
    for row in final_rows:
        collected[tuple(row.values())] = row
    print(f"  Terkumpul {len(collected)} baris setelah scroll sampai bawah.")
    return list(collected.values()), final_signature


def _wait_for_page_change(
    page: Page,
    previous_rows: str,
    timeout_seconds: int = WAIT_TIMEOUT_SECONDS,
) -> Locator | None:
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        for table in _candidate_tables(page):
            _, current_rows = _read_table_page(table)
            if current_rows and current_rows != previous_rows:
                return table
        page.wait_for_timeout(300)
    return None


def _save_result(rows: list[dict[str, str]]) -> str:
    timestamp = time.strftime("%Y%m%d_%H%M%S")
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


def _load_checkpoint() -> dict[str, list[dict[str, str]]]:
    if not CHECKPOINT_FILE.exists():
        return {}

    with CHECKPOINT_FILE.open(encoding="utf-8") as checkpoint_file:
        checkpoint = json.load(checkpoint_file)

    if not isinstance(checkpoint, dict) or checkpoint.get("version") != 1:
        raise ValueError(
            f"Format checkpoint tidak dikenal: {CHECKPOINT_FILE}. "
            "Periksa atau pindahkan file checkpoint sebelum melanjutkan."
        )

    groups = checkpoint.get("groups")
    if not isinstance(groups, dict):
        raise ValueError(
            f"Data kelompok pada checkpoint tidak valid: {CHECKPOINT_FILE}."
        )

    completed: dict[str, list[dict[str, str]]] = {}
    for group_key, group_record in groups.items():
        if (
            not isinstance(group_key, str)
            or not isinstance(group_record, dict)
            or group_record.get("status") != "success"
            or not isinstance(group_record.get("rows"), list)
            or any(not isinstance(row, dict) for row in group_record["rows"])
        ):
            raise ValueError(
                f"Checkpoint kelompok '{group_key}' tidak valid: "
                f"{CHECKPOINT_FILE}."
            )
        completed[group_key] = group_record["rows"]
    return completed


def _save_checkpoint(
    completed_groups: dict[str, list[dict[str, str]]],
) -> None:
    checkpoint = {
        "version": 1,
        "groups": {
            group_key: {"status": "success", "rows": rows}
            for group_key, rows in completed_groups.items()
        },
    }
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            "w",
            encoding="utf-8",
            dir=SCRIPT_DIR,
            prefix=".scraper_p2k2_detail_status_",
            suffix=".tmp",
            delete=False,
        ) as temp_file:
            temp_path = Path(temp_file.name)
            json.dump(checkpoint, temp_file, ensure_ascii=False, indent=2)
            temp_file.write("\n")
        temp_path.replace(CHECKPOINT_FILE)
    finally:
        if temp_path and temp_path.exists():
            temp_path.unlink()


def _find_next_button(page: Page) -> Locator | None:
    next_selector = (
        ".MuiTablePagination-actions button[aria-label*='next page' i], "
        "button[aria-label='Go to next page'], "
        "button.next, a.next, "
        "button[aria-label='Next'], button[aria-label*='next' i], "
        "button[aria-label*='next page' i], "
        "button[title*='next page' i], "
        "button[title*='next' i], a[title*='next' i], "
        ".pagination li:last-child a, "
        ".mdi-chevron-right, .fa-chevron-right"
    )
    dialog = page.locator(".MuiDialog-root:visible").last
    if dialog.count():
        button = dialog.locator(next_selector).last
        if button.count():
            return button

    button = page.locator(next_selector).last
    return button if button.count() else None


def _find_group_table(page: Page) -> Locator | None:
    tables = page.locator(
        "table:visible, .MuiDataGrid-root:visible, .MuiTable-root:visible"
    )
    for index in range(tables.count()):
        table = tables.nth(index)
        in_dialog = table.evaluate(
            "element => Boolean(element.closest('.MuiDialog-root'))"
        )
        if in_dialog:
            continue
        rows = _data_rows(table)
        if rows.count():
            return table
    return None


def _find_group_detail_button(row: Locator) -> Locator | None:
    action_cell = row.locator("td, [role='cell'], .MuiTableCell-root").last
    button = action_cell.locator(
        "a.btn-info, button.btn-info, "
        "a:has(.fa-info-circle), button:has(.fa-info-circle), "
        "a:has(.mdi-information), button:has(.mdi-information), "
        "a:has(.mdi-eye), button:has(.mdi-eye), "
        "a:has(.fa-eye), button:has(.fa-eye), "
        "[title*='detail' i], [aria-label*='detail' i], "
        "[title*='informasi' i], [aria-label*='informasi' i]"
    ).first
    if button.count():
        return button

    return None


def _group_name_from_row(row: Locator) -> str:
    group_name = row.evaluate(
        """row => {
            const root = row.closest(
                'table, [role="grid"], .MuiDataGrid-root, .MuiTable-root'
            );
            if (!root) return '';

            const headers = [
                ...root.querySelectorAll(
                    'thead th, [role="columnheader"], ' +
                    '.MuiDataGrid-columnHeader'
                )
            ].map(header => header.innerText.trim().toLowerCase());
            const nameIndex = headers.findIndex(header =>
                header.includes('nama') && header.includes('kelompok')
            );
            const cells = [
                ...row.querySelectorAll(
                    ':scope > td, :scope > [role="cell"], ' +
                    ':scope > [role="gridcell"], ' +
                    ':scope > .MuiTableCell-root'
                )
            ];
            return nameIndex >= 0 && cells[nameIndex]
                ? cells[nameIndex].innerText.trim()
                : '';
        }"""
    )
    if not group_name:
        raise RuntimeError(
            "Nama kelompok tidak ditemukan pada kolom tabel daftar kelompok."
        )
    return group_name


def _scroll_group_list(table: Locator) -> dict[str, bool]:
    return table.evaluate(
        """element => {
            const candidates = [];
            const seen = new Set();
            const addIfScrollable = candidate => {
                if (!candidate || seen.has(candidate)) return;
                seen.add(candidate);
                const style = getComputedStyle(candidate);
                if (
                    candidate.scrollHeight > candidate.clientHeight + 2 &&
                    /(auto|scroll)/.test(style.overflowY)
                ) {
                    candidates.push(candidate);
                }
            };
            let ancestor = element.parentElement;
            while (ancestor && ancestor !== document.body) {
                addIfScrollable(ancestor);
                ancestor = ancestor.parentElement;
            }
            const documentScroller = document.scrollingElement;
            if (
                documentScroller &&
                documentScroller.scrollHeight > documentScroller.clientHeight + 2
            ) {
                candidates.push(documentScroller);
            }

            const target = candidates.find(candidate =>
                candidate.scrollTop + candidate.clientHeight <
                candidate.scrollHeight - 2
            );
            if (!target) return {atBottom: true, moved: false};

            const row = element.querySelector(
                'tbody tr, .MuiDataGrid-row, .MuiTableRow-root'
            );
            const rowHeight = row
                ? row.getBoundingClientRect().height
                : target.clientHeight * 0.3;
            const step = Math.max(
                1,
                Math.floor(Math.min(target.clientHeight * 0.45, rowHeight * 0.75))
            );
            const before = target.scrollTop;
            target.scrollTop = Math.min(
                before + step,
                target.scrollHeight - target.clientHeight
            );
            return {
                atBottom: false,
                moved: target.scrollTop > before
            };
        }"""
    )


def _next_group_page_button(page: Page) -> Locator | None:
    selector = (
        ".MuiTablePagination-actions button[aria-label*='next page' i], "
        "button[aria-label='Go to next page'], "
        "button[aria-label*='next page' i], "
        ".pagination li:last-child a"
    )
    buttons = page.locator(selector)
    for index in range(buttons.count() - 1, -1, -1):
        button = buttons.nth(index)
        in_dialog = button.evaluate(
            "element => Boolean(element.closest('.MuiDialog-root'))"
        )
        if not in_dialog:
            return button
    return None


def _wait_for_group_page_change(
    page: Page,
    old_first_row: str,
) -> Locator | None:
    deadline = time.monotonic() + WAIT_TIMEOUT_SECONDS
    while time.monotonic() < deadline:
        table = _find_group_table(page)
        if table is not None:
            rows = _data_rows(table)
            if rows.count() and rows.first.inner_text().strip() != old_first_row:
                return table
        page.wait_for_timeout(300)
    return None


def _scrape_open_group(
    page: Page,
    all_data: list[dict[str, str]],
    group_data: list[dict[str, str]],
    group_name: str,
) -> None:
    table = _wait_for_data_table(page)
    if table is None:
        raise RuntimeError(
            "Tabel anggota tidak ditemukan dalam detail kelompok yang dibuka."
        )

    member_page = 1
    while True:
        print(f"  Mengambil anggota halaman {member_page}...")
        _scroll_table_to_top(table)
        page_data, previous_rows = _scroll_table_to_bottom_and_collect(page, table)
        for member in page_data:
            member["NAMA KELOMPOK"] = group_name
        group_data.extend(page_data)

        next_button = _find_next_button(page)
        if next_button is None or _is_disabled(next_button):
            break

        next_button.click()
        table = _wait_for_page_change(page, previous_rows)
        if table is None:
            raise RuntimeError(
                "Tabel anggota tidak berubah setelah tombol Next diklik. "
                f"Status kelompok selesai tetap tersimpan di {CHECKPOINT_FILE}."
            )
        member_page += 1

    if not group_data:
        raise RuntimeError(
            "Tidak ada anggota yang berhasil dibaca. Kelompok tidak akan "
            "ditandai selesai di checkpoint JSON."
        )


def _close_group_detail(page: Page) -> None:
    dialog = page.locator(".MuiDialog-root:visible").last
    if not dialog.count():
        return

    close_button = dialog.locator(
        "button[aria-label*='close' i], "
        "button[title*='close' i], "
        ".MuiDialogTitle-root button"
    ).last
    if close_button.count():
        close_button.click()
    else:
        page.keyboard.press("Escape")

    try:
        dialog.wait_for(state="hidden", timeout=5000)
    except PlaywrightTimeoutError as error:
        raise RuntimeError(
            "Dialog detail kelompok tidak tertutup; tidak dapat melanjutkan "
            "ke kelompok berikutnya."
        ) from error


def _is_disabled(button: Locator) -> bool:
    button_class = button.get_attribute("class") or ""
    parent_class = button.locator("xpath=..").get_attribute("class") or ""
    return (
        button.is_disabled()
        or button.get_attribute("aria-disabled") == "true"
        or "disabled" in button_class.split()
        or "disabled" in parent_class.split()
    )


def _open_homepage(page: Page) -> None:
    try:
        page.goto(BASE_URL, wait_until="commit", timeout=60000)
    except PlaywrightTimeoutError as error:
        if page.url in ("", "about:blank"):
            raise RuntimeError(
                "Beranda SIKS belum berhasil dibuka. Periksa koneksi internet "
                "lalu jalankan ulang skrip."
            ) from error
        print(
            "[PERINGATAN] Navigasi melewati batas waktu, tetapi browser sudah "
            f"berpindah ke {page.url}. Tetap mencoba melanjutkan."
        )

    try:
        page.wait_for_load_state("domcontentloaded", timeout=15000)
    except PlaywrightTimeoutError:
        print(
            "[PERINGATAN] Halaman belum menyelesaikan pemuatan awal. "
            "Browser tetap dibuka agar halaman bisa dilanjutkan secara manual."
        )


def scrape_p2k2_details():
    if not COOKIE_FILE.is_file():
        raise FileNotFoundError(f"File sesi login tidak ditemukan: {COOKIE_FILE}")

    completed_groups = _load_checkpoint()
    all_data = [
        row
        for group_rows in completed_groups.values()
        for row in group_rows
    ]
    if completed_groups:
        print(
            f"Checkpoint JSON memuat {len(completed_groups)} kelompok selesai "
            f"({len(all_data)} baris anggota). Kelompok tersebut akan dilewati."
        )

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False)
        try:
            context = browser.new_context(storage_state=str(COOKIE_FILE))
            page = context.new_page()
            _open_homepage(page)

            print(
                "[PETUNJUK] Sesi dari cookie_siks.json dimuat. Login bila "
                "diminta, lalu buka halaman daftar kelompok P2K2."
            )
            input(
                "Tekan ENTER setelah halaman daftar kelompok P2K2 terbuka..."
            )

            group_table = _find_group_table(page)
            if group_table is None:
                raise RuntimeError(
                    "Tabel daftar kelompok tidak ditemukan setelah Enter. "
                    "Pastikan halaman daftar P2K2 sedang terbuka."
                )

            group_page = 1
            while True:
                print(f"Memproses daftar kelompok halaman {group_page}...")
                processed_groups: set[str] = set()
                stuck_scrolls = 0
                while True:
                    rows = _data_rows(group_table)
                    for row_index in range(rows.count()):
                        row = rows.nth(row_index)
                        group_key = " ".join(row.inner_text().split())
                        if not group_key or group_key in processed_groups:
                            continue
                        group_name = _group_name_from_row(row)
                        if group_key in completed_groups:
                            for member in completed_groups[group_key]:
                                member["NAMA KELOMPOK"] = group_name
                            _save_checkpoint(completed_groups)
                            print(f"Melewati kelompok yang sudah selesai: {group_key}")
                            processed_groups.add(group_key)
                            continue

                        detail_button = _find_group_detail_button(row)
                        if detail_button is None:
                            print(
                                "Tombol detail tidak dikenali otomatis untuk "
                                f"kelompok: {group_key}"
                            )
                            input(
                                "Klik tombol detail kelompok ini di browser, "
                                "lalu tekan ENTER..."
                            )
                        else:
                            detail_button.click()

                        print(f"Membuka kelompok: {group_key}")
                        page.locator(".MuiDialog-root:visible").wait_for(
                            state="visible",
                            timeout=WAIT_TIMEOUT_SECONDS * 1000,
                        )
                        group_data: list[dict[str, str]] = []
                        _scrape_open_group(
                            page, all_data, group_data, group_name
                        )
                        completed_groups[group_key] = group_data
                        _save_checkpoint(completed_groups)
                        all_data.extend(group_data)
                        _close_group_detail(page)
                        processed_groups.add(group_key)
                        print(
                            f"Kelompok selesai: {len(group_data)} anggota. "
                            f"Checkpoint tersimpan di {CHECKPOINT_FILE}"
                        )
                        group_table = _find_group_table(page)
                        if group_table is None:
                            raise RuntimeError(
                                "Tabel daftar kelompok hilang setelah dialog "
                                "detail ditutup."
                            )
                        rows = _data_rows(group_table)

                    scroll_result = _scroll_group_list(group_table)
                    if scroll_result["atBottom"]:
                        break
                    if not scroll_result["moved"]:
                        stuck_scrolls += 1
                        if stuck_scrolls >= 5:
                            raise RuntimeError(
                                "Daftar kelompok tidak dapat digulir lebih "
                                "jauh, tetapi belum mencapai posisi bawah."
                            )
                    else:
                        stuck_scrolls = 0
                        page.wait_for_timeout(300)

                next_group_button = _next_group_page_button(page)
                if (
                    next_group_button is None
                    or _is_disabled(next_group_button)
                ):
                    print("Semua halaman daftar kelompok sudah diproses.")
                    break

                old_first_group = _data_rows(group_table).first.inner_text().strip()
                next_group_button.click()
                group_table = _wait_for_group_page_change(page, old_first_group)
                if group_table is None:
                    raise RuntimeError(
                        "Daftar kelompok tidak berubah setelah tombol Next "
                        "diklik. Hasil anggota yang sudah selesai tetap "
                        f"tersimpan di {CHECKPOINT_FILE}."
                    )
                group_page += 1

            if all_data:
                input(
                    "Proses selesai. Tekan ENTER untuk menyimpan hasil XLSX..."
                )
                result_path = _save_result(all_data)
                print(f"Sukses! {len(all_data)} baris disimpan ke {result_path}.")
            else:
                print(
                    "Tidak ada baris yang bisa diekspor. Pastikan tabel yang "
                    "terbuka adalah tabel anggota dan memiliki minimal 5 kolom."
                )
                input("Browser tetap terbuka. Tekan ENTER untuk menutup...")
        finally:
            browser.close()


if __name__ == "__main__":
    scrape_p2k2_details()
