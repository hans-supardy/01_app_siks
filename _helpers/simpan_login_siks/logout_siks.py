import os
import re

from playwright.sync_api import sync_playwright
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

COOKIE_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "cookie_siks.json",
)
SIKS_URL = "https://siks.kemensos.go.id/"
LOGIN_PLACEHOLDER = "Masukkan NIK atau NIP Anda..."
LOGOUT_PATTERN = re.compile(r"\b(keluar|log\s*out|sign\s*out)\b", re.IGNORECASE)
PROFILE_PATTERN = re.compile(
    r"\b(akun|profil|account|profile|user|person)\b|account_circle|account_box",
    re.IGNORECASE,
)


def login_page_visible(page):
    return page.get_by_placeholder(LOGIN_PLACEHOLDER).is_visible()


def cari_tombol_logout(page):
    for role in ("menuitem", "button", "link"):
        matches = page.get_by_role(role, name=LOGOUT_PATTERN)
        visible_matches = [
            matches.nth(index)
            for index in range(matches.count())
            if matches.nth(index).is_visible()
        ]
        if len(visible_matches) > 1:
            raise RuntimeError(
                "Ditemukan lebih dari satu kontrol Keluar/Logout yang terlihat."
            )
        if visible_matches:
            return visible_matches[0]
    return None


def buka_menu_akun(page):
    controls = page.locator(
        "button, [role='button'], [aria-haspopup='menu']"
    )
    candidates = []
    for index in range(controls.count()):
        control = controls.nth(index)
        if not control.is_visible():
            continue
        metadata = control.evaluate(
            """element => [
                element.innerText,
                element.getAttribute('aria-label'),
                element.getAttribute('title'),
                ...[...element.querySelectorAll(
                    '[data-testid], [aria-label], [title], .material-icons'
                )].flatMap(child => [
                    child.innerText,
                    child.getAttribute('data-testid'),
                    child.getAttribute('aria-label'),
                    child.getAttribute('title')
                ])
            ].filter(Boolean).join(' ')"""
        )
        if PROFILE_PATTERN.search(metadata):
            candidates.append(control)

    if len(candidates) != 1:
        raise RuntimeError(
            "Menu akun/profil tidak dapat ditentukan secara tunggal "
            f"(ditemukan {len(candidates)} kandidat)."
        )
    candidates[0].click()


def logout_siks():
    if not os.path.isfile(COOKIE_FILE):
        raise FileNotFoundError(
            f"Sesi tersimpan tidak ditemukan: {COOKIE_FILE}"
        )

    with sync_playwright() as playwright:
        print("[INFO] Membuka SIKS dengan sesi tersimpan...")
        try:
            browser = playwright.chromium.launch(headless=False)
        except Exception as error:
            raise RuntimeError(
                "Gagal membuka browser Chromium. Pastikan browser Playwright "
                "sudah terpasang dengan perintah: "
                "python -m playwright install chromium"
            ) from error
        try:
            context = browser.new_context(storage_state=COOKIE_FILE)
            page = context.new_page()
            page.goto(
                SIKS_URL,
                wait_until="domcontentloaded",
                timeout=60000,
            )

            if not login_page_visible(page):
                tombol_logout = cari_tombol_logout(page)
                if tombol_logout is None:
                    print("[INFO] Membuka menu akun/profil...")
                    buka_menu_akun(page)
                    tombol_logout = cari_tombol_logout(page)
                if tombol_logout is None:
                    raise RuntimeError(
                        "Tombol Keluar/Logout tidak ditemukan. "
                        "Sesi lokal tetap dipertahankan."
                    )

                tombol_logout.click()
                dialogs = page.locator(
                    "[role='dialog']:visible, .MuiDialog-root:visible"
                )
                for index in range(dialogs.count()):
                    dialog = dialogs.nth(index)
                    konfirmasi = dialog.get_by_role(
                        "button",
                        name=re.compile(
                            r"^(ya|ya,?\s*(keluar|log\s*out)|"
                            r"(keluar|log\s*out|sign\s*out|konfirmasi))$",
                            re.IGNORECASE,
                        ),
                    )
                    tombol_konfirmasi = [
                        konfirmasi.nth(button_index)
                        for button_index in range(konfirmasi.count())
                        if konfirmasi.nth(button_index).is_visible()
                    ]
                    if len(tombol_konfirmasi) > 1:
                        raise RuntimeError(
                            "Konfirmasi Logout memiliki beberapa tombol "
                            "yang cocok."
                        )
                    if tombol_konfirmasi:
                        tombol_konfirmasi[0].click()
                        break

                try:
                    page.get_by_placeholder(LOGIN_PLACEHOLDER).wait_for(
                        state="visible",
                        timeout=20000,
                    )
                except PlaywrightTimeoutError as error:
                    raise RuntimeError(
                        "Tombol Logout sudah ditekan, tetapi halaman login "
                        "tidak terdeteksi. File cookie dipertahankan."
                    ) from error

            context.clear_cookies()
            page.evaluate(
                "() => { localStorage.clear(); sessionStorage.clear(); }"
            )
            os.remove(COOKIE_FILE)
            print(
                "[SUKSES] Halaman login terdeteksi; sesi lokal dan "
                "cookie_siks.json telah dihapus."
            )
        finally:
            browser.close()


if __name__ == "__main__":
    logout_siks()
