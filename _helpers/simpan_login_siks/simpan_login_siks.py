import os

from playwright.sync_api import sync_playwright

COOKIE_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "cookie_siks.json"
)

# Ganti dengan NIK atau username Anda yang sebenarnya
USERNAME_ANDA = "3527091405870002"

with sync_playwright() as p:
  print("[INFO] Membuka browser otomatis untuk merekam login...")
  try:
    browser = p.chromium.launch(headless=False)
  except Exception as error:
    raise RuntimeError(
        "Gagal membuka browser Chromium. Pastikan browser Playwright sudah "
        "terpasang dengan perintah: python -m playwright install chromium"
    ) from error
  try:
    page = browser.new_page()

    print("[INFO] Menuju ke halaman login SIKS-NG...")
    page.goto("https://siks.kemensos.go.id/", wait_until="commit", timeout=60000)

    # Menunggu elemen input muncul
    input_nik = page.get_by_placeholder("Masukkan NIK atau NIP Anda...")
    input_nik.wait_for(state="visible", timeout=60000)

    # Mengisi username/NIK secara otomatis
    print("[INFO] Mengisi username secara otomatis...")
    input_nik.fill(USERNAME_ANDA)

    print("\n[PETUNJUK] Login secara manual di jendela browser:")
    print("1. Username sudah terisi, silahkan lengkapi password.")
    print("2. Selesaikan CAPTCHA dan OTP jika diminta.")
    print("3. Setelah login berhasil, tekan ENTER di terminal ini.")
    input(
        "\n[TUNGGU] Tekan ENTER setelah berhasil login untuk menyimpan sesi..."
    )

    print("\n[PROSES] Menyimpan sesi login...")
    page.context.storage_state(path=COOKIE_FILE)
    print("[SUKSES] Sesi disimpan ke file 'cookie_siks.json'.")
  finally:
    browser.close()

  print("[SELESAI] Sesi aman disimpan.")
