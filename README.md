# Conference Registration & ABA KHQR (Excel Database + GitHub Pages)

Web registration app with dynamic ABA KHQR payment, Telegram confirmation, PIN-protected Admin dashboard, and live Excel database storage (`registrations.xlsx`).

## Files

- `index.html` — Frontend registration & Admin dashboard (works on GitHub Pages and with `server.py`).
- `server.py` — Local Python HTTP + REST API server storing registrations directly in `registrations.xlsx`.
- `registrations.xlsx` — Excel database file stored locally on `H:\Registration` (excluded from Git via `.gitignore` to protect attendee data).
- `start_server.bat` — Double-click launcher for Windows.
- `requirements.txt` — Python dependencies (`openpyxl`).

## 1. Run Locally with Live Excel Database (`H:\Registration\registrations.xlsx`)

1. Install Python dependency (if not already installed):
   ```powershell
   pip install -r requirements.txt
   ```
2. Double-click `start_server.bat` (or run `python server.py`), then open:
   - Local PC: `http://localhost:8000`
   - Same Wi-Fi / Office LAN: shown in the terminal (e.g. `http://192.168.x.x:8000`)

Every time an attendee clicks **បន្ត → ទូទាត់ ABA**, their registration is automatically saved into `registrations.xlsx` with status `Pending`.

> **Tip:** Close `registrations.xlsx` in Microsoft Excel before saving/updating from the web Admin panel so Excel does not lock the file.

## 2. Admin Access & Default PIN

- Click the **Admin (Excel)** tab (or open `?admin=1`).
- Default PIN: `1234`
- **Before publishing**, change the PIN in:
  1. `index.html` (`CFG.adminPin`)
  2. `server.py` (`ADMIN_PIN` or set environment variable `$env:ADMIN_PIN="your-new-pin"` before starting `server.py`)

## 3. Publish to GitHub Pages

1. Create a repository on GitHub (do **not** commit `registrations.xlsx`; `.gitignore` already excludes `*.xlsx`).
2. Push this folder to GitHub:
   ```powershell
   git remote add origin https://github.com/<USERNAME>/<REPO>.git
   git push -u origin main
   ```
3. In GitHub: **Settings → Pages → Build and deployment → Deploy from a branch → `main` / `/ (root)`**.
4. Your public registration form will be live at `https://<USERNAME>.github.io/<REPO>/`.

### Syncing GitHub Pages Registrations to `H:\Registration\registrations.xlsx`

Because GitHub Pages is static hosting, public visitors cannot write directly to your `H:\` drive unless you connect a tunnel URL:
- **Simple workflow (Telegram):** Attendees submit on GitHub Pages, pay via ABA KHQR, and send their message + slip to Telegram. Open `http://localhost:8000` → **Admin (Excel)** → paste the Telegram message(s) → click **បន្ថែម / បញ្ជាក់ការបង់ប្រាក់ (Paid)** to save straight into `H:\Registration\registrations.xlsx`.
- **Direct live sync (optional):** Expose `server.py` with Cloudflare Tunnel (`cloudflared tunnel --url http://localhost:8000`), then paste that tunnel URL inside **Admin → កំណត់ការភ្ជាប់ Server** (or set `CFG.apiUrl` in `index.html`).
