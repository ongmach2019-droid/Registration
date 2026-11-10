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

## 3. AI Card Scan (Gemini Vision + QR/OCR Fallback)

On the Registration tab, attendees can upload or capture photos (front/back, up to 3 images) of a Dentist ID, Student ID, or Business Card to auto-fill Name, Phone, Telegram, Organization, and Role:

- **With Gemini API Key (best accuracy for Khmer + English):**
  - **Local Server (`server.py`):** Open **Admin (Excel) → ⚙️ កំណត់ការភ្ជាប់ Server & Gemini AI** and save your Google AI Studio API key (`GEMINI_API_KEY` is stored in local `.env`, excluded from Git). Public attendees using `server.py` can scan without seeing or entering any key.
  - **GitHub Pages (browser-only):** Enter a Gemini API key under **🔑 កំណត់ Gemini AI Key** (stored in your browser's `localStorage` only).
- **Without API Key (Offline QR + OCR Fallback):** Automatically reads QR contact codes (`vCard` / `MECARD` / `t.me/` / `tel:`) and runs client-side OCR (`Tesseract.js`) to extract fields and classify roles.

## 4. Publish to GitHub Pages

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

## 5. Deploy on Render (`render.com`)

`render.yaml` is included for one-click Python Web Service deployment on Render:

1. Push the project files (`index.html`, `server.py`, `requirements.txt`, `render.yaml`, `README.md`, `.gitignore`) to GitHub (**never** upload `.env` or `registrations.xlsx`).
2. In Render Dashboard: **New + → Blueprint** (or **Web Service**) → connect your GitHub repository (`Registration`).
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `python server.py`
3. Add Environment Variables in Render (**Environment** tab, never commit keys to Git):
   - `ADMIN_PIN` — your secret Admin PIN (do not leave `1234`)
   - `GEMINI_API_KEY` — your Google AI Studio Gemini API key
   - `GEMINI_MODEL` — `gemini-3-flash-preview`
   - `TELEGRAM_BOT_TOKEN` — your Telegram Bot token from `@BotFather`
   - `TELEGRAM_CHAT_ID` — your Telegram Chat ID (send `/start` to your bot, then use **Admin → ⚙️ → 🔍 រក Chat ID ពី Bot**)
4. **Data persistence on Render:** Render's Free plan uses temporary storage that resets when the service restarts or redeploys. Use the **✈ ផ្ញើ Excel Backup ទៅ Telegram** and **📥 ទាញយក Excel (.xlsx)** buttons in Admin (or attach a Render Persistent Disk and set `REGISTRATION_DB=/data/registrations.xlsx`) to keep your Excel data safe.
