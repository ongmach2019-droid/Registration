"""Local web server and Excel database backend for Conference Registration.

Serves index.html and persists registration records directly into
registrations.xlsx in the same directory using openpyxl.
"""

from __future__ import annotations

import argparse
from datetime import date, datetime
import hmac
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import os
from pathlib import Path
import re
import socket
import sys
import tempfile
import threading
import time
from typing import Any
import urllib.error
from urllib.parse import quote, unquote, urlparse
import urllib.request

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent
ENV_PATH = BASE_DIR / ".env"
SYNC_STATE_PATH = BASE_DIR / ".sync_state.json"
DB_PATH = Path(os.environ.get("REGISTRATION_DB", BASE_DIR / "registrations.xlsx"))
HTML_PATH = BASE_DIR / "index.html"
HOST = os.environ.get("HOST", "0.0.0.0")
PORT = int(os.environ.get("PORT", "8000"))
MAX_BODY_BYTES = 2 * 1024 * 1024
MAX_SCAN_BODY_BYTES = 8 * 1024 * 1024
DEFAULT_GEMINI_MODEL = "gemini-3-flash-preview"
FALLBACK_GEMINI_MODELS = ("gemini-3-flash-preview", "gemini-flash-latest", "gemini-3.1-flash-lite")

CARD_SCAN_PROMPT = (
    "The attached image(s) are photos (possibly front and back) of one person's ID card, "
    "dentist/professional card, student card, or business card (in Khmer and/or English). "
    "Extract the holder's details accurately. Return ONLY JSON matching the schema. "
    "Rules: 'role' is 'Dentist' if the card shows Dr., DDS, DMD, ទន្តបណ្ឌិត, dentist, dental clinic, or dental licence; "
    "'Technician' if it shows dental technician, dental lab, ceramist, or ជាងធ្មេញ; "
    "'Student' if it shows student, university, school ID, or និស្សិត; otherwise 'Unknown'. "
    "Include 'tg' only if a Telegram username/link is shown (e.g. @username). "
    "Use empty string for any field not clearly visible; never guess."
)


def _load_env_file() -> None:
    """Load key=value pairs from .env into os.environ without overriding existing vars."""
    if not ENV_PATH.exists():
        return
    try:
        for raw_line in ENV_PATH.read_text(encoding="utf-8").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, val = line.split("=", 1)
            key = key.strip()
            val = val.strip().strip('"').strip("'")
            if key:
                os.environ.setdefault(key, val)
    except OSError:
        pass


_load_env_file()
ADMIN_PIN = os.environ.get("ADMIN_PIN", "1234")

SHEET_NAME = "Registrations"
COLUMNS = [
    ("ref", "Ref", 14),
    ("name", "Name", 24),
    ("phone", "Phone", 16),
    ("tg", "Telegram", 18),
    ("org", "Organization", 26),
    ("type", "Type", 14),
    ("ticket", "Ticket", 26),
    ("amount", "Amount (USD)", 15),
    ("paid", "Paid", 12),
    ("date", "Date", 14),
]

HEADER_ALIASES = {
    "ref": "ref",
    "name": "name",
    "phone": "phone",
    "telegram": "tg",
    "tg": "tg",
    "organization": "org",
    "org": "org",
    "type": "type",
    "ticket": "ticket",
    "amount (usd)": "amount",
    "amount": "amount",
    "paid": "paid",
    "date": "date",
}

PRICING: dict[str, dict[str, float]] = {
    "Dentist": {"From Framework to Final Glaze": 35, "Main Conference": 50, "Hands-On": 80, "Main + Hands-On": 120},
    "Student": {"From Framework to Final Glaze": 35, "Main Conference": 20, "Hands-On": 40, "Main + Hands-On": 55},
    "Technician": {"From Framework to Final Glaze": 35, "Full-day hands-on course": 150},
}

DB_LOCK = threading.Lock()


class ExcelLockedError(Exception):
    """Raised when registrations.xlsx is open/locked by another program (e.g. Excel)."""


class RegistrationConflictError(Exception):
    """Raised when attempting an invalid modification to an existing registration."""


class AiConfigError(Exception):
    """Raised when Gemini AI API key is missing or invalidly configured."""


class AiUpstreamError(Exception):
    """Raised when the upstream Gemini API request fails."""


class TelegramConfigError(Exception):
    """Raised when Telegram bot token or chat ID is missing."""


def _get_admin_pin() -> str:
    return os.environ.get("ADMIN_PIN", ADMIN_PIN).strip() or "1234"


def _get_gemini_key() -> str:
    return os.environ.get("GEMINI_API_KEY", "").strip()


def _get_gemini_model() -> str:
    model = os.environ.get("GEMINI_MODEL", DEFAULT_GEMINI_MODEL).strip()
    if not re.fullmatch(r"[A-Za-z0-9._-]+", model):
        return DEFAULT_GEMINI_MODEL
    return model


def _get_telegram_token() -> str:
    return os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()


def _get_telegram_chat_id() -> str:
    return os.environ.get("TELEGRAM_CHAT_ID", "").strip()


def _get_render_url() -> str:
    return os.environ.get("RENDER_URL", "").strip().rstrip("/")


def _get_render_pin() -> str:
    return os.environ.get("RENDER_ADMIN_PIN", "").strip() or _get_admin_pin()


def _get_sync_interval() -> int:
    try:
        return max(5, int(os.environ.get("SYNC_INTERVAL", "10")))
    except ValueError:
        return 10


def _mask_secret(value: str) -> str:
    if not value:
        return ""
    if len(value) <= 8:
        return "*" * len(value)
    return f"{value[:4]}...{value[-4:]}"


def _save_env_settings(updates: dict[str, str | None]) -> None:
    """Persist runtime settings into local .env and os.environ."""
    entries: dict[str, str] = {}
    if ENV_PATH.exists():
        for raw_line in ENV_PATH.read_text(encoding="utf-8").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            entries[k.strip()] = v.strip()

    for env_key, raw_val in updates.items():
        if raw_val is None:
            continue
        cleaned = str(raw_val).strip().replace("\n", "").replace("\r", "")
        if env_key == "GEMINI_MODEL":
            cleaned = cleaned or DEFAULT_GEMINI_MODEL
            if not re.fullmatch(r"[A-Za-z0-9._-]+", cleaned):
                raise ValueError("Invalid Gemini model name.")
        if cleaned:
            entries[env_key] = cleaned
            os.environ[env_key] = cleaned
        else:
            entries.pop(env_key, None)
            os.environ.pop(env_key, None)

    lines = [f"{k}={v}" for k, v in entries.items()]
    ENV_PATH.write_text("\n".join(lines) + ("\n" if lines else ""), encoding="utf-8")


def _telegram_request(method: str, payload: dict[str, Any] | None = None) -> dict[str, Any]:
    token = _get_telegram_token()
    if not token:
        raise TelegramConfigError("TELEGRAM_BOT_TOKEN is not configured.")
    url = f"https://api.telegram.org/bot{token}/{method}"
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    headers = {"Content-Type": "application/json"} if data is not None else {}
    req = urllib.request.Request(url, data=data, headers=headers, method="POST" if data else "GET")
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        raise AiUpstreamError(f"Telegram API error ({exc.code}): {detail}") from exc
    except (urllib.error.URLError, TimeoutError) as exc:
        raise AiUpstreamError(f"Could not reach Telegram API: {exc}") from exc


def list_telegram_chats() -> list[dict[str, str]]:
    """Fetch recent chats that have messaged the configured Telegram bot."""
    resp = _telegram_request("getUpdates")
    seen: dict[str, dict[str, str]] = {}
    for update in resp.get("result", []):
        msg = update.get("message") or update.get("channel_post") or {}
        chat = msg.get("chat") or {}
        cid = chat.get("id")
        if cid is None:
            continue
        key = str(cid)
        title = chat.get("title") or " ".join(
            p for p in [chat.get("first_name", ""), chat.get("last_name", "")] if p
        ).strip()
        seen[key] = {
            "id": key,
            "name": title or chat.get("username") or key,
            "username": chat.get("username") or "",
            "type": chat.get("type") or "",
        }
    return list(seen.values())


def _format_registration_alert(row: dict[str, Any], header: str) -> str:
    status = "Paid ✓" if row.get("paid") else "Pending"
    return (
        f"{header}\n"
        f"Ref: {row.get('ref', '')}\n"
        f"Name: {row.get('name', '')}\n"
        f"Phone: {row.get('phone', '')}\n"
        f"Telegram: {row.get('tg', '')}\n"
        f"Org: {row.get('org', '')}\n"
        f"Type: {row.get('type', '')}\n"
        f"Ticket: {row.get('ticket', '')}\n"
        f"Amount: ${row.get('amount', 0)}\n"
        f"Status: {status}"
    )


def _notify_telegram_async(text: str) -> None:
    """Send a Telegram notification in the background without blocking HTTP responses."""
    if not (_get_telegram_token() and _get_telegram_chat_id()):
        return

    def _send() -> None:
        try:
            _telegram_request("sendMessage", {"chat_id": _get_telegram_chat_id(), "text": text})
        except Exception as exc:
            sys.stderr.write(f"[telegram] notify skipped: {exc}\n")

    threading.Thread(target=_send, daemon=True).start()


def send_telegram_excel_backup() -> dict[str, Any]:
    """Upload current registrations.xlsx to the configured Telegram chat."""
    token = _get_telegram_token()
    chat_id = _get_telegram_chat_id()
    if not token or not chat_id:
        raise TelegramConfigError("Set both TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID first.")

    data = export_excel_bytes()
    rows = list_registrations()
    paid_count = sum(1 for r in rows if r.get("paid"))
    caption = f"📊 Registrations Backup ({len(rows)} total, {paid_count} paid)"

    boundary = "----RegBoundary" + os.urandom(12).hex()
    parts = [
        f"--{boundary}\r\nContent-Disposition: form-data; name=\"chat_id\"\r\n\r\n{chat_id}\r\n".encode("utf-8"),
        f"--{boundary}\r\nContent-Disposition: form-data; name=\"caption\"\r\n\r\n{caption}\r\n".encode("utf-8"),
        (
            f"--{boundary}\r\nContent-Disposition: form-data; name=\"document\"; "
            f"filename=\"{DB_PATH.name}\"\r\n"
            "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n"
        ).encode("utf-8")
        + data
        + b"\r\n",
        f"--{boundary}--\r\n".encode("utf-8"),
    ]
    body = b"".join(parts)
    req = urllib.request.Request(
        f"https://api.telegram.org/bot{token}/sendDocument",
        data=body,
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        raise AiUpstreamError(f"Telegram backup error ({exc.code}): {detail}") from exc


def scan_card_with_gemini(images: list[dict[str, Any]]) -> dict[str, str]:
    """Send 1-3 card images to Gemini Vision API and return structured holder fields."""
    api_key = _get_gemini_key()
    if not api_key:
        raise AiConfigError("GEMINI_API_KEY is not configured on the server.")
    if not isinstance(images, list) or not (1 <= len(images) <= 3):
        raise ValueError("Provide 1 to 3 card images.")

    parts: list[dict[str, Any]] = [{"text": CARD_SCAN_PROMPT}]
    allowed_mimes = {"image/jpeg", "image/png", "image/webp"}
    for img in images:
        if not isinstance(img, dict):
            raise ValueError("Invalid image item.")
        mime = str(img.get("mimeType") or "image/jpeg").strip().lower()
        if mime not in allowed_mimes:
            mime = "image/jpeg"
        b64 = str(img.get("data") or "").strip()
        if b64.startswith("data:") and "," in b64:
            b64 = b64.split(",", 1)[1]
        if not b64:
            raise ValueError("Empty image data.")
        parts.append({"inline_data": {"mime_type": mime, "data": b64}})

    payload = {
        "contents": [{"parts": parts}],
        "generationConfig": {
            "temperature": 0.1,
            "responseMimeType": "application/json",
            "responseSchema": {
                "type": "OBJECT",
                "properties": {
                    "name": {"type": "STRING"},
                    "phone": {"type": "STRING"},
                    "tg": {"type": "STRING"},
                    "org": {"type": "STRING"},
                    "role": {
                        "type": "STRING",
                        "enum": ["Dentist", "Student", "Technician", "Unknown"],
                    },
                },
                "required": ["name", "phone", "tg", "org", "role"],
            },
        },
    }
    models_to_try = list(dict.fromkeys([_get_gemini_model(), *FALLBACK_GEMINI_MODELS]))
    raw_resp: dict[str, Any] | None = None
    last_error = ""
    for model in models_to_try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
        req = urllib.request.Request(
            url,
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json", "x-goog-api-key": api_key},
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                raw_resp = json.loads(resp.read().decode("utf-8"))
                break
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="ignore")
            try:
                msg = json.loads(detail).get("error", {}).get("message") or detail
            except json.JSONDecodeError:
                msg = detail or str(exc)
            last_error = f"Gemini API error ({exc.code}): {msg}"
            if exc.code in {404, 429, 503}:
                continue
            raise AiUpstreamError(last_error) from exc
        except (urllib.error.URLError, TimeoutError) as exc:
            last_error = f"Could not reach Gemini API: {exc}"
            continue

    if raw_resp is None:
        raise AiUpstreamError(last_error or "Gemini API request failed.")

    text_chunks = [
        p.get("text", "")
        for cand in raw_resp.get("candidates", [])
        for p in cand.get("content", {}).get("parts", [])
        if isinstance(p, dict) and p.get("text")
    ]
    if not text_chunks:
        raise AiUpstreamError("Gemini returned an empty response.")

    raw_text = "\n".join(text_chunks).strip()
    if raw_text.startswith("```"):
        raw_text = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw_text, flags=re.S)
    parsed = json.loads(raw_text)
    role = _clean_text(parsed.get("role"))
    if role not in {"Dentist", "Student", "Technician", "Unknown"}:
        role = "Unknown"
    return {
        "name": _clean_text(parsed.get("name")),
        "phone": _clean_text(parsed.get("phone")),
        "tg": _clean_text(parsed.get("tg")),
        "org": _clean_text(parsed.get("org")),
        "role": role,
    }


def _parse_paid(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    text = str(value or "").strip().lower()
    return text in {"paid", "yes", "true", "1", "✓", "paid ✓"}


def _parse_amount(value: Any) -> float:
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = "".join(ch for ch in str(value or "") if ch.isdigit() or ch == ".")
    try:
        return float(cleaned) if cleaned else 0.0
    except ValueError:
        return 0.0


def _clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _normalize_row(
    raw: dict[str, Any],
    *,
    allow_paid: bool = True,
    validate_pricing: bool = False,
) -> dict[str, Any]:
    ref = _clean_text(raw.get("ref"))
    name = _clean_text(raw.get("name"))
    if not ref or not name:
        raise ValueError("Both 'ref' and 'name' are required.")

    phone = _clean_text(raw.get("phone"))
    role = _clean_text(raw.get("type"))
    ticket = _clean_text(raw.get("ticket"))

    if validate_pricing:
        if len(phone) < 8:
            raise ValueError("Phone number must be at least 8 characters.")
        if role not in PRICING or ticket not in PRICING[role]:
            raise ValueError("Invalid role or ticket selection.")
        amount = float(PRICING[role][ticket])
    else:
        amount = _parse_amount(raw.get("amount"))

    paid = _parse_paid(raw.get("paid")) if allow_paid else False
    reg_date = _clean_text(raw.get("date")) or date.today().isoformat()

    return {
        "ref": ref,
        "name": name,
        "phone": phone,
        "tg": _clean_text(raw.get("tg") or raw.get("telegram")),
        "org": _clean_text(raw.get("org") or raw.get("organization")),
        "type": role,
        "ticket": ticket,
        "amount": amount,
        "paid": paid,
        "date": reg_date[:10],
    }


def _apply_workbook_styles(ws: Any, row_count: int) -> None:
    header_fill = PatternFill(start_color="0B5CAB", end_color="0B5CAB", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    body_font = Font(name="Calibri", size=11, color="14213D")
    paid_fill = PatternFill(start_color="D1FADF", end_color="D1FADF", fill_type="solid")
    paid_font = Font(name="Calibri", size=11, bold=True, color="027A48")
    pending_fill = PatternFill(start_color="FEF0C7", end_color="FEF0C7", fill_type="solid")
    pending_font = Font(name="Calibri", size=11, bold=True, color="B54708")
    thin_side = Side(border_style="thin", color="E4E7EC")
    cell_border = Border(bottom=thin_side)

    ws.row_dimensions[1].height = 26
    for idx, (_, label, width) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=idx, value=label)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.column_dimensions[get_column_letter(idx)].width = width

    for row_idx in range(2, row_count + 2):
        ws.row_dimensions[row_idx].height = 21
        for col_idx, (key, _, _) in enumerate(COLUMNS, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = cell_border
            if key == "phone":
                cell.number_format = "@"
                cell.font = body_font
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif key == "amount":
                cell.number_format = "$#,##0.00"
                cell.font = body_font
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif key == "paid":
                is_paid = _parse_paid(cell.value)
                cell.fill = paid_fill if is_paid else pending_fill
                cell.font = paid_font if is_paid else pending_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.font = body_font
                cell.alignment = Alignment(horizontal="left", vertical="center")

    ws.freeze_panes = "A2"
    last_col = get_column_letter(len(COLUMNS))
    ws.auto_filter.ref = f"A1:{last_col}{max(1, row_count + 1)}"


def _write_rows_unlocked(rows: list[dict[str, Any]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME

    for row_idx, item in enumerate(rows, start=2):
        for col_idx, (key, _, _) in enumerate(COLUMNS, start=1):
            if key == "paid":
                value: Any = "Paid" if item.get("paid") else "Pending"
            elif key == "amount":
                value = _parse_amount(item.get("amount"))
            else:
                value = _clean_text(item.get(key))
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            if key == "phone":
                cell.data_type = "s"

    _apply_workbook_styles(ws, len(rows))

    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(prefix=".tmp_reg_", suffix=".xlsx", dir=str(DB_PATH.parent))
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        wb.save(tmp_path)
        os.replace(tmp_path, DB_PATH)
    except PermissionError as exc:
        raise ExcelLockedError(
            f"Cannot update '{DB_PATH.name}' because it is currently open in Excel or locked. "
            "Please close Excel and try again."
        ) from exc
    finally:
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass


def _read_rows_unlocked() -> list[dict[str, Any]]:
    if not DB_PATH.exists():
        _write_rows_unlocked([])
        return []

    try:
        wb = load_workbook(DB_PATH, data_only=True)
    except PermissionError as exc:
        raise ExcelLockedError(
            f"Cannot read '{DB_PATH.name}' because it is locked by another process."
        ) from exc

    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
    raw_headers = [_clean_text(cell.value).lower() for cell in ws[1]] if ws.max_row >= 1 else []
    col_map: dict[int, str] = {}
    for idx, header in enumerate(raw_headers):
        mapped = HEADER_ALIASES.get(header)
        if mapped:
            col_map[idx] = mapped

    if not col_map:
        col_map = {idx: key for idx, (key, _, _) in enumerate(COLUMNS)}

    rows: list[dict[str, Any]] = []
    for row_values in ws.iter_rows(min_row=2, values_only=True):
        if not any(v is not None and str(v).strip() != "" for v in row_values):
            continue
        record: dict[str, Any] = {}
        for idx, value in enumerate(row_values):
            key = col_map.get(idx)
            if key:
                record[key] = value
        if not _clean_text(record.get("ref")) and not _clean_text(record.get("name")):
            continue
        rows.append(
            {
                "ref": _clean_text(record.get("ref")),
                "name": _clean_text(record.get("name")),
                "phone": _clean_text(record.get("phone")),
                "tg": _clean_text(record.get("tg")),
                "org": _clean_text(record.get("org")),
                "type": _clean_text(record.get("type")),
                "ticket": _clean_text(record.get("ticket")),
                "amount": _parse_amount(record.get("amount")),
                "paid": _parse_paid(record.get("paid")),
                "date": _clean_text(record.get("date"))[:10],
            }
        )
    wb.close()
    return rows


def ensure_database() -> int:
    """Ensure registrations.xlsx exists and return current row count."""
    with DB_LOCK:
        rows = _read_rows_unlocked()
        return len(rows)


def list_registrations() -> list[dict[str, Any]]:
    """Return all registration rows sorted newest first."""
    with DB_LOCK:
        rows = _read_rows_unlocked()
    indexed = list(enumerate(rows))
    indexed.sort(key=lambda pair: (pair[1].get("date") or "", pair[0]), reverse=True)
    return [row for _, row in indexed]


def upsert_registration(
    raw: dict[str, Any],
    *,
    allow_paid: bool,
    validate_pricing: bool,
) -> dict[str, Any]:
    """Create or update a single registration row in registrations.xlsx."""
    item = _normalize_row(raw, allow_paid=allow_paid, validate_pricing=validate_pricing)
    with DB_LOCK:
        rows = _read_rows_unlocked()
        for idx, existing in enumerate(rows):
            if existing["ref"] == item["ref"]:
                if not allow_paid and existing.get("paid"):
                    raise RegistrationConflictError("Registration is already marked as Paid.")
                rows[idx] = item
                _write_rows_unlocked(rows)
                return item
        rows.append(item)
        _write_rows_unlocked(rows)
        return item


def bulk_upsert_registrations(raw_items: list[dict[str, Any]]) -> dict[str, Any]:
    """Upsert multiple registration records into registrations.xlsx."""
    normalized: list[dict[str, Any]] = []
    for raw in raw_items:
        if isinstance(raw, dict):
            normalized.append(_normalize_row(raw, allow_paid=True, validate_pricing=False))

    with DB_LOCK:
        rows = _read_rows_unlocked()
        by_ref = {row["ref"]: idx for idx, row in enumerate(rows)}
        saved = 0
        for item in normalized:
            if item["ref"] in by_ref:
                rows[by_ref[item["ref"]]] = item
            else:
                by_ref[item["ref"]] = len(rows)
                rows.append(item)
            saved += 1
        _write_rows_unlocked(rows)

    indexed = list(enumerate(rows))
    indexed.sort(key=lambda pair: (pair[1].get("date") or "", pair[0]), reverse=True)
    return {"saved": saved, "rows": [row for _, row in indexed]}


def set_registration_paid(ref: str, paid: bool) -> dict[str, Any] | None:
    """Update the paid flag of a registration by Ref."""
    with DB_LOCK:
        rows = _read_rows_unlocked()
        for idx, row in enumerate(rows):
            if row["ref"] == ref:
                rows[idx]["paid"] = bool(paid)
                _write_rows_unlocked(rows)
                return rows[idx]
    return None


def delete_registration(ref: str) -> bool:
    """Delete a registration row by Ref."""
    with DB_LOCK:
        rows = _read_rows_unlocked()
        remaining = [row for row in rows if row["ref"] != ref]
        if len(remaining) == len(rows):
            return False
        _write_rows_unlocked(remaining)

    if _get_render_url():
        with SYNC_LOCK:
            state = _load_sync_state()
            deleted = set(state.get("deleted", []))
            deleted.add(ref)
            state["deleted"] = sorted(deleted)
            state.get("rows", {}).pop(ref, None)
            _save_sync_state(state)
        _delete_remote_async(ref)
    return True


def export_excel_bytes() -> bytes:
    """Read raw bytes of registrations.xlsx safely under lock."""
    with DB_LOCK:
        if not DB_PATH.exists():
            _write_rows_unlocked([])
        try:
            return DB_PATH.read_bytes()
        except PermissionError as exc:
            raise ExcelLockedError(
                f"Cannot read '{DB_PATH.name}' because it is locked by Excel."
            ) from exc


SYNC_LOCK = threading.Lock()
_SYNC_STATUS: dict[str, Any] = {"running": False, "lastSyncAt": "", "lastError": "", "lastAction": ""}
_SYNC_THREAD_STARTED = False


def _load_sync_state() -> dict[str, Any]:
    if not SYNC_STATE_PATH.exists():
        return {"rows": {}, "deleted": []}
    try:
        data = json.loads(SYNC_STATE_PATH.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            return {"rows": data.get("rows") or {}, "deleted": data.get("deleted") or []}
    except (OSError, json.JSONDecodeError):
        pass
    return {"rows": {}, "deleted": []}


def _save_sync_state(state: dict[str, Any]) -> None:
    SYNC_STATE_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def _row_sig(row: dict[str, Any]) -> tuple[Any, ...]:
    return (
        _clean_text(row.get("ref")),
        _clean_text(row.get("name")),
        _clean_text(row.get("phone")),
        _clean_text(row.get("tg")),
        _clean_text(row.get("org")),
        _clean_text(row.get("type")),
        _clean_text(row.get("ticket")),
        round(_parse_amount(row.get("amount")), 2),
        bool(row.get("paid")),
        _clean_text(row.get("date"))[:10],
    )


def _render_request(method: str, path: str, body: dict[str, Any] | None = None) -> dict[str, Any]:
    base = _get_render_url()
    if not base:
        raise ValueError("RENDER_URL is not configured.")
    pin = _get_render_pin()
    data = json.dumps(body, ensure_ascii=False).encode("utf-8") if body is not None else None
    headers = {"X-Admin-Pin": pin, "Content-Type": "application/json"}
    req = urllib.request.Request(f"{base}{path}", data=data, headers=headers, method=method)
    with urllib.request.urlopen(req, timeout=20) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _delete_remote_async(ref: str) -> None:
    def _run() -> None:
        try:
            _render_request("DELETE", f"/api/registrations/{quote(ref, safe='')}")
        except Exception:
            pass

    threading.Thread(target=_run, daemon=True).start()


def sync_with_render_once() -> dict[str, Any]:
    """Perform safe two-way sync between Render server and local registrations.xlsx."""
    render_url = _get_render_url()
    if not render_url:
        return {"ok": False, "error": "RENDER_URL is not configured."}

    with SYNC_LOCK:
        _SYNC_STATUS["running"] = True
        try:
            remote_resp = _render_request("GET", "/api/registrations")
            remote_list = [
                _normalize_row(r, allow_paid=True, validate_pricing=False)
                for r in (remote_resp.get("rows") or [])
                if isinstance(r, dict) and _clean_text(r.get("ref"))
            ]
            remote_by_ref = {r["ref"]: r for r in remote_list}

            with DB_LOCK:
                local_list = _read_rows_unlocked()
                local_by_ref = {r["ref"]: r for r in local_list}
                state = _load_sync_state()
                prev_by_ref: dict[str, dict[str, Any]] = state.get("rows", {})
                deleted_set: set[str] = set(state.get("deleted", []))

                merged_list: list[dict[str, Any]] = []
                push_to_render: list[dict[str, Any]] = []
                delete_from_render: list[str] = []
                local_changed = False
                pulled = 0

                for local_row in local_list:
                    ref = local_row["ref"]
                    if ref in deleted_set:
                        # Row was deleted; ensure it is purged from local Excel
                        local_changed = True
                        continue

                    if ref in remote_by_ref:
                        remote_row = remote_by_ref[ref]
                        prev_row = prev_by_ref.get(ref)
                        if _row_sig(local_row) == _row_sig(remote_row):
                            merged_list.append(local_row)
                            continue

                        if prev_row and _row_sig(local_row) != _row_sig(prev_row) and _row_sig(remote_row) == _row_sig(prev_row):
                            merged_list.append(local_row)
                            push_to_render.append(local_row)
                        elif prev_row and _row_sig(remote_row) != _row_sig(prev_row) and _row_sig(local_row) == _row_sig(prev_row):
                            merged_list.append(remote_row)
                            local_changed = True
                            pulled += 1
                        else:
                            chosen = dict(remote_row)
                            if local_row.get("paid") or remote_row.get("paid"):
                                chosen["paid"] = True
                            if _row_sig(chosen) != _row_sig(local_row):
                                local_changed = True
                                pulled += 1
                            if _row_sig(chosen) != _row_sig(remote_row):
                                push_to_render.append(chosen)
                            merged_list.append(chosen)
                    else:
                        # Row exists locally, but NOT on remote
                        if ref in prev_by_ref:
                            # It previously existed on both, but is now gone from remote -> remote DELETED it!
                            deleted_set.add(ref)
                            local_changed = True
                            # Do not add to merged_list, do not push to render
                        else:
                            # Brand new row created locally
                            merged_list.append(local_row)
                            push_to_render.append(local_row)

                for remote_row in reversed(remote_list):
                    ref = remote_row["ref"]
                    if ref in local_by_ref:
                        continue

                    if ref in deleted_set:
                        delete_from_render.append(ref)
                        continue

                    if ref in prev_by_ref:
                        # Previously existed on both, but now gone from local -> local DELETED it!
                        deleted_set.add(ref)
                        delete_from_render.append(ref)
                    else:
                        # Brand new row created on remote
                        merged_list.append(remote_row)
                        local_changed = True
                        pulled += 1

                if local_changed:
                    _write_rows_unlocked(merged_list)

            for ref_del in delete_from_render:
                try:
                    _render_request("DELETE", f"/api/registrations/{quote(ref_del, safe='')}")
                except Exception:
                    pass

            if push_to_render:
                _render_request("POST", "/api/registrations/bulk", {"rows": push_to_render})

            _save_sync_state(
                {
                    "rows": {r["ref"]: r for r in merged_list},
                    "deleted": sorted(deleted_set),
                }
            )
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            deleted_cnt = len(delete_from_render)
            action = f"Synced ({len(merged_list)} total, +{pulled} pulled, +{len(push_to_render)} pushed, -{deleted_cnt} remote-deleted)"
            _SYNC_STATUS.update({"lastSyncAt": now_str, "lastError": "", "lastAction": action})
            if pulled or push_to_render or deleted_cnt or local_changed:
                sys.stdout.write(f"[render-sync] {action} -> {DB_PATH.name}\n")
            return {"ok": True, "total": len(merged_list), "pulled": pulled, "pushed": len(push_to_render), "lastSyncAt": now_str}
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="ignore")
            msg = f"HTTP {exc.code}: {detail}"
            _SYNC_STATUS["lastError"] = msg
            return {"ok": False, "error": msg}
        except Exception as exc:
            _SYNC_STATUS["lastError"] = str(exc)
            return {"ok": False, "error": str(exc)}
        finally:
            _SYNC_STATUS["running"] = False


def _trigger_sync_async() -> None:
    if not _get_render_url():
        return
    threading.Thread(target=sync_with_render_once, daemon=True).start()


def ensure_sync_thread() -> None:
    """Start background Render sync loop once if RENDER_URL is configured."""
    global _SYNC_THREAD_STARTED
    if _SYNC_THREAD_STARTED or not _get_render_url():
        return
    _SYNC_THREAD_STARTED = True

    def _loop() -> None:
        while True:
            if _get_render_url():
                sync_with_render_once()
            time.sleep(_get_sync_interval())

    threading.Thread(target=_loop, daemon=True).start()


class RegistrationHandler(BaseHTTPRequestHandler):
    """HTTP handler serving index.html and JSON/Excel REST APIs."""

    server_version = "RegistrationExcelServer/1.0"

    def log_message(self, format: str, *args: Any) -> None:
        sys.stdout.write(f"[{self.address_string()}] {format % args}\n")

    def _send_cors_headers(self) -> None:
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, PATCH, DELETE, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, X-Admin-Pin")

    def _send_json(self, status: int, payload: dict[str, Any]) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self._send_cors_headers()
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _send_bytes(
        self,
        status: int,
        content_type: str,
        data: bytes,
        *,
        filename: str | None = None,
    ) -> None:
        self.send_response(status)
        self._send_cors_headers()
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        if filename:
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.end_headers()
        self.wfile.write(data)

    def _is_admin(self, pin_override: str | None = None) -> bool:
        supplied = pin_override if pin_override is not None else self.headers.get("X-Admin-Pin", "")
        return hmac.compare_digest(str(supplied), str(_get_admin_pin()))

    def _admin_settings_payload(self) -> dict[str, Any]:
        gem_key = _get_gemini_key()
        tg_tok = _get_telegram_token()
        return {
            "ok": True,
            "geminiConfigured": bool(gem_key),
            "geminiMask": _mask_secret(gem_key),
            "geminiModel": _get_gemini_model(),
            "telegramBotConfigured": bool(tg_tok),
            "telegramTokenMask": _mask_secret(tg_tok),
            "telegramChatId": _get_telegram_chat_id(),
            "renderUrl": _get_render_url(),
            "renderPinConfigured": bool(os.environ.get("RENDER_ADMIN_PIN", "").strip()),
            "syncStatus": dict(_SYNC_STATUS),
        }

    def _require_admin(self) -> bool:
        if self._is_admin():
            return True
        self._send_json(HTTPStatus.UNAUTHORIZED, {"ok": False, "error": "Invalid Admin PIN."})
        return False

    def _read_json(self, max_bytes: int = MAX_BODY_BYTES) -> Any:
        raw_len = self.headers.get("Content-Length", "0")
        try:
            length = int(raw_len)
        except ValueError:
            raise ValueError("Invalid Content-Length.")
        if length < 0 or length > max_bytes:
            raise ValueError("Request body too large.")
        raw = self.rfile.read(length) if length > 0 else b"{}"
        return json.loads(raw.decode("utf-8"))

    def do_OPTIONS(self) -> None:
        self.send_response(HTTPStatus.NO_CONTENT)
        self._send_cors_headers()
        self.end_headers()

    def do_GET(self) -> None:
        path = urlparse(self.path).path
        try:
            if path in {"/", "/index.html", "/admin", "/admin/"}:
                if not HTML_PATH.exists():
                    self._send_json(HTTPStatus.NOT_FOUND, {"ok": False, "error": "index.html not found."})
                    return
                self._send_bytes(HTTPStatus.OK, "text/html; charset=utf-8", HTML_PATH.read_bytes())
                return

            if path == "/api/health":
                count = ensure_database()
                self._send_json(
                    HTTPStatus.OK,
                    {
                        "ok": True,
                        "db": DB_PATH.name,
                        "dbPath": str(DB_PATH),
                        "count": count,
                        "aiConfigured": bool(_get_gemini_key()),
                        "geminiModel": _get_gemini_model(),
                        "botConfigured": bool(_get_telegram_token()),
                        "telegramConfigured": bool(_get_telegram_token() and _get_telegram_chat_id()),
                        "syncEnabled": bool(_get_render_url()),
                        "syncStatus": dict(_SYNC_STATUS),
                    },
                )
                return

            if path == "/api/admin/settings":
                if not self._require_admin():
                    return
                self._send_json(HTTPStatus.OK, self._admin_settings_payload())
                return

            if path == "/api/admin/telegram-chats":
                if not self._require_admin():
                    return
                chats = list_telegram_chats()
                self._send_json(HTTPStatus.OK, {"ok": True, "chats": chats})
                return

            if path == "/api/registrations":
                if not self._require_admin():
                    return
                rows = list_registrations()
                self._send_json(HTTPStatus.OK, {"ok": True, "rows": rows})
                return

            if path == "/api/export":
                if not self._require_admin():
                    return
                data = export_excel_bytes()
                self._send_bytes(
                    HTTPStatus.OK,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    data,
                    filename="registrations.xlsx",
                )
                return

            self._send_json(HTTPStatus.NOT_FOUND, {"ok": False, "error": "Not found."})
        except TelegramConfigError as exc:
            self._send_json(HTTPStatus.SERVICE_UNAVAILABLE, {"ok": False, "error": str(exc)})
        except AiUpstreamError as exc:
            self._send_json(HTTPStatus.BAD_GATEWAY, {"ok": False, "error": str(exc)})
        except ExcelLockedError as exc:
            self._send_json(HTTPStatus.LOCKED, {"ok": False, "error": str(exc)})

    def do_POST(self) -> None:
        path = urlparse(self.path).path
        try:
            limit = MAX_SCAN_BODY_BYTES if path == "/api/scan-card" else MAX_BODY_BYTES
            payload = self._read_json(max_bytes=limit)

            if path == "/api/admin/verify":
                pin = payload.get("pin") if isinstance(payload, dict) else None
                if self._is_admin(pin_override=pin if pin is not None else None):
                    self._send_json(HTTPStatus.OK, {"ok": True})
                else:
                    self._send_json(HTTPStatus.UNAUTHORIZED, {"ok": False, "error": "Invalid Admin PIN."})
                return

            if path == "/api/admin/settings":
                if not self._require_admin():
                    return
                if not isinstance(payload, dict):
                    raise ValueError("Expected JSON object.")
                _save_env_settings(
                    {
                        "GEMINI_API_KEY": payload.get("geminiApiKey") if "geminiApiKey" in payload else None,
                        "GEMINI_MODEL": payload.get("geminiModel") if "geminiModel" in payload else None,
                        "TELEGRAM_BOT_TOKEN": payload.get("telegramBotToken") if "telegramBotToken" in payload else None,
                        "TELEGRAM_CHAT_ID": payload.get("telegramChatId") if "telegramChatId" in payload else None,
                        "RENDER_URL": payload.get("renderUrl").rstrip("/") if isinstance(payload.get("renderUrl"), str) else None,
                        "RENDER_ADMIN_PIN": payload.get("renderAdminPin") if "renderAdminPin" in payload else None,
                    }
                )
                ensure_sync_thread()
                if "renderUrl" in payload or "renderAdminPin" in payload:
                    sync_with_render_once()
                self._send_json(HTTPStatus.OK, self._admin_settings_payload())
                return

            if path == "/api/admin/sync-now":
                if not self._require_admin():
                    return
                res = sync_with_render_once()
                status = HTTPStatus.OK if res.get("ok") else HTTPStatus.BAD_GATEWAY
                self._send_json(status, {**res, "syncStatus": dict(_SYNC_STATUS)})
                return

            if path == "/api/admin/telegram-backup":
                if not self._require_admin():
                    return
                send_telegram_excel_backup()
                self._send_json(HTTPStatus.OK, {"ok": True})
                return

            if path == "/api/scan-card":
                if not isinstance(payload, dict):
                    raise ValueError("Expected JSON object.")
                result = scan_card_with_gemini(payload.get("images") or [])
                self._send_json(
                    HTTPStatus.OK,
                    {"ok": True, "provider": "gemini", "model": _get_gemini_model(), "data": result},
                )
                return

            if path == "/api/registrations":
                if not isinstance(payload, dict):
                    raise ValueError("Expected JSON object.")
                saved = upsert_registration(payload, allow_paid=False, validate_pricing=True)
                _notify_telegram_async(_format_registration_alert(saved, "📋 New Registration (Pending)"))
                _trigger_sync_async()
                self._send_json(HTTPStatus.CREATED, {"ok": True, "row": saved})
                return

            if path == "/api/registrations/bulk":
                if not self._require_admin():
                    return
                items = payload.get("rows") if isinstance(payload, dict) else payload
                if not isinstance(items, list):
                    raise ValueError("Expected a list of registrations.")
                result = bulk_upsert_registrations(items)
                _trigger_sync_async()
                self._send_json(HTTPStatus.OK, {"ok": True, **result})
                return

            self._send_json(HTTPStatus.NOT_FOUND, {"ok": False, "error": "Not found."})
        except (json.JSONDecodeError, ValueError) as exc:
            self._send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": str(exc)})
        except (AiConfigError, TelegramConfigError) as exc:
            self._send_json(HTTPStatus.SERVICE_UNAVAILABLE, {"ok": False, "error": str(exc)})
        except AiUpstreamError as exc:
            self._send_json(HTTPStatus.BAD_GATEWAY, {"ok": False, "error": str(exc)})
        except RegistrationConflictError as exc:
            self._send_json(HTTPStatus.CONFLICT, {"ok": False, "error": str(exc)})
        except ExcelLockedError as exc:
            self._send_json(HTTPStatus.LOCKED, {"ok": False, "error": str(exc)})

    def do_PATCH(self) -> None:
        path = urlparse(self.path).path
        if not path.startswith("/api/registrations/"):
            self._send_json(HTTPStatus.NOT_FOUND, {"ok": False, "error": "Not found."})
            return
        if not self._require_admin():
            return
        ref = unquote(path.removeprefix("/api/registrations/")).strip()
        try:
            payload = self._read_json()
            if not isinstance(payload, dict) or "paid" not in payload:
                raise ValueError("Expected JSON object with 'paid' boolean.")
            updated = set_registration_paid(ref, bool(payload["paid"]))
            if not updated:
                self._send_json(HTTPStatus.NOT_FOUND, {"ok": False, "error": "Registration not found."})
                return
            header = "✅ Payment Confirmed" if updated.get("paid") else "⏳ Reverted to Pending"
            _notify_telegram_async(_format_registration_alert(updated, header))
            _trigger_sync_async()
            self._send_json(HTTPStatus.OK, {"ok": True, "row": updated})
        except (json.JSONDecodeError, ValueError) as exc:
            self._send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": str(exc)})
        except ExcelLockedError as exc:
            self._send_json(HTTPStatus.LOCKED, {"ok": False, "error": str(exc)})

    def do_DELETE(self) -> None:
        path = urlparse(self.path).path
        if not path.startswith("/api/registrations/"):
            self._send_json(HTTPStatus.NOT_FOUND, {"ok": False, "error": "Not found."})
            return
        if not self._require_admin():
            return
        ref = unquote(path.removeprefix("/api/registrations/")).strip()
        try:
            if not delete_registration(ref):
                self._send_json(HTTPStatus.NOT_FOUND, {"ok": False, "error": "Registration not found."})
                return
            self._send_json(HTTPStatus.OK, {"ok": True})
        except ExcelLockedError as exc:
            self._send_json(HTTPStatus.LOCKED, {"ok": False, "error": str(exc)})


def _lan_ip() -> str:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.connect(("8.8.8.8", 80))
            return sock.getsockname()[0]
    except OSError:
        return "127.0.0.1"


def main() -> None:
    parser = argparse.ArgumentParser(description="Conference Registration Excel DB Server")
    parser.add_argument("--init-only", action="store_true", help="Initialize Excel DB and exit")
    parser.add_argument("--port", type=int, default=PORT, help="Port to listen on (default: 8000)")
    args = parser.parse_args()

    count = ensure_database()
    print(f"Excel database ready: {DB_PATH} ({count} registration(s))")
    if args.init_only:
        return

    if _get_render_url():
        print(f"Render sync enabled -> {_get_render_url()} (every {_get_sync_interval()}s)")
        ensure_sync_thread()

    httpd = ThreadingHTTPServer((HOST, args.port), RegistrationHandler)
    lan_ip = _lan_ip()
    print(f"Local URL : http://localhost:{args.port}")
    print(f"Network   : http://{lan_ip}:{args.port}")
    print("Press Ctrl+C to stop.")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping server...")
    finally:
        httpd.server_close()


if __name__ == "__main__":
    main()
