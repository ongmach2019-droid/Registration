"""Local web server and Excel database backend for Conference Registration.

Serves index.html and persists registration records directly into
registrations.xlsx in the same directory using openpyxl.
"""

from __future__ import annotations

import argparse
from datetime import date
import hmac
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import os
from pathlib import Path
import socket
import sys
import tempfile
import threading
from typing import Any
from urllib.parse import unquote, urlparse

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = Path(os.environ.get("REGISTRATION_DB", BASE_DIR / "registrations.xlsx"))
HTML_PATH = BASE_DIR / "index.html"
ADMIN_PIN = os.environ.get("ADMIN_PIN", "1234")
HOST = os.environ.get("HOST", "0.0.0.0")
PORT = int(os.environ.get("PORT", "8000"))
MAX_BODY_BYTES = 2 * 1024 * 1024

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
    "Dentist": {"Main Conference": 50, "Hands-On": 80, "Main + Hands-On": 120},
    "Student": {"Main Conference": 20, "Hands-On": 40, "Main + Hands-On": 55},
    "Technician": {"Full-day hands-on course": 150},
}

DB_LOCK = threading.Lock()


class ExcelLockedError(Exception):
    """Raised when registrations.xlsx is open/locked by another program (e.g. Excel)."""


class RegistrationConflictError(Exception):
    """Raised when attempting an invalid modification to an existing registration."""


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
        return hmac.compare_digest(str(supplied), str(ADMIN_PIN))

    def _require_admin(self) -> bool:
        if self._is_admin():
            return True
        self._send_json(HTTPStatus.UNAUTHORIZED, {"ok": False, "error": "Invalid Admin PIN."})
        return False

    def _read_json(self) -> Any:
        raw_len = self.headers.get("Content-Length", "0")
        try:
            length = int(raw_len)
        except ValueError:
            raise ValueError("Invalid Content-Length.")
        if length < 0 or length > MAX_BODY_BYTES:
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
            if path in {"/", "/index.html"}:
                if not HTML_PATH.exists():
                    self._send_json(HTTPStatus.NOT_FOUND, {"ok": False, "error": "index.html not found."})
                    return
                self._send_bytes(HTTPStatus.OK, "text/html; charset=utf-8", HTML_PATH.read_bytes())
                return

            if path == "/api/health":
                count = ensure_database()
                self._send_json(
                    HTTPStatus.OK,
                    {"ok": True, "db": DB_PATH.name, "dbPath": str(DB_PATH), "count": count},
                )
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
        except ExcelLockedError as exc:
            self._send_json(HTTPStatus.LOCKED, {"ok": False, "error": str(exc)})

    def do_POST(self) -> None:
        path = urlparse(self.path).path
        try:
            payload = self._read_json()

            if path == "/api/admin/verify":
                pin = payload.get("pin") if isinstance(payload, dict) else None
                if self._is_admin(pin_override=pin if pin is not None else None):
                    self._send_json(HTTPStatus.OK, {"ok": True})
                else:
                    self._send_json(HTTPStatus.UNAUTHORIZED, {"ok": False, "error": "Invalid Admin PIN."})
                return

            if path == "/api/registrations":
                if not isinstance(payload, dict):
                    raise ValueError("Expected JSON object.")
                saved = upsert_registration(payload, allow_paid=False, validate_pricing=True)
                self._send_json(HTTPStatus.CREATED, {"ok": True, "row": saved})
                return

            if path == "/api/registrations/bulk":
                if not self._require_admin():
                    return
                items = payload.get("rows") if isinstance(payload, dict) else payload
                if not isinstance(items, list):
                    raise ValueError("Expected a list of registrations.")
                result = bulk_upsert_registrations(items)
                self._send_json(HTTPStatus.OK, {"ok": True, **result})
                return

            self._send_json(HTTPStatus.NOT_FOUND, {"ok": False, "error": "Not found."})
        except (json.JSONDecodeError, ValueError) as exc:
            self._send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": str(exc)})
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
