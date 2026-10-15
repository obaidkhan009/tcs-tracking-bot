#!/usr/bin/env python3
"""Shipment tracking bot.

Looks up a tracking number in your Excel sheet, asks the courier's website for
the live status, prints it in plain words, and writes the result back into the
sheet.

Usage
  python tracker.py shipments.xlsx                    # interactive: type numbers, get status
  python tracker.py shipments.xlsx --track 7723...    # one number
  python tracker.py shipments.xlsx --all              # refresh every row
  python tracker.py shipments.xlsx --all --pending    # refresh only rows not yet delivered
  python tracker.py --track 7723...                   # no sheet, just check the courier

Options
  --sheet NAME        worksheet to use (default: first sheet)
  --courier NAME      courier to use when the sheet has no Courier column (default: TCS)
  --no-write          do not write results back to the sheet
  --eta-days N        add an estimate (booking date + N days); off by default
  --history           also print the full checkpoint history
"""

import argparse
import sys
import time
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from couriers import TrackingError, TrackingResult, get_courier

# Header names we look for in the sheet (case-insensitive, spaces/underscores ignored).
TRACKING_HEADERS = ("trackingnumber", "tracking", "trackingno", "tracking#", "cn", "cnnumber",
                    "cnno", "consignment", "consignmentno", "consignmentnumber", "awb", "awbno")
COURIER_HEADERS = ("courier", "carrier", "courierservice", "shipper", "service")

# Columns the bot writes. Created at the end of the sheet if missing.
OUTPUT_COLUMNS = ["Status", "Last Update", "Origin", "Destination", "Booked On",
                  "Delivered On", "Received By", "Days In Transit", "ETA / Note", "Checked At"]

FILL = {
    "DELIVERED": PatternFill("solid", fgColor="C6EFCE"),
    "IN TRANSIT": PatternFill("solid", fgColor="FFEB9C"),
    "RETURNED": PatternFill("solid", fgColor="F4B084"),
    "NOT FOUND": PatternFill("solid", fgColor="FFC7CE"),
    "ERROR": PatternFill("solid", fgColor="FFC7CE"),
}


def norm(text) -> str:
    return "".join(ch for ch in str(text or "").lower() if ch.isalnum() or ch == "#")


def clean_number(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return "".join(str(value).split())


# ---------------------------------------------------------------- Excel helpers

class Sheet:
    def __init__(self, path: Path, sheet_name: str | None, write: bool):
        self.path = path
        self.write = write
        self.wb = load_workbook(path)
        self.ws = self.wb[sheet_name] if sheet_name else self.wb.worksheets[0]
        self.header_row, self.headers = self._find_headers()
        self.col_tracking = self._find_col(TRACKING_HEADERS)
        self.col_courier = self._find_col(COURIER_HEADERS)
        if self.col_tracking is None:
            raise SystemExit(
                f"Could not find a tracking-number column in '{self.ws.title}'. "
                f"Headers seen: {[h for h in self.headers.values()]}. "
                f"Name the column 'Tracking Number' (or CN / Consignment No / AWB)."
            )
        self.out_cols = {}
        if write:
            self._ensure_output_columns()

    def _find_headers(self):
        """Header row = first row (within the first 10) that contains a tracking-like header."""
        for r in range(1, min(self.ws.max_row, 10) + 1):
            row = {c: self.ws.cell(r, c).value for c in range(1, self.ws.max_column + 1)}
            if any(norm(v) in TRACKING_HEADERS for v in row.values()):
                return r, {c: str(v).strip() for c, v in row.items() if v not in (None, "")}
        row = {c: self.ws.cell(1, c).value for c in range(1, self.ws.max_column + 1)}
        return 1, {c: str(v).strip() for c, v in row.items() if v not in (None, "")}

    def _find_col(self, candidates):
        for c, h in self.headers.items():
            if norm(h) in candidates:
                return c
        return None

    def _ensure_output_columns(self):
        existing = {norm(h): c for c, h in self.headers.items()}
        next_col = max(self.headers) + 1 if self.headers else 1
        for name in OUTPUT_COLUMNS:
            col = existing.get(norm(name))
            if col is None:
                col = next_col
                next_col += 1
                cell = self.ws.cell(self.header_row, col, name)
                cell.font = Font(bold=True)
                self.headers[col] = name
                width = max(12, len(name) + 2)
                self.ws.column_dimensions[get_column_letter(col)].width = width
            self.out_cols[name] = col

    def rows(self):
        """Yield (row_index, tracking_number, courier_name) for every data row."""
        for r in range(self.header_row + 1, self.ws.max_row + 1):
            number = clean_number(self.ws.cell(r, self.col_tracking).value)
            if not number:
                continue
            courier = self.ws.cell(r, self.col_courier).value if self.col_courier else None
            yield r, number, (str(courier).strip() if courier else None)

    def find(self, number: str):
        """Rows whose tracking number matches (exact first, then 'contains')."""
        target = clean_number(number)
        exact = [(r, n, c) for r, n, c in self.rows() if n == target]
        if exact:
            return exact
        return [(r, n, c) for r, n, c in self.rows() if target and target in n]

    def status_of(self, row: int) -> str:
        col = self.out_cols.get("Status") or next(
            (c for c, h in self.headers.items() if norm(h) == "status"), None)
        return str(self.ws.cell(row, col).value or "").strip().upper() if col else ""

    def write_result(self, row: int, res: TrackingResult, eta_days: int):
        if not self.write:
            return
        fmt = lambda d: d.strftime("%d %b %Y %I:%M %p") if d else ""
        values = {
            "Status": res.state,
            "Last Update": res.summary,
            "Origin": res.origin,
            "Destination": res.destination,
            "Booked On": fmt(res.booking_date),
            "Delivered On": fmt(res.delivered_on),
            "Received By": res.received_by,
            "Days In Transit": res.days_in_transit if res.days_in_transit is not None else "",
            "ETA / Note": res.eta_text(eta_days),
            "Checked At": datetime.now().strftime("%d %b %Y %I:%M %p"),
        }
        for name, value in values.items():
            cell = self.ws.cell(row, self.out_cols[name], value)
            if name == "Status":
                cell.fill = FILL.get(res.state, PatternFill())
                cell.font = Font(bold=True)

    def save(self):
        if not self.write:
            return
        try:
            self.wb.save(self.path)
        except PermissionError:
            raise SystemExit(
                f"Could not save {self.path.name}. Close it in Excel and run again.")


# ---------------------------------------------------------------- Output

def print_result(res: TrackingResult, eta_days: int, history: bool = False):
    icon = {"DELIVERED": "✅", "IN TRANSIT": "🚚", "RETURNED": "↩️", "NOT FOUND": "❓",
            "ERROR": "⚠️"}.get(res.state, "•")
    print(f"\n{icon}  {res.courier} {res.number}  —  {res.state}")
    if not res.found:
        print(f"   {res.summary}")
        return
    if res.origin or res.destination:
        print(f"   Route:     {res.origin or '?'}  →  {res.destination or '?'}")
    if res.booking_date:
        print(f"   Booked:    {res.booking_date.strftime('%d %b %Y %I:%M %p')}")
    print(f"   Latest:    {res.summary}")
    eta = res.eta_text(eta_days)
    if eta:
        label = {"DELIVERED": "Delivery", "RETURNED": "Outcome"}.get(res.state, "Note")
        print(f"   {label}:  {eta}")
    if history and res.checkpoints:
        print("   History:")
        for cp in res.checkpoints:
            when = cp.time.strftime("%d %b %Y %I:%M %p") if cp.time else cp.raw_time
            extra = f"  ({cp.detail})" if cp.detail else ""
            print(f"     {when:<22} {cp.status}{extra}")


def check_one(number: str, courier_name: str | None, default_courier: str) -> TrackingResult:
    try:
        courier = get_courier(courier_name or default_courier)
        return courier.track(number)
    except TrackingError as exc:
        return TrackingResult(courier=(courier_name or default_courier).upper(), number=number,
                              found=False, state="ERROR", summary=str(exc))


# ---------------------------------------------------------------- Modes

def run_lookup(sheet: Sheet | None, number: str, args) -> None:
    """Find the number in the sheet (if any), check it with the courier, write back."""
    if sheet is None:
        res = check_one(number, None, args.courier)
        print_result(res, args.eta_days, args.history)
        return

    matches = sheet.find(number)
    if not matches:
        print(f"\n'{number}' is not in {sheet.path.name}. Checking with {args.courier.upper()} anyway.")
        res = check_one(number, None, args.courier)
        print_result(res, args.eta_days, args.history)
        return

    for row, full_number, courier_name in matches:
        res = check_one(full_number, courier_name, args.courier)
        print(f"\nRow {row} in '{sheet.ws.title}':", end="")
        print_result(res, args.eta_days, args.history)
        sheet.write_result(row, res, args.eta_days)
    sheet.save()
    if sheet.write:
        print(f"\nSaved to {sheet.path.name}.")


def run_all(sheet: Sheet, args) -> None:
    rows = list(sheet.rows())
    if args.pending:
        rows = [r for r in rows if sheet.status_of(r[0]) not in ("DELIVERED", "RETURNED")]
    if not rows:
        print("Nothing to check.")
        return
    print(f"Checking {len(rows)} shipment(s) in '{sheet.ws.title}'...")
    counts = {}
    for i, (row, number, courier_name) in enumerate(rows, 1):
        res = check_one(number, courier_name, args.courier)
        counts[res.state] = counts.get(res.state, 0) + 1
        print(f"[{i}/{len(rows)}] row {row}  {number:<16} {res.state:<11} {res.summary}")
        sheet.write_result(row, res, args.eta_days)
        if i < len(rows):
            time.sleep(args.delay)
    sheet.save()
    print("\nSummary: " + ", ".join(f"{k}: {v}" for k, v in sorted(counts.items())))
    if sheet.write:
        print(f"Saved to {sheet.path.name}.")


def run_interactive(sheet: Sheet | None, args) -> None:
    where = f"{sheet.path.name} (sheet '{sheet.ws.title}')" if sheet else "no sheet, courier only"
    print(f"Tracking bot ready — {where}. Default courier: {args.courier.upper()}.")
    print("Type a tracking number and press Enter. Type 'all' to refresh every row, 'q' to quit.\n")
    while True:
        try:
            text = input("Tracking number> ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            break
        if not text:
            continue
        if text.lower() in ("q", "quit", "exit"):
            break
        if text.lower() == "all" and sheet:
            run_all(sheet, args)
            continue
        run_lookup(sheet, text, args)
        print()


IS_FROZEN = bool(getattr(sys, "frozen", False))   # True when running as a packaged .exe/app


def pick_excel_for_app() -> Path | None:
    """Packaged app with no arguments: use the only .xlsx next to the app, or ask."""
    app_dir = Path(sys.executable).resolve().parent if IS_FROZEN else Path.cwd()
    candidates = sorted(p for p in app_dir.glob("*.xlsx") if not p.name.startswith("~$"))
    if len(candidates) == 1:
        print(f"Using {candidates[0].name} (found next to the app).")
        return candidates[0]
    if candidates:
        print("Excel files found next to the app:")
        for i, c in enumerate(candidates, 1):
            print(f"  {i}. {c.name}")
        choice = input("Pick a number, or drag another file here and press Enter: ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(candidates):
            return candidates[int(choice) - 1]
        text = choice
    else:
        text = input("Drag your Excel file into this window and press Enter "
                     "(or just press Enter to check numbers without a sheet): ").strip()
    text = text.strip("'\"")
    if not text:
        return None
    path = Path(text).expanduser()
    if not path.exists():
        print(f"File not found: {path}")
        return pick_excel_for_app()
    return path


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("excel", nargs="?", help="path to the .xlsx file")
    p.add_argument("--track", metavar="NUMBER", help="tracking number to look up")
    p.add_argument("--all", action="store_true", help="refresh every row in the sheet")
    p.add_argument("--pending", action="store_true", help="with --all: skip rows already DELIVERED")
    p.add_argument("--sheet", help="worksheet name (default: first sheet)")
    p.add_argument("--courier", default="tcs", help="default courier when the sheet has none (default: tcs)")
    p.add_argument("--no-write", action="store_true", help="do not write results into the sheet")
    p.add_argument("--eta-days", type=int, default=0,
                   help="add an estimated arrival of booking date + N days (off by default; "
                        "TCS does not publish an expected date)")
    p.add_argument("--delay", type=float, default=1.0, help="seconds to wait between requests in --all")
    p.add_argument("--history", action="store_true", help="print full checkpoint history")
    args = p.parse_args(argv)

    sheet = None
    if args.excel:
        path = Path(args.excel).expanduser()
        if not path.exists():
            raise SystemExit(f"File not found: {path}")
        sheet = Sheet(path, args.sheet, write=not args.no_write)
    elif IS_FROZEN and not args.track:
        print("TCS Tracking Bot\n")
        path = pick_excel_for_app()
        if path:
            sheet = Sheet(path, args.sheet, write=not args.no_write)
    elif args.all:
        raise SystemExit("--all needs an Excel file.")

    if args.all:
        run_all(sheet, args)
    elif args.track:
        run_lookup(sheet, args.track, args)
    else:
        run_interactive(sheet, args)


def _entry():
    """Keep the window open on errors when double-clicked as a packaged app."""
    try:
        main()
    except SystemExit as exc:
        if exc.code not in (None, 0):
            print(exc.code)
        if IS_FROZEN and exc.code not in (None, 0):
            input("\nPress Enter to close.")
        return exc.code
    except KeyboardInterrupt:
        return 0
    except Exception as exc:  # noqa: BLE001
        print(f"\nSomething went wrong: {exc}")
        if IS_FROZEN:
            input("\nPress Enter to close.")
        return 1


if __name__ == "__main__":
    sys.exit(_entry())
