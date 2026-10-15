"""Interactive TCS tracking bot UI for Google Colab (also works in Jupyter).

Run `%run colab_app.py` in a notebook cell. It:
  1. connects Google Drive and lists every .xlsx found there,
  2. lets you pick the sheet from a dropdown (or upload one from your computer),
  3. shows a search box: type a tracking number, press Enter, see the live TCS
     status, and the sheet is updated in place,
  4. has a "Refresh all rows" button and a "Download updated file" button.
"""

import os
import sys
import time
from argparse import Namespace
from pathlib import Path

import ipywidgets as w
from IPython.display import display

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

from tracker import Sheet, run_all, run_lookup  # noqa: E402

try:
    from google.colab import drive as gdrive, files as gfiles  # type: ignore
    IN_COLAB = True
except ImportError:  # local Jupyter / testing
    gdrive = gfiles = None
    IN_COLAB = False

DRIVE_ROOTS = ["/content/drive/MyDrive", "/content/drive/Shareddrives"]
MAX_DEPTH = 4
MAX_FILES = 300
SKIP_DIRS = {".Trash", "Colab Notebooks", ".ipynb_checkpoints", "node_modules", ".git"}


def find_excel_files(roots):
    """All .xlsx under the roots (limited depth), newest first."""
    found = []
    for root in roots:
        root = Path(root)
        if not root.exists():
            continue
        for dirpath, dirnames, filenames in os.walk(root):
            depth = len(Path(dirpath).relative_to(root).parts)
            dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS and not d.startswith(".")]
            if depth >= MAX_DEPTH:
                dirnames[:] = []
            for f in filenames:
                if f.lower().endswith(".xlsx") and not f.startswith("~$"):
                    found.append(Path(dirpath) / f)
            if len(found) >= MAX_FILES:
                break
    found.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return found


def label_for(path: Path, roots):
    for root in roots:
        try:
            return str(path.relative_to(root))
        except ValueError:
            continue
    return str(path)


def connect_drive():
    if not IN_COLAB:
        return [Path.cwd()]
    gdrive.mount("/content/drive", force_remount=False)
    return DRIVE_ROOTS


def build_ui():
    roots = connect_drive()
    excel_files = find_excel_files(roots)

    options = [(label_for(p, roots), str(p)) for p in excel_files]
    sheet_dd = w.Dropdown(options=options or [("No Excel files found in Drive", "")],
                          description="Sheet:", layout=w.Layout(width="70%"),
                          style={"description_width": "60px"})
    upload_btn = w.Button(description="Upload from computer", icon="upload")
    rescan_btn = w.Button(description="Rescan Drive", icon="refresh")

    number_tx = w.Text(placeholder="type tracking number, press Enter",
                       description="Track:", layout=w.Layout(width="50%"),
                       continuous_update=False,
                       style={"description_width": "60px"})
    check_btn = w.Button(description="Check", button_style="primary", icon="search")
    history_cb = w.Checkbox(value=False, description="show full history", indent=False)

    all_btn = w.Button(description="Refresh all rows", icon="list")
    pending_cb = w.Checkbox(value=True, description="skip delivered / returned", indent=False)
    download_btn = w.Button(description="Download updated file", icon="download")

    out = w.Output(layout=w.Layout(border="1px solid #ddd", padding="6px", max_height="420px",
                                   overflow="auto"))

    def args():
        return Namespace(courier="tcs", eta_days=0, history=history_cb.value,
                         pending=pending_cb.value, delay=1.0)

    def current_sheet():
        path = sheet_dd.value
        if not path:
            raise SystemExit("Pick a sheet first (or upload one).")
        return Sheet(Path(path), None, write=True)

    def on_check(_=None):
        number = number_tx.value.strip()
        with out:
            out.clear_output(wait=True)
            if not number:
                print("Type a tracking number first.")
                return
            try:
                if sheet_dd.value:
                    run_lookup(current_sheet(), number, args())
                else:
                    run_lookup(None, number, args())
            except SystemExit as exc:
                print(exc)
            except Exception as exc:  # noqa: BLE001
                print("Something went wrong:", exc)
        number_tx.value = ""

    def on_all(_=None):
        with out:
            out.clear_output(wait=True)
            try:
                run_all(current_sheet(), args())
            except SystemExit as exc:
                print(exc)
            except Exception as exc:  # noqa: BLE001
                print("Something went wrong:", exc)

    def on_upload(_=None):
        with out:
            out.clear_output(wait=True)
            if not IN_COLAB:
                print("Upload works only inside Google Colab.")
                return
            uploaded = gfiles.upload()
            names = [n for n in uploaded if n.lower().endswith(".xlsx")]
            if not names:
                print("No .xlsx uploaded.")
                return
            path = str(Path(names[0]).resolve())
            sheet_dd.options = [(f"(uploaded) {names[0]}", path)] + list(sheet_dd.options)
            sheet_dd.value = path
            print("Using uploaded file:", names[0])

    def on_rescan(_=None):
        with out:
            out.clear_output(wait=True)
            files_now = find_excel_files(roots)
            sheet_dd.options = [(label_for(p, roots), str(p)) for p in files_now] or \
                [("No Excel files found in Drive", "")]
            print(f"Found {len(files_now)} Excel file(s).")

    def on_download(_=None):
        with out:
            out.clear_output(wait=True)
            path = sheet_dd.value
            if not path:
                print("Pick a sheet first.")
            elif path.startswith("/content/drive/"):
                print("This file lives in Google Drive and is already updated there:", path)
            elif IN_COLAB:
                gfiles.download(path)
            else:
                print("Updated file is at:", path)

    check_btn.on_click(on_check)

    def on_enter(change):          # Enter (or leaving the box) submits; ignore the clear
        if change["new"].strip():
            on_check()

    number_tx.observe(on_enter, names="value")
    all_btn.on_click(on_all)
    upload_btn.on_click(on_upload)
    rescan_btn.on_click(on_rescan)
    download_btn.on_click(on_download)

    title = w.HTML("<h3 style='margin:4px 0'>TCS Tracking Bot</h3>"
                   "<div style='color:#666'>Pick your sheet, type a tracking number, press Enter. "
                   "The sheet is updated in place after every check.</div>")
    ui = w.VBox([
        title,
        w.HBox([sheet_dd, rescan_btn, upload_btn]),
        w.HBox([number_tx, check_btn, history_cb]),
        w.HBox([all_btn, pending_cb, download_btn]),
        out,
    ])
    display(ui)
    with out:
        if excel_files:
            print(f"Connected. {len(excel_files)} Excel file(s) found in Drive. "
                  f"Selected: {options[0][0]}")
        else:
            print("Connected to Drive but no .xlsx found. Use 'Upload from computer' "
                  "or put the file in My Drive and click 'Rescan Drive'.")
    return ui, {"check": on_check, "all": on_all, "sheet": sheet_dd, "number": number_tx,
                "out": out, "history": history_cb, "pending": pending_cb}


if __name__ == "__main__":
    APP = build_ui()
