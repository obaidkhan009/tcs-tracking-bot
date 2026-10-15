# Shipment Tracking Bot (TCS)

Type a tracking number, the bot finds it in your Excel sheet, checks the TCS
website for the live status, tells you whether it is delivered or returned,
and writes the answer back into the sheet.

## For non-developers: download and run (no install)

1. Open the **Releases** page:
   https://github.com/obaidkhan009/tcs-tracking-bot/releases/latest
2. Download the file for your computer:
   - Windows: `TCS-Tracker-Windows.exe`
   - Mac with Apple chip (M1/M2/M3/M4): `TCS-Tracker-macOS-AppleSilicon.zip`
   - Mac with Intel chip: `TCS-Tracker-macOS-Intel.zip`
3. Put the app in a folder together with your Excel file (one .xlsx in that
   folder). Close the Excel file.
4. Double-click the app. It finds the Excel file next to it and asks:
   `Tracking number>`. Type a number, press Enter, read the status.
   Type `all` to refresh every row, `q` to quit. Results are saved into the
   Excel file (Status, Last Update, Delivered On, Received By, and so on).

First-run warnings (the app is not signed with a paid certificate):
- **Windows**: SmartScreen says "Windows protected your PC". Click
  *More info*, then *Run anyway*.
- **Mac**: unzip, then right-click `TCS-Tracker` and choose *Open*, then *Open*
  again. After the first time, double-click works. If it still refuses, open
  Terminal and run `xattr -d com.apple.quarantine ~/Downloads/TCS-Tracker`.

## Run it in Google Colab (no install, works in the browser)

Open this link and click **Runtime > Run all**:
https://colab.research.google.com/github/obaidkhan009/tcs-tracking-bot/blob/main/TCS_Tracker_Colab.ipynb

It asks for Google Drive access, lists every Excel file in your Drive in a
dropdown, and shows a search box. Type a tracking number, press Enter, read the
status. The sheet in Drive is updated in place after every check, so you never
re-upload or download anything. Buttons: *Refresh all rows*, *Upload from
computer* (for a file not in Drive) and *Download updated file* (for uploads).

Have a blank Colab notebook already? Paste this into a cell and run it:

```
!git clone -q https://github.com/obaidkhan009/tcs-tracking-bot.git /content/tcs-tracking-bot 2>/dev/null || git -C /content/tcs-tracking-bot pull -q
!pip install -q -r /content/tcs-tracking-bot/requirements.txt
%cd -q /content/tcs-tracking-bot
%run colab_app.py
```

## Does it change my sheet?

Yes, in place. On the first run it adds the status columns at the end of your
existing sheet. Every later run overwrites those columns for the rows checked.
It never creates a new file or a new worksheet, and your original columns are
untouched. Keep a copy of the file if you want the pre-bot version.

## For developers: run from source

### Setup (one time)

```
cd Tracking_Bot
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

(Already done on this Mac. Skip if `.venv` exists.)

## Your Excel sheet

Any .xlsx works as long as one column header is the tracking number. Recognised
header names: `Tracking Number`, `Tracking`, `CN`, `CN No`, `Consignment No`, `AWB`.

Optional `Courier` column. If it is missing, every row is treated as TCS.

Example (see `shipments_sample.xlsx`):

| Order ID | Customer   | City   | Courier | Tracking Number |
|----------|------------|--------|---------|-----------------|
| ORD-1001 | Ahmed Khan | Lahore | TCS     | 772319512345    |

The bot adds these columns at the end and fills them per row:
`Status`, `Last Update`, `Origin`, `Destination`, `Booked On`, `Delivered On`,
`Received By`, `Days In Transit`, `ETA / Note`, `Checked At`.
Status is colour coded: green = delivered, yellow = in transit, red = not found.

Close the file in Excel before running, otherwise it cannot be saved.

## Run it

Interactive (the "bot" mode). Type numbers one by one:

```
.venv/bin/python tracker.py shipments.xlsx
Tracking number> 772319512345
Tracking number> all      # refresh every row
Tracking number> q
```

Or double-click `Track.command` in Finder and drop your .xlsx path when asked.

One number:

```
.venv/bin/python tracker.py shipments.xlsx --track 772319512345
```

Refresh the whole sheet (skips rows already delivered with `--pending`):

```
.venv/bin/python tracker.py shipments.xlsx --all
.venv/bin/python tracker.py shipments.xlsx --all --pending
```

Check a number without any sheet:

```
.venv/bin/python tracker.py --track 772319512345 --history
```

Options: `--sheet NAME`, `--no-write`, `--history`, `--eta-days 3`, `--delay 1`.

## What the status means

- **DELIVERED**: TCS reports delivery. Shows date, time and who received it.
- **RETURNED**: TCS sent the parcel back to the sender (their "Returned To
  Sender" outcome).
- **IN TRANSIT**: latest checkpoint from TCS (booked, departed, arrived at hub,
  out for delivery, undelivered attempt) plus days since booking.
  TCS does not publish an expected delivery date, so by default no estimate is
  shown. If you want one, add `--eta-days 3`: the note then shows booking date
  plus 3 days as an estimate and says how many days overdue once that passes.
- **NOT FOUND**: TCS says "No Data Found/Invalid CN". Usually a typo, or the
  parcel is not yet scanned into the TCS system.
- **ERROR**: network problem or unsupported courier. Run again later.

## How it talks to TCS

It calls the same endpoint the TCS tracking page uses
(`https://www.tcsexpress.com/apibridge`, action `trackapinew`), so the result
matches the website. If TCS changes their site, update `couriers/tcs.py`.

## Adding another courier (Leopards, M&P, etc.)

Create `couriers/<name>.py` with a `track(number) -> TrackingResult` function
and register it in `couriers/__init__.py`. Put the courier name in the
`Courier` column of the sheet.
