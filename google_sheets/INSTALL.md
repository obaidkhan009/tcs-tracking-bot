# TCS Tracking inside Google Sheets

The search box and the status live inside the spreadsheet. Nothing to install
on the computer; it works for anyone who can edit the sheet.

## 1. Make the sheet a Google Sheet

The file must be a native Google Sheet, not an .xlsx. In Google Drive,
right-click the Excel file > Open with > Google Sheets, then in that window
File > Save as Google Sheets. Use the new Google Sheet from now on.

The sheet needs one column header with the tracking number, such as
`Tracking Number`, `CN`, `Consignment No` or `AWB`.

## 2. Add the script (one time, about 2 minutes)

1. In the Google Sheet: Extensions > Apps Script.
2. Delete the few lines in the editor, paste the whole content of `Code.gs`
   from this folder, press the save icon (or Ctrl/Cmd+S).
3. Go back to the spreadsheet tab and reload the page.
4. A new menu **TCS Tracking** appears next to Help.
5. First use: Google asks for permission. Click Continue > choose your account >
   Advanced > "Go to Untitled project (unsafe)" > Allow. The warning appears
   because the script is your own, not a published add-on.

## 3. Use it

- **TCS Tracking > Search box (sidebar)**: a panel opens on the right. Type a
  tracking number, press Enter. The status shows in the panel and the row in
  the sheet is updated at the same time.
- **TCS Tracking > Enable 'Track' sheet**: creates a sheet called *Track* with
  a yellow cell B2. Type a number there and press Enter; the status, route,
  dates and history appear below it, and the matching row in the data sheet is
  updated. This is the "search inside the sheet" option.
- **Check selected rows**: select some rows in the data sheet, run it.
- **Check pending rows**: refreshes every row not yet DELIVERED or RETURNED.
- **Check all rows**: refreshes everything.

The script adds these columns at the end of the data sheet and fills them:
Status, Last Update, Origin, Destination, Booked On, Delivered On,
Received By, Days In Transit, ETA / Note, Checked At. Status is colour coded.

## Limits

- Google allows a script to run about 6 minutes per call. "Check all rows"
  handles roughly 300 rows per run; for bigger sheets use "Check pending rows"
  a few times.
- A number not in the sheet is still checked and shown, but nothing is written.
