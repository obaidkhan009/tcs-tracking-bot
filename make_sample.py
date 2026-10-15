"""Creates shipments_sample.xlsx so you can see the expected sheet layout.
Replace the numbers with real TCS consignment numbers."""

from openpyxl import Workbook
from openpyxl.styles import Font

wb = Workbook()
ws = wb.active
ws.title = "Shipments"

headers = ["Order ID", "Customer", "City", "Courier", "Tracking Number"]
ws.append(headers)
for c in range(1, len(headers) + 1):
    ws.cell(1, c).font = Font(bold=True)

rows = [
    ["ORD-1001", "Ahmed Khan", "Lahore", "TCS", "772319512345"],
    ["ORD-1002", "Sara Malik", "Karachi", "TCS", "772319512346"],
    ["ORD-1003", "Bilal Raza", "Islamabad", "TCS", "772319512347"],
]
for r in rows:
    ws.append(r)

for col, width in zip("ABCDE", (12, 16, 14, 10, 18)):
    ws.column_dimensions[col].width = width

wb.save("shipments_sample.xlsx")
print("Created shipments_sample.xlsx")
