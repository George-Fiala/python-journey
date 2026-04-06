# PO Flow — Purchase Order Automation Tool

A command-line Python tool that automates the full purchase order lifecycle in a manufacturing/engineering environment — from creation to goods receipt, with Excel integration, PDF generation, and Outlook email automation.

Built to solve a real problem: manual PO creation was slow, error-prone, and disconnected from a shared tracking workbook. This tool reduced PO creation time from ~4 minutes to under 1 minute per order (~85% reduction in admin time).

---

## What it does

**Create a Purchase Order**
- Guided CLI input with fuzzy autocomplete for machines, areas, suppliers, and reason codes
- Auto-generates a PO number based on cost centre, initials, date, and sequence
- Writes the order to the correct monthly sheet in a local Excel workbook
- Adds the order to a "To Receive" backlog sheet for goods receipt tracking
- Generates a styled A4 PDF (via ReportLab) with supplier details, line items, VAT, and totals
- Opens a pre-filled Outlook email with the PDF attached, ready to send

**Mark as Received (Goods Receipt)**
- Marks a PO as received in the To Receive backlog
- Updates the delivered date in the monthly sheet
- Records the invoice number

**Full Sync with Shared Workbook**
- Reads the main shared workbook (e.g. SharePoint/Teams sync folder)
- Imports any orders missing from the local workbook
- Syncs delivery dates and costs back to local sheets
- Detects and removes duplicate rows and misrouted orders
- Runs a verification pass to ensure all orders are on the correct monthly sheet

---

## Key features

- **Fuzzy autocomplete** — typo-tolerant input for machines, areas, suppliers
- **Cost centre resolution** — automatically maps machine + area to the correct cost centre code
- **Supplier database** — JSON-based supplier lookup with aliases, contact details, and addresses
- **Part number support** — optional part number field shown in PDF and Excel summary
- **VAT handling** — configurable VAT rate per order (default 20%)
- **Savings tracking** — optional cost initiative / savings field per order
- **Excel auto-fit and totals** — monthly sheets update spend and savings totals automatically
- **Retry logic** — gracefully handles locked Excel files (prompts to close and retry)

---

## Tech stack

| Library | Purpose |
|---|---|
| `openpyxl` | Read/write Excel workbooks |
| `reportlab` | Generate styled A4 PDF purchase orders |
| `pywin32` | Outlook COM automation (Windows only) |
| `difflib` | Fuzzy matching for autocomplete |
| `json` | Supplier database |

---

## Setup

**1. Clone or download this repository**

**2. Install dependencies**
```bash
pip install openpyxl reportlab pywin32
```

**3. Edit the CONFIG section** at the top of `po_flow.py`:

```python
WORKBOOK_PATH      = r"C:\Users\YOUR_USERNAME\OneDrive\Desktop\Parts Ordered Home.xlsx"
SUPPLIERS_JSON_PATH = r"C:\Users\YOUR_USERNAME\OneDrive\Desktop\Python project\suppliers.json"
MAIN_WORKBOOK_PATH  = r"C:\Users\YOUR_USERNAME\Company\Parts order files\Parts Ordered.xlsx"
PO_PDF_OUTPUT_DIR   = r"C:\Users\YOUR_USERNAME\OneDrive\Desktop\PO PDFs"

INITIALS   = "XX"   # Your initials, used in PO number generation
LOGO_PATH  = r""    # Optional: path to your company logo PNG
```

**4. Customise machines and cost centres**

Edit `KNOWN_MACHINES`, `KNOWN_AREAS`, `MACHINE_ALIAS_RAW`, and `CC_LOOKUP` to match the machines and cost centre codes at your site. The file ships with generic example entries.

**5. Set up your suppliers JSON**

Create a `suppliers.json` file in the format:
```json
{
  "suppliers": [
    {
      "name": "Supplier Name",
      "supplier_id": "SUP001",
      "aliases": ["short name", "nickname"],
      "email": "orders@supplier.com",
      "phone": "+44 1234 567890",
      "contact_name": "John Smith",
      "address_lines": ["123 Street", "City", "Postcode"]
    }
  ]
}
```

**6. Run**
```bash
python po_flow.py
```

---

## Excel workbook structure

The tool expects (and creates if missing) the following sheet structure:

| Sheet | Purpose |
|---|---|
| `January 2025`, `February 2025`, ... | Monthly order log — one sheet per month |
| `To Receive` | Goods receipt backlog — all open and completed orders |
| `Cost centre` | Optional: machine/area → cost centre lookup table |

Monthly sheets are auto-detected by header names — the tool handles column variations gracefully.

---

## PO number format

```
{CostCentreCode}-{Initials}{DD}{MM}{Seq}{YY}

Example:  100-XX040412026
          ↑   ↑  ↑  ↑↑  ↑
          CC  In Day Mo Seq Year
```

---

## Notes

- Designed for **Windows** (Outlook integration via `pywin32`, `.venv` auto-launch)
- Outlook integration is optional — set `AUTO_OPEN_OUTLOOK = False` to disable
- The tool will **never overwrite** existing data — it always appends to the next empty row
- All currency values stored as floats; displayed as `£#,##0.00` in Excel

---

## Background

This tool was built to automate procurement admin in an engineering stores environment managing 15+ machines and 30+ suppliers. Before automation, each purchase order required manual Excel entry, a separate PDF, and a manually composed 
email — around 4 minutes per order. With PO Flow, the same process takes under 1 minute (~50 seconds on average).

The savings tracking feature was added to support monthly cost initiative reporting, where alternative suppliers and negotiated prices are logged against each order.

---

## Roadmap / planned features

- [ ] Delivery Note OCR — auto-match scanned PDFs from email to open POs using Tesseract
- [ ] Spend Analyser — CSV import with anomaly detection and category breakdown
- [ ] Web dashboard — local Flask UI instead of CLI

---

## License

MIT — free to use, adapt, and build on.

