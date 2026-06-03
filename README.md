# Amazon Ads Audit Report Generator

Python CLI that reads an official **Amazon Ads Bulk Sheet** (`.xlsx`) and produces a styled **Audit Report** workbook with three sheets: **OPD**, **ASIN wise report.**, and **Keywords**.

## Installation

Requires Python 3.9+.

```bash
pip install openpyxl pandas numpy flask
```

Or from this folder:

```bash
pip install -r requirements.txt
```

## Web UI (manual upload)

1. Install dependencies (see above).
2. Start the server:

```bash
python web_app.py
```

Or on Windows, double-click `run_web.bat`.

3. Open **http://127.0.0.1:5000** in your browser.
4. Upload your Amazon bulk `.xlsx` file and click **Generate & download**.

Optional flags:

```bash
python web_app.py --host 0.0.0.0 --port 8080
```

Use `--host 0.0.0.0` only on a trusted network if others need access on your LAN.

## Download the bulk sheet from Amazon Ads

1. Sign in to [Amazon Advertising](https://advertising.amazon.com/) (Seller Central / Vendor Central ads console).
2. Go to **Bulk operations** (sometimes under **Campaigns** → **Bulk operations** or **Settings**).
3. Choose **Download spreadsheet** / **Export**.
4. Select the date range and ad products you want (Sponsored Products, Brands, Display).
5. Include performance columns (Spend, Sales, Orders, etc.) for the same period you want in the audit.
6. Save the downloaded `.xlsx` file — the tool expects the standard sheet names:
   - Sponsored Products Campaigns
   - Sponsored Brands Campaigns
   - SB Multi Ad Group Campaigns
   - Sponsored Display Campaigns
   - (Search term sheets are optional; keyword analysis uses campaign keyword rows.)

## Usage

```bash
python amazon_ads_audit.py bulk_sheet.xlsx
```

Custom output path:

```bash
python amazon_ads_audit.py bulk_sheet.xlsx --output report.xlsx
```

### Example (your files)

```bash
python amazon_ads_audit.py "c:\Users\Morning\Downloads\bulk-a21vhz1tv3zuzi-20260501-20260527-1779886481295.xlsx" --output "c:\Users\Morning\Downloads\BlackBuck_Audit_Report_Generated.xlsx"
```

Default output name: `{bulk_sheet_stem}_Audit_Report.xlsx` in the same folder as the input.

## Output sheets

| Sheet | Contents |
|-------|----------|
| **OPD** | Account summary (SP, SB, SBV, SD CPC, SD vCPM), SP auto/manual, placements, auto & manual targets, match types, SB targeting, SD contextual/audience |
| **ASIN wise report.** | Per ASIN/SKU metrics sorted by Spend |
| **Keywords** | Exact, Phrase, and Broad columns side-by-side, sorted by Spend within each match type |

Styling: dark blue titles (`#1F3864`), section headers (`#2E75B6`), light blue column headers (`#BDD7EE`), alternating rows, ₹ currency, percentage formats for ACoS/CTR/Conv Rate, frozen header rows.

## Common errors and fixes

| Error | Cause | Fix |
|-------|--------|-----|
| `ModuleNotFoundError: openpyxl` | Dependencies not installed | Run `pip install openpyxl pandas numpy` |
| `Input file not found` | Wrong path or quotes | Use the full path to the `.xlsx`; on Windows wrap paths with spaces in quotes |
| `No recognized campaign sheets found` | Not an Amazon bulk export | Re-download from Bulk Operations; file must contain SP/SB/SD campaign sheets |
| `Input must be an Excel file` | Wrong extension | Use `.xlsx` from Amazon (not CSV) |
| Empty SB/SD sections | No spend on those ad types in the date range | Normal if the account only runs SP; verify date range in the bulk download |
| ASIN sheet empty | No Product Ad rows with ASIN/SKU | Ensure the bulk export includes product ads and performance data |
| Numbers differ from Amazon UI | Bulk uses export date range & attribution window | Match UI date range; tool uses `Sales` / `Orders` columns (auto-detects `14 Day Total Sales (#)` etc.) |
| `Permission denied` on save | Output file open in Excel | Close the workbook and run again |

## How metrics are aggregated

To avoid double-counting, the tool uses Amazon entity levels consistently:

- **Account SP / Auto vs Manual:** `Campaign` rows
- **Placements:** `Bidding Adjustment` rows (Top of Search, Rest of Search, Product Pages)
- **Targets / match types:** `Keyword` and `Product Targeting` rows
- **ASIN report:** `Product Ad` rows (SP + Display)
- **SB / SD splits:** Campaign or targeting entities with fallbacks when campaign-level metrics are zero

## License

Use freely for internal reporting.
