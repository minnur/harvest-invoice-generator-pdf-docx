# Invoice Generator

Generates professional invoices (DOCX + PDF) from Harvest time-tracking CSV exports. Uses a DOCX template for layout/styling and injects data via [docxtpl](https://docxtpl.readthedocs.io/). Supports both single-file and batch mode with automatic invoice numbering based on CSV date ranges.

## Prerequisites

```
pip install docxtpl
```

| Package | Purpose |
|---|---|
| [docxtpl](https://pypi.org/project/docxtpl/) | Jinja2 template rendering for DOCX files |
| [python-docx](https://pypi.org/project/python-docx/) | DOCX document creation (installed with docxtpl) |

macOS is required for PDF generation (uses Pages via AppleScript).

## Template

The default template is `templates/invoice-default.docx`. You can open it in Word or Pages to customize:

- Page margins, fonts, colors
- "INVOICE" title styling
- Header table layout (sender/recipient info)
- Footer paragraph styling

The template uses Jinja2 placeholders (`{{r left_header }}`, `{{ itemized_table }}`, `{{ payable_to }}`, etc.) that get replaced with data at generation time. The itemized line-items table (headers, data rows, summary rows) is built entirely in Python.

To use a custom template per client, set the `template` field in their config JSON to the path of their template DOCX.

## Configuration

Edit `config.json` (or a client-specific config) to set invoice details:

| Field | Description |
|---|---|
| `rate` | Hourly rate |
| `output_dir` | Output directory for generated files (empty = same directory as the CSV input) |
| `template` | Path to a custom template DOCX (empty = default `templates/invoice-default.docx`) |
| `csv_columns` | Harvest CSV column indices: `date`, `project_code`, `notes`, `hours`, `ref_url` |
| `sender` | Your name, address, and email |
| `invoice.date` | Invoice date (single-file mode only; batch mode derives this from filenames) |
| `invoice.number` | Invoice number — in batch mode the year prefix is extracted (e.g. `2026-1` -> prefix `2026`) |
| `invoice.for` | Description of services |
| `bill_to` | Client company name lines and address lines |
| `footer` | Payable to, contact name, contact info |

Example `config.json`:

```json
{
  "rate": 12.34,
  "output_dir": "/Users/myname/Invoices/ClientName/Invoices",
  "template": "",
  "csv_columns": {
    "date": 0,
    "project_code": 3,
    "notes": 5,
    "hours": 6,
    "ref_url": 14
  },
  "sender": {
    "name": "Jane Doe",
    "address": ["123 Main St", "San Francisco, CA 94101"],
    "email": "jane@example.com"
  },
  "invoice": {
    "date": "March 1, 2026",
    "number": "2026-1",
    "for": "Contractor services"
  },
  "bill_to": {
    "company": ["Acme Corp"],
    "address": ["456 Market St,", "San Francisco, CA 94105"]
  },
  "footer": {
    "payable_to": "Jane Doe",
    "contact_name": "Jane Doe",
    "contact_info": "(555) 123-4567, jane@example.com"
  }
}
```

## Usage

```
python3 generate_invoice.py <harvest_csv_or_directory> [config.json]
```

The config file defaults to `config.json` in the script directory if not specified.

### Single-file mode

Pass a CSV file path to generate one invoice using the date and number from your config:

```bash
# CSV in script directory, default config
python3 generate_invoice.py harvest_time_report_from2026-02-16to2026-02-28.csv

# Custom config file
python3 generate_invoice.py harvest_time_report_from2026-02-16to2026-02-28.csv /path/to/config.json

# Absolute path to CSV
python3 generate_invoice.py /Users/myname/Invoices/ClientName/harvest-csv/harvest_time_report_from2026-02-16to2026-02-28.csv
```

### Batch mode

Pass a directory containing Harvest CSVs to generate invoices for all of them automatically:

```bash
python3 generate_invoice.py /Users/myname/Invoices/ClientName/harvest-csv /Users/myname/Invoices/ClientName/clientname.json
```

In batch mode the script will:

1. Find all CSV files matching the pattern `harvest_time_report_from{YYYY-MM-DD}to{YYYY-MM-DD}.csv`
2. Sort them chronologically by start date
3. Assign sequential invoice numbers using the year prefix from your config (e.g. `2026-1`, `2026-2`, ...)
4. Use the **end date** from each filename as the invoice date (e.g. `to2026-02-28` -> "February 28, 2026")
5. Generate DOCX + PDF for each CSV

### Output

Each invoice produces two files in the output directory:

```
Jane-Doe-invoice-2026-1.docx
Jane-Doe-invoice-2026-1.pdf
```

If a DOCX/PDF pair already exists for an invoice number, that invoice is skipped. Delete the existing files to regenerate.

## Harvest CSV filenames

Harvest exports use this naming convention:

```
harvest_time_report_from{START_DATE}to{END_DATE}.csv
```

Examples:

```
harvest_time_report_from2026-01-06to2026-02-16.csv
harvest_time_report_from2026-02-16to2026-02-28.csv
harvest_time_report_from2026-03-01to2026-03-15.csv
```

## Recommended directory structure

Organize invoices per client with separate directories for CSV exports and generated invoices:

```
Invoices/
├── Generator/
│   ├── generate_invoice.py
│   ├── config.json              # default/template config
│   └── templates/
│       └── invoice-default.docx     # default template
│
├── ClientA/
│   ├── clienta.json             # client-specific config
│   ├── harvest-csv/             # Harvest CSV exports
│   │   ├── harvest_time_report_from2026-01-06to2026-02-16.csv
│   │   ├── harvest_time_report_from2026-02-16to2026-02-28.csv
│   │   └── harvest_time_report_from2026-03-01to2026-03-15.csv
│   └── Invoices/                # generated output (set as output_dir)
│       ├── Jane-Doe-invoice-2026-1.docx
│       ├── Jane-Doe-invoice-2026-1.pdf
│       ├── Jane-Doe-invoice-2026-2.docx
│       ├── Jane-Doe-invoice-2026-2.pdf
│       ├── Jane-Doe-invoice-2026-3.docx
│       └── Jane-Doe-invoice-2026-3.pdf
│
└── ClientB/
    ├── clientb.json
    ├── harvest-csv/
    └── Invoices/
```

Set `output_dir` in each client config to point to their `Invoices/` directory:

```json
{
  "output_dir": "/Users/myname/Invoices/ClientA/Invoices"
}
```

## Workflow

1. Export time reports from Harvest and save CSVs into the client's `harvest-csv/` directory
2. Run the generator in batch mode:
   ```bash
   python3 generate_invoice.py /Users/myname/Invoices/ClientA/harvest-csv /Users/myname/Invoices/ClientA/clienta.json
   ```
3. Find generated DOCX and PDF files in the client's `Invoices/` directory
