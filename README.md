# Invoice Generator

## Prerequisites

```
pip install python-docx
```

## Configuration

Edit `config.json` to set invoice details:

- **rate** — hourly rate
- **output_prefix** — output filename prefix
- **sender** — name, address, email
- **invoice** — date, number, description (for)
- **bill_to** — company name lines, address lines
- **footer** — payable to, contact name, contact info

## Usage

```
python3 generate_invoice.py <harvest_csv_filename>
```

Example:

```
python3 generate_invoice.py harvest_time_report_from2026-02-16to2026-02-28.csv
```

Generates both:

- `Firstname-Lastname-invoice-2026-1.docx` — Word document
- `Firstname-Lastname-invoice-2026-1.pdf` — PDF (converted via Pages)
