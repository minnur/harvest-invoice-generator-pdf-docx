import csv
import json
import sys
import os
import subprocess
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

if len(sys.argv) < 2:
    print('Usage: python3 generate_invoice.py <harvest_csv_filename>')
    sys.exit(1)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(SCRIPT_DIR, 'config.json')) as f:
    cfg = json.load(f)

RATE = cfg['rate']
HARVEST_CSV = os.path.join(SCRIPT_DIR, sys.argv[1])
OUTPUT_PREFIX = cfg['output_prefix']
INVOICE_NUM = cfg['invoice']['number']
OUTPUT_DOCX = os.path.join(SCRIPT_DIR, f'{OUTPUT_PREFIX}-{INVOICE_NUM}.docx')
OUTPUT_PDF = os.path.join(SCRIPT_DIR, f'{OUTPUT_PREFIX}-{INVOICE_NUM}.pdf')

# --- Step 1: Read Harvest CSV and build invoice rows ---
invoice_rows = []
with open(HARVEST_CSV, 'r') as f:
    reader = csv.reader(f)
    header = next(reader)
    for row in reader:
        date = row[0]
        project_code = row[3].strip()
        notes = row[5].strip()
        hours = float(row[6])
        ref_url = row[14].strip() if len(row) > 14 else ''
        amount = round(hours * RATE, 2)

        # Build description: Project Code first, then notes, then reference
        parts = []
        if project_code:
            parts.append(f'[{project_code}]')
        parts.append(notes)
        desc = ' '.join(parts)
        if ref_url:
            desc += f' Reference: {ref_url}'

        invoice_rows.append({
            'date': date,
            'description': desc,
            'project_code': project_code,
            'notes': notes,
            'ref_url': ref_url,
            'rate': f'${RATE:.2f}',
            'hours': hours,
            'amount': amount,
        })

total_hours = sum(r['hours'] for r in invoice_rows)
total_amount = sum(r['amount'] for r in invoice_rows)

print(f'{len(invoice_rows)} line items, {total_hours:.2f} hours, ${total_amount:,.2f}')

# --- Step 3: Generate DOCX matching Pages template ---
doc = Document()

for section in doc.sections:
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(10)
style.paragraph_format.space_before = Pt(0)
style.paragraph_format.space_after = Pt(2)

def remove_table_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else tbl._add_tblPr()
    borders = tblPr.find(qn('w:tblBorders'))
    if borders is not None:
        tblPr.remove(borders)
    # Add empty borders to ensure no borders
    borders_elm = tblPr.makeelement(qn('w:tblBorders'), {})
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = borders_elm.makeelement(qn(f'w:{border_name}'), {
            qn('w:val'): 'none', qn('w:sz'): '0', qn('w:space'): '0', qn('w:color'): 'auto'
        })
        borders_elm.append(border)
    tblPr.append(borders_elm)

def shade_cell(cell, color='F2F2F2'):
    shading = cell._element.get_or_add_tcPr()
    shading_elm = shading.makeelement(qn('w:shd'), {qn('w:fill'): color, qn('w:val'): 'clear'})
    shading.append(shading_elm)

# ===== TITLE =====
title = doc.add_paragraph()
title.paragraph_format.space_after = Pt(16)
run = title.add_run('INVOICE')
run.bold = True
run.font.size = Pt(32)

# ===== HEADER (2-column borderless table) =====
header_table = doc.add_table(rows=1, cols=2)
header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
remove_table_borders(header_table)

# Left: sender info
left_cell = header_table.cell(0, 0)
left_cell.width = Inches(3.4)
p = left_cell.paragraphs[0]
p.paragraph_format.space_after = Pt(0)
run = p.add_run(cfg['sender']['name'])
run.bold = True
run.font.size = Pt(11)
for line in cfg['sender']['address'] + [f"Email: {cfg['sender']['email']}"]:
    p.add_run('\n')
    run = p.add_run(line)
    run.font.size = Pt(10)

# Right: invoice details + bill to
right_cell = header_table.cell(0, 1)
right_cell.width = Inches(3.4)
p = right_cell.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
p.paragraph_format.space_after = Pt(0)

run = p.add_run('DATE: ')
run.font.size = Pt(10)
run.font.color.rgb = RGBColor(128, 128, 128)
run = p.add_run(cfg['invoice']['date'])
run.bold = True
run.font.size = Pt(10)

p.add_run('\n')
run = p.add_run('INVOICE#: ')
run.font.size = Pt(10)
run.font.color.rgb = RGBColor(128, 128, 128)
run = p.add_run(cfg['invoice']['number'])
run.bold = True
run.font.size = Pt(10)

p.add_run('\n')
run = p.add_run('FOR: ')
run.font.size = Pt(10)
run.font.color.rgb = RGBColor(128, 128, 128)
run = p.add_run(cfg['invoice']['for'])
run.font.size = Pt(10)

p.add_run('\n')
run = p.add_run('BILL TO:')
run.font.size = Pt(10)
run.font.color.rgb = RGBColor(128, 128, 128)

for line in cfg['bill_to']['company']:
    p.add_run('\n')
    run = p.add_run(line)
    run.bold = True
    run.font.size = Pt(10)

for line in cfg['bill_to']['address']:
    p.add_run('\n')
    run = p.add_run(line)
    run.font.size = Pt(10)

doc.add_paragraph('')  # spacer

# ===== INVOICE TABLE =====
table = doc.add_table(rows=1, cols=5)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.style = 'Table Grid'

col_widths = [Inches(1.0), Inches(2.8), Inches(0.8), Inches(0.9), Inches(1.0)]

# Header row
header_cells = table.rows[0].cells
for i, h in enumerate(['DATE', 'DESCRIPTION', 'RATE', 'HOURS', 'AMOUNT']):
    header_cells[i].width = col_widths[i]
    p = header_cells[i].paragraphs[0]
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(h)
    run.bold = True
    run.font.size = Pt(9)
    shade_cell(header_cells[i], 'E8E8E8')

header_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
header_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
header_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
header_cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

# Data rows
for r in invoice_rows:
    row = table.add_row()
    for i in range(5):
        row.cells[i].width = col_widths[i]

    # Date
    p = row.cells[0].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(r['date'])
    run.font.size = Pt(9)

    # Description
    p = row.cells[1].paragraphs[0]
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)

    if r['project_code']:
        run = p.add_run(f"[{r['project_code']}]")
        run.bold = True
        run.font.size = Pt(9)
        run = p.add_run(f" {r['notes']}")
        run.font.size = Pt(9)
    else:
        run = p.add_run(r['notes'])
        run.font.size = Pt(9)

    if r['ref_url']:
        p.add_run('\n')
        ref_run = p.add_run(f"Reference: {r['ref_url']}")
        ref_run.font.size = Pt(7)
        ref_run.font.color.rgb = RGBColor(120, 120, 120)

    # Rate
    p = row.cells[2].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(r['rate'])
    run.font.size = Pt(9)

    # Hours
    p = row.cells[3].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(f"{r['hours']:.2f}")
    run.font.size = Pt(9)

    # Amount
    p = row.cells[4].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(f"${r['amount']:.2f}")
    run.font.size = Pt(9)

# --- Summary rows ---
# Subtotal
row = table.add_row()
row.cells[0].merge(row.cells[2])
p = row.cells[0].paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
p.paragraph_format.space_before = Pt(4)
p.paragraph_format.space_after = Pt(4)
run = p.add_run('Subtotal')
run.bold = False
run.font.size = Pt(9)
run.font.color.rgb = RGBColor(128, 128, 128)

p = row.cells[3].paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(4)
p.paragraph_format.space_after = Pt(4)
run = p.add_run(f'{total_hours:.2f}')
run.bold = True
run.font.size = Pt(9)
run.font.color.rgb = RGBColor(0, 0, 0)

p = row.cells[4].paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(4)
p.paragraph_format.space_after = Pt(4)
run = p.add_run(f'${total_amount:,.2f}')
run.bold = True
run.font.size = Pt(9)
run.font.color.rgb = RGBColor(0, 0, 0)

# Tax Rate, Sales Tax, Other
for label in ['Tax Rate', 'Sales Tax', 'Other']:
    row = table.add_row()
    row.cells[0].merge(row.cells[3])
    p = row.cells[0].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(label)
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(128, 128, 128)

    p = row.cells[4].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run('-')
    run.font.size = Pt(9)

# Total
row = table.add_row()
row.cells[0].merge(row.cells[3])
p = row.cells[0].paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
p.paragraph_format.space_before = Pt(6)
p.paragraph_format.space_after = Pt(6)
run = p.add_run('Total')
run.bold = False
run.font.size = Pt(9)
run.font.color.rgb = RGBColor(128, 128, 128)

p = row.cells[4].paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(6)
p.paragraph_format.space_after = Pt(6)
run = p.add_run(f'${total_amount:,.2f}')
run.bold = True
run.font.size = Pt(9)
run.font.color.rgb = RGBColor(0, 0, 0)

# ===== FOOTER =====
doc.add_paragraph('')
p = doc.add_paragraph()
p.add_run('Make all checks payable to ')
run = p.add_run(cfg['footer']['payable_to'])
run.bold = True
p.add_run('.')

p = doc.add_paragraph()
p.add_run('If you have any questions concerning this invoice, contact ')
run = p.add_run(cfg['footer']['contact_name'])
run.bold = True

p = doc.add_paragraph()
p.add_run(cfg['footer']['contact_info'])

doc.save(OUTPUT_DOCX)

print(f'DOCX saved: {OUTPUT_DOCX}')

# Convert DOCX to PDF via Pages
applescript = f'''
tell application "Pages"
    set doc to open POSIX file "{OUTPUT_DOCX}"
    delay 2
    export doc to POSIX file "{OUTPUT_PDF}" as PDF
    close doc saving no
end tell
'''
subprocess.run(['osascript', '-e', applescript], check=True)
subprocess.run(['SetFile', '-a', 'e', OUTPUT_PDF], check=True)
print(f'PDF saved: {OUTPUT_PDF}')
