# FedEx Zone Tools

Tools for parsing FedEx zone PDFs and generating rate sheets.

## Scripts

### parse_fedex_zones.py

Extracts destination ZIP ranges and zones from FedEx zone locator PDFs.

```bash
# Process all PDFs in 'inputs' directory
python parse_fedex_zones.py

# Process a single PDF
python parse_fedex_zones.py <pdf_file>

# Process single PDF with custom output directory
python parse_fedex_zones.py <pdf_file> <output_dir>
```

**Input:** FedEx zone locator PDF files (placed in `inputs/` directory)

**Output:** Excel files in `output/` directory named by origin ZIP range (e.g., `01700-01899.xlsx`)

Output file columns:
- `Start Postal Code` - Destination ZIP range start
- `End Postal Code` - Destination ZIP range end
- `Zone` - FedEx zone for that destination

---

### generate_rate_sheet.py

Generates rate sheets by combining zone data with a rate sheet template. Processes SSL/Postal Code pairs and produces one output file per SSL.

```bash
python generate_rate_sheet.py \
  --input <ssl_postal_codes.xlsx> \
  --rate-sheet <rate_template.xlsx> \
  --country-name "United States" \
  --country-symbol "US" \
  --client-name "ClientName" \
  --carrier "FedEx" \
  --carrier-account "123456"
```

**Arguments:**

| Argument | Description |
|----------|-------------|
| `--input` | Excel file with SSL and Postal Code columns |
| `--rate-sheet` | Rate sheet template with a `Zones` tab |
| `--country-name` | Country name (e.g., "United States") |
| `--country-symbol` | Country symbol (e.g., "US") |
| `--client-name` | Client name for output filename |
| `--carrier` | Carrier name (e.g., "FedEx") |
| `--carrier-account` | Carrier account number |

**Input Files:**

1. **SSL/Postal Code file** (`--input`)
   - Required columns: `SSL`, `Postal Code`
   - Each row maps an origin postal code to an SSL

2. **Rate sheet template** (`--rate-sheet`)
   - Must contain a `Zones` tab with headers on row 3
   - Expected columns: `Country Name`, `Country Symbol`, `Zone`, `City`, `Start Postal Code`, `End Postal Code`

**Output:**

- One file per SSL in the current directory
- Filename format: `YYYYMMDD-{SSL}-{clientName}-{carrier}-{carrierAccount}.xlsx`
- Example: `20260116-SSL001-Acme-FedEx-789012.xlsx`

**Processing:**

1. Groups input by SSL
2. For each postal code, finds the matching zone file in `output/` (by ZIP range)
3. Loads zone data and adds Country Name/Symbol columns
4. Appends all zone data to the rate sheet template's `Zones` tab
5. Saves the output file

**Example:**

```bash
python generate_rate_sheet.py \
  --input arista_us_ssls.xlsx \
  --rate-sheet "United States International Rates 2025(1)_converted.xlsx" \
  --country-name "United States" \
  --country-symbol "UNITED_STATES" \
  --client-name "Arista" \
  --carrier "FedEx" \
  --carrier-account "758360300"
```

## Directory Structure

```
fedex_zones/
├── inputs/              # Place FedEx zone PDFs here
│   ├── archive/         # Successfully processed PDFs moved here
│   └── failed_parsing/  # Failed PDFs moved here
├── output/              # Generated zone files (XXXXX-XXXXX.xlsx)
├── parse_fedex_zones.py
├── generate_rate_sheet.py
└── README.md
```

## Requirements

- Python 3.10+
- pandas
- openpyxl
- pdfplumber (for PDF parsing)

Install dependencies:

```bash
pip install pandas openpyxl pdfplumber
```
