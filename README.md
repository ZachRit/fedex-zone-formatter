# FedEx Rate Sheet Tool

A unified CLI tool for parsing FedEx zone PDFs and generating rate sheets.

## Installation

```bash
pip install -r requirements.txt
```

## Usage

The tool provides five commands:

### 1. Find PDF URLs for Postal Codes

Find valid FedEx zone PDF URLs for a list of postal codes:

```bash
python fedex_rate_tool.py find-pdfs --input inputs/postal_codes.xlsx --output urls.txt
```

**Arguments:**
| Argument | Description |
|----------|-------------|
| `--input`, `-i` | Excel file with postal codes (columns: `Postal Codes`, `Postal Code`, or `ZIP`) |
| `--output`, `-o` | Output file for URLs (default: `valid_pdf_urls.txt`) |

### 2. Parse US Zone PDFs

Parse FedEx zone PDFs for US destinations:

```bash
# Process all PDFs in a directory
python fedex_rate_tool.py parse-us-zones --input inputs/ --output outputs/

# Process a single PDF
python fedex_rate_tool.py parse-us-zones --input 01700-01899.pdf --output outputs/
```

**Arguments:**
| Argument | Description |
|----------|-------------|
| `--input`, `-i` | PDF file or directory containing PDFs |
| `--output`, `-o` | Output directory (default: `outputs`) |

**Output:** Excel files named by ZIP range (e.g., `01700-01899.xlsx`) with columns:
- `Start Postal Code` - Destination ZIP range start
- `End Postal Code` - Destination ZIP range end
- `Zone` - FedEx zone for that destination

### 3. Parse Canadian Rate PDFs

Parse FedEx domestic rate PDFs for Canada:

```bash
python fedex_rate_tool.py parse-ca-rates \
  --input inputs/CA_EN_2026_Domestic_Rate_Guide_Express.pdf \
  --origin M5V \
  --output outputs/CA_Express_Rates_2026.xlsx
```

**Arguments:**
| Argument | Description |
|----------|-------------|
| `--input`, `-i` | Path to Canadian rate PDF |
| `--origin`, `-p` | Origin postal code (FSA) for zone calculations |
| `--output`, `-o` | Output Excel file (default: `CA_Express_Rates.xlsx`) |

**Output:** Excel file with:
- `Zones` tab - Destination postal codes with calculated zones
- Service tabs (First Overnight, Priority Overnight, etc.) - Rate tables by weight and zone

### 4. Generate Rate Sheets

Generate rate sheets by combining zone data with a template:

```bash
python fedex_rate_tool.py generate \
  --ssl-file inputs/arista_us_ssls.xlsx \
  --template template.xlsx \
  --zones-dir outputs \
  --country-name "United States" \
  --country-symbol "US" \
  --client-name "Arista" \
  --carrier "FedEx" \
  --carrier-account "758360300" \
  --output outputs/
```

**Arguments:**
| Argument | Description |
|----------|-------------|
| `--ssl-file` | Excel file with `SSL` and `Postal Code` columns |
| `--template` | Rate sheet template with a `Zones` tab |
| `--zones-dir` | Directory containing zone files (default: `outputs`) |
| `--country-name` | Country name (default: `United States`) |
| `--country-symbol` | Country symbol (default: `US`) |
| `--client-name` | Client name for output filename |
| `--carrier` | Carrier name (default: `FedEx`) |
| `--carrier-account` | Carrier account number |
| `--output`, `-o` | Output directory (default: `outputs`) |

**Output:** One file per SSL: `YYYYMMDD-{SSL}-{clientName}-{carrier}-{carrierAccount}.xlsx`

### 5. Fix Rate Sheets

Clean and deduplicate rate sheets:

```bash
python fedex_rate_tool.py fix --input outputs/fix_zones/ --output outputs/cleaned/
```

**Arguments:**
| Argument | Description |
|----------|-------------|
| `--input`, `-i` | Directory containing rate sheets to fix |
| `--output`, `-o` | Output directory for cleaned files |

**Operations:**
- Fixes zone header formatting (`Zone 02` -> `Zone 2`)
- Deduplicates the Zones tab
- Moves originals to `{input}/processed/`

## Directory Structure

```
fedex_zones/
├── fedex_rate_tool.py      # Unified CLI tool
├── requirements.txt
├── README.md
├── .gitignore
├── inputs/                  # Input PDFs and data files (gitignored)
│   ├── *.pdf               # FedEx rate guide PDFs
│   ├── arista_us_ssls.xlsx # SSL/postal code mappings
│   ├── ca_zones.xlsx       # Canadian zone mappings
│   └── ...
├── outputs/                 # Generated files (gitignored)
│   ├── *.xlsx              # Generated zone and rate files
│   └── fix_zones/          # Files to be fixed
└── archive/                 # Old scripts for reference (gitignored)
    ├── find_pdf_ranges.py
    ├── find_ranges_*.py
    ├── parse_fedex_zones.py
    ├── parse_ca_fedex_rate_sheets.py
    ├── generate_rate_sheet.py
    └── fix_rate_sheets.py
```

## Requirements

- Python 3.10+
- pandas
- openpyxl
- pdfplumber
- requests

## Quick Examples

```bash
# Parse Canadian 2026 rates for Toronto (M5V)
python fedex_rate_tool.py parse-ca-rates \
  -i inputs/CA_EN_2026_Domestic_Rate_Guide_Express.pdf \
  -p M5V \
  -o outputs/CA_Express_Rates_2026_M5V.xlsx

# Find URLs for all postal codes in a file
python fedex_rate_tool.py find-pdfs \
  -i inputs/Origin_Postal_Codes.xlsx \
  -o urls.txt

# Process all US zone PDFs
python fedex_rate_tool.py parse-us-zones \
  -i inputs/ \
  -o outputs/

# Generate rate sheets for Arista
python fedex_rate_tool.py generate \
  --ssl-file inputs/arista_us_ssls.xlsx \
  --template "inputs/United States International Rates 2025(1)_converted.xlsx" \
  --client-name Arista \
  --carrier-account 758360300 \
  --output outputs/
```
