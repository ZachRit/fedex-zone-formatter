"""
FedEx Zone PDF Parser

Extracts destination ZIP ranges and zones from FedEx zone locator PDFs
and outputs to xlsx format.

Usage:
    python parse_fedex_zones.py                    # Process all PDFs in 'inputs' directory
    python parse_fedex_zones.py <pdf_file>         # Process a single PDF file
    python parse_fedex_zones.py <pdf_file> <dir>   # Process single PDF with custom output dir
"""

import re
import shutil
import sys
from pathlib import Path

import pdfplumber
import pandas as pd


def extract_text_from_pdf(pdf_path: str) -> str:
    """Extract all text from PDF using pdfplumber."""
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() or ""
    return text


def parse_contiguous_us(text: str) -> list[tuple[str, str]]:
    """
    Parse the Contiguous U.S. section.
    Format: ZIP range followed by zone (single value).
    Returns list of (zip_range, zone) tuples.
    """
    results = []

    # Pattern matches: 5 digits or 5 digits-5 digits, followed by zone (number, NA, or *)
    pattern = r'(\d{5}(?:-\d{5})?)\s+(\d+|NA|\*)'

    # Find the Contiguous U.S. section - it ends at Alaska, Hawaii section
    contiguous_match = re.search(r'Contiguous U\.S\.(.*?)Alaska,\s*Hawaii', text, re.DOTALL | re.IGNORECASE)

    if contiguous_match:
        contiguous_text = contiguous_match.group(1)
        matches = re.findall(pattern, contiguous_text)

        for zip_range, zone in matches:
            # Skip entries that look like they have two zone values (Alaska/Hawaii format)
            # These would be in the wrong section
            results.append((zip_range, zone))

    return results


def parse_alaska_hawaii_pr(text: str) -> list[tuple[str, str]]:
    """
    Parse the Alaska, Hawaii, and Puerto Rico section.
    Format: ZIP range followed by Express zone and Ground zone.
    Returns list of (zip_range, express_zone) tuples (ignoring ground zone).
    """
    results = []

    # Pattern matches: ZIP range, then two zone values (Express and Ground)
    pattern = r'(\d{5}(?:-\d{5})?)\s+(\d+|NA|\*)\s+(\d+|NA|\*)'

    # Find the Alaska, Hawaii, and Puerto Rico section
    alaska_match = re.search(r'Alaska,\s*Hawaii,?\s*and\s*Puerto\s*Rico(.*?)$', text, re.DOTALL | re.IGNORECASE)

    if alaska_match:
        alaska_text = alaska_match.group(1)
        matches = re.findall(pattern, alaska_text)

        for zip_range, express_zone, ground_zone in matches:
            # Use Express zone only (second column)
            results.append((zip_range, express_zone))

    return results


def split_zip_range(zip_range: str) -> tuple[str, str]:
    """
    Split ZIP range into start and end.
    "00000-00399" -> ("00000", "00399")
    "96700" -> ("96700", "96700")
    """
    if '-' in zip_range:
        start, end = zip_range.split('-')
        return (start.strip(), end.strip())
    else:
        return (zip_range.strip(), zip_range.strip())


def normalize_zone(zone: str) -> str:
    """
    Normalize zone value.
    NA or * -> empty string
    Otherwise return as-is.
    """
    if zone.upper() == 'NA' or zone == '*':
        return ''
    return zone


def validate_fedex_pdf(text: str) -> bool:
    """
    Validate that the PDF appears to be a FedEx zone locator.

    Returns True if valid, False otherwise.
    """
    # Check for key indicators of a FedEx zone locator PDF
    has_fedex = 'fedex' in text.lower()
    has_zone = 'zone' in text.lower()
    has_contiguous = 'contiguous' in text.lower() or 'u.s.' in text.lower()

    # Must have at least some ZIP-zone data
    zip_pattern = r'\d{5}(?:-\d{5})?\s+\d+'
    has_zip_data = bool(re.search(zip_pattern, text))

    return has_fedex and has_zone and has_zip_data


def process_pdf(input_path: str, output_dir: str = 'output') -> str:
    """
    Process a FedEx zone PDF and output to xlsx.

    Args:
        input_path: Path to input PDF file
        output_dir: Directory for output files (created if doesn't exist)

    Returns:
        Path to output xlsx file

    Raises:
        ValueError: If PDF is not a valid FedEx zone locator format
    """
    input_path = Path(input_path)
    output_dir = Path(output_dir)

    # Create output directory if needed
    output_dir.mkdir(parents=True, exist_ok=True)

    # Extract text from PDF
    text = extract_text_from_pdf(str(input_path))

    # Validate PDF format
    if not validate_fedex_pdf(text):
        raise ValueError("PDF does not appear to be a valid FedEx zone locator")

    # Parse both sections
    contiguous_data = parse_contiguous_us(text)
    alaska_data = parse_alaska_hawaii_pr(text)

    # Validate we got some data
    if not contiguous_data and not alaska_data:
        raise ValueError("No zone data could be extracted from PDF")

    # Combine all data
    all_data = contiguous_data + alaska_data

    # Build output rows
    rows = []
    for zip_range, zone in all_data:
        start_zip, end_zip = split_zip_range(zip_range)
        normalized_zone = normalize_zone(zone)
        rows.append({
            'Start Postal Code': start_zip,
            'End Postal Code': end_zip,
            'Zone': normalized_zone
        })

    # Create DataFrame
    df = pd.DataFrame(rows)

    # Sort by Start Postal Code
    df = df.sort_values('Start Postal Code').reset_index(drop=True)

    # Output filename = input filename with .xlsx extension
    output_filename = input_path.stem + '.xlsx'
    output_path = output_dir / output_filename

    # Write to xlsx, keeping ZIP codes as text to preserve leading zeros
    with pd.ExcelWriter(str(output_path), engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Zones')
        # Format columns as text to preserve leading zeros
        worksheet = writer.sheets['Zones']
        for row in worksheet.iter_rows(min_row=2, max_row=len(df) + 1, min_col=1, max_col=2):
            for cell in row:
                cell.number_format = '@'  # Text format

    return str(output_path)


def process_directory(input_dir: str = 'inputs', output_dir: str = 'output') -> dict:
    """
    Process all PDF files in a directory.

    Successfully parsed files are moved to input_dir/archive.
    Failed files are moved to input_dir/failed_parsing.

    Args:
        input_dir: Directory containing PDF files to process
        output_dir: Directory for output xlsx files

    Returns:
        Dictionary with 'success' and 'failed' counts
    """
    input_dir = Path(input_dir)
    archive_dir = input_dir / 'archive'
    failed_dir = input_dir / 'failed_parsing'

    # Create directories if needed
    input_dir.mkdir(parents=True, exist_ok=True)
    archive_dir.mkdir(parents=True, exist_ok=True)
    failed_dir.mkdir(parents=True, exist_ok=True)

    # Find all PDF files in input directory (not in subdirectories)
    pdf_files = list(input_dir.glob('*.pdf'))

    if not pdf_files:
        print(f"No PDF files found in {input_dir}")
        return {'success': 0, 'failed': 0}

    print(f"Found {len(pdf_files)} PDF file(s) to process")
    print()

    success_count = 0
    failed_count = 0

    for pdf_path in pdf_files:
        print(f"Processing: {pdf_path.name}")
        try:
            output_path = process_pdf(str(pdf_path), output_dir)
            print(f"  Output: {output_path}")

            # Move to archive
            dest_path = archive_dir / pdf_path.name
            shutil.move(str(pdf_path), str(dest_path))
            print(f"  Moved to: {dest_path}")
            success_count += 1

        except Exception as e:
            print(f"  ERROR: {e}")

            # Move to failed_parsing
            dest_path = failed_dir / pdf_path.name
            shutil.move(str(pdf_path), str(dest_path))
            print(f"  Moved to: {dest_path}")
            failed_count += 1

        print()

    return {'success': success_count, 'failed': failed_count}


def main():
    """Main entry point."""
    if len(sys.argv) < 2:
        # No arguments: process all PDFs in 'inputs' directory
        print("Processing all PDFs in 'inputs' directory...")
        print()
        results = process_directory('inputs', 'output')
        print("=" * 40)
        print(f"Complete: {results['success']} succeeded, {results['failed']} failed")
        return

    pdf_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else 'output'

    if not Path(pdf_path).exists():
        print(f"Error: File not found: {pdf_path}")
        sys.exit(1)

    print(f"Processing: {pdf_path}")
    output_path = process_pdf(pdf_path, output_dir)
    print(f"Output written to: {output_path}")


if __name__ == '__main__':
    main()
