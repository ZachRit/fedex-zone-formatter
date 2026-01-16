"""
Find FedEx zone PDF ranges using concurrent requests.
"""

import requests
from concurrent.futures import ThreadPoolExecutor, as_completed
import sys

def check_range(session, lower, upper):
    """Check if a PDF range exists."""
    url = f'https://www.fedex.com/ratetools/documents2/{lower:05d}-{upper:05d}.pdf'
    try:
        response = session.head(url, timeout=10, allow_redirects=False)
        # 200 means PDF exists, 302/404 means not found
        return response.status_code == 200, lower, upper, url
    except Exception as e:
        return False, lower, upper, url

def find_range_for_lower(lower):
    """Find the valid range for a given lower bound."""
    session = requests.Session()
    session.headers.update({'User-Agent': 'Mozilla/5.0'})

    sizes = [100, 200, 300, 400, 500, 600, 700, 800, 900, 1000]
    for size in sizes:
        upper = lower + size - 1
        exists, _, _, url = check_range(session, lower, upper)
        if exists:
            return (lower, upper, url)
    return None

def main():
    # Known starting points and ranges
    known_ranges = [
        (1700, 1899),   # 01700-01899
        (1900, 1999),   # 01900-01999
        (2000, 2499),   # 02000-02499
        (2500, 2599),   # 02500-02599
    ]

    # Generate candidate lower bounds
    # Start from 2600 and go up to 99300 in increments
    # We'll test each potential starting point

    found_ranges = list(known_ranges)
    current_lower = 2600
    max_postal = 99300

    print("Finding FedEx zone PDF ranges...")
    print("Known ranges:", [f"{l:05d}-{u:05d}" for l, u in known_ranges])
    print("=" * 60)

    while current_lower <= max_postal:
        print(f"Testing from {current_lower:05d}...", end='', flush=True)
        result = find_range_for_lower(current_lower)

        if result:
            lower, upper, url = result
            found_ranges.append((lower, upper))
            print(f" FOUND: {lower:05d}-{upper:05d}")
            current_lower = upper + 1
        else:
            print(f" not found, skipping")
            current_lower += 100

    print("\n" + "=" * 60)
    print(f"Found {len(found_ranges)} total PDF ranges:")
    print("=" * 60)

    # Output all URLs
    base_url = "https://www.fedex.com/ratetools/documents2/"
    urls = []
    for lower, upper in sorted(found_ranges):
        url = f"{base_url}{lower:05d}-{upper:05d}.pdf"
        urls.append(url)
        print(url)

    # Save to file
    with open('valid_pdf_urls.txt', 'w') as f:
        for url in urls:
            f.write(url + '\n')

    print(f"\nSaved {len(urls)} URLs to valid_pdf_urls.txt")

    # Check coverage of postal codes
    import pandas as pd
    df = pd.read_excel('Origin_Postal_Codes.xlsx')
    postal_codes = sorted([int(str(pc).zfill(5)) for pc in df['Postal Codes'].tolist()])

    covered = []
    uncovered = []
    for pc in postal_codes:
        is_covered = False
        for lower, upper in found_ranges:
            if lower <= pc <= upper:
                is_covered = True
                break
        if is_covered:
            covered.append(pc)
        else:
            uncovered.append(pc)

    print(f"\nPostal code coverage: {len(covered)}/{len(postal_codes)}")
    if uncovered:
        print(f"Uncovered postal codes: {[f'{pc:05d}' for pc in uncovered]}")

if __name__ == '__main__':
    main()
