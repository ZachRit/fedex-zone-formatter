"""
Smart PDF range finder - focuses on postal codes we need to cover.
"""

import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed

def check_url(url, timeout=8):
    """Check if a URL exists."""
    try:
        response = requests.head(url, timeout=timeout, allow_redirects=False,
                                 headers={'User-Agent': 'Mozilla/5.0'})
        return response.status_code == 200
    except:
        return False

def find_range_containing(postal_code):
    """Find the PDF range that contains a given postal code."""
    pc = int(postal_code)

    # Try different starting points near the postal code
    # The lower bound is typically a round number
    possible_lowers = []

    # Try multiples of 100 below the postal code
    base = (pc // 100) * 100
    for offset in range(0, 2000, 100):
        lower = base - offset
        if lower >= 0:
            possible_lowers.append(lower)

    # For each possible lower, try different upper bounds
    sizes = [100, 200, 300, 400, 500, 600, 700, 800, 900, 1000]

    for lower in possible_lowers:
        for size in sizes:
            upper = lower + size - 1
            if lower <= pc <= upper:  # Only test if postal code would be in range
                url = f'https://www.fedex.com/ratetools/documents2/{lower:05d}-{upper:05d}.pdf'
                if check_url(url):
                    return (lower, upper, url)

    return None

def main():
    # Load postal codes
    df = pd.read_excel('Origin_Postal_Codes.xlsx')
    postal_codes = sorted(set([int(str(pc).zfill(5)) for pc in df['Postal Codes'].tolist()]))

    print(f"Need to find ranges for {len(postal_codes)} postal codes")
    print(f"Postal codes: {[f'{pc:05d}' for pc in postal_codes[:10]]}... to {postal_codes[-1]:05d}")
    print("=" * 60)

    found_ranges = {}  # lower -> (upper, url)
    postal_to_range = {}  # postal_code -> range

    # Process each postal code
    for i, pc in enumerate(postal_codes):
        print(f"[{i+1}/{len(postal_codes)}] Finding range for {pc:05d}...", end='', flush=True)

        # Check if we already have a range that covers this postal code
        already_covered = False
        for lower, (upper, url) in found_ranges.items():
            if lower <= pc <= upper:
                postal_to_range[pc] = (lower, upper)
                print(f" already covered by {lower:05d}-{upper:05d}")
                already_covered = True
                break

        if already_covered:
            continue

        # Find the range
        result = find_range_containing(pc)
        if result:
            lower, upper, url = result
            found_ranges[lower] = (upper, url)
            postal_to_range[pc] = (lower, upper)
            print(f" FOUND: {lower:05d}-{upper:05d}")
        else:
            print(f" NOT FOUND!")

    print("\n" + "=" * 60)
    print(f"Found {len(found_ranges)} unique PDF ranges:")
    print("=" * 60)

    # Sort and output all ranges
    sorted_ranges = sorted(found_ranges.items())
    urls = []
    for lower, (upper, url) in sorted_ranges:
        urls.append(url)
        print(url)

    # Save to file
    with open('valid_pdf_urls.txt', 'w') as f:
        for url in urls:
            f.write(url + '\n')

    print(f"\nSaved {len(urls)} URLs to valid_pdf_urls.txt")

    # Check coverage
    uncovered = [pc for pc in postal_codes if pc not in postal_to_range]
    if uncovered:
        print(f"\nWARNING: {len(uncovered)} postal codes not covered:")
        for pc in uncovered:
            print(f"  {pc:05d}")
    else:
        print(f"\nAll {len(postal_codes)} postal codes are covered!")

if __name__ == '__main__':
    main()
