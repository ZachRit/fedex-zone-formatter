"""
Find all FedEx zone PDF ranges by sequential testing.
"""

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

def create_session():
    """Create a requests session with retry logic."""
    session = requests.Session()
    retries = Retry(total=3, backoff_factor=0.5, status_forcelist=[500, 502, 503, 504])
    adapter = HTTPAdapter(max_retries=retries)
    session.mount('https://', adapter)
    session.mount('http://', adapter)
    return session

def check_pdf_exists(session, start, end):
    """Check if a PDF exists for the given range."""
    url = f'https://www.fedex.com/ratetools/documents2/{start:05d}-{end:05d}.pdf'
    try:
        response = session.head(url, timeout=15, allow_redirects=True)
        content_type = response.headers.get('Content-Type', '')
        # Check for PDF content type or 200 status with no error page
        if response.status_code == 200 and ('pdf' in content_type.lower() or 'octet' in content_type.lower()):
            return True, url
        return False, url
    except Exception as e:
        print(f"  Error checking {url}: {e}")
        return False, url

def find_next_range(session, lower_bound):
    """
    Find the next valid PDF range starting from lower_bound.
    Try incrementing upper bound by 100, 200, 300, etc. until we find a valid PDF.
    """
    # Try upper bounds in increments of 100
    for increment in range(100, 1100, 100):  # Try up to 1000 range
        upper_bound = lower_bound + increment - 1
        exists, url = check_pdf_exists(session, lower_bound, upper_bound)
        if exists:
            return (lower_bound, upper_bound, url)
    return None

def find_all_ranges(postal_codes, start_lower=0, known_ranges=None):
    """
    Find all PDF ranges needed to cover the given postal codes.

    Args:
        postal_codes: List of postal codes to cover (as integers)
        start_lower: Starting lower bound for search
        known_ranges: List of known (lower, upper) tuples to skip

    Returns:
        List of (lower, upper, url) tuples for all found ranges
    """
    session = create_session()

    found_ranges = []
    current_lower = start_lower
    max_postal = max(postal_codes)

    print(f"Finding ranges from {current_lower:05d} to cover up to {max_postal:05d}")
    print("=" * 60)

    while current_lower <= max_postal:
        print(f"\nSearching for range starting at {current_lower:05d}...")

        result = find_next_range(session, current_lower)

        if result:
            lower, upper, url = result
            found_ranges.append(result)
            print(f"  FOUND: {lower:05d}-{upper:05d}.pdf")

            # Check which postal codes this covers
            covered = [pc for pc in postal_codes if lower <= pc <= upper]
            if covered:
                print(f"  Covers postal codes: {', '.join(str(pc).zfill(5) for pc in covered)}")

            # Move to next range
            current_lower = upper + 1
        else:
            print(f"  No range found starting at {current_lower:05d}, skipping...")
            # Try jumping forward
            current_lower += 100

    return found_ranges

def main():
    import pandas as pd

    # Load postal codes
    df = pd.read_excel('Origin_Postal_Codes.xlsx')
    postal_codes = sorted([int(str(pc).zfill(5)) for pc in df['Postal Codes'].tolist()])

    print(f"Need to cover {len(postal_codes)} postal codes")
    print(f"Range: {min(postal_codes):05d} to {max(postal_codes):05d}")
    print()

    # Start from 00000 to find all ranges
    # We know 01700-01899 is the first range that covers 01801
    ranges = find_all_ranges(postal_codes, start_lower=0)

    print("\n" + "=" * 60)
    print("SUMMARY - All found PDF ranges:")
    print("=" * 60)

    for lower, upper, url in ranges:
        print(f"{lower:05d}-{upper:05d}: {url}")

    # Check coverage
    all_covered = []
    for lower, upper, _ in ranges:
        all_covered.extend([pc for pc in postal_codes if lower <= pc <= upper])

    uncovered = set(postal_codes) - set(all_covered)
    if uncovered:
        print(f"\nWARNING: {len(uncovered)} postal codes not covered:")
        for pc in sorted(uncovered):
            print(f"  {pc:05d}")
    else:
        print(f"\nAll {len(postal_codes)} postal codes are covered!")

    # Save URLs to file
    with open('pdf_urls.txt', 'w') as f:
        for lower, upper, url in ranges:
            f.write(f"{url}\n")
    print(f"\nURLs saved to pdf_urls.txt")

if __name__ == '__main__':
    main()
