"""
Fast PDF range finder using HEAD requests with short timeout.
"""

import requests
import sys

def check_url(session, lower, upper):
    """Check if URL exists using HEAD request."""
    url = f'https://www.fedex.com/ratetools/documents2/{lower:05d}-{upper:05d}.pdf'
    try:
        response = session.head(url, timeout=5, allow_redirects=True)
        # 200 = exists, 404 or redirect to error page = doesn't exist
        if response.status_code == 200:
            ct = response.headers.get('Content-Type', '')
            if 'pdf' in ct.lower() or 'octet' in ct.lower() or not ct:
                return True
        return False
    except:
        return False

def find_upper_bound(session, lower):
    """Find the upper bound for a range starting at lower."""
    # Try increments of 100 up to 1000
    for size in [100, 200, 300, 400, 500, 600, 700, 800, 900, 1000]:
        upper = lower + size - 1
        if check_url(session, lower, upper):
            return upper
    return None

def main():
    session = requests.Session()
    session.headers.update({'User-Agent': 'Mozilla/5.0'})

    # Start from first known range
    current_lower = 1700
    max_postal = 99300  # A bit beyond max postal code

    found_ranges = []

    print("Finding FedEx zone PDF ranges...")
    print("=" * 60)

    while current_lower <= max_postal:
        sys.stdout.write(f"\rSearching from {current_lower:05d}...")
        sys.stdout.flush()

        upper = find_upper_bound(session, current_lower)

        if upper:
            range_str = f"{current_lower:05d}-{upper:05d}"
            found_ranges.append(range_str)
            print(f"\rFOUND: {range_str}.pdf" + " " * 20)
            current_lower = upper + 1
        else:
            # No range found, skip forward
            current_lower += 100

    print("\n" + "=" * 60)
    print(f"Found {len(found_ranges)} PDF ranges:")
    print("=" * 60)

    # Output URLs
    base_url = "https://www.fedex.com/ratetools/documents2/"
    for r in found_ranges:
        print(f"{base_url}{r}.pdf")

    # Save to file
    with open('valid_pdf_urls.txt', 'w') as f:
        for r in found_ranges:
            f.write(f"{base_url}{r}.pdf\n")

    print(f"\nSaved {len(found_ranges)} URLs to valid_pdf_urls.txt")

if __name__ == '__main__':
    main()
