# FedEx Zone Chart Download Process

## Quick Method (Direct URL)

The FedEx zone charts are available at predictable URLs based on ZIP code ranges.

### URL Pattern
```
https://www.fedex.com/ratetools/documents2/[START]-[END].pdf
```

### How to Determine the ZIP Range

The ZIP codes are grouped by their **first 2 digits**. Each file covers a 300-code range:

**Formula:**
1. Take the first 2 digits of the ZIP code
2. Append `000` for the start
3. Append `299` for the end

**Examples:**
- ZIP 43201 → First 2 digits: `43` → File: `43000-43299.pdf`
- ZIP 19966 → First 2 digits: `19` → File: `19000-19299.pdf`

**Direct URLs:**
- `https://www.fedex.com/ratetools/documents2/43000-43299.pdf`
- `https://www.fedex.com/ratetools/documents2/19000-19299.pdf`

### Quick Reference Formula
```
ZIP code XXXXX → File: XX000-XX299.pdf
```

## Alternative: Form Method (Slower)

If the direct URL doesn't work, use the form:

1. Navigate to: https://www.fedex.com/ratetools/RateToolsMain.do
2. Select "Yes" for zone information
3. Select "Domestic"
4. Enter the ZIP code
5. The form will populate `downloadFileName` with the correct PDF path
6. Navigate directly to: `https://www.fedex.com/ratetools/[downloadFileName]`

## JavaScript to Get Filename from Form

After filling the form, run this in the browser console to get the download filename:
```javascript
document.querySelector('input[name="downloadFileName"]').value
```

## Batch Processing Notes

For processing multiple ZIP codes:
1. First determine the unique ZIP ranges needed (many ZIPs share the same range file)
2. Download each unique range file once
3. Files cover 300 consecutive ZIP codes each
