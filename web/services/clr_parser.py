"""
CLR Parser — Reads Amazon Category Listing Report (.xlsm) files
and extracts product data, ITK values, and family groupings.
"""
import openpyxl
from collections import Counter


def parse_clr(filepath):
    """Parse a CLR file and return structured product data."""
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb['Template']

    # Read headers from Row 4
    headers = []
    for cell in ws[4]:
        headers.append(str(cell.value).strip() if cell.value else None)

    # Read attribute references from Row 5
    attr_refs = []
    for cell in ws[5]:
        attr_refs.append(str(cell.value).strip() if cell.value else None)

    # Find key column indices
    col_map = _find_columns(headers)

    # Read product data starting from Row 6+
    # Find the actual first data row (skip attribute reference echoes)
    products = []
    for row in ws.iter_rows(min_row=6, values_only=True):
        row_data = list(row)

        # Skip empty rows
        sku_idx = col_map.get('sku')
        if sku_idx is None or sku_idx >= len(row_data) or not row_data[sku_idx]:
            continue

        sku_val = str(row_data[sku_idx]).strip()

        # Skip attribute reference echo rows
        if sku_val.startswith('contribution_sku') or '#' in sku_val and '.value' in sku_val:
            continue

        # Skip rows where product_type looks like an attribute reference
        pt_idx = col_map.get('product_type')
        if pt_idx is not None and pt_idx < len(row_data) and row_data[pt_idx]:
            pt_val = str(row_data[pt_idx]).strip()
            if 'product_type#' in pt_val and '.value' in pt_val:
                continue

        product = {
            'row_data': row_data,
            'sku': sku_val,
            'product_type': _get_val(row_data, col_map.get('product_type')),
            'itk': _get_val(row_data, col_map.get('itk')),
            'parentage': _get_val(row_data, col_map.get('parentage')),
            'parent_sku': _get_val(row_data, col_map.get('parent_sku')),
        }
        products.append(product)

    wb.close()

    # Categorize products
    parents = []
    children = []
    standalone = []

    for p in products:
        parentage = (p['parentage'] or '').lower()
        if 'parent' in parentage and 'child' not in parentage:
            parents.append(p)
        elif 'child' in parentage and p['parent_sku']:
            children.append(p)
        else:
            standalone.append(p)

    return {
        'headers': headers,
        'attr_refs': attr_refs,
        'products': products,
        'parents': parents,
        'children': children,
        'standalone': standalone,
        'col_map': col_map,
        'total_products': len(products),
        'total_parents': len(parents),
        'total_children': len(children),
        'total_standalone': len(standalone),
    }


def extract_itk_summary(clr_data):
    """Extract unique ITK values with counts, sorted by count descending.

    For display, uses the parenthesized slug when available in the ITK value.
    e.g. "Blended Vitamin & Mineral Supplements (multiple-vitamin-mineral-combinations)"
         -> display as "(multiple-vitamin-mineral-combinations)"
    Browse-path ITKs without a slug are shown as-is.
    """
    import re
    itk_counter = Counter()
    # Map raw ITK -> display name
    itk_display = {}

    for p in clr_data['products']:
        itk = p['itk']
        if itk:
            itk_counter[itk] += 1
            if itk not in itk_display:
                # Extract parenthesized slug if present
                paren_match = re.search(r'\(([^)]+)\)\s*$', itk)
                if paren_match:
                    itk_display[itk] = f"({paren_match.group(1)})"
                else:
                    itk_display[itk] = itk
        else:
            itk_counter['(no ITK)'] += 1
            itk_display['(no ITK)'] = '(no ITK)'

    # Return tuples of (display_name, raw_itk, count)
    result = []
    for itk, count in sorted(itk_counter.items(), key=lambda x: x[1], reverse=True):
        result.append((itk_display.get(itk, itk), itk, count))

    return result


def _find_columns(headers):
    """Find indices of key columns by header name."""
    col_map = {}
    header_lower = [(h.lower() if h else '') for h in headers]

    mappings = {
        'sku': ['sku'],
        'product_type': ['product type'],
        'itk': ['item type keyword'],
        'parentage': ['parentage level', 'parentage'],
        'parent_sku': ['parent sku'],
    }

    for key, names in mappings.items():
        for i, h in enumerate(header_lower):
            if h in names:
                col_map[key] = i
                break

    return col_map


def _get_val(row_data, idx):
    """Safely get a string value from row data at the given index."""
    if idx is None or idx >= len(row_data) or row_data[idx] is None:
        return None
    val = str(row_data[idx]).strip()
    return val if val else None
