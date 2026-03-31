#!/usr/bin/env python3
"""
AMAZON DEALS PLANNER GENERATOR — Multi-Brand Universal Script
============================================================================
Creates comprehensive Amazon Deals planning spreadsheets for ANY brand.
Integrates Manage Inventory (FBA Raw Data) + Fee Preview files.
Correct fee calculations: Upfront = lump sum per deal, Variable = per unit % of revenue.

Usage:
    python amazon_deals_planner_generator.py --brand "YourBrand" \
        --inventory "path/to/inventory.csv" \
        --fees "path/to/fees.csv" \
        --output "path/to/output.xlsx"
"""

import argparse
import csv
import sys
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ═════════════════════════════════════════════════════════════════════════════
# CONFIGURATION & CONSTANTS
# ═════════════════════════════════════════════════════════════════════════════

# Colors
NAVY = "1F4E79"
BLUE = "2E75B6"
LT_BLUE = "BDD7EE"
YEL = "FFF2CC"
GRN = "E2EFDA"
RED_L = "FFE0E0"
WHT = "FFFFFF"
ALT = "F5F9FF"
GRY = "F2F2F2"
ORG = "FCE4D6"
LIGHT_ORANGE = "FCE4D6"
ORANGE_BG = "F4B942"

FONT_NAME = "Arial"
MAX_PRODUCT_ROWS = 150

# Column helpers
def col_letter(n):
    return get_column_letter(n)

def fill(c):
    return PatternFill("solid", fgColor=c, start_color=c)

def bdr(color="9DC3E6"):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def font(bold=False, color="000000", sz=10, italic=False):
    return Font(name=FONT_NAME, bold=bold, color=color, size=sz, italic=italic)

# ═════════════════════════════════════════════════════════════════════════════
# DATA LOADING & PARSING
# ═════════════════════════════════════════════════════════════════════════════

def clean_header(h):
    """Strip BOM, quotes, non-alphanumeric prefixes, and whitespace from CSV header."""
    import re
    h = h.strip().lstrip('\ufeff')
    # Remove any leading non-letter characters (BOM artifacts, ?, etc.)
    h = re.sub(r'^[^a-zA-Z]+', '', h)
    # Remove surrounding quotes
    h = h.strip('"').strip("'").strip()
    return h

def load_inventory(filepath):
    """Load Amazon Manage Inventory (FBA Active Listings Report)."""
    products = []
    try:
        for enc in ("utf-8-sig", "utf-8", "latin-1"):
            try:
                with open(filepath, encoding=enc) as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        cleaned = {clean_header(k): v.strip().strip('"') if v else "" for k, v in row.items() if k}
                        if cleaned.get("sku"):
                            products.append(cleaned)
                if products:
                    break
            except (UnicodeDecodeError, UnicodeError):
                continue
    except Exception as e:
        print(f"❌ Error loading inventory: {e}")
        sys.exit(1)
    return products

def load_fees(filepath):
    """Load Amazon Fee Preview file and map SKU -> FBA fee."""
    fees_map = {}
    try:
        for enc in ("utf-8-sig", "utf-8", "latin-1"):
            try:
                with open(filepath, encoding=enc) as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        cleaned = {clean_header(k): v.strip().strip('"') if v else "" for k, v in row.items() if k}
                        sku = cleaned.get("sku", "").strip()
                        if sku:
                            fee_str = cleaned.get("expected-fulfillment-fee-per-unit", "0").strip()
                            try:
                                fee = float(fee_str) if fee_str and fee_str != "--" else 0.0
                                fees_map[sku] = fee
                            except ValueError:
                                fees_map[sku] = 0.0
                if fees_map:
                    break
            except (UnicodeDecodeError, UnicodeError):
                continue
    except Exception as e:
        print(f"⚠️  Warning: Could not load fees file: {e}")
    return fees_map

# ═════════════════════════════════════════════════════════════════════════════
# CATEGORY & COST DETECTION (BRAND-AGNOSTIC)
# ═════════════════════════════════════════════════════════════════════════════

def detect_category(product_name, sku, brand_name):
    """
    Auto-detect product category from name/SKU patterns.
    Returns: (category_name, suggested_cogs)
    """
    n = product_name.lower()
    s = sku.upper()

    # Try to identify by keyword patterns
    # Adjust these patterns based on your products

    if any(x in n for x in ["bundle", "pack of", "set of"]):
        count = "bundle"
        if "12" in n or "12pk" in s:
            return ("Bundle — 12-Pack", 25.0)
        elif "3" in n or "3pk" in s:
            return ("Bundle — 3-Pack", 8.0)
        else:
            return ("Bundle", 6.0)

    if "single" in n or "1 pack" in n or "1pk" in s or "-1pk" in s:
        # Single/1-pack items
        if any(x in n for x in ["premium", "deluxe", "professional"]):
            return ("Premium — 1-Pack", 8.0)
        else:
            return ("Single — 1-Pack", 3.0)

    if "3" in n or "3pk" in s or "pack of 3" in n:
        return ("Multi-Pack — 3-Pack", 8.0)

    if "12" in n or "12pk" in s or "pack of 12" in n:
        return ("Multi-Pack — 12-Pack", 30.0)

    if any(x in n for x in ["deluxe", "premium", "pro"]):
        return ("Premium", 15.0)

    # Fallback: infer from price (rough heuristic)
    return ("Standard", 5.0)

# ═════════════════════════════════════════════════════════════════════════════
# WORKBOOK GENERATION
# ═════════════════════════════════════════════════════════════════════════════

def create_workbook(brand_name, products, fees_map):
    """Create the complete Excel workbook."""

    wb = Workbook()

    # Process products
    product_data = []
    for p in products:
        sku = p.get("sku", "").strip()
        name = p.get("product-name", "")
        asin = p.get("asin", "")

        # Parse price
        try:
            price = float((p.get("your-price", "0") or "0").replace(",", ""))
        except (ValueError, TypeError):
            price = 0.0

        # Parse quantities
        def safe_int(key):
            try:
                return int(p.get(key, "0") or 0)
            except (ValueError, TypeError):
                return 0

        ful = safe_int("afn-fulfillable-quantity")
        wh = safe_int("afn-warehouse-quantity")
        unsel = safe_int("afn-unsellable-quantity")
        total = safe_int("afn-total-quantity")
        inb = safe_int("afn-inbound-working-quantity") + safe_int("afn-inbound-shipped-quantity")

        # Get FBA fee from fees_map, fallback to estimate
        fba_fee = fees_map.get(sku, 0.0)
        if fba_fee <= 0:
            # Estimate based on price
            fba_fee = min(price * 0.15, 8.0) if price > 0 else 3.5

        # Auto-detect category & COGS
        category, cogs = detect_category(name, sku, brand_name)

        product_data.append({
            "sku": sku,
            "asin": asin,
            "name": name,
            "price": price,
            "category": category,
            "cogs": cogs,
            "fba_fee": fba_fee,
            "ful": ful,
            "wh": wh,
            "unsel": unsel,
            "total": total,
            "inb": inb,
        })

    N = len(product_data)
    print(f"✓ Loaded {N} products for {brand_name}")

    # Sheet names
    SH_RAW = "RAW DATA INPUT"
    SH_CAT = "CATEGORY & COSTS"
    SH_PROD = "PRODUCTS DASHBOARD"
    SH_CALC = "DEAL CALCULATOR"
    SH_PLAN = "DEALS PLANNER"
    SH_CAL = "PEAK CALENDAR"

    # ─────────────────────────────────────────────────────────────────────────
    # SHEET 1: RAW DATA INPUT
    # ─────────────────────────────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = SH_RAW
    ws1.sheet_view.showGridLines = False

    # Title
    ws1.merge_cells("A1:V1")
    title_cell = ws1.cell(1, 1, f"📥  RAW DATA INPUT — {brand_name} Inventory")
    title_cell.font = Font(name=FONT_NAME, bold=True, color=WHT, size=13)
    title_cell.fill = fill(NAVY)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 28

    # Instructions
    ws1.merge_cells("A2:V2")
    inst = ws1.cell(2, 1,
        "HOW TO UPDATE: 1) Download Amazon Active Listings Report. 2) Copy all rows (including headers). "
        "3) Paste starting at cell A3. All other sheets auto-update.")
    inst.font = Font(name=FONT_NAME, italic=True, color="555555", size=9)
    inst.fill = fill(GRY)
    inst.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws1.row_dimensions[2].height = 20

    # Headers (row 3) — Amazon Manage Inventory standard columns
    headers = ["sku","fnsku","asin","product-name","condition","your-price",
     "mfn-listing-exists","mfn-fulfillable-quantity","afn-listing-exists",
     "afn-warehouse-quantity","afn-fulfillable-quantity","afn-unsellable-quantity",
     "afn-reserved-quantity","afn-total-quantity","per-unit-volume",
     "afn-inbound-working-quantity","afn-inbound-shipped-quantity",
     "afn-inbound-receiving-quantity","afn-researching-quantity",
     "afn-reserved-future-supply","afn-future-supply-buyable","store"]

    for i, h in enumerate(headers, 1):
        c = ws1.cell(3, i, h)
        c.font = Font(name=FONT_NAME, bold=True, color=WHT, size=9)
        c.fill = fill(BLUE)
        c.border = bdr()
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws1.row_dimensions[3].height = 30

    # Populate with current data
    for idx, p in enumerate(product_data):
        r = idx + 4
        vals = [p["sku"], p["asin"] if p["sku"] != p["asin"] else "", p["asin"],
                p["name"], "New", p["price"], "No", "", "Yes", p["wh"], p["ful"],
                p["unsel"], "", p["total"], "", "", "", "", "", "", "", ""]
        bg = WHT if idx % 2 == 0 else ALT
        for ci, v in enumerate(vals, 1):
            c = ws1.cell(r, ci, v)
            c.font = Font(name=FONT_NAME, size=9)
            c.fill = fill(bg)
            c.border = bdr()
            c.alignment = Alignment(horizontal="right" if isinstance(v, (int, float)) else "left",
                                     vertical="center")
            if ci == 6:
                c.number_format = "$#,##0.00"

    ws1.freeze_panes = "A4"
    col_widths = [12, 14, 14, 55, 8, 12, 14, 18, 14, 16, 18, 17, 17, 14, 12, 20, 20, 21, 19, 22, 20, 10]
    for i, w in enumerate(col_widths, 1):
        ws1.column_dimensions[col_letter(i)].width = w

    # ─────────────────────────────────────────────────────────────────────────
    # SHEET 2: CATEGORY & COSTS
    # ─────────────────────────────────────────────────────────────────────────
    ws2 = wb.create_sheet(SH_CAT)
    ws2.sheet_view.showGridLines = False

    # Title
    ws2.merge_cells("A1:I1")
    title = ws2.cell(1, 1, f"🏷️  CATEGORY & COSTS — {brand_name}")
    title.font = Font(name=FONT_NAME, bold=True, color=WHT, size=13)
    title.fill = fill(NAVY)
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 28

    # Subtitle
    ws2.merge_cells("A2:I2")
    sub = ws2.cell(2, 1, "SKUs auto-fill. Fill YELLOW cells (Brand, Category, COGS, FBA Fee). Auto-detected values provided as defaults.")
    sub.font = Font(name=FONT_NAME, italic=True, color="555555", size=9)
    sub.fill = fill(GRY)
    sub.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws2.row_dimensions[2].height = 18

    # Headers (row 3)
    hdrs = ["SKU", "Product Name (Auto)", "Brand", "Category", "ASIN (Auto)",
            "COGS / Unit ($)", "FBA Fee / Unit ($)", "Min Deal Price 15% ($)", "Notes"]
    for i, h in enumerate(hdrs, 1):
        c = ws2.cell(3, i, h)
        c.font = Font(name=FONT_NAME, bold=True, color=WHT, size=10)
        c.fill = fill(BLUE)
        c.border = bdr()
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws2.row_dimensions[3].height = 30

    # Data rows
    for idx, pd in enumerate(product_data):
        r = idx + 4
        bg = WHT if idx % 2 == 0 else ALT

        # SKU
        c = ws2.cell(r, 1, pd["sku"])
        c.font = Font(name=FONT_NAME, bold=True, size=10)
        c.fill = fill(bg)
        c.border = bdr()

        # Product name (formula from Raw Data)
        c = ws2.cell(r, 2, f"=IFERROR(INDEX('{SH_RAW}'!$A:$Z,MATCH(A{r},'{SH_RAW}'!$A:$A,0),MATCH(\"product-name\",'{SH_RAW}'!$3:$3,0)),\"\")")
        c.font = Font(name=FONT_NAME, size=10)
        c.fill = fill(LT_BLUE)
        c.border = bdr()

        # Brand (editable, pre-filled)
        c = ws2.cell(r, 3, brand_name)
        c.font = Font(name=FONT_NAME, size=10)
        c.fill = fill(YEL)
        c.border = bdr()
        c.alignment = Alignment(horizontal="center")

        # Category (editable, pre-filled with detected)
        c = ws2.cell(r, 4, pd["category"])
        c.font = Font(name=FONT_NAME, size=10)
        c.fill = fill(YEL)
        c.border = bdr()

        # ASIN (formula from Raw Data)
        c = ws2.cell(r, 5, f"=IFERROR(INDEX('{SH_RAW}'!$A:$Z,MATCH(A{r},'{SH_RAW}'!$A:$A,0),MATCH(\"asin\",'{SH_RAW}'!$3:$3,0)),\"\")")
        c.font = Font(name=FONT_NAME, size=10)
        c.fill = fill(LT_BLUE)
        c.border = bdr()

        # COGS (editable, pre-filled)
        c = ws2.cell(r, 6, pd["cogs"])
        c.font = Font(name=FONT_NAME, bold=True, size=10)
        c.number_format = "$#,##0.00"
        c.fill = fill(YEL)
        c.border = bdr()
        c.alignment = Alignment(horizontal="right")

        # FBA Fee (editable, pre-filled from fees_map)
        c = ws2.cell(r, 7, pd["fba_fee"])
        c.font = Font(name=FONT_NAME, bold=True, size=10)
        c.number_format = "$#,##0.00"
        c.fill = fill(YEL)
        c.border = bdr()
        c.alignment = Alignment(horizontal="right")

        # Min deal price (calculated, 85% of list)
        c = ws2.cell(r, 8, f"=IFERROR(VALUE(INDEX('{SH_RAW}'!$A:$Z,MATCH(A{r},'{SH_RAW}'!$A:$A,0),MATCH(\"your-price\",'{SH_RAW}'!$3:$3,0)))*0.85,0)")
        c.font = Font(name=FONT_NAME, bold=True, size=10)
        c.number_format = "$#,##0.00"
        c.fill = fill(GRN)
        c.border = bdr()
        c.alignment = Alignment(horizontal="right")

        # Notes
        c = ws2.cell(r, 9, "")
        c.fill = fill(bg)
        c.border = bdr()

    set_col_widths(ws2, [16, 55, 14, 22, 14, 14, 14, 18, 35])

    # ─────────────────────────────────────────────────────────────────────────
    # SHEET 3: PRODUCTS DASHBOARD
    # ─────────────────────────────────────────────────────────────────────────
    ws3 = wb.create_sheet(SH_PROD)
    ws3.sheet_view.showGridLines = False
    ws3.freeze_panes = "A4"

    # Title
    ws3.merge_cells("A1:O1")
    title = ws3.cell(1, 1, f"📦  PRODUCTS DASHBOARD — {brand_name}")
    title.font = Font(name=FONT_NAME, bold=True, color=WHT, size=13)
    title.fill = fill(NAVY)
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 28

    # Subtitle
    ws3.merge_cells("A2:O2")
    sub = ws3.cell(2, 1, "Auto-populated from RAW DATA and CATEGORY & COSTS. Shows eligibility (7+ units + price), inventory, and minimum deal price.")
    sub.font = Font(name=FONT_NAME, italic=True, color="555555", size=9)
    sub.fill = fill(GRY)
    ws3.row_dimensions[2].height = 18

    # Headers
    prod_hdrs = [("SKU", NAVY), ("Brand", NAVY), ("Category", BLUE),
                 ("Product", BLUE), ("List Price ($)", BLUE), ("FBA Fulfillable", BLUE),
                 ("Warehouse", BLUE), ("Unsellable", BLUE), ("Total", BLUE), ("Inbound", BLUE),
                 ("COGS ($)", NAVY), ("FBA Fee ($)", NAVY),
                 ("Min Deal 15% ($)", NAVY), ("Eligible?", NAVY), ("Notes", BLUE)]
    for i, (h, bg) in enumerate(prod_hdrs, 1):
        c = ws3.cell(3, i, h)
        c.font = Font(name=FONT_NAME, bold=True, color=WHT, size=10)
        c.fill = fill(bg)
        c.border = bdr()
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws3.row_dimensions[3].height = 36

    # Data rows
    for idx, pd in enumerate(product_data):
        r = idx + 4
        bg = WHT if idx % 2 == 0 else ALT

        # SKU
        ws3.cell(r, 1, pd["sku"]).font = Font(name=FONT_NAME, bold=True, size=10)
        ws3.cell(r, 1).fill = fill(bg); ws3.cell(r, 1).border = bdr()

        # Brand (from Cat & Costs)
        ws3.cell(r, 2, f"=IFERROR(VLOOKUP(A{r},'{SH_CAT}'!$A:$C,3,0),\"\")")
        ws3.cell(r, 2).fill = fill(LT_BLUE); ws3.cell(r, 2).border = bdr()

        # Category (from Cat & Costs)
        ws3.cell(r, 3, f"=IFERROR(VLOOKUP(A{r},'{SH_CAT}'!$A:$D,4,0),\"\")")
        ws3.cell(r, 3).fill = fill(LT_BLUE); ws3.cell(r, 3).border = bdr()

        # Product name
        ws3.cell(r, 4, f"=IFERROR(INDEX('{SH_RAW}'!$A:$Z,MATCH(A{r},'{SH_RAW}'!$A:$A,0),MATCH(\"product-name\",'{SH_RAW}'!$3:$3,0)),\"\")")
        ws3.cell(r, 4).fill = fill(bg); ws3.cell(r, 4).border = bdr()
        ws3.row_dimensions[r].height = 28

        # List price
        ws3.cell(r, 5, f"=IFERROR(VALUE(INDEX('{SH_RAW}'!$A:$Z,MATCH(A{r},'{SH_RAW}'!$A:$A,0),MATCH(\"your-price\",'{SH_RAW}'!$3:$3,0))),0)")
        ws3.cell(r, 5).number_format = "$#,##0.00"
        ws3.cell(r, 5).fill = fill(bg); ws3.cell(r, 5).border = bdr()
        ws3.cell(r, 5).alignment = Alignment(horizontal="right")

        # FBA Fulfillable
        ws3.cell(r, 6, f"=IFERROR(VALUE(INDEX('{SH_RAW}'!$A:$Z,MATCH(A{r},'{SH_RAW}'!$A:$A,0),MATCH(\"afn-fulfillable-quantity\",'{SH_RAW}'!$3:$3,0))),0)")
        ws3.cell(r, 6).number_format = "#,##0"
        ws3.cell(r, 6).fill = fill(bg); ws3.cell(r, 6).border = bdr()
        ws3.cell(r, 6).alignment = Alignment(horizontal="right")

        # Warehouse
        ws3.cell(r, 7, f"=IFERROR(VALUE(INDEX('{SH_RAW}'!$A:$Z,MATCH(A{r},'{SH_RAW}'!$A:$A,0),MATCH(\"afn-warehouse-quantity\",'{SH_RAW}'!$3:$3,0))),0)")
        ws3.cell(r, 7).number_format = "#,##0"
        ws3.cell(r, 7).fill = fill(bg); ws3.cell(r, 7).border = bdr()
        ws3.cell(r, 7).alignment = Alignment(horizontal="right")

        # Unsellable
        ws3.cell(r, 8, f"=IFERROR(VALUE(INDEX('{SH_RAW}'!$A:$Z,MATCH(A{r},'{SH_RAW}'!$A:$A,0),MATCH(\"afn-unsellable-quantity\",'{SH_RAW}'!$3:$3,0))),0)")
        ws3.cell(r, 8).number_format = "#,##0"
        ws3.cell(r, 8).fill = fill(bg); ws3.cell(r, 8).border = bdr()
        ws3.cell(r, 8).alignment = Alignment(horizontal="right")

        # Total
        ws3.cell(r, 9, f"=IFERROR(VALUE(INDEX('{SH_RAW}'!$A:$Z,MATCH(A{r},'{SH_RAW}'!$A:$A,0),MATCH(\"afn-total-quantity\",'{SH_RAW}'!$3:$3,0))),0)")
        ws3.cell(r, 9).number_format = "#,##0"
        ws3.cell(r, 9).fill = fill(bg); ws3.cell(r, 9).border = bdr()
        ws3.cell(r, 9).alignment = Alignment(horizontal="right")

        # Inbound
        ws3.cell(r, 10, f"=IFERROR(VALUE(INDEX('{SH_RAW}'!$A:$Z,MATCH(A{r},'{SH_RAW}'!$A:$A,0),MATCH(\"afn-inbound-working-quantity\",'{SH_RAW}'!$3:$3,0)))+VALUE(INDEX('{SH_RAW}'!$A:$Z,MATCH(A{r},'{SH_RAW}'!$A:$A,0),MATCH(\"afn-inbound-shipped-quantity\",'{SH_RAW}'!$3:$3,0))),0)")
        ws3.cell(r, 10).number_format = "#,##0"
        ws3.cell(r, 10).fill = fill(bg); ws3.cell(r, 10).border = bdr()
        ws3.cell(r, 10).alignment = Alignment(horizontal="right")

        # COGS (from Cat & Costs)
        ws3.cell(r, 11, f"=IFERROR(VLOOKUP(A{r},'{SH_CAT}'!$A:$F,6,0),0)")
        ws3.cell(r, 11).number_format = "$#,##0.00"
        ws3.cell(r, 11).fill = fill(GRN); ws3.cell(r, 11).border = bdr()
        ws3.cell(r, 11).alignment = Alignment(horizontal="right")

        # FBA Fee (from Cat & Costs)
        ws3.cell(r, 12, f"=IFERROR(VLOOKUP(A{r},'{SH_CAT}'!$A:$G,7,0),0)")
        ws3.cell(r, 12).number_format = "$#,##0.00"
        ws3.cell(r, 12).fill = fill(GRN); ws3.cell(r, 12).border = bdr()
        ws3.cell(r, 12).alignment = Alignment(horizontal="right")

        # Min deal price (85% of list)
        ws3.cell(r, 13, f"=IFERROR(E{r}*0.85,0)")
        ws3.cell(r, 13).number_format = "$#,##0.00"
        ws3.cell(r, 13).fill = fill(GRN); ws3.cell(r, 13).border = bdr()
        ws3.cell(r, 13).alignment = Alignment(horizontal="right")
        ws3.cell(r, 13).font = Font(name=FONT_NAME, bold=True, size=10)

        # Eligibility
        ws3.cell(r, 14, f"=IF(AND(F{r}>=7,E{r}>0),\"✅ Eligible\",IF(E{r}=0,\"⚠️ No Price\",\"❌ Low Stock\"))")
        ws3.cell(r, 14).alignment = Alignment(horizontal="center")
        ws3.cell(r, 14).fill = fill(GRN); ws3.cell(r, 14).border = bdr()
        ws3.cell(r, 14).font = Font(name=FONT_NAME, bold=True, size=10)

        # Notes
        note_formula = (f"=IF(F{r}=0,\"No fulfillable stock\","
                       f"IF(F{r}<7,\"Need \"&(7-F{r})&\" more (min 7)\","
                       f"IF(E{r}=0,\"Set price\","
                       f"\"Ready ✓\")))")
        ws3.cell(r, 15, note_formula)
        ws3.cell(r, 15).fill = fill(bg); ws3.cell(r, 15).border = bdr()
        ws3.cell(r, 15).font = Font(name=FONT_NAME, italic=True, size=9, color="444444")

    set_col_widths(ws3, [16, 12, 22, 54, 13, 13, 13, 11, 11, 11, 13, 13, 16, 14, 30])

    # ─────────────────────────────────────────────────────────────────────────
    # SHEET 4: DEAL CALCULATOR (All Promo Types)
    # ─────────────────────────────────────────────────────────────────────────
    ws4 = wb.create_sheet(SH_CALC)
    ws4.sheet_view.showGridLines = False

    # Title
    ws4.merge_cells("A1:J1")
    title = ws4.cell(1, 1, f"🧮  DEAL CALCULATOR — All 13 Promotion Types  (v2 – fixed)")
    title.font = Font(name=FONT_NAME, bold=True, color=WHT, size=13)
    title.fill = fill(NAVY)
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws4.row_dimensions[1].height = 28

    ws4.merge_cells("A2:J2")
    sub = ws4.cell(2, 1, "Select SKU, enter discount %, units sold, COGS, and FBA fee. View all deal types with correct fee calculations: UPFRONT FEE = lump sum per deal period. VARIABLE FEE = % of sales.")
    sub.font = Font(name=FONT_NAME, italic=True, color="555555", size=9)
    sub.fill = fill(GRY)
    ws4.row_dimensions[2].height = 20

    # Input section
    ws4.merge_cells("A3:J3")
    hdr = ws4.cell(3, 1, "⚙️  PRODUCT SELECTION & SHARED INPUTS")
    hdr.font = Font(name=FONT_NAME, bold=True, color=WHT, size=11)
    hdr.fill = fill(BLUE)
    ws4.row_dimensions[3].height = 22

    # Inputs
    def input_cell(r, label, c, val=None, formula=None, fmt=None, bg=YEL):
        ws4.merge_cells(f"A{r}:B{r}")
        lbl = ws4.cell(r, 1, label)
        lbl.font = Font(name=FONT_NAME, bold=True, size=10)
        lbl.fill = fill(GRY); lbl.border = bdr()
        lbl.alignment = Alignment(horizontal="left", indent=1)

        cell = ws4.cell(r, c, formula if formula else val)
        cell.font = Font(name=FONT_NAME, bold=True if formula else False, size=10)
        cell.fill = fill(bg); cell.border = bdr()
        cell.alignment = Alignment(horizontal="right")
        if fmt: cell.number_format = fmt
        ws4.merge_cells(f"C{r}:J{r}") if c == 3 else None
        return cell

    input_cell(4, "Select SKU", 3, product_data[0]["sku"] if product_data else "", bg=YEL)

    # SKU dropdown — references SKU column from PRODUCTS DASHBOARD
    last_prod_row = 3 + len(product_data)
    dv_sku = DataValidation(
        type="list",
        formula1=f"'PRODUCTS DASHBOARD'!$A$4:$A${last_prod_row}",
        allow_blank=True,
        showDropDown=False  # False = SHOW the dropdown arrow in Excel
    )
    dv_sku.error = "Please select a valid SKU from the list"
    dv_sku.errorTitle = "Invalid SKU"
    dv_sku.prompt = "Pick a SKU to analyze"
    dv_sku.promptTitle = "Select SKU"
    dv_sku.showInputMessage = True
    dv_sku.showErrorMessage = True
    ws4.add_data_validation(dv_sku)
    dv_sku.sqref = "C4"

    input_cell(5, "Product", 3, formula="=IFERROR(VLOOKUP(C4,'PRODUCTS DASHBOARD'!$A:$D,4,0),\"\")", bg=LT_BLUE)
    input_cell(6, "List Price", 3, formula="=IFERROR(VLOOKUP(C4,'PRODUCTS DASHBOARD'!$A:$E,5,0),0)", fmt="$#,##0.00", bg=LT_BLUE)
    input_cell(7, "Discount %", 3, 0.20, fmt="0.0%", bg=YEL)
    input_cell(8, "Deal Price ($)", 3, formula="=C6*(1-C7)", fmt="$#,##0.00", bg=GRN)
    input_cell(9, "Estimated Units", 3, 30, fmt="#,##0", bg=YEL)
    input_cell(10, "COGS per Unit", 3, formula="=IFERROR(VLOOKUP(C4,'CATEGORY & COSTS'!$A:$F,6,0),0)", fmt="$#,##0.00", bg=LT_BLUE)
    input_cell(11, "FBA Fee per Unit", 3, formula="=IFERROR(VLOOKUP(C4,'CATEGORY & COSTS'!$A:$G,7,0),0)", fmt="$#,##0.00", bg=LT_BLUE)
    input_cell(12, "Deal Duration (days)", 3, 7, fmt="#,##0", bg=YEL)

    # Promo types table header
    ws4.merge_cells("A14:J14")
    hdr = ws4.cell(14, 1, "📊  ALL 13 PROMOTION TYPES — SIDE-BY-SIDE COMPARISON")
    hdr.font = Font(name=FONT_NAME, bold=True, color=WHT, size=11)
    hdr.fill = fill(NAVY)
    ws4.row_dimensions[14].height = 22

    # Table headers
    tbl_hdrs = [("Type", NAVY), ("Period", BLUE), ("Duration", BLUE), ("Upfront Fee ($)", NAVY),
                ("Var Fee %", BLUE), ("Total Fee ($)", NAVY), ("Revenue ($)", BLUE),
                ("COGS+FBA ($)", BLUE), ("Net Profit ($)", NAVY), ("Margin", NAVY)]
    for i, (h, bg) in enumerate(tbl_hdrs, 1):
        c = ws4.cell(15, i, h)
        c.font = Font(name=FONT_NAME, bold=True, color=WHT if bg==NAVY or bg==BLUE else "000000", size=9)
        c.fill = fill(bg); c.border = bdr()
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    ws4.row_dimensions[15].height = 28

    # Promo data: (name, period, duration_text, upfront_fee, var_fee_expr)
    # var_fee_expr: formula expression WITHOUT leading "=" (use # for current row)
    # Variable fee is % of REVENUE (col G), NOT costs
    promos = [
        ("Lightning Deal", "Non-Peak", "=TEXT(C12,\"0.0\")&\" days\"", "=C12*70",  "MIN(G#*0.01,2000)"),
        ("Best Deal 3-day", "Non-Peak", "3 days",   210,   "MIN(G#*0.01,2000)"),
        ("Best Deal 7-day", "Non-Peak", "7 days",   490,   "MIN(G#*0.01,2000)"),
        ("Best Deal 14-day","Non-Peak", "14 days",  980,   "MIN(G#*0.01,2000)"),
        ("Lightning Deal",  "Prime Day","4-12 hrs",  500,   "0"),
        ("Best Deal",       "Prime Day","1-14 days", 1000,  "0"),
        ("Lightning Deal",  "BFCM",     "4-12 hrs",  500,   "0"),
        ("Best Deal",       "BFCM",     "1-14 days", 1000,  "0"),
        ("Lightning Deal",  "Prime Big Deal Days","4-12 hrs", 500, "0"),
        ("Best Deal",       "Prime Big Deal Days","1-14 days",1000,"0"),
        ("Prime Exclusive", "BFCM",     "Event",     245,   "0"),
        ("Coupon",          "Any",      "Any",        5,    "G#*0.025"),
        ("Regular Discount","Any",      "Any",        0,    "0"),
    ]

    for idx, (ptype, period, dur, fee_up, fee_var_expr) in enumerate(promos):
        r = 16 + idx
        bg = ALT if idx % 2 == 0 else WHT

        ws4.cell(r, 1, ptype).fill = fill(bg); ws4.cell(r, 1).border = bdr()
        ws4.cell(r, 1).font = Font(name=FONT_NAME, bold=True, size=10)

        ws4.cell(r, 2, period).fill = fill(bg); ws4.cell(r, 2).border = bdr()
        ws4.cell(r, 2).font = Font(name=FONT_NAME, size=10)

        # Duration: formula or plain text
        ws4.cell(r, 3, dur)
        ws4.cell(r, 3).fill = fill(bg); ws4.cell(r, 3).border = bdr()
        ws4.cell(r, 3).font = Font(name=FONT_NAME, size=9)

        # Upfront fee: write as number (not string) or formula
        if isinstance(fee_up, str) and fee_up.startswith("="):
            ws4.cell(r, 4, fee_up)
        else:
            ws4.cell(r, 4, int(fee_up) if isinstance(fee_up, (int, float)) else fee_up)
        c = ws4.cell(r, 4)
        c.number_format = "$#,##0.00"; c.fill = fill(bg); c.border = bdr()
        c.alignment = Alignment(horizontal="right")

        # Variable fee % (display label)
        var_pct = "1% cap $2k" if "MIN" in fee_var_expr else ("2.5%" if "0.025" in fee_var_expr else "None")
        ws4.cell(r, 5, var_pct).fill = fill(bg); ws4.cell(r, 5).border = bdr()
        ws4.cell(r, 5).font = Font(name=FONT_NAME, size=9)
        ws4.cell(r, 5).alignment = Alignment(horizontal="center")

        # Total fee = Upfront + Variable (clean formula, no double "=")
        var_resolved = fee_var_expr.replace("#", str(r))
        c = ws4.cell(r, 6, f"=D{r}+{var_resolved}")
        c.number_format = "$#,##0.00"; c.fill = fill(LIGHT_ORANGE); c.border = bdr()
        c.alignment = Alignment(horizontal="right"); c.font = Font(bold=True, size=10)

        # Revenue
        c = ws4.cell(r, 7, f"=C8*C9")
        c.number_format = "$#,##0.00"; c.fill = fill(GRN); c.border = bdr()
        c.alignment = Alignment(horizontal="right")

        # COGS + FBA
        c = ws4.cell(r, 8, f"=(C10+C11)*C9")
        c.number_format = "$#,##0.00"; c.fill = fill(GRN); c.border = bdr()
        c.alignment = Alignment(horizontal="right")

        # Net profit
        c = ws4.cell(r, 9, f"=G{r}-F{r}-H{r}")
        c.number_format = "$#,##0.00"; c.fill = fill(GRN); c.border = bdr()
        c.alignment = Alignment(horizontal="right"); c.font = Font(bold=True, size=10)

        # Margin
        c = ws4.cell(r, 10, f"=IFERROR(I{r}/G{r},0)")
        c.number_format = "0.0%"; c.fill = fill(GRN); c.border = bdr()
        c.alignment = Alignment(horizontal="right"); c.font = Font(bold=True, size=10)

    set_col_widths(ws4, [20, 16, 14, 14, 12, 14, 14, 14, 14, 12])
    ws4.freeze_panes = "A15"

    # ─────────────────────────────────────────────────────────────────────────
    # SHEET 5: DEALS PLANNER (All Products × Period/Type Selector)
    # ─────────────────────────────────────────────────────────────────────────
    ws5 = wb.create_sheet(SH_PLAN)
    ws5.sheet_view.showGridLines = False

    # Title
    ws5.merge_cells("A1:P1")
    title = ws5.cell(1, 1, f"📅  DEALS PLANNER — All Products with Period & Deal Type Selector")
    title.font = Font(name=FONT_NAME, bold=True, color=WHT, size=13)
    title.fill = fill(NAVY)
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws5.row_dimensions[1].height = 28

    ws5.merge_cells("A2:P2")
    sub = ws5.cell(2, 1, "Set PERIOD and DEAL TYPE in the selectors — table recalculates for ALL products. Yellow = editable inputs. UPFRONT FEE charged once per deal. VARIABLE FEE applies per unit as % of revenue.")
    sub.font = Font(name=FONT_NAME, italic=True, color="555555", size=9)
    sub.fill = fill(GRY)
    ws5.row_dimensions[2].height = 20

    # Selector section header
    ws5.merge_cells("A3:P3")
    hdr = ws5.cell(3, 1, "⚙️  GLOBAL SETTINGS — Change to Recalculate All Products")
    hdr.font = Font(name=FONT_NAME, bold=True, color=WHT, size=11)
    hdr.fill = fill(NAVY); ws5.row_dimensions[3].height = 22

    # Selectors (row 4)
    ws5.cell(4, 1, "Period:").font = Font(bold=True); ws5.cell(4, 1).fill = fill(GRY)
    ws5.cell(4, 2, "Non-Peak").fill = fill(YEL); ws5.cell(4, 2).font = Font(bold=True)
    ws5.cell(4, 3, "Deal Type:").font = Font(bold=True); ws5.cell(4, 3).fill = fill(GRY)
    ws5.cell(4, 4, "Best Deal").fill = fill(YEL); ws5.cell(4, 4).font = Font(bold=True)
    ws5.cell(4, 5, "Discount %:").font = Font(bold=True); ws5.cell(4, 5).fill = fill(GRY)
    ws5.cell(4, 6, 0.20).number_format = "0.0%"; ws5.cell(4, 6).fill = fill(YEL); ws5.cell(4, 6).font = Font(bold=True)
    ws5.cell(4, 7, "Duration (days):").font = Font(bold=True); ws5.cell(4, 7).fill = fill(GRY)
    ws5.cell(4, 8, 7).number_format = "#,##0"; ws5.cell(4, 8).fill = fill(YEL); ws5.cell(4, 8).font = Font(bold=True)

    # Dropdowns
    dv_period = DataValidation(type="list", formula1='"Non-Peak,Prime Day,Black Friday/Cyber Monday,Prime Big Deal Days"', allow_blank=False)
    ws5.add_data_validation(dv_period); dv_period.sqref = "B4"

    dv_dtype = DataValidation(type="list", formula1='"Lightning Deal,Best Deal,Coupon,Regular Discount,Prime Exclusive"', allow_blank=False)
    ws5.add_data_validation(dv_dtype); dv_dtype.sqref = "D4"

    ws5.freeze_panes = "A7"

    # Product table headers (row 6)
    tbl_hdrs = [("Category", NAVY), ("SKU", NAVY), ("Product", BLUE), ("Price ($)", BLUE),
                ("FBA Ful", BLUE), ("Eligible?", NAVY), ("Disc %", NAVY), ("Deal Price ($)", BLUE),
                ("Est Units", NAVY), ("Revenue ($)", BLUE), ("Upfront Fee ($)", NAVY),
                ("Variable Fee ($)", NAVY), ("Total Fee ($)", NAVY), ("COGS+FBA ($)", BLUE),
                ("Net Profit ($)", NAVY), ("Margin", NAVY)]
    for i, (h, bg) in enumerate(tbl_hdrs, 1):
        c = ws5.cell(6, i, h)
        c.font = Font(name=FONT_NAME, bold=True, color=WHT if bg in [NAVY, BLUE] else "000000", size=9)
        c.fill = fill(bg); c.border = bdr()
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    ws5.row_dimensions[6].height = 32

    # Data rows
    for idx, pd in enumerate(product_data):
        r = idx + 7
        bg = WHT if idx % 2 == 0 else ALT

        # Category
        ws5.cell(r, 1, f"=IFERROR(VLOOKUP(B{r},'{SH_CAT}'!$A:$D,4,0),\"\")")
        ws5.cell(r, 1).fill = fill(bg); ws5.cell(r, 1).border = bdr(); ws5.cell(r, 1).font = Font(size=9)

        # SKU
        ws5.cell(r, 2, pd["sku"])
        ws5.cell(r, 2).fill = fill(bg); ws5.cell(r, 2).border = bdr()
        ws5.cell(r, 2).font = Font(bold=True, size=10)

        # Product name
        ws5.cell(r, 3, f"=IFERROR(VLOOKUP(B{r},'{SH_PROD}'!$A:$D,4,0),\"\")")
        ws5.cell(r, 3).fill = fill(bg); ws5.cell(r, 3).border = bdr()
        ws5.cell(r, 3).font = Font(size=9); ws5.row_dimensions[r].height = 26
        ws5.cell(r, 3).alignment = Alignment(wrap_text=True)

        # Price
        ws5.cell(r, 4, f"=IFERROR(VLOOKUP(B{r},'{SH_PROD}'!$A:$E,5,0),0)")
        ws5.cell(r, 4).number_format = "$#,##0.00"; ws5.cell(r, 4).fill = fill(bg); ws5.cell(r, 4).border = bdr()
        ws5.cell(r, 4).alignment = Alignment(horizontal="right")

        # FBA Fulfillable
        ws5.cell(r, 5, f"=IFERROR(VLOOKUP(B{r},'{SH_PROD}'!$A:$F,6,0),0)")
        ws5.cell(r, 5).number_format = "#,##0"; ws5.cell(r, 5).fill = fill(bg); ws5.cell(r, 5).border = bdr()
        ws5.cell(r, 5).alignment = Alignment(horizontal="right")

        # Eligible
        ws5.cell(r, 6, f"=IFERROR(VLOOKUP(B{r},'{SH_PROD}'!$A:$N,14,0),\"\")")
        ws5.cell(r, 6).fill = fill(bg); ws5.cell(r, 6).border = bdr()
        ws5.cell(r, 6).font = Font(bold=True, size=9); ws5.cell(r, 6).alignment = Alignment(horizontal="center")

        # Discount % (per-product override, default to global)
        ws5.cell(r, 7, f"=$B$4")
        ws5.cell(r, 7).number_format = "0.0%"; ws5.cell(r, 7).fill = fill(YEL); ws5.cell(r, 7).border = bdr()
        ws5.cell(r, 7).alignment = Alignment(horizontal="right"); ws5.cell(r, 7).font = Font(bold=True, size=10)

        # Deal price
        ws5.cell(r, 8, f"=D{r}*(1-G{r})")
        ws5.cell(r, 8).number_format = "$#,##0.00"; ws5.cell(r, 8).fill = fill(GRN); ws5.cell(r, 8).border = bdr()
        ws5.cell(r, 8).alignment = Alignment(horizontal="right"); ws5.cell(r, 8).font = Font(bold=True, size=10)

        # Est units
        ws5.cell(r, 9, 20)
        ws5.cell(r, 9).number_format = "#,##0"; ws5.cell(r, 9).fill = fill(YEL); ws5.cell(r, 9).border = bdr()
        ws5.cell(r, 9).alignment = Alignment(horizontal="right"); ws5.cell(r, 9).font = Font(bold=True, size=10)

        # Revenue
        ws5.cell(r, 10, f"=H{r}*I{r}")
        ws5.cell(r, 10).number_format = "$#,##0.00"; ws5.cell(r, 10).fill = fill(GRN); ws5.cell(r, 10).border = bdr()
        ws5.cell(r, 10).alignment = Alignment(horizontal="right")

        # Upfront fee
        upfront_f = (f'=IF($B$4="Non-Peak",$H$4*$I$4,IF($D$4="Lightning Deal",500,IF($D$4="Best Deal",1000,'
                    f'IF($D$4="Coupon",5,IF($D$4="Prime Exclusive",245,0)))))')
        ws5.cell(r, 11, upfront_f)
        ws5.cell(r, 11).number_format = "$#,##0.00"; ws5.cell(r, 11).fill = fill(LIGHT_ORANGE); ws5.cell(r, 11).border = bdr()
        ws5.cell(r, 11).alignment = Alignment(horizontal="right")

        # Variable fee
        var_fee_f = (f'=IF($B$4="Non-Peak",MIN(J{r}*0.01,2000),IF($D$4="Coupon",J{r}*0.025,0))')
        ws5.cell(r, 12, var_fee_f)
        ws5.cell(r, 12).number_format = "$#,##0.00"; ws5.cell(r, 12).fill = fill(LIGHT_ORANGE); ws5.cell(r, 12).border = bdr()
        ws5.cell(r, 12).alignment = Alignment(horizontal="right")

        # Total fee
        ws5.cell(r, 13, f"=K{r}+L{r}")
        ws5.cell(r, 13).number_format = "$#,##0.00"; ws5.cell(r, 13).fill = fill(LIGHT_ORANGE); ws5.cell(r, 13).border = bdr()
        ws5.cell(r, 13).alignment = Alignment(horizontal="right"); ws5.cell(r, 13).font = Font(bold=True, size=10)

        # COGS + FBA
        ws5.cell(r, 14, f"=(IFERROR(VLOOKUP(B{r},'{SH_CAT}'!$A:$F,6,0),0)+IFERROR(VLOOKUP(B{r},'{SH_CAT}'!$A:$G,7,0),0))*I{r}")
        ws5.cell(r, 14).number_format = "$#,##0.00"; ws5.cell(r, 14).fill = fill(GRN); ws5.cell(r, 14).border = bdr()
        ws5.cell(r, 14).alignment = Alignment(horizontal="right")

        # Net profit
        ws5.cell(r, 15, f"=J{r}-M{r}-N{r}")
        ws5.cell(r, 15).number_format = "$#,##0.00"; ws5.cell(r, 15).fill = fill(GRN); ws5.cell(r, 15).border = bdr()
        ws5.cell(r, 15).alignment = Alignment(horizontal="right"); ws5.cell(r, 15).font = Font(bold=True, size=10)

        # Margin
        ws5.cell(r, 16, f"=IFERROR(O{r}/J{r},0)")
        ws5.cell(r, 16).number_format = "0.0%"; ws5.cell(r, 16).fill = fill(GRN); ws5.cell(r, 16).border = bdr()
        ws5.cell(r, 16).alignment = Alignment(horizontal="right"); ws5.cell(r, 16).font = Font(bold=True, size=10)

    set_col_widths(ws5, [22, 14, 52, 13, 12, 14, 12, 14, 14, 14, 14, 14, 14, 14, 14, 12])

    # ─────────────────────────────────────────────────────────────────────────
    # SHEET 6: PEAK CALENDAR
    # ─────────────────────────────────────────────────────────────────────────
    ws6 = wb.create_sheet(SH_CAL)
    ws6.sheet_view.showGridLines = False

    # Title
    ws6.merge_cells("A1:I1")
    title = ws6.cell(1, 1, f"📆  PEAK EVENT CALENDAR — {brand_name} 2025/2026")
    title.font = Font(name=FONT_NAME, bold=True, color=WHT, size=13)
    title.fill = fill(NAVY)
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws6.row_dimensions[1].height = 28

    ws6.merge_cells("A2:I2")
    sub = ws6.cell(2, 1, "Planning guide for peak events with fee structures and submission deadlines. Customize strategy notes for your specific brand below.")
    sub.font = Font(name=FONT_NAME, italic=True, color="555555", size=9)
    sub.fill = fill(GRY)
    ws6.row_dimensions[2].height = 18

    # Headers
    cal_hdrs = [("Event", NAVY), ("Dates", BLUE), ("Best Deal Fee", NAVY), ("Lightning Fee", NAVY),
               ("Coupon Fee", BLUE), ("Min Discount", BLUE), ("Strategy Notes", BLUE), ("Brand Opportunity", NAVY)]
    for i, (h, bg) in enumerate(cal_hdrs, 1):
        c = ws6.cell(3, i, h)
        c.font = Font(name=FONT_NAME, bold=True, color=WHT if bg in [NAVY] else "000000", size=10)
        c.fill = fill(bg); c.border = bdr()
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    ws6.row_dimensions[3].height = 28

    # Generic peak events (user can customize)
    events_generic = [
        ("Prime Day", "~July (estimated)", "$1,000", "$500", "$5+2.5%", "15%",
         "Submit 4+ weeks early. Strong impulse buys. Best Deals for sustained visibility.",
         "🔥 HIGH — Peak summer shopping event."),
        ("Black Friday/BFCM", "~Late November", "$1,000", "$500", "$245", "15%",
         "Submit 6+ weeks early. Biggest traffic day. Combine Best Deal + Coupon + PEPDP.",
         "🔥 HIGHEST — Top gift-giving event."),
        ("Prime Big Deal Days", "~October (estimated)", "$1,000", "$500", "$5+2.5%", "15%",
         "Submit 4+ weeks ahead. Pre-holiday mindset. Good for multi-packs as gifts.",
         "🟡 MEDIUM — Pre-holiday shopping push."),
        ("Cyber Monday", "~December 2", "$1,000", "$500", "$245", "15%",
         "Span Best Deal across BFCM + Cyber Monday for single $1,000 fee. Layer coupons.",
         "🔥 HIGH — Online-focused shoppers."),
        ("Non-Peak Best Deals", "Year-round", "$70/day", "$70/day", "$5+2.5%", "15%",
         "Best ROI for margins. Run 7–14 day deals at $70/day. Volume drives profitability.",
         "🟡 MEDIUM — Consistent, predictable fee structure."),
    ]

    for idx, (evt, dates, bd_fee, ld_fee, cpn_fee, disc, notes, opp) in enumerate(events_generic):
        r = idx + 4
        bg = WHT if idx % 2 == 0 else ALT
        for ci, val in enumerate([evt, dates, bd_fee, ld_fee, cpn_fee, disc, notes, opp], 1):
            c = ws6.cell(r, ci, val)
            c.font = Font(name=FONT_NAME, size=9, bold=(ci in [1, 8]))
            c.fill = fill(bg); c.border = bdr()
            c.alignment = Alignment(horizontal="left" if ci in [1, 6, 7] else "center", wrap_text=True)
        ws6.row_dimensions[r].height = 48

    set_col_widths(ws6, [24, 22, 12, 12, 14, 12, 46, 38])

    return wb

def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[col_letter(i)].width = w

# ═════════════════════════════════════════════════════════════════════════════
# MAIN EXECUTION
# ═════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="Generate Amazon Deals Planner for any brand",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python amazon_deals_planner_generator.py --brand "Grillbot" \\
    --inventory inventory.csv --fees fees.csv --output grillbot_planner.xlsx

  python amazon_deals_planner_generator.py --brand "Wicked" \\
    --inventory wicked_inventory.csv --fees wicked_fees.csv \\
    --output wicked_planner.xlsx
        """)

    parser.add_argument("--brand", required=True, help="Brand name (e.g., 'Grillbot', 'Wicked')")
    parser.add_argument("--inventory", required=True, help="Path to Amazon Manage Inventory CSV")
    parser.add_argument("--fees", required=True, help="Path to Amazon Fee Preview CSV")
    parser.add_argument("--output", required=True, help="Output Excel file path")

    args = parser.parse_args()

    # Validate inputs
    inv_path = Path(args.inventory)
    fees_path = Path(args.fees)
    out_path = Path(args.output)

    if not inv_path.exists():
        print(f"❌ Inventory file not found: {inv_path}")
        sys.exit(1)
    if not fees_path.exists():
        print(f"❌ Fees file not found: {fees_path}")
        sys.exit(1)

    print(f"\n📊  Amazon Deals Planner Generator")
    print(f"{'='*60}")
    print(f"Brand:        {args.brand}")
    print(f"Inventory:    {inv_path}")
    print(f"Fees:         {fees_path}")
    print(f"Output:       {out_path}")
    print(f"{'='*60}\n")

    # Load data
    print("Loading inventory...")
    products = load_inventory(inv_path)
    print(f"✓ Loaded {len(products)} products")

    print("Loading FBA fees...")
    fees_map = load_fees(fees_path)
    print(f"✓ Loaded fees for {len(fees_map)} SKUs")

    # Create workbook
    print("\nGenerating workbook...")
    wb = create_workbook(args.brand, products, fees_map)

    # Save
    print(f"Saving to {out_path}...")
    wb.save(out_path)

    print(f"\n✅ Complete! Workbook saved to:\n   {out_path.resolve()}")
    print(f"\nSheets created:")
    print(f"  • RAW DATA INPUT — Paste inventory data here")
    print(f"  • CATEGORY & COSTS — Configure COGS, FBA fees, categories (yellow cells)")
    print(f"  • PRODUCTS DASHBOARD — Auto-populated overview")
    print(f"  • DEAL CALCULATOR — All 13 promo types side-by-side")
    print(f"  • DEALS PLANNER — All products with period/type selector")
    print(f"  • PEAK CALENDAR — Planning guide & event calendar")
    print(f"\n📌 IMPORTANT:")
    print(f"  • Yellow cells = user input (edit as needed)")
    print(f"  • Green cells = auto-calculated (formulas)")
    print(f"  • UPFRONT FEE = charged once per deal period")
    print(f"  • VARIABLE FEE = charged per unit as % of revenue")
    print(f"\n")

if __name__ == "__main__":
    main()
