#!/usr/bin/env python3
"""
AMAZON DEALS PLANNER GENERATOR — Deal Recommendation Template Based
============================================================================
Input:  Amazon Deals Recommendation Template (.xlsx)  +  Fee Preview CSV
Output: Excel workbook with DEALS PLANNER + PEAK CALENDAR

Only reads data from the 'Deal Recommendation Template' tab of the Amazon file.
Fees come from the Amazon Fee Preview report (estimated-fee-total per SKU).
"""

import csv
import re
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ═════════════════════════════════════════════════════════════════════════════
# CONSTANTS & HELPERS
# ═════════════════════════════════════════════════════════════════════════════

NAVY       = "1F4E79"
BLUE       = "2E75B6"
LT_BLUE    = "BDD7EE"
YEL        = "FFF2CC"
GRN        = "E2EFDA"
GRN_DARK   = "70AD47"
RED_L      = "FFE0E0"
WHT        = "FFFFFF"
ALT        = "F5F9FF"
GRY        = "F2F2F2"
ORG        = "FCE4D6"
ORG_DARK   = "F4B942"
PURPLE     = "7030A0"

FONT_NAME  = "Calibri"


def col_letter(n):
    return get_column_letter(n)


def fill(c):
    return PatternFill("solid", fgColor=c, start_color=c)


def bdr(color="9DC3E6"):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)


def font(bold=False, color="000000", sz=10, italic=False, name=None):
    return Font(name=name or FONT_NAME, bold=bold, color=color, size=sz, italic=italic)


def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[col_letter(i)].width = w


# ═════════════════════════════════════════════════════════════════════════════
# DATA LOADING
# ═════════════════════════════════════════════════════════════════════════════

def clean_header(h):
    """Strip BOM, quotes, non-alphanumeric prefixes from CSV/header strings."""
    h = h.strip().lstrip('\ufeff')
    h = re.sub(r'^[^a-zA-Z]+', '', h)
    return h.strip('"').strip("'").strip()


def load_deal_recommendations(filepath):
    """
    Read the 'Deal Recommendation Template' sheet from Amazon's .xlsx file.
    Headers are in row 4, field-name row in row 5, data starts row 6.
    Schedule is inherited from the first row of each recommendation_id group.
    Returns a list of dicts.
    """
    recs = []
    try:
        wb_src = load_workbook(filepath, data_only=True)

        # Find target sheet
        target = None
        for name in wb_src.sheetnames:
            if "deal recommendation template" in name.lower():
                target = wb_src[name]
                break
        if target is None:
            # Fall back: find any sheet where row 4 has "Deal Type"
            for name in wb_src.sheetnames:
                ws = wb_src[name]
                for col in range(1, 25):
                    v = ws.cell(4, col).value
                    if v and "deal type" in str(v).lower():
                        target = ws
                        break
                if target:
                    break

        if target is None:
            print("⚠️  'Deal Recommendation Template' sheet not found")
            return []

        # Map column positions from row 4 (display headers)
        headers_row4 = []
        for col in range(1, target.max_column + 1):
            v = target.cell(4, col).value
            headers_row4.append(str(v).lower().strip() if v else "")

        def find_col(needle):
            for i, h in enumerate(headers_row4):
                if needle in h:
                    return i  # 0-based
            return -1

        ci = {
            "parent_asin":    find_col("parent asin"),
            "deal_asin":      find_col("deal asin"),
            "product_name":   find_col("product name"),
            "deal_type":      find_col("deal type"),
            "rec_id":         find_col("recommendation id"),
            "sku":            find_col("sku"),
            "participating":  find_col("participating"),
            "schedule":       find_col("schedule"),
            "seller_price":   find_col("seller price"),
            "deal_price":     find_col("deal price"),
            "committed":      find_col("committed units"),
            "seller_qty":     find_col("seller quantity"),
        }

        # Track schedule inheritance by recommendation_id
        rec_schedule_map = {}

        for row in range(6, target.max_row + 1):
            def cv(key):
                idx = ci.get(key, -1)
                if idx < 0:
                    return None
                return target.cell(row, idx + 1).value

            sku = str(cv("sku") or "").strip()
            if not sku:
                continue

            deal_type = str(cv("deal_type") or "").strip()
            rec_id    = str(cv("rec_id")    or "").strip()

            # Schedule: only real strings, not formula strings
            raw_sched = cv("schedule")
            if raw_sched and not str(raw_sched).startswith("="):
                schedule = str(raw_sched).strip()
                if rec_id and rec_id not in rec_schedule_map:
                    rec_schedule_map[rec_id] = schedule
            else:
                schedule = rec_schedule_map.get(rec_id, "")

            def to_float(v):
                try:
                    return float(v) if v not in (None, "", "--") else 0.0
                except (ValueError, TypeError):
                    return 0.0

            def to_int(v):
                try:
                    return int(float(v)) if v not in (None, "", "--") else 0
                except (ValueError, TypeError):
                    return 0

            recs.append({
                "parent_asin":    str(cv("parent_asin") or "").strip(),
                "deal_asin":      str(cv("deal_asin")   or "").strip(),
                "product_name":   str(cv("product_name") or "").strip(),
                "deal_type":      deal_type,
                "recommendation_id": rec_id,
                "sku":            sku,
                "participating":  str(cv("participating") or "Yes").strip(),
                "schedule":       schedule,
                "seller_price":   to_float(cv("seller_price")),
                "deal_price":     to_float(cv("deal_price")),
                "committed_units": to_int(cv("committed")),
                "seller_quantity": to_int(cv("seller_qty")),
            })

        print(f"✓ Loaded {len(recs)} deal recommendations from '{target.title}'")
        return recs

    except Exception as e:
        print(f"❌ Error loading deal recommendations: {e}")
        import traceback; traceback.print_exc()
        return []


def load_fees(filepath):
    """
    Load Amazon Fee Preview CSV.
    Prefers 'estimated-fee-total' (referral + fulfillment).
    Falls back to 'expected-fulfillment-fee-per-unit' if total not available.
    Returns dict: SKU -> float
    """
    fees_map = {}
    if not filepath:
        return fees_map
    try:
        for enc in ("utf-8-sig", "utf-8", "latin-1"):
            try:
                with open(filepath, encoding=enc) as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        cleaned = {
                            clean_header(k): v.strip().strip('"') if v else ""
                            for k, v in row.items() if k
                        }
                        sku = cleaned.get("sku", "").strip()
                        if sku:
                            fee_str = cleaned.get("estimated-fee-total", "").strip()
                            if not fee_str or fee_str == "--":
                                fee_str = cleaned.get("expected-fulfillment-fee-per-unit", "0").strip()
                            try:
                                fees_map[sku] = float(fee_str) if fee_str and fee_str != "--" else 0.0
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
# FEE / SCHEDULE HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def is_prime_day(schedule):
    return bool(schedule and "prime day" in str(schedule).lower())


def parse_days_from_schedule(schedule):
    """Extract duration in days from 'Mon (2026-04-20 - 2026-04-26)' patterns."""
    if not schedule:
        return 7
    m = re.search(r'(\d{4}-\d{2}-\d{2})\s*[-–]\s*(\d{4}-\d{2}-\d{2})', str(schedule))
    if m:
        try:
            d1 = datetime.strptime(m.group(1), "%Y-%m-%d")
            d2 = datetime.strptime(m.group(2), "%Y-%m-%d")
            return max(1, (d2 - d1).days + 1)
        except Exception:
            pass
    return 7


def compute_upfront_fee(deal_type, schedule):
    """One-time upfront fee per deal submission."""
    if is_prime_day(schedule):
        return 100.0
    if deal_type == "Lightning Deal":
        return 70.0          # single slot (≤12 hrs)
    else:                    # Best Deal
        days = parse_days_from_schedule(schedule)
        return 70.0 * days


def var_fee_rate(schedule):
    return 0.015 if is_prime_day(schedule) else 0.01


def var_fee_cap(schedule):
    return 5000.0 if is_prime_day(schedule) else 2000.0


# ═════════════════════════════════════════════════════════════════════════════
# WORKBOOK CREATION
# ═════════════════════════════════════════════════════════════════════════════

def create_workbook(brand_name, recommendations, fees_map):
    """
    Build the Amazon Deals Planner Excel workbook.
    Sheets: DEALS PLANNER, PEAK CALENDAR
    """
    wb = Workbook()

    # ─────────────────────────────────────────────────────────────────────────
    # SHEET 1 — DEALS PLANNER
    # ─────────────────────────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "DEALS PLANNER"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A5"

    NUM_COLS = 17   # A through Q
    last_col  = col_letter(NUM_COLS)

    # ── Row 1: Main title ──────────────────────────────────────────────────
    ws.merge_cells(f"A1:{last_col}1")
    c = ws.cell(1, 1, f"🛒  DEALS PLANNER — {brand_name}")
    c.font      = Font(name=FONT_NAME, bold=True, color=WHT, size=16)
    c.fill      = fill(NAVY)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    # ── Row 2: Subtitle ───────────────────────────────────────────────────
    ws.merge_cells(f"A2:{last_col}2")
    c = ws.cell(2, 1,
        "★ All deals are Amazon-recommended  |  "
        "Source: Amazon Deals Recommendation Template  |  "
        "Fees: Amazon Fee Preview (Referral + FBA)")
    c.font      = Font(name=FONT_NAME, italic=True, color=WHT, size=10)
    c.fill      = fill(BLUE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 18

    # ── Row 3: COGS note ──────────────────────────────────────────────────
    ws.merge_cells(f"A3:{last_col}3")
    c = ws.cell(3, 1,
        "  ★ = Amazon Recommended Deal  |  "
        "COGS / Unit (column N) is optional — fill yellow cells to calculate true net profit & margin.")
    c.font      = Font(name=FONT_NAME, italic=True, color="595959", size=9)
    c.fill      = fill(YEL)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[3].height = 16

    # ── Row 4: Column Headers ─────────────────────────────────────────────
    HEADERS = [
        # (label, background)
        ("SKU",                         LT_BLUE),
        ("Product Name",                LT_BLUE),
        ("Deal Type ★",                 LT_BLUE),
        ("Schedule / Period",           LT_BLUE),
        ("List Price ($)",              LT_BLUE),
        ("Deal Price ($)",              LT_BLUE),
        ("Disc %",                      LT_BLUE),
        ("Committed Units",             LT_BLUE),
        ("Deal Revenue ($)",            BLUE),
        ("Est. Amazon Fee / Unit ($)",  NAVY),
        ("Deal Var Fee / Unit ($)",     NAVY),
        ("Total Fees / Unit ($)",       NAVY),
        ("Total Fees ($)",              NAVY),
        ("COGS / Unit* ($)",            YEL),
        ("Total COGS* ($)",             YEL),
        ("SKU Profit* ($)",             GRN_DARK),
        ("Margin*",                     GRN_DARK),
    ]
    ws.row_dimensions[4].height = 38
    for col_i, (hdr, bg) in enumerate(HEADERS, 1):
        c = ws.cell(4, col_i, hdr)
        txt_col = WHT if bg in (NAVY, BLUE, GRN_DARK) else "1F1F1F"
        c.font      = Font(name=FONT_NAME, bold=True, color=txt_col, size=9)
        c.fill      = fill(bg)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = bdr()

    # ── Group recommendations by recommendation_id (preserve order) ───────
    rec_groups  = {}   # rid -> [rec, ...]
    rec_order   = []
    for rec in recommendations:
        rid = rec["recommendation_id"]
        if rid not in rec_groups:
            rec_groups[rid]  = []
            rec_order.append(rid)
        rec_groups[rid].append(rec)

    ordered_recs = []
    for rid in rec_order:
        ordered_recs.extend(rec_groups[rid])

    # ── Data Rows ─────────────────────────────────────────────────────────
    DATA_START = 5
    row = DATA_START

    for i, rec in enumerate(ordered_recs):
        r        = row
        sku      = rec["sku"]
        schedule = rec["schedule"]
        dtype    = rec["deal_type"]
        sp       = rec["seller_price"]
        dp       = rec["deal_price"]
        units    = rec["committed_units"]
        amz_fee  = fees_map.get(sku, 0.0)
        rate     = var_fee_rate(schedule)

        row_bg = ALT if i % 2 == 0 else WHT

        def wc(col_i, value, fmt=None, bold=False, bg_ov=None, align="center", italic=False):
            c = ws.cell(r, col_i, value)
            c.font      = Font(name=FONT_NAME, bold=bold, italic=italic,
                               color="000000", size=9)
            c.fill      = fill(bg_ov if bg_ov is not None else row_bg)
            c.alignment = Alignment(horizontal=align, vertical="center",
                                    wrap_text=(col_i == 2))
            c.border    = bdr()
            if fmt:
                c.number_format = fmt
            return c

        wc(1,  sku,                 align="left")
        wc(2,  rec["product_name"][:90],  align="left")
        wc(3,  dtype)
        wc(4,  schedule or "Non-Peak")
        wc(5,  sp,                  "$#,##0.00")
        wc(6,  dp,                  "$#,##0.00")
        wc(7,  f"=IFERROR((E{r}-F{r})/E{r},0)", "0%")
        wc(8,  units,               "#,##0")
        wc(9,  f"=F{r}*H{r}",      "$#,##0.00")                         # Deal Revenue
        wc(10, amz_fee if amz_fee > 0 else 0, "$#,##0.00")              # Est. Amazon Fee/Unit
        wc(11, f"=F{r}*{rate}",    "$#,##0.00")                         # Deal Var Fee/Unit
        wc(12, f"=J{r}+K{r}",      "$#,##0.00")                         # Total Fees/Unit
        wc(13, f"=L{r}*H{r}",      "$#,##0.00")                         # Total Fees
        wc(14, "",                  "$#,##0.00", bg_ov=YEL)             # COGS/Unit (user fills)
        wc(15, f"=IF(N{r}>0,N{r}*H{r},0)", "$#,##0.00", bg_ov=YEL)    # Total COGS
        wc(16, f"=I{r}-M{r}-O{r}", "$#,##0.00", bg_ov=GRN,
           bold=True)                                                     # SKU Profit*
        wc(17, f"=IFERROR(P{r}/I{r},0)", "0.0%", bg_ov=GRN)            # Margin*

        ws.row_dimensions[r].height = 20
        row += 1

    DATA_END = row - 1

    # ── Gap before summary ────────────────────────────────────────────────
    row += 1
    SUMMARY_TITLE_ROW = row

    # ── DEAL SUMMARY title bar ────────────────────────────────────────────
    ws.merge_cells(f"A{row}:{last_col}{row}")
    c = ws.cell(row, 1, "📊  DEAL SUMMARY — Upfront fee charged ONCE per deal group (not per SKU)")
    c.font      = Font(name=FONT_NAME, bold=True, color=WHT, size=12)
    c.fill      = fill(NAVY)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 24
    row += 1

    # Helper: write a single summary row
    def s_row(label, value, fmt="$#,##0.00", note="", label_bold=False,
              val_bold=False, bg=GRY, note_bg=None):
        ws.merge_cells(f"A{row}:E{row}")
        c = ws.cell(row, 1, f"  {label}")
        c.font      = Font(name=FONT_NAME, bold=label_bold, color="1F1F1F", size=10)
        c.fill      = fill(bg)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

        ws.merge_cells(f"F{row}:I{row}")
        c = ws.cell(row, 6, value)
        c.font      = Font(name=FONT_NAME, bold=val_bold, color="1F1F1F", size=10)
        c.number_format = fmt
        c.fill      = fill(bg)
        c.alignment = Alignment(horizontal="right", vertical="center")

        if note:
            ws.merge_cells(f"J{row}:{last_col}{row}")
            c = ws.cell(row, 10, f"  {note}")
            c.font      = Font(name=FONT_NAME, italic=True, color="595959", size=8)
            c.fill      = fill(note_bg or bg)
            c.alignment = Alignment(horizontal="left", vertical="center")

        ws.row_dimensions[row].height = 18

    # ── Per-group breakdown ───────────────────────────────────────────────
    # Pre-compute row mapping: sku+rec_id → data row number
    data_row_map = {}
    for i, rec in enumerate(ordered_recs):
        data_row_map[(rec["sku"], rec["recommendation_id"])] = DATA_START + i

    for g_idx, rid in enumerate(rec_order):
        group   = rec_groups[rid]
        first   = group[0]
        sched   = first["schedule"]
        dtype   = first["deal_type"]
        prime   = is_prime_day(sched)
        rate    = var_fee_rate(sched)
        cap     = var_fee_cap(sched)
        upfront = compute_upfront_fee(dtype, sched)

        # Compute group totals from actual data
        grp_revenue   = sum(r["deal_price"] * r["committed_units"] for r in group)
        grp_raw_var   = sum(r["deal_price"] * r["committed_units"] * rate for r in group)
        grp_var_capped = min(grp_raw_var, cap)
        grp_amz_fees  = sum(fees_map.get(r["sku"], 0.0) * r["committed_units"] for r in group)

        # Group header row
        period_tag = "🔥 PRIME DAY" if prime else "📅 Non-Peak"
        bg_grp = ORG if prime else LT_BLUE

        ws.merge_cells(f"A{row}:{last_col}{row}")
        c = ws.cell(row, 1,
            f"  Deal Group {g_idx + 1}:  {dtype}  │  {sched or 'Non-Peak'}  │  {period_tag}  │  {len(group)} SKU(s)")
        c.font      = Font(name=FONT_NAME, bold=True, color="1F1F1F", size=10)
        c.fill      = fill(bg_grp)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row].height = 20
        row += 1

        s_row("Deal Revenue", grp_revenue, bg=GRY, note=f"{len(group)} SKU(s)")
        row += 1

        upfront_note = (
            "$100 flat — Prime Day" if prime else
            (f"$70 flat — Lightning Deal (single slot)" if dtype == "Lightning Deal"
             else f"${upfront:,.0f}  =  {parse_days_from_schedule(sched)} days × $70/day")
        )
        s_row("(Less)  Upfront Fee — ONE-TIME per deal", upfront,
              note=upfront_note, bg=RED_L)
        row += 1

        var_note = (
            f"{rate*100:.0f}% of deal price  |  "
            f"Cap ${cap:,.0f}  |  "
            f"Raw ${grp_raw_var:,.2f}"
            + ("  ✓ CAPPED" if grp_raw_var > cap else "")
        )
        s_row("(Less)  Variable Deal Fees (capped)", grp_var_capped,
              note=var_note, bg=RED_L)
        row += 1

        s_row("(Less)  Est. Amazon Fees  (Referral + FBA)", grp_amz_fees,
              note="From Amazon Fee Preview — referral + FBA fulfillment combined", bg=RED_L)
        row += 1

        grp_rows = [data_row_map[(r["sku"], rid)] for r in group]
        cogs_formula = f"=SUM({','.join(f'O{dr}' for dr in grp_rows)})"
        s_row("(Less)  Total COGS*", cogs_formula,
              note="Fill COGS/Unit (yellow, column N) in the rows above", bg=YEL)
        row += 1

        # Group sub-profit
        grp_profit_ex = grp_revenue - upfront - grp_var_capped - grp_amz_fees
        ws.merge_cells(f"A{row}:E{row}")
        c = ws.cell(row, 1, "  → Est. Group Profit*  (before COGS)")
        c.font      = Font(name=FONT_NAME, bold=True, color="1F4E79", size=10)
        c.fill      = fill(GRN)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

        ws.merge_cells(f"F{row}:I{row}")
        c = ws.cell(row, 6, grp_profit_ex)
        c.font         = Font(name=FONT_NAME, bold=True, color="1F4E79", size=10)
        c.number_format = "$#,##0.00"
        c.fill          = fill(GRN)
        c.alignment     = Alignment(horizontal="right", vertical="center")

        ws.merge_cells(f"J{row}:{last_col}{row}")
        c = ws.cell(row, 10, "  * Subtract COGS (fill column N above) for true net profit")
        c.font      = Font(name=FONT_NAME, italic=True, color="595959", size=8)
        c.fill      = fill(GRN)
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[row].height = 20
        row += 2   # blank separator between groups

    # ── GRAND TOTALS ──────────────────────────────────────────────────────
    ws.merge_cells(f"A{row}:{last_col}{row}")
    c = ws.cell(row, 1, "💰  GRAND TOTALS — All Deals Combined")
    c.font      = Font(name=FONT_NAME, bold=True, color=WHT, size=12)
    c.fill      = fill(BLUE)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 24
    row += 1

    total_revenue  = sum(r["deal_price"] * r["committed_units"] for r in ordered_recs)
    total_upfront  = sum(
        compute_upfront_fee(rec_groups[rid][0]["deal_type"], rec_groups[rid][0]["schedule"])
        for rid in rec_order
    )
    all_raw_var = {
        rid: sum(r["deal_price"] * r["committed_units"] * var_fee_rate(rec_groups[rid][0]["schedule"])
                 for r in rec_groups[rid])
        for rid in rec_order
    }
    total_capped_var = sum(min(v, var_fee_cap(rec_groups[rid][0]["schedule"]))
                           for rid, v in all_raw_var.items())
    total_amz        = sum(fees_map.get(r["sku"], 0.0) * r["committed_units"] for r in ordered_recs)
    grand_ex_cogs    = total_revenue - total_upfront - total_capped_var - total_amz

    def g_row(label, value, fmt="$#,##0.00", note="", bold=False, bg=GRY):
        ws.merge_cells(f"A{row}:E{row}")
        c = ws.cell(row, 1, f"  {label}")
        c.font      = Font(name=FONT_NAME, bold=bold, color="1F1F1F", size=10)
        c.fill      = fill(bg)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

        ws.merge_cells(f"F{row}:I{row}")
        c = ws.cell(row, 6, value)
        c.font         = Font(name=FONT_NAME, bold=bold, color="1F1F1F", size=10)
        c.number_format = fmt
        c.fill          = fill(bg)
        c.alignment     = Alignment(horizontal="right", vertical="center")

        if note:
            ws.merge_cells(f"J{row}:{last_col}{row}")
            c = ws.cell(row, 10, f"  {note}")
            c.font      = Font(name=FONT_NAME, italic=True, color="595959", size=8)
            c.fill      = fill(bg)
            c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[row].height = 18

    g_row("Total Deal Revenue",
          total_revenue, bold=True,
          note=f"{len(ordered_recs)} SKU(s) across {len(rec_order)} deal group(s)")
    row += 1

    g_row("(Less)  Total Upfront Fees",
          -total_upfront, bg=RED_L,
          note=f"{len(rec_order)} deal group(s) — one upfront fee per group")
    row += 1

    g_row("(Less)  Total Variable Deal Fees (capped per group)",
          -total_capped_var, bg=RED_L,
          note="Each group capped separately ($2K Non-Peak / $5K Prime Day)")
    row += 1

    g_row("(Less)  Total Est. Amazon Fees  (Referral + FBA)",
          -total_amz, bg=RED_L,
          note="Referral + fulfillment from Amazon Fee Preview")
    row += 1

    g_row("(Less)  Total COGS*",
          f"=-SUM(O{DATA_START}:O{DATA_END})",
          note="Fill COGS/Unit (yellow column N) in data rows above", bg=YEL)
    row += 1

    # NET PROFIT bar
    ws.merge_cells(f"A{row}:E{row}")
    c = ws.cell(row, 1, "  💰  NET DEAL PROFIT*")
    c.font      = Font(name=FONT_NAME, bold=True, color=WHT, size=13)
    c.fill      = fill(NAVY)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws.merge_cells(f"F{row}:I{row}")
    c = ws.cell(row, 6, grand_ex_cogs)
    c.font         = Font(name=FONT_NAME, bold=True, color=WHT, size=13)
    c.number_format = "$#,##0.00"
    c.fill          = fill(NAVY)
    c.alignment     = Alignment(horizontal="right", vertical="center")

    ws.merge_cells(f"J{row}:{last_col}{row}")
    c = ws.cell(row, 10, "  * Before COGS — fill column N above for true net profit")
    c.font      = Font(name=FONT_NAME, italic=True, color="9DC3E6", size=9)
    c.fill      = fill(NAVY)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 28
    row += 1

    # Net Margin
    net_margin = grand_ex_cogs / total_revenue if total_revenue > 0 else 0
    ws.merge_cells(f"A{row}:E{row}")
    c = ws.cell(row, 1, "  Net Margin*  (before COGS)")
    c.font      = Font(name=FONT_NAME, bold=True, color=WHT, size=11)
    c.fill      = fill(BLUE)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws.merge_cells(f"F{row}:I{row}")
    c = ws.cell(row, 6, net_margin)
    c.font         = Font(name=FONT_NAME, bold=True, color=WHT, size=11)
    c.number_format = "0.0%"
    c.fill          = fill(BLUE)
    c.alignment     = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[row].height = 22
    row += 2

    # HOW IT WORKS note
    ws.merge_cells(f"A{row}:{last_col}{row}")
    c = ws.cell(row, 1,
        "ℹ️  HOW IT WORKS:  All deals are sourced from Amazon's Deals Recommendation Template (Amazon-vetted). "
        "Upfront fee = one charge per deal group submission (not per SKU). "
        "Variable fee = Deal Price × rate per unit (1% Non-Peak / 1.5% Prime Day), capped at $2,000 or $5,000 per group. "
        "Est. Amazon Fee = referral + FBA fulfillment combined (from Fee Preview). "
        "COGS/Unit is optional — fill yellow column N to see true net profit.")
    c.font      = Font(name=FONT_NAME, italic=True, color="595959", size=8)
    c.fill      = fill(GRY)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    ws.row_dimensions[row].height = 36

    # ── Column widths ─────────────────────────────────────────────────────
    set_col_widths(ws, [18, 42, 14, 30, 12, 12, 8, 12, 14, 14, 14, 14, 14, 12, 12, 14, 10])

    # ─────────────────────────────────────────────────────────────────────────
    # SHEET 2 — PEAK CALENDAR
    # ─────────────────────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("PEAK CALENDAR")
    ws2.sheet_view.showGridLines = False

    ws2.merge_cells("A1:H1")
    c = ws2.cell(1, 1, f"📆  PEAK EVENT CALENDAR — {brand_name} 2025/2026")
    c.font      = Font(name=FONT_NAME, bold=True, color=WHT, size=13)
    c.fill      = fill(NAVY)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 28

    ws2.merge_cells("A2:H2")
    c = ws2.cell(2, 1,
        "Planning guide for peak events with fee structures and submission deadlines. "
        "Customize strategy notes for your brand.")
    c.font      = Font(name=FONT_NAME, italic=True, color="555555", size=9)
    c.fill      = fill(GRY)
    ws2.row_dimensions[2].height = 18

    cal_hdrs = [
        ("Event", NAVY), ("Dates", BLUE), ("Best Deal Fee", NAVY),
        ("Lightning Fee", NAVY), ("Coupon Fee", BLUE), ("Min Discount", BLUE),
        ("Strategy Notes", BLUE), ("Brand Opportunity", NAVY),
    ]
    for i, (h, bg) in enumerate(cal_hdrs, 1):
        c = ws2.cell(3, i, h)
        c.font      = Font(name=FONT_NAME, bold=True,
                           color=WHT if bg == NAVY else "1F1F1F", size=10)
        c.fill      = fill(bg)
        c.border    = bdr()
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    ws2.row_dimensions[3].height = 28

    events = [
        ("Prime Day", "~July (est.)", 100, 100, "$5+2.5%", 0.15,
         "Submit by Apr 30 to save $50/deal. $100 flat upfront + 1.5% variable (capped $5K). "
         "Same fee for both LD and BD.",
         "🔥 HIGH — Peak summer shopping."),
        ("Black Friday / BFCM", "~Late November", "TBD", "TBD", 245, 0.15,
         "Submit 6+ weeks early. Biggest traffic day. Variable fee ~1.5% capped $5K. "
         "Upfront TBD — check Seller Central.",
         "🔥 HIGHEST — Top gift-giving event."),
        ("Prime Big Deal Days", "~October (est.)", "TBD", "TBD", "$5+2.5%", 0.15,
         "Submit 4+ weeks ahead. Variable fee ~1.5% capped $5K. Upfront TBD.",
         "🟡 MEDIUM — Pre-holiday push."),
        ("Cyber Monday", "~December 2", 1000, 500, 245, 0.15,
         "Span Best Deal across BFCM + Cyber Monday for single $1,000 fee. Layer coupons.",
         "🔥 HIGH — Online-focused buyers."),
        ("Non-Peak Deals", "Year-round", "$70/day", "$70 flat", "$5+2.5%", 0.15,
         "Best ROI for margins. Run 7–14 day Best Deals at $70/day. 1% variable, capped $2K.",
         "🟡 MEDIUM — Consistent, predictable fees."),
    ]

    fmts = [None, None, "$#,##0", "$#,##0", "$#,##0", "0%", None, None]
    for idx, evt_row in enumerate(events):
        r = idx + 4
        bg = WHT if idx % 2 == 0 else ALT
        for ci, (val, fmt) in enumerate(zip(evt_row, fmts), 1):
            c = ws2.cell(r, ci, val)
            c.font      = Font(name=FONT_NAME, size=9, bold=(ci in [1, 8]))
            c.fill      = fill(bg)
            c.border    = bdr()
            c.alignment = Alignment(
                horizontal="left" if ci in [1, 6, 7] else "center",
                wrap_text=True)
            if fmt and isinstance(val, (int, float)):
                c.number_format = fmt
        ws2.row_dimensions[r].height = 48

    set_col_widths(ws2, [24, 18, 12, 12, 14, 12, 46, 36])

    print(f"✓ Workbook generated: {len(ordered_recs)} SKUs | {len(rec_order)} deal group(s)")
    return wb


# ═════════════════════════════════════════════════════════════════════════════
# MAIN (CLI)
# ═════════════════════════════════════════════════════════════════════════════

def main():
    import argparse
    parser = argparse.ArgumentParser(
        description="Generate Amazon Deals Planner from Deals Recommendation Template")
    parser.add_argument("--brand",    required=True,
                        help="Brand name (e.g. 'Grillbot')")
    parser.add_argument("--deals",    required=True,
                        help="Path to Amazon Deals Recommendation Template (.xlsx)")
    parser.add_argument("--fees",     required=False, default=None,
                        help="Path to Amazon Fee Preview CSV (optional)")
    parser.add_argument("--output",   required=True,
                        help="Output .xlsx path")
    args = parser.parse_args()

    print(f"\n  Amazon Deals Planner — {args.brand}")
    print("  Loading deal recommendations...")
    recs = load_deal_recommendations(args.deals)
    if not recs:
        print("❌  No recommendations found. Check the file and sheet name.")
        sys.exit(1)

    fees_map = {}
    if args.fees:
        print("  Loading fee preview...")
        fees_map = load_fees(args.fees)
        print(f"  Fees loaded: {len(fees_map)} SKUs")

    wb = create_workbook(args.brand, recs, fees_map)
    wb.save(args.output)
    print(f"\n  ✅  Saved: {args.output}\n")


if __name__ == "__main__":
    main()
