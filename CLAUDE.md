# Amazon Deals Planner

A free tool for Amazon sellers that generates a formatted Excel deals planning workbook from Amazon's Deals Recommendation Template and Fee Preview CSV.

**Live URL:** https://amazon-deals-planner.vercel.app
**Stack:** Python · Flask · openpyxl · Supabase · Vercel (primary) + Render (legacy redirect)

---

## Repository Structure

```
app.py                # Flask application — routes, Supabase wiring, file handling
planner.py            # Core engine — workbook generation, fee calculation logic
requirements.txt      # flask, gunicorn, openpyxl, supabase, python-dotenv
templates/
  index.html          # Single-page landing + upload form (Jinja2)
static/
  style.css           # Landing page styles
  favicon.svg         # Dark rounded square with orange 'A' glyph
  favicon.ico         # ICO fallback for older browsers
supabase_setup.sql    # Run once in Supabase SQL Editor — creates leads table + RLS
render.yaml           # Render.com config (legacy — now only redirects to Vercel)
vercel.json           # Vercel serverless config (primary hosting)
Procfile              # gunicorn start command for Render
.env.example          # Template for local .env
DEPLOY.md             # Step-by-step deployment guide
```

---

## Architecture

```
Browser → Vercel (primary host)
              │
              ├── GET  /          → templates/index.html
              └── POST /generate  → process uploads → return .xlsx download
                        │
                        ├── planner.load_deal_recommendations(deals.xlsx)
                        ├── planner.load_fees(fees.csv)
                        ├── planner.create_workbook(brand, recs, fees)
                        │
                        └── Supabase (optional)
                                ├── leads table  (email, brand, marketplace, …)
                                └── uploads bucket (private, brand/email/type_ts.ext)

Render (legacy) → 301 redirect to Vercel when RENDER env var is detected
```

---

## Key Files

### `app.py`

Flask application with two routes:

- `GET /` — renders `templates/index.html`
- `POST /generate` — accepts multipart form upload

**Form fields:** `email` (required), `brand` (required), `store_url`, `marketplace`, `num_skus`
**Files:** `deals` — Amazon Deals Recommendation Template `.xlsx` (required), `fees` — Fee Preview `.csv` or `.xlsx` (required)

**Request flow:**
1. Validates `email`, `brand`, `deals` file are present
2. Saves uploaded files to Supabase Storage (`uploads` bucket, path: `brand/email/type_YYYYMMDD_HHMMSS.ext`)
3. Writes files to `tempfile.NamedTemporaryFile` for processing
4. Calls `load_deal_recommendations()` and `load_fees()` from `planner.py`
5. Calls `create_workbook()` → saves to a temp path → `send_file()` as attachment
6. Calls `save_lead()` to record to Supabase `leads` table (falls back to `leads_log.csv` locally)
7. Deletes temp files via `response.call_on_close`

**Supabase is optional.** If `SUPABASE_URL`/`SUPABASE_KEY` are absent, `supabase = None`, leads append to `leads_log.csv`, uploads are skipped. The generation flow is unaffected.

**Render redirect.** When the `RENDER` env var is set (Render sets this automatically), a `before_request` hook 301-redirects every request to `https://amazon-deals-planner.vercel.app`. This block is skipped on Vercel because Vercel never sets `RENDER`.

---

### `planner.py`

The core engine. Three public functions plus helpers.

#### `load_deal_recommendations(filepath) → list[dict]`

Reads Amazon's Deals Recommendation Template `.xlsx`.
- Target sheet: first sheet whose name contains `"deal recommendation template"` (case-insensitive); falls back to any sheet where row 4 contains "Deal Type"
- **Row 4** = display headers used for column mapping
- **Row 6+** = data rows; row 5 is a field-name row that is skipped
- Column positions are resolved by substring match on lowercased header strings
- Schedule inheritance: the schedule string is only present on the first row of each `recommendation_id` group; subsequent rows in the group inherit it via `rec_schedule_map`
- Cells starting with `=` are formula strings (openpyxl reads them literally) — treated as empty

Returns list of dicts: `parent_asin`, `deal_asin`, `product_name`, `deal_type`, `recommendation_id`, `sku`, `participating`, `schedule`, `seller_price`, `deal_price`, `committed_units`, `seller_quantity`

#### `load_fees(filepath) → (sku_map, asin_map)`

Reads Amazon Fee Preview CSV. Tries encodings: `utf-8-sig` → `utf-8` → `latin-1`.
- Preferred column: `estimated-fee-total` (referral + fulfillment combined)
- Fallback column: `expected-fulfillment-fee-per-unit`
- Multi-marketplace files: filters to rows where `amazon-store == "US"` when US rows are present
- `clean_header()` strips BOM, quotes, and non-alphanumeric prefixes from CSV headers
- Returns `(sku_map, asin_map)` — both `dict[str, float]`

Use `lookup_fee(sku, asin, sku_map, asin_map)` to resolve a fee: tries SKU first, ASIN as fallback, returns `(fee, source)` where source is `"sku"`, `"asin"`, or `"missing"`.

#### `create_workbook(brand_name, recommendations, fees_data) → Workbook`

Builds the output Excel workbook. Accepts `fees_data` as either the `(sku_map, asin_map)` tuple from `load_fees()` or a legacy plain `dict`.

**Sheet 1: DEALS PLANNER** (17 columns A–Q, frozen at row 5)

| Row | Content |
|-----|---------|
| 1 | Main title bar (navy) |
| 2 | Subtitle / source note (blue) |
| 3 | COGS instruction note (yellow) |
| 4 | Column headers |
| 5+ | One data row per SKU |
| … | Missing-fee warning (if any SKUs lacked fee data) |
| … | DEAL SUMMARY: per-group breakdown |
| … | GRAND TOTALS + NET DEAL PROFIT |

Data row columns: SKU, Product Name, Deal Type, Schedule, List Price, Deal Price, Disc%, Committed Units, Deal Revenue, Est. Amazon Fee/Unit, Deal Var Fee/Unit, Total Fees/Unit, Total Fees, COGS/Unit (user fills yellow), Total COGS, SKU Profit, Margin

Cell highlights:
- Red row (`RED_L`) — SKU missing from Fee Preview
- Orange cell (`ORG`) on Var Fee/Unit — cap was applied (prorated rate)
- Yellow cells (`YEL`) — COGS/Unit, user-editable
- Green cells (`GRN`) — profit/margin outputs

**Sheet 2: PEAK CALENDAR**

Static reference table of Amazon peak events (Prime Day, BFCM, Cyber Monday, non-peak) with fee structures, minimum discounts, and strategy notes.

---

### Fee Calculation Logic

| Fee Type | Non-Peak | Prime Day |
|----------|----------|-----------|
| Upfront — Lightning Deal | $70 flat | $100 flat |
| Upfront — Best Deal | $70 × days in schedule | $100 flat |
| Variable rate | 1.0% of deal price/unit | 1.5% of deal price/unit |
| Variable cap (per group) | $2,000 | $5,000 |

**Cap proration:** When a group's raw variable fee (`group_revenue × rate`) exceeds the cap, the effective rate is scaled down:
```python
effective_rate = cap / group_total_revenue
```
This ensures per-row formulas and the group summary sum to the same capped figure.

**Schedule parsing:** `parse_days_from_schedule()` extracts duration from `"Mon (2026-04-20 - 2026-04-26)"` patterns. Defaults to 7 days when no date range is found.

**Prime Day detection:** `is_prime_day(schedule)` checks for `"prime day"` in the schedule string (case-insensitive).

---

## Local Development

```bash
git clone https://github.com/Munkey-Beep/amazon-deals-planner.git
cd amazon-deals-planner
python -m venv venv
source venv/bin/activate      # Windows: venv\Scripts\activate
pip install -r requirements.txt

cp .env.example .env
# Edit .env — Supabase credentials are optional

python app.py
# → http://localhost:5000
```

Without Supabase credentials the app is fully functional — leads write to `leads_log.csv`.

---

## Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `SUPABASE_URL` | No | Supabase project URL (`https://xxx.supabase.co`) |
| `SUPABASE_KEY` | No | Supabase anon public key |
| `FLASK_DEBUG` | No | `true` enables auto-reload in dev (default: `false`) |
| `PORT` | No | Server port (default: `5000`) |
| `RENDER` | Auto | Set by Render automatically; triggers 301 redirect to Vercel |

---

## Database Schema

Single table: `leads` (Supabase / Postgres)

```sql
id           BIGSERIAL PRIMARY KEY
email        TEXT NOT NULL
brand        TEXT NOT NULL
store_url    TEXT DEFAULT ''
marketplace  TEXT DEFAULT 'US'
num_skus     TEXT DEFAULT ''
num_products INTEGER DEFAULT 0   -- number of SKUs processed
num_fees     INTEGER DEFAULT 0   -- number of fee entries loaded
created_at   TIMESTAMPTZ DEFAULT NOW()
```

RLS policies:
- Anonymous key (`anon`) → INSERT only
- Authenticated users → SELECT

Run `supabase_setup.sql` once in the Supabase SQL Editor to create the table, indexes, and policies. The storage bucket (`uploads`, private, 10 MB limit) must be created manually in the Supabase dashboard.

---

## Deployment

### Primary: Vercel

`vercel.json` routes all requests through `app.py` as a Python serverless function:

```json
{
  "builds": [{ "src": "app.py", "use": "@vercel/python", "config": { "maxLambdaSize": "15mb" } }],
  "routes": [{ "src": "/(.*)", "dest": "app.py" }]
}
```

15 MB lambda size is required to fit openpyxl + supabase + dependencies.
Set `SUPABASE_URL` and `SUPABASE_KEY` in Vercel project settings → Environment Variables.

### Legacy: Render

`render.yaml` defines a free Python web service. When the `RENDER` env var is detected at runtime, the Flask app only redirects — it does not process requests. The Render deployment is kept live purely as a permanent redirect for bookmarked/cached URLs.

gunicorn command: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120`

See `DEPLOY.md` for the complete setup walkthrough (Supabase, Render, Vercel, custom domain).

---

## Key Conventions

- **Temp files only in production.** All uploaded files go to `tempfile.NamedTemporaryFile`. Never write to the working directory on the server.
- **Supabase is always optional.** Wrap every Supabase call in try/except. The generation flow must succeed without it.
- **Fee CSV encoding.** The Fee Preview CSV may have BOM, inconsistent quoting, or Latin-1 encoding. Always use `clean_header()` on CSV headers and try multiple encodings via the loop in `load_fees()`.
- **Formula cells in xlsx.** `openpyxl` with `data_only=True` still returns formula strings (e.g., `"=A1"`) for cells that haven't been recalculated. Skip any cell value starting with `=` when parsing Amazon's template.
- **Group-level variable fee cap.** The $2K/$5K cap applies per `recommendation_id` group, not per SKU. Pre-compute `group_effective_rates` before writing data rows — do not compute per-row.
- **Column order is fixed.** `create_workbook` uses hard-coded column indices (A=1 through Q=17). The summary section references data rows by calculated row-number offsets. Inserting or removing columns will break all summary formulas.
- **16 MB upload limit** is set in `app.config['MAX_CONTENT_LENGTH']`. This is intentional; do not raise it without considering Vercel's request size limits.
- **No streaming.** `send_file()` sends the complete workbook. For very large inputs (hundreds of SKUs) this is still fast enough; openpyxl generation is the bottleneck, not I/O.

---

## No Automated Tests

There is no test suite. To verify changes:

1. Run locally: `python app.py`
2. Upload a real Amazon Deals Recommendation Template `.xlsx` + Fee Preview `.csv`
3. Open the generated workbook and confirm:
   - Discount % formula resolves correctly
   - Variable fee cells are orange when the cap fires
   - Red rows appear for SKUs absent from the Fee Preview
   - Deal Summary group profit = revenue − upfront − variable − Amazon fees − COGS
   - Grand total NET DEAL PROFIT matches sum of group profits
4. Test without Supabase env vars to confirm graceful fallback to `leads_log.csv`
