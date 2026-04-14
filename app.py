"""
Amazon Deals Planner — Web Application
Flask + Supabase backend for free Amazon seller tool
"""

import os
import csv
import re
import io
import tempfile
import uuid
import traceback
from datetime import datetime
from pathlib import Path

from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename

# Planner generator
from planner import load_deal_recommendations, load_fees, create_workbook

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB max upload

# ── Supabase (optional — works without it for local dev) ──
SUPABASE_URL = os.environ.get("SUPABASE_URL", "")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "")
supabase = None

if SUPABASE_URL and SUPABASE_KEY:
    try:
        from supabase import create_client
        supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
        print(f"  Supabase client created (URL: {SUPABASE_URL[:40]}...)")
        # Verify connection with a test read
        try:
            test = supabase.table("leads").select("id").limit(1).execute()
            print(f"  Supabase leads table OK — connection verified")
        except Exception as e:
            print(f"  Supabase leads table check failed: {e}")
            traceback.print_exc()
        # Auto-create the uploads bucket if it doesn't exist
        try:
            supabase.storage.create_bucket("uploads", options={"public": False})
            print("  Created 'uploads' storage bucket")
        except Exception:
            pass  # Bucket already exists — that's fine
    except Exception as e:
        print(f"  Supabase not available: {e}")
        traceback.print_exc()


def save_lead(data):
    """Save lead info to Supabase (or log locally if Supabase unavailable)."""
    record = {
        "email": data.get("email", ""),
        "brand": data.get("brand", ""),
        "store_url": data.get("store_url", ""),
        "marketplace": data.get("marketplace", ""),
        "num_skus": data.get("num_skus", ""),
        "num_products": data.get("num_products", 0),
        "num_fees": data.get("num_fees", 0),
        # Let the database set created_at via DEFAULT NOW()
    }

    if supabase:
        try:
            result = supabase.table("leads").insert(record).execute()
            print(f"  Lead saved: {record['email']} ({record['brand']}) — {result}")
        except Exception as e:
            print(f"  Failed to save lead: {e}")
            traceback.print_exc()
    else:
        # Local fallback — append to CSV
        log_path = Path("leads_log.csv")
        write_header = not log_path.exists()
        with open(log_path, "a", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=record.keys())
            if write_header:
                writer.writeheader()
            writer.writerow(record)
        print(f"  Lead logged locally: {record['email']} ({record['brand']})")


def save_upload(file_obj, email, brand, file_type, ext="xlsx"):
    """Save uploaded file to Supabase storage (or skip if unavailable)."""
    if not supabase:
        return

    try:
        filename = f"{brand}/{email}/{file_type}_{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}.{ext}"
        content_type = (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            if ext == "xlsx" else "text/csv"
        )
        file_bytes = file_obj.read()
        file_obj.seek(0)  # Reset for later use

        supabase.storage.from_("uploads").upload(
            filename,
            file_bytes,
            {"content-type": content_type}
        )
        print(f"  Uploaded: {filename}")
    except Exception as e:
        print(f"  Upload storage failed: {e}")


# ── Redirect (Render → Vercel migration) ──────────────────────────────────
# Set REDIRECT_TO=https://your-app.vercel.app on the Render service.
# All requests will 301 to the equivalent path on the new host.
# Leave unset on Vercel so the app runs normally there.

REDIRECT_TO = os.environ.get("REDIRECT_TO", "").rstrip("/")

if REDIRECT_TO:
    @app.before_request
    def redirect_to_new_host():
        target = REDIRECT_TO + request.full_path.rstrip("?")
        from flask import redirect as flask_redirect
        return flask_redirect(target, code=301)

# ── Routes ──

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    try:
        # Get form data
        email = request.form.get("email", "").strip()
        brand = request.form.get("brand", "Brand").strip()
        store_url = request.form.get("store_url", "").strip()
        marketplace = request.form.get("marketplace", "US").strip()
        num_skus = request.form.get("num_skus", "").strip()

        if not email or not brand:
            return "Email and brand name are required", 400

        # Get files
        deals_file = request.files.get("deals")
        fees_file  = request.files.get("fees")

        if not deals_file:
            return "Deals Recommendation Template (.xlsx) is required", 400

        if not fees_file or not fees_file.filename:
            return "Fee Preview CSV is required for accurate referral and fulfillment fee calculations", 400

        # Save uploads to Supabase storage
        save_upload(deals_file, email, brand, "deals_recommendations", ext="xlsx")
        if fees_file and fees_file.filename:
            fees_ext = "csv" if fees_file.filename.lower().endswith(".csv") else "xlsx"
            save_upload(fees_file, email, brand, "fee_preview", ext=fees_ext)

        # Save deals file to temp
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False, mode="wb") as tmp_deals:
            deals_file.save(tmp_deals)
            deals_path = tmp_deals.name

        fees_path = None
        if fees_file and fees_file.filename:
            fees_ext = ".csv" if fees_file.filename.lower().endswith(".csv") else ".xlsx"
            with tempfile.NamedTemporaryFile(suffix=fees_ext, delete=False, mode="wb") as tmp_fees:
                fees_file.save(tmp_fees)
                fees_path = tmp_fees.name

        # Process
        recommendations = load_deal_recommendations(deals_path)
        fees_map = load_fees(fees_path) if fees_path else ({}, {})

        if not recommendations:
            os.unlink(deals_path)
            if fees_path:
                os.unlink(fees_path)
            return (
                "No deal recommendations found. "
                "Make sure it's the Amazon Deals Recommendation Template with the "
                "'Deal Recommendation Template' tab.", 400
            )

        # Generate workbook
        wb = create_workbook(brand, recommendations, fees_map)

        # Save to temp
        output_path = os.path.join(tempfile.gettempdir(), f"{uuid.uuid4().hex}.xlsx")
        wb.save(output_path)

        # Save lead data
        save_lead({
            "email": email,
            "brand": brand,
            "store_url": store_url,
            "marketplace": marketplace,
            "num_skus": num_skus,
            "num_products": len(recommendations),
            "num_fees": len(fees_map[0]) if isinstance(fees_map, tuple) else len(fees_map),
        })

        # Clean up temp files
        os.unlink(deals_path)
        if fees_path:
            os.unlink(fees_path)

        # Send file
        response = send_file(
            output_path,
            as_attachment=True,
            download_name=f"{brand}_Deals_Planner.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Clean up output after sending
        @response.call_on_close
        def cleanup():
            try:
                os.unlink(output_path)
            except:
                pass

        return response

    except Exception as e:
        print(f"  Error generating planner: {e}")
        return f"Error generating planner: {str(e)}", 500


# ── Main ──

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "false").lower() == "true"

    print(f"\n  Amazon Deals Planner Web App")
    print(f"  http://localhost:{port}")
    print(f"  Supabase: {'Connected' if supabase else 'Not configured (using local log)'}\n")

    app.run(host="0.0.0.0", port=port, debug=debug)
