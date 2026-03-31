# Deployment Guide: Amazon Deals Planner Web App

Get your free Amazon seller tool live in under 30 minutes.

## Step 1: Create a GitHub Repo (5 min)

1. Go to github.com and create a new repository
2. Name it `amazon-deals-planner` (public or private)
3. Upload all files from this folder to the repo
4. Or use git:

```bash
cd amazon-deals-planner-web
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/YOUR_USERNAME/amazon-deals-planner.git
git push -u origin main
```

## Step 2: Set Up Supabase (10 min)

1. Go to **supabase.com** and create a free account
2. Click **New Project** and name it `amazon-deals-planner`
3. Choose a region close to your users (e.g., US East)
4. Wait for the project to initialize (~2 min)

### Create the database table:
5. Go to **SQL Editor** in the left sidebar
6. Click **New Query**
7. Copy and paste the contents of `supabase_setup.sql`
8. Click **Run**

### Create storage bucket:
9. Go to **Storage** in the left sidebar
10. Click **New Bucket**
11. Name it `uploads`
12. Set it to **Private** (not public)
13. Set file size limit to **10 MB**

### Get your credentials:
14. Go to **Settings > API**
15. Copy your **Project URL** (looks like `https://xxxxx.supabase.co`)
16. Copy your **anon public** key (the long string)
17. Save both for the next step

## Step 3: Deploy on Render (10 min)

1. Go to **render.com** and create a free account
2. Click **New > Web Service**
3. Connect your GitHub account and select the `amazon-deals-planner` repo
4. Render auto-detects the config. Verify these settings:
   - **Name**: `amazon-deals-planner`
   - **Runtime**: Python
   - **Build command**: `pip install -r requirements.txt`
   - **Start command**: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120`
   - **Plan**: Free

5. Under **Environment Variables**, add:
   - `SUPABASE_URL` = your Supabase project URL
   - `SUPABASE_KEY` = your Supabase anon key

6. Click **Deploy Web Service**
7. Wait for the build to complete (~3-5 min)
8. Your app is live at: `https://amazon-deals-planner.onrender.com`

## Step 4: Custom Domain (Optional)

### On Render:
1. Go to your service > **Settings > Custom Domains**
2. Add your domain (e.g., `planner.yourbrand.com`)
3. Follow the DNS instructions

### On your domain registrar:
4. Add a CNAME record pointing to `amazon-deals-planner.onrender.com`

## Viewing Your Collected Data

### Via Supabase Dashboard:
1. Go to supabase.com > your project
2. Click **Table Editor** > **leads**
3. Browse all sign-ups with email, brand, marketplace, etc.

### Export to CSV:
1. In Table Editor, click the export button
2. Download as CSV for your CRM or spreadsheet

### Via Supabase Storage:
1. Go to **Storage** > **uploads** bucket
2. Browse uploaded CSV files organized by brand/email

## Architecture Overview

```
User visits site
    |
    v
[Landing Page] -- HTML/CSS/JS (static)
    |
    v
[Lead Form] -- email, brand, store URL, marketplace
    |
    v
[File Upload] -- Manage Inventory CSV + Fee Preview CSV
    |
    v
[Flask Backend]
    |--- Saves lead to Supabase (leads table)
    |--- Saves CSV files to Supabase Storage
    |--- Generates Excel planner
    |
    v
[Download] -- User gets their planner.xlsx
```

## Local Development

```bash
# Clone
git clone https://github.com/YOUR_USERNAME/amazon-deals-planner.git
cd amazon-deals-planner

# Create virtual environment
python -m venv venv
source venv/bin/activate  # Mac/Linux
# venv\Scripts\activate   # Windows

# Install dependencies
pip install -r requirements.txt

# Set environment (copy and edit)
cp .env.example .env
# Edit .env with your Supabase credentials

# Run
python app.py
# Open http://localhost:5000
```

## Costs

Everything in this stack has a generous free tier:

| Service | Free Tier |
|---------|-----------|
| Render | 750 hours/month, auto-sleep after 15 min inactivity |
| Supabase | 500 MB database, 1 GB storage, 50K API requests/month |
| GitHub | Unlimited public/private repos |

You can serve hundreds of sellers per month at zero cost.

## Scaling Up (When You Need It)

| Need | Solution | Cost |
|------|----------|------|
| Faster response | Render Starter plan | $7/month |
| More storage | Supabase Pro | $25/month |
| Custom domain SSL | Included free on Render | $0 |
| Email notifications | Add SendGrid/Resend | Free tier |
| Analytics | Add Plausible or Umami | Free self-hosted |

## Security Notes

- Supabase Row Level Security is enabled on the leads table
- Uploaded files are stored in a private bucket
- No passwords or sensitive data are collected
- HTTPS is automatic on Render
- File uploads are capped at 16 MB

## Troubleshooting

**Build fails on Render:**
- Check that `requirements.txt` is in the root directory
- Verify Python version compatibility

**Supabase not saving leads:**
- Check environment variables are set correctly
- Verify the `leads` table was created (run the SQL again)
- Check Supabase dashboard for API errors

**Large files timing out:**
- Increase worker timeout in Procfile (`--timeout 300`)
- Consider upgrading to Render Starter for better performance

**App sleeps on free tier:**
- Render free tier auto-sleeps after 15 min of no traffic
- First request after sleep takes ~30 seconds to wake up
- Upgrade to Starter ($7/mo) to keep it always running
