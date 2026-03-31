-- Run this in Supabase SQL Editor (Dashboard > SQL Editor > New Query)

-- Leads table: stores every user who generates a planner
CREATE TABLE IF NOT EXISTS leads (
    id BIGSERIAL PRIMARY KEY,
    email TEXT NOT NULL,
    brand TEXT NOT NULL,
    store_url TEXT DEFAULT '',
    marketplace TEXT DEFAULT 'US',
    num_skus TEXT DEFAULT '',
    num_products INTEGER DEFAULT 0,
    num_fees INTEGER DEFAULT 0,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

-- Index for quick lookups
CREATE INDEX IF NOT EXISTS idx_leads_email ON leads (email);
CREATE INDEX IF NOT EXISTS idx_leads_created ON leads (created_at DESC);

-- Enable Row Level Security (recommended)
ALTER TABLE leads ENABLE ROW LEVEL SECURITY;

-- Policy: allow inserts from the anon key (your web app)
CREATE POLICY "Allow anonymous inserts" ON leads
    FOR INSERT
    TO anon
    WITH CHECK (true);

-- Policy: only authenticated users (you) can read leads
CREATE POLICY "Only authenticated can read" ON leads
    FOR SELECT
    TO authenticated
    USING (true);

-- Storage bucket for uploaded CSV files
-- Go to: Dashboard > Storage > New Bucket
-- Name: uploads
-- Public: OFF (private)
-- File size limit: 10MB
