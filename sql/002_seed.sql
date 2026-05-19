-- ============================================================
-- TaxlyCMS — Seed Data (Demo / First Run)
-- Run AFTER 001_schema.sql
--
-- Usage:
--   psql $DATABASE_URL -f sql/002_seed.sql
--
-- NOTE: This is for demo purposes only.
--       Change all passwords before going live!
-- ============================================================

-- Default pricing plans
INSERT INTO plans (id, name, description, price_monthly, price_annual, max_users, max_companies, max_documents, features, is_active)
VALUES
  (gen_random_uuid()::text, 'Starter',      'Ideal for small firms',      2999.00, 29990.00,  3,  3,  50,  '["companies","directors","tasks","documents"]'::jsonb,         1),
  (gen_random_uuid()::text, 'Professional', 'Most popular — full access', 6999.00, 69990.00, 10, 10, 500, '["companies","directors","tasks","documents","reports","bulk_upload","dsc","meetings"]'::jsonb, 1),
  (gen_random_uuid()::text, 'Enterprise',   'For large CA firms',            0.00,     0.00, 50, 50, 999, '["all"]'::jsonb, 1)
ON CONFLICT DO NOTHING;

-- Default tenant
INSERT INTO tenants (id, name, slug, email, phone, address, plan_id, status, max_users, max_companies)
SELECT
  'default-tenant-001',
  'Taxly India Private Limited',
  'taxlyindia',
  'info@taxlyindia.com',
  '+91 88829 35471',
  'L-30B, LGF, Malviya Nagar, New Delhi-110017',
  p.id,
  'active',
  10,
  10
FROM plans p WHERE p.name = 'Professional' LIMIT 1
ON CONFLICT (id) DO NOTHING;

-- Platform Admin (is_platform_admin = 1, no tenant)
-- Password: platform@2025  (SHA-256)
INSERT INTO users (id, name, email, password, role, is_active, is_platform_admin, tenant_id)
VALUES (
  gen_random_uuid()::text,
  'Platform Admin',
  'platform@taxlycms.in',
  encode(digest('platform@2025', 'sha256'), 'hex'),
  'superadmin', 1, 1, NULL
) ON CONFLICT (email) DO NOTHING;

-- Tenant users  (password = SHA-256 of the plain text shown)
-- admin123 / manager123 / staff123
INSERT INTO users (id, name, email, password, role, is_active, is_platform_admin, tenant_id) VALUES
  (gen_random_uuid()::text, 'Super Admin',  'admin@compli.in',   encode(digest('admin123',   'sha256'), 'hex'), 'superadmin', 1, 0, 'default-tenant-001'),
  (gen_random_uuid()::text, 'Priya Sharma', 'manager@compli.in', encode(digest('manager123', 'sha256'), 'hex'), 'manager',    1, 0, 'default-tenant-001'),
  (gen_random_uuid()::text, 'Rahul Verma',  'staff@compli.in',   encode(digest('staff123',   'sha256'), 'hex'), 'staff',      1, 0, 'default-tenant-001')
ON CONFLICT (email) DO NOTHING;

-- ⚠️  Change all passwords immediately after first login!
-- Platform Admin: platform@taxlycms.in / platform@2025
-- Tenant Admin:   admin@compli.in      / admin123
-- Manager:        manager@compli.in    / manager123
-- Staff:          staff@compli.in      / staff123
