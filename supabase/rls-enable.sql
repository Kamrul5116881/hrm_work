-- ============================================================================
-- Hastizam HRM — Supabase hardening (Phase 3)
-- Run this in the Supabase SQL Editor AFTER the data migration is verified
-- and the app has been cut over.
--
-- Why this is safe: the Express/Vercel APIs connect as the table OWNER via
-- Prisma, and table owners bypass RLS unless FORCE ROW LEVEL SECURITY is set
-- (we deliberately do not use FORCE). What RLS blocks is anonymous/authed
-- access through the Supabase REST API (PostgREST) — which nothing uses yet,
-- so enabling it now is a pure security win with zero app impact.
--
-- Policies for supabase-js clients are intentionally NOT added here; they
-- belong to a later phase once client-side access patterns are decided.
-- ============================================================================

alter table public."Employee"         enable row level security;
alter table public."Attendance"       enable row level security;
alter table public."LeaveRequest"     enable row level security;
alter table public."Overtime"         enable row level security;
alter table public."Payroll"          enable row level security;
alter table public."PayrollApproval"  enable row level security;
alter table public."HrRule"           enable row level security;
alter table public."User"             enable row level security;
alter table public."Increment"        enable row level security;
alter table public."Separation"       enable row level security;

-- Sanity check: no table should end up RLS-disabled.
select relname as table_name, relrowsecurity as rls_enabled
from pg_class
join pg_namespace n on n.oid = relnamespace
where n.nspname = 'public' and relkind = 'r'
order by relname;
