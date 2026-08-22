# Supabase Migration Runbook — Hastizam HRM

> **STATUS: COMPLETED 2026-08-23.** App is live on Supabase (Sydney, project
> `bswqsduzxdllfpeaodas`) via the `hrm_app` role + session pooler. Neon remains
> as `NEON_DATABASE_URL` backup in `.env`. Remaining optional phase: RLS SQL in
> `supabase/rls-enable.sql` (only needed before exposing supabase-js to clients),
> and updating `DATABASE_URL` in Vercel env vars for production deploys.

Goal: move the Prisma database from Neon to **Supabase Postgres** with zero UI changes,
zero data loss, and a one-line rollback.

Why it's safe: Supabase *is* PostgreSQL. The app talks to the DB only through Prisma,
so swapping `DATABASE_URL` swaps the backend. Auth keeps using bcrypt against the
migrated `User` table — no auth code changes in this phase.

---

## Phase 1 — Prepare (you)

In your `.env` (edit the file yourself — don't paste credentials into chat):

```env
# existing line — leave as-is until cutover:
DATABASE_URL="postgresql://...neon..."

# NEW: Supabase session-pooler connection (Settings → Database → Connection string → Session pooler, port 5432)
TARGET_DATABASE_URL="postgresql://postgres.<project-ref>:<URL-ENCODED-PASSWORD>@aws-0-<region>.pooler.supabase.com:5432/postgres"
```

Notes:
- URL-encode special characters in the password (`@`→`%40`, `#`→`%23`, etc.).
- Use the **session pooler** host/port 5432 — works with IPv4 and supports Prisma's prepared statements.
  (Transaction pooler 6543 is only for serverless functions like `api/*.js` on Vercel.)
- Optional, for later phases only: `SUPABASE_URL`, `SUPABASE_ANON_KEY` (never the service_role key).

## Phase 2 — Schema + data (assistant runs these)

```powershell
# 2.1 Create the same schema on Supabase (non-destructive; fresh project is empty)
$env:DATABASE_URL = $env:TARGET_DATABASE_URL; npx prisma db push

# 2.2 Dry run — counts only, writes nothing
$env:SOURCE_DATABASE_URL = <neon>; $env:TARGET_DATABASE_URL = <supabase>
node scripts/migrate-to-supabase.mjs --dry-run

# 2.3 Real migration (copies all 10 tables in FK-safe order, verifies counts + admin user)
node scripts/migrate-to-supabase.mjs
```

The script fails loudly if any table count mismatches or `admin@example.com` is missing.

## Phase 3 — Cutover (one line)

```env
# .env
NEON_DATABASE_URL="postgresql://...old neon url kept as backup..."
DATABASE_URL="postgresql://...supabase session pooler..."   # TARGET_DATABASE_URL value moves here
```

Smoke tests after restart (`npm run server`, then through Vite :5173):
- `POST /api/auth` correct creds → 200 + sanitized user
- wrong password → 401 · wrong email → 401 · empty body → 400
- Login via UI → dashboard loads with real employee/payroll data

## Phase 4 — Harden (SQL Editor in Supabase dashboard)

Run `supabase/rls-enable.sql`. This enables Row Level Security on all 10 tables so the
anon-facing Supabase REST API can't read anything. Our server connects as table owner
and bypasses RLS, so app behavior is unchanged. Client-facing policies come later,
when/where supabase-js is introduced.

## Rollback

Revert `.env` to the Neon `DATABASE_URL` line and restart. Nothing destructive happens
to either database at any point; Neon stays intact as an archive even after cutover.
