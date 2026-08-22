---
name: supabase
description: Use when working with Supabase in the Hastizam HRM project — project setup, connecting Prisma to Supabase Postgres, migrating auth from bcrypt/Express, Row Level Security policies, storage buckets, or edge functions. Covers Supabase-specific gotchas for this repo's React/Vite frontend, Prisma 6.19 data layer, and Express/Vercel API.
---

# Supabase in Hastizam HRM

## Repo context first

- Data layer today: **Prisma 6.19 → PostgreSQL** via `DATABASE_URL` in `.env` (schema: `prisma/schema.prisma`). Models include Employee, Attendance, Payroll, LeaveRequest, User, and more — all holding **live data**.
- Auth today: `POST /api/auth` (Express `server.js` on :3001; Vercel mirror `api/auth.js`) verifies **bcryptjs** hashes (`bcrypt.compare`, cost 12) against `User.passwordHash`. Frontend stores only a sanitized user in `sessionStorage["hrm_session_user"]`.
- Hard rules that also apply to any Supabase work: **never run destructive SQL** (`drop table`, `truncate`, `migrate reset`), never swap bcrypt hashing for another scheme unless explicitly asked, never expose `passwordHash` or service keys to the client. See `AGENTS.md`.

## Connecting Prisma to Supabase

- The Express server (`server.js`) is a long-running process → use the **session pooler**: host `aws-0-<region>.pooler.supabase.com`, port **5432**. Works over IPv4 and supports Prisma's prepared statements. This is what `DATABASE_URL` should be after cutover.
- The Vercel serverless handlers (`api/*.js`) are ephemeral → if they ever connect directly, use the **transaction pooler** port **6543** with `?pgbouncer=true&connection_limit=1`.
- Schema operations (`prisma db push`, `prisma migrate`) also use session pooler 5432; keep a `DIRECT_URL` var only if a future setup needs a truly direct `db.<ref>.supabase.co` host.
- Prisma `db push` is preferred in this repo (no migration history exists); it is non-destructive by default but still review the SQL it emits before confirming.

## Auth migration rules

- Existing passwords are bcrypt(12) hashes. If moving login to Supabase Auth, hashes can be imported into Supabase (GoTrue supports bcrypt) — do **not** regenerate or reset user passwords as part of a migration.
- Keep `/api/auth` semantics intact while both paths exist: email normalized `trim().toLowerCase()`, inactive users rejected with 403, sanitized user object returned, 401 for bad credentials.
- If using `@supabase/supabase-js` auth from the React app instead, still bridge to the existing session shape (`hrm_session_user`) so `AppShell`'s stage guard keeps working.

## Row Level Security (critical)

- Every new table created on Supabase: `alter table ... enable row level security;` plus explicit policies. A table without RLS is open to any anon key holder.
- Frontend may use ONLY the **anon** key. The **service_role** key bypasses RLS — server-side only (never in `VITE_*` vars, never in git).
- Anything exposed to the Vite frontend needs the `VITE_` prefix; server-only secrets stay unprefixed in `.env`.

## Client usage pattern

```js
// src/services/supabaseClient.js — single shared client
import { createClient } from "@supabase/supabase-js";

export const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
);
```

Install once: `npm i @supabase/supabase-js`. Don't create multiple clients per component.

## Workflow notes

- Prefer Supabase CLI (`supabase link`, `supabase db push`) for schema changes when available; otherwise use Prisma against the direct URL.
- Edge Functions are optional here — the repo already has Vercel serverless handlers in `api/`; don't duplicate endpoints into Edge Functions without being asked.
- Storage: buckets are private by default; access goes through policies like tables.
