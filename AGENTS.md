# AGENTS.md — Hastizam HRM

## Commands

```bash
npm install            # postinstall runs `prisma generate` automatically
npm run dev            # Vite dev server on :5173
npm run server         # Express backend on :3001 — REQUIRED for any /api call in dev
npm run build          # vite build (no typecheck step; plain JS project)
npm run lint           # eslint
node create-user.js    # upserts admin user (admin@example.com / Admin@123) via bcrypt(12)
```

- No test suite exists. Verify changes with `npm run build` plus manual API calls (POST `/api/auth` correct/wrong creds).
- Backend order matters: start `npm run server` before testing login through Vite; `vite.config.js` proxies `/api/*` → `http://localhost:3001`.

## Architecture

- **The entire frontend lives in one file:** `src/App.jsx` (~6,100 lines) — all components, including both login pages and the Workspace.
- **Two login components exist:** `TopHRLoginPage` is the ACTIVE one rendered by `App`. `LoginPage` is dead/unused but kept functional; do not delete its UI or wire it in without asking.
- **`api/*.js` files are Vercel serverless handlers**, not executed by Vite. Locally, only `server.js` serves endpoints (`/api/auth`, `/api/health`). The other routes (`employees`, `attendance`, `leave`, `payroll`, `state`) exist only when deployed to Vercel.
- `src/services/hrmApi.js` is currently **unused** by `App.jsx`.
- The Workspace persists HR data (employees, attendance, payroll) in **browser localStorage** under its own keys, not through the API yet.

## Authentication rules (do not break)

- Passwords use **bcryptjs**: `create-user.js` hashes with 12 rounds; auth MUST verify with `bcrypt.compare()`. Never swap to SHA-256 or any other scheme.
- `/api/auth` normalizes email (`trim().toLowerCase()`), rejects inactive users (403), and returns a sanitized user only — **never return `passwordHash` to the client**.
- The dashboard transition (`goTo("app")`) may ONLY happen after `/api/auth` responds success. Never call `onLogin()` unconditionally from a click handler.
- Form submit handlers must `e.preventDefault()` — a native form submit reloads the page mid-fetch and causes flaky auth behavior.
- Session persistence stores only the sanitized user object in `sessionStorage` key `hrm_session_user`. Never store credentials client-side.
- Logout lives in ONE place: the **global Workspace top bar** (top-right — visible on every module). `LogoutButton`/`LogoutConfirmDialog` are defined near the bottom of `App.jsx`. Confirming calls `useAuth().logout()` → clears `hrm_session_user`, shows a toast, and App drops back to stage `"login"`. A derived `activeStage` guard in `AppShell` returns to login whenever `user` is null while `stage === "app"`. Do NOT add a second logout inside the HRApp sidebar.
- NEVER wipe all of localStorage/sessionStorage on logout: localStorage holds live HR business data (employees, attendance, payroll). Only auth keys may be cleared (`clearSessionData()` in App.jsx).
- Responsive rules for the HR sidebar live in HRApp's `<style>` block; global top-bar/toast rules live in `GLOBAL_LAYOUT_CSS` (near `STAGE_TRANSITION_CSS`). On ≤900px the HR nav becomes a horizontal scroller — do not reintroduce `.hrm-sidebar > div:last-child { display:none }`, it hides the whole nav.

## Database safety

- **Live database is now Supabase Postgres** (`DATABASE_URL` → session pooler `aws-0-ap-southeast-2.pooler.supabase.com:5432`, project ref `bswqsduzxdllfpeaodas`, migrated 2026-08-23). `NEON_DATABASE_URL` is the untouched pre-migration backup — rollback = swap the two lines.
- The app connects as the dedicated **`hrm_app`** DB role (owns all tables, created via Management API). The `postgres` role password was NOT changed and is unknown locally — do not attempt to connect as `postgres`. Role management happens through the Supabase Management API (`/v1/projects/{ref}/database/query`) using the CLI's access token from Windows Credential Manager ("Supabase CLI:supabase").
- **Never run `prisma migrate reset`, drop tables, or destructive commands** on either database.
- Schema changes: `npx prisma db push` (no migration files in this project).
- Keep Prisma pinned at `6.19.0`; don't bump React/Vite/Prisma unless there's a genuine compatibility issue.
- Migration tooling kept for reference: `scripts/migrate-to-supabase.mjs` (needs explicit SOURCE/TARGET URLs), runbook `SUPABASE_MIGRATION.md`. RLS hardening SQL (`supabase/rls-enable.sql`) is prepared but NOT yet applied — apply only when client-side supabase-js access becomes real.

## Gotchas

- `README-DATABASE.md` is partially stale: `.env.example` does not exist, and its documented `/api/health` response shape doesn't match what `server.js` actually returns (`{"success":true,"message":"HRM API is running"}`). Trust the code over that README.
- `vercel.json` is intentionally `{}`; the `api/` directory deploys as-is on Vercel.
- `npm run build` emits a pre-existing >500 kB chunk-size warning — harmless, don't chase it.

## Owner constraints

- The existing UI (login page design, dashboard, colors, layout, animations) is considered locked. When fixing auth or backend wiring, change logic only — preserve all JSX/design exactly.
