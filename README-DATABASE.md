# Hastizam HRM — Database setup

This package keeps the existing React/Vite HRM UI and adds Neon PostgreSQL + Prisma + Vercel API scaffolding.

## Local setup

1. Copy `.env.example` to `.env`.
2. Put your Neon connection string into `DATABASE_URL`.
3. Run:

```bash
npm install
npx prisma generate
npx prisma db push
npm run dev
```

## Vercel

Add `DATABASE_URL` under Project → Settings → Environment Variables. Do not use a `VITE_` prefix.

After deployment, test:

```text
https://YOUR-DOMAIN.vercel.app/api/health
```

A successful response is:

```json
{"ok":true,"database":"connected","service":"hrm-api"}
```

## Included API routes

- `/api/health`
- `/api/employees`
- `/api/attendance`
- `/api/leave`
- `/api/payroll`
- `/api/state`

`/api/state` is a compatibility endpoint matching the existing HRM state shape, so the current UI can be migrated from browser-only storage after the database connection is verified.

## Prisma models

Employee, Attendance, LeaveRequest, Overtime, Payroll, PayrollApproval, HrRule, User, Increment, Separation.

## Important

The API is a database foundation, not the final production security layer. Before real HR data is used, add authentication, role-based authorization, audit logging, CSRF/origin controls where appropriate, and stricter request validation.
