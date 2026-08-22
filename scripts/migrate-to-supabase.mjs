/**
 * Migrate Hastizam HRM data: SOURCE_DATABASE_URL -> TARGET_DATABASE_URL
 *
 * - Preserves all primary keys and foreign keys verbatim (cuid()s are plain text).
 * - Inserts in FK-safe order (parents before children).
 * - Verifies row counts per table afterwards and fails loudly on mismatch.
 *
 * Usage:
 *   SOURCE_DATABASE_URL=... TARGET_DATABASE_URL=... node scripts/migrate-to-supabase.mjs --dry-run
 *   SOURCE_DATABASE_URL=... TARGET_DATABASE_URL=... node scripts/migrate-to-supabase.mjs
 *
 * NEVER points at a database implicitly: both URLs are REQUIRED.
 */
import { PrismaClient } from "@prisma/client";

const args = process.argv.slice(2);
const DRY_RUN = args.includes("--dry-run");

const SOURCE = process.env.SOURCE_DATABASE_URL;
const TARGET = process.env.TARGET_DATABASE_URL;

if (!SOURCE || !TARGET) {
  console.error(
    "ERROR: both SOURCE_DATABASE_URL and TARGET_DATABASE_URL must be set.\n" +
      "Example:\n" +
      '  $env:SOURCE_DATABASE_URL="<neon-url>"; $env:TARGET_DATABASE_URL="<supabase-url>"; node scripts/migrate-to-supabase.mjs --dry-run'
  );
  process.exit(1);
}

if (SOURCE === TARGET) {
  console.error("ERROR: source and target URLs are identical.");
  process.exit(1);
}

const source = new PrismaClient({ datasourceUrl: SOURCE });
const target = new PrismaClient({ datasourceUrl: TARGET });

/* FK-safe order: standalone tables first, employees before their children. */
const MODELS = [
  "hrRule",
  "payrollApproval",
  "user",
  "employee",
  "attendance",
  "leaveRequest",
  "overtime",
  "payroll",
  "increment",
  "separation",
];

function chunk(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

async function main() {
  console.log(`Dry run: ${DRY_RUN ? "YES" : "NO"}\n`);

  // Sanity: source must be reachable and non-empty somewhere.
  const srcCounts = {};
  let totalSource = 0;
  for (const model of MODELS) {
    srcCounts[model] = await source[model].count();
    totalSource += srcCounts[model];
    console.log(`source ${model.padEnd(16)} ${srcCounts[model]}`);
  }
  console.log(`source TOTAL     ${totalSource}\n`);

  if (totalSource === 0) {
    console.error("ERROR: source database has no rows — aborting (wrong URL?).");
    process.exit(1);
  }

  if (DRY_RUN) {
    const tgtCounts = {};
    for (const model of MODELS) tgtCounts[model] = await target[model].count();
    console.log("Dry run complete. Target current counts:");
    for (const model of MODELS) console.log(`target ${model.padEnd(16)} ${tgtCounts[model]}`);
    return;
  }

  const t0 = Date.now();
  await target.$transaction(
    async (tx) => {
      // Idempotent: clear any partial prior attempt, children first.
      for (const model of [...MODELS].reverse()) {
        await tx[model].deleteMany();
      }
      for (const model of MODELS) {
        const rows = await source[model].findMany();
        if (rows.length === 0) continue;
        for (const part of chunk(rows, 400)) {
          await tx[model].createMany({ data: part, skipDuplicates: false });
        }
        console.log(`copied ${model.padEnd(16)} ${rows.length}`);
      }
    },
    { maxWait: 20000, timeout: 300000 }
  );

  // ---- Verification ----
  console.log("\nVerifying row counts...");
  let ok = true;
  for (const model of MODELS) {
    const tgt = await target[model].count();
    const match = tgt === srcCounts[model];
    if (!match) ok = false;
    console.log(`${match ? "OK  " : "FAIL"} ${model.padEnd(16)} src=${srcCounts[model]} tgt=${tgt}`);
  }

  const admin = await target.user.findUnique({
    where: { email: "admin@example.com" },
    select: { email: true, role: true, status: true },
  });
  console.log(`\nAdmin user in target: ${admin ? JSON.stringify(admin) : "MISSING"}`);
  if (!admin) ok = false;

  if (!ok) {
    console.error("\nVERIFICATION FAILED — do NOT cut over. Target may hold partial data.");
    process.exit(1);
  }

  console.log(`\nSUCCESS — copied and verified in ${((Date.now() - t0) / 1000).toFixed(1)}s.`);
}

main()
  .catch((e) => {
    console.error("MIGRATION ERROR:", e);
    process.exitCode = 1;
  })
  .finally(async () => {
    await source.$disconnect();
    await target.$disconnect();
  });
