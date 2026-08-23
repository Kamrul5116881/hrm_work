import express from "express";
import fs from "node:fs";
import path from "node:path";
import { PrismaClient } from "@prisma/client";
import bcrypt from "bcryptjs";
import { createToken, requireAuth, canWrite } from "./api/_lib/auth.js";
import stateHandler from "./api/state.js";
import ledgerHandler from "./api/ledger.js";

/* Minimal .env loader for local dev (Vercel injects env vars natively).
   Must run before the first request that needs AUTH_SECRET. */
try {
  const envPath = path.join(process.cwd(), ".env");
  if (fs.existsSync(envPath)) {
    for (const line of fs.readFileSync(envPath, "utf8").split(/\r?\n/)) {
      const m = line.match(/^([A-Z0-9_]+)=(.*)$/);
      if (m && process.env[m[1]] === undefined) {
        let v = m[2].trim();
        if ((v.startsWith('"') && v.endsWith('"')) || (v.startsWith("'") && v.endsWith("'"))) v = v.slice(1, -1);
        process.env[m[1]] = v;
      }
    }
  }
} catch {}

const app = express();
const prisma = new PrismaClient();

app.use(express.json({ limit: "5mb" }));

app.post("/api/auth", async (req, res) => {
  try {
    const { email, password } = req.body || {};

    if (!email || !password) {
      return res.status(400).json({
        success: false,
        message: "Email and password are required.",
      });
    }

    const loginValue = String(email).trim().toLowerCase();

    const user = await prisma.user.findUnique({
      where: {
        email: loginValue,
      },
    });

    if (!user) {
      return res.status(401).json({
        success: false,
        message: "Invalid username/email or password.",
      });
    }

    const passwordMatch = await bcrypt.compare(
      String(password),
      user.passwordHash
    );

    if (!passwordMatch) {
      return res.status(401).json({
        success: false,
        message: "Invalid username/email or password.",
      });
    }

    if (user.status !== "Active") {
      return res.status(403).json({
        success: false,
        message: "This account is inactive.",
      });
    }

    return res.json({
      success: true,
      message: "Login successful.",
      token: createToken(user),
      user: {
        id: user.id,
        name: user.name,
        email: user.email,
        role: user.role,
        status: user.status,
      },
    });

  } catch (error) {
    console.error("AUTH ERROR:", error);

    return res.status(500).json({
      success: false,
      message: "Authentication server error.",
    });
  }
});

app.get("/api/health", (req, res) => {
  res.json({
    success: true,
    message: "HRM API is running",
  });
});

/* HR state persistence — same handler as the Vercel function,
   protected by session-token auth. Writes are role-gated inside. */
app.get("/api/state", requireAuth, (req, res) => stateHandler(req, res));
app.post("/api/state", requireAuth, (req, res) => stateHandler(req, res));

/* Accounting Ledger (TDS/VDS transactions) */
app.get("/api/ledger", requireAuth, (req, res) => ledgerHandler(req, res));
app.post("/api/ledger", requireAuth, (req, res) => ledgerHandler(req, res));

const PORT = process.env.PORT || 3001;

app.listen(PORT, () => {
  console.log(`HRM API running on http://localhost:${PORT}`);
});