import crypto from "node:crypto";

/**
 * Stateless HMAC session tokens for the custom bcrypt auth system.
 * Format: base64url(JSON payload) + "." + base64url(HMAC-SHA256)
 *
 * AUTH_SECRET must be identical across server.js and every Vercel function.
 */

function getSecret() {
  const secret = process.env.AUTH_SECRET;
  if (!secret || secret.length < 32) {
    throw new Error("AUTH_SECRET missing or too short (min 32 chars)");
  }
  return secret;
}

export function createToken(user, ttlMs = 12 * 60 * 60 * 1000) {
  const payload = {
    sub: user.id,
    email: user.email,
    role: String(user.role || "").toLowerCase(),
    status: user.status,
    exp: Date.now() + ttlMs,
  };
  const body = Buffer.from(JSON.stringify(payload)).toString("base64url");
  const sig = crypto
    .createHmac("sha256", getSecret())
    .update(body)
    .digest("base64url");
  return `${body}.${sig}`;
}

export function verifyToken(token) {
  try {
    if (!token || typeof token !== "string") return null;
    const [body, sig] = token.split(".");
    if (!body || !sig) return null;
    const expected = crypto
      .createHmac("sha256", getSecret())
      .update(body)
      .digest("base64url");
    const a = Buffer.from(sig);
    const b = Buffer.from(expected);
    if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) return null;
    const payload = JSON.parse(Buffer.from(body, "base64url").toString());
    if (!payload.exp || Date.now() > payload.exp) return null;
    if (payload.status !== "Active") return null;
    return payload;
  } catch {
    return null;
  }
}

/** Express/Vercel middleware: requires a valid Bearer token.
 *  Returns true and sets req.auth on success; sends 401 and returns false otherwise. */
export function requireAuth(req, res, next) {
  const header = req.headers.authorization || "";
  const token = header.startsWith("Bearer ") ? header.slice(7) : null;
  const payload = verifyToken(token);
  if (!payload) {
    res.status(401).json({
      success: false,
      error: "Authentication required. Please sign in again.",
    });
    return false;
  }
  req.auth = payload;
  if (typeof next === "function") next();
  return true;
}

const WRITE_ROLES = new Set([
  "admin",
  "super admin",
  "super_admin",
  "hr",
  "hr manager",
  "hr_manager",
]);

export function canWrite(role) {
  return WRITE_ROLES.has(String(role || "").toLowerCase());
}
