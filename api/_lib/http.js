export function methodGuard(req, res, methods) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,PUT,PATCH,DELETE,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  if (req.method === "OPTIONS") { res.status(204).end(); return false; }
  if (!methods.includes(req.method)) { res.status(405).json({ error: "Method not allowed" }); return false; }
  return true;
}
export function body(req) {
  if (!req.body) return {};
  if (typeof req.body === "object") return req.body;
  try { return JSON.parse(req.body); } catch { return {}; }
}
export function errorMessage(error) { return process.env.NODE_ENV === "production" ? "Request failed" : (error?.message || String(error)); }
