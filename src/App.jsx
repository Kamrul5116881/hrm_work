import React, { useState, useEffect, useMemo, useCallback, useRef, createContext, useContext } from "react";
import {
  LayoutDashboard,
  BookOpen,
  ListFilter,
  FileBarChart,
  Plus,
  Search,
  Pencil,
  Trash2,
  X,
  Download,
  Printer,
  ChevronDown,
  ArrowUpDown,
  CircleDollarSign,
  Receipt,
  Landmark,
  Wallet,
  AlertCircle,
  Upload,
  ArrowRight,
  LogIn,
  LogOut,
  Users,
  ClipboardList,
  CheckCircle2,
  Lock,
  Unlock,
  UserCircle,
  CalendarCheck,
  Umbrella,
  History,
  Settings as SettingsIcon,
  ArrowLeft,
  AlertTriangle,
  Scissors,
  Ruler,
  Package,
  TrendingUp,
  TrendingDown,
  ChevronRight,
  UsersRound,
  UserRound,
  LockKeyhole,
  Eye,
  EyeOff,
} from "lucide-react";
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  PieChart,
  Pie,
  Cell,
  LineChart,
  Line,
} from "recharts";
import * as XLSX from "xlsx";

/* ---------------------------------------------------------------------
   DESIGN TOKENS — "Ledger" system
   Paper ivory ground, deep ledger-green ink, brick-red rule (the classic
   double-entry cashbook margin line) as the one signature accent.
--------------------------------------------------------------------- */
const T = {
  paper: "#F6F2E9",
  paperDeep: "#EFE8D8",
  ink: "#1B2A29",
  inkSoft: "#4B5D5A",
  muted: "#8A8272",
  line: "#D9D0B8",
  card: "#FFFEFB",
  accent: "#0F6B4C",
  accentDeep: "#0B4F38",
  rule: "#B3472B",
  ruleSoft: "#E7CFC5",
  gold: "#B08D2B",
};

const FONT_CSS = `
@import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:opsz,wght@8..60,400;8..60,600;8..60,700&family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
`;

const PAYMENT_TYPES = ["Advance", "Payment", "Final Payment"];
const CONDITIONS = [
  "Including",
  "Excluding",
  "TDS Including VDS Excluding",
  "N/A",
];
const MONTH_NAMES = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];

const emptyForm = {
  id: null,
  date: "",
  cheque: "",
  chequeDate: "",
  vendor: "",
  paymentType: "Payment",
  particular: "",
  vendorAddress: "",
  tin: "",
  bin: "",
  mainGL: "",
  subGL: "",
  sectionRef: "",
  invoiceAmount: "",
  condition: "Excluding",
  tdsRate: "",
  vdsRate: "",
  remarks: "",
};

/* ---------------------------------------------------------------------
   CALCULATION ENGINE — mirrors Work!R:AA exactly
--------------------------------------------------------------------- */
function calc(rec) {
  const N = parseFloat(rec.invoiceAmount) || 0;
  const P = parseFloat(rec.tdsRate) || 0; // fraction e.g. 0.075
  const Q = parseFloat(rec.vdsRate) || 0;
  const O = rec.condition;

  const R = O === "Including" ? null : 0; // placeholder, computed below w/ V
  let tdsIncluding = 0,
    tdsExcluding = 0,
    tdsInclVdsExcl = 0;
  let vdsIncluding = 0,
    vdsExcluding = 0,
    vdsInclVdsExcl = 0;

  // Excel evaluates V (VDS-Including) independent of R, and R depends on V —
  // both only fire in the "Including" branch, so compute together.
  if (O === "Including") {
    vdsIncluding = (N / (1 + Q)) * Q;
    tdsIncluding = (N - vdsIncluding) * P;
  }
  if (O === "Excluding") {
    tdsExcluding = (N / (1 - P)) * P;
    vdsExcluding = (N + tdsExcluding) * Q;
  }
  if (O === "TDS Including VDS Excluding") {
    tdsInclVdsExcl = N * P;
    vdsInclVdsExcl = N * Q;
  }

  const tdsTotal = tdsIncluding + tdsExcluding + tdsInclVdsExcl;
  const vdsTotal = vdsIncluding + vdsExcluding + vdsInclVdsExcl;

  let netPayment = 0;
  if (O === "Excluding") netPayment = N;
  else if (O === "TDS Including VDS Excluding") netPayment = N - tdsInclVdsExcl;
  else if (O === "Including") netPayment = N - tdsIncluding - vdsIncluding;
  else netPayment = N; // N/A — no withholding

  return {
    invoiceAmount: N,
    tdsAmount: round2(tdsTotal),
    vdsAmount: round2(vdsTotal),
    netPayment: round2(netPayment),
  };
}
function round2(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}
function money(n) {
  const v = Number(n) || 0;
  return v.toLocaleString("en-BD", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}
function fmtDate(d) {
  if (!d) return "—";
  const dt = new Date(d);
  if (isNaN(dt)) return d;
  return dt.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
}
function monthKey(d) {
  const dt = new Date(d);
  if (isNaN(dt)) return "Unknown";
  return dt.toLocaleDateString("en-US", { month: "short", year: "numeric" });
}

/* ---------------------------------------------------------------------
   EXCEL IMPORT — reads a workbook laid out like the real Work sheet
   (A Sl | B Date | C Cheque | D Cheque Date | E Vendor | F Payment Type |
    G Particular | H Vendor Address | I TIN | J BIN | K Main GL | L Sub-GL |
    M Section Ref | N Invoice Amount | O Condition | P TDS Rate | Q VDS Rate
    ... AD Remarks). Rows are detected by having a vendor name AND an
    invoice amount — header/blank rows are skipped automatically.
--------------------------------------------------------------------- */
/* ---------------------------------------------------------------------
   EXCEL IMPORT — reads a workbook and maps columns by their HEADER TEXT
   first (so it works with the real Work sheet's headers — "Name of
   Vendor", "TDS VDS Condition ", "TDS Rate", etc. — AND with files this
   app itself exports later), falling back to the original Work-sheet
   column positions only if no header row can be confidently detected.
--------------------------------------------------------------------- */
function excelDateToStr(v) {
  if (v instanceof Date && !isNaN(v)) {
    const y = v.getFullYear(),
      m = String(v.getMonth() + 1).padStart(2, "0"),
      d = String(v.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  if (typeof v === "string" && v.trim()) return v.trim();
  return "";
}

// TDS/VDS rates can come out of Excel as a true fraction (0.075), a bare
// percent number (7.5), or text ("7.5%"). All three should mean 7.5%.
function normalizeRate(v) {
  if (v === "" || v === null || v === undefined) return "";
  if (typeof v === "string") {
    const trimmed = v.trim();
    if (!trimmed) return "";
    const isPercentText = trimmed.endsWith("%");
    const n = parseFloat(trimmed);
    if (isNaN(n)) return "";
    if (isPercentText) return n / 100;
    return n > 1 ? n / 100 : n;
  }
  if (typeof v === "number") {
    return v > 1 ? v / 100 : v;
  }
  return "";
}

// Match imported condition text against the exact strings calc() checks
// with strict equality — trims stray whitespace ("Excluding " → "Excluding")
// and is case-insensitive, so a real-world sloppy export still calculates.
function normalizeCondition(v) {
  const s = (v ?? "").toString().trim();
  if (!s) return "N/A";
  const found = CONDITIONS.find((c) => c.trim().toLowerCase() === s.toLowerCase());
  return found || s;
}

// Header text → field name. Matched case-insensitively with whitespace
// collapsed, so "Payment Type ", "payment  type", etc. all match.
const FIELD_HEADER_ALIASES = {
  date: ["date"],
  cheque: ["cheque", "cheque no", "cheque no.", "cheque number"],
  chequeDate: ["cheque date"],
  vendor: ["name of vendor", "vendor", "vendor name"],
  paymentType: ["payment type"],
  particular: ["particular", "particulars"],
  vendorAddress: ["vendor address"],
  tin: ["tin"],
  bin: ["bin"],
  mainGL: ["main gl"],
  subGL: ["sub-gl", "sub gl"],
  sectionRef: ["section ref"],
  invoiceAmount: ["invoice amount", "invoice"],
  condition: ["tds vds condition", "tds/vds condition", "condition"],
  tdsRate: ["tds rate"],
  vdsRate: ["vds rate"],
  remarks: ["remarks", "remarks 1"],
};
function normHeader(s) {
  return (s ?? "").toString().trim().toLowerCase().replace(/\s+/g, " ");
}
// Scans the first few rows for the one that looks most like a header row
// (matches the most known field names) — handles the real Work sheet's
// two-row header (a title row, then the real headers) without hardcoding
// "row 2".
function detectHeaderMap(rows) {
  let best = { rowIndex: -1, score: 0, map: {} };
  const scanRows = Math.min(rows.length, 6);
  for (let r = 0; r < scanRows; r++) {
    const row = rows[r] || [];
    const map = {};
    let score = 0;
    row.forEach((cell, c) => {
      const h = normHeader(cell);
      if (!h) return;
      for (const [field, aliases] of Object.entries(FIELD_HEADER_ALIASES)) {
        if (map[field] !== undefined) continue; // first matching column wins
        if (aliases.includes(h)) {
          map[field] = c;
          score++;
          break;
        }
      }
    });
    if (score > best.score) best = { rowIndex: r, score, map };
  }
  // Require several confident matches before trusting header-based mapping;
  // otherwise fall back to the known positional layout.
  return best.score >= 4 ? best : null;
}

async function parseExcelFile(file, existingRecords) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const sheetName = wb.SheetNames.includes("Work") ? "Work" : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

  // Positional fallback — the original Work sheet's exact column order.
  const fallbackMap = {
    date: 1, cheque: 2, chequeDate: 3, vendor: 4, paymentType: 5, particular: 6,
    vendorAddress: 7, tin: 8, bin: 9, mainGL: 10, subGL: 11, sectionRef: 12,
    invoiceAmount: 13, condition: 14, tdsRate: 15, vdsRate: 16, remarks: 29,
  };
  const detected = detectHeaderMap(rows);
  const colMap = detected ? { ...fallbackMap, ...detected.map } : fallbackMap;
  const dataStart = detected ? detected.rowIndex + 1 : 0;
  const cell = (row, field) => {
    const idx = colMap[field];
    return idx === undefined ? "" : row[idx];
  };

  let sl = existingRecords.reduce((m, r) => Math.max(m, r.sl || 0), 0);
  const out = [];
  rows.slice(dataStart).forEach((row, i) => {
    const vendor = (cell(row, "vendor") || "").toString().trim();
    const invoice = parseFloat(cell(row, "invoiceAmount"));
    if (!vendor || !invoice) return; // skip header/blank/total rows
    sl += 1;
    out.push({
      id: `imp-${Date.now()}-${i}`,
      sl,
      date: excelDateToStr(cell(row, "date")),
      cheque: (cell(row, "cheque") || "").toString(),
      chequeDate: excelDateToStr(cell(row, "chequeDate")),
      vendor,
      paymentType: (cell(row, "paymentType") || "Payment").toString().trim() || "Payment",
      particular: (cell(row, "particular") || "").toString(),
      vendorAddress: (cell(row, "vendorAddress") || "").toString(),
      tin: (cell(row, "tin") || "").toString(),
      bin: (cell(row, "bin") || "").toString(),
      mainGL: (cell(row, "mainGL") || "").toString(),
      subGL: (cell(row, "subGL") || "").toString(),
      sectionRef: (cell(row, "sectionRef") || "").toString(),
      invoiceAmount: invoice,
      condition: normalizeCondition(cell(row, "condition")),
      tdsRate: normalizeRate(cell(row, "tdsRate")),
      vdsRate: normalizeRate(cell(row, "vdsRate")),
      remarks: (cell(row, "remarks") || "").toString(),
    });
  });
  return out;
}

/* ---------------------------------------------------------------------
   SEED DATA — the 29 real rows read from the Work sheet
--------------------------------------------------------------------- */
const SEED = [
  [
    "2026-07-01",
    "2210003",
    "2026-07-01",
    "Rangs Industries Ltd",
    "Payment",
    "For the purpose of TV purchase for the admin block",
    "8768",
    "Excluding",
    0.05,
    "",
  ],
  [
    "2026-07-01",
    "2210004",
    "2026-07-01",
    "AB Power Engineering Ltd.",
    "Payment",
    "",
    "470980",
    "Excluding",
    "",
    "",
  ],
  [
    "2026-07-02",
    "2201489",
    "2026-07-02",
    "GTCBL",
    "Payment",
    "Professional fees for the month of Jun'26",
    "687960",
    "Excluding",
    0.075,
    0.15,
  ],
  [
    "2026-07-02",
    "2201491",
    "2026-07-02",
    "M/S Hansa",
    "Payment",
    "Consultancy fee",
    "419230",
    "Excluding",
    0.075,
    0.15,
  ],
  [
    "2026-07-02",
    "2201492",
    "2026-07-02",
    "M/S Hansa",
    "Payment",
    "Consultancy fee",
    "68265",
    "Excluding",
    0.075,
    0.15,
  ],
  [
    "2026-07-02",
    "2201493",
    "2026-07-02",
    "Nazrul Islam",
    "Payment",
    "rent",
    "455625",
    "Excluding",
    0.1,
    0.15,
  ],
  [
    "2026-07-05",
    "2201494",
    "2026-07-05",
    "Unity Enterprise",
    "Payment",
    "Dining Table supply",
    "81000",
    "Excluding",
    0.05,
    "",
  ],
  [
    "2026-07-05",
    "2201495",
    "2026-07-05",
    "Masud & Brothers",
    "Payment",
    "Paid mirazul islam for pureit supply",
    "32835",
    "Excluding",
    0.005,
    "",
  ],
  [
    "2026-07-06",
    "2201496",
    "2026-07-06",
    "Unity Enterprise",
    "Payment",
    "Dining Table supply",
    "27750",
    "Excluding",
    0.05,
    "",
  ],
  [
    "2026-07-06",
    "2201500",
    "2026-07-08",
    "F. Rahman Construction",
    "Payment",
    "Supply of fixture and fittings",
    "70000",
    "Excluding",
    0.05,
    0.15,
  ],
  [
    "2026-07-12",
    "2210007",
    "2026-07-12",
    "Bhai Bhai Fire Fighting Company",
    "Payment",
    "Supply of safety materials",
    "55650",
    "Excluding",
    0.05,
    "",
  ],
  [
    "2026-07-12",
    "2210008",
    "2026-07-12",
    "Azad Ad",
    "Payment",
    "Supply of safety materials",
    "52280",
    "Excluding",
    0.05,
    0.15,
  ],
  [
    "2026-07-13",
    "2210009",
    "2026-07-13",
    "Vai Vai Enterprise",
    "Payment",
    "Airline installation for air compressor",
    "12745",
    "Excluding",
    0.05,
    0.15,
  ],
  [
    "2026-07-13",
    "2210010",
    "2026-07-13",
    "F. Rahman Construction",
    "Payment",
    "Civil work payment",
    "43000",
    "Excluding",
    0.05,
    0.15,
  ],
  [
    "2026-07-19",
    "2210012",
    "2026-07-19",
    "Unity Enterprise",
    "Payment",
    "Supply of stationary Items, Hydraulic oil, PVC sign board etc",
    "397836",
    "Excluding",
    0.05,
    "",
  ],
  [
    "2026-07-19",
    "2210013",
    "2026-07-19",
    "Binary Kraft",
    "Payment",
    "Admin Block – Roller Blinds and Glass sticker work – Completed",
    "106000",
    "Excluding",
    0.05,
    0.15,
  ],
  [
    "2026-07-19",
    "2210015",
    "2026-07-19",
    "AB Power Engineering Ltd.",
    "Payment",
    "Remaining supply of electrical materials and labour charges",
    "291136",
    "Excluding",
    0.05,
    "",
  ],
  [
    "2026-07-22",
    "2210017",
    "2026-07-22",
    "AB Power Engineering Ltd.",
    "Payment",
    "Electric materials supply",
    "91675",
    "Excluding",
    0.05,
    "",
  ],
  [
    "2026-07-23",
    "2210018",
    "2026-07-23",
    "Fess Trade International",
    "Payment",
    "Floor cleaning machine purchase",
    "83200",
    "Excluding",
    0.05,
    0.15,
  ],
  [
    "2026-07-23",
    "2210019",
    "2026-07-23",
    "M/S Hansa",
    "Payment",
    "Leather Purchase",
    "150652",
    "Excluding",
    0.05,
    0.15,
  ],
  [
    "2026-07-23",
    "2210020",
    "2026-07-23",
    "M/S Leather Sewing and Accessories",
    "Payment",
    "machineries spare parts supplied",
    "16000",
    "Excluding",
    0.05,
    0.15,
  ],
  [
    "2026-07-23",
    "2210021",
    "2026-07-23",
    "F. Rahman Construction",
    "Payment",
    "manpower hired during buyer visit",
    "10000",
    "Excluding",
    0.1,
    0.15,
  ],
  [
    "2026-07-23",
    "2210022",
    "2026-07-23",
    "M/S Ajad Ad",
    "Payment",
    "Safety sign board purchased.",
    "7960",
    "Excluding",
    0.05,
    0.15,
  ],
  [
    "2026-07-23",
    "2210023",
    "2026-07-23",
    "K.A.R. Associates",
    "Payment",
    "Dehumidifier purchased for Hall-A materials stores",
    "143750",
    "Excluding",
    0.05,
    0.15,
  ],
  [
    "2026-07-26",
    "",
    "",
    "Sikdar insurance PLC",
    "Payment",
    "Bill paid for marine insurance",
    "129528",
    "",
    "",
    "",
  ],
  [
    "2026-07-27",
    "2210025",
    "2026-07-27",
    "Transplace logistic Ltd.",
    "Payment",
    "Bill paid for Shipping agent",
    "488585",
    "Excluding",
    0.01,
    "",
  ],
  [
    "2026-07-28",
    "2210027",
    "",
    "Sentry Security Services Ltd",
    "Payment",
    "Security Service Bill for the month of June'26",
    "84096",
    "Excluding",
    0.02,
    "",
  ],
].map((r, i) => ({
  id: `seed-${i}`,
  sl: i + 1,
  date: r[0],
  cheque: r[1],
  chequeDate: r[2],
  vendor: r[3],
  paymentType: r[4],
  particular: r[5],
  vendorAddress: "",
  tin: "",
  bin: "",
  mainGL: "",
  subGL: "",
  sectionRef: "",
  invoiceAmount: r[6],
  condition: r[7] || "N/A",
  tdsRate: r[8] || "",
  vdsRate: r[9] || "",
  remarks: "",
}));

/* ---------------------------------------------------------------------
   STORAGE — writes to BOTH Claude's artifact storage (when available)
   AND the browser's localStorage on every save, and reads back from
   whichever one actually has data. This redundancy means a save
   survives even if one of the two mechanisms is unavailable or
   silently failing in a given environment (Claude preview vs. a local
   dev server vs. a browser with storage restrictions).
--------------------------------------------------------------------- */
const STORAGE_KEY = "ledger:transactions:v1";

function hasArtifactStorage() {
  return (
    typeof window !== "undefined" &&
    window.storage &&
    typeof window.storage.get === "function" &&
    typeof window.storage.set === "function"
  );
}
function hasLocalStorage() {
  try {
    const k = "__ledger_ls_probe__";
    window.localStorage.setItem(k, "1");
    window.localStorage.removeItem(k);
    return true;
  } catch (e) {
    return false;
  }
}

async function loadFromStorage(key) {
  let fromArtifact = null;
  if (hasArtifactStorage()) {
    try {
      const res = await window.storage.get(key);
      if (res && res.value) fromArtifact = JSON.parse(res.value);
    } catch (e) {
      /* key not present yet in artifact storage — not an error */
    }
  }
  let fromLocal = null;
  if (hasLocalStorage()) {
    try {
      const v = window.localStorage.getItem(key);
      if (v) fromLocal = JSON.parse(v);
    } catch (e) {
      console.error("localStorage read failed", e);
    }
  }
  if (fromArtifact !== null) return fromArtifact;
  if (fromLocal !== null) return fromLocal;
  return null;
}

async function saveToStorage(key, data) {
  const payload = JSON.stringify(data);
  let savedSomewhere = false;
  if (hasArtifactStorage()) {
    try {
      await window.storage.set(key, payload, false);
      savedSomewhere = true;
    } catch (e) {
      console.error("artifact storage save failed:", e);
    }
  }
  if (hasLocalStorage()) {
    try {
      window.localStorage.setItem(key, payload);
      savedSomewhere = true;
    } catch (e) {
      console.error("localStorage save failed:", e);
    }
  }
  return savedSomewhere;
}

/* Ledger records load/save now live in the Supabase-backed STORAGE section below. */

/* ---------------------------------------------------------------------
   SHARED UI BITS
--------------------------------------------------------------------- */
function Field({ label, children, span }) {
  return (
    <label
      style={{
        display: "flex",
        flexDirection: "column",
        gap: 5,
        gridColumn: span ? `span ${span}` : undefined,
      }}
    >
      <span
        style={{
          fontFamily: "'IBM Plex Sans'",
          fontSize: 11.5,
          fontWeight: 600,
          letterSpacing: ".04em",
          textTransform: "uppercase",
          color: T.inkSoft,
        }}
      >
        {label}
      </span>
      {children}
    </label>
  );
}
const inputStyle = {
  fontFamily: "'IBM Plex Mono', monospace",
  fontSize: 14,
  color: T.ink,
  background: T.card,
  border: `1px solid ${T.line}`,
  borderRadius: 9,
  padding: "10px 12px",
  outline: "none",
  width: "100%",
  boxSizing: "border-box",
};
function TInput(props) {
  return <input {...props} style={{ ...inputStyle, ...(props.style || {}) }} />;
}
function TSelect({ children, ...props }) {
  return (
    <div style={{ position: "relative" }}>
      <select
        {...props}
        style={{
          ...inputStyle,
          appearance: "none",
          paddingRight: 30,
          cursor: "pointer",
        }}
      >
        {children}
      </select>
      <ChevronDown
        size={14}
        color={T.muted}
        style={{
          position: "absolute",
          right: 10,
          top: 12,
          pointerEvents: "none",
        }}
      />
    </div>
  );
}
function StatCard({ icon: Icon, label, value, sub, accent }) {
  return (
    <div className="hr-stat-card"
      style={{
        background: T.card,
        border: `1px solid ${T.line}`,
        borderRadius: 10,
        padding: "18px 20px",
        display: "flex",
        flexDirection: "column",
        gap: 10,
        position: "relative",
        overflow: "hidden",
      }}
    >
      <div
        style={{
          position: "absolute",
          left: 0,
          top: 0,
          bottom: 0,
          width: 3,
          background: accent || T.accent,
        }}
      />
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 8,
          color: T.inkSoft,
        }}
      >
        <Icon size={15} strokeWidth={2} />
        <span
          style={{
            fontFamily: "'IBM Plex Sans'",
            fontSize: 12,
            fontWeight: 600,
            letterSpacing: ".04em",
            textTransform: "uppercase",
          }}
        >
          {label}
        </span>
      </div>
      <div
        style={{
          fontFamily: "'Source Serif 4'",
          fontSize: 26,
          fontWeight: 600,
          color: T.ink,
        }}
      >
        {value}
      </div>
      {sub && (
        <div
          style={{
            fontFamily: "'IBM Plex Sans'",
            fontSize: 12,
            color: T.muted,
          }}
        >
          {sub}
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------------
   MASTER DATA / ENTRY FORM
--------------------------------------------------------------------- */
function EntryForm({ records, initial, onSave, onCancel }) {
  const [form, setForm] = useState(initial || emptyForm);
  const isEdit = !!(initial && initial.id);
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));
  const nextSl = useMemo(() => {
    if (isEdit) return initial.sl;
    return records.reduce((m, r) => Math.max(m, r.sl || 0), 0) + 1;
  }, [records, isEdit, initial]);

  const preview = calc(form);
  const vendors = useMemo(
    () =>
      Array.from(new Set(records.map((r) => r.vendor).filter(Boolean))).sort(),
    [records],
  );
  const [vendorOpen, setVendorOpen] = useState(false);

  function submit(e) {
    e.preventDefault();
    if (!form.date || !form.vendor || !form.invoiceAmount) return;
    onSave({ ...form, sl: nextSl, id: isEdit ? form.id : `tx-${Date.now()}` });
  }

  return (
    <form
      onSubmit={submit}
      style={{
        background: T.card,
        border: `1px solid ${T.line}`,
        borderRadius: 12,
        padding: 26,
      }}
    >
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "baseline",
          marginBottom: 20,
        }}
      >
        <div>
          <div
            style={{
              fontFamily: "'IBM Plex Sans'",
              fontSize: 11,
              fontWeight: 600,
              color: T.rule,
              letterSpacing: ".08em",
              textTransform: "uppercase",
            }}
          >
            Sl. No. {String(nextSl).padStart(4, "0")}
          </div>
          <h2
            style={{
              fontFamily: "'Source Serif 4'",
              fontSize: 22,
              margin: "2px 0 0",
              color: T.ink,
            }}
          >
            {isEdit ? "Edit entry" : "New ledger entry"}
          </h2>
        </div>
        {onCancel && (
          <button
            type="button"
            onClick={onCancel}
            style={{
              background: "none",
              border: "none",
              cursor: "pointer",
              color: T.muted,
            }}
          >
            <X size={20} />
          </button>
        )}
      </div>

      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(4, 1fr)",
          gap: 16,
        }}
      >
        <Field label="Date">
          <TInput
            type="date"
            value={form.date}
            onChange={set("date")}
            required
          />
        </Field>
        <Field label="Cheque No.">
          <TInput
            value={form.cheque}
            onChange={set("cheque")}
            placeholder="2210005"
          />
        </Field>
        <Field label="Cheque Date">
          <TInput
            type="date"
            value={form.chequeDate}
            onChange={set("chequeDate")}
          />
        </Field>
        <Field label="Payment Type">
          <TSelect value={form.paymentType} onChange={set("paymentType")}>
            {PAYMENT_TYPES.map((p) => (
              <option key={p} value={p}>
                {p}
              </option>
            ))}
          </TSelect>
        </Field>

        <Field label="Vendor Name" span={2}>
          <div style={{ position: "relative" }}>
            <TInput
              value={form.vendor}
              onChange={(e) => {
                set("vendor")(e);
                setVendorOpen(true);
              }}
              onFocus={() => setVendorOpen(true)}
              onBlur={() => setTimeout(() => setVendorOpen(false), 120)}
              placeholder="Search or type a vendor…"
              required
              autoComplete="off"
            />
            {vendorOpen && form.vendor && (
              <div
                style={{
                  position: "absolute",
                  top: "104%",
                  left: 0,
                  right: 0,
                  background: T.card,
                  border: `1px solid ${T.line}`,
                  borderRadius: 6,
                  maxHeight: 160,
                  overflowY: "auto",
                  zIndex: 5,
                  boxShadow: "0 6px 16px rgba(27,42,41,.12)",
                }}
              >
                {vendors
                  .filter(
                    (v) =>
                      v.toLowerCase().includes(form.vendor.toLowerCase()) &&
                      v !== form.vendor,
                  )
                  .slice(0, 6)
                  .map((v) => (
                    <div
                      key={v}
                      onMouseDown={() => setForm((f) => ({ ...f, vendor: v }))}
                      style={{
                        padding: "8px 12px",
                        fontFamily: "'IBM Plex Sans'",
                        fontSize: 13.5,
                        cursor: "pointer",
                        color: T.ink,
                      }}
                      onMouseEnter={(e) =>
                        (e.currentTarget.style.background = T.paperDeep)
                      }
                      onMouseLeave={(e) =>
                        (e.currentTarget.style.background = "transparent")
                      }
                    >
                      {v}
                    </div>
                  ))}
              </div>
            )}
          </div>
        </Field>
        <Field label="Particular" span={2}>
          <TInput
            value={form.particular}
            onChange={set("particular")}
            placeholder="Nature of the payment"
          />
        </Field>

        <Field label="Vendor Address" span={2}>
          <TInput value={form.vendorAddress} onChange={set("vendorAddress")} />
        </Field>
        <Field label="TIN">
          <TInput value={form.tin} onChange={set("tin")} />
        </Field>
        <Field label="BIN">
          <TInput value={form.bin} onChange={set("bin")} />
        </Field>

        <Field label="Main GL">
          <TInput value={form.mainGL} onChange={set("mainGL")} />
        </Field>
        <Field label="Sub-GL">
          <TInput value={form.subGL} onChange={set("subGL")} />
        </Field>
        <Field label="Section Ref">
          <TInput value={form.sectionRef} onChange={set("sectionRef")} />
        </Field>
        <Field label="Invoice Amount">
          <TInput
            type="number"
            step="0.01"
            value={form.invoiceAmount}
            onChange={set("invoiceAmount")}
            placeholder="500000"
            required
          />
        </Field>

        <Field label="TDS / VDS Condition">
          <TSelect value={form.condition} onChange={set("condition")}>
            {CONDITIONS.map((c) => (
              <option key={c} value={c}>
                {c}
              </option>
            ))}
          </TSelect>
        </Field>
        <Field label="TDS Rate">
          <TInput
            type="number"
            step="0.001"
            min="0"
            max="1"
            value={form.tdsRate}
            onChange={set("tdsRate")}
            placeholder="0.075 = 7.5%"
          />
        </Field>
        <Field label="VDS Rate">
          <TInput
            type="number"
            step="0.001"
            min="0"
            max="1"
            value={form.vdsRate}
            onChange={set("vdsRate")}
            placeholder="0.15 = 15%"
          />
        </Field>
        <Field label="Remarks">
          <TInput value={form.remarks} onChange={set("remarks")} />
        </Field>
      </div>

      {/* live calculation preview */}
      <div
        style={{
          marginTop: 22,
          borderTop: `1px dashed ${T.line}`,
          paddingTop: 18,
        }}
      >
        <div
          style={{
            fontFamily: "'IBM Plex Sans'",
            fontSize: 11,
            fontWeight: 600,
            color: T.inkSoft,
            letterSpacing: ".05em",
            textTransform: "uppercase",
            marginBottom: 10,
          }}
        >
          Calculated automatically
        </div>
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "repeat(4, 1fr)",
            gap: 14,
          }}
        >
          {[
            ["Invoice Amount", preview.invoiceAmount, T.ink],
            ["TDS Amount", preview.tdsAmount, T.rule],
            ["VDS Amount", preview.vdsAmount, T.rule],
            ["Net Payment", preview.netPayment, T.accent],
          ].map(([lbl, val, color]) => (
            <div
              key={lbl}
              style={{
                background: T.paperDeep,
                borderRadius: 8,
                padding: "12px 14px",
              }}
            >
              <div
                style={{
                  fontFamily: "'IBM Plex Sans'",
                  fontSize: 11,
                  color: T.muted,
                }}
              >
                {lbl}
              </div>
              <div
                style={{
                  fontFamily: "'IBM Plex Mono'",
                  fontSize: 17,
                  fontWeight: 600,
                  color,
                }}
              >
                ৳ {money(val)}
              </div>
            </div>
          ))}
        </div>
      </div>

      <div
        style={{
          display: "flex",
          gap: 10,
          marginTop: 22,
          justifyContent: "flex-end",
        }}
      >
        {onCancel && (
          <button
            type="button"
            onClick={onCancel}
            style={{
              padding: "10px 18px",
              borderRadius: 7,
              border: `1px solid ${T.line}`,
              background: T.card,
              color: T.inkSoft,
              fontFamily: "'IBM Plex Sans'",
              fontSize: 13.5,
              fontWeight: 600,
              cursor: "pointer",
            }}
          >
            Cancel
          </button>
        )}
        <button
          type="submit"
          style={{
            padding: "10px 22px",
            borderRadius: 7,
            border: "none",
            background: T.accent,
            color: "#fff",
            fontFamily: "'IBM Plex Sans'",
            fontSize: 13.5,
            fontWeight: 600,
            cursor: "pointer",
            display: "flex",
            alignItems: "center",
            gap: 8,
          }}
        >
          <Plus size={15} /> {isEdit ? "Save changes" : "Save entry"}
        </button>
      </div>
    </form>
  );
}

/* ---------------------------------------------------------------------
   TRANSACTIONS TABLE
--------------------------------------------------------------------- */
function Transactions({ records, onEdit, onDelete, onImport }) {
  const fileInputRef = React.useRef(null);
  const [importBusy, setImportBusy] = useState(false);
  const [q, setQ] = useState("");
  const [vendorF, setVendorF] = useState("");
  const [typeF, setTypeF] = useState("");
  const [condF, setCondF] = useState("");
  const [sort, setSort] = useState({ key: "sl", dir: "desc" });

  const vendors = useMemo(
    () =>
      Array.from(new Set(records.map((r) => r.vendor).filter(Boolean))).sort(),
    [records],
  );

  const filtered = useMemo(() => {
    let rows = records.map((r) => ({ ...r, ...calc(r) }));
    if (q) {
      const s = q.toLowerCase();
      rows = rows.filter((r) =>
        [r.vendor, r.particular, r.cheque, String(r.sl)].some((v) =>
          (v || "").toString().toLowerCase().includes(s),
        ),
      );
    }
    if (vendorF) rows = rows.filter((r) => r.vendor === vendorF);
    if (typeF) rows = rows.filter((r) => r.paymentType === typeF);
    if (condF) rows = rows.filter((r) => r.condition === condF);
    rows.sort((a, b) => {
      const { key, dir } = sort;
      let av = a[key],
        bv = b[key];
      if (key === "date") {
        av = new Date(av || 0).getTime();
        bv = new Date(bv || 0).getTime();
      }
      if (typeof av === "string") {
        av = (av || "").toLowerCase();
        bv = (bv || "").toLowerCase();
      }
      if (av < bv) return dir === "asc" ? -1 : 1;
      if (av > bv) return dir === "asc" ? 1 : -1;
      return 0;
    });
    return rows;
  }, [records, q, vendorF, typeF, condF, sort]);

  const totals = useMemo(
    () =>
      filtered.reduce(
        (a, r) => ({
          invoice: a.invoice + r.invoiceAmount,
          tds: a.tds + r.tdsAmount,
          vds: a.vds + r.vdsAmount,
          net: a.net + r.netPayment,
        }),
        { invoice: 0, tds: 0, vds: 0, net: 0 },
      ),
    [filtered],
  );

  function toggleSort(key) {
    setSort((s) =>
      s.key === key
        ? { key, dir: s.dir === "asc" ? "desc" : "asc" }
        : { key, dir: "asc" },
    );
  }
  async function handleFileChosen(e) {
    const file = e.target.files && e.target.files[0];
    e.target.value = ""; // allow re-selecting the same file later
    if (!file) return;
    setImportBusy(true);
    try {
      const imported = await parseExcelFile(file, records);
      if (imported.length === 0) {
        window.alert(
          "No rows with both a vendor name and an invoice amount were found — check that the sheet matches the Work layout (Vendor in column E, Invoice Amount in column N).",
        );
        return;
      }
      const proceed = window.confirm(
        `Import ${imported.length} row(s) from "${file.name}"?\n\nThis will replace all data currently shown — the old data will not come back unless you import it again.`,
      );
      if (!proceed) return;
      onImport(imported);
    } catch (err) {
      window.alert(
        "Couldn't read that file — make sure it's a .xlsx or .xls export of your ledger.",
      );
      console.error(err);
    } finally {
      setImportBusy(false);
    }
  }
  function exportExcel() {
    const data = filtered.map((r) => ({
      Sl: r.sl,
      Date: r.date,
      Cheque: r.cheque,
      Vendor: r.vendor,
      "Payment Type": r.paymentType,
      Particular: r.particular,
      "Invoice Amount": r.invoiceAmount,
      Condition: r.condition,
      "TDS Rate": r.tdsRate,
      "VDS Rate": r.vdsRate,
      "TDS Amount": r.tdsAmount,
      "VDS Amount": r.vdsAmount,
      "Net Payment": r.netPayment,
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Transactions");
    XLSX.writeFile(wb, "transactions.xlsx");
  }

  const Th = ({ label, k, align }) => (
    <th
      onClick={() => toggleSort(k)}
      style={{
        textAlign: align || "left",
        padding: "10px 12px",
        cursor: "pointer",
        userSelect: "none",
        fontFamily: "'IBM Plex Sans'",
        fontSize: 11,
        fontWeight: 600,
        letterSpacing: ".04em",
        textTransform: "uppercase",
        color: T.inkSoft,
        whiteSpace: "nowrap",
      }}
    >
      <span style={{ display: "inline-flex", alignItems: "center", gap: 4 }}>
        {label} {sort.key === k && <ArrowUpDown size={11} />}
      </span>
    </th>
  );

  return (
    <div>
      <div
        style={{
          display: "flex",
          flexWrap: "wrap",
          gap: 10,
          marginBottom: 16,
          alignItems: "center",
        }}
      >
        <div style={{ position: "relative", flex: "1 1 220px" }}>
          <Search
            size={14}
            color={T.muted}
            style={{ position: "absolute", left: 11, top: 11 }}
          />
          <TInput
            placeholder="Search vendor, particular, cheque, Sl…"
            value={q}
            onChange={(e) => setQ(e.target.value)}
            style={{ paddingLeft: 32 }}
          />
        </div>
        <TSelect
          value={vendorF}
          onChange={(e) => setVendorF(e.target.value)}
          style={{ maxWidth: 190 }}
        >
          <option value="">All vendors</option>
          {vendors.map((v) => (
            <option key={v} value={v}>
              {v}
            </option>
          ))}
        </TSelect>
        <TSelect
          value={typeF}
          onChange={(e) => setTypeF(e.target.value)}
          style={{ maxWidth: 170 }}
        >
          <option value="">All payment types</option>
          {PAYMENT_TYPES.map((v) => (
            <option key={v} value={v}>
              {v}
            </option>
          ))}
        </TSelect>
        <TSelect
          value={condF}
          onChange={(e) => setCondF(e.target.value)}
          style={{ maxWidth: 220 }}
        >
          <option value="">All conditions</option>
          {CONDITIONS.map((v) => (
            <option key={v} value={v}>
              {v}
            </option>
          ))}
        </TSelect>
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFileChosen}
          style={{ display: "none" }}
        />
        <button
          onClick={() => fileInputRef.current && fileInputRef.current.click()}
          disabled={importBusy}
          style={{
            display: "flex",
            alignItems: "center",
            gap: 6,
            padding: "9px 14px",
            borderRadius: 7,
            border: `1px solid ${T.line}`,
            background: T.card,
            cursor: importBusy ? "wait" : "pointer",
            fontFamily: "'IBM Plex Sans'",
            fontSize: 13,
            fontWeight: 600,
            color: T.inkSoft,
          }}
        >
          <Upload size={14} /> {importBusy ? "Reading…" : "Import Excel"}
        </button>
        <button
          onClick={exportExcel}
          style={{
            display: "flex",
            alignItems: "center",
            gap: 6,
            padding: "9px 14px",
            borderRadius: 7,
            border: `1px solid ${T.line}`,
            background: T.card,
            cursor: "pointer",
            fontFamily: "'IBM Plex Sans'",
            fontSize: 13,
            fontWeight: 600,
            color: T.inkSoft,
          }}
        >
          <Download size={14} /> Export
        </button>
        <button
          onClick={() => window.print()}
          style={{
            display: "flex",
            alignItems: "center",
            gap: 6,
            padding: "9px 14px",
            borderRadius: 7,
            border: `1px solid ${T.line}`,
            background: T.card,
            cursor: "pointer",
            fontFamily: "'IBM Plex Sans'",
            fontSize: 13,
            fontWeight: 600,
            color: T.inkSoft,
          }}
        >
          <Printer size={14} /> Print
        </button>
      </div>

      <div
        style={{
          background: T.card,
          border: `1px solid ${T.line}`,
          borderRadius: 12,
          overflow: "hidden",
        }}
      >
        <div style={{ overflowX: "auto" }}>
          <table
            style={{
              width: "100%",
              borderCollapse: "collapse",
              minWidth: 1100,
            }}
          >
            <thead>
              <tr style={{ borderBottom: `2px solid ${T.rule}` }}>
                <Th label="Sl" k="sl" />
                <Th label="Date" k="date" />
                <Th label="Cheque" k="cheque" />
                <Th label="Vendor" k="vendor" />
                <Th label="Type" k="paymentType" />
                <Th label="Particular" k="particular" />
                <Th label="Invoice" k="invoiceAmount" align="right" />
                <Th label="Condition" k="condition" />
                <Th label="TDS" k="tdsAmount" align="right" />
                <Th label="VDS" k="vdsAmount" align="right" />
                <Th label="Net Payment" k="netPayment" align="right" />
                <th style={{ padding: "10px 12px" }} />
              </tr>
            </thead>
            <tbody>
              {filtered.map((r, i) => (
                <tr
                  key={r.id}
                  style={{
                    borderBottom: `1px solid ${T.line}`,
                    background: i % 2 ? T.paper : T.card,
                  }}
                >
                  <td
                    style={{
                      padding: "9px 12px",
                      fontFamily: "'IBM Plex Mono'",
                      fontSize: 12.5,
                      color: T.muted,
                      borderLeft: `2px solid ${T.rule}`,
                    }}
                  >
                    {String(r.sl).padStart(3, "0")}
                  </td>
                  <td
                    style={{
                      padding: "9px 12px",
                      fontFamily: "'IBM Plex Sans'",
                      fontSize: 13,
                    }}
                  >
                    {fmtDate(r.date)}
                  </td>
                  <td
                    style={{
                      padding: "9px 12px",
                      fontFamily: "'IBM Plex Mono'",
                      fontSize: 13,
                    }}
                  >
                    {r.cheque || "—"}
                  </td>
                  <td
                    style={{
                      padding: "9px 12px",
                      fontFamily: "'IBM Plex Sans'",
                      fontSize: 13,
                      fontWeight: 500,
                    }}
                  >
                    {r.vendor}
                  </td>
                  <td style={{ padding: "9px 12px" }}>
                    <span
                      style={{
                        fontFamily: "'IBM Plex Sans'",
                        fontSize: 11,
                        fontWeight: 600,
                        padding: "3px 8px",
                        borderRadius: 20,
                        background: T.paperDeep,
                        color: T.inkSoft,
                      }}
                    >
                      {r.paymentType}
                    </span>
                  </td>
                  <td
                    style={{
                      padding: "9px 12px",
                      fontFamily: "'IBM Plex Sans'",
                      fontSize: 12.5,
                      color: T.inkSoft,
                      maxWidth: 220,
                      overflow: "hidden",
                      textOverflow: "ellipsis",
                      whiteSpace: "nowrap",
                    }}
                    title={r.particular}
                  >
                    {r.particular || "—"}
                  </td>
                  <td
                    style={{
                      padding: "9px 12px",
                      fontFamily: "'IBM Plex Mono'",
                      fontSize: 13,
                      textAlign: "right",
                    }}
                  >
                    {money(r.invoiceAmount)}
                  </td>
                  <td
                    style={{
                      padding: "9px 12px",
                      fontFamily: "'IBM Plex Sans'",
                      fontSize: 12,
                    }}
                  >
                    {r.condition}
                  </td>
                  <td
                    style={{
                      padding: "9px 12px",
                      fontFamily: "'IBM Plex Mono'",
                      fontSize: 13,
                      textAlign: "right",
                      color: T.rule,
                    }}
                  >
                    {money(r.tdsAmount)}
                  </td>
                  <td
                    style={{
                      padding: "9px 12px",
                      fontFamily: "'IBM Plex Mono'",
                      fontSize: 13,
                      textAlign: "right",
                      color: T.rule,
                    }}
                  >
                    {money(r.vdsAmount)}
                  </td>
                  <td
                    style={{
                      padding: "9px 12px",
                      fontFamily: "'IBM Plex Mono'",
                      fontSize: 13.5,
                      textAlign: "right",
                      fontWeight: 600,
                      color: T.accentDeep,
                    }}
                  >
                    {money(r.netPayment)}
                  </td>
                  <td style={{ padding: "9px 12px" }}>
                    <div style={{ display: "flex", gap: 6 }}>
                      <button
                        onClick={() => onEdit(r)}
                        style={{
                          background: "none",
                          border: "none",
                          cursor: "pointer",
                          color: T.muted,
                        }}
                      >
                        <Pencil size={14} />
                      </button>
                      <button
                        onClick={() => onDelete(r.id)}
                        style={{
                          background: "none",
                          border: "none",
                          cursor: "pointer",
                          color: T.rule,
                        }}
                      >
                        <Trash2 size={14} />
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
              {filtered.length === 0 && (
                <tr>
                  <td
                    colSpan={12}
                    style={{
                      padding: 40,
                      textAlign: "center",
                      fontFamily: "'IBM Plex Sans'",
                      color: T.muted,
                      fontSize: 13,
                    }}
                  >
                    {records.length === 0
                      ? "No entries yet. Add one from Master Data, or import your Excel file."
                      : "No entries match these filters."}
                  </td>
                </tr>
              )}
            </tbody>
            {filtered.length > 0 && (
              <tfoot>
                <tr
                  style={{
                    borderTop: `2px solid ${T.rule}`,
                    background: T.paperDeep,
                  }}
                >
                  <td
                    colSpan={6}
                    style={{
                      padding: "10px 12px",
                      fontFamily: "'IBM Plex Sans'",
                      fontSize: 12,
                      fontWeight: 700,
                      color: T.ink,
                    }}
                  >
                    Total ({filtered.length} entries)
                  </td>
                  <td
                    style={{
                      padding: "10px 12px",
                      fontFamily: "'IBM Plex Mono'",
                      fontSize: 13,
                      textAlign: "right",
                      fontWeight: 700,
                    }}
                  >
                    {money(totals.invoice)}
                  </td>
                  <td />
                  <td
                    style={{
                      padding: "10px 12px",
                      fontFamily: "'IBM Plex Mono'",
                      fontSize: 13,
                      textAlign: "right",
                      fontWeight: 700,
                      color: T.rule,
                    }}
                  >
                    {money(totals.tds)}
                  </td>
                  <td
                    style={{
                      padding: "10px 12px",
                      fontFamily: "'IBM Plex Mono'",
                      fontSize: 13,
                      textAlign: "right",
                      fontWeight: 700,
                      color: T.rule,
                    }}
                  >
                    {money(totals.vds)}
                  </td>
                  <td
                    style={{
                      padding: "10px 12px",
                      fontFamily: "'IBM Plex Mono'",
                      fontSize: 13.5,
                      textAlign: "right",
                      fontWeight: 700,
                      color: T.accentDeep,
                    }}
                  >
                    {money(totals.net)}
                  </td>
                  <td />
                </tr>
              </tfoot>
            )}
          </table>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------
   DASHBOARD
--------------------------------------------------------------------- */
function Dashboard({ records }) {
  const allRows = useMemo(
    () => records.map((r) => ({ ...r, ...calc(r) })),
    [records],
  );

  const years = useMemo(() => {
    const s = new Set();
    allRows.forEach((r) => {
      const d = new Date(r.date);
      if (!isNaN(d)) s.add(d.getFullYear());
    });
    return Array.from(s).sort((a, b) => b - a);
  }, [allRows]);

  const [year, setYear] = useState("All");
  const [month, setMonth] = useState("All");

  const rows = useMemo(() => {
    if (year === "All" && month === "All") return allRows;
    return allRows.filter((r) => {
      const d = new Date(r.date);
      if (isNaN(d)) return false; // undated rows only show under All/All
      if (year !== "All" && d.getFullYear() !== Number(year)) return false;
      if (month !== "All" && d.getMonth() !== Number(month)) return false;
      return true;
    });
  }, [allRows, year, month]);

  const totals = rows.reduce(
    (a, r) => ({
      invoice: a.invoice + r.invoiceAmount,
      tds: a.tds + r.tdsAmount,
      vds: a.vds + r.vdsAmount,
      net: a.net + r.netPayment,
    }),
    { invoice: 0, tds: 0, vds: 0, net: 0 },
  );

  const byMonth = useMemo(() => {
    const m = {};
    rows.forEach((r) => {
      const k = monthKey(r.date);
      m[k] = (m[k] || 0) + r.netPayment;
    });
    return Object.entries(m).map(([month, net]) => ({ month, net }));
  }, [rows]);

  const byType = useMemo(() => {
    const m = {};
    rows.forEach((r) => {
      m[r.paymentType] = (m[r.paymentType] || 0) + r.netPayment;
    });
    return Object.entries(m).map(([name, value]) => ({ name, value }));
  }, [rows]);

  const byVendor = useMemo(() => {
    const m = {};
    rows.forEach((r) => {
      m[r.vendor] = (m[r.vendor] || 0) + r.netPayment;
    });
    return Object.entries(m)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 6);
  }, [rows]);

  const pieColors = [T.accent, T.gold, T.rule];

  return (
    <div>
      <div
        className="no-print"
        style={{
          display: "flex",
          alignItems: "center",
          gap: 10,
          marginBottom: 16,
          flexWrap: "wrap",
        }}
      >
        <span
          style={{
            fontFamily: "'IBM Plex Sans'",
            fontSize: 12.5,
            fontWeight: 600,
            color: T.inkSoft,
          }}
        >
          Period
        </span>
        <TSelect value={year} onChange={(e) => setYear(e.target.value)} style={{ maxWidth: 130 }}>
          <option value="All">All Years</option>
          {years.map((y) => (
            <option key={y} value={y}>{y}</option>
          ))}
        </TSelect>
        <TSelect value={month} onChange={(e) => setMonth(e.target.value)} style={{ maxWidth: 160 }}>
          <option value="All">All Months</option>
          {MONTH_NAMES.map((m, i) => (
            <option key={m} value={i}>{m}</option>
          ))}
        </TSelect>
        {(year !== "All" || month !== "All") && (
          <button
            onClick={() => { setYear("All"); setMonth("All"); }}
            style={{
              padding: "8px 14px", borderRadius: 20, border: `1px solid ${T.line}`,
              background: T.card, color: T.inkSoft, fontFamily: "'IBM Plex Sans'",
              fontSize: 12.5, fontWeight: 600, cursor: "pointer",
            }}
          >
            Clear filter
          </button>
        )}
      </div>

      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(4, 1fr)",
          gap: 14,
          marginBottom: 22,
        }}
      >
        <StatCard
          icon={Receipt}
          label="Total Invoice Amount"
          value={`৳ ${money(totals.invoice)}`}
          sub={`${rows.length} transactions`}
        />
        <StatCard
          icon={Landmark}
          label="Total TDS"
          value={`৳ ${money(totals.tds)}`}
          accent={T.rule}
          sub="Withheld at source"
        />
        <StatCard
          icon={CircleDollarSign}
          label="Total VDS"
          value={`৳ ${money(totals.vds)}`}
          accent={T.rule}
          sub="Value Deducted at Source"
        />
        <StatCard
          icon={Wallet}
          label="Total Payment"
          value={`৳ ${money(totals.net)}`}
          accent={T.accent}
          sub="Net disbursed to vendors"
        />
      </div>

      <div
        style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr", gap: 16 }}
      >
        <div
          style={{
            background: T.card,
            border: `1px solid ${T.line}`,
            borderRadius: 12,
            padding: 20,
          }}
        >
          <div
            style={{
              fontFamily: "'Source Serif 4'",
              fontSize: 16,
              fontWeight: 600,
              marginBottom: 12,
              color: T.ink,
            }}
          >
            Net payment by month
          </div>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={byMonth}>
              <CartesianGrid stroke={T.line} vertical={false} />
              <XAxis
                dataKey="month"
                tick={{
                  fontFamily: "IBM Plex Sans",
                  fontSize: 11,
                  fill: T.inkSoft,
                }}
                axisLine={{ stroke: T.line }}
                tickLine={false}
              />
              <YAxis
                tick={{
                  fontFamily: "IBM Plex Mono",
                  fontSize: 10,
                  fill: T.muted,
                }}
                axisLine={false}
                tickLine={false}
                width={70}
                tickFormatter={(v) => `৳${(v / 1000).toFixed(0)}k`}
              />
              <Tooltip
                formatter={(v) => `৳ ${money(v)}`}
                contentStyle={{
                  fontFamily: "IBM Plex Sans",
                  fontSize: 12,
                  border: `1px solid ${T.line}`,
                  borderRadius: 8,
                }}
              />
              <Bar dataKey="net" fill={T.accent} radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>

        <div
          style={{
            background: T.card,
            border: `1px solid ${T.line}`,
            borderRadius: 12,
            padding: 20,
          }}
        >
          <div
            style={{
              fontFamily: "'Source Serif 4'",
              fontSize: 16,
              fontWeight: 600,
              marginBottom: 12,
              color: T.ink,
            }}
          >
            By payment type
          </div>
          <ResponsiveContainer width="100%" height={220}>
            <PieChart>
              <Pie
                data={byType}
                dataKey="value"
                nameKey="name"
                innerRadius={50}
                outerRadius={80}
                paddingAngle={2}
              >
                {byType.map((_, i) => (
                  <Cell key={i} fill={pieColors[i % pieColors.length]} />
                ))}
              </Pie>
              <Tooltip
                formatter={(v) => `৳ ${money(v)}`}
                contentStyle={{
                  fontFamily: "IBM Plex Sans",
                  fontSize: 12,
                  border: `1px solid ${T.line}`,
                  borderRadius: 8,
                }}
              />
            </PieChart>
          </ResponsiveContainer>
          <div
            style={{
              display: "flex",
              flexDirection: "column",
              gap: 6,
              marginTop: 6,
            }}
          >
            {byType.map((t, i) => (
              <div
                key={t.name}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 8,
                  fontFamily: "'IBM Plex Sans'",
                  fontSize: 12.5,
                  color: T.inkSoft,
                }}
              >
                <span
                  style={{
                    width: 9,
                    height: 9,
                    borderRadius: 9,
                    background: pieColors[i % pieColors.length],
                  }}
                />
                {t.name} — ৳ {money(t.value)}
              </div>
            ))}
          </div>
        </div>
      </div>

      <div
        style={{
          background: T.card,
          border: `1px solid ${T.line}`,
          borderRadius: 12,
          padding: 20,
          marginTop: 16,
        }}
      >
        <div
          style={{
            fontFamily: "'Source Serif 4'",
            fontSize: 16,
            fontWeight: 600,
            marginBottom: 14,
            color: T.ink,
          }}
        >
          Top vendors by net payment
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {byVendor.map(([name, val]) => {
            const pct = totals.net ? (val / totals.net) * 100 : 0;
            return (
              <div key={name}>
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    fontFamily: "'IBM Plex Sans'",
                    fontSize: 12.5,
                    marginBottom: 4,
                  }}
                >
                  <span style={{ color: T.ink, fontWeight: 500 }}>{name}</span>
                  <span
                    style={{ fontFamily: "'IBM Plex Mono'", color: T.inkSoft }}
                  >
                    ৳ {money(val)}
                  </span>
                </div>
                <div
                  style={{
                    height: 6,
                    borderRadius: 4,
                    background: T.paperDeep,
                    overflow: "hidden",
                  }}
                >
                  <div
                    style={{
                      height: "100%",
                      width: `${pct}%`,
                      background: T.accent,
                      borderRadius: 4,
                    }}
                  />
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------
   REPORTS
--------------------------------------------------------------------- */
function Reports({ records }) {
  const [groupBy, setGroupBy] = useState("vendor");
  const rows = useMemo(
    () => records.map((r) => ({ ...r, ...calc(r) })),
    [records],
  );

  const grouped = useMemo(() => {
    const m = {};
    rows.forEach((r) => {
      const key =
        groupBy === "vendor"
          ? r.vendor
          : groupBy === "paymentType"
            ? r.paymentType
            : groupBy === "month"
              ? monthKey(r.date)
              : r.condition;
      if (!m[key])
        m[key] = { key, count: 0, invoice: 0, tds: 0, vds: 0, net: 0 };
      m[key].count++;
      m[key].invoice += r.invoiceAmount;
      m[key].tds += r.tdsAmount;
      m[key].vds += r.vdsAmount;
      m[key].net += r.netPayment;
    });
    return Object.values(m).sort((a, b) => b.net - a.net);
  }, [rows, groupBy]);

  function exportExcel() {
    const ws = XLSX.utils.json_to_sheet(
      grouped.map((g) => ({
        Group: g.key,
        Entries: g.count,
        "Invoice Amount": round2(g.invoice),
        TDS: round2(g.tds),
        VDS: round2(g.vds),
        "Net Payment": round2(g.net),
      })),
    );
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Report");
    XLSX.writeFile(wb, `report-by-${groupBy}.xlsx`);
  }

  const labels = {
    vendor: "Vendor",
    paymentType: "Payment Type",
    month: "Month",
    condition: "TDS/VDS Condition",
  };

  return (
    <div>
      <div
        style={{
          display: "flex",
          gap: 10,
          marginBottom: 16,
          alignItems: "center",
          flexWrap: "wrap",
        }}
      >
        <span
          style={{
            fontFamily: "'IBM Plex Sans'",
            fontSize: 13,
            color: T.inkSoft,
            fontWeight: 600,
          }}
        >
          Group by
        </span>
        {Object.entries(labels).map(([k, l]) => (
          <button
            key={k}
            onClick={() => setGroupBy(k)}
            style={{
              padding: "7px 14px",
              borderRadius: 20,
              border: `1px solid ${groupBy === k ? T.accent : T.line}`,
              background: groupBy === k ? T.accent : T.card,
              color: groupBy === k ? "#fff" : T.inkSoft,
              fontFamily: "'IBM Plex Sans'",
              fontSize: 12.5,
              fontWeight: 600,
              cursor: "pointer",
            }}
          >
            {l}
          </button>
        ))}
        <div style={{ flex: 1 }} />
        <button
          onClick={exportExcel}
          style={{
            display: "flex",
            alignItems: "center",
            gap: 6,
            padding: "9px 14px",
            borderRadius: 7,
            border: `1px solid ${T.line}`,
            background: T.card,
            cursor: "pointer",
            fontFamily: "'IBM Plex Sans'",
            fontSize: 13,
            fontWeight: 600,
            color: T.inkSoft,
          }}
        >
          <Download size={14} /> Export Excel
        </button>
        <button
          onClick={() => window.print()}
          style={{
            display: "flex",
            alignItems: "center",
            gap: 6,
            padding: "9px 14px",
            borderRadius: 7,
            border: `1px solid ${T.line}`,
            background: T.card,
            cursor: "pointer",
            fontFamily: "'IBM Plex Sans'",
            fontSize: 13,
            fontWeight: 600,
            color: T.inkSoft,
          }}
        >
          <Printer size={14} /> Print
        </button>
      </div>

      <div
        style={{
          background: T.card,
          border: `1px solid ${T.line}`,
          borderRadius: 12,
          overflow: "hidden",
        }}
      >
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr style={{ borderBottom: `2px solid ${T.rule}` }}>
              <th
                style={{
                  textAlign: "left",
                  padding: "10px 14px",
                  fontFamily: "'IBM Plex Sans'",
                  fontSize: 11,
                  fontWeight: 600,
                  textTransform: "uppercase",
                  color: T.inkSoft,
                }}
              >
                {labels[groupBy]}
              </th>
              <th
                style={{
                  textAlign: "right",
                  padding: "10px 14px",
                  fontFamily: "'IBM Plex Sans'",
                  fontSize: 11,
                  fontWeight: 600,
                  textTransform: "uppercase",
                  color: T.inkSoft,
                }}
              >
                Entries
              </th>
              <th
                style={{
                  textAlign: "right",
                  padding: "10px 14px",
                  fontFamily: "'IBM Plex Sans'",
                  fontSize: 11,
                  fontWeight: 600,
                  textTransform: "uppercase",
                  color: T.inkSoft,
                }}
              >
                Invoice
              </th>
              <th
                style={{
                  textAlign: "right",
                  padding: "10px 14px",
                  fontFamily: "'IBM Plex Sans'",
                  fontSize: 11,
                  fontWeight: 600,
                  textTransform: "uppercase",
                  color: T.inkSoft,
                }}
              >
                TDS
              </th>
              <th
                style={{
                  textAlign: "right",
                  padding: "10px 14px",
                  fontFamily: "'IBM Plex Sans'",
                  fontSize: 11,
                  fontWeight: 600,
                  textTransform: "uppercase",
                  color: T.inkSoft,
                }}
              >
                VDS
              </th>
              <th
                style={{
                  textAlign: "right",
                  padding: "10px 14px",
                  fontFamily: "'IBM Plex Sans'",
                  fontSize: 11,
                  fontWeight: 600,
                  textTransform: "uppercase",
                  color: T.inkSoft,
                }}
              >
                Net Payment
              </th>
            </tr>
          </thead>
          <tbody>
            {grouped.map((g, i) => (
              <tr
                key={g.key}
                style={{
                  borderBottom: `1px solid ${T.line}`,
                  background: i % 2 ? T.paper : T.card,
                }}
              >
                <td
                  style={{
                    padding: "10px 14px",
                    fontFamily: "'IBM Plex Sans'",
                    fontSize: 13,
                    fontWeight: 500,
                    borderLeft: `2px solid ${T.rule}`,
                  }}
                >
                  {g.key}
                </td>
                <td
                  style={{
                    padding: "10px 14px",
                    textAlign: "right",
                    fontFamily: "'IBM Plex Mono'",
                    fontSize: 13,
                    color: T.muted,
                  }}
                >
                  {g.count}
                </td>
                <td
                  style={{
                    padding: "10px 14px",
                    textAlign: "right",
                    fontFamily: "'IBM Plex Mono'",
                    fontSize: 13,
                  }}
                >
                  {money(g.invoice)}
                </td>
                <td
                  style={{
                    padding: "10px 14px",
                    textAlign: "right",
                    fontFamily: "'IBM Plex Mono'",
                    fontSize: 13,
                    color: T.rule,
                  }}
                >
                  {money(g.tds)}
                </td>
                <td
                  style={{
                    padding: "10px 14px",
                    textAlign: "right",
                    fontFamily: "'IBM Plex Mono'",
                    fontSize: 13,
                    color: T.rule,
                  }}
                >
                  {money(g.vds)}
                </td>
                <td
                  style={{
                    padding: "10px 14px",
                    textAlign: "right",
                    fontFamily: "'IBM Plex Mono'",
                    fontSize: 13.5,
                    fontWeight: 600,
                    color: T.accentDeep,
                  }}
                >
                  {money(g.net)}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}


/* =========================================================================
   HR MANAGEMENT MODULE
   Full Employee Data / Attendance / Leave / Payroll / Payslip / Probation
   system, merged in as-is (see hr_renamed_body.jsx for the rename log —
   only its internal T/FONT_CSS/Field/TInput/TSelect/StatCard/money/
   fmtDate/Dashboard/Reports/App identifiers were renamed with an HR_/Hr
   prefix so they don't collide with the ledger's own copies of those
   names; all logic, styling, and behavior are untouched).
========================================================================= */


/* =========================================================================
   DESIGN TOKENS — "Docket" system
   A garments-floor work-docket aesthetic: denim indigo as structure,
   bobbin-thread amber as the one live accent, unbleached-canvas ground.
   Signature motif: a stitched / perforated edge (like a cutting ticket
   torn from a pad) on the primary record cards.
========================================================================= */
const HR_T = {
  canvas: "#F1EDE4",
  canvasDeep: "#E7E0D0",
  panel: "#FFFFFF",
  ink: "#1E2433",
  inkSoft: "#5B6478",
  muted: "#8D8B7F",
  line: "#DAD3BF",
  indigo: "#2C3E70",
  indigoDeep: "#1E2C52",
  amber: "#C8842C",
  amberSoft: "#F3DDB6",
  good: "#3C7A5D",
  bad: "#B4472F",
  chip: "#EDE6D3",
};

const HR_FONT_CSS = `
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
`;

const STATUSES = ["New Join", "Regular", "Resigned", "Terminated", "Suspended", "On Leave", "Inactive"];
const ACTIVE_STATUSES = ["New Join", "Regular", "On Leave"];
const MARITAL = ["Single", "Married", "Widowed", "Divorced"];
const HR_MONTH_NAMES = ["January","February","March","April","May","June","July","August","September","October","November","December"];

/* =========================================================================
   PAYROLL CALCULATION ENGINE
   Reproduces, formula-for-formula, the logic found in the user's real
   "Salary, July 2026.xlsx" reference sheet. Every constant below (31-day
   proration, 30-day absence divisor, /104 OT rate, the fixed-allowance
   subtraction in the Basic formula) is taken directly from that sheet's
   cell formulas, not invented. Divisors live in `rules` so HR/Admin can
   retune them later without touching this function's shape.
========================================================================= */
const DEFAULT_RULES = {
  payDaysDivisor: 31,     // Pay Salary = Gross / payDaysDivisor * Payable Days (present+weekend+leave)
  absentDaysDivisor: 30,  // Absent Amount = Basic / absentDaysDivisor * Absent Days
  otDivisor: 104,         // OT Rate = Basic / otDivisor
  basicDivisor: 1.5,      // Reserved (legacy formula: Basic = (Gross - fixed) / 1.5) — payroll now uses Basic/House Rent exactly as entered in Employee Data
};

function computePayroll(emp, att, rules) {
  const present = num(att.present);
  const weekend = num(att.weekend);
  const leave = num(att.leave);
  const absent = num(att.absent);
  const otHours = num(att.otHours);
  const advance = num(att.advance);
  const arrear = num(att.arrear);
  const tds = num(att.tds);

  // Salary components are taken exactly as entered in Employee Data.
  // If a monthly attendance row overrides a component (e.g. gross, medical),
  // that override is used; otherwise the employee master value is used.
  // Basic and House Rent are NEVER recomputed from gross — they are the
  // verbatim amounts from the employee record.
  const basic = num(att.basic ?? emp.basic);
  const houseRent = num(att.houseRent ?? emp.houseRent);
  const medical = num(att.medical ?? emp.medical);
  const conveyance = num(att.conveyance ?? emp.conveyance);
  const food = num(att.food ?? emp.food);
  const gross = num(att.gross ?? emp.gross);

  const totalDays = present + weekend + leave + absent;
  const payableDays = present + weekend + leave;

  // Pay Salary is pro-rated on PAYABLE days (present + weekend + leave) — absent days are not paid at gross rate.
  // Net is then paySalary - absent deduction (basic/30*absent) + OT, so absent is penalised once at basic rate.
  const paySalary = (gross / rules.payDaysDivisor) * payableDays;
  const absentAmount = (basic / rules.absentDaysDivisor) * absent;
  const otRate = basic / rules.otDivisor;
  const otAmount = otRate * otHours;
  const actualAmount = paySalary - absentAmount + otAmount;
  const payBeforeTds = actualAmount - advance + arrear;
  const payAmount = payBeforeTds - tds;

  return {
    basic: r2(basic), houseRent: r2(houseRent), medical: r2(medical),
    conveyance: r2(conveyance), food: r2(food), gross: r2(gross),
    totalDays, payableDays, present, weekend, leave, absent,
    paySalary: r2(paySalary), absentAmount: r2(absentAmount),
    otRate: r2(otRate), otHours, otAmount: r2(otAmount),
    actualAmount: r2(actualAmount), advance, arrear,
    payBeforeTds: r2(payBeforeTds), tds,
    payAmount: r2(payAmount),
  };
}
function num(v) { const n = parseFloat(v); return isNaN(n) ? 0 : n; }
function r2(n) { return Math.round((n + Number.EPSILON) * 100) / 100; }
function hrMoney(n) {
  const v = Number(n) || 0;
  return v.toLocaleString("en-BD", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function hrFmtDate(d) {
  if (!d) return "—";
  const dt = new Date(d);
  if (isNaN(dt)) return d;
  return dt.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
}
function calcAge(dob) {
  if (!dob) return "";
  const d = new Date(dob);
  if (isNaN(d)) return "";
  const diff = Date.now() - d.getTime();
  return Math.floor(diff / (365.25 * 24 * 3600 * 1000));
}
function monthLabel(key) {
  if (!key) return "";
  const [y, m] = key.split("-");
  return `${HR_MONTH_NAMES[Number(m) - 1]} ${y}`;
}
function currentMonthKey() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

/* =========================================================================
   EXCEL HEADER MAPPING — matches by header text (case/whitespace
   insensitive) so column order in the user's real files never breaks
   import, per their explicit requirement.
========================================================================= */
function hrNormHeader(s) {
  return (s ?? "")
    .toString()
    .trim()
    .toLowerCase()
    .replace(/[\u00a0]/g, " ")
    .replace(/[\/\\_-]+/g, " ")
    .replace(/[().,:]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function hrDetectHeaderMap(rows, aliasTable, minScore) {
  let best = { rowIndex: -1, score: 0, map: {} };
  const scanRows = Math.min(rows.length, 6);
  for (let r = 0; r < scanRows; r++) {
    const row = rows[r] || [];
    const map = {};
    let score = 0;
    row.forEach((cell, c) => {
      const h = hrNormHeader(cell);
      if (!h) return;
      for (const [field, aliases] of Object.entries(aliasTable)) {
        if (map[field] !== undefined) continue;
        const normalizedAliases = aliases.map(hrNormHeader);
        if (normalizedAliases.includes(h)) { map[field] = c; score++; break; }
      }
    });
    if (score > best.score) best = { rowIndex: r, score, map };
  }
  return best.score >= minScore ? best : null;
}
function hrExcelDateToStr(v) {
  if (v instanceof Date && !isNaN(v)) {
    const y = v.getFullYear(), m = String(v.getMonth() + 1).padStart(2, "0"), d = String(v.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  if (typeof v === "string" && v.trim()) return v.trim();
  return "";
}

const EMPLOYEE_FIELD_ALIASES = {
  employeeId: ["employee id", "emp id", "id"],
  name: ["employee name", "name"],
  joiningDate: ["joining date"],
  jobTitle: ["job title"],
  section: ["section"],
  basic: ["basic salary", "basic"],
  houseRent: ["house rent"],
  conveyance: ["conveyance"],
  food: ["food allowance", "food"],
  medical: ["medical allowance", "medical"],
  gross: ["gross salary", "gross"],
  casualLeaveAlloc: ["casual leave allocation", "casual leave"],
  medicalLeaveAlloc: ["medical leave allocation", "medical leave"],
  motherName: ["mother name"],
  fatherName: ["father name"],
  dob: ["date of birth", "dob"],
  nid: ["nid number", "nid"],
  maritalStatus: ["marital status"],
  status: ["employee status", "status"],
};
async function parseEmployeeExcel(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  const detected = hrDetectHeaderMap(rows, EMPLOYEE_FIELD_ALIASES, 3);
  if (!detected) return { rows: [], errors: [{ row: 0, msg: "Could not find a recognizable header row (need at least Employee ID + Name)." }] };
  const map = detected.map;
  const cell = (row, f) => (map[f] === undefined ? "" : row[map[f]]);
  const out = [];
  const errors = [];
  rows.slice(detected.rowIndex + 1).forEach((row, i) => {
    const employeeId = (cell(row, "employeeId") || "").toString().trim();
    const name = (cell(row, "name") || "").toString().trim();
    if (!employeeId && !name) return; // blank row
    const rowNum = detected.rowIndex + 2 + i;
    if (!employeeId) { errors.push({ row: rowNum, msg: `Missing Employee ID for "${name}"` }); return; }
    if (!name) { errors.push({ row: rowNum, msg: `Missing Name for Employee ID ${employeeId}` }); return; }
    out.push({
      employeeId, name,
      joiningDate: hrExcelDateToStr(cell(row, "joiningDate")),
      jobTitle: (cell(row, "jobTitle") || "").toString(),
      section: (cell(row, "section") || "").toString(),
      basic: num(cell(row, "basic")), houseRent: num(cell(row, "houseRent")),
      conveyance: num(cell(row, "conveyance")), food: num(cell(row, "food")),
      medical: num(cell(row, "medical")), gross: num(cell(row, "gross")),
      casualLeaveAlloc: num(cell(row, "casualLeaveAlloc")) || 10,
      medicalLeaveAlloc: num(cell(row, "medicalLeaveAlloc")) || 14,
      motherName: (cell(row, "motherName") || "").toString(),
      fatherName: (cell(row, "fatherName") || "").toString(),
      dob: hrExcelDateToStr(cell(row, "dob")),
      nid: (cell(row, "nid") || "").toString(),
      maritalStatus: (cell(row, "maritalStatus") || "").toString(),
      status: STATUSES.find((s) => s.toLowerCase() === (cell(row, "status") || "").toString().trim().toLowerCase()) || "New Join",
    });
  });
  return { rows: out, errors };
}

const ATTENDANCE_FIELD_ALIASES = {
  status: ["employee staus", "employee status", "status"],
  employeeId: ["emp id", "employee id", "employee no", "emp no", "employee number"],
  name: ["name", "employee name"],
  section: ["section"],
  jobTitle: ["job title"],
  joiningDate: ["joining date"],
  present: ["present days", "present"],
  weekend: ["weekend / holidays", "weekend/holidays", "weekend"],
  leave: ["leave"],
  absent: ["absent days", "absent  days", "absent"],
  basic: ["basic"],
  houseRent: ["house rent"],
  medical: ["medical"],
  conveyance: ["conveyance"],
  food: ["food allowance"],
  gross: ["gross"],
  otHours: [
    "ot hours", "ot hrs", "ot hour", "ot hr", "ot",
    "o t hours", "o t hrs", "overtime", "overtime hours",
    "overtime hrs", "overtime hour", "overtime hr"
  ],
  totalDays: ["total days", "total day"],
  payableDays: ["payable days", "payable day"],
  advance: ["advance amount", "advance amount (tk.)", "advance"],
  arrear: ["arrear", "arrear (tk.)"],
};
async function parseAttendanceExcel(file, employees) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  const detected = hrDetectHeaderMap(rows, ATTENDANCE_FIELD_ALIASES, 5);
  const byId = new Map(employees.map((e) => [e.employeeId.toString().trim(), e]));
  const byNumericId = new Map();
  employees.forEach((e) => {
    const raw = e.employeeId == null ? "" : e.employeeId.toString().trim();
    const numeric = raw.replace(/\.0+$/, "");
    if (/^\d+$/.test(numeric)) {
      const key = numeric.replace(/^0+(?=\d)/, "");
      if (!byNumericId.has(key)) byNumericId.set(key, e);
      else byNumericId.set(key, null); // ambiguous IDs are never auto-matched
    }
  });
  if (!detected) {
    return { records: [], summary: { total: 0, imported: 0, duplicate: 0, notFound: 0, missing: 0 }, errors: [{ row: 0, msg: "Could not detect the attendance header row." }] };
  }
  const map = detected.map;
  const cell = (row, f) => (map[f] === undefined ? "" : row[map[f]]);
  const seen = new Set();
  const records = [];
  let total = 0, duplicate = 0, notFound = 0, missing = 0;
  const errors = [];
  rows.slice(detected.rowIndex + 1).forEach((row, i) => {
    const employeeId = (cell(row, "employeeId") || "").toString().trim();
    const name = (cell(row, "name") || "").toString().trim();
    if (!employeeId && !name) return;
    total += 1;
    const rowNum = detected.rowIndex + 2 + i;
    if (!employeeId) { missing += 1; errors.push({ row: rowNum, msg: `Missing Employee ID for "${name || "unnamed row"}"` }); return; }
    if (seen.has(employeeId)) { duplicate += 1; errors.push({ row: rowNum, msg: `Duplicate Employee ID ${employeeId} in this file` }); return; }
    let matchedEmployee = byId.get(employeeId);
    if (!matchedEmployee) {
      const numericId = employeeId.replace(/\.0+$/, "").replace(/^0+(?=\d)/, "");
      if (/^\d+$/.test(numericId)) matchedEmployee = byNumericId.get(numericId) || null;
    }
    if (!matchedEmployee) { notFound += 1; errors.push({ row: rowNum, msg: `Employee ID ${employeeId} not found in Employee Database` }); return; }
    const matchedId = matchedEmployee.employeeId.toString().trim();
    if (seen.has(matchedId)) { duplicate += 1; errors.push({ row: rowNum, msg: `Duplicate Employee ID ${employeeId} in this file` }); return; }
    seen.add(matchedId);
    records.push({
      employeeId: matchedId,
      present: num(cell(row, "present")), weekend: num(cell(row, "weekend")),
      leave: num(cell(row, "leave")), absent: num(cell(row, "absent")),
      otHours: num(cell(row, "otHours")),
      advance: num(cell(row, "advance")), arrear: num(cell(row, "arrear")), tds: 0,
      basic: cell(row, "basic") !== "" ? num(cell(row, "basic")) : undefined,
      houseRent: cell(row, "houseRent") !== "" ? num(cell(row, "houseRent")) : undefined,
      medical: cell(row, "medical") !== "" ? num(cell(row, "medical")) : undefined,
      conveyance: cell(row, "conveyance") !== "" ? num(cell(row, "conveyance")) : undefined,
      food: cell(row, "food") !== "" ? num(cell(row, "food")) : undefined,
      gross: cell(row, "gross") !== "" ? num(cell(row, "gross")) : undefined,
    });
  });
  return { records, summary: { total, imported: records.length, duplicate, notFound, missing }, errors };
}

/* =========================================================================
   STORAGE
   Supabase (via /api/* API + Prisma) is the source of truth.
   Pattern for every dataset: instant local-cache paint, then a network
   refresh; saves mirror locally immediately but return TRUE only after
   the database confirms the write.
========================================================================= */
const SKEY = "hrms:v1";
const HRM_SESSION_TOKEN_KEY = "hrm_session_token";

function hrHasArtifactStorage() {
  return typeof window !== "undefined" && window.storage && typeof window.storage.get === "function";
}
function hrHasLocalStorage() {
  try { const k = "__hrms_probe__"; window.localStorage.setItem(k, "1"); window.localStorage.removeItem(k); return true; }
  catch (e) { return false; }
}
function hrGetAuthToken() {
  try { return window.sessionStorage.getItem(HRM_SESSION_TOKEN_KEY); } catch { return null; }
}
function readLocalJson(key) {
  try {
    if (hrHasLocalStorage()) {
      const v = window.localStorage.getItem(key);
      if (v) return JSON.parse(v);
    }
  } catch (e) { console.error(e); }
  return null;
}
function writeLocalJson(key, value) {
  let ok = false;
  const payload = JSON.stringify(value);
  if (hrHasLocalStorage()) { try { window.localStorage.setItem(key, payload); ok = true; } catch (e) { console.error(e); } }
  return ok;
}
async function readArtifactJson(key) {
  if (!hrHasArtifactStorage()) return null;
  try { const res = await window.storage.get(key); if (res && res.value) return JSON.parse(res.value); } catch {}
  return null;
}
async function writeArtifactJson(key, value) {
  if (!hrHasArtifactStorage()) return false;
  try { await window.storage.set(key, JSON.stringify(value), false); return true; } catch (e) { console.error(e); return false; }
}
async function apiGet(path) {
  const token = hrGetAuthToken();
  if (!token) throw Object.assign(new Error("not signed in"), { code: "NOAUTH" });
  const res = await fetch(path, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) {
    const err = new Error(`GET ${path} failed (${res.status})`);
    err.status = res.status;
    throw err;
  }
  return res.json();
}
async function apiPost(path, bodyObj) {
  const token = hrGetAuthToken();
  if (!token) throw Object.assign(new Error("not signed in"), { code: "NOAUTH" });
  const res = await fetch(path, {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
    body: JSON.stringify(bodyObj),
  });
  if (!res.ok) {
    const err = new Error(`POST ${path} failed (${res.status})`);
    err.status = res.status;
    throw err;
  }
  return res.json();
}

/* ---------- HR state ---------- */
function normalizeLoadedState(raw) {
  if (!raw || typeof raw !== "object") return null;
  if (!Array.isArray(raw.employees) && !Array.isArray(raw.attendanceRecords)) return null;
  // payrollRecords (server-derived Payroll ledger) is not needed client-side — skip caching it.
  const { payrollRecords, ...clean } = raw;
  return {
    employees: clean.employees || [],
    attendanceRecords: clean.attendanceRecords || [],
    payrollApprovals: clean.payrollApprovals || {},
    rules: clean.rules || undefined,
    ...clean,
  };
}
async function loadCachedState() {
  const artifact = await readArtifactJson(SKEY);
  if (artifact) return artifact;
  return readLocalJson(SKEY);
}
async function fetchServerState() {
  const data = await apiGet("/api/state");
  const normalized = normalizeLoadedState(data.state);
  if (normalized) writeLocalJson(SKEY, normalized); // refresh offline cache
  return normalized === null ? (data.state ? {} : null) : normalized;
}
async function loadState() {
  const cached = await loadCachedState();
  if (cached) return cached;
  try { return await fetchServerState(); } catch (e) { console.error("State fetch failed:", e.message); return null; }
}
let hrSaveChain = Promise.resolve();
let hrPendingSave; // latest state waiting to be sent; undefined = nothing pending
async function saveState(state) {
  /* Saves are serialized: only ONE POST is ever in flight, and each queued
     run sends the LATEST state (intermediate keystroke snapshots are
     coalesced). Overlapping full-state replaces used to commit out of
     order on slow links — an older snapshot could land last and wipe
     newer payroll entries (they "vanished" after a reload). */
  writeLocalJson(SKEY, state);
  void writeArtifactJson(SKEY, state);
  hrPendingSave = state;
  const attempt = async () => {
    if (hrPendingSave === undefined) return true; // superseded by a newer save that already carried this data
    const payload = hrPendingSave;
    hrPendingSave = undefined;
    try {
      const data = await apiPost("/api/state", { state: payload });
      return data.ok === true;
    } catch (e) {
      if (e.status === 401) console.error("Save rejected: session expired.");
      else if (e.status === 403) console.error("Save rejected: your role cannot modify HR data.");
      else console.error("Database unreachable during save:", e.message);
      return false;
    }
  };
  const result = hrSaveChain.then(attempt, attempt);
  hrSaveChain = result.then(() => {}, () => {});
  return result;
}

/* Computed salary rows (what the Payroll tab renders) are materialized into
   the Supabase "Payroll" table by the SERVER on every save — see api/state.js.
   The client does not need to send them. */

/* ---------- Ledger records ---------- */
async function loadCachedLedger(key) {
  const artifact = await readArtifactJson(key);
  if (artifact) return artifact;
  return readLocalJson(key);
}
async function fetchServerLedger() {
  const data = await apiGet("/api/ledger");
  if (!Array.isArray(data.records)) return null;
  writeLocalJson(STORAGE_KEY, data.records);
  return data.records;
}
async function loadRecords() {
  const cached = await loadCachedLedger(STORAGE_KEY);
  if (cached !== null) return cached;
  try { return await fetchServerLedger(); } catch (e) { console.error("Ledger fetch failed:", e.message); return null; }
}
async function saveRecords(records) {
  writeLocalJson(STORAGE_KEY, records);
  await writeArtifactJson(STORAGE_KEY, records);
  try {
    const data = await apiPost("/api/ledger", { records });
    return data.ok === true;
  } catch (e) {
    if (e.status === 401) console.error("Ledger save rejected: session expired.");
    else if (e.status === 403) console.error("Ledger save rejected: role not permitted.");
    else console.error("Database unreachable during ledger save:", e.message);
    return false;
  }
}

/* ---------- Global save-status bus (drives the on-screen sync chip) ---------- */
const syncListeners = new Set();
function emitSync(status) { syncListeners.forEach((f) => { try { f(status); } catch {} }); }
function useSyncStatus() {
  const [status, setStatus] = useState("idle");
  useEffect(() => {
    const fn = (s) => setStatus(s);
    syncListeners.add(fn);
    return () => syncListeners.delete(fn);
  }, []);
  useEffect(() => {
    if (status === "saved") { const t = setTimeout(() => setStatus("idle"), 2500); return () => clearTimeout(t); }
  }, [status]);
  return status;
}

/* =========================================================================
   SHARED UI
========================================================================= */
function HrField({ label, children, span }) {
  return (
    <label style={{ display: "flex", flexDirection: "column", gap: 5, gridColumn: span ? `span ${span}` : undefined }}>
      <span style={{ fontFamily: "'IBM Plex Sans'", fontSize: 11, fontWeight: 600, letterSpacing: ".04em", textTransform: "uppercase", color: HR_T.inkSoft }}>{label}</span>
      {children}
    </label>
  );
}
const hrInputStyle = {
  fontFamily: "'IBM Plex Mono', monospace", fontSize: 13.5, color: HR_T.ink, background: HR_T.panel,
  border: `1px solid ${HR_T.line}`, borderRadius: 6, padding: "9px 11px", outline: "none", width: "100%", boxSizing: "border-box",
};
function HrTInput(props) { return <input {...props} style={{ ...hrInputStyle, ...(props.style || {}) }} />; }
function HrTSelect({ children, ...props }) {
  return (
    <div style={{ position: "relative" }}>
      <select {...props} style={{ ...hrInputStyle, appearance: "none", paddingRight: 30, cursor: "pointer" }}>{children}</select>
      <ChevronDown size={14} color={HR_T.muted} style={{ position: "absolute", right: 10, top: 12, pointerEvents: "none" }} />
    </div>
  );
}
function Badge({ children, tone }) {
  const tones = {
    good: { bg: "#E4F0EA", fg: HR_T.good }, bad: { bg: "#F6E4DE", fg: HR_T.bad },
    neutral: { bg: HR_T.chip, fg: HR_T.inkSoft }, amber: { bg: HR_T.amberSoft, fg: "#8A5A16" },
  };
  const c = tones[tone] || tones.neutral;
  return (
    <span style={{ fontFamily: "'IBM Plex Sans'", fontSize: 11, fontWeight: 600, padding: "3px 9px", borderRadius: 20, background: c.bg, color: c.fg, whiteSpace: "nowrap" }}>
      {children}
    </span>
  );
}
function statusTone(s) {
  if (s === "Regular" || s === "New Join") return "good";
  if (s === "Resigned" || s === "Terminated") return "bad";
  return "neutral";
}
// Signature motif: a perforated / stitched divider, evoking a torn work docket.
function StitchDivider() {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 4, margin: "18px 0", opacity: 0.6 }}>
      {Array.from({ length: 46 }).map((_, i) => (
        <div key={i} style={{ width: 3, height: 1.5, background: HR_T.muted, borderRadius: 1 }} />
      ))}
    </div>
  );
}
function DocketCard({ children, style }) {
  return (
    <div style={{
      background: HR_T.panel, border: `1px solid ${HR_T.line}`, borderRadius: 4, position: "relative",
      boxShadow: "0 1px 0 rgba(30,36,51,.03)", ...style,
    }}>
      <div style={{ position: "absolute", left: -1, top: 0, bottom: 0, width: 4, background: HR_T.indigo, borderRadius: "4px 0 0 4px" }} />
      {children}
    </div>
  );
}
function HrStatCard({ icon: Icon, label, value, sub, tone }) {
  const accent = tone === "amber" ? HR_T.amber : tone === "bad" ? HR_T.bad : tone === "good" ? HR_T.good : HR_T.indigo;
  return (
    <DocketCard style={{ padding: "17px 19px", display: "flex", flexDirection: "column", gap: 9 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, color: HR_T.inkSoft }}>
        <Icon size={15} strokeWidth={2} />
        <span style={{ fontFamily: "'IBM Plex Sans'", fontSize: 11.5, fontWeight: 600, letterSpacing: ".04em", textTransform: "uppercase" }}>{label}</span>
      </div>
      <div style={{ fontFamily: "'Space Grotesk'", fontSize: 25, fontWeight: 700, color: accent }}>{value}</div>
      {sub && <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 11.5, color: HR_T.muted }}>{sub}</div>}
    </DocketCard>
  );
}
function Btn({ children, variant = "default", ...props }) {
  const styles = {
    default: { background: HR_T.panel, border: `1px solid ${HR_T.line}`, color: HR_T.inkSoft },
    primary: { background: HR_T.indigo, border: "none", color: "#fff" },
    amber: { background: HR_T.amber, border: "none", color: "#fff" },
    danger: { background: "none", border: "none", color: HR_T.bad },
    ghost: { background: "none", border: "none", color: HR_T.inkSoft },
  };
  return (
    <button {...props} style={{
      display: "inline-flex", alignItems: "center", gap: 7, padding: "9px 15px", borderRadius: 7,
      fontFamily: "'IBM Plex Sans'", fontSize: 13, fontWeight: 600, cursor: "pointer",
      ...styles[variant], ...(props.style || {}),
    }}>{children}</button>
  );
}
function ImportSummaryBox({ summary, errors }) {
  if (!summary) return null;
  return (
    <div style={{ background: HR_T.canvasDeep, border: `1px solid ${HR_T.line}`, borderRadius: 8, padding: 16, marginTop: 14 }}>
      <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 12, fontWeight: 700, color: HR_T.ink, marginBottom: 10, display: "flex", alignItems: "center", gap: 7 }}>
        <ClipboardList size={14} /> Import summary
      </div>
      <div style={{ display: "flex", gap: 18, flexWrap: "wrap", fontFamily: "'IBM Plex Mono'", fontSize: 12.5 }}>
        <span>Total rows: <b>{summary.total}</b></span>
        <span style={{ color: HR_T.good }}>Imported: <b>{summary.imported}</b></span>
        {summary.duplicate > 0 && <span style={{ color: HR_T.amber }}>Duplicate ID: <b>{summary.duplicate}</b></span>}
        {summary.notFound > 0 && <span style={{ color: HR_T.bad }}>ID not found: <b>{summary.notFound}</b></span>}
        {summary.missing > 0 && <span style={{ color: HR_T.bad }}>Missing data: <b>{summary.missing}</b></span>}
      </div>
      {errors && errors.length > 0 && (
        <div style={{ marginTop: 10, maxHeight: 140, overflowY: "auto" }}>
          {errors.slice(0, 30).map((e, i) => (
            <div key={i} style={{ fontFamily: "'IBM Plex Sans'", fontSize: 11.5, color: HR_T.bad, padding: "3px 0" }}>
              Row {e.row}: {e.msg}
            </div>
          ))}
          {errors.length > 30 && <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 11.5, color: HR_T.muted }}>…and {errors.length - 30} more</div>}
        </div>
      )}
    </div>
  );
}

/* =========================================================================
   EMPLOYEE FORM
========================================================================= */
const emptyEmployee = {
  id: null, employeeId: "", name: "", joiningDate: "", jobTitle: "", section: "",
  basic: "", houseRent: "", conveyance: "", food: "", medical: "", gross: "",
  casualLeaveAlloc: 10, medicalLeaveAlloc: 14, motherName: "", fatherName: "",
  dob: "", nid: "", maritalStatus: "Single", status: "New Join",
};
function EmployeeForm({ employees, initial, onSave, onCancel }) {
  const [form, setForm] = useState(initial || emptyEmployee);
  const isEdit = !!(initial && initial.id);
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));
  const [err, setErr] = useState("");

  function submit(e) {
    e.preventDefault();
    if (!form.employeeId || !form.name) { setErr("Employee ID and Name are required."); return; }
    const dup = employees.find((emp) => emp.employeeId === form.employeeId && emp.id !== form.id);
    if (dup) { setErr(`Employee ID ${form.employeeId} already exists (${dup.name}).`); return; }
    onSave({ ...form, id: isEdit ? form.id : `emp-${Date.now()}` });
  }

  return (
    <form onSubmit={submit}>
      <DocketCard style={{ padding: 26 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 18 }}>
          <h2 style={{ fontFamily: "'Space Grotesk'", fontSize: 21, margin: 0, color: HR_T.ink }}>
            {isEdit ? "Edit employee" : "New employee"}
          </h2>
          {onCancel && <Btn variant="ghost" type="button" onClick={onCancel}><X size={18} /></Btn>}
        </div>

        {err && <div style={{ background: "#F6E4DE", color: HR_T.bad, padding: "9px 12px", borderRadius: 6, fontFamily: "'IBM Plex Sans'", fontSize: 12.5, marginBottom: 14 }}>{err}</div>}

        <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 11, fontWeight: 700, color: HR_T.indigo, textTransform: "uppercase", letterSpacing: ".05em", marginBottom: 10 }}>Employment</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14 }}>
          <HrField label="Employee ID"><HrTInput value={form.employeeId} onChange={set("employeeId")} required disabled={isEdit} /></HrField>
          <HrField label="Employee Name" span={2}><HrTInput value={form.name} onChange={set("name")} required /></HrField>
          <HrField label="Status">
            <HrTSelect value={form.status} onChange={set("status")}>{STATUSES.map((s) => <option key={s} value={s}>{s}</option>)}</HrTSelect>
          </HrField>
          <HrField label="Joining Date"><HrTInput type="date" value={form.joiningDate} onChange={set("joiningDate")} /></HrField>
          <HrField label="Job Title"><HrTInput value={form.jobTitle} onChange={set("jobTitle")} /></HrField>
          <HrField label="Section"><HrTInput value={form.section} onChange={set("section")} /></HrField>
        </div>

        <StitchDivider />
        <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 11, fontWeight: 700, color: HR_T.indigo, textTransform: "uppercase", letterSpacing: ".05em", marginBottom: 10 }}>Salary components</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14 }}>
          <HrField label="Basic Salary"><HrTInput type="number" value={form.basic} onChange={set("basic")} /></HrField>
          <HrField label="House Rent"><HrTInput type="number" value={form.houseRent} onChange={set("houseRent")} /></HrField>
          <HrField label="Conveyance"><HrTInput type="number" value={form.conveyance} onChange={set("conveyance")} /></HrField>
          <HrField label="Food Allowance"><HrTInput type="number" value={form.food} onChange={set("food")} /></HrField>
          <HrField label="Medical Allowance"><HrTInput type="number" value={form.medical} onChange={set("medical")} /></HrField>
          <HrField label="Gross Salary"><HrTInput type="number" value={form.gross} onChange={set("gross")} /></HrField>
          <HrField label="Casual Leave Allocation"><HrTInput type="number" value={form.casualLeaveAlloc} onChange={set("casualLeaveAlloc")} /></HrField>
          <HrField label="Medical Leave Allocation"><HrTInput type="number" value={form.medicalLeaveAlloc} onChange={set("medicalLeaveAlloc")} /></HrField>
        </div>

        <StitchDivider />
        <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 11, fontWeight: 700, color: HR_T.indigo, textTransform: "uppercase", letterSpacing: ".05em", marginBottom: 10 }}>Personal</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14 }}>
          <HrField label="Mother Name"><HrTInput value={form.motherName} onChange={set("motherName")} /></HrField>
          <HrField label="Father Name"><HrTInput value={form.fatherName} onChange={set("fatherName")} /></HrField>
          <HrField label="Date of Birth"><HrTInput type="date" value={form.dob} onChange={set("dob")} /></HrField>
          <HrField label="Age"><HrTInput value={calcAge(form.dob)} disabled style={{ color: HR_T.muted }} /></HrField>
          <HrField label="NID Number"><HrTInput value={form.nid} onChange={set("nid")} /></HrField>
          <HrField label="Marital Status">
            <HrTSelect value={form.maritalStatus} onChange={set("maritalStatus")}>{MARITAL.map((s) => <option key={s} value={s}>{s}</option>)}</HrTSelect>
          </HrField>
        </div>

        <div style={{ display: "flex", gap: 10, marginTop: 24, justifyContent: "flex-end" }}>
          {onCancel && <Btn type="button" onClick={onCancel}>Cancel</Btn>}
          <Btn variant="primary" type="submit"><Plus size={15} /> {isEdit ? "Save changes" : "Add employee"}</Btn>
        </div>
      </DocketCard>
    </form>
  );
}

/* =========================================================================
   EMPLOYEES TABLE
========================================================================= */
function EmployeesTable({ employees, onEdit, onDelete, onOpenProfile, onImport, fileInputRef, importBusy, setImportBusy }) {
  const [q, setQ] = useState("");
  const [sectionF, setSectionF] = useState("");
  const [titleF, setTitleF] = useState("");
  const [statusF, setStatusF] = useState("");
  const [sort, setSort] = useState({ key: "employeeId", dir: "asc" });

  const sections = useMemo(() => Array.from(new Set(employees.map((e) => e.section).filter(Boolean))).sort(), [employees]);
  const titles = useMemo(() => Array.from(new Set(employees.map((e) => e.jobTitle).filter(Boolean))).sort(), [employees]);

  const filtered = useMemo(() => {
    let rows = [...employees];
    if (q) {
      const s = q.toLowerCase();
      rows = rows.filter((e) => [e.employeeId, e.name, e.section, e.jobTitle].some((v) => (v || "").toString().toLowerCase().includes(s)));
    }
    if (sectionF) rows = rows.filter((e) => e.section === sectionF);
    if (titleF) rows = rows.filter((e) => e.jobTitle === titleF);
    if (statusF) rows = rows.filter((e) => e.status === statusF);
    rows.sort((a, b) => {
      const { key, dir } = sort;
      let av = a[key], bv = b[key];
      if (typeof av === "string") { av = (av || "").toLowerCase(); bv = (bv || "").toLowerCase(); }
      if (av < bv) return dir === "asc" ? -1 : 1;
      if (av > bv) return dir === "asc" ? 1 : -1;
      return 0;
    });
    return rows;
  }, [employees, q, sectionF, titleF, statusF, sort]);

  function toggleSort(key) { setSort((s) => (s.key === key ? { key, dir: s.dir === "asc" ? "desc" : "asc" } : { key, dir: "asc" })); }

  function exportExcel(rows) {
    const data = rows.map((e) => ({
      SL: e.sl, "Employee Name": e.name, "Employee ID": e.employeeId, "Joining Date": e.joiningDate,
      "Job Title": e.jobTitle, Section: e.section, "Basic Salary": e.basic, "House Rent": e.houseRent,
      Conveyance: e.conveyance, "Food Allowance": e.food, "Medical Allowance": e.medical, "Gross Salary": e.gross,
      "Casual Leave Allocation": e.casualLeaveAlloc, "Medical Leave Allocation": e.medicalLeaveAlloc,
      "Mother Name": e.motherName, "Father Name": e.fatherName, "Date of Birth": e.dob, Age: calcAge(e.dob),
      "NID Number": e.nid, "Marital Status": e.maritalStatus, "Employee Status": e.status,
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Employees");
    XLSX.writeFile(wb, "employees.xlsx");
  }

  const Th = ({ label, k, align }) => (
    <th onClick={() => toggleSort(k)} style={{
      textAlign: align || "left", padding: "10px 12px", cursor: "pointer", userSelect: "none",
      fontFamily: "'IBM Plex Sans'", fontSize: 10.5, fontWeight: 700, letterSpacing: ".04em",
      textTransform: "uppercase", color: HR_T.inkSoft, whiteSpace: "nowrap",
    }}>
      <span style={{ display: "inline-flex", alignItems: "center", gap: 4 }}>{label} {sort.key === k && <ArrowUpDown size={10} />}</span>
    </th>
  );

  async function handleFile(e) {
    const file = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!file) return;
    setImportBusy(true);
    try { await onImport(file); } finally { setImportBusy(false); }
  }

  return (
    <div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 10, marginBottom: 16, alignItems: "center" }}>
        <div style={{ position: "relative", flex: "1 1 220px" }}>
          <Search size={14} color={HR_T.muted} style={{ position: "absolute", left: 11, top: 11 }} />
          <HrTInput placeholder="Search ID, name, section, title…" value={q} onChange={(e) => setQ(e.target.value)} style={{ paddingLeft: 32 }} />
        </div>
        <HrTSelect value={sectionF} onChange={(e) => setSectionF(e.target.value)} style={{ maxWidth: 170 }}>
          <option value="">All sections</option>{sections.map((s) => <option key={s} value={s}>{s}</option>)}
        </HrTSelect>
        <HrTSelect value={titleF} onChange={(e) => setTitleF(e.target.value)} style={{ maxWidth: 170 }}>
          <option value="">All job titles</option>{titles.map((s) => <option key={s} value={s}>{s}</option>)}
        </HrTSelect>
        <HrTSelect value={statusF} onChange={(e) => setStatusF(e.target.value)} style={{ maxWidth: 160 }}>
          <option value="">All statuses</option>{STATUSES.map((s) => <option key={s} value={s}>{s}</option>)}
        </HrTSelect>
        <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
        <Btn onClick={() => fileInputRef.current && fileInputRef.current.click()} disabled={importBusy}>
          <Upload size={14} /> {importBusy ? "Reading…" : "Import Excel"}
        </Btn>
        <Btn onClick={() => exportExcel(filtered)}><Download size={14} /> Export</Btn>
      </div>

      <DocketCard style={{ overflow: "hidden" }}>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 1100 }}>
            <thead>
              <tr style={{ borderBottom: `2px solid ${HR_T.indigo}` }}>
                <Th label="Emp ID" k="employeeId" /><Th label="Name" k="name" /><Th label="Section" k="section" />
                <Th label="Job Title" k="jobTitle" /><Th label="Joining" k="joiningDate" />
                <Th label="Gross" k="gross" align="right" /><Th label="Status" k="status" /><th style={{ padding: "10px 12px" }} />
              </tr>
            </thead>
            <tbody>
              {filtered.map((e, i) => (
                <tr key={e.id} style={{ borderBottom: `1px solid ${HR_T.line}`, background: i % 2 ? HR_T.canvas : HR_T.panel, cursor: "pointer" }}
                    onClick={() => onOpenProfile(e)}>
                  <td style={{ padding: "9px 12px", fontFamily: "'IBM Plex Mono'", fontSize: 12.5, borderLeft: `2px solid ${HR_T.indigo}` }}>{e.employeeId}</td>
                  <td style={{ padding: "9px 12px", fontFamily: "'IBM Plex Sans'", fontSize: 13, fontWeight: 500 }}>{e.name}</td>
                  <td style={{ padding: "9px 12px", fontFamily: "'IBM Plex Sans'", fontSize: 12.5, color: HR_T.inkSoft }}>{e.section || "—"}</td>
                  <td style={{ padding: "9px 12px", fontFamily: "'IBM Plex Sans'", fontSize: 12.5, color: HR_T.inkSoft }}>{e.jobTitle || "—"}</td>
                  <td style={{ padding: "9px 12px", fontFamily: "'IBM Plex Sans'", fontSize: 12.5 }}>{hrFmtDate(e.joiningDate)}</td>
                  <td style={{ padding: "9px 12px", fontFamily: "'IBM Plex Mono'", fontSize: 13, textAlign: "right" }}>{hrMoney(e.gross)}</td>
                  <td style={{ padding: "9px 12px" }}><Badge tone={statusTone(e.status)}>{e.status}</Badge></td>
                  <td style={{ padding: "9px 12px" }} onClick={(ev) => ev.stopPropagation()}>
                    <div style={{ display: "flex", gap: 6 }}>
                      <Btn variant="ghost" onClick={() => onEdit(e)}><Pencil size={14} /></Btn>
                      <Btn variant="danger" onClick={() => onDelete(e.id)}><Trash2 size={14} /></Btn>
                    </div>
                  </td>
                </tr>
              ))}
              {filtered.length === 0 && (
                <tr><td colSpan={8} style={{ padding: 40, textAlign: "center", fontFamily: "'IBM Plex Sans'", color: HR_T.muted, fontSize: 13 }}>
                  {employees.length === 0 ? "No employees yet. Add one, or import your Excel file." : "No employees match these filters."}
                </td></tr>
              )}
            </tbody>
          </table>
        </div>
      </DocketCard>
    </div>
  );
}

/* =========================================================================
   EMPLOYEE PROFILE
========================================================================= */
function EmployeeProfile({ employee, attendanceByEmployee, payrollByEmployee, leaveBalance, onBack, onEdit }) {
  const att = attendanceByEmployee(employee.employeeId);
  const pay = payrollByEmployee(employee.employeeId);
  const lb = leaveBalance(employee);

  return (
    <div>
      <Btn variant="ghost" onClick={onBack} style={{ marginBottom: 14 }}><ArrowLeft size={15} /> Back to employees</Btn>

      <DocketCard style={{ padding: 24, marginBottom: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div>
            <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 11, color: HR_T.amber, fontWeight: 700, letterSpacing: ".06em", textTransform: "uppercase" }}>
              {employee.employeeId} · {employee.section || "—"}
            </div>
            <h2 style={{ fontFamily: "'Space Grotesk'", fontSize: 26, margin: "3px 0 4px", color: HR_T.ink }}>{employee.name}</h2>
            <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 13, color: HR_T.inkSoft }}>{employee.jobTitle || "—"} · Joined {hrFmtDate(employee.joiningDate)}</div>
          </div>
          <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
            <Badge tone={statusTone(employee.status)}>{employee.status}</Badge>
            <Btn onClick={() => onEdit(employee)}><Pencil size={14} /> Edit</Btn>
          </div>
        </div>
      </DocketCard>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16, marginBottom: 18 }}>
        <DocketCard style={{ padding: 20 }}>
          <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 11, fontWeight: 700, color: HR_T.indigo, textTransform: "uppercase", marginBottom: 12 }}>Personal</div>
          {[["Mother's Name", employee.motherName], ["Father's Name", employee.fatherName], ["Date of Birth", hrFmtDate(employee.dob)],
            ["Age", calcAge(employee.dob)], ["NID Number", employee.nid], ["Marital Status", employee.maritalStatus]].map(([l, v]) => (
            <div key={l} style={{ display: "flex", justifyContent: "space-between", fontFamily: "'IBM Plex Sans'", fontSize: 12.5, padding: "5px 0", borderBottom: `1px dashed ${HR_T.line}` }}>
              <span style={{ color: HR_T.muted }}>{l}</span><span style={{ color: HR_T.ink, fontWeight: 500 }}>{v || "—"}</span>
            </div>
          ))}
        </DocketCard>
        <DocketCard style={{ padding: 20 }}>
          <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 11, fontWeight: 700, color: HR_T.indigo, textTransform: "uppercase", marginBottom: 12 }}>Salary components</div>
          {[["Basic", employee.basic], ["House Rent", employee.houseRent], ["Conveyance", employee.conveyance],
            ["Food Allowance", employee.food], ["Medical Allowance", employee.medical]].map(([l, v]) => (
            <div key={l} style={{ display: "flex", justifyContent: "space-between", fontFamily: "'IBM Plex Sans'", fontSize: 12.5, padding: "5px 0", borderBottom: `1px dashed ${HR_T.line}` }}>
              <span style={{ color: HR_T.muted }}>{l}</span><span style={{ fontFamily: "'IBM Plex Mono'", color: HR_T.ink }}>{hrMoney(v)}</span>
            </div>
          ))}
          <div style={{ display: "flex", justifyContent: "space-between", fontFamily: "'IBM Plex Sans'", fontSize: 13, fontWeight: 700, padding: "8px 0 0" }}>
            <span>Gross</span><span style={{ fontFamily: "'IBM Plex Mono'", color: HR_T.indigo }}>৳ {hrMoney(employee.gross)}</span>
          </div>
        </DocketCard>
        <DocketCard style={{ padding: 20 }}>
          <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 11, fontWeight: 700, color: HR_T.indigo, textTransform: "uppercase", marginBottom: 12 }}>Leave balance</div>
          {[["Casual", lb.casualAlloc, lb.casualUsed], ["Medical", lb.medicalAlloc, lb.medicalUsed]].map(([l, alloc, used]) => (
            <div key={l} style={{ marginBottom: 12 }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontFamily: "'IBM Plex Sans'", fontSize: 12.5, marginBottom: 4 }}>
                <span style={{ color: HR_T.ink, fontWeight: 500 }}>{l} leave</span>
                <span style={{ fontFamily: "'IBM Plex Mono'", color: HR_T.inkSoft }}>{used} / {alloc} used</span>
              </div>
              <div style={{ height: 6, borderRadius: 4, background: HR_T.canvasDeep, overflow: "hidden" }}>
                <div style={{ height: "100%", width: `${alloc ? Math.min(100, (used / alloc) * 100) : 0}%`, background: used > alloc ? HR_T.bad : HR_T.amber, borderRadius: 4 }} />
              </div>
            </div>
          ))}
          <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 11, color: HR_T.muted, marginTop: 4 }}>
            Note: the imported attendance file records one combined "Leave" figure per month, not split by type — used counts are tracked against combined leave taken and shown per type against each allocation.
          </div>
        </DocketCard>
      </div>

      <DocketCard style={{ padding: 20, marginBottom: 18 }}>
        <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 11, fontWeight: 700, color: HR_T.indigo, textTransform: "uppercase", marginBottom: 12 }}>Attendance history</div>
        <SimpleTable
          cols={["Month", "Present", "Weekend/Hol.", "Leave", "Absent", "Total Days", "Payable Days"]}
          rows={att.map((a) => [monthLabel(a.month), a.present, a.weekend, a.leave, a.absent, a.present + a.weekend + a.leave + a.absent, a.present + a.weekend + a.leave])}
          empty="No attendance imported for this employee yet."
        />
      </DocketCard>

      <DocketCard style={{ padding: 20 }}>
        <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 11, fontWeight: 700, color: HR_T.indigo, textTransform: "uppercase", marginBottom: 12 }}>Payroll history</div>
        <SimpleTable
          cols={["Month", "Payable Days", "Gross", "OT", "Pay Amount", "Status"]}
          rows={pay.map((p) => [monthLabel(p.month), p.payableDays, `৳ ${hrMoney(p.gross)}`, `৳ ${hrMoney(p.otAmount)}`, `৳ ${hrMoney(p.payAmount)}`, p.approved ? "Approved" : "Draft"])}
          empty="No payroll generated for this employee yet."
        />
      </DocketCard>
    </div>
  );
}
function SimpleTable({ cols, rows, empty }) {
  return (
    <div style={{ overflowX: "auto" }}>
      <table style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead>
          <tr style={{ borderBottom: `1px solid ${HR_T.line}` }}>
            {cols.map((c) => <th key={c} style={{ textAlign: "left", padding: "7px 10px", fontFamily: "'IBM Plex Sans'", fontSize: 10.5, fontWeight: 700, textTransform: "uppercase", color: HR_T.inkSoft }}>{c}</th>)}
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => (
            <tr key={i} style={{ borderBottom: `1px dashed ${HR_T.line}` }}>
              {r.map((v, j) => <td key={j} style={{ padding: "7px 10px", fontFamily: "'IBM Plex Mono'", fontSize: 12.5, color: HR_T.ink }}>{v}</td>)}
            </tr>
          ))}
          {rows.length === 0 && <tr><td colSpan={cols.length} style={{ padding: 18, textAlign: "center", fontFamily: "'IBM Plex Sans'", color: HR_T.muted, fontSize: 12.5 }}>{empty}</td></tr>}
        </tbody>
      </table>
    </div>
  );
}

/* =========================================================================
   ATTENDANCE MANAGEMENT
========================================================================= */
function AttendanceManagement({ employees, attendanceRecords, onImport, monthPicker, setMonthPicker }) {
  const fileInputRef = useRef(null);
  const [importBusy, setImportBusy] = useState(false);
  const [lastResult, setLastResult] = useState(null);

  const monthsAvailable = useMemo(() => Array.from(new Set(attendanceRecords.map((r) => r.month))).sort().reverse(), [attendanceRecords]);
  const rowsForMonth = useMemo(() => {
    const recs = attendanceRecords.filter((r) => r.month === monthPicker);
    const byId = new Map(employees.map((e) => [e.employeeId, e]));
    return recs.map((r) => ({ ...r, emp: byId.get(r.employeeId) })).filter((r) => r.emp);
  }, [attendanceRecords, monthPicker, employees]);

  async function handleFile(e) {
    const file = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!file) return;
    if (employees.length === 0) { window.alert("Import your Employee Database first — attendance is matched to it by Employee ID."); return; }
    setImportBusy(true);
    try {
      const { records, summary, errors } = await parseAttendanceExcel(file, employees);
      setLastResult({ summary, errors });
      if (records.length === 0) return;
      const month = window.prompt("Which month does this file cover? (YYYY-MM)", currentMonthKey());
      if (!month || !/^\d{4}-\d{2}$/.test(month)) return;
      const existing = attendanceRecords.some((r) => r.month === month);
      if (existing) {
        const replace = window.confirm(`Attendance for ${monthLabel(month)} was already imported. Replace it with this file?`);
        if (!replace) return;
      }
      onImport(month, records);
      setMonthPicker(month);
    } finally {
      setImportBusy(false);
    }
  }

  return (
    <div>
      <DocketCard style={{ padding: 20, marginBottom: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
          <div>
            <div style={{ fontFamily: "'Space Grotesk'", fontSize: 17, color: HR_T.ink, marginBottom: 4 }}>Monthly attendance import</div>
            <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 12.5, color: HR_T.inkSoft }}>Matched to the Employee Database by Employee ID — names are never used for matching.</div>
          </div>
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
          <Btn variant="primary" onClick={() => fileInputRef.current && fileInputRef.current.click()} disabled={importBusy}>
            <Upload size={15} /> {importBusy ? "Reading…" : "Import monthly Excel"}
          </Btn>
        </div>
        {lastResult && <ImportSummaryBox summary={lastResult.summary} errors={lastResult.errors} />}
      </DocketCard>

      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
        <span style={{ fontFamily: "'IBM Plex Sans'", fontSize: 12.5, fontWeight: 600, color: HR_T.inkSoft }}>Viewing</span>
        <HrTSelect value={monthPicker} onChange={(e) => setMonthPicker(e.target.value)} style={{ maxWidth: 200 }}>
          {monthsAvailable.length === 0 && <option value="">No months imported</option>}
          {monthsAvailable.map((m) => <option key={m} value={m}>{monthLabel(m)}</option>)}
        </HrTSelect>
      </div>

      <DocketCard style={{ overflow: "hidden" }}>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 900 }}>
            <thead>
              <tr style={{ borderBottom: `2px solid ${HR_T.indigo}` }}>
                {["Emp ID", "Name", "Present", "Weekend/Hol.", "Leave", "Absent", "Total Days", "Payable Days", "OT Hrs"].map((h) => (
                  <th key={h} style={{ textAlign: h === "Name" || h === "Emp ID" ? "left" : "right", padding: "9px 12px", fontFamily: "'IBM Plex Sans'", fontSize: 10.5, fontWeight: 700, textTransform: "uppercase", color: HR_T.inkSoft }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rowsForMonth.map((r, i) => (
                <tr key={r.employeeId} style={{ borderBottom: `1px solid ${HR_T.line}`, background: i % 2 ? HR_T.canvas : HR_T.panel }}>
                  <td style={{ padding: "8px 12px", fontFamily: "'IBM Plex Mono'", fontSize: 12.5, borderLeft: `2px solid ${HR_T.indigo}` }}>{r.employeeId}</td>
                  <td style={{ padding: "8px 12px", fontFamily: "'IBM Plex Sans'", fontSize: 13 }}>{r.emp.name}</td>
                  <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: "'IBM Plex Mono'", fontSize: 13 }}>{r.present}</td>
                  <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: "'IBM Plex Mono'", fontSize: 13 }}>{r.weekend}</td>
                  <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: "'IBM Plex Mono'", fontSize: 13 }}>{r.leave}</td>
                  <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: "'IBM Plex Mono'", fontSize: 13, color: r.absent > 0 ? HR_T.bad : HR_T.ink }}>{r.absent}</td>
                  <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: "'IBM Plex Mono'", fontSize: 13, fontWeight: 600 }}>{r.present + r.weekend + r.leave + r.absent}</td>
                  <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: "'IBM Plex Mono'", fontSize: 13 }}>{r.present + r.weekend + r.leave}</td>
                  <td style={{ padding: "8px 12px", textAlign: "right", fontFamily: "'IBM Plex Mono'", fontSize: 13 }}>{r.otHours}</td>
                </tr>
              ))}
              {rowsForMonth.length === 0 && (
                <tr><td colSpan={9} style={{ padding: 36, textAlign: "center", fontFamily: "'IBM Plex Sans'", color: HR_T.muted, fontSize: 13 }}>No attendance records for this month.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </DocketCard>
    </div>
  );
}

/* =========================================================================
   LEAVE MANAGEMENT
========================================================================= */
function LeaveManagement({ employees, attendanceRecords }) {
  const [q, setQ] = useState("");

  const rows = useMemo(() => {
    return employees
      .filter((e) => !q || e.name.toLowerCase().includes(q.toLowerCase()) || e.employeeId.includes(q))
      .map((e) => {
        const recs = attendanceRecords.filter((r) => r.employeeId === e.employeeId);
        const totalUsed = recs.reduce((s, r) => s + num(r.leave), 0);
        const alloc = num(e.casualLeaveAlloc) + num(e.medicalLeaveAlloc);
        return { emp: e, totalUsed, alloc, remaining: alloc - totalUsed, months: recs.length };
      });
  }, [employees, attendanceRecords, q]);

  const monthOptions = useMemo(() => {
    return Array.from(new Set(attendanceRecords.map((r) => r.month).filter(Boolean)))
      .sort()
      .reverse();
  }, [attendanceRecords]);

  function exportLeaveExcel() {
    if (!window.XLSX) {
      window.alert("Excel export library is not available.");
      return;
    }

    const byEmployee = new Map(employees.map((e) => [e.employeeId, e]));

    /* Every imported attendance month becomes a separate column.
       The value is the Leave used in that month. */
    const monthlyColumns = monthOptions.map((month) => ({
      key: month,
      label: monthLabel(month),
    }));

    const exportRows = employees
      .filter((e) => !q || e.name.toLowerCase().includes(q.toLowerCase()) || e.employeeId.includes(q))
      .map((e) => {
        const recs = attendanceRecords.filter((r) => r.employeeId === e.employeeId);

        const row = {
          "Employee ID": e.employeeId,
          "Employee Name": e.name,
          "Casual Leave Allocation": num(e.casualLeaveAlloc),
          "Medical Leave Allocation": num(e.medicalLeaveAlloc),
          "Total Leave Allocation": num(e.casualLeaveAlloc) + num(e.medicalLeaveAlloc),
        };

        let totalUsed = 0;

        monthlyColumns.forEach(({ key, label }) => {
          const used = recs
            .filter((r) => r.month === key)
            .reduce((sum, r) => sum + num(r.leave), 0);

          row[label] = r2(used);
          totalUsed += used;
        });

        row["Total Leave Used"] = r2(totalUsed);
        row["Remaining Leave"] = r2(
          num(e.casualLeaveAlloc) + num(e.medicalLeaveAlloc) - totalUsed
        );

        return row;
      });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(exportRows);

    ws["!cols"] = [
      { wch: 16 },
      { wch: 28 },
      { wch: 22 },
      { wch: 23 },
      { wch: 22 },
      ...monthlyColumns.map(() => ({ wch: 16 })),
      { wch: 18 },
      { wch: 18 },
    ];

    XLSX.utils.book_append_sheet(wb, ws, "Leave Management");
    XLSX.writeFile(wb, `leave-management-${currentMonthKey()}.xlsx`);
  }

  return (
    <div>
      <DocketCard style={{ padding: 20, marginBottom: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12 }}>
          <div>
            <div style={{ fontFamily: "'Space Grotesk'", fontSize: 17, color: HR_T.ink, marginBottom: 4 }}>
              Leave Management
            </div>
            <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 12.5, color: HR_T.inkSoft }}>
              Track total leave used by employee and see exactly how many leave days were used in each month.
            </div>
          </div>

          <Btn variant="primary" onClick={exportLeaveExcel}>
            <Download size={15} /> Export Leave Excel
          </Btn>
        </div>
      </DocketCard>

      <div style={{ marginBottom: 14, maxWidth: 320 }}>
        <div style={{ position: "relative" }}>
          <Search size={14} color={HR_T.muted} style={{ position: "absolute", left: 11, top: 11 }} />
          <HrTInput
            placeholder="Search employee…"
            value={q}
            onChange={(e) => setQ(e.target.value)}
            style={{ paddingLeft: 32 }}
          />
        </div>
      </div>

      <DocketCard style={{ overflow: "hidden" }}>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 900 }}>
            <thead>
              <tr style={{ borderBottom: `2px solid ${HR_T.indigo}` }}>
                {[
                  "Employee ID",
                  "Name",
                  "Total Allocation",
                  ...monthOptions.map((m) => monthLabel(m)),
                  "Total Used",
                  "Remaining"
                ].map((h) => (
                  <th
                    key={h}
                    style={{
                      textAlign: h === "Name" || h === "Employee ID" ? "left" : "right",
                      padding: "9px 12px",
                      fontFamily: "'IBM Plex Sans'",
                      fontSize: 10.5,
                      fontWeight: 700,
                      textTransform: "uppercase",
                      color: HR_T.inkSoft,
                      whiteSpace: "nowrap"
                    }}
                  >
                    {h}
                  </th>
                ))}
              </tr>
            </thead>

            <tbody>
              {rows.map((r, i) => {
                const recs = attendanceRecords.filter((a) => a.employeeId === r.emp.employeeId);

                return (
                  <tr
                    key={r.emp.employeeId}
                    style={{
                      borderBottom: `1px solid ${HR_T.line}`,
                      background: i % 2 ? HR_T.canvas : HR_T.panel
                    }}
                  >
                    <td style={{ padding: "9px 12px", fontFamily: "'IBM Plex Mono'", fontSize: 12.5 }}>
                      {r.emp.employeeId}
                    </td>

                    <td style={{ padding: "9px 12px", fontFamily: "'IBM Plex Sans'", fontSize: 13 }}>
                      {r.emp.name}
                    </td>

                    <td style={{ padding: "9px 12px", textAlign: "right", fontFamily: "'IBM Plex Mono'", fontSize: 13 }}>
                      {r.alloc}
                    </td>

                    {monthOptions.map((month) => {
                      const used = recs
                        .filter((a) => a.month === month)
                        .reduce((sum, a) => sum + num(a.leave), 0);

                      return (
                        <td
                          key={month}
                          style={{
                            padding: "9px 12px",
                            textAlign: "right",
                            fontFamily: "'IBM Plex Mono'",
                            fontSize: 13
                          }}
                        >
                          {r2(used)}
                        </td>
                      );
                    })}

                    <td style={{ padding: "9px 12px", textAlign: "right", fontFamily: "'IBM Plex Mono'", fontSize: 13, fontWeight: 600 }}>
                      {r2(r.totalUsed)}
                    </td>

                    <td
                      style={{
                        padding: "9px 12px",
                        textAlign: "right",
                        fontFamily: "'IBM Plex Mono'",
                        fontSize: 13,
                        fontWeight: 600,
                        color: r.remaining < 0 ? HR_T.bad : HR_T.ink
                      }}
                    >
                      {r2(r.remaining)}
                    </td>
                  </tr>
                );
              })}

              {rows.length === 0 && (
                <tr>
                  <td
                    colSpan={4 + monthOptions.length}
                    style={{
                      padding: 36,
                      textAlign: "center",
                      fontFamily: "'IBM Plex Sans'",
                      color: HR_T.muted,
                      fontSize: 13
                    }}
                  >
                    No employees found.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </DocketCard>
    </div>
  );
}

/* =========================================================================
   PAYROLL MANAGEMENT + HISTORY
========================================================================= */
function PayrollManagement({ employees, attendanceRecords, payrollApprovals, onApprove, onReopen, onUpdateAttendance, rules, monthPicker, setMonthPicker }) {
  const monthsAvailable = useMemo(() => Array.from(new Set(attendanceRecords.map((r) => r.month))).sort().reverse(), [attendanceRecords]);
  const byId = useMemo(() => new Map(employees.map((e) => [e.employeeId, e])), [employees]);

  const rows = useMemo(() => {
    return attendanceRecords
      .filter((r) => r.month === monthPicker)
      .map((r) => {
        const emp = byId.get(r.employeeId);
        if (!emp) return null;
        return { emp, att: r, calc: computePayroll(emp, r, rules) };
      })
      .filter(Boolean);
  }, [attendanceRecords, monthPicker, byId, rules]);

  const isApproved = payrollApprovals[monthPicker]?.approved;
  const totals = rows.reduce((a, r) => ({
    gross: a.gross + r.calc.gross, pay: a.pay + r.calc.paySalary, ot: a.ot + r.calc.otAmount, net: a.net + r.calc.payAmount,
  }), { gross: 0, pay: 0, ot: 0, net: 0 });

  function updateManualAmount(employeeId, field, value) {
    if (isApproved) return;
    const numericValue = value === "" ? 0 : num(value);
    onUpdateAttendance(monthPicker, employeeId, field, numericValue);
  }

  function openPayslip(r) {
    if (!isApproved) {
      alert("Please approve & lock the payroll before issuing the payslip.");
      return;
    }

    const esc = (v) => String(v ?? "").replace(/[&<>"']/g, (ch) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
    }[ch]));
    const moneyHtml = (v) => `৳ ${hrMoney(v)}`;
    const month = monthLabel(monthPicker);
    const printedAt = new Date().toLocaleString("en-BD", { dateStyle: "medium", timeStyle: "short" });

    const html = `<!doctype html>
<html>
<head>
<meta charset="utf-8"/>
<title>Payslip - ${esc(r.emp.employeeId)} - ${esc(month)}</title>
<style>
  *{box-sizing:border-box} body{margin:0;background:#eee;font-family:Arial,Helvetica,sans-serif;color:#1e2433}
  .sheet{width:210mm;min-height:297mm;margin:18px auto;background:#fff;padding:18mm 16mm}
  .top{display:flex;justify-content:space-between;align-items:flex-start;border-bottom:2px solid #2c3e70;padding-bottom:14px}
  .brand{font-size:22px;font-weight:800;color:#2c3e70}.sub{font-size:11px;color:#777;margin-top:4px}
  .title{text-align:right}.title h1{margin:0;font-size:22px;color:#1e2c52}.title div{font-size:12px;color:#666;margin-top:5px}
  .info{display:grid;grid-template-columns:1fr 1fr;gap:7px 28px;margin:18px 0;padding:12px;background:#f7f5ef;border:1px solid #dad3bf}
  .info div{font-size:12px}.label{color:#777;display:inline-block;min-width:100px}.value{font-weight:700}
  .section-title{font-size:12px;font-weight:800;text-transform:uppercase;color:#2c3e70;margin:18px 0 7px}
  table{width:100%;border-collapse:collapse}th,td{padding:8px 9px;border-bottom:1px solid #ddd;font-size:12px}th{text-align:left;background:#f1ede4;color:#5b6478}td.amount{text-align:right;font-family:monospace}
  .net{margin-top:16px;border:2px solid #2c3e70;padding:12px 14px;display:flex;justify-content:space-between;align-items:center}.net span{font-weight:800;color:#2c3e70}.net strong{font-size:20px}
  .footer{margin-top:55px;display:grid;grid-template-columns:1fr 1fr;gap:50px;font-size:11px;color:#666}.sig{border-top:1px solid #555;padding-top:6px;text-align:center}
  .printbar{text-align:center;margin:12px}.printbar button{background:#2c3e70;color:#fff;border:0;padding:9px 18px;border-radius:5px;cursor:pointer}
  @media print{body{background:#fff}.sheet{margin:0;width:auto;min-height:auto;padding:12mm}.printbar{display:none}@page{size:A4;margin:0}}
</style>
</head>
<body>
<div class="printbar"><button onclick="window.print()">Print / Save as PDF</button></div>
<div class="sheet">
  <div class="top">
    <div><div class="brand">Hastizam Limited</div><div class="sub">Garments HR &amp; Payroll</div></div>
    <div class="title"><h1>SALARY PAYSLIP</h1><div>${esc(month)}</div></div>
  </div>

  <div class="info">
    <div><span class="label">Employee ID</span><span class="value">${esc(r.emp.employeeId)}</span></div>
    <div><span class="label">Employee Name</span><span class="value">${esc(r.emp.name)}</span></div>
    <div><span class="label">Section</span><span class="value">${esc(r.emp.section || "—")}</span></div>
    <div><span class="label">Job Title</span><span class="value">${esc(r.emp.jobTitle || "—")}</span></div>
    <div><span class="label">Joining Date</span><span class="value">${esc(r.emp.joiningDate || "—")}</span></div>
    <div><span class="label">Payable Days</span><span class="value">${esc(r.calc.payableDays)}</span></div>
  </div>

  <div class="section-title">Earnings</div>
  <table><thead><tr><th>Description</th><th style="text-align:right">Amount</th></tr></thead><tbody>
    <tr><td>Basic Salary</td><td class="amount">${moneyHtml(r.calc.basic)}</td></tr>
    <tr><td>House Rent</td><td class="amount">${moneyHtml(r.calc.houseRent)}</td></tr>
    <tr><td>Medical Allowance</td><td class="amount">${moneyHtml(r.calc.medical)}</td></tr>
    <tr><td>Conveyance</td><td class="amount">${moneyHtml(r.calc.conveyance)}</td></tr>
    <tr><td>Food Allowance</td><td class="amount">${moneyHtml(r.calc.food)}</td></tr>
    <tr><td>Gross Salary</td><td class="amount"><b>${moneyHtml(r.calc.gross)}</b></td></tr>
    <tr><td>Overtime</td><td class="amount">${moneyHtml(r.calc.otAmount)}</td></tr>
    <tr><td>Arrear</td><td class="amount">${moneyHtml(r.calc.arrear)}</td></tr>
  </tbody></table>

  <div class="section-title">Deductions</div>
  <table><thead><tr><th>Description</th><th style="text-align:right">Amount</th></tr></thead><tbody>
    <tr><td>Absent Deduction</td><td class="amount">${moneyHtml(r.calc.absentAmount)}</td></tr>
    <tr><td>Advance</td><td class="amount">${moneyHtml(r.calc.advance)}</td></tr>
    <tr><td>TDS Deduction</td><td class="amount">${moneyHtml(r.calc.tds)}</td></tr>
  </tbody></table>

  <div class="net"><span>NET PAYABLE SALARY</span><strong>${moneyHtml(r.calc.payAmount)}</strong></div>

  <div class="footer"><div class="sig">Employee Signature</div><div class="sig">Authorized Signature</div></div>
  <div style="margin-top:24px;text-align:center;font-size:10px;color:#999">Generated: ${esc(printedAt)} · Payroll Status: Approved</div>
</div>
</body></html>`;

    const win = window.open("", "_blank", "width=900,height=1000");
    if (!win) {
      alert("Popup blocked. Please allow popups for this HRM page and try again.");
      return;
    }
    win.document.open();
    win.document.write(html);
    win.document.close();
    win.focus();
  }

  function exportExcel() {
    const data = rows.map((r, i) => ({
      SL: i + 1, "Employee Status": r.emp.status, "Employee ID": r.emp.employeeId, Name: r.emp.name,
      Section: r.emp.section, "Job Title": r.emp.jobTitle, "Joining Date": r.emp.joiningDate,
      "Present Days": r.calc.present, "Weekend/Holidays": r.calc.weekend, Leave: r.calc.leave, "Absent Days": r.calc.absent,
      "Total Days": r.calc.totalDays, "Payable Days": r.calc.payableDays, Basic: r.calc.basic, "House Rent": r.calc.houseRent,
      Medical: r.calc.medical, Conveyance: r.calc.conveyance, "Food Allowance": r.calc.food, "Gross Salary": r.calc.gross,
      Overtime: r.calc.otAmount, "Advance": r.calc.advance, "Arrear": r.calc.arrear, "TDS Deduction": r.calc.tds, "Pay Amount Before TDS": r.calc.payBeforeTds, "Net Payable Salary": r.calc.payAmount,
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Payroll");
    XLSX.writeFile(wb, `payroll-${monthPicker || "export"}.xlsx`);
  }

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 16, flexWrap: "wrap" }}>
        <span style={{ fontFamily: "'IBM Plex Sans'", fontSize: 12.5, fontWeight: 600, color: HR_T.inkSoft }}>Payroll month</span>
        <HrTSelect value={monthPicker} onChange={(e) => setMonthPicker(e.target.value)} style={{ maxWidth: 200 }}>
          {monthsAvailable.length === 0 && <option value="">Import attendance first</option>}
          {monthsAvailable.map((m) => <option key={m} value={m}>{monthLabel(m)}</option>)}
        </HrTSelect>
        {isApproved ? (
          <Badge tone="good"><Lock size={11} style={{ marginRight: 4, verticalAlign: -1 }} />Locked & approved</Badge>
        ) : (
          <Badge tone="amber">Draft — pending review</Badge>
        )}
        <div style={{ flex: 1 }} />
        <Btn onClick={exportExcel}><Download size={14} /> Export Excel</Btn>
        {isApproved ? (
          <Btn onClick={() => onReopen(monthPicker)}><Unlock size={14} /> Reopen for correction</Btn>
        ) : (
          <Btn variant="primary" onClick={() => onApprove(monthPicker)} disabled={rows.length === 0}><CheckCircle2 size={14} /> Approve & lock payroll</Btn>
        )}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 14, marginBottom: 18 }}>
        <HrStatCard icon={Users} label="Employees" value={rows.length} />
        <HrStatCard icon={Wallet} label="Total Gross" value={`৳ ${hrMoney(totals.gross)}`} />
        <HrStatCard icon={TrendingUp} label="Total Overtime" value={`৳ ${hrMoney(totals.ot)}`} tone="amber" />
        <HrStatCard icon={CheckCircle2} label="Total Net Payable" value={`৳ ${hrMoney(totals.net)}`} tone="good" />
      </div>

      <DocketCard style={{ overflow: "hidden" }}>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 1650 }}>
            <thead>
              <tr style={{ borderBottom: `2px solid ${HR_T.indigo}` }}>
                {["Emp ID","Name","Section","Payable Days","Basic","House Rent","Gross","Absent Amt","OT Amt","Advance","Arrear","TDS Deduction","Pay Amount","Net Pay","Payslip"].map((h) => (
                  <th key={h} style={{ textAlign: ["Name","Emp ID","Section"].includes(h) ? "left" : "right", padding: "9px 11px", fontFamily: "'IBM Plex Sans'", fontSize: 10.5, fontWeight: 700, textTransform: "uppercase", color: HR_T.inkSoft, whiteSpace: "nowrap" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((r, i) => (
                <tr key={r.emp.employeeId} style={{ borderBottom: `1px solid ${HR_T.line}`, background: i % 2 ? HR_T.canvas : HR_T.panel }}>
                  <td style={{ padding: "8px 11px", fontFamily: "'IBM Plex Mono'", fontSize: 12, borderLeft: `2px solid ${HR_T.indigo}` }}>{r.emp.employeeId}</td>
                  <td style={{ padding: "8px 11px", fontFamily: "'IBM Plex Sans'", fontSize: 12.5 }}>{r.emp.name}</td>
                  <td style={{ padding: "8px 11px", fontFamily: "'IBM Plex Sans'", fontSize: 12, color: HR_T.inkSoft }}>{r.emp.section || "—"}</td>
                  <td style={{ padding: "8px 11px", textAlign: "right", fontFamily: "'IBM Plex Mono'", fontSize: 12.5 }}>{r.calc.payableDays}</td>
                  <td style={{ padding: "8px 11px", textAlign: "right", fontFamily: "'IBM Plex Mono'", fontSize: 12.5 }}>{hrMoney(r.calc.basic)}</td>
                  <td style={{ padding: "8px 11px", textAlign: "right", fontFamily: "'IBM Plex Mono'", fontSize: 12.5 }}>{hrMoney(r.calc.houseRent)}</td>
                  <td style={{ padding: "8px 11px", textAlign: "right", fontFamily: "'IBM Plex Mono'", fontSize: 12.5 }}>{hrMoney(r.calc.gross)}</td>
                  <td style={{ padding: "8px 11px", textAlign: "right", fontFamily: "'IBM Plex Mono'", fontSize: 12.5, color: HR_T.bad }}>{hrMoney(r.calc.absentAmount)}</td>
                  <td style={{ padding: "8px 11px", textAlign: "right", fontFamily: "'IBM Plex Mono'", fontSize: 12.5, color: HR_T.amber }}>{hrMoney(r.calc.otAmount)}</td>
                  <td style={{ padding: "6px 8px", textAlign: "right" }}>
                    <input type="number" min="0" step="0.01" value={r.att.advance ?? 0} disabled={isApproved} onChange={(e) => updateManualAmount(r.emp.employeeId, "advance", e.target.value)} style={{ width: 92, padding: "7px 8px", border: `1px solid ${isApproved ? HR_T.line : HR_T.indigo}`, borderRadius: 5, background: isApproved ? HR_T.canvas : "#fff", color: HR_T.ink, textAlign: "right", fontFamily: "'IBM Plex Mono'", fontSize: 12 }} />
                  </td>
                  <td style={{ padding: "6px 8px", textAlign: "right" }}>
                    <input type="number" step="0.01" value={r.att.arrear ?? 0} disabled={isApproved} onChange={(e) => updateManualAmount(r.emp.employeeId, "arrear", e.target.value)} style={{ width: 92, padding: "7px 8px", border: `1px solid ${isApproved ? HR_T.line : HR_T.indigo}`, borderRadius: 5, background: isApproved ? HR_T.canvas : "#fff", color: HR_T.ink, textAlign: "right", fontFamily: "'IBM Plex Mono'", fontSize: 12 }} />
                  </td>
                  <td style={{ padding: "6px 8px", textAlign: "right" }}>
                    <input type="number" min="0" step="0.01" value={r.att.tds ?? 0} disabled={isApproved} onChange={(e) => updateManualAmount(r.emp.employeeId, "tds", e.target.value)} style={{ width: 92, padding: "7px 8px", border: `1px solid ${isApproved ? HR_T.line : HR_T.indigo}`, borderRadius: 5, background: isApproved ? HR_T.canvas : "#fff", color: HR_T.ink, textAlign: "right", fontFamily: "'IBM Plex Mono'", fontSize: 12 }} />
                  </td>
                  <td style={{ padding: "8px 11px", textAlign: "right", fontFamily: "'IBM Plex Mono'", fontSize: 12.5 }}>{hrMoney(r.calc.payBeforeTds)}</td>
                  <td style={{ padding: "8px 11px", textAlign: "right", fontFamily: "'IBM Plex Mono'", fontSize: 13, fontWeight: 700, color: HR_T.indigo }}>{hrMoney(r.calc.payAmount)}</td>
                  <td style={{ padding: "6px 8px", textAlign: "center" }}>
                    <Btn size="sm" onClick={() => openPayslip(r)} disabled={!isApproved}>
                      <Printer size={13} /> Payslip
                    </Btn>
                  </td>
                </tr>
              ))}
              {rows.length === 0 && <tr><td colSpan={15} style={{ padding: 36, textAlign: "center", fontFamily: "'IBM Plex Sans'", color: HR_T.muted, fontSize: 13 }}>No attendance imported for this month yet — payroll needs attendance first.</td></tr>}
            </tbody>
          </table>
        </div>
      </DocketCard>
    </div>
  );
}

function PayrollHistory({ employees, attendanceRecords, payrollApprovals, rules }) {
  const byId = useMemo(() => new Map(employees.map((e) => [e.employeeId, e])), [employees]);
  const months = useMemo(() => Array.from(new Set(attendanceRecords.map((r) => r.month))).sort().reverse(), [attendanceRecords]);

  const byMonth = months.map((m) => {
    const recs = attendanceRecords.filter((r) => r.month === m);
    const net = recs.reduce((s, r) => { const emp = byId.get(r.employeeId); return emp ? s + computePayroll(emp, r, rules).payAmount : s; }, 0);
    return { month: m, count: recs.length, net, approved: !!payrollApprovals[m]?.approved };
  });

  return (
    <div>
      <ResponsiveContainer width="100%" height={220}>
        <LineChart data={[...byMonth].reverse()}>
          <CartesianGrid stroke={HR_T.line} vertical={false} />
          <XAxis dataKey="month" tickFormatter={monthLabel} tick={{ fontFamily: "IBM Plex Sans", fontSize: 11, fill: HR_T.inkSoft }} axisLine={{ stroke: HR_T.line }} tickLine={false} />
          <YAxis tick={{ fontFamily: "IBM Plex Mono", fontSize: 10, fill: HR_T.muted }} axisLine={false} tickLine={false} width={70} tickFormatter={(v) => `৳${(v/1000).toFixed(0)}k`} />
          <Tooltip formatter={(v) => `৳ ${hrMoney(v)}`} labelFormatter={monthLabel} contentStyle={{ fontFamily: "IBM Plex Sans", fontSize: 12, border: `1px solid ${HR_T.line}`, borderRadius: 8 }} />
          <Line type="monotone" dataKey="net" stroke={HR_T.indigo} strokeWidth={2.5} dot={{ r: 4, fill: HR_T.amber }} />
        </LineChart>
      </ResponsiveContainer>

      <DocketCard style={{ overflow: "hidden", marginTop: 18 }}>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr style={{ borderBottom: `2px solid ${HR_T.indigo}` }}>
                {["Month", "Employees", "Total Net Payable", "Status"].map((h) => (
                  <th key={h} style={{ textAlign: h === "Month" ? "left" : "right", padding: "9px 14px", fontFamily: "'IBM Plex Sans'", fontSize: 10.5, fontWeight: 700, textTransform: "uppercase", color: HR_T.inkSoft }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {byMonth.map((m, i) => (
                <tr key={m.month} style={{ borderBottom: `1px solid ${HR_T.line}`, background: i % 2 ? HR_T.canvas : HR_T.panel }}>
                  <td style={{ padding: "10px 14px", fontFamily: "'IBM Plex Sans'", fontSize: 13, fontWeight: 500, borderLeft: `2px solid ${HR_T.indigo}` }}>{monthLabel(m.month)}</td>
                  <td style={{ padding: "10px 14px", textAlign: "right", fontFamily: "'IBM Plex Mono'", fontSize: 13 }}>{m.count}</td>
                  <td style={{ padding: "10px 14px", textAlign: "right", fontFamily: "'IBM Plex Mono'", fontSize: 13.5, fontWeight: 700, color: HR_T.indigo }}>৳ {hrMoney(m.net)}</td>
                  <td style={{ padding: "10px 14px", textAlign: "right" }}>{m.approved ? <Badge tone="good">Locked</Badge> : <Badge tone="amber">Draft</Badge>}</td>
                </tr>
              ))}
              {byMonth.length === 0 && <tr><td colSpan={4} style={{ padding: 36, textAlign: "center", fontFamily: "'IBM Plex Sans'", color: HR_T.muted, fontSize: 13 }}>No payroll history yet.</td></tr>}
            </tbody>
          </table>
        </div>
      </DocketCard>
    </div>
  );
}

/* =========================================================================
   DASHBOARD
========================================================================= */
function HRDashboard({ employees, attendanceRecords, rules }) {
  const [dashMode, setDashMode] = useState("monthly"); // "monthly" | "yearly"
  const [dashMonth, setDashMonth] = useState(""); // month key or year, depending on dashMode
  const total = employees.length;
  const active = employees.filter((e) => ACTIVE_STATUSES.includes(e.status)).length;
  const newJoiners = employees.filter((e) => e.status === "New Join").length;
  const resigned = employees.filter((e) => e.status === "Resigned" || e.status === "Terminated").length;

  /* Dashboard period filter: Monthly = single-month stats, Yearly = aggregates across the whole year. */
  const dashMonths = useMemo(() => Array.from(new Set(attendanceRecords.map((r) => r.month).filter(Boolean))).sort().reverse(), [attendanceRecords]);
  const dashYears = useMemo(() => Array.from(new Set(dashMonths.map((m) => m.slice(0, 4)))).sort().reverse(), [dashMonths]);
  useEffect(() => {
    if (!dashMonths.length) return;
    const pool = dashMode === "yearly" ? dashYears : dashMonths;
    if (!pool.includes(dashMonth)) setDashMonth(pool[0] || "");
  }, [dashMonths, dashYears, dashMode, dashMonth]);
  function changeDashMode(mode) {
    if (mode === dashMode) return;
    if (!dashMonth) { setDashMode(mode); return; }
    const year = dashMonth.slice(0, 4);
    if (mode === "yearly") {
      setDashMode("yearly");
      setDashMonth(year);
    } else {
      const inYear = dashMonths.filter((m) => m.startsWith(year));
      setDashMode("monthly");
      setDashMonth(inYear[0] || dashMonths[0] || "");
    }
  }
  const scopeLabel = !dashMonth ? "—" : dashMode === "yearly" ? dashMonth : monthLabel(dashMonth);
  const scopedRecs = useMemo(() => {
    if (!dashMonth) return [];
    return dashMode === "yearly"
      ? attendanceRecords.filter((r) => String(r.month || "").startsWith(dashMonth))
      : attendanceRecords.filter((r) => r.month === dashMonth);
  }, [attendanceRecords, dashMode, dashMonth]);

  const byId = useMemo(() => new Map(employees.map((e) => [e.employeeId, e])), [employees]);
  const onLeaveCount = new Set(scopedRecs.filter((r) => num(r.leave) > 0).map((r) => r.employeeId)).size;
  const absentCount = dashMode === "yearly"
    ? scopedRecs.reduce((s, r) => s + num(r.absent), 0)
    : scopedRecs.filter((r) => num(r.absent) > 0).length;

  const totals = scopedRecs.reduce((a, r) => {
    const emp = byId.get(r.employeeId);
    if (!emp) return a;
    const c = computePayroll(emp, r, rules);
    return { payroll: a.payroll + c.payAmount, ot: a.ot + c.otAmount, leave: a.leave + num(r.leave) };
  }, { payroll: 0, ot: 0, leave: 0 });

  const bySection = useMemo(() => {
    const m = {};
    employees.forEach((e) => { const k = e.section || "Unassigned"; m[k] = (m[k] || 0) + 1; });
    return Object.entries(m).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [employees]);

  const statusPie = useMemo(() => {
    const m = {};
    employees.forEach((e) => { m[e.status] = (m[e.status] || 0) + 1; });
    return Object.entries(m).map(([name, value]) => ({ name, value }));
  }, [employees]);
  const pieColors = [HR_T.indigo, HR_T.amber, HR_T.good, HR_T.bad, "#7A88AE", "#D9AE6E", HR_T.muted];

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14, flexWrap: "wrap" }}>
        <span style={{ fontFamily: "'IBM Plex Sans'", fontSize: 11, fontWeight: 700, color: HR_T.muted, letterSpacing: ".05em", textTransform: "uppercase" }}>Dashboard period</span>
        {[["monthly", "Monthly"], ["yearly", "Yearly"]].map(([m, label]) => (
          <button key={m} onClick={() => changeDashMode(m)} style={{
            padding: "5px 14px", borderRadius: 20, cursor: "pointer",
            fontFamily: "'IBM Plex Sans'", fontSize: 12, fontWeight: 600,
            border: `1px solid ${dashMode === m ? HR_T.indigo : HR_T.line}`,
            background: dashMode === m ? HR_T.indigo : "#fff",
            color: dashMode === m ? "#fff" : HR_T.inkSoft,
          }}>{label}</button>
        ))}
        <HrTSelect value={dashMonth} onChange={(e) => setDashMonth(e.target.value)} style={{ maxWidth: 170 }}>
          {(dashMode === "yearly" ? dashYears : dashMonths).map((v) => (
            <option key={v} value={v}>{dashMode === "yearly" ? v : monthLabel(v)}</option>
          ))}
        </HrTSelect>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14, marginBottom: 18 }}>
        <HrStatCard icon={Users} label="Total Employees" value={total} sub={`${active} active`} />
        <HrStatCard icon={Package} label="New Joiners" value={newJoiners} tone="amber" />
        <HrStatCard icon={TrendingDown} label="Resigned / Terminated" value={resigned} tone="bad" />
        <HrStatCard icon={Umbrella} label={`On Leave (${scopeLabel})`} value={onLeaveCount} />
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14, marginBottom: 22 }}>
        <HrStatCard icon={AlertCircle} label={dashMode === "yearly" ? `Absent Days (${scopeLabel})` : `Absent (${scopeLabel})`} value={absentCount} tone="bad" />
        <HrStatCard icon={Wallet} label={`Payroll (${scopeLabel})`} value={`৳ ${hrMoney(totals.payroll)}`} tone="good" />
        <HrStatCard icon={TrendingUp} label={`Overtime (${scopeLabel})`} value={`৳ ${hrMoney(totals.ot)}`} tone="amber" />
        <HrStatCard icon={CalendarCheck} label={`Leave Days Used (${scopeLabel})`} value={totals.leave} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr", gap: 16 }}>
        <DocketCard style={{ padding: 20 }}>
          <div style={{ fontFamily: "'Space Grotesk'", fontSize: 16, fontWeight: 600, marginBottom: 12, color: HR_T.ink }}>Section-wise employee count</div>
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={bySection}>
              <CartesianGrid stroke={HR_T.line} vertical={false} />
              <XAxis dataKey="name" tick={{ fontFamily: "IBM Plex Sans", fontSize: 10.5, fill: HR_T.inkSoft }} axisLine={{ stroke: HR_T.line }} tickLine={false} interval={0} angle={-25} textAnchor="end" height={70} />
              <YAxis tick={{ fontFamily: "IBM Plex Mono", fontSize: 10, fill: HR_T.muted }} axisLine={false} tickLine={false} allowDecimals={false} />
              <Tooltip contentStyle={{ fontFamily: "IBM Plex Sans", fontSize: 12, border: `1px solid ${HR_T.line}`, borderRadius: 8 }} />
              <Bar dataKey="value" fill={HR_T.indigo} radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </DocketCard>
        <DocketCard style={{ padding: 20 }}>
          <div style={{ fontFamily: "'Space Grotesk'", fontSize: 16, fontWeight: 600, marginBottom: 12, color: HR_T.ink }}>By employee status</div>
          <ResponsiveContainer width="100%" height={200}>
            <PieChart>
              <Pie data={statusPie} dataKey="value" nameKey="name" innerRadius={46} outerRadius={76} paddingAngle={2}>
                {statusPie.map((_, i) => <Cell key={i} fill={pieColors[i % pieColors.length]} />)}
              </Pie>
              <Tooltip contentStyle={{ fontFamily: "IBM Plex Sans", fontSize: 12, border: `1px solid ${HR_T.line}`, borderRadius: 8 }} />
            </PieChart>
          </ResponsiveContainer>
          <div style={{ display: "flex", flexDirection: "column", gap: 5, marginTop: 6 }}>
            {statusPie.map((t, i) => (
              <div key={t.name} style={{ display: "flex", alignItems: "center", gap: 8, fontFamily: "'IBM Plex Sans'", fontSize: 12, color: HR_T.inkSoft }}>
                <span style={{ width: 9, height: 9, borderRadius: 9, background: pieColors[i % pieColors.length] }} />
                {t.name} — {t.value}
              </div>
            ))}
          </div>
        </DocketCard>
      </div>
    </div>
  );
}

/* =========================================================================
   REPORTS
========================================================================= */
function HRReports({ employees, attendanceRecords, rules }) {
  const byId = useMemo(() => new Map(employees.map((e) => [e.employeeId, e])), [employees]);
  const months = useMemo(() => Array.from(new Set(attendanceRecords.map((r) => r.month))).sort().reverse(), [attendanceRecords]);
  const [month, setMonth] = useState("");
  useEffect(() => { if (!month && months.length) setMonth(months[0]); }, [months, month]);

  function download(kind) {
    const wb = XLSX.utils.book_new();
    if (kind === "attendance") {
      const rows = attendanceRecords.filter((r) => r.month === month).map((r) => {
        const e = byId.get(r.employeeId);
        return { "Employee ID": r.employeeId, Name: e?.name, Section: e?.section, Present: r.present, "Weekend/Holiday": r.weekend, Leave: r.leave, Absent: r.absent, "OT Hours": r.otHours };
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "Attendance Summary");
      XLSX.writeFile(wb, `attendance-summary-${month}.xlsx`);
    } else if (kind === "leave") {
      const rows = employees.map((e) => {
        const used = attendanceRecords.filter((r) => r.employeeId === e.employeeId).reduce((s, r) => s + num(r.leave), 0);
        const alloc = num(e.casualLeaveAlloc) + num(e.medicalLeaveAlloc);
        return { "Employee ID": e.employeeId, Name: e.name, "Casual Allocation": e.casualLeaveAlloc, "Medical Allocation": e.medicalLeaveAlloc, "Total Used": used, Remaining: alloc - used };
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "Leave Summary");
      XLSX.writeFile(wb, "leave-summary.xlsx");
    } else if (kind === "sectionPayroll") {
      const recs = attendanceRecords.filter((r) => r.month === month);
      const bySection = {};
      recs.forEach((r) => {
        const e = byId.get(r.employeeId); if (!e) return;
        const c = computePayroll(e, r, rules);
        const k = e.section || "Unassigned";
        bySection[k] = (bySection[k] || 0) + c.payAmount;
      });
      const clean = Object.entries(bySection).map(([section, net]) => ({ Section: section, "Total Net Payable": r2(net) }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(clean), "Section Payroll");
      XLSX.writeFile(wb, `section-payroll-${month}.xlsx`);
    }
  }

  const reportCards = [
    { key: "attendance", title: "Attendance Summary", desc: "Present, weekend, leave, absent and OT hours per employee for the selected month.", icon: CalendarCheck },
    { key: "leave", title: "Leave Summary", desc: "Allocation, used and remaining leave for every employee, all-time.", icon: Umbrella },
    { key: "sectionPayroll", title: "Section-wise Payroll", desc: "Total net payable grouped by section for the selected month.", icon: Ruler },
  ];

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 18 }}>
        <span style={{ fontFamily: "'IBM Plex Sans'", fontSize: 12.5, fontWeight: 600, color: HR_T.inkSoft }}>Month (for monthly reports)</span>
        <HrTSelect value={month} onChange={(e) => setMonth(e.target.value)} style={{ maxWidth: 200 }}>
          {months.length === 0 && <option value="">No months imported</option>}
          {months.map((m) => <option key={m} value={m}>{monthLabel(m)}</option>)}
        </HrTSelect>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 16 }}>
        {reportCards.map((c) => (
          <DocketCard key={c.key} style={{ padding: 20, display: "flex", flexDirection: "column", gap: 12 }}>
            <c.icon size={20} color={HR_T.indigo} />
            <div style={{ fontFamily: "'Space Grotesk'", fontSize: 15.5, fontWeight: 600, color: HR_T.ink }}>{c.title}</div>
            <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 12.5, color: HR_T.inkSoft, lineHeight: 1.5, flex: 1 }}>{c.desc}</div>
            <Btn onClick={() => download(c.key)}><Download size={14} /> Download Excel</Btn>
          </DocketCard>
        ))}
      </div>
    </div>
  );
}

/* =========================================================================
   HR SETTINGS
========================================================================= */
function HRSettings({ rules, setRules, sections, jobTitles }) {
  const [local, setLocal] = useState(rules);
  return (
    <div>
      <DocketCard style={{ padding: 24, marginBottom: 18 }}>
        <div style={{ fontFamily: "'Space Grotesk'", fontSize: 17, marginBottom: 4, color: HR_T.ink }}>Payroll rule constants</div>
        <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 12.5, color: HR_T.inkSoft, marginBottom: 16 }}>
          These are the exact divisors read from your reference sheet. Change them here if your company's rules change — the calculation engine picks up the new values immediately, without any code changes.
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14 }}>
          <HrField label="Pay Salary days divisor"><HrTInput type="number" value={local.payDaysDivisor} onChange={(e) => setLocal((s) => ({ ...s, payDaysDivisor: num(e.target.value) }))} /></HrField>
          <HrField label="Absent amount divisor"><HrTInput type="number" value={local.absentDaysDivisor} onChange={(e) => setLocal((s) => ({ ...s, absentDaysDivisor: num(e.target.value) }))} /></HrField>
          <HrField label="OT rate divisor"><HrTInput type="number" value={local.otDivisor} onChange={(e) => setLocal((s) => ({ ...s, otDivisor: num(e.target.value) }))} /></HrField>
          <HrField label="Basic salary divisor"><HrTInput type="number" step="0.01" value={local.basicDivisor} onChange={(e) => setLocal((s) => ({ ...s, basicDivisor: num(e.target.value) }))} /></HrField>
        </div>
        <div style={{ marginTop: 18 }}>
          <Btn variant="primary" onClick={() => setRules(local)}><CheckCircle2 size={14} /> Save rules</Btn>
        </div>
      </DocketCard>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
        <DocketCard style={{ padding: 22 }}>
          <div style={{ fontFamily: "'Space Grotesk'", fontSize: 15.5, marginBottom: 10, color: HR_T.ink }}>Sections on file</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
            {sections.length === 0 && <span style={{ fontFamily: "'IBM Plex Sans'", fontSize: 12.5, color: HR_T.muted }}>None yet — sections appear here once employees are added.</span>}
            {sections.map((s) => <Badge key={s} tone="neutral">{s}</Badge>)}
          </div>
        </DocketCard>
        <DocketCard style={{ padding: 22 }}>
          <div style={{ fontFamily: "'Space Grotesk'", fontSize: 15.5, marginBottom: 10, color: HR_T.ink }}>Job titles on file</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
            {jobTitles.length === 0 && <span style={{ fontFamily: "'IBM Plex Sans'", fontSize: 12.5, color: HR_T.muted }}>None yet — job titles appear here once employees are added.</span>}
            {jobTitles.map((s) => <Badge key={s} tone="neutral">{s}</Badge>)}
          </div>
        </DocketCard>
      </div>

      <div style={{ marginTop: 16, background: HR_T.canvasDeep, border: `1px solid ${HR_T.line}`, borderRadius: 8, padding: 16, fontFamily: "'IBM Plex Sans'", fontSize: 12, color: HR_T.inkSoft, lineHeight: 1.6 }}>
        This prototype covers Employee Data, Attendance, Leave and Payroll end-to-end. Not built yet, but the Employee-ID-keyed data model here is designed to extend to: Overtime rules editor, role-based permissions (Super Admin / HR-Admin / HR User / Payroll User / Viewer), audit trail with created/updated-by, payslip generation, bank salary sheet export, and the other modules listed in your requirements doc — each can be added as a new tab reading the same employee and attendance records, without restructuring what's here.
      </div>
    </div>
  );
}


/* =========================================================================
   PROBATION & INCREMENT TRACKER
   New Join -> 3 Month Check -> 6 Month Probation Completion -> Regular
   1 Year -> Increment Due
========================================================================= */
function parseHRDate(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value)) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    // Excel serial date
    const excelEpoch = new Date(1899, 11, 30);
    const d = new Date(excelEpoch.getTime() + value * 86400000);
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  const s = String(value).trim();
  if (!s) return null;

  // YYYY-MM-DD / YYYY/MM/DD
  let m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (m) {
    const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    return isNaN(d) ? null : d;
  }

  // DD-MM-YYYY / DD/MM/YYYY
  m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
  if (m) {
    const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
    return isNaN(d) ? null : d;
  }

  const parsed = new Date(s);
  if (isNaN(parsed)) return null;
  return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

function addMonthsSafe(date, months) {
  const base = parseHRDate(date);
  if (!base) return null;
  const d = new Date(base.getFullYear(), base.getMonth(), base.getDate());
  const day = d.getDate();
  d.setDate(1);
  d.setMonth(d.getMonth() + months);
  const last = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
  d.setDate(Math.min(day, last));
  return d;
}
function daysBetween(a, b) {
  const x = parseHRDate(a); const y = parseHRDate(b);
  if (!x || !y) return null;
  x.setHours(0,0,0,0); y.setHours(0,0,0,0);
  return Math.round((y - x) / 86400000);
}
function milestoneInfo(joiningDate, months, today = new Date()) {
  const join = parseHRDate(joiningDate);
  if (!join) return null;
  const due = addMonthsSafe(join, months);
  if (!due) return null;
  const days = daysBetween(today, due);
  return { due, days, completed: days <= 0 };
}
function fmtISODate(d) {
  const dt = parseHRDate(d);
  if (!dt) return "";
  return `${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,"0")}-${String(dt.getDate()).padStart(2,"0")}`;
}
function milestoneLabel(info) {
  if (!info) return "—";
  if (info.days < 0) return `Completed ${Math.abs(info.days)} days ago`;
  if (info.days === 0) return "Due today";
  return `${info.days} days left`;
}

function ProbationIncrementTracker({ employees, onMakeRegular }) {
  const [q, setQ] = useState("");
  const [windowDays, setWindowDays] = useState(30);
  const today = useMemo(() => new Date(), []);

  const rows = useMemo(() => {
    return employees.map((e) => {
      const join = parseHRDate(e.joiningDate);
      if (!join) return null;
      const m3 = milestoneInfo(join, 3, today);
      const m6 = milestoneInfo(join, 6, today);
      const y1 = milestoneInfo(join, 12, today);
      return { emp: e, join, m3, m6, y1 };
    }).filter(Boolean);
  }, [employees, today]);

  const filtered = useMemo(() => {
    const query = q.trim().toLowerCase();
    return rows.filter(r => {
      if (query && ![r.emp.employeeId, r.emp.name, r.emp.section, r.emp.jobTitle, r.emp.status].some(v => (v || "").toString().toLowerCase().includes(query))) return false;
      return true;
    });
  }, [rows, q]);

  // Show everyone who has completed the milestone plus anyone due within the selected window.
  const threeMonth = filtered.filter(r => r.m3 && r.m3.days <= windowDays);
  const sixMonth = filtered.filter(r => r.m6 && r.m6.days <= windowDays);
  const oneYear = filtered.filter(r => r.y1 && r.y1.days <= windowDays);
  const tenDays = filtered.filter(r => r.y1 && r.y1.days >= 0 && r.y1.days <= 10);
  const probationComplete = filtered.filter(r => r.m6 && r.m6.days <= 0 && r.emp.status === "New Join");

  function exportExcel() {
    const milestoneStatus = (info) => {
      if (!info) return "—";
      if (info.completed) return "Completed";
      return `${info.days} days remaining`;
    };

    const data = filtered.map((r, i) => ({
      SL: i + 1,
      "Employee ID": r.emp.employeeId,
      "Employee Name": r.emp.name,
      Status: r.emp.status,
      Section: r.emp.section,
      "Job Title": r.emp.jobTitle,
      "Joining Date": r.emp.joiningDate,
      "3 Month": milestoneStatus(r.m3),
      "6 Month / Regular": milestoneStatus(r.m6),
      "1 Year / Increment": milestoneStatus(r.y1),
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Probation & Increment");
    XLSX.writeFile(wb, `probation-increment-${fmtISODate(today)}.xlsx`);
  }

  const Card = ({title, count, tone, children}) => (
    <DocketCard style={{ padding: 18 }}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
        <div style={{fontFamily:"'IBM Plex Sans'",fontSize:11,fontWeight:700,color:tone||HR_T.indigo,textTransform:"uppercase",letterSpacing:".05em"}}>{title}</div>
        <div style={{fontFamily:"'IBM Plex Mono'",fontSize:18,fontWeight:700,color:HR_T.ink}}>{count}</div>
      </div>
      {children}
    </DocketCard>
  );

  const miniRows = (arr, type) => arr.slice().sort((a,b) => (a[type]?.days ?? 99999) - (b[type]?.days ?? 99999)).slice(0,8).map(r => (
    <div key={r.emp.id} style={{display:"grid",gridTemplateColumns:"90px 1fr 120px 110px",gap:10,padding:"9px 0",borderBottom:`1px solid ${HR_T.line}`,fontSize:12}}>
      <span style={{fontFamily:"'IBM Plex Mono'"}}>{r.emp.employeeId}</span>
      <span>{r.emp.name}</span>
      <span style={{textAlign:"right",fontFamily:"'IBM Plex Mono'"}}>{hrFmtDate(r[type]?.due)}</span>
      <span style={{textAlign:"right",color:(r[type]?.days ?? 0) <= 10 && (r[type]?.days ?? 0) >= 0 ? HR_T.amber : HR_T.inkSoft}}>{milestoneLabel(r[type])}</span>
    </div>
  ));

  return (
    <div>
      <DocketCard style={{padding:18,marginBottom:16}}>
        <div style={{display:"flex",justifyContent:"space-between",gap:12,alignItems:"center",flexWrap:"wrap"}}>
          <div>
            <div style={{fontFamily:"'Space Grotesk'",fontSize:20,fontWeight:700,color:HR_T.ink}}>Probation & Increment Tracker</div>
            <div style={{fontFamily:"'IBM Plex Sans'",fontSize:12,color:HR_T.inkSoft,marginTop:4}}>Track 3-month review, 6-month probation completion / Regular status and 1-year increment dates.</div>
          </div>
          <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
            <HrTInput placeholder="Search ID / name…" value={q} onChange={e=>setQ(e.target.value)} />
            <HrTSelect value={windowDays} onChange={e=>setWindowDays(Number(e.target.value))}>
              <option value={30}>±30 days</option><option value={60}>±60 days</option><option value={90}>±90 days</option>
            </HrTSelect>
            <Btn variant="primary" onClick={exportExcel}><Download size={14}/> Export Excel</Btn>
          </div>
        </div>
      </DocketCard>

      <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:12,marginBottom:16}}>
        <Card title="3 Month Review" count={threeMonth.length}>{threeMonth.length ? miniRows(threeMonth,"m3") : <div style={{color:HR_T.muted,fontSize:12}}>No employees in this window.</div>}</Card>
        <Card title="6 Month Completion" count={sixMonth.length} tone={HR_T.amber}>{sixMonth.length ? miniRows(sixMonth,"m6") : <div style={{color:HR_T.muted,fontSize:12}}>No employees in this window.</div>}</Card>
        <Card title="Ready for Regular" count={probationComplete.length} tone={HR_T.good}>{probationComplete.length ? probationComplete.slice(0,8).map(r=><div key={r.emp.id} style={{padding:"9px 0",borderBottom:`1px solid ${HR_T.line}`,fontSize:12,display:"flex",justifyContent:"space-between",gap:8}}><span>{r.emp.employeeId} · {r.emp.name}</span><Btn variant="ghost" onClick={()=>onMakeRegular(r.emp)} style={{padding:"4px 7px",fontSize:10}}><CheckCircle2 size={12}/> Regular</Btn></div>) : <div style={{color:HR_T.muted,fontSize:12}}>No New Join employee has completed 6 months.</div>}</Card>
        <Card title="1 Year / Increment" count={oneYear.length}>{oneYear.length ? miniRows(oneYear,"y1") : <div style={{color:HR_T.muted,fontSize:12}}>No employees in this window.</div>}</Card>
        <Card title="10 Days Left" count={tenDays.length} tone={HR_T.bad}>{tenDays.length ? miniRows(tenDays,"y1") : <div style={{color:HR_T.muted,fontSize:12}}>No increment due within 10 days.</div>}</Card>
      </div>

      <DocketCard style={{padding:18,overflowX:"auto"}}>
        <div style={{fontFamily:"'IBM Plex Sans'",fontSize:11,fontWeight:700,color:HR_T.indigo,textTransform:"uppercase",marginBottom:12}}>Complete employee milestone list</div>
        <table style={{width:"100%",borderCollapse:"collapse",minWidth:1100}}>
          <thead><tr style={{borderBottom:`2px solid ${HR_T.indigo}`}}>
            {["Employee ID","Name","Status","Joining Date","3 Month","6 Month / Regular","1 Year / Increment","Days to 1 Year"].map((h)=><th key={h} style={{textAlign:h==="Name"?"left":"right",padding:"9px 10px",fontSize:10.5,color:HR_T.inkSoft,textTransform:"uppercase"}}>{h}</th>)}
          </tr></thead>
          <tbody>{filtered.sort((a,b)=>(a.y1?.days??99999)-(b.y1?.days??99999)).map(r=><tr key={r.emp.id} style={{borderBottom:`1px solid ${HR_T.line}`}}>
            <td style={{padding:"9px 10px",fontFamily:"'IBM Plex Mono'",fontSize:12}}>{r.emp.employeeId}</td>
            <td style={{padding:"9px 10px",fontSize:12.5}}>{r.emp.name}</td>
            <td style={{padding:"9px 10px",textAlign:"right",fontSize:12}}>{r.emp.status}</td>
            <td style={{padding:"9px 10px",textAlign:"right",fontSize:12}}>{hrFmtDate(r.emp.joiningDate)}</td>
            <td style={{padding:"9px 10px",textAlign:"right",fontFamily:"'IBM Plex Mono'",fontSize:12}}>{r.m3 ? `${hrFmtDate(r.m3.due)} · ${milestoneLabel(r.m3)}` : "—"}</td>
            <td style={{padding:"9px 10px",textAlign:"right",fontFamily:"'IBM Plex Mono'",fontSize:12,color:r.m6?.days<=0&&r.emp.status==="New Join"?HR_T.good:HR_T.ink}}>{r.m6 ? `${hrFmtDate(r.m6.due)} · ${milestoneLabel(r.m6)}` : "—"}</td>
            <td style={{padding:"9px 10px",textAlign:"right",fontFamily:"'IBM Plex Mono'",fontSize:12,color:r.y1?.days>=0&&r.y1?.days<=10?HR_T.bad:HR_T.ink}}>{r.y1 ? `${hrFmtDate(r.y1.due)} · ${milestoneLabel(r.y1)}` : "—"}</td>
            <td style={{padding:"9px 10px",textAlign:"right",fontFamily:"'IBM Plex Mono'",fontWeight:700,color:r.y1?.days>=0&&r.y1?.days<=10?HR_T.bad:HR_T.ink}}>{r.y1 ? `${r.y1.days} days` : "—"}</td>
          </tr>)}</tbody>
        </table>
        {filtered.length===0 && <div style={{padding:30,textAlign:"center",color:HR_T.muted,fontSize:12}}>No employees with a Joining Date found.</div>}
      </DocketCard>
    </div>
  );
}

/* =========================================================================
   LOGOUT — global top-bar button + styled confirm dialog.
   Rendered by the Workspace shell so it is available on every module.
   The button only opens the dialog; the session is cleared and the app
   returns to the login stage ONLY after the user confirms (onConfirm).
========================================================================= */
function LogoutButton({ onConfirm }) {
  const [confirmOpen, setConfirmOpen] = useState(false);

  return (
    <>
      <button
        type="button"
        className="hrm-header-logout-btn"
        onClick={() => setConfirmOpen(true)}
      >
        <LogOut size={15} /> Logout
      </button>

      {confirmOpen && (
        <LogoutConfirmDialog
          onCancel={() => setConfirmOpen(false)}
          onConfirm={() => {
            setConfirmOpen(false);
            onConfirm();
          }}
        />
      )}
    </>
  );
}

function LogoutConfirmDialog({ onCancel, onConfirm }) {
  useEffect(() => {
    function onKeyDown(e) {
      if (e.key === "Escape") onCancel();
    }
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [onCancel]);

  return (
    <div
      onClick={onCancel}
      style={{
        position: "fixed",
        inset: 0,
        zIndex: 1000,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        padding: 16,
        background: "rgba(20,27,48,.48)",
        fontFamily: "'IBM Plex Sans', sans-serif",
      }}
    >
      <div
        role="dialog"
        aria-modal="true"
        aria-label="Confirm logout"
        onClick={(e) => e.stopPropagation()}
        style={{
          width: "min(370px, 100%)",
          background: "#fff",
          border: `1px solid ${HR_T.line}`,
          borderRadius: 12,
          boxShadow: "0 28px 70px rgba(20,27,48,.35)",
          padding: "26px 26px 22px",
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div
            style={{
              width: 40,
              height: 40,
              borderRadius: 10,
              background: HR_T.amberSoft,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              flexShrink: 0,
            }}
          >
            <LogOut size={18} color={HR_T.amber} />
          </div>
          <div>
            <div
              style={{
                fontFamily: "'Space Grotesk', sans-serif",
                fontSize: 17,
                fontWeight: 700,
                color: HR_T.ink,
              }}
            >
              Confirm logout
            </div>
            <div style={{ fontSize: 11.5, color: HR_T.inkSoft, marginTop: 2 }}>
              Hastizam Limited · HR &amp; Payroll
            </div>
          </div>
        </div>

        <p
          style={{
            margin: "16px 0 20px",
            fontSize: 13,
            lineHeight: 1.6,
            color: HR_T.inkSoft,
          }}
        >
          Are you sure you want to logout? Your unsaved changes will remain on
          this device and you will need to sign in again.
        </p>

        <div style={{ display: "flex", justifyContent: "flex-end", gap: 9 }}>
          <button
            type="button"
            autoFocus
            onClick={onCancel}
            style={{
              padding: "8px 16px",
              borderRadius: 7,
              border: `1px solid ${HR_T.line}`,
              background: "#fff",
              color: HR_T.inkSoft,
              fontFamily: "'IBM Plex Sans', sans-serif",
              fontSize: 12.5,
              fontWeight: 600,
              cursor: "pointer",
            }}
          >
            No
          </button>
          <button
            type="button"
            onClick={onConfirm}
            style={{
              padding: "8px 18px",
              borderRadius: 7,
              border: "none",
              background: HR_T.amber,
              color: "#fff",
              fontFamily: "'IBM Plex Sans', sans-serif",
              fontSize: 12.5,
              fontWeight: 600,
              cursor: "pointer",
              display: "flex",
              alignItems: "center",
              gap: 7,
            }}
          >
            <LogOut size={13} /> Yes, Logout
          </button>
        </div>
      </div>
    </div>
  );
}

/* =========================================================================
   APP SHELL
======================================================================== */
const NAV = [
  { id: "dashboard", label: "HRDashboard", icon: LayoutDashboard },
  { id: "employees", label: "Employee Data", icon: Users },
  { id: "attendance", label: "Attendance", icon: CalendarCheck },
  { id: "leave", label: "Leave Management", icon: Umbrella },
  { id: "payroll", label: "Payroll", icon: Wallet },
  { id: "payrollHistory", label: "Payroll History", icon: History },
  { id: "probation", label: "Probation & Increment", icon: TrendingUp },
  { id: "reports", label: "HRReports", icon: FileBarChart },
  { id: "settings", label: "HR Settings", icon: SettingsIcon },
];

function HRApp() {
  const [tab, setTab] = useState("dashboard");
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [loginUser, setLoginUser] = useState("");
  const [loginPass, setLoginPass] = useState("");
  const [showLoginPass, setShowLoginPass] = useState(false);
  const [state, setState] = useState(null); // { employees, attendanceRecords, payrollApprovals, rules }
  const [profileEmployee, setProfileEmployee] = useState(null);
  const [editingEmployee, setEditingEmployee] = useState(null);
  const [showEmployeeForm, setShowEmployeeForm] = useState(false);
  const [saveStatus, setSaveStatus] = useState("idle");
  const [attMonth, setAttMonth] = useState("");
  const [payMonth, setPayMonth] = useState("");
  const [periodMode, setPeriodMode] = useState("monthly"); // "monthly" | "yearly"
  const [periodValue, setPeriodValue] = useState(""); // month key ("2026-09") or year ("2026")
  const fileInputRef = useRef(null);
  const [importBusy, setImportBusy] = useState(false);
  const [empImportResult, setEmpImportResult] = useState(null);

  const loadGenRef = useRef(0);
  const hrAppliedRef = useRef(false);
  useEffect(() => {
    const gen = ++loadGenRef.current;
    let cancelled = false;
    // Phase 1: instant paint from local cache (no waiting on the network)
    loadCachedState().then((cached) => {
      if (cancelled || gen !== loadGenRef.current || hrAppliedRef.current) return;
      hrAppliedRef.current = true;
      setState(cached ? { rules: DEFAULT_RULES, payrollApprovals: {}, ...cached } : { employees: [], attendanceRecords: [], payrollApprovals: {}, rules: DEFAULT_RULES });
    });
    // Phase 2: authoritative refresh from Supabase
    (async () => {
      try {
        const server = await fetchServerState();
        if (cancelled || gen !== loadGenRef.current) return;
        if (server) {
          hrAppliedRef.current = true;
          setState({ rules: DEFAULT_RULES, payrollApprovals: {}, ...server });
        }
      } catch (e) {
        console.error("HR state refresh failed — showing cached data:", e.message);
      }
    })();
    return () => { cancelled = true; };
  }, []);

  const persist = useCallback(async (next) => {
    loadGenRef.current++; // any in-flight server refresh is now stale
    hrAppliedRef.current = true;
    setState(next);
    setSaveStatus("saving");
    emitSync("saving");
    const ok = await saveState(next);
    setSaveStatus(ok ? "saved" : "error");
    emitSync(ok ? "saved" : "error");
  }, []);

  useEffect(() => {
    if (!state) return;
    const months = Array.from(new Set(state.attendanceRecords.map((r) => r.month))).sort().reverse();
    if (!attMonth && months[0]) setAttMonth(months[0]);
    if (!payMonth && months[0]) setPayMonth(months[0]);
  }, [state, attMonth, payMonth]);

  /* ---- Period filter: Monthly shows one month, Yearly aggregates the whole year.
     Drives every data view (dashboard, leave, history, reports) and the month
     pickers of the Attendance & Payroll tabs. ---- */
  const periodMonths = useMemo(() => Array.from(new Set((state ? state.attendanceRecords : []).map((r) => r.month).filter(Boolean))).sort().reverse(), [state]);
  const periodYears = useMemo(() => Array.from(new Set(periodMonths.map((m) => m.slice(0, 4)))).sort().reverse(), [periodMonths]);
  const visibleRecords = useMemo(() => {
    const all = state ? state.attendanceRecords : [];
    if (!periodValue) return all;
    return periodMode === "monthly" ? all.filter((r) => r.month === periodValue) : all.filter((r) => String(r.month || "").startsWith(periodValue));
  }, [state, periodMode, periodValue]);
  useEffect(() => {
    if (!periodMonths.length) return;
    const pool = periodMode === "yearly" ? periodYears : periodMonths;
    if (!pool.includes(periodValue)) setPeriodValue(pool[0] || "");
  }, [periodMonths, periodYears, periodMode, periodValue]);
  useEffect(() => {
    if (periodMode === "monthly" && periodValue) { setAttMonth(periodValue); setPayMonth(periodValue); }
  }, [periodMode, periodValue]);
  function changePeriodMode(mode) {
    if (mode === periodMode) return;
    if (!periodValue) { setPeriodMode(mode); return; }
    const year = periodValue.slice(0, 4);
    if (mode === "yearly") {
      setPeriodMode("yearly");
      setPeriodValue(year);
    } else {
      const inYear = periodMonths.filter((m) => m.startsWith(year));
      setPeriodMode("monthly");
      setPeriodValue(inYear[0] || periodMonths[0] || "");
    }
  }

  if (!state) {
    return <div style={{ padding: 40, fontFamily: "'IBM Plex Sans'", color: HR_T.muted }}>Loading HR system…</div>;
  }

  function handleSaveEmployee(emp) {
    const exists = state.employees.some((e) => e.id === emp.id);
    const employees = exists ? state.employees.map((e) => (e.id === emp.id ? emp : e)) : [...state.employees, emp];
    persist({ ...state, employees });
    setShowEmployeeForm(false); setEditingEmployee(null);
    if (profileEmployee && profileEmployee.id === emp.id) setProfileEmployee(emp);
    setTab("employees");
  }
  function handleMakeRegular(emp) {
    const updated = state.employees.map((e) => e.id === emp.id ? { ...e, status: "Regular" } : e);
    persist({ ...state, employees: updated });
  }
  function handleDeleteEmployee(id) {
    if (window.confirm("Delete this employee record? Their attendance and payroll history will remain but will no longer link to a name.")) {
      persist({ ...state, employees: state.employees.filter((e) => e.id !== id) });
    }
  }
  async function handleImportEmployees(file) {
    const { rows, errors } = await parseEmployeeExcel(file);
    setEmpImportResult({ summary: { total: rows.length + errors.length, imported: 0, duplicate: 0, notFound: 0, missing: errors.length }, errors });
    if (rows.length === 0) { window.alert("No valid employee rows found in that file."); return; }
    const existingIds = new Set(state.employees.map((e) => e.employeeId));
    const dupIds = rows.filter((r) => existingIds.has(r.employeeId)).map((r) => r.employeeId);
    let toAdd = rows;
    if (dupIds.length > 0) {
      const choice = window.confirm(`${dupIds.length} Employee ID(s) already exist (e.g. ${dupIds.slice(0, 5).join(", ")}). Click OK to UPDATE those existing records, or Cancel to SKIP them and only add new employees.`);
      if (choice) {
        const byId = new Map(state.employees.map((e) => [e.employeeId, e]));
        toAdd = rows.map((r) => {
          const existing = byId.get(r.employeeId);
          return existing ? { ...existing, ...r, id: existing.id } : { ...r, id: `emp-${Date.now()}-${r.employeeId}` };
        });
        const untouched = state.employees.filter((e) => !rows.some((r) => r.employeeId === e.employeeId));
        persist({ ...state, employees: [...untouched, ...toAdd] });
        setEmpImportResult({ summary: { total: rows.length, imported: rows.length, duplicate: dupIds.length, notFound: 0, missing: errors.length }, errors });
        return;
      } else {
        toAdd = rows.filter((r) => !existingIds.has(r.employeeId));
      }
    }
    const withIds = toAdd.map((r) => ({ ...r, id: `emp-${Date.now()}-${r.employeeId}` }));
    persist({ ...state, employees: [...state.employees, ...withIds] });
    setEmpImportResult({ summary: { total: rows.length, imported: withIds.length, duplicate: dupIds.length, notFound: 0, missing: errors.length }, errors });
  }
  function handleImportAttendance(month, records) {
    const others = state.attendanceRecords.filter((r) => r.month !== month);
    const withMonth = records.map((r) => ({ ...r, month }));
    const approvals = { ...state.payrollApprovals };
    delete approvals[month]; // re-importing reopens the month
    persist({ ...state, attendanceRecords: [...others, ...withMonth], payrollApprovals: approvals });
  }
  function handleUpdateAttendance(month, employeeId, field, value) {
    const updated = state.attendanceRecords.map((r) => {
      if (r.month !== month || r.employeeId !== employeeId) return r;
      return { ...r, [field]: value };
    });
    persist({ ...state, attendanceRecords: updated });
  }
  function handleApprovePayroll(month) {
    persist({ ...state, payrollApprovals: { ...state.payrollApprovals, [month]: { approved: true, approvedAt: new Date().toISOString() } } });
  }
  function handleReopenPayroll(month) {
    if (!window.confirm(`Reopen ${monthLabel(month)} payroll for correction? It will show as a draft until approved again.`)) return;
    const approvals = { ...state.payrollApprovals };
    delete approvals[month];
    persist({ ...state, payrollApprovals: approvals });
  }
  function setRules(rules) { persist({ ...state, rules }); }

  const attendanceByEmployee = (employeeId) => state.attendanceRecords.filter((r) => r.employeeId === employeeId).sort((a, b) => a.month.localeCompare(b.month));
  const payrollByEmployee = (employeeId) => attendanceByEmployee(employeeId).map((att) => {
    const emp = state.employees.find((e) => e.employeeId === employeeId);
    const calc = computePayroll(emp, att, state.rules);
    return { month: att.month, ...calc, approved: !!state.payrollApprovals[att.month]?.approved };
  });
  const leaveBalance = (emp) => {
    const recs = attendanceByEmployee(emp.employeeId);
    const totalUsed = recs.reduce((s, r) => s + num(r.leave), 0);
    // No type split available from the source data — see LeaveManagement note.
    const casualUsed = Math.min(totalUsed, num(emp.casualLeaveAlloc));
    const medicalUsed = totalUsed - casualUsed;
    return { casualAlloc: num(emp.casualLeaveAlloc), medicalAlloc: num(emp.medicalLeaveAlloc), casualUsed, medicalUsed };
  };

  const sections = Array.from(new Set(state.employees.map((e) => e.section).filter(Boolean))).sort();
  const jobTitles = Array.from(new Set(state.employees.map((e) => e.jobTitle).filter(Boolean))).sort();


  return (
    <div style={{ background: HR_T.canvas, minHeight: "100vh", width: "100vw", fontFamily: "'IBM Plex Sans', sans-serif" }}>
      <style>
        {HR_FONT_CSS}
        {`
        html, body, #root { width: 100%; min-height: 100%; margin: 0; padding: 0; }
        * { box-sizing: border-box; }
        html { overflow-x: hidden; }
        body { overflow-x: hidden; -webkit-tap-highlight-color: transparent; }
        button, input, select, textarea { -webkit-tap-highlight-color: transparent; }
        input:focus, select:focus { border-color: ${HR_T.indigo} !important; box-shadow: 0 0 0 3px ${HR_T.amberSoft}88; }
        input:disabled { color: ${HR_T.muted}; background: ${HR_T.canvasDeep}; }

        .hrm-shell { width: 100%; min-height: 100vh; }
        .hrm-sidebar {
          flex: 0 0 226px;
          z-index: 20;
        }
        .hrm-main {
          width: calc(100% - 226px);
          overflow: hidden;
        }
        .hrm-nav button {
          min-height: 38px;
          transition: background .18s ease, transform .18s ease;
        }
        .hrm-nav button:hover { background: rgba(255,255,255,.07) !important; }
        table { border-spacing: 0; }
        th, td { vertical-align: middle; }
        button { touch-action: manipulation; }
        @media (max-width: 900px) {
          .hrm-shell { display: block !important; }
          .hrm-sidebar {
            width: 100% !important;
            min-height: auto !important;
            height: auto !important;
            position: relative !important;
            top: auto !important;
            padding: 14px 12px 12px !important;
            box-shadow: 0 4px 18px rgba(20,27,48,.14);
          }
          .hrm-sidebar > div:first-child {
            margin-bottom: 12px !important;
            padding: 0 4px !important;
          }
          .hrm-nav {
            display: flex !important;
            flex-direction: row !important;
            gap: 6px !important;
            overflow-x: auto;
            overflow-y: hidden;
            padding: 2px 1px 6px;
            scrollbar-width: thin;
            -webkit-overflow-scrolling: touch;
          }
          .hrm-nav button {
            flex: 0 0 auto;
            white-space: nowrap;
            padding: 9px 11px !important;
            border-left: none !important;
            border-bottom: 2px solid transparent !important;
            border-radius: 8px !important;
          }
          .hrm-main {
            width: 100% !important;
            padding: 18px 14px 28px !important;
            overflow: visible !important;
          }
        }
        @media (max-width: 600px) {
          .hrm-main { padding: 14px 10px 24px !important; }
          .hrm-main > div:first-child {
            align-items: flex-start !important;
            gap: 12px !important;
            margin-bottom: 15px !important;
          }
          .hrm-main h1 {
            font-size: 21px !important;
            line-height: 1.15 !important;
          }
          .hrm-main > div:first-child > div:first-child > div {
            font-size: 9px !important;
          }
          .hrm-main button {
            min-height: 38px;
          }
          table {
            font-size: 12px;
          }
          th {
            font-size: 9.5px !important;
            padding: 8px 9px !important;
          }
          td {
            font-size: 12px !important;
            padding: 8px 9px !important;
          }
          input, select {
            min-height: 40px;
          }
        }
        @media (max-width: 430px) {
          .hrm-sidebar { padding: 11px 9px 8px !important; }
          .hrm-nav button { font-size: 11.5px !important; padding: 8px 10px !important; }
          .hrm-main { padding: 12px 8px 22px !important; }
          .hrm-main h1 { font-size: 19px !important; }
        }
        `}
      </style>

      <div className="hrm-shell" style={{ display: "flex" }}>
        <div className="hrm-sidebar" style={{ width: 226, minHeight: "100vh", background: HR_T.indigoDeep, padding: "24px 15px", position: "sticky", top: 0, alignSelf: "flex-start" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 9, marginBottom: 30, padding: "0 6px" }}>
            <div style={{ width: 32, height: 32, borderRadius: 6, background: HR_T.amber, display: "flex", alignItems: "center", justifyContent: "center" }}>
              <Scissors size={16} color="#fff" />
            </div>
            <div>
              <div style={{ fontFamily: "'Space Grotesk'", fontSize: 15, color: "#fff", fontWeight: 700, lineHeight: 1.1 }}>Hastizam Limited</div>
              <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 10, color: "#9AA6C4" }}>Garments HR &amp; Payroll</div>
            </div>
          </div>
          <nav className="hrm-nav" style={{ display: "flex", flexDirection: "column", gap: 2 }}>
            {NAV.map((n) => (
              <button key={n.id} onClick={() => { setTab(n.id); setProfileEmployee(null); setShowEmployeeForm(false); }} style={{
                display: "flex", alignItems: "center", gap: 10, padding: "9px 12px", borderRadius: 6,
                background: tab === n.id ? "rgba(255,255,255,.08)" : "transparent", border: "none",
                color: tab === n.id ? "#fff" : "#AEB8D4", fontFamily: "'IBM Plex Sans'", fontSize: 13, fontWeight: tab === n.id ? 600 : 500,
                cursor: "pointer", textAlign: "left", borderLeft: tab === n.id ? `2px solid ${HR_T.amber}` : "2px solid transparent",
              }}><n.icon size={15} /> {n.label}</button>
            ))}
          </nav>
        </div>

        <div className="hrm-main" style={{ flex: 1, padding: "26px 32px", minWidth: 0 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
            <div>
              <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 11, color: HR_T.muted, fontWeight: 600, letterSpacing: ".05em", textTransform: "uppercase" }}>Employee ID-linked HR ecosystem</div>
              <h1 style={{ fontFamily: "'Space Grotesk'", fontSize: 25, color: HR_T.ink, margin: "2px 0 0" }}>
                {profileEmployee ? "Employee Profile" : showEmployeeForm ? (editingEmployee ? "Edit Employee" : "New Employee") : NAV.find((n) => n.id === tab)?.label}
              </h1>
            </div>
            {tab === "employees" && !showEmployeeForm && !profileEmployee && (
              <Btn variant="primary" onClick={() => { setEditingEmployee(null); setShowEmployeeForm(true); }}><Plus size={15} /> New Employee</Btn>
            )}
          </div>

          {["dashboard", "attendance", "leave", "payroll", "payrollHistory", "reports"].includes(tab) && (
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 18, flexWrap: "wrap" }}>
              <span style={{ fontFamily: "'IBM Plex Sans'", fontSize: 11, fontWeight: 700, color: HR_T.muted, letterSpacing: ".05em", textTransform: "uppercase" }}>View by</span>
              {[["monthly", "Monthly"], ["yearly", "Yearly"]].map(([m, label]) => (
                <button key={m} onClick={() => changePeriodMode(m)} style={{
                  padding: "6px 16px", borderRadius: 20, cursor: "pointer",
                  fontFamily: "'IBM Plex Sans'", fontSize: 12.5, fontWeight: 600,
                  border: `1px solid ${periodMode === m ? HR_T.indigo : HR_T.line}`,
                  background: periodMode === m ? HR_T.indigo : "#fff",
                  color: periodMode === m ? "#fff" : HR_T.inkSoft,
                }}>{label}</button>
              ))}
              <HrTSelect value={periodValue} onChange={(e) => setPeriodValue(e.target.value)} style={{ maxWidth: 170 }}>
                {(periodMode === "monthly" ? periodMonths : periodYears).map((v) => (
                  <option key={v} value={v}>{periodMode === "monthly" ? monthLabel(v) : v}</option>
                ))}
              </HrTSelect>
            </div>
          )}

          {tab === "dashboard" && <HRDashboard employees={state.employees} attendanceRecords={visibleRecords} rules={state.rules} />}

          {tab === "employees" && !showEmployeeForm && !profileEmployee && (
            <>
              <EmployeesTable
                employees={state.employees}
                onEdit={(e) => { setEditingEmployee(e); setShowEmployeeForm(true); }}
                onDelete={handleDeleteEmployee}
                onOpenProfile={(e) => setProfileEmployee(e)}
                onImport={handleImportEmployees}
                fileInputRef={fileInputRef}
                importBusy={importBusy}
                setImportBusy={setImportBusy}
              />
              <ImportSummaryBox summary={empImportResult?.summary} errors={empImportResult?.errors} />
            </>
          )}
          {tab === "employees" && showEmployeeForm && (
            <EmployeeForm employees={state.employees} initial={editingEmployee} onSave={handleSaveEmployee} onCancel={() => { setShowEmployeeForm(false); setEditingEmployee(null); }} />
          )}
          {tab === "employees" && profileEmployee && (
            <EmployeeProfile
              employee={profileEmployee}
              attendanceByEmployee={attendanceByEmployee}
              payrollByEmployee={payrollByEmployee}
              leaveBalance={leaveBalance}
              onBack={() => setProfileEmployee(null)}
              onEdit={(e) => { setEditingEmployee(e); setShowEmployeeForm(true); setProfileEmployee(null); }}
            />
          )}

          {tab === "attendance" && (
            <AttendanceManagement employees={state.employees} attendanceRecords={state.attendanceRecords} onImport={handleImportAttendance} monthPicker={attMonth} setMonthPicker={setAttMonth} />
          )}
          {tab === "leave" && <LeaveManagement employees={state.employees} attendanceRecords={visibleRecords} />}
          {tab === "payroll" && (
            <PayrollManagement
              employees={state.employees} attendanceRecords={state.attendanceRecords} payrollApprovals={state.payrollApprovals}
              onApprove={handleApprovePayroll} onUpdateAttendance={handleUpdateAttendance} onReopen={handleReopenPayroll} rules={state.rules}
              monthPicker={payMonth} setMonthPicker={setPayMonth}
            />
          )}
          {tab === "payrollHistory" && <PayrollHistory employees={state.employees} attendanceRecords={visibleRecords} payrollApprovals={state.payrollApprovals} rules={state.rules} />}
          {tab === "probation" && <ProbationIncrementTracker employees={state.employees} onMakeRegular={handleMakeRegular} />}
          {tab === "reports" && <HRReports employees={state.employees} attendanceRecords={visibleRecords} rules={state.rules} />}
          {tab === "settings" && <HRSettings rules={state.rules} setRules={setRules} sections={sections} jobTitles={jobTitles} />}
        </div>
      </div>
    </div>
  );
}
function WelcomePage({ onContinue }) {
  return (
    <div className="hr-landing" style={{
      minHeight: "100vh", width: "100%", display: "flex", alignItems: "stretch",
      background: T.paper, fontFamily: "'IBM Plex Sans', sans-serif",
    }}>
      <div style={{
        flex: "0 0 42%", minWidth: 320, background: T.ink, display: "flex",
        flexDirection: "column", justifyContent: "center", padding: "60px 56px",
        position: "relative", overflow: "hidden",
      }}>
        <div style={{ position: "absolute", left: 0, top: 0, bottom: 0, width: 4, background: T.rule }} />
        <div style={{
          width: 52, height: 52, borderRadius: 10, background: T.rule, display: "flex",
          alignItems: "center", justifyContent: "center", fontFamily: "'Source Serif 4'",
          color: "#fff", fontWeight: 700, fontSize: 26, marginBottom: 28,
        }}>H</div>
        <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 12, fontWeight: 600, letterSpacing: ".08em", textTransform: "uppercase", color: "#8FA39D", marginBottom: 10 }}>
          Hastizam Limited
        </div>
        <h1 style={{ fontFamily: "'Source Serif 4'", fontSize: 40, lineHeight: 1.15, color: "#fff", margin: "0 0 18px" }}>
          Hastizam Ledger
        </h1>
        <p style={{ fontFamily: "'IBM Plex Sans'", fontSize: 14.5, lineHeight: 1.7, color: "#B7C6C2", maxWidth: 400, margin: 0 }}>
          A single home for TDS, VDS, and vendor payment records — master data entry,
          live withholding calculations, filterable transactions, and reports, all
          in one place.
        </p>
        <div style={{ display: "flex", gap: 22, marginTop: 40 }}>
          {["TDS", "VDS", "Payments"].map((t) => (
            <div key={t} style={{ fontFamily: "'IBM Plex Mono'", fontSize: 11.5, color: "#8FA39D", letterSpacing: ".05em", borderTop: "1px solid rgba(255,255,255,.15)", paddingTop: 8 }}>
              {t}
            </div>
          ))}
        </div>
      </div>

      <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", padding: 40 }}>
        <div style={{ maxWidth: 380, textAlign: "left" }}>
          <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 11.5, fontWeight: 600, letterSpacing: ".05em", textTransform: "uppercase", color: T.rule, marginBottom: 10 }}>
            Welcome to Hastizam HR
          </div>
          <h2 style={{ fontFamily: "'Source Serif 4'", fontSize: 28, color: T.ink, margin: "0 0 14px" }}>
            Manage your people with confidence
          </h2>
          <p style={{ fontFamily: "'IBM Plex Sans'", fontSize: 14, lineHeight: 1.7, color: T.inkSoft, margin: "0 0 30px" }}>
            Employee records, attendance, leave, payroll and history — designed to work together without changing your existing HR workflow.
          </p>
          <button onClick={onContinue} style={{
            display: "inline-flex", alignItems: "center", gap: 10, padding: "13px 28px",
            borderRadius: 8, border: "none", background: T.accent, color: "#fff",
            fontFamily: "'IBM Plex Sans'", fontSize: 14.5, fontWeight: 600, cursor: "pointer",
          }}>
            Get Started <ArrowRight size={16} />
          </button>
          <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 11.5, color: T.muted, marginTop: 18 }}>
            Hastizam Ledger · TDS / VDS accounting prototype
          </div>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------
   LOGIN PAGE — ledger-style login (currently unused; App uses
   TopHRLoginPage below). Wired to POST /api/auth.
--------------------------------------------------------------------- */
function LoginPage({ onLogin, onBack }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [remember, setRemember] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  async function submit(e) {
  e.preventDefault();

  if (loading) return;

  if (!email.trim() || !password) {
    setError("Please enter username/email and password.");
    return;
  }

  setLoading(true);
  setError("");

  try {
    const response = await fetch("/api/auth", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        email: email.trim(),
        password,
      }),
    });

    const data = await response.json().catch(() => ({}));

    if (!response.ok || data.success !== true) {
      setError(data.message || "Invalid username/email or password.");
      return;
    }

    // IMPORTANT:
    // Dashboard ONLY opens after successful authentication
    onLogin(data.user, data.token);

  } catch (err) {
    console.error("AUTH ERROR:", err);
    setError("Unable to communicate with authentication server.");
  } finally {
    setLoading(false);
  }
}

  return (
    <div
      className="hr-login"
      style={{
        minHeight: "100vh",
        width: "100%",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        background: T.paper,
        fontFamily: "'IBM Plex Sans', sans-serif",
        padding: 24,
        boxSizing: "border-box",
      }}
    >
      <form
        onSubmit={submit}
        style={{
          width: "100%",
          maxWidth: 380,
          background: T.card,
          border: `1px solid ${T.line}`,
          borderRadius: 14,
          padding: "38px 34px",
          position: "relative",
          overflow: "hidden",
          boxSizing: "border-box",
          boxShadow: "0 18px 45px rgba(0,0,0,0.08)",
        }}
      >
        {/* Top line */}
        <div
          style={{
            position: "absolute",
            left: 0,
            top: 0,
            right: 0,
            height: 3,
            background: T.rule,
          }}
        />

        {/* Logo / Brand */}
        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: 10,
            marginBottom: 26,
          }}
        >
          <div
            style={{
              width: 34,
              height: 34,
              borderRadius: 8,
              background: T.ink,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontFamily: "'Source Serif 4'",
              color: "#fff",
              fontWeight: 700,
              fontSize: 17,
              flexShrink: 0,
            }}
          >
            H
          </div>

          <div>
            <div
              style={{
                fontFamily: "'Source Serif 4'",
                fontSize: 16,
                fontWeight: 600,
                color: T.ink,
                lineHeight: 1.1,
              }}
            >
              Hastizam HR Management
            </div>

            <div
              style={{
                fontFamily: "'IBM Plex Sans'",
                fontSize: 10.5,
                color: T.muted,
                marginTop: 3,
              }}
            >
              Secure sign in to continue
            </div>
          </div>
        </div>

        {/* Email */}
        <Field label="Username or Email">
          <TInput
            type="email"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            placeholder="you@hastizam.com"
            autoComplete="username"
            disabled={loading}
          />
        </Field>

        <div style={{ height: 14 }} />

        {/* Password */}
        <Field label="Password">
          <TInput
            type="password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            placeholder="••••••••"
            autoComplete="current-password"
            disabled={loading}
          />
        </Field>

        {/* Error Message */}
        {error && (
  <div
    style={{
      marginTop: 10,
      marginBottom: 10,
      padding: "10px 12px",
      borderRadius: 8,
      background: "#fef2f2",
      border: "1px solid #fecaca",
      color: "#dc2626",
      fontSize: 12,
    }}
  >
    {error}
  </div>
)}

        {/* Remember / Forgot */}
        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            margin: "16px 0 22px",
          }}
        >
          <label
            style={{
              display: "flex",
              alignItems: "center",
              gap: 7,
              fontFamily: "'IBM Plex Sans'",
              fontSize: 12.5,
              color: T.inkSoft,
              cursor: "pointer",
            }}
          >
            <input
              type="checkbox"
              checked={remember}
              onChange={(e) => setRemember(e.target.checked)}
              disabled={loading}
              style={{
                accentColor: T.accent,
                width: 14,
                height: 14,
              }}
            />

            Remember me
          </label>

          <a
            href="#"
            onClick={(e) => e.preventDefault()}
            style={{
              fontFamily: "'IBM Plex Sans'",
              fontSize: 12.5,
              color: T.rule,
              textDecoration: "none",
              fontWeight: 600,
            }}
          >
            Forgot password?
          </a>
        </div>

        {/* Login Button */}
        <button
          type="submit"
          disabled={loading}
          style={{
            width: "100%",
            padding: "12px 0",
            borderRadius: 8,
            border: "none",
            background: loading ? "#94a3b8" : T.accent,
            color: "#fff",
            fontFamily: "'IBM Plex Sans'",
            fontSize: 14.5,
            fontWeight: 600,
            cursor: loading ? "not-allowed" : "pointer",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            gap: 8,
            opacity: loading ? 0.8 : 1,
          }}
        >
          {loading ? (
            <>
              <span
                style={{
                  width: 14,
                  height: 14,
                  border: "2px solid rgba(255,255,255,0.4)",
                  borderTopColor: "#fff",
                  borderRadius: "50%",
                  display: "inline-block",
                  animation: "hrmLoginSpin 0.7s linear infinite",
                }}
              />

              Signing in...
            </>
          ) : (
            <>
              <LogIn size={16} />
              Login
            </>
          )}
        </button>

        {/* Back */}
        <div
          style={{
            textAlign: "center",
            marginTop: 18,
          }}
        >
          <button
            type="button"
            onClick={onBack}
            disabled={loading}
            style={{
              background: "none",
              border: "none",
              cursor: loading ? "not-allowed" : "pointer",
              fontFamily: "'IBM Plex Sans'",
              fontSize: 12,
              color: T.muted,
            }}
          >
            ← Back
          </button>
        </div>

        {/* Security note */}
        <div
          style={{
            marginTop: 18,
            paddingTop: 16,
            borderTop: `1px dashed ${T.line}`,
            fontFamily: "'IBM Plex Sans'",
            fontSize: 10.5,
            color: T.muted,
            textAlign: "center",
            lineHeight: 1.5,
          }}
        >
          Your credentials are verified securely against the
          HRM database.
        </div>
      </form>

      <style>
        {`
          @keyframes hrmLoginSpin {
            from {
              transform: rotate(0deg);
            }
            to {
              transform: rotate(360deg);
            }
          }

          @media (max-width: 480px) {
            .hr-login {
              padding: 14px !important;
            }

            .hr-login form {
              padding: 30px 22px !important;
            }
          }
        `}
      </style>
    </div>
  );
}

/* ---------------------------------------------------------------------
   APP SHELL (the existing dashboard/entry/transactions/reports app —
   unchanged, just renamed so it can be mounted after Welcome → Login)
--------------------------------------------------------------------- */
function LedgerApp() {
  const [tab, setTab] = useState("dashboard");
  const [records, setRecords] = useState(null); // null = loading
  const [editing, setEditing] = useState(null);
  const [showForm, setShowForm] = useState(false);
  const [isDefaultData, setIsDefaultData] = useState(true); // true = still showing sample data, never saved
  const [saveStatus, setSaveStatus] = useState("idle"); // idle | saving | saved | error
  const [lastSavedAt, setLastSavedAt] = useState(null);

  const ledgerLoadGenRef = useRef(0);
  const ledgerAppliedRef = useRef(false);
  useEffect(() => {
    const gen = ++ledgerLoadGenRef.current;
    let cancelled = false;
    // Phase 1: instant paint from local cache
    loadCachedLedger(STORAGE_KEY).then((cached) => {
      if (cancelled || gen !== ledgerLoadGenRef.current || ledgerAppliedRef.current) return;
      ledgerAppliedRef.current = true;
      if (cached !== null) {
        setRecords(cached);
        setIsDefaultData(false);
      } else {
        setRecords([]);
        setIsDefaultData(true);
      }
    });
    // Phase 2: authoritative refresh from Supabase
    (async () => {
      try {
        const server = await fetchServerLedger();
        if (cancelled || gen !== ledgerLoadGenRef.current || server === null) return;
        ledgerAppliedRef.current = true;
        setRecords(server);
        setIsDefaultData(false);
      } catch (e) {
        console.error("Ledger refresh failed — showing cached data:", e.message);
      }
    })();
    return () => { cancelled = true; };
  }, []);

  const persist = useCallback(async (next) => {
    ledgerLoadGenRef.current++; // invalidate in-flight refresh
    ledgerAppliedRef.current = true;
    setRecords(next);
    setIsDefaultData(false);
    setSaveStatus("saving");
    emitSync("saving");
    const ok = await saveRecords(next);
    setSaveStatus(ok ? "saved" : "error");
    emitSync(ok ? "saved" : "error");
    if (ok) setLastSavedAt(new Date());
  }, []);

  function handleSave(rec) {
    const withCalc = { ...rec };
    if (records.some((r) => r.id === rec.id)) {
      persist(records.map((r) => (r.id === rec.id ? withCalc : r)));
    } else {
      persist([...records, withCalc]);
    }
    setShowForm(false);
    setEditing(null);
    setTab("transactions");
  }
  function handleDelete(id) {
    if (window.confirm("Delete this ledger entry? This cannot be undone.")) {
      persist(records.filter((r) => r.id !== id));
    }
  }
  function handleImport(imported) {
    persist(imported);
  }
  const nav = [
    { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
    { id: "entry", label: "Master Data", icon: BookOpen },
    { id: "transactions", label: "Transactions", icon: ListFilter },
    { id: "reports", label: "Reports", icon: FileBarChart },
  ];

  if (records === null) {
    return (
      <div
        style={{ padding: 40, fontFamily: "'IBM Plex Sans'", color: T.muted }}
      >
        Loading ledger…
      </div>
    );
  }

  return (
    <div
      className="ledger-shell"
      style={{
        background: T.paper,
        minHeight: "100vh",
        width: "100vw",
        fontFamily: "'IBM Plex Sans', sans-serif",
      }}
    >
      <style>
        {FONT_CSS}
        {`
      html, body, #root {
      width: 100%;
      min-height: 100%;
      margin: 0;
      padding: 0;
      }
        * { box-sizing: border-box; }
        input:focus, select:focus { border-color: ${T.accent} !important; box-shadow: 0 0 0 3px ${T.ruleSoft}55; }
        @media print {
          .no-print { display: none !important; }
          body { background: white !important; }
        }
      `}
      </style>

      <div className="ledger-shell" style={{ display: "flex", width: "100%", minHeight: "100vh" }}>
        {/* Sidebar */}
        <div
          className="no-print ledger-sidebar"
          style={{
            width: 220,
            minHeight: "100vh",
            background: T.ink,
            padding: "26px 16px",
            position: "sticky",
            top: 0,
            alignSelf: "flex-start",
          }}
        >
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 8,
              marginBottom: 34,
              padding: "0 6px",
            }}
          >
            <div
              style={{
                width: 30,
                height: 30,
                borderRadius: 6,
                background: T.rule,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontFamily: "'Source Serif 4'",
                color: "#fff",
                fontWeight: 700,
                fontSize: 16,
              }}
            >
              H
            </div>
            <div>
              <div
                style={{
                  fontFamily: "'Source Serif 4'",
                  fontSize: 15.5,
                  color: "#fff",
                  fontWeight: 600,
                  lineHeight: 1.1,
                }}
              >
                Hastizam Ledger
              </div>
              <div
                style={{
                  fontFamily: "'IBM Plex Sans'",
                  fontSize: 10.5,
                  color: "#8FA39D",
                }}
              >
                TDS · VDS · Payments
              </div>
            </div>
          </div>
          <nav style={{ display: "flex", flexDirection: "column", gap: 3 }}>
            {nav.map((n) => (
              <button
                key={n.id}
                onClick={() => {
                  setTab(n.id);
                  setShowForm(false);
                }}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 10,
                  padding: "10px 12px",
                  borderRadius: 7,
                  background:
                    tab === n.id ? "rgba(255,255,255,.08)" : "transparent",
                  border: "none",
                  color: tab === n.id ? "#fff" : "#9FB2AC",
                  fontFamily: "'IBM Plex Sans'",
                  fontSize: 13.5,
                  fontWeight: tab === n.id ? 600 : 500,
                  cursor: "pointer",
                  textAlign: "left",
                  borderLeft:
                    tab === n.id
                      ? `2px solid ${T.rule}`
                      : "2px solid transparent",
                }}
              >
                <n.icon size={16} /> {n.label}
              </button>
            ))}
          </nav>
          
        </div>

        {/* Main */}
        <div
          className="ledger-main"
          style={{
            flex: 1,
            padding: "28px 34px",
            minWidth: 0,
            width: "100%",
          }}
        >
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              marginBottom: 22,
            }}
          >
            <div>
              <div
                style={{
                  fontFamily: "'IBM Plex Sans'",
                  fontSize: 11.5,
                  color: T.muted,
                  fontWeight: 600,
                  letterSpacing: ".05em",
                  textTransform: "uppercase",
                }}
              >
                Hastizam Limited
              </div>
              <h1
                style={{
                  fontFamily: "'Source Serif 4'",
                  fontSize: 27,
                  color: T.ink,
                  margin: "2px 0 0",
                }}
              >
                {nav.find((n) => n.id === tab)?.label}
              </h1>
            </div>
            {tab !== "entry" && (
              <button
                className="no-print"
                onClick={() => {
                  setEditing(null);
                  setTab("entry");
                }}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 8,
                  padding: "10px 18px",
                  borderRadius: 8,
                  border: "none",
                  background: T.accent,
                  color: "#fff",
                  fontFamily: "'IBM Plex Sans'",
                  fontSize: 13.5,
                  fontWeight: 600,
                  cursor: "pointer",
                }}
              >
                <Plus size={15} /> New Entry
              </button>
            )}
          </div>

          {tab === "dashboard" && <Dashboard records={records} />}
          {tab === "transactions" && (
            <Transactions
              records={records}
              onEdit={(r) => {
                setEditing(r);
                setTab("entry");
              }}
              onDelete={handleDelete}
              onImport={handleImport}
            />
          )}
          {tab === "reports" && <Reports records={records} />}
          {tab === "entry" && (
            <EntryForm
              records={records}
              initial={editing}
              onSave={handleSave}
              onCancel={() => {
                setEditing(null);
                setTab("transactions");
              }}
            />
          )}
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------
   WORKSPACE SWITCHER — thin top bar over the two systems. Defaults to
   the ledger (unchanged first-open experience); switching to HR does
   not affect the ledger's state, and vice versa.
--------------------------------------------------------------------- */
/* ---------------------------------------------------------------------
   WORKSPACE — GLOBAL LAYOUT SHELL. The thin top bar is shared by every
   module (Accounting Ledger, HR Management, and anything added to the
   switcher below), so the authenticated-user chip and the Logout button
   here are visible across the whole application. Any future module
   registered in this bar automatically gets global logout for free.
--------------------------------------------------------------------- */
function Workspace({ onLogout }) {
  const { user } = useAuth();
  const [system, setSystem] = useState("ledger"); // ledger | hr
  const initials = user?.name
    ? user.name.split(/\s+/).filter(Boolean).map((w) => w[0]).slice(0, 2).join("").toUpperCase()
    : "";
  return (
    <div>
      <div className="no-print hrm-topbar" style={{
        display: "flex", alignItems: "center", gap: 8, padding: "10px 20px",
        background: T.accentDeep, borderBottom: `1px solid ${T.line}`,
      }}>
        <span className="hrm-topbar-brand-label" style={{ fontFamily: "'IBM Plex Sans'", fontSize: 11.5, color: "#CFE3DA", fontWeight: 600, marginRight: 6 }}>Hastizam Systems</span>
        {[["ledger", "Accounting Ledger"], ["hr", "HR Management"]].map(([id, label]) => (
          <button key={id} onClick={() => setSystem(id)} style={{
            padding: "6px 14px", borderRadius: 20, border: "none", cursor: "pointer",
            fontFamily: "'IBM Plex Sans'", fontSize: 12.5, fontWeight: 600,
            background: system === id ? "#fff" : "transparent",
            color: system === id ? T.accentDeep : "#CFE3DA",
          }}>
            {label}
          </button>
        ))}

        {/* GLOBAL IDENTITY + LOGOUT — top-right, on every module */}
        <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 12 }}>
          {user && (
            <>
              <div
                title={user.name}
                style={{
                  width: 28, height: 28, borderRadius: "50%", background: HR_T.amber,
                  color: "#fff", fontFamily: "'IBM Plex Sans'", fontSize: 11, fontWeight: 700,
                  display: "flex", alignItems: "center", justifyContent: "center",
                }}
              >
                {initials || "?"}
              </div>
              <div className="hrm-topbar-user-meta">
                <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 12, fontWeight: 600, color: "#fff", lineHeight: 1.15 }}>
                  {user.name}
                </div>
                <div style={{ fontFamily: "'IBM Plex Sans'", fontSize: 10, color: "#CFE3DA" }}>
                  {user.role}
                </div>
              </div>
            </>
          )}
          {onLogout && <LogoutButton onConfirm={onLogout} />}
        </div>
      </div>
      {system === "ledger" ? <LedgerApp /> : <HRApp />}
    </div>
  );
}

/* ---------------------------------------------------------------------
   TOP-LEVEL APP — drives Login → Workspace with a fade/slide
   transition between stages. The workspace is only reachable after
   /api/auth confirms the credentials.
--------------------------------------------------------------------- */
const STAGE_TRANSITION_CSS = `
@keyframes stageFadeIn {
  from { opacity: 0; transform: translateY(14px); }
  to   { opacity: 1; transform: translateY(0); }
}
@keyframes stageFadeOut {
  from { opacity: 1; transform: translateY(0); }
  to   { opacity: 0; transform: translateY(-10px); }
}
.stage-enter { animation: stageFadeIn 420ms ease both; }
.stage-leave { animation: stageFadeOut 260ms ease both; }
`;

/* -------------------------------------------------------------------------
   GLOBAL LAYOUT CSS — top-bar identity + logout (visible on every module)
   and the post-logout toast. Used by the Workspace shell.
------------------------------------------------------------------------- */
const GLOBAL_LAYOUT_CSS = `
@keyframes hrmToastIn {
  from { opacity: 0; transform: translateY(10px); }
  to   { opacity: 1; transform: translateY(0); }
}
.hrm-header-logout-btn {
  display: flex;
  align-items: center;
  gap: 7px;
  padding: 6px 14px;
  border-radius: 20px;
  background: transparent;
  border: 1px solid rgba(255,255,255,.35);
  color: #E8EDF7;
  font-family: 'IBM Plex Sans';
  font-size: 12.5px;
  font-weight: 600;
  cursor: pointer;
  white-space: nowrap;
  transition: background .18s ease, color .18s ease, border-color .18s ease;
}
.hrm-header-logout-btn:hover { background: #fff; border-color: #fff; color: ${HR_T.indigoDeep}; }
.hrm-toast {
  position: fixed;
  right: 22px;
  bottom: 22px;
  z-index: 1200;
  display: flex;
  align-items: center;
  gap: 9px;
  max-width: calc(100vw - 44px);
  padding: 11px 16px;
  background: #fff;
  border: 1px solid ${HR_T.line};
  border-left: 3px solid ${HR_T.good};
  border-radius: 9px;
  box-shadow: 0 14px 40px rgba(20,27,48,.22);
  animation: hrmToastIn 260ms ease both;
}
@media (max-width: 900px) {
  .hrm-topbar { flex-wrap: wrap; row-gap: 8px; padding: 10px 14px !important; }
}
@media (max-width: 600px) {
  .hrm-topbar { padding: 8px 12px !important; gap: 6px; }
  .hrm-topbar-user-meta { display: none; }
  .hrm-topbar-brand-label { display: none; }
  .hrm-toast { left: 16px; right: 16px; bottom: 14px; }
}
`;

/* =========================================================================
   TOP-LEVEL LOGIN GATE — the HR Management login screen (copied verbatim
   from HRApp's own login gate), used as the app's first screen instead
   of the ledger's Welcome/Login pages. HRApp's internal login screen is
   untouched below; this is a separate copy so opening the HR tab later
   doesn't ask to log in a second time.
========================================================================= */
function TopHRLoginPage({ onLogin }) {
  const [loginUser, setLoginUser] = useState("");
  const [loginPass, setLoginPass] = useState("");
  const [showLoginPass, setShowLoginPass] = useState(false);
  const [loading, setLoading] = useState(false);
  const [loginError, setLoginError] = useState("");

  async function handleLogin() {
    if (loading) return;

    if (!loginUser.trim() || !loginPass) {
      setLoginError("Please enter username/email and password.");
      return;
    }

    try {
      setLoading(true);
      setLoginError("");

      const response = await fetch("/api/auth", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          email: loginUser.trim(),
          password: loginPass,
        }),
      });

      let data;

      try {
        data = await response.json();
      } catch {
        data = {};
      }

      if (!response.ok || !data.success) {
        setLoginError(
          data.message || "Invalid username/email or password."
        );
        return;
      }

      // IMPORTANT:
      // Dashboard will open ONLY after successful authentication.
      onLogin(data.user, data.token);

    } catch (error) {
      console.error("Login error:", error);
      setLoginError(
        "Unable to connect to authentication server."
      );
    } finally {
      setLoading(false);
    }
  }

  function submitLogin(e) {
    e.preventDefault();
    handleLogin();
  }

  return (
    <div className="hrm-login-page">
      <style>{`
        .hrm-login-page {
          min-height: 100vh;
          width: 100%;
          display: flex;
          align-items: center;
          justify-content: center;
          padding: 24px;
          background:
            radial-gradient(circle at 15% 15%, rgba(231,166,42,.16), transparent 32%),
            radial-gradient(circle at 85% 85%, rgba(38,58,112,.20), transparent 35%),
            #f4f1e9;
          font-family: 'IBM Plex Sans', sans-serif;
        }

        .hrm-login-shell {
          width: min(960px, 100%);
          min-height: 560px;
          display: grid;
          grid-template-columns: 1.08fr .92fr;
          overflow: hidden;
          border-radius: 22px;
          background: #fff;
          border: 1px solid rgba(38,58,112,.14);
          box-shadow: 0 28px 80px rgba(28,39,68,.18);
        }

        .hrm-login-brand {
          position: relative;
          padding: 48px;
          color: #fff;
          background:
            linear-gradient(145deg, #17264e 0%, #263a70 62%, #304982 100%);
          display: flex;
          flex-direction: column;
          justify-content: space-between;
          overflow: hidden;
        }

        .hrm-login-brand:before,
        .hrm-login-brand:after {
          content: "";
          position: absolute;
          border-radius: 50%;
          pointer-events: none;
        }

        .hrm-login-brand:before {
          width: 300px;
          height: 300px;
          right: -130px;
          top: -130px;
          background: rgba(231,166,42,.12);
        }

        .hrm-login-brand:after {
          width: 240px;
          height: 240px;
          left: -110px;
          bottom: -120px;
          border: 1px solid rgba(255,255,255,.08);
        }

        .hrm-logo-mark {
          width: 48px;
          height: 48px;
          border-radius: 13px;
          background: #e7a62a;
          display: flex;
          align-items: center;
          justify-content: center;
          box-shadow: 0 10px 28px rgba(231,166,42,.25);
        }

        .hrm-login-title {
          font-family: 'Space Grotesk', sans-serif;
          font-size: clamp(34px, 5vw, 52px);
          line-height: .98;
          letter-spacing: -.04em;
          margin: 22px 0 14px;
        }

        .hrm-login-copy {
          max-width: 450px;
          color: #cbd4eb;
          line-height: 1.7;
          font-size: 14px;
        }

        .hrm-login-features {
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 10px;
          margin-top: 28px;
        }

        .hrm-login-feature {
          padding: 12px;
          border: 1px solid rgba(255,255,255,.09);
          background: rgba(255,255,255,.045);
          border-radius: 10px;
          color: #dce3f4;
          font-size: 11px;
        }

        .hrm-login-form {
          padding: 48px;
          display: flex;
          flex-direction: column;
          justify-content: center;
        }

        .hrm-login-form h2 {
          font-family: 'Space Grotesk', sans-serif;
          color: #202b43;
          font-size: 27px;
          margin: 0 0 7px;
        }

        .hrm-login-form .sub {
          color: #7a8499;
          font-size: 12px;
          margin-bottom: 27px;
        }

        .hrm-login-label {
          display: block;
          color: #46516a;
          font-size: 11px;
          font-weight: 700;
          margin: 0 0 7px;
        }

        .hrm-login-input-wrap {
          position: relative;
          margin-bottom: 16px;
        }

        .hrm-login-input {
          width: 100%;
          height: 46px;
          border: 1px solid #d7dce6;
          border-radius: 9px;
          padding: 0 42px 0 40px;
          outline: none;
          font: 13px 'IBM Plex Sans', sans-serif;
          color: #202b43;
          background: #fbfbfa;
          box-sizing: border-box;
        }

        .hrm-login-input:focus {
          border-color: #263a70;
          box-shadow: 0 0 0 3px rgba(231,166,42,.16);
        }

        .hrm-login-input:disabled {
          opacity: .65;
        }

        .hrm-login-icon {
          position: absolute;
          left: 13px;
          top: 15px;
          color: #8b95a8;
        }

        .hrm-login-eye {
          position: absolute;
          right: 7px;
          top: 6px;
          width: 34px;
          height: 34px;
          padding: 0;
          border: 0;
          background: transparent;
          color: #7f899d;
          display: flex;
          align-items: center;
          justify-content: center;
          cursor: pointer;
        }

        .hrm-login-button {
          width: 100%;
          height: 46px;
          margin-top: 5px;
          border: 0;
          border-radius: 9px;
          background: #263a70;
          color: #fff;
          font: 700 13px 'IBM Plex Sans', sans-serif;
          cursor: pointer;
          box-shadow: 0 10px 24px rgba(38,58,112,.18);
          transition: transform .15s ease, background .15s ease;
        }

        .hrm-login-button:hover:not(:disabled) {
          background: #304982;
          transform: translateY(-1px);
        }

        .hrm-login-button:disabled {
          opacity: .65;
          cursor: not-allowed;
          transform: none;
        }

        .hrm-login-error {
          margin: -4px 0 12px;
          padding: 10px 12px;
          border-radius: 8px;
          background: #fef2f2;
          border: 1px solid #fecaca;
          color: #dc2626;
          font-size: 11.5px;
          line-height: 1.45;
        }

        .hrm-demo-note {
          margin-top: 15px;
          padding: 10px 12px;
          border-radius: 8px;
          background: #f6f2e8;
          color: #7b6b48;
          font-size: 10.5px;
          line-height: 1.5;
        }

        .hrm-login-footer {
          margin-top: 25px;
          color: #a0a8b7;
          font-size: 10px;
          text-align: center;
        }

        @media (max-width: 760px) {
          .hrm-login-page {
            padding: 12px;
          }

          .hrm-login-shell {
            grid-template-columns: 1fr;
            min-height: auto;
            max-width: 500px;
          }

          .hrm-login-brand {
            min-height: 280px;
            padding: 28px 24px;
          }

          .hrm-login-title {
            font-size: 36px;
          }

          .hrm-login-features {
            grid-template-columns: 1fr 1fr;
          }

          .hrm-login-form {
            padding: 30px 24px;
          }
        }

        @media (max-width: 420px) {
          .hrm-login-brand {
            padding: 23px 18px;
          }

          .hrm-login-form {
            padding: 26px 18px;
          }

          .hrm-login-features {
            grid-template-columns: 1fr;
          }

          .hrm-login-title {
            font-size: 32px;
          }
        }
      `}</style>

      <div className="hrm-login-shell">

        {/* LEFT SIDE */}
        <div className="hrm-login-brand">

          <div>

            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 12,
                position: "relative",
                zIndex: 1,
              }}
            >
              <div className="hrm-logo-mark">
                <UsersRound size={23} color="#fff" />
              </div>

              <div>
                <div
                  style={{
                    fontFamily: "'Space Grotesk'",
                    fontSize: 16,
                    fontWeight: 800,
                    letterSpacing: ".01em",
                  }}
                >
                  Hastizam Limited
                </div>

                <div
                  style={{
                    color: "#aebbd7",
                    fontSize: 9.5,
                    marginTop: 2,
                  }}
                >
                  Garments HR & Payroll
                </div>
              </div>
            </div>

            <div
              style={{
                position: "relative",
                zIndex: 1,
              }}
            >
              <div className="hrm-login-title">
                Human
                <br />
                Resource
                <br />
                <span style={{ color: "#e7a62a" }}>
                  Management
                </span>
              </div>

              <div className="hrm-login-copy">
                A clean workspace for employee data, attendance,
                leave, payroll, payslips and workforce milestones.
              </div>

              <div className="hrm-login-features">
                <div className="hrm-login-feature">
                  Employee Management
                </div>

                <div className="hrm-login-feature">
                  Attendance & Leave
                </div>

                <div className="hrm-login-feature">
                  Payroll & Payslip
                </div>

                <div className="hrm-login-feature">
                  Probation & Increment
                </div>
              </div>
            </div>

          </div>

          <div
            style={{
              position: "relative",
              zIndex: 1,
              color: "#8f9dbc",
              fontSize: 10,
            }}
          >
            © {new Date().getFullYear()} Hastizam Limited · HR & Payroll Workspace
          </div>

        </div>

        {/* RIGHT SIDE */}
        <div className="hrm-login-form">

          <div
            style={{
              marginBottom: 8,
              color: "#e7a62a",
              fontSize: 10,
              fontWeight: 800,
              letterSpacing: ".08em",
              textTransform: "uppercase",
            }}
          >
            Secure Workspace
          </div>

          <h2>Welcome back</h2>

          <div className="sub">
            Sign in to continue to your HRM dashboard.
          </div>

          <label className="hrm-login-label">
            Username
          </label>

          <div className="hrm-login-input-wrap">

            <UserRound
              size={15}
              className="hrm-login-icon"
            />

            <input
              className="hrm-login-input"
              value={loginUser}
              onChange={(e) => {
                setLoginUser(e.target.value);
                setLoginError("");
              }}
              placeholder="Enter username"
              autoComplete="username"
              disabled={loading}
            />

          </div>

          <label className="hrm-login-label">
            Password
          </label>

          <div className="hrm-login-input-wrap">

            <LockKeyhole
              size={15}
              className="hrm-login-icon"
            />

            <input
              className="hrm-login-input"
              type={showLoginPass ? "text" : "password"}
              value={loginPass}
              onChange={(e) => {
                setLoginPass(e.target.value);
                setLoginError("");
              }}
              placeholder="Enter password"
              autoComplete="current-password"
              disabled={loading}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  handleLogin();
                }
              }}
            />

            <button
              className="hrm-login-eye"
              type="button"
              onClick={() =>
                setShowLoginPass((v) => !v)
              }
              aria-label="Toggle password visibility"
            >
              {showLoginPass ? (
                <EyeOff size={15} />
              ) : (
                <Eye size={15} />
              )}
            </button>

          </div>

          {/* LOGIN ERROR */}
          {loginError && (
            <div className="hrm-login-error">
              {loginError}
            </div>
          )}

          <form onSubmit={submitLogin}>
            <button
  type="submit"
  className="hrm-login-button"
  disabled={loading}
>
  {loading ? "Signing in..." : "Sign In"}
</button>
          </form>

          <div className="hrm-demo-note">
            <strong>Authorized users only.</strong>{" "}
            Please use your registered HRM account to continue.
          </div>

          <div className="hrm-login-footer">
            Designed for a clean, responsive HR & Payroll experience
          </div>

        </div>

      </div>
    </div>
  );
}

/* -------------------------------------------------------------------------
   SESSION PERSISTENCE — stores ONLY the sanitized authenticated user
   returned by /api/auth (id, name, email, role, status).
   Never stores email/password or any credential material.
   sessionStorage survives page refreshes but ends when the tab closes.
------------------------------------------------------------------------- */
const HRM_SESSION_KEY = "hrm_session_user";

function readHrmSessionUser() {
  try {
    const raw = window.sessionStorage.getItem(HRM_SESSION_KEY);
    if (!raw) return null;
    const user = JSON.parse(raw);
    if (user && user.id && user.email) return user;
    return null;
  } catch {
    return null;
  }
}

function saveHrmSessionUser(user) {
  try {
    window.sessionStorage.setItem(
      HRM_SESSION_KEY,
      JSON.stringify({
        id: user.id,
        name: user.name,
        email: user.email,
        role: user.role,
        status: user.status,
      })
    );
  } catch (e) {
    console.error("Session save failed:", e);
  }
}

function clearHrmSessionUser() {
  try {
    window.sessionStorage.removeItem(HRM_SESSION_KEY);
  } catch (e) {
    console.error("Session clear failed:", e);
  }
}

/* Clears ONLY authentication/session material. NEVER wipe all of
   localStorage here — the Workspace persists live HR business data
   (employees, attendance, payroll) in localStorage under its own keys,
   and no auth cookies exist in this architecture. */
function clearSessionData() {
  clearHrmSessionUser();
  try { window.sessionStorage.removeItem(HRM_SESSION_TOKEN_KEY); } catch {}
}

/* -------------------------------------------------------------------------
   AUTH CONTEXT — single source of truth for the authenticated user.
   login(user)/logout() keep sessionStorage and React state in sync.
------------------------------------------------------------------------- */
const AuthContext = createContext(null);

function AuthProvider({ children }) {
  const [user, setUser] = useState(() => readHrmSessionUser());

  function login(nextUser, nextToken) {
    if (!nextUser || !nextUser.id || !nextUser.email) return false;
    if (!nextToken) return false; // never accept a session without its token
    saveHrmSessionUser(nextUser);
    try { window.sessionStorage.setItem(HRM_SESSION_TOKEN_KEY, nextToken); } catch (e) { console.error("Token save failed:", e); }
    setUser(nextUser);
    return true;
  }

  function logout() {
    clearSessionData();
    setUser(null);
  }

  const value = useMemo(() => ({ user, login, logout }), [user]);
  return <AuthContext.Provider value={value}>{children}</AuthContext.Provider>;
}

function useAuth() {
  const ctx = useContext(AuthContext);
  if (!ctx) throw new Error("useAuth must be used inside <AuthProvider>");
  return ctx;
}

/* ---------- On-screen sync indicator ---------- */
function SyncChip() {
  const status = useSyncStatus();
  if (status === "idle") return null;
  const map = {
    saving: { border: HR_T.line, color: HR_T.inkSoft, dot: HR_T.amber, label: "Saving to Supabase…" },
    saved: { border: "#BFDCC8", color: HR_T.good, dot: HR_T.good, label: "All changes saved" },
    error: { border: "#F3C4B8", color: HR_T.bad, dot: HR_T.bad, label: "Save failed — data is NOT stored. Check your connection." },
  };
  const s = map[status] || map.saving;
  return (
    <div
      role="status"
      style={{
        position: "fixed",
        left: 22,
        bottom: 22,
        zIndex: 1150,
        display: "flex",
        alignItems: "center",
        gap: 8,
        padding: "9px 14px",
        borderRadius: 9,
        background: "#fff",
        border: `1px solid ${s.border}`,
        boxShadow: "0 10px 30px rgba(20,27,48,.16)",
        fontFamily: "'IBM Plex Sans', sans-serif",
        fontSize: 12,
        fontWeight: 600,
        color: s.color,
      }}
    >
      <span style={{ width: 8, height: 8, borderRadius: "50%", background: s.dot, display: "inline-block" }} />
      {s.label}
    </div>
  );
}

export default function App() {
  return (
    <AuthProvider>
      <AppShell />
    </AuthProvider>
  );
}

function AppShell() {
  const { user, login, logout } = useAuth();
  const [stage, setStage] = useState(() => (user ? "app" : "login")); // login | app
  const [leaving, setLeaving] = useState(false);
  const [toast, setToast] = useState(null);
  const toastTimerRef = useRef(null);

  function goTo(next) {
    setLeaving(true);
    setTimeout(() => {
      setStage(next);
      setLeaving(false);
    }, 260);
  }

  function showToast(message) {
    if (toastTimerRef.current) clearTimeout(toastTimerRef.current);
    setToast({ id: Date.now(), message });
    toastTimerRef.current = setTimeout(() => setToast(null), 3200);
  }

  function handleAuthenticatedLogin(nextUser, nextToken) {
    if (!login(nextUser, nextToken)) return;
    goTo("app");
  }

  function handleLogout() {
    logout();
    goTo("login");
    showToast("You have been signed out successfully.");
  }

  // Protected-route equivalent: if the session disappears for ANY reason
  // (any module calling useAuth().logout() directly, cleared storage),
  // the app drops back to the login stage immediately.
  const activeStage = stage === "app" && !user ? "login" : stage;

  return (
    <div style={{ minHeight: "100vh", width: "100%" }}>
      <style>{FONT_CSS}{STAGE_TRANSITION_CSS}{GLOBAL_LAYOUT_CSS}{`
      html, body, #root { width: 100%; min-height: 100%; margin: 0; padding: 0; }
      body { overflow-x: hidden; }
      .hr-app, .hr-shell, .ledger-shell { width: 100%; min-height: 100vh; }
      .hr-main, .ledger-main { min-width: 0; width: 100%; max-width: none !important; flex: 1 1 auto; }
      @media (max-width: 900px) {
        .hr-sidebar, .ledger-sidebar { width: 190px !important; padding-left: 10px !important; padding-right: 10px !important; }
        .hr-main, .ledger-main { padding: 22px 20px !important; }
        [style*="grid-template-columns"] { grid-template-columns: repeat(2, minmax(0, 1fr)) !important; }
      }
      @media (max-width: 700px) {
        .hr-shell, .ledger-shell { display: block !important; }
        .hr-sidebar, .ledger-sidebar { position: relative !important; top: auto !important; width: 100% !important; min-height: auto !important; padding: 12px !important; }
        .hr-sidebar nav, .ledger-sidebar nav { flex-direction: row !important; flex-wrap: nowrap !important; overflow-x: auto !important; gap: 4px !important; scrollbar-width: thin; }
        .hr-sidebar nav button, .ledger-sidebar nav button { flex: 0 0 auto !important; white-space: nowrap !important; }
        .hr-main, .ledger-main { padding: 16px 12px !important; }
        [style*="grid-template-columns"] { grid-template-columns: minmax(0, 1fr) !important; }
        [style*="justify-content: space-between"] { flex-wrap: wrap !important; gap: 12px !important; }
      }
      @media (max-width: 480px) {
        .hr-main, .ledger-main { padding: 12px 10px !important; }
        .hr-sidebar, .ledger-sidebar { padding: 10px !important; }
        [style*="grid-template-columns"] { grid-template-columns: minmax(0, 1fr) !important; }
        table { font-size: 12px; }
      }
      @media print {
        .no-print { display: none !important; }
        body { background: white !important; }
      }

      html, body, #root { width: 100%; min-height: 100%; margin: 0; padding: 0; }
      body { overflow-x: hidden; }
      .welcome-page, .login-page { width: 100%; min-height: 100vh; }
      @media (max-width: 760px) {
        .welcome-page { flex-direction: column !important; }
        .welcome-page > div:first-child { flex: 0 0 auto !important; width: 100% !important; min-width: 0 !important; padding: 38px 28px !important; }
        .welcome-page > div:last-child { padding: 34px 22px !important; }
      }
      @media (max-width: 480px) {
        .welcome-page > div:first-child { padding: 30px 20px !important; }
        .welcome-page > div:last-child { padding: 28px 16px !important; }
        .login-page { padding: 14px !important; }
        .login-page form { padding: 30px 22px !important; }
      }
      @media (max-height: 650px) and (min-width: 761px) {
        .welcome-page > div:first-child { padding-top: 28px !important; padding-bottom: 28px !important; }
        .login-page form { padding-top: 26px !important; padding-bottom: 26px !important; }
      }
    `}</style>
      <div key={activeStage} className={leaving ? "stage-leave" : "stage-enter"}>
        {activeStage === "login" && (
  <TopHRLoginPage
    onLogin={handleAuthenticatedLogin}
  />
)}
        {activeStage === "app" && <Workspace onLogout={handleLogout} />}
      </div>
      {toast && (
        <div key={toast.id} className="hrm-toast" role="status">
          <CheckCircle2 size={15} color={HR_T.good} />
          <span style={{ fontFamily: "'IBM Plex Sans', sans-serif", fontSize: 12.5, fontWeight: 600, color: HR_T.ink }}>
            {toast.message}
          </span>
        </div>
      )}
      <SyncChip />
    </div>
  );
}