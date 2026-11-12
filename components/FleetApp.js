import { useEffect, useMemo, useRef, useState } from "react";
import {
  BarChart, Bar, XAxis, YAxis, ResponsiveContainer, Tooltip, CartesianGrid, Cell,
} from "recharts";
import {
  Truck, Snowflake, Droplet, Sun, Wrench, Calendar, AlertTriangle, CheckCircle2,
  User, Phone, Search, X, Clock3, CreditCard, ChevronRight, ClipboardList, Gauge,
  Plus, Trash2, Pencil, LogOut, Gauge as GaugeIcon, Settings2, FileText, Lock, Hourglass, Fuel, Package, Box, Moon, Download, Camera, Upload, ShoppingCart, Palette,
} from "lucide-react";
import { supabase } from "../lib/supabaseClient";

/* ---------------------------------------------------------------
   CONFIG
---------------------------------------------------------------- */

const TODAY = new Date(new Date().toDateString());

const TEMP_META = {
  freezer: { label: "ห้องเย็น -18°C", icon: Snowflake, color: "#0E8FA0" },
  chiller: { label: "แช่เย็น 0-4°C", icon: Droplet, color: "#5B9FE0" },
  ambient: { label: "อุณหภูมิห้อง", icon: Sun, color: "#E0A85B" },
};
const THEMES = {
  light: {
    bg: "#F4F6F9", surface: "#FFFFFF", surface2: "#F1F4F8", border: "#E2E8F0",
    text: "#1E293B", textMuted: "#64748B", accent: "#0E8FA0",
    headerBg: "rgba(255,255,255,0.85)", chartTooltipBg: "#FFFFFF", chartAxis: "#64748B",
  },
  dark: {
    bg: "#0F1620", surface: "#182231", surface2: "#1F2A3B", border: "#2C3B4E",
    text: "#EAF1F7", textMuted: "#93A4B8", accent: "#3FC3D6",
    headerBg: "rgba(15,22,32,0.85)", chartTooltipBg: "#182231", chartAxis: "#93A4B8",
  },
};

const REPAIR_TYPES = ["เครื่องยนต์", "ระบบทำความเย็น", "ระบบไฟฟ้า", "เบรก", "ยาง", "แบตเตอรี่", "ตัวถัง", "อื่นๆ"];
const GARAGE_OPTIONS = ["ศูนย์บริการ ฮีโน่ทีพี จำกัด", "อู่ช่างอ้วน ตั้งวาล์ว เซอร์วิส จำกัด"];
const STATION_OPTIONS = ["ปตท. ลาดหลุมแก้ว"];
const PR_CATEGORIES = ["อะไหล่", "ยาง/แบตเตอรี่", "น้ำมันหล่อลื่น/น้ำยา", "อุปกรณ์/เครื่องมือ", "จ้างซ่อม/บริการ", "อื่นๆ"];
const PR_STATUSES = [
  { key: "รออนุมัติ", color: "#D97706" },
  { key: "อนุมัติแล้ว", color: "#0E8FA0" },
  { key: "เปิด PO แล้ว", color: "#6366F1" },
  { key: "รับของแล้ว", color: "#16A34A" },
  { key: "ยกเลิก", color: "#64748B" },
];
const DRIVER_LICENSE_TYPES = ["ท.2", "ท.3", "ท.4"];
const THAI_MONTHS = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];

/* ---------------------------------------------------------------
   HELPERS
---------------------------------------------------------------- */

function daysBetween(a, b) { return Math.round((b - a) / (1000 * 60 * 60 * 24)); }
function addDays(dateStr, n) { const d = new Date(dateStr); d.setDate(d.getDate() + n); return d; }
function fmtDate(d) { if (!d) return "-"; const date = typeof d === "string" ? new Date(d) : d; return `${date.getDate()} ${THAI_MONTHS[date.getMonth()]} ${date.getFullYear() + 543}`; }
function fmtMoney(n) { return Number(n || 0).toLocaleString("th-TH"); }

// เลข PR / PO: เก็บเฉพาะ "ตัวเลข" แล้วแสดงคำนำหน้า PR / PO ให้เอง (ข้อมูลเก่าที่พิมพ์ PR- นำหน้ามาแล้วก็รองรับ)
function refNumber(prefix, v) {
  return String(v ?? "").replace(new RegExp(`^\\s*${prefix}(?![A-Za-z])[\\s\\-._:/]*`, "i"), "").trim();
}
function refLabel(prefix, v) {
  const n = refNumber(prefix, v);
  return n ? `${prefix} ${n}` : "-";
}

// ส่งออกข้อมูลเป็นไฟล์ Excel (.xlsx) - โหลดไลบรารีเฉพาะตอนกดใช้งานจริง
async function exportToExcel(rows, filename, sheetName) {
  const XLSX = await import("xlsx");
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, filename);
}

/* ---------- รูปภาพพนักงานขับรถ (Supabase Storage แบบส่วนตัว) ---------- */
const DRIVER_BUCKET = "driver-photos";

// ย่อและบีบอัดรูปก่อนอัปโหลด (ประหยัดพื้นที่ + อัปโหลดเร็วบนมือถือ)
async function compressImage(file, maxDim = 1280, quality = 0.82) {
  if (!file.type || !file.type.startsWith("image/")) throw new Error("ไฟล์ที่เลือกไม่ใช่รูปภาพ");
  const url = URL.createObjectURL(file);
  try {
    const img = await new Promise((resolve, reject) => {
      const im = new Image();
      im.onload = () => resolve(im);
      im.onerror = () => reject(new Error("เปิดไฟล์รูปไม่ได้ ลองถ่ายใหม่หรือเลือกไฟล์ JPG/PNG"));
      im.src = url;
    });
    const scale = Math.min(1, maxDim / Math.max(img.naturalWidth, img.naturalHeight));
    const w = Math.max(1, Math.round(img.naturalWidth * scale));
    const h = Math.max(1, Math.round(img.naturalHeight * scale));
    const canvas = document.createElement("canvas");
    canvas.width = w;
    canvas.height = h;
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "#FFFFFF";
    ctx.fillRect(0, 0, w, h);
    ctx.drawImage(img, 0, 0, w, h);
    const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/jpeg", quality));
    if (!blob) throw new Error("บีบอัดรูปไม่สำเร็จ");
    return blob;
  } finally {
    URL.revokeObjectURL(url);
  }
}

async function uploadDriverImage(file, kind) {
  const blob = await compressImage(file);
  const rand = Math.random().toString(36).slice(2, 8);
  const path = `${kind}/${Date.now()}-${rand}.jpg`;
  const { error } = await supabase.storage.from(DRIVER_BUCKET).upload(path, blob, { contentType: "image/jpeg", upsert: false });
  if (error) throw error;
  return path;
}

async function removeDriverImages(paths) {
  const list = (paths || []).filter(Boolean);
  if (!list.length) return;
  try { await supabase.storage.from(DRIVER_BUCKET).remove(list); } catch (e) { /* ลบไม่สำเร็จไม่เป็นไร ไม่กระทบข้อมูลหลัก */ }
}

// ขอลิงก์ชั่วคราว (หมดอายุใน 1 ชม.) สำหรับดูรูปในที่เก็บไฟล์ส่วนตัว และต่ออายุให้อัตโนมัติ
function useSignedUrls(paths) {
  const [map, setMap] = useState({});
  const [tick, setTick] = useState(0);
  const key = Array.from(new Set((paths || []).filter(Boolean))).sort().join("|");

  useEffect(() => {
    const t = setInterval(() => setTick((x) => x + 1), 50 * 60 * 1000);
    return () => clearInterval(t);
  }, []);

  useEffect(() => {
    let cancelled = false;
    const list = key ? key.split("|") : [];
    if (!list.length) { setMap({}); return undefined; }
    (async () => {
      try {
        const { data, error } = await supabase.storage.from(DRIVER_BUCKET).createSignedUrls(list, 3600);
        if (cancelled || error || !data) return;
        const next = {};
        data.forEach((row) => { if (row.signedUrl && row.path) next[row.path] = row.signedUrl; });
        setMap(next);
      } catch (e) { /* ถ้าดึงรูปไม่ได้ ก็แค่ไม่แสดงรูป */ }
    })();
    return () => { cancelled = true; };
  }, [key, tick]);

  return map;
}
// cost === null หมายถึง "รอประเมินราคา"
function isPendingCost(c) { return c === null || c === undefined; }
function fmtCost(c) { return isPendingCost(c) ? "รอประเมินราคา" : `${fmtMoney(c)} บ.`; }
function sumCost(list) { return list.reduce((s, r) => s + (isPendingCost(r.cost) ? 0 : Number(r.cost || 0)), 0); }
function todayISO() { return new Date().toISOString().slice(0, 10); }

// คำนวณสถานะ "ครบรอบ" จากวันที่ล่าสุด + จำนวนวันต่อรอบ (ใช้กับ PM / คาลิเบรท)
function computeDue(lastDateStr, intervalDays) {
  if (!lastDateStr) return { next: null, daysLeft: null, status: "unknown" };
  const next = addDays(lastDateStr, intervalDays || 0);
  const daysLeft = daysBetween(TODAY, next);
  let status = "ok";
  if (daysLeft < 0) status = "overdue";
  else if (daysLeft <= 7) status = "soon";
  return { next, daysLeft, status };
}

// คำนวณสถานะ "วันหมดอายุ" ตรงๆ (ใช้กับภาษี/พรบ)
function computeExpiry(dateStr, warnDays = 30) {
  if (!dateStr) return { daysLeft: null, status: "unknown" };
  const d = new Date(dateStr);
  const daysLeft = daysBetween(TODAY, d);
  let status = "ok";
  if (daysLeft < 0) status = "overdue";
  else if (daysLeft <= warnDays) status = "soon";
  return { daysLeft, status };
}

function computeVehicle(v) {
  const pm = computeDue(v.last_pm, v.pm_interval);
  const coolingPm = computeDue(v.cooling_pm_last, v.cooling_pm_interval);
  const calibration = computeDue(v.calibration_last, v.calibration_interval);
  const tax = computeExpiry(v.tax_expiry, 30);
  return {
    ...v,
    pm, coolingPm, calibration, tax,
    nextPM: pm.next, daysLeft: pm.daysLeft, pmStatus: pm.status,
  };
}

async function fetchVehicles() {
  const { data, error } = await supabase.from("vehicles").select("*").order("id");
  if (error) throw error;
  return data;
}
async function fetchRepairs() {
  const { data, error } = await supabase.from("repairs").select("*").order("date", { ascending: false });
  if (error) throw error;
  return data;
}
async function fetchPurchaseRequests() {
  const { data, error } = await supabase.from("purchase_requests").select("*").order("request_date", { ascending: false });
  if (error) throw error;
  return data;
}
async function fetchFuelLogs() {
  const { data, error } = await supabase.from("fuel_logs").select("*").order("date", { ascending: false });
  if (error) throw error;
  return data;
}
async function fetchDrivers() {
  const { data, error } = await supabase.from("drivers").select("*").order("name");
  if (error) throw error;
  return data;
}

/* ---------------------------------------------------------------
   SMALL UI PIECES
---------------------------------------------------------------- */

function PlateBadge({ plate, size = "md" }) {
  const parts = (plate || "").split("-");
  const prefix = parts[0] || plate || "-";
  const num = parts[1] || "";
  const big = size === "lg";
  return (
    <div style={{ background: "#F4F2EC", border: "2px solid #1E3A5F", borderRadius: 6, padding: big ? "6px 12px" : "4px 8px", display: "inline-flex", flexDirection: "column", alignItems: "center", lineHeight: 1.1, minWidth: big ? 96 : 76, boxShadow: "0 1px 0 rgba(0,0,0,0.25) inset" }}>
      <span style={{ fontFamily: "'JetBrains Mono', monospace", fontWeight: 700, fontSize: big ? 20 : 15, color: "#141414", letterSpacing: 1 }}>{num ? `${prefix}-${num}` : prefix}</span>
      <span style={{ fontFamily: "'Inter', sans-serif", fontSize: big ? 10 : 8, color: "#1E3A5F", fontWeight: 600 }}>กรุงเทพมหานคร</span>
    </div>
  );
}

function TempChip({ temp }) {
  const meta = TEMP_META[temp] || TEMP_META.ambient;
  const Icon = meta.icon;
  return (
    <span className="inline-flex items-center gap-1 rounded-full px-2 py-1 text-xs font-medium" style={{ background: `${meta.color}22`, color: meta.color, border: `1px solid ${meta.color}55` }}>
      <Icon size={13} />{meta.label}
    </span>
  );
}

function StatusChip({ status, reason }) {
  const ready = status === "ready";
  return (
    <span className="inline-flex items-center gap-1 rounded-full px-2 py-1 text-xs font-semibold" style={{ background: ready ? "rgba(22,163,74,0.15)" : "rgba(220,38,38,0.15)", color: ready ? "#16A34A" : "#DC2626", border: `1px solid ${ready ? "#16A34A55" : "#DC262655"}` }} title={reason || ""}>
      {ready ? <CheckCircle2 size={13} /> : <AlertTriangle size={13} />}
      {ready ? "พร้อมใช้งาน" : "ไม่พร้อมใช้งาน"}
    </span>
  );
}

// ใช้ได้กับทุกประเภทกำหนดการ (PM เครื่องยนต์ / PM ตู้เย็น / คาลิเบรท / ภาษี)
function DueChip({ status, daysLeft }) {
  if (status === "unknown" || daysLeft === null || daysLeft === undefined) {
    return <span className="inline-flex items-center gap-1 rounded-full px-2 py-1 text-xs font-semibold" style={{ background: "#64748B1E", color: "#64748B", border: "1px solid #64748B55" }}>ไม่มีข้อมูล</span>;
  }
  const conf = {
    overdue: { color: "#DC2626", label: `เกินกำหนด ${Math.abs(daysLeft)} วัน` },
    soon: { color: "#D97706", label: `ใกล้ครบกำหนด (${daysLeft} วัน)` },
    ok: { color: "#64748B", label: `อีก ${daysLeft} วัน` },
  }[status];
  return (
    <span className="inline-flex items-center gap-1 rounded-full px-2 py-1 text-xs font-semibold" style={{ background: `${conf.color}1E`, color: conf.color, border: `1px solid ${conf.color}55` }}>
      <Clock3 size={13} />{conf.label}
    </span>
  );
}

function Card({ children, style }) {
  return <div style={{ background: "var(--surface)", border: "1px solid var(--border)", borderRadius: 14, boxShadow: "0 1px 2px rgba(15,23,42,0.04), 0 2px 10px rgba(15,23,42,0.04)", backdropFilter: "var(--card-blur, none)", WebkitBackdropFilter: "var(--card-blur, none)", ...style }}>{children}</div>;
}

function SectionTitle({ icon: Icon, children, sub }) {
  return (
    <div className="mb-4">
      <div className="flex items-center gap-2">
        <Icon size={18} style={{ color: "var(--accent-frost)" }} />
        <h2 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 19, fontWeight: 700, color: "var(--text)" }}>{children}</h2>
      </div>
      {sub && <p style={{ color: "var(--text-muted)", fontSize: 13, marginTop: 4 }}>{sub}</p>}
    </div>
  );
}

function SubHeading({ children }) {
  return (
    <div style={{ fontSize: 12, fontWeight: 700, color: "var(--accent-frost)", textTransform: "uppercase", letterSpacing: 0.5, margin: "4px 0 2px" }}>
      {children}
    </div>
  );
}

function Field({ label, children }) {
  return (
    <label className="flex flex-col gap-1">
      <span style={{ fontSize: 12, color: "var(--text-muted)", fontWeight: 600 }}>{label}</span>
      {children}
    </label>
  );
}

// ช่องกรอกเลข PR/PO: มีคำนำหน้า (PR / PO) ติดอยู่ด้านหน้า พิมพ์เฉพาะตัวเลข
function RefInput({ prefix, value, onChange, placeholder }) {
  return (
    <div className="flex items-stretch" style={{ border: "1px solid var(--border)", borderRadius: 8, overflow: "hidden", background: "var(--surface-2)" }}>
      <span style={{ padding: "0 12px", display: "inline-flex", alignItems: "center", background: "rgba(14,143,160,0.14)", color: "var(--accent-frost)", fontWeight: 700, fontSize: 13, borderRight: "1px solid var(--border)" }}>{prefix}</span>
      <input
        value={value}
        onChange={(e) => onChange(refNumber(prefix, e.target.value))}
        placeholder={placeholder || "พิมพ์เฉพาะเลข"}
        style={{ flex: 1, minWidth: 0, background: "transparent", border: "none", outline: "none", color: "var(--text)", fontSize: 14, padding: "9px 10px" }}
      />
    </div>
  );
}

// ช่องเลือกจากลิสต์ที่ตั้งไว้ล่วงหน้า พร้อมตัวเลือก "อื่นๆ (ระบุเอง)" เผื่อไม่มีในลิสต์
function SelectOrOther({ label, options, value, onChange, placeholder }) {
  const isCustom = value !== "" && !options.includes(value);
  const [customMode, setCustomMode] = useState(isCustom);

  return (
    <Field label={label}>
      <div className="flex flex-col gap-2">
        <select
          style={inputStyle}
          value={customMode ? "__other__" : value}
          onChange={(e) => {
            if (e.target.value === "__other__") { setCustomMode(true); onChange(""); }
            else { setCustomMode(false); onChange(e.target.value); }
          }}
        >
          <option value="">-- เลือก --</option>
          {options.map((o) => <option key={o} value={o}>{o}</option>)}
          <option value="__other__">อื่นๆ (ระบุเอง)</option>
        </select>
        {customMode && (
          <input style={inputStyle} placeholder={placeholder || "พิมพ์ชื่อเอง"} value={value} onChange={(e) => onChange(e.target.value)} />
        )}
      </div>
    </Field>
  );
}

const inputStyle = { background: "var(--surface-2)", border: "1px solid var(--border)", borderRadius: 8, padding: "8px 10px", color: "var(--text)", fontSize: 13, outline: "none", width: "100%" };
const iconBtnStyle = { color: "var(--text-muted)", padding: 4, borderRadius: 6, cursor: "pointer" };

function ModalShell({ title, onClose, children, width = 560 }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(8,10,14,0.7)", zIndex: 50, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ maxWidth: width, width: "100%", maxHeight: "90vh", overflowY: "auto" }}>
        <Card style={{ padding: 22, background: "var(--surface-solid, var(--surface))" }}>
          <div className="flex items-center justify-between mb-5">
            <span style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 17, fontWeight: 700, color: "var(--text)" }}>{title}</span>
            <button onClick={onClose} style={{ color: "var(--text-muted)" }}><X size={18} /></button>
          </div>
          {children}
        </Card>
      </div>
    </div>
  );
}

const DELETE_PASSWORD = "500412";

function ConfirmDelete({ label, onConfirm, onCancel, requirePassword = false }) {
  const [pwd, setPwd] = useState("");
  const [err, setErr] = useState("");

  function handleConfirm() {
    if (requirePassword) {
      if (pwd !== DELETE_PASSWORD) {
        setErr("รหัสผ่านไม่ถูกต้อง");
        return;
      }
    }
    onConfirm();
  }

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(8,10,14,0.55)", zIndex: 60, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }} onClick={onCancel}>
      <div onClick={(e) => e.stopPropagation()} style={{ maxWidth: 400, width: "100%" }}>
        <Card style={{ padding: 20, background: "var(--surface-solid, var(--surface))" }}>
          <div className="flex items-center gap-2 mb-3"><AlertTriangle size={18} style={{ color: "#DC2626" }} /><span style={{ fontSize: 15, fontWeight: 700, color: "var(--text)" }}>ยืนยันการลบ</span></div>
          <p style={{ fontSize: 13, color: "var(--text-muted)", marginBottom: 16 }}>{label}</p>

          {requirePassword && (
            <div style={{ marginBottom: 16 }}>
              <div className="flex items-center gap-2 rounded-lg px-3 py-2 mb-2" style={{ background: "rgba(217,119,6,0.1)", border: "1px solid #D9770655" }}>
                <Lock size={14} style={{ color: "#D97706" }} />
                <span style={{ fontSize: 12, color: "#D97706" }}>รายการนี้ต้องใส่รหัสผ่านเพื่อยืนยันการลบ</span>
              </div>
              <label className="flex flex-col gap-1">
                <span style={{ fontSize: 12, color: "var(--text-muted)", fontWeight: 600 }}>รหัสผ่านสำหรับลบข้อมูล</span>
                <input
                  style={inputStyle}
                  type="password"
                  value={pwd}
                  autoFocus
                  placeholder="กรอกรหัสผ่าน"
                  onChange={(e) => { setPwd(e.target.value); setErr(""); }}
                  onKeyDown={(e) => { if (e.key === "Enter") handleConfirm(); }}
                />
              </label>
              {err && <div style={{ fontSize: 12, color: "#DC2626", marginTop: 6 }}>{err}</div>}
            </div>
          )}

          <div className="flex items-center justify-end gap-2">
            <button onClick={onCancel} style={{ ...inputStyle, width: "auto", padding: "8px 16px", cursor: "pointer" }}>ยกเลิก</button>
            <button onClick={handleConfirm} style={{ background: "#DC2626", color: "#fff", border: "none", borderRadius: 8, padding: "8px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>ลบ</button>
          </div>
        </Card>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   VEHICLE FORM (add + edit) - ครบทุกข้อมูล
---------------------------------------------------------------- */

function VehicleFormModal({ initial, onClose, onSave, existingPlates }) {
  const isEdit = !!initial;
  const [form, setForm] = useState(() => initial ? {
    ...initial,
    year: String(initial.year || ""),
    pm_interval: String(initial.pm_interval),
    wheels: String(initial.wheels ?? 6),
    length_m: String(initial.length_m ?? ""),
    fuel_tank_liters: String(initial.fuel_tank_liters ?? ""),
    cooling_pm_interval: String(initial.cooling_pm_interval ?? 90),
    calibration_interval: String(initial.calibration_interval ?? 180),
    cooling_pm_last: initial.cooling_pm_last || todayISO(),
    calibration_last: initial.calibration_last || todayISO(),
    tire_changed: initial.tire_changed || todayISO(),
    battery_changed: initial.battery_changed || todayISO(),
    has_cargo: !!initial.has_cargo,
    cargo_note: initial.cargo_note || "",
    has_tail_lift: !!initial.has_tail_lift,
    tax_expiry: initial.tax_expiry || todayISO(),
    cooling_brand: initial.cooling_brand || "",
  } : {
    id: "", brand: "", model: "", year: "2024", temp: "chiller",
    status: "ready", reason: "", last_pm: todayISO(), pm_interval: "90", driver: "",
    cooling_brand: "", wheels: "6", length_m: "6", fuel_tank_liters: "150",
    tax_expiry: todayISO(), tire_changed: todayISO(), battery_changed: todayISO(),
    has_cargo: false, cargo_note: "",
    has_tail_lift: false,
    cooling_pm_last: todayISO(), cooling_pm_interval: "90",
    calibration_last: todayISO(), calibration_interval: "180",
  });
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);
  const update = (key, val) => setForm((f) => ({ ...f, [key]: val }));

  async function handleSubmit(e) {
    e.preventDefault();
    const plate = form.id.trim();
    if (!plate) return setError("กรุณากรอกทะเบียนรถ");
    const otherPlates = existingPlates.filter((p) => !isEdit || p !== initial.id);
    if (otherPlates.includes(plate)) return setError("มีทะเบียนนี้อยู่แล้วในระบบ");
    if (!form.brand.trim() || !form.model.trim()) return setError("กรุณากรอกยี่ห้อและรุ่น");
    if (!form.last_pm) return setError("กรุณาเลือกวันที่ PM เครื่องยนต์ล่าสุด");

    setSaving(true);
    setError("");
    try {
      await onSave({
        oldId: isEdit ? initial.id : null,
        id: plate, brand: form.brand.trim(), model: form.model.trim(),
        year: Number(form.year) || new Date().getFullYear(), temp: form.temp, status: form.status,
        reason: form.status === "not_ready" ? form.reason.trim() : "",
        last_pm: form.last_pm, pm_interval: Number(form.pm_interval) || 90, driver: form.driver.trim() || "-",
        cooling_brand: form.cooling_brand.trim(),
        wheels: Number(form.wheels) || 0,
        length_m: Number(form.length_m) || 0,
        fuel_tank_liters: Number(form.fuel_tank_liters) || 0,
        tax_expiry: form.tax_expiry || null,
        tire_changed: form.tire_changed || null,
        battery_changed: form.battery_changed || null,
        has_cargo: !!form.has_cargo,
        cargo_note: form.has_cargo ? form.cargo_note.trim() : "",
        has_tail_lift: !!form.has_tail_lift,
        cooling_pm_last: form.cooling_pm_last || null,
        cooling_pm_interval: Number(form.cooling_pm_interval) || 90,
        calibration_last: form.calibration_last || null,
        calibration_interval: Number(form.calibration_interval) || 180,
      });
      onClose();
    } catch (err) {
      setError(err.message || "บันทึกไม่สำเร็จ กรุณาลองใหม่อีกครั้ง");
    } finally {
      setSaving(false);
    }
  }

  return (
    <ModalShell title={isEdit ? "แก้ไขข้อมูลรถ" : "เพิ่มรถคันใหม่"} onClose={onClose}>
      <form onSubmit={handleSubmit} className="flex flex-col gap-3">
        <SubHeading>ข้อมูลทั่วไป</SubHeading>
        <div className="grid grid-cols-2 gap-3">
          <Field label="ทะเบียนรถ *"><input style={inputStyle} placeholder="เช่น 70-1234" value={form.id} onChange={(e) => update("id", e.target.value)} /></Field>
          <Field label="คนขับประจำ"><input style={inputStyle} placeholder="ชื่อ-นามสกุล" value={form.driver} onChange={(e) => update("driver", e.target.value)} /></Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="ยี่ห้อ *"><input style={inputStyle} placeholder="เช่น Isuzu" value={form.brand} onChange={(e) => update("brand", e.target.value)} /></Field>
          <Field label="รุ่น *"><input style={inputStyle} placeholder="เช่น FRR90" value={form.model} onChange={(e) => update("model", e.target.value)} /></Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="ปีรถ"><input style={inputStyle} type="number" value={form.year} onChange={(e) => update("year", e.target.value)} /></Field>
          <Field label="สถานะรถ">
            <select style={inputStyle} value={form.status} onChange={(e) => update("status", e.target.value)}>
              <option value="ready">พร้อมใช้งาน</option>
              <option value="not_ready">ไม่พร้อมใช้งาน</option>
            </select>
          </Field>
        </div>
        {form.status === "not_ready" && (
          <Field label="เหตุผลที่ไม่พร้อมใช้งาน"><input style={inputStyle} placeholder="เช่น รอซ่อมเครื่องยนต์" value={form.reason} onChange={(e) => update("reason", e.target.value)} /></Field>
        )}

        <label className="flex items-center gap-2 rounded-lg px-3 py-2.5" style={{ background: form.has_cargo ? "rgba(217,119,6,0.1)" : "var(--surface-2)", border: `1px solid ${form.has_cargo ? "#D9770655" : "var(--border)"}`, cursor: "pointer" }}>
          <input type="checkbox" checked={form.has_cargo} onChange={(e) => update("has_cargo", e.target.checked)} style={{ width: 15, height: 15, accentColor: "#D97706", cursor: "pointer" }} />
          <Package size={14} style={{ color: form.has_cargo ? "#D97706" : "var(--text-muted)" }} />
          <span style={{ fontSize: 13, color: form.has_cargo ? "#D97706" : "var(--text-muted)", fontWeight: 600 }}>มีสินค้า/ตะกร้าค้างอยู่ในตู้</span>
        </label>
        {form.has_cargo && (
          <Field label="รายละเอียดสินค้า/ตะกร้าที่ค้างอยู่"><input style={inputStyle} placeholder="เช่น ตะกร้าเปล่า 20 ใบ, สินค้าเบเกอรี่รอส่ง" value={form.cargo_note} onChange={(e) => update("cargo_note", e.target.value)} /></Field>
        )}

        <SubHeading>สเปครถ / ระบบทำความเย็น</SubHeading>
        <div className="grid grid-cols-2 gap-3">
          <Field label="ประเภทตู้บรรทุก">
            <select style={inputStyle} value={form.temp} onChange={(e) => update("temp", e.target.value)}>
              <option value="freezer">ห้องเย็น -18°C</option>
              <option value="chiller">แช่เย็น 0-4°C</option>
              <option value="ambient">อุณหภูมิห้อง</option>
            </select>
          </Field>
          <Field label="ยี่ห้อเครื่องทำความเย็น"><input style={inputStyle} placeholder="เช่น Thermo King, Carrier" value={form.cooling_brand} onChange={(e) => update("cooling_brand", e.target.value)} /></Field>
        </div>
        <div className="grid grid-cols-3 gap-3">
          <Field label="จำนวนล้อ"><input style={inputStyle} type="number" value={form.wheels} onChange={(e) => update("wheels", e.target.value)} /></Field>
          <Field label="ความยาว (เมตร)"><input style={inputStyle} type="number" step="0.1" value={form.length_m} onChange={(e) => update("length_m", e.target.value)} /></Field>
          <Field label="ถังน้ำมัน (ลิตร)"><input style={inputStyle} type="number" value={form.fuel_tank_liters} onChange={(e) => update("fuel_tank_liters", e.target.value)} /></Field>
        </div>
        <label className="flex items-center gap-2 rounded-lg px-3 py-2.5" style={{ background: form.has_tail_lift ? "rgba(14,143,160,0.1)" : "var(--surface-2)", border: `1px solid ${form.has_tail_lift ? "#0E8FA055" : "var(--border)"}`, cursor: "pointer" }}>
          <input type="checkbox" checked={form.has_tail_lift} onChange={(e) => update("has_tail_lift", e.target.checked)} style={{ width: 15, height: 15, accentColor: "#0E8FA0", cursor: "pointer" }} />
          <span style={{ fontSize: 13, color: form.has_tail_lift ? "var(--accent-frost)" : "var(--text-muted)", fontWeight: 600 }}>มีลิฟท์ท้าย (Tail Lift)</span>
        </label>

        <SubHeading>กำหนดการบำรุงรักษา</SubHeading>
        <div className="grid grid-cols-2 gap-3">
          <Field label="PM เครื่องยนต์ล่าสุด *"><input style={inputStyle} type="date" value={form.last_pm} onChange={(e) => update("last_pm", e.target.value)} /></Field>
          <Field label="รอบ PM เครื่องยนต์ (วัน)"><input style={inputStyle} type="number" value={form.pm_interval} onChange={(e) => update("pm_interval", e.target.value)} /></Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="PM ตู้เย็นล่าสุด"><input style={inputStyle} type="date" value={form.cooling_pm_last} onChange={(e) => update("cooling_pm_last", e.target.value)} /></Field>
          <Field label="รอบ PM ตู้เย็น (วัน)"><input style={inputStyle} type="number" value={form.cooling_pm_interval} onChange={(e) => update("cooling_pm_interval", e.target.value)} /></Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="คาลิเบรทตู้เย็นล่าสุด"><input style={inputStyle} type="date" value={form.calibration_last} onChange={(e) => update("calibration_last", e.target.value)} /></Field>
          <Field label="รอบคาลิเบรท (วัน)"><input style={inputStyle} type="number" value={form.calibration_interval} onChange={(e) => update("calibration_interval", e.target.value)} /></Field>
        </div>
        <Field label="วันหมดอายุภาษี/พ.ร.บ."><input style={inputStyle} type="date" value={form.tax_expiry} onChange={(e) => update("tax_expiry", e.target.value)} /></Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="เปลี่ยนยางล่าสุด"><input style={inputStyle} type="date" value={form.tire_changed} onChange={(e) => update("tire_changed", e.target.value)} /></Field>
          <Field label="เปลี่ยนแบตเตอรี่ล่าสุด"><input style={inputStyle} type="date" value={form.battery_changed} onChange={(e) => update("battery_changed", e.target.value)} /></Field>
        </div>

        {error && <div style={{ fontSize: 12, color: "#DC2626", background: "rgba(220,38,38,0.1)", border: "1px solid #DC262655", borderRadius: 8, padding: "8px 10px" }}>{error}</div>}
        <div className="flex items-center justify-end gap-2 mt-2">
          <button type="button" onClick={onClose} style={{ ...inputStyle, width: "auto", padding: "8px 16px", cursor: "pointer" }}>ยกเลิก</button>
          <button type="submit" disabled={saving} style={{ background: "var(--accent-frost)", color: "#FFFFFF", border: "none", borderRadius: 8, padding: "8px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer", opacity: saving ? 0.6 : 1 }}>
            {saving ? "กำลังบันทึก..." : "บันทึก"}
          </button>
        </div>
      </form>
    </ModalShell>
  );
}

/* ---------------------------------------------------------------
   REPAIR FORM (add + edit)
---------------------------------------------------------------- */

const REPAIR_SUGGESTIONS = {
  "เครื่องยนต์": ["เปลี่ยนถ่ายน้ำมันเครื่อง", "เปลี่ยนกรองน้ำมันเชื้อเพลิง", "ซ่อมเครื่องยนต์"],
  "ระบบทำความเย็น": ["เติมน้ำยาแอร์ตู้เย็น", "ล้างคอยล์เย็น", "เปลี่ยนคอมเพรสเซอร์"],
  "ระบบไฟฟ้า": ["เปลี่ยนหลอดไฟ", "ซ่อมสายไฟ", "เปลี่ยนไดชาร์จ"],
  "เบรก": ["เปลี่ยนผ้าเบรก", "เปลี่ยนน้ำมันเบรก", "ซ่อมระบบเบรก"],
  "ยาง": ["เปลี่ยนยาง", "ปะยาง", "สลับยาง/ตั้งศูนย์ถ่วงล้อ"],
  "แบตเตอรี่": ["เปลี่ยนแบตเตอรี่", "ชาร์จแบตเตอรี่"],
  "ตัวถัง": ["ซ่อมตัวถัง", "ซ่อมประตูตู้", "ทำสี"],
  "อื่นๆ": [],
};

function TapChip({ active, onClick, children, tone = "accent", color: customColor, style }) {
  const color = customColor || (tone === "amber" ? "#D97706" : tone === "green" ? "#16A34A" : "var(--accent-frost)");
  const tint = customColor ? `${customColor}22` : (tone === "amber" ? "rgba(217,119,6,0.12)" : tone === "green" ? "rgba(22,163,74,0.12)" : "rgba(14,143,160,0.14)");
  return (
    <button
      type="button"
      onClick={onClick}
      style={{
        minHeight: 38, padding: "6px 14px", borderRadius: 999, fontSize: 13, cursor: "pointer",
        fontWeight: active ? 700 : 500,
        background: active ? tint : "var(--surface-2)",
        border: `1.5px solid ${active ? color : "var(--border)"}`,
        color: active ? color : "var(--text-muted)",
        transition: "all .12s",
        ...style,
      }}
    >
      {children}
    </button>
  );
}

function StepLabel({ n, children }) {
  return (
    <div className="flex items-center gap-2" style={{ marginTop: 4 }}>
      <span style={{ width: 22, height: 22, borderRadius: 999, background: "var(--accent-frost)", color: "#FFFFFF", fontSize: 12, fontWeight: 700, display: "inline-flex", alignItems: "center", justifyContent: "center" }}>{n}</span>
      <span style={{ fontSize: 14, fontWeight: 700, color: "var(--text)" }}>{children}</span>
    </div>
  );
}

function RepairFormModal({ plate, initial, onClose, onSave }) {
  const isEdit = !!initial;
  const [form, setForm] = useState(() => initial ? {
    ...initial,
    cost: initial.cost === null || initial.cost === undefined ? "" : String(initial.cost),
    pendingEstimate: initial.cost === null || initial.cost === undefined,
    garage: !initial.garage || initial.garage === "-" ? "" : initial.garage,
    pr_number: refNumber("PR", initial.pr_number),
    po_number: refNumber("PO", initial.po_number),
  } : {
    date: todayISO(), type: REPAIR_TYPES[0], description: "", cost: "", garage: "", status: "เสร็จสิ้น", pendingEstimate: false, pr_number: "", po_number: "",
  });
  const [garageCustom, setGarageCustom] = useState(() => !!(initial && initial.garage && initial.garage !== "-" && !GARAGE_OPTIONS.includes(initial.garage)));
  const [showRefs, setShowRefs] = useState(() => !!(initial && (initial.pr_number || initial.po_number)));
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);
  const update = (key, val) => setForm((f) => ({ ...f, [key]: val }));

  const yesterdayISO = new Date(Date.now() - 86400000).toISOString().slice(0, 10);
  const descMissing = !!error && !form.description.trim();
  const suggestions = REPAIR_SUGGESTIONS[form.type] || [];

  function addSuggestion(text) {
    const cur = form.description.trim();
    if (!cur) update("description", text);
    else if (!cur.includes(text)) update("description", `${cur}, ${text}`);
  }

  async function handleSubmit(e) {
    e.preventDefault();
    if (!form.date) return setError("กรุณาเลือกวันที่ซ่อม");
    if (!form.description.trim()) return setError("กรุณากรอกว่าซ่อมอะไร (พิมพ์เอง หรือแตะตัวอย่างด้านล่างช่องได้)");
    setSaving(true);
    setError("");
    try {
      await onSave({
        id: isEdit ? initial.id : undefined,
        plate, date: form.date, type: form.type, description: form.description.trim(),
        cost: form.pendingEstimate ? null : (Number(form.cost) || 0),
        garage: form.garage.trim() || "-", status: form.status,
        pr_number: form.pr_number.trim(),
        po_number: form.po_number.trim(),
      });
      onClose();
    } catch (err) {
      setError(err.message || "บันทึกไม่สำเร็จ กรุณาลองใหม่อีกครั้ง");
    } finally {
      setSaving(false);
    }
  }

  return (
    <ModalShell title={isEdit ? "แก้ไขรายการซ่อม" : "เพิ่มรายการซ่อม"} onClose={onClose} width={640}>
      <div style={{ marginBottom: 14 }}><PlateBadge plate={plate} /></div>
      <form onSubmit={handleSubmit} className="flex flex-col gap-4">

        {/* 1. ซ่อมอะไร */}
        <div className="flex flex-col gap-2">
          <StepLabel n={1}>ซ่อมอะไร</StepLabel>
          <div className="flex flex-wrap gap-2">
            {REPAIR_TYPES.map((t) => (
              <TapChip key={t} active={form.type === t} onClick={() => update("type", t)}>{t}</TapChip>
            ))}
          </div>
          <input
            style={{ ...inputStyle, padding: "11px 12px", fontSize: 14, borderColor: descMissing ? "#DC2626" : "var(--border)" }}
            placeholder="ทำอะไรไปบ้าง เช่น เปลี่ยนผ้าเบรกหน้า"
            value={form.description}
            onChange={(e) => update("description", e.target.value)}
          />
          {suggestions.length > 0 && (
            <div className="flex flex-wrap items-center gap-2">
              <span style={{ fontSize: 11, color: "var(--text-muted)" }}>แตะเพื่อเติม:</span>
              {suggestions.map((sg) => (
                <button key={sg} type="button" onClick={() => addSuggestion(sg)} style={{ fontSize: 12, padding: "4px 10px", borderRadius: 999, background: "transparent", border: "1px dashed var(--border)", color: "var(--text-muted)", cursor: "pointer" }}>
                  + {sg}
                </button>
              ))}
            </div>
          )}
        </div>

        {/* 2. เมื่อไหร่ / ที่ไหน */}
        <div className="flex flex-col gap-2">
          <StepLabel n={2}>เมื่อไหร่ / ที่ไหน</StepLabel>
          <div className="flex flex-wrap items-center gap-2">
            <input style={{ ...inputStyle, width: "auto", minWidth: 160, padding: "9px 10px" }} type="date" value={form.date} onChange={(e) => update("date", e.target.value)} />
            <TapChip active={form.date === todayISO()} onClick={() => update("date", todayISO())}>วันนี้</TapChip>
            <TapChip active={form.date === yesterdayISO} onClick={() => update("date", yesterdayISO)}>เมื่อวาน</TapChip>
          </div>
          <div className="flex flex-col gap-2">
            {GARAGE_OPTIONS.map((g) => (
              <TapChip key={g} active={!garageCustom && form.garage === g} onClick={() => { setGarageCustom(false); update("garage", g); }} style={{ borderRadius: 10, textAlign: "left", width: "100%" }}>
                {g}
              </TapChip>
            ))}
            <TapChip active={garageCustom} onClick={() => { setGarageCustom(true); if (GARAGE_OPTIONS.includes(form.garage)) update("garage", ""); }} style={{ borderRadius: 10, textAlign: "left", width: "100%" }}>
              อื่นๆ (พิมพ์ชื่อเอง)
            </TapChip>
            {garageCustom && (
              <input style={{ ...inputStyle, padding: "10px 12px" }} autoFocus placeholder="พิมพ์ชื่ออู่/ศูนย์บริการ" value={form.garage} onChange={(e) => update("garage", e.target.value)} />
            )}
          </div>
        </div>

        {/* 3. ค่าใช้จ่ายและสถานะ */}
        <div className="flex flex-col gap-2">
          <StepLabel n={3}>ค่าใช้จ่ายและสถานะงาน</StepLabel>
          <div className="grid grid-cols-2 gap-2">
            <TapChip active={form.status === "เสร็จสิ้น"} tone="green" onClick={() => update("status", "เสร็จสิ้น")} style={{ borderRadius: 10, minHeight: 44 }}>
              เสร็จสิ้นแล้ว
            </TapChip>
            <TapChip active={form.status === "กำลังดำเนินการ"} tone="amber" onClick={() => update("status", "กำลังดำเนินการ")} style={{ borderRadius: 10, minHeight: 44 }}>
              กำลังดำเนินการ
            </TapChip>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <div className="flex items-center gap-2" style={{ ...inputStyle, width: "auto", flex: "1 1 180px", padding: "6px 12px", opacity: form.pendingEstimate ? 0.5 : 1 }}>
              <input
                type="number" inputMode="decimal" min="0"
                style={{ background: "transparent", border: "none", outline: "none", color: "var(--text)", fontSize: 20, fontWeight: 700, width: "100%" }}
                value={form.pendingEstimate ? "" : form.cost}
                disabled={form.pendingEstimate}
                placeholder={form.pendingEstimate ? "รอประเมินราคา" : "0"}
                onChange={(e) => update("cost", e.target.value)}
              />
              <span style={{ fontSize: 13, color: "var(--text-muted)" }}>บาท</span>
            </div>
            <TapChip active={form.pendingEstimate} tone="amber" onClick={() => update("pendingEstimate", !form.pendingEstimate)}>
              <span className="inline-flex items-center gap-1"><Hourglass size={14} />รอประเมินราคา</span>
            </TapChip>
          </div>
        </div>

        {/* 4. PR / PO (ไม่บังคับ) */}
        <div className="flex flex-col gap-2">
          <button type="button" onClick={() => setShowRefs((v) => !v)} className="flex items-center gap-1" style={{ fontSize: 13, color: "var(--accent-frost)", fontWeight: 600, background: "transparent", border: "none", cursor: "pointer", alignSelf: "flex-start", padding: 0 }}>
            <ChevronRight size={14} style={{ transform: showRefs ? "rotate(90deg)" : "none", transition: "transform .15s" }} />
            เพิ่มเลข PR / PO (ไม่บังคับ)
          </button>
          {showRefs && (
            <div className="grid grid-cols-2 gap-3">
              <Field label="เลขที่ PR"><RefInput prefix="PR" value={form.pr_number} onChange={(v) => update("pr_number", v)} /></Field>
              <Field label="เลขที่ PO"><RefInput prefix="PO" value={form.po_number} onChange={(v) => update("po_number", v)} /></Field>
            </div>
          )}
        </div>

        {error && <div style={{ fontSize: 12, color: "#DC2626", background: "rgba(220,38,38,0.1)", border: "1px solid #DC262655", borderRadius: 8, padding: "8px 10px" }}>{error}</div>}

        <div className="flex items-center justify-end gap-2" style={{ position: "sticky", bottom: 0, background: "var(--surface-solid, var(--surface))", paddingTop: 12, marginTop: 2, borderTop: "1px solid var(--border)" }}>
          <button type="button" onClick={onClose} style={{ ...inputStyle, width: "auto", padding: "10px 18px", cursor: "pointer" }}>ยกเลิก</button>
          <button type="submit" disabled={saving} style={{ background: "var(--accent-frost)", color: "#FFFFFF", border: "none", borderRadius: 8, padding: "10px 26px", fontSize: 14, fontWeight: 700, cursor: "pointer", opacity: saving ? 0.6 : 1 }}>
            {saving ? "กำลังบันทึก..." : "บันทึก"}
          </button>
        </div>
      </form>
    </ModalShell>
  );
}

/* ---------------------------------------------------------------
   DRIVER FORM (add + edit)
---------------------------------------------------------------- */

function ImagePicker({ label, existingUrl, file, removed, onPick, onRemove, round = false }) {
  const [preview, setPreview] = useState(null);
  useEffect(() => {
    if (!file) { setPreview(null); return undefined; }
    const u = URL.createObjectURL(file);
    setPreview(u);
    return () => URL.revokeObjectURL(u);
  }, [file]);
  const shown = preview || (!removed ? existingUrl : null);

  return (
    <div className="flex flex-col items-center gap-2 rounded-lg p-3" style={{ background: "var(--surface-2)", border: "1px solid var(--border)" }}>
      <span style={{ fontSize: 12, color: "var(--text-muted)", fontWeight: 600, textAlign: "center" }}>{label}</span>
      <div style={{ width: round ? 76 : 120, height: round ? 76 : 80, borderRadius: round ? 999 : 8, overflow: "hidden", background: "var(--border)", display: "flex", alignItems: "center", justifyContent: "center" }}>
        {shown ? <img src={shown} alt={label} style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : <Camera size={22} style={{ color: "var(--text-muted)" }} />}
      </div>
      <div className="flex items-center gap-2 flex-wrap justify-center">
        <label className="rounded-lg px-3 py-1.5" style={{ background: "var(--accent-frost)", color: "#FFFFFF", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
          {shown ? "เปลี่ยนรูป" : "เลือก/ถ่ายรูป"}
          <input type="file" accept="image/*" style={{ display: "none" }} onChange={(e) => { const f = e.target.files && e.target.files[0]; if (f) onPick(f); e.target.value = ""; }} />
        </label>
        {shown && (
          <button type="button" onClick={onRemove} className="rounded-lg px-3 py-1.5" style={{ background: "transparent", border: "1px solid #DC262655", color: "#DC2626", fontSize: 12, fontWeight: 600 }}>
            ลบรูป
          </button>
        )}
      </div>
    </div>
  );
}

function ImageViewer({ url, title, onClose }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(8,10,14,0.82)", zIndex: 70, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ maxWidth: 900, width: "100%" }}>
        <div className="flex items-center justify-between mb-2">
          <span style={{ color: "#FFFFFF", fontWeight: 600, fontSize: 14 }}>{title}</span>
          <button onClick={onClose} style={{ color: "#FFFFFF" }}><X size={20} /></button>
        </div>
        <img src={url} alt={title} style={{ width: "100%", maxHeight: "80vh", objectFit: "contain", borderRadius: 10, background: "#000" }} />
      </div>
    </div>
  );
}

function DriverFormModal({ initial, onClose, onSave, vehiclePlates, urlMap = {} }) {
  const isEdit = !!initial;
  const [form, setForm] = useState(() => initial ? {
    ...initial,
    years: String(initial.years),
    company_license_no: initial.company_license_no || "",
    company_license_expiry: initial.company_license_expiry || "",
  } : {
    name: "", license: "", license_type: "ท.2", expiry: "2027-01-01", phone: "", vehicle: "สำรอง", years: "1",
    company_license_no: "", company_license_expiry: "",
  });
  const [files, setFiles] = useState({ photo: null, license: null, company: null });
  const [removed, setRemoved] = useState({ photo: false, license: false, company: false });
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);
  const update = (key, val) => setForm((f) => ({ ...f, [key]: val }));
  const pickFile = (k, f) => { setFiles((x) => ({ ...x, [k]: f })); setRemoved((x) => ({ ...x, [k]: false })); };
  const removeFile = (k) => { setFiles((x) => ({ ...x, [k]: null })); setRemoved((x) => ({ ...x, [k]: true })); };

  async function handleSubmit(e) {
    e.preventDefault();
    if (!form.name.trim()) return setError("กรุณากรอกชื่อ-นามสกุล");
    if (!form.license.trim()) return setError("กรุณากรอกเลขที่ใบขับขี่");
    if (!form.expiry) return setError("กรุณาเลือกวันหมดอายุใบขับขี่");
    setSaving(true);
    setError("");
    const uploaded = []; // เก็บไว้ลบคืนถ้าบันทึกข้อมูลไม่สำเร็จ
    try {
      const resolvePath = async (field, kind, oldPath) => {
        if (files[field]) { const path = await uploadDriverImage(files[field], kind); uploaded.push(path); return path; }
        if (removed[field]) return null;
        return oldPath || null;
      };
      const photo_path = await resolvePath("photo", "face", initial && initial.photo_path);
      const license_photo_path = await resolvePath("license", "license", initial && initial.license_photo_path);
      const company_license_photo_path = await resolvePath("company", "company-license", initial && initial.company_license_photo_path);

      await onSave({
        id: isEdit ? initial.id : undefined,
        name: form.name.trim(), license: form.license.trim(), license_type: form.license_type,
        expiry: form.expiry, phone: form.phone.trim(), vehicle: form.vehicle, years: Number(form.years) || 0,
        photo_path, license_photo_path, company_license_photo_path,
        company_license_no: form.company_license_no.trim(),
        company_license_expiry: form.company_license_expiry || null,
      });

      // บันทึกสำเร็จแล้ว ค่อยลบไฟล์เก่าที่ถูกแทนที่/ลบออก
      if (initial) {
        const old = [];
        if (initial.photo_path && initial.photo_path !== photo_path) old.push(initial.photo_path);
        if (initial.license_photo_path && initial.license_photo_path !== license_photo_path) old.push(initial.license_photo_path);
        if (initial.company_license_photo_path && initial.company_license_photo_path !== company_license_photo_path) old.push(initial.company_license_photo_path);
        await removeDriverImages(old);
      }
      onClose();
    } catch (err) {
      await removeDriverImages(uploaded);
      setError(err.message || "บันทึกไม่สำเร็จ กรุณาลองใหม่อีกครั้ง");
    } finally {
      setSaving(false);
    }
  }

  return (
    <ModalShell title={isEdit ? "แก้ไขข้อมูลพนักงานขับรถ" : "เพิ่มพนักงานขับรถ"} onClose={onClose} width={680}>
      <form onSubmit={handleSubmit} className="flex flex-col gap-3">
        <SubHeading>รูปภาพ</SubHeading>
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
          <ImagePicker label="รูปหน้าคนขับ" round file={files.photo} removed={removed.photo} existingUrl={initial && initial.photo_path ? urlMap[initial.photo_path] : null} onPick={(f) => pickFile("photo", f)} onRemove={() => removeFile("photo")} />
          <ImagePicker label="รูปใบขับขี่" file={files.license} removed={removed.license} existingUrl={initial && initial.license_photo_path ? urlMap[initial.license_photo_path] : null} onPick={(f) => pickFile("license", f)} onRemove={() => removeFile("license")} />
          <ImagePicker label="รูปใบขับขี่ของเครือ" file={files.company} removed={removed.company} existingUrl={initial && initial.company_license_photo_path ? urlMap[initial.company_license_photo_path] : null} onPick={(f) => pickFile("company", f)} onRemove={() => removeFile("company")} />
        </div>
        <div style={{ fontSize: 11, color: "var(--text-muted)" }}>รูปจะถูกย่อขนาดอัตโนมัติ และเก็บแบบส่วนตัว (ต้องล็อกอินถึงจะเห็นได้)</div>

        <SubHeading>ข้อมูลทั่วไป</SubHeading>
        <Field label="ชื่อ-นามสกุล *"><input style={inputStyle} value={form.name} onChange={(e) => update("name", e.target.value)} /></Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="เลขที่ใบขับขี่ *"><input style={inputStyle} placeholder="เช่น 31-045678-9" value={form.license} onChange={(e) => update("license", e.target.value)} /></Field>
          <Field label="ประเภทใบขับขี่">
            <select style={inputStyle} value={form.license_type} onChange={(e) => update("license_type", e.target.value)}>
              {DRIVER_LICENSE_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
            </select>
          </Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="วันหมดอายุใบขับขี่ *"><input style={inputStyle} type="date" value={form.expiry} onChange={(e) => update("expiry", e.target.value)} /></Field>
          <Field label="เบอร์โทรศัพท์"><input style={inputStyle} placeholder="08x-xxx-xxxx" value={form.phone} onChange={(e) => update("phone", e.target.value)} /></Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="รถที่รับผิดชอบ">
            <select style={inputStyle} value={form.vehicle} onChange={(e) => update("vehicle", e.target.value)}>
              <option value="สำรอง">พนักงานสำรอง (ไม่ประจำรถ)</option>
              {vehiclePlates.map((p) => <option key={p} value={p}>{p}</option>)}
            </select>
          </Field>
          <Field label="ประสบการณ์ (ปี)"><input style={inputStyle} type="number" value={form.years} onChange={(e) => update("years", e.target.value)} /></Field>
        </div>
        <SubHeading>ใบขับขี่ของเครือ</SubHeading>
        <div className="grid grid-cols-2 gap-3">
          <Field label="เลขที่ใบขับขี่ของเครือ"><input style={inputStyle} placeholder="ถ้ามี" value={form.company_license_no} onChange={(e) => update("company_license_no", e.target.value)} /></Field>
          <Field label="วันหมดอายุใบขับขี่ของเครือ"><input style={inputStyle} type="date" value={form.company_license_expiry} onChange={(e) => update("company_license_expiry", e.target.value)} /></Field>
        </div>
        {error && <div style={{ fontSize: 12, color: "#DC2626", background: "rgba(220,38,38,0.1)", border: "1px solid #DC262655", borderRadius: 8, padding: "8px 10px" }}>{error}</div>}
        <div className="flex items-center justify-end gap-2 mt-2">
          <button type="button" onClick={onClose} style={{ ...inputStyle, width: "auto", padding: "8px 16px", cursor: "pointer" }}>ยกเลิก</button>
          <button type="submit" disabled={saving} style={{ background: "var(--accent-frost)", color: "#FFFFFF", border: "none", borderRadius: 8, padding: "8px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer", opacity: saving ? 0.6 : 1 }}>
            {saving ? "กำลังบันทึก..." : "บันทึก"}
          </button>
        </div>
      </form>
    </ModalShell>
  );
}

/* ---------------------------------------------------------------
   DASHBOARD
---------------------------------------------------------------- */

function Dashboard({ vehicles, repairs, onGoToVehicle, dark }) {
  const axisColor = dark ? "#93A4B8" : "#64748B";
  const gridColor = dark ? "#2C3B4E" : "#E2E8F0";
  const tooltipBg = dark ? "#1F2A3B" : "#FFFFFF";
  const tooltipBorder = dark ? "#334155" : "#E2E8F0";
  const labelColor = dark ? "#EAF1F7" : "#1E293B";
  const cursorFill = dark ? "rgba(255,255,255,0.05)" : "rgba(15,23,42,0.05)";
  const readyCount = vehicles.filter((v) => v.status === "ready").length;
  const notReadyCount = vehicles.length - readyCount;

  // รวมรายการแจ้งเตือนทุกประเภทเข้าด้วยกัน: PM เครื่องยนต์ / PM ตู้เย็น / คาลิเบรท / ภาษี
  const alerts = useMemo(() => {
    const list = [];
    vehicles.forEach((v) => {
      if (v.pm.status === "soon" || v.pm.status === "overdue") list.push({ plate: v.id, brand: v.brand, model: v.model, label: "PM เครื่องยนต์", status: v.pm.status, daysLeft: v.pm.daysLeft });
      if (v.coolingPm.status === "soon" || v.coolingPm.status === "overdue") list.push({ plate: v.id, brand: v.brand, model: v.model, label: "PM ตู้เย็น", status: v.coolingPm.status, daysLeft: v.coolingPm.daysLeft });
      if (v.calibration.status === "soon" || v.calibration.status === "overdue") list.push({ plate: v.id, brand: v.brand, model: v.model, label: "คาลิเบรทตู้เย็น", status: v.calibration.status, daysLeft: v.calibration.daysLeft });
      if (v.tax.status === "soon" || v.tax.status === "overdue") list.push({ plate: v.id, brand: v.brand, model: v.model, label: "ภาษี/พ.ร.บ.", status: v.tax.status, daysLeft: v.tax.daysLeft });
    });
    return list.sort((a, b) => a.daysLeft - b.daysLeft);
  }, [vehicles]);

  const recentRepairs = [...repairs].sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 6);

  const statusData = [
    { name: "พร้อมใช้งาน", value: readyCount, color: "#16A34A" },
    { name: "ไม่พร้อมใช้งาน", value: notReadyCount, color: "#DC2626" },
  ];

  const monthlyCost = useMemo(() => {
    const map = {};
    repairs.forEach((r) => {
      const d = new Date(r.date);
      const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
      map[key] = (map[key] || 0) + Number(r.cost || 0);
    });
    return Object.keys(map).sort().map((key) => {
      const [y, m] = key.split("-");
      return { month: `${THAI_MONTHS[Number(m) - 1]} '${String(Number(y) + 543).slice(2)}`, cost: map[key] };
    });
  }, [repairs]);

  const totalSpent = sumCost(repairs);

  // งานซ่อมที่กำลังดำเนินการอยู่ (ยังไม่เสร็จ)
  const inProgress = useMemo(
    () => repairs.filter((r) => r.status === "กำลังดำเนินการ").sort((a, b) => new Date(b.date) - new Date(a.date)),
    [repairs]
  );
  const pendingEstimateCount = inProgress.filter((r) => isPendingCost(r.cost)).length;

  const kpis = [
    { label: "รถทั้งหมด", value: vehicles.length, icon: Truck, color: "#0E8FA0" },
    { label: "พร้อมใช้งาน", value: readyCount, icon: CheckCircle2, color: "#16A34A" },
    { label: "กำลังซ่อม", value: inProgress.length, icon: Wrench, color: "#D97706" },
    { label: "รายการต้องดำเนินการ", value: alerts.length, icon: Clock3, color: "#DC2626" },
  ];

  return (
    <div>
      <SectionTitle icon={Gauge} sub="ภาพรวมสถานะรถบรรทุกควบคุมอุณหภูมิทั้งหมด ณ วันนี้">แดชบอร์ดภาพรวม</SectionTitle>

      <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
        {kpis.map((k) => (
          <Card key={k.label} style={{ padding: 18 }}>
            <div className="flex items-center justify-between mb-3"><k.icon size={20} style={{ color: k.color }} /></div>
            <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 30, fontWeight: 700, color: "var(--text)" }}>{k.value}</div>
            <div style={{ fontSize: 13, color: "var(--text-muted)", marginTop: 2 }}>{k.label}</div>
          </Card>
        ))}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-4 mb-6">
        <Card style={{ padding: 20 }}>
          <div style={{ fontSize: 13, color: "var(--text-muted)", marginBottom: 12, fontWeight: 600 }}>สถานะความพร้อมของรถ</div>
          <ResponsiveContainer width="100%" height={160}>
            <BarChart data={statusData} layout="vertical" margin={{ left: 0, right: 20 }}>
              <XAxis type="number" hide />
              <YAxis type="category" dataKey="name" width={110} tick={{ fill: axisColor, fontSize: 12 }} axisLine={false} tickLine={false} />
              <Tooltip contentStyle={{ background: tooltipBg, border: `1px solid ${tooltipBorder}`, borderRadius: 8, fontSize: 12 }} labelStyle={{ color: labelColor }} cursor={{ fill: cursorFill }} />
              <Bar dataKey="value" radius={[0, 6, 6, 0]} barSize={28}>{statusData.map((d, i) => <Cell key={i} fill={d.color} />)}</Bar>
            </BarChart>
          </ResponsiveContainer>
          <div style={{ marginTop: 8, fontSize: 12, color: "var(--text-muted)" }}>ค่าใช้จ่ายซ่อมบำรุงสะสมทั้งหมด: <span style={{ color: "var(--accent-frost)", fontWeight: 700 }}>{fmtMoney(totalSpent)} บาท</span></div>
        </Card>

        <Card style={{ padding: 20, gridColumn: "span 2 / span 2" }}>
          <div style={{ fontSize: 13, color: "var(--text-muted)", marginBottom: 12, fontWeight: 600 }}>ค่าใช้จ่ายซ่อมบำรุงรายเดือน (บาท)</div>
          <ResponsiveContainer width="100%" height={180}>
            <BarChart data={monthlyCost}>
              <CartesianGrid strokeDasharray="3 3" stroke={gridColor} vertical={false} />
              <XAxis dataKey="month" tick={{ fill: axisColor, fontSize: 12 }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fill: axisColor, fontSize: 11 }} axisLine={false} tickLine={false} />
              <Tooltip formatter={(v) => `${fmtMoney(v)} บาท`} contentStyle={{ background: tooltipBg, border: `1px solid ${tooltipBorder}`, borderRadius: 8, fontSize: 12 }} labelStyle={{ color: labelColor }} cursor={{ fill: cursorFill }} />
              <Bar dataKey="cost" radius={[6, 6, 0, 0]} fill="#0E8FA0" barSize={28} />
            </BarChart>
          </ResponsiveContainer>
        </Card>
      </div>

      <Card style={{ padding: 20, marginBottom: 16 }}>
        <div className="flex items-center justify-between flex-wrap gap-2 mb-3">
          <div className="flex items-center gap-2">
            <Wrench size={16} style={{ color: "#D97706" }} />
            <span style={{ fontSize: 14, fontWeight: 600, color: "var(--text)" }}>งานซ่อมที่กำลังดำเนินการ ({inProgress.length})</span>
          </div>
          {pendingEstimateCount > 0 && (
            <span className="inline-flex items-center gap-1 rounded-full px-2.5 py-1" style={{ background: "rgba(217,119,6,0.12)", color: "#D97706", border: "1px solid #D9770655", fontSize: 12, fontWeight: 600 }}>
              <Hourglass size={13} />รอประเมินราคา {pendingEstimateCount} รายการ
            </span>
          )}
        </div>
        <div className="flex flex-col gap-2" style={{ maxHeight: 300, overflowY: "auto" }}>
          {inProgress.length === 0 && <p style={{ color: "var(--text-muted)", fontSize: 13 }}>ไม่มีงานซ่อมที่กำลังดำเนินการอยู่ตอนนี้</p>}
          {inProgress.map((r) => (
            <div key={r.id} onClick={() => onGoToVehicle(r.plate)} className="flex items-center justify-between rounded-lg px-3 py-2.5" style={{ background: "var(--surface-2)", borderLeft: `3px solid ${isPendingCost(r.cost) ? "#D97706" : "#0E8FA0"}`, cursor: "pointer", transition: "background .15s" }} onMouseEnter={(e) => (e.currentTarget.style.background = "rgba(14,143,160,0.08)")} onMouseLeave={(e) => (e.currentTarget.style.background = "var(--surface-2)")}>
              <div className="flex items-center gap-3">
                <PlateBadge plate={r.plate} />
                <div>
                  <div style={{ fontSize: 13, color: "var(--text)", fontWeight: 500 }}>{r.type} · {r.description}</div>
                  <div style={{ fontSize: 11, color: "var(--text-muted)" }}>{r.garage} · เริ่ม {fmtDate(r.date)}</div>
                </div>
              </div>
              <div className="flex items-center gap-2" style={{ whiteSpace: "nowrap" }}>
                {isPendingCost(r.cost) ? (
                  <span className="inline-flex items-center gap-1 rounded-full px-2 py-1" style={{ background: "rgba(217,119,6,0.12)", color: "#D97706", border: "1px solid #D9770655", fontSize: 11, fontWeight: 600 }}>
                    <Hourglass size={12} />รอประเมินราคา
                  </span>
                ) : (
                  <div style={{ fontSize: 13, color: "var(--text)", fontWeight: 700 }}>{fmtMoney(r.cost)} บ.</div>
                )}
                <ChevronRight size={15} style={{ color: "var(--text-muted)" }} />
              </div>
            </div>
          ))}
        </div>
      </Card>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <Card style={{ padding: 20 }}>
          <div className="flex items-center gap-2 mb-3"><AlertTriangle size={16} style={{ color: "#D97706" }} /><span style={{ fontSize: 14, fontWeight: 600, color: "var(--text)" }}>รายการที่ใกล้/เกินกำหนด ({alerts.length})</span></div>
          <div className="flex flex-col gap-2" style={{ maxHeight: 320, overflowY: "auto" }}>
            {alerts.length === 0 && <p style={{ color: "var(--text-muted)", fontSize: 13 }}>ไม่มีรายการที่ต้องดำเนินการตอนนี้</p>}
            {alerts.map((a, i) => (
              <div key={i} onClick={() => onGoToVehicle(a.plate)} className="flex items-center justify-between rounded-lg px-3 py-2" style={{ background: "var(--surface-2)", cursor: "pointer", transition: "background .15s" }} onMouseEnter={(e) => (e.currentTarget.style.background = "rgba(14,143,160,0.08)")} onMouseLeave={(e) => (e.currentTarget.style.background = "var(--surface-2)")}>
                <div className="flex items-center gap-3">
                  <PlateBadge plate={a.plate} />
                  <div>
                    <div style={{ fontSize: 12, color: "var(--text)" }}>{a.label}</div>
                    <div style={{ fontSize: 11, color: "var(--text-muted)" }}>{a.brand} {a.model}</div>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  <DueChip status={a.status} daysLeft={a.daysLeft} />
                  <ChevronRight size={15} style={{ color: "var(--text-muted)" }} />
                </div>
              </div>
            ))}
          </div>
        </Card>

        <Card style={{ padding: 20 }}>
          <div className="flex items-center gap-2 mb-3"><Wrench size={16} style={{ color: "#0E8FA0" }} /><span style={{ fontSize: 14, fontWeight: 600, color: "var(--text)" }}>ประวัติการซ่อมล่าสุด</span></div>
          <div className="flex flex-col gap-2">
            {recentRepairs.length === 0 && <p style={{ color: "var(--text-muted)", fontSize: 13 }}>ยังไม่มีประวัติการซ่อม</p>}
            {recentRepairs.map((r) => (
              <div key={r.id} onClick={() => onGoToVehicle(r.plate)} className="flex items-center justify-between rounded-lg px-3 py-2" style={{ background: "var(--surface-2)", cursor: "pointer", transition: "background .15s" }} onMouseEnter={(e) => (e.currentTarget.style.background = "rgba(14,143,160,0.08)")} onMouseLeave={(e) => (e.currentTarget.style.background = "var(--surface-2)")}>
                <div><div style={{ fontSize: 13, color: "var(--text)", fontWeight: 500 }}>{r.plate} · {r.type}</div><div style={{ fontSize: 12, color: "var(--text-muted)" }}>{r.description}</div></div>
                <div className="flex items-center gap-2">
                  <div style={{ textAlign: "right" }}>
                    <div style={{ fontSize: 12, color: "var(--text-muted)", fontFamily: "'JetBrains Mono', monospace" }}>{fmtDate(r.date)}</div>
                    <div style={{ fontSize: 12, color: isPendingCost(r.cost) ? "#D97706" : "var(--accent-frost)", fontWeight: 600 }}>{fmtCost(r.cost)}</div>
                  </div>
                  <ChevronRight size={15} style={{ color: "var(--text-muted)" }} />
                </div>
              </div>
            ))}
          </div>
        </Card>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   VEHICLES VIEW
---------------------------------------------------------------- */

function VehiclesView({ vehicles, repairs, onAdd, onUpdate, onDelete, onAddRepair, onUpdateRepair, onDeleteRepair, initialPlate }) {
  const [query, setQuery] = useState("");
  const [selected, setSelected] = useState(initialPlate || null);
  const [formState, setFormState] = useState(null);
  const [deleteTarget, setDeleteTarget] = useState(null);
  const [repairForm, setRepairForm] = useState(null);
  const [deleteRepairTarget, setDeleteRepairTarget] = useState(null);

  const filtered = vehicles.filter((v) => v.id.includes(query) || v.brand.toLowerCase().includes(query.toLowerCase()) || v.model.toLowerCase().includes(query.toLowerCase()));
  const selVehicle = vehicles.find((v) => v.id === selected);
  const selRepairs = selVehicle ? repairs.filter((r) => r.plate === selVehicle.id).sort((a, b) => new Date(b.date) - new Date(a.date)) : [];
  const selTotalCost = sumCost(selRepairs);

  const groupedByMonth = useMemo(() => {
    const map = {};
    selRepairs.forEach((r) => {
      const d = new Date(r.date);
      const key = `${THAI_MONTHS[d.getMonth()]} ${d.getFullYear() + 543}`;
      if (!map[key]) map[key] = [];
      map[key].push(r);
    });
    return map;
  }, [selRepairs]);

  return (
    <div>
      <div className="flex items-center justify-between flex-wrap gap-3 mb-4">
        <SectionTitle icon={Truck} sub="สเปครถ กำหนดการบำรุงรักษา และประวัติการซ่อมรายทะเบียน">รถทั้งหมด</SectionTitle>
        <button onClick={() => setFormState("add")} className="flex items-center gap-2 rounded-lg px-4 py-2" style={{ background: "var(--accent-frost)", color: "#FFFFFF", fontSize: 13, fontWeight: 700, height: 38 }}>
          <Plus size={16} />เพิ่มรถคันใหม่
        </button>
      </div>

      <div className="flex items-center gap-2 mb-4">
        <div className="flex items-center gap-2 rounded-lg px-3 py-2 flex-1 max-w-sm" style={{ background: "var(--surface-2)", border: "1px solid var(--border)" }}>
          <Search size={15} style={{ color: "var(--text-muted)" }} />
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="ค้นหาทะเบียน / ยี่ห้อ / รุ่น" style={{ background: "transparent", outline: "none", color: "var(--text)", fontSize: 13, width: "100%" }} />
        </div>
      </div>

      <Card style={{ overflow: "hidden", marginBottom: selVehicle ? 20 : 0 }}>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 1160 }}>
            <thead>
              <tr style={{ background: "var(--surface-2)" }}>
                {["ทะเบียน", "ยี่ห้อ / รุ่น", "ประเภทตู้", "ล้อ", "ขนาดตู้ (ยาว)", "ถังน้ำมัน", "ลิฟท์ท้าย", "สถานะรถ", "คนขับ", "จัดการ", ""].map((h) => (
                  <th key={h} style={{ textAlign: "left", padding: "10px 14px", fontSize: 12, color: "var(--text-muted)", fontWeight: 600 }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map((v) => (
                <tr key={v.id} style={{ borderTop: "1px solid var(--border)", background: selected === v.id ? "rgba(14,143,160,0.08)" : "transparent" }}>
                  <td style={{ padding: "10px 14px", cursor: "pointer" }} onClick={() => setSelected(v.id === selected ? null : v.id)}><PlateBadge plate={v.id} /></td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: "var(--text)", cursor: "pointer" }} onClick={() => setSelected(v.id === selected ? null : v.id)}>{v.brand} {v.model} <span style={{ color: "var(--text-muted)" }}>'{String(v.year).slice(2)}</span></td>
                  <td style={{ padding: "10px 14px", cursor: "pointer" }} onClick={() => setSelected(v.id === selected ? null : v.id)}><TempChip temp={v.temp} /></td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: "var(--text)", cursor: "pointer" }} onClick={() => setSelected(v.id === selected ? null : v.id)}>{v.wheels || "-"} ล้อ</td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: "var(--text)", cursor: "pointer" }} onClick={() => setSelected(v.id === selected ? null : v.id)}>{v.length_m ? `${v.length_m} ม.` : "-"}</td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: "var(--text)", cursor: "pointer" }} onClick={() => setSelected(v.id === selected ? null : v.id)}>{v.fuel_tank_liters ? `${v.fuel_tank_liters} ล.` : "-"}</td>
                  <td style={{ padding: "10px 14px", cursor: "pointer" }} onClick={() => setSelected(v.id === selected ? null : v.id)}>
                    {v.has_tail_lift ? (
                      <span className="chip" style={{ background: "rgba(14,143,160,0.14)", color: "var(--accent-frost)", display: "inline-flex", alignItems: "center", gap: 4, borderRadius: 999, padding: "4px 9px", fontSize: 11, fontWeight: 600 }}>มี</span>
                    ) : (
                      <span style={{ fontSize: 12, color: "var(--text-muted)" }}>ไม่มี</span>
                    )}
                  </td>
                  <td style={{ padding: "10px 14px", cursor: "pointer" }} onClick={() => setSelected(v.id === selected ? null : v.id)}><StatusChip status={v.status} reason={v.reason} /></td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: "var(--text-muted)", cursor: "pointer" }} onClick={() => setSelected(v.id === selected ? null : v.id)}>{v.driver}</td>
                  <td style={{ padding: "10px 14px" }}>
                    <div className="flex items-center gap-1">
                      <button onClick={() => setFormState(v)} title="แก้ไข" style={iconBtnStyle}><Pencil size={15} /></button>
                      <button onClick={() => setDeleteTarget(v)} title="ลบ" style={{ ...iconBtnStyle, color: "#DC2626" }}><Trash2 size={15} /></button>
                    </div>
                  </td>
                  <td style={{ padding: "10px 14px", cursor: "pointer" }} onClick={() => setSelected(v.id === selected ? null : v.id)}>
                    <ChevronRight size={16} style={{ color: "var(--text-muted)", transform: selected === v.id ? "rotate(90deg)" : "none", transition: "transform .15s" }} />
                  </td>
                </tr>
              ))}
              {filtered.length === 0 && (
                <tr><td colSpan={11} style={{ padding: "24px 14px", textAlign: "center", color: "var(--text-muted)", fontSize: 13 }}>ไม่พบรถที่ตรงกับคำค้นหา</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </Card>

      {selVehicle && (
        <Card style={{ padding: 22 }}>
          <div className="flex items-start justify-between mb-5">
            <div className="flex items-center gap-4">
              <PlateBadge plate={selVehicle.id} size="lg" />
              <div>
                <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 18, fontWeight: 700, color: "var(--text)" }}>{selVehicle.brand} {selVehicle.model} · {selVehicle.year}</div>
                <div className="flex items-center gap-2 mt-1"><TempChip temp={selVehicle.temp} /><StatusChip status={selVehicle.status} reason={selVehicle.reason} /></div>
              </div>
            </div>
            <button onClick={() => setSelected(null)} style={{ color: "var(--text-muted)" }}><X size={18} /></button>
          </div>

          {selVehicle.status === "not_ready" && selVehicle.reason && (
            <div className="flex items-center gap-2 rounded-lg px-3 py-2 mb-3" style={{ background: "rgba(220,38,38,0.1)", border: "1px solid #DC262655" }}>
              <AlertTriangle size={14} style={{ color: "#DC2626" }} /><span style={{ fontSize: 13, color: "#DC2626" }}>{selVehicle.reason}</span>
            </div>
          )}

          {selVehicle.has_cargo ? (
            <div className="flex items-center gap-2 rounded-lg px-3 py-2 mb-5" style={{ background: "rgba(217,119,6,0.1)", border: "1px solid #D9770655" }}>
              <Package size={14} style={{ color: "#D97706" }} /><span style={{ fontSize: 13, color: "#D97706" }}>มีสินค้า/ตะกร้าค้างอยู่ในตู้: {selVehicle.cargo_note || "ไม่ได้ระบุรายละเอียด"}</span>
            </div>
          ) : (
            <div className="flex items-center gap-2 rounded-lg px-3 py-2 mb-5" style={{ background: "var(--surface-2)" }}>
              <Box size={14} style={{ color: "var(--text-muted)" }} /><span style={{ fontSize: 13, color: "var(--text-muted)" }}>ตู้ว่าง ไม่มีสินค้าค้าง</span>
            </div>
          )}

          <SubHeading>สเปครถ</SubHeading>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-4">
            <div className="rounded-lg px-3 py-3" style={{ background: "var(--surface-2)" }}><div style={{ fontSize: 11, color: "var(--text-muted)" }}>ยี่ห้อเครื่องเย็น</div><div style={{ fontSize: 13, color: "var(--text)", marginTop: 2 }}>{selVehicle.cooling_brand || "-"}</div></div>
            <div className="rounded-lg px-3 py-3" style={{ background: "var(--surface-2)" }}><div style={{ fontSize: 11, color: "var(--text-muted)" }}>จำนวนล้อ</div><div style={{ fontSize: 13, color: "var(--text)", marginTop: 2 }}>{selVehicle.wheels || "-"} ล้อ</div></div>
            <div className="rounded-lg px-3 py-3" style={{ background: "var(--surface-2)" }}><div style={{ fontSize: 11, color: "var(--text-muted)" }}>ความยาว</div><div style={{ fontSize: 13, color: "var(--text)", marginTop: 2 }}>{selVehicle.length_m || "-"} ม.</div></div>
            <div className="rounded-lg px-3 py-3" style={{ background: "var(--surface-2)" }}><div style={{ fontSize: 11, color: "var(--text-muted)" }}>ถังน้ำมัน</div><div style={{ fontSize: 13, color: "var(--text)", marginTop: 2 }}>{selVehicle.fuel_tank_liters || "-"} ลิตร</div></div>
            <div className="rounded-lg px-3 py-3" style={{ background: "var(--surface-2)" }}><div style={{ fontSize: 11, color: "var(--text-muted)" }}>ลิฟท์ท้าย</div><div style={{ fontSize: 13, color: selVehicle.has_tail_lift ? "var(--accent-frost)" : "var(--text-muted)", marginTop: 2, fontWeight: selVehicle.has_tail_lift ? 700 : 400 }}>{selVehicle.has_tail_lift ? "มี" : "ไม่มี"}</div></div>
          </div>

          <SubHeading>กำหนดการบำรุงรักษา</SubHeading>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-4">
            <div className="rounded-lg px-3 py-3" style={{ background: "var(--surface-2)" }}>
              <div style={{ fontSize: 11, color: "var(--text-muted)" }}>PM เครื่องยนต์</div>
              <div style={{ fontSize: 12, color: "var(--text)", fontFamily: "'JetBrains Mono', monospace", marginTop: 2 }}>{fmtDate(selVehicle.pm.next)}</div>
              <div style={{ marginTop: 4 }}><DueChip status={selVehicle.pm.status} daysLeft={selVehicle.pm.daysLeft} /></div>
            </div>
            <div className="rounded-lg px-3 py-3" style={{ background: "var(--surface-2)" }}>
              <div style={{ fontSize: 11, color: "var(--text-muted)" }}>PM ตู้เย็น</div>
              <div style={{ fontSize: 12, color: "var(--text)", fontFamily: "'JetBrains Mono', monospace", marginTop: 2 }}>{fmtDate(selVehicle.coolingPm.next)}</div>
              <div style={{ marginTop: 4 }}><DueChip status={selVehicle.coolingPm.status} daysLeft={selVehicle.coolingPm.daysLeft} /></div>
            </div>
            <div className="rounded-lg px-3 py-3" style={{ background: "var(--surface-2)" }}>
              <div style={{ fontSize: 11, color: "var(--text-muted)" }}>คาลิเบรทตู้เย็น</div>
              <div style={{ fontSize: 12, color: "var(--text)", fontFamily: "'JetBrains Mono', monospace", marginTop: 2 }}>{fmtDate(selVehicle.calibration.next)}</div>
              <div style={{ marginTop: 4 }}><DueChip status={selVehicle.calibration.status} daysLeft={selVehicle.calibration.daysLeft} /></div>
            </div>
            <div className="rounded-lg px-3 py-3" style={{ background: "var(--surface-2)" }}>
              <div style={{ fontSize: 11, color: "var(--text-muted)" }}>ภาษี/พ.ร.บ.</div>
              <div style={{ fontSize: 12, color: "var(--text)", fontFamily: "'JetBrains Mono', monospace", marginTop: 2 }}>{fmtDate(selVehicle.tax_expiry)}</div>
              <div style={{ marginTop: 4 }}><DueChip status={selVehicle.tax.status} daysLeft={selVehicle.tax.daysLeft} /></div>
            </div>
          </div>

          <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-6">
            <div className="rounded-lg px-3 py-3" style={{ background: "var(--surface-2)" }}><div style={{ fontSize: 11, color: "var(--text-muted)" }}>เปลี่ยนยางล่าสุด</div><div style={{ fontSize: 13, color: "var(--text)", fontFamily: "'JetBrains Mono', monospace", marginTop: 2 }}>{fmtDate(selVehicle.tire_changed)}</div></div>
            <div className="rounded-lg px-3 py-3" style={{ background: "var(--surface-2)" }}><div style={{ fontSize: 11, color: "var(--text-muted)" }}>เปลี่ยนแบตล่าสุด</div><div style={{ fontSize: 13, color: "var(--text)", fontFamily: "'JetBrains Mono', monospace", marginTop: 2 }}>{fmtDate(selVehicle.battery_changed)}</div></div>
            <div className="rounded-lg px-3 py-3" style={{ background: "var(--surface-2)" }}><div style={{ fontSize: 11, color: "var(--text-muted)" }}>คนขับประจำ</div><div style={{ fontSize: 13, color: "var(--text)", marginTop: 2 }}>{selVehicle.driver}</div></div>
            <div className="rounded-lg px-3 py-3" style={{ background: "rgba(14,143,160,0.1)", border: "1px solid #0E8FA055" }}><div style={{ fontSize: 11, color: "var(--accent-frost)" }}>ค่าซ่อมสะสมทั้งหมด</div><div style={{ fontSize: 15, color: "var(--accent-frost)", fontWeight: 700, marginTop: 2 }}>{fmtMoney(selTotalCost)} บ.</div></div>
          </div>

          <div className="flex items-center justify-between mb-3">
            <div className="flex items-center gap-2"><ClipboardList size={15} style={{ color: "var(--accent-frost)" }} /><span style={{ fontSize: 14, fontWeight: 600, color: "var(--text)" }}>ประวัติการซ่อม ({selRepairs.length} รายการ)</span></div>
            <div className="flex items-center gap-2">
              <button
                onClick={async () => {
                  const rows = selRepairs.map((r) => ({
                    "ทะเบียน": r.plate, "วันที่": fmtDate(r.date), "ประเภท": r.type, "รายละเอียด": r.description,
                    "เลขที่ PR": r.pr_number ? refLabel("PR", r.pr_number) : "", "เลขที่ PO": r.po_number ? refLabel("PO", r.po_number) : "",
                    "อู่/ศูนย์บริการ": r.garage, "สถานะ": r.status,
                    "ค่าใช้จ่าย (บาท)": isPendingCost(r.cost) ? "รอประเมินราคา" : Number(r.cost || 0),
                  }));
                  await exportToExcel(rows, `ประวัติการซ่อม-${selVehicle.id}-${todayISO()}.xlsx`, "ประวัติการซ่อม");
                }}
                disabled={selRepairs.length === 0}
                className="flex items-center gap-1 rounded-lg px-3 py-1.5"
                style={{ background: "#16A34A", border: "none", color: "#FFFFFF", fontSize: 12, fontWeight: 600, opacity: selRepairs.length === 0 ? 0.6 : 1 }}
              >
                <Download size={14} />Excel
              </button>
              <button onClick={() => setRepairForm("add")} className="flex items-center gap-1 rounded-lg px-3 py-1.5" style={{ background: "var(--surface-2)", border: "1px solid var(--border)", color: "var(--accent-frost)", fontSize: 12, fontWeight: 600 }}>
                <Plus size={14} />เพิ่มรายการซ่อม
              </button>
            </div>
          </div>

          {Object.keys(groupedByMonth).length === 0 && <p style={{ fontSize: 13, color: "var(--text-muted)" }}>ยังไม่มีประวัติการซ่อมสำหรับคันนี้</p>}

          <div className="flex flex-col gap-5">
            {Object.entries(groupedByMonth).map(([month, items]) => (
              <div key={month}>
                <div style={{ fontSize: 12, fontWeight: 700, color: "var(--accent-frost)", marginBottom: 8, textTransform: "uppercase", letterSpacing: 0.5 }}>{month} · รวม {fmtMoney(sumCost(items))} บาท</div>
                <div className="flex flex-col gap-2">
                  {items.map((r) => (
                    <div key={r.id} className="flex items-center justify-between rounded-lg px-3 py-2" style={{ background: "var(--surface-2)" }}>
                      <div className="flex items-center gap-3">
                        <span className="rounded-full px-2 py-1 text-xs font-medium" style={{ background: "rgba(14,143,160,0.12)", color: "var(--accent-frost)", whiteSpace: "nowrap" }}>{r.type}</span>
                        <div><div style={{ fontSize: 13, color: "var(--text)" }}>{r.description}</div><div style={{ fontSize: 11, color: "var(--text-muted)" }}>{r.garage} · {r.status}{r.pr_number ? ` · ${refLabel("PR", r.pr_number)}` : ""}{r.po_number ? ` · ${refLabel("PO", r.po_number)}` : ""}</div></div>
                      </div>
                      <div className="flex items-center gap-3">
                        <div style={{ textAlign: "right" }}>
                          <div style={{ fontSize: 12, color: "var(--text-muted)", fontFamily: "'JetBrains Mono', monospace" }}>{fmtDate(r.date)}</div>
                          <div style={{ fontSize: 13, color: isPendingCost(r.cost) ? "#D97706" : "var(--text)", fontWeight: 600 }}>{fmtCost(r.cost)}</div>
                        </div>
                        <div className="flex items-center gap-1">
                          <button onClick={() => setRepairForm(r)} title="แก้ไข" style={iconBtnStyle}><Pencil size={14} /></button>
                          <button onClick={() => setDeleteRepairTarget(r)} title="ลบ" style={{ ...iconBtnStyle, color: "#DC2626" }}><Trash2 size={14} /></button>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </Card>
      )}

      {formState && (
        <VehicleFormModal
          initial={formState === "add" ? null : formState}
          existingPlates={vehicles.map((v) => v.id)}
          onClose={() => setFormState(null)}
          onSave={async (data) => { if (data.oldId) await onUpdate(data.oldId, data); else await onAdd(data); }}
        />
      )}

      {deleteTarget && (
        <ConfirmDelete
          requirePassword
          label={`ต้องการลบรถทะเบียน ${deleteTarget.id} ใช่หรือไม่? ประวัติการซ่อมของรถคันนี้จะถูกลบไปด้วย`}
          onCancel={() => setDeleteTarget(null)}
          onConfirm={async () => { await onDelete(deleteTarget.id); if (selected === deleteTarget.id) setSelected(null); setDeleteTarget(null); }}
        />
      )}

      {repairForm && selVehicle && (
        <RepairFormModal
          plate={selVehicle.id}
          initial={repairForm === "add" ? null : repairForm}
          onClose={() => setRepairForm(null)}
          onSave={async (data) => { if (repairForm === "add") await onAddRepair(data); else await onUpdateRepair(data); }}
        />
      )}

      {deleteRepairTarget && (
        <ConfirmDelete
          label={`ต้องการลบรายการซ่อม "${deleteRepairTarget.description}" ใช่หรือไม่?`}
          onCancel={() => setDeleteRepairTarget(null)}
          onConfirm={async () => { await onDeleteRepair(deleteRepairTarget.id); setDeleteRepairTarget(null); }}
        />
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   MAINTENANCE VIEW (PM เครื่องยนต์ / PM ตู้เย็น / คาลิเบรท / ภาษี)
---------------------------------------------------------------- */

function MaintenanceView({ vehicles }) {
  const rows = useMemo(() => {
    return [...vehicles].map((v) => {
      const worst = Math.min(
        v.pm.daysLeft ?? Infinity,
        v.coolingPm.daysLeft ?? Infinity,
        v.calibration.daysLeft ?? Infinity,
        v.tax.daysLeft ?? Infinity
      );
      return { ...v, worst };
    }).sort((a, b) => a.worst - b.worst);
  }, [vehicles]);

  return (
    <div>
      <SectionTitle icon={Calendar} sub="รวม PM เครื่องยนต์ / PM ตู้เย็น / คาลิเบรทตู้เย็น / ภาษี-พ.ร.บ. ของรถทุกคันไว้ที่เดียว เรียงตามรายการที่ใกล้ครบกำหนดที่สุด">
        กำหนดการบำรุงรักษา
      </SectionTitle>
      <div className="flex items-center gap-4 mb-4 flex-wrap" style={{ fontSize: 12, color: "var(--text-muted)" }}>
        <span className="flex items-center gap-1"><span style={{ width: 8, height: 8, borderRadius: 99, background: "#DC2626", display: "inline-block" }} /> เกินกำหนด</span>
        <span className="flex items-center gap-1"><span style={{ width: 8, height: 8, borderRadius: 99, background: "#D97706", display: "inline-block" }} /> ใกล้ครบกำหนด</span>
        <span className="flex items-center gap-1"><span style={{ width: 8, height: 8, borderRadius: 99, background: "#64748B", display: "inline-block" }} /> ปกติ</span>
      </div>
      <Card style={{ overflow: "hidden" }}>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 1260 }}>
            <thead>
              <tr style={{ background: "var(--surface-2)" }}>
                {["ทะเบียน", "ยี่ห้อ / รุ่น", "ล้อ", "ยี่ห้อเครื่องเย็น", "PM เครื่องยนต์", "PM ตู้เย็น", "คาลิเบรทตู้เย็น", "ภาษี/พ.ร.บ.", "เปลี่ยนยางล่าสุด", "เปลี่ยนแบตล่าสุด"].map((h) => (
                  <th key={h} style={{ textAlign: "left", padding: "10px 14px", fontSize: 12, color: "var(--text-muted)", fontWeight: 600 }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((v) => (
                <tr key={v.id} style={{ borderTop: "1px solid var(--border)" }}>
                  <td style={{ padding: "10px 14px" }}><PlateBadge plate={v.id} /></td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: "var(--text)" }}>{v.brand} {v.model}</td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: "var(--text-muted)" }}>{v.wheels || "-"} ล้อ</td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: "var(--text-muted)" }}>{v.cooling_brand || "-"}</td>
                  <td style={{ padding: "10px 14px" }}>
                    <DueChip status={v.pm.status} daysLeft={v.pm.daysLeft} />
                    <div style={{ fontSize: 11, color: "var(--text-muted)", marginTop: 4, fontFamily: "'JetBrains Mono', monospace" }}>{fmtDate(v.pm.next)}</div>
                  </td>
                  <td style={{ padding: "10px 14px" }}>
                    <DueChip status={v.coolingPm.status} daysLeft={v.coolingPm.daysLeft} />
                    <div style={{ fontSize: 11, color: "var(--text-muted)", marginTop: 4, fontFamily: "'JetBrains Mono', monospace" }}>{fmtDate(v.coolingPm.next)}</div>
                  </td>
                  <td style={{ padding: "10px 14px" }}>
                    <DueChip status={v.calibration.status} daysLeft={v.calibration.daysLeft} />
                    <div style={{ fontSize: 11, color: "var(--text-muted)", marginTop: 4, fontFamily: "'JetBrains Mono', monospace" }}>{fmtDate(v.calibration.next)}</div>
                  </td>
                  <td style={{ padding: "10px 14px" }}>
                    <DueChip status={v.tax.status} daysLeft={v.tax.daysLeft} />
                    <div style={{ fontSize: 11, color: "var(--text-muted)", marginTop: 4, fontFamily: "'JetBrains Mono', monospace" }}>{fmtDate(v.tax_expiry)}</div>
                  </td>
                  <td style={{ padding: "10px 14px", fontSize: 12, color: "var(--text-muted)", fontFamily: "'JetBrains Mono', monospace" }}>{fmtDate(v.tire_changed)}</td>
                  <td style={{ padding: "10px 14px", fontSize: 12, color: "var(--text-muted)", fontFamily: "'JetBrains Mono', monospace" }}>{fmtDate(v.battery_changed)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

/* ---------------------------------------------------------------
   REPAIRS LOG VIEW (ประวัติการซ่อมทั้งหมด - ทุกคันรวมกัน)
---------------------------------------------------------------- */

function RepairsLogView({ vehicles, repairs, onUpdateRepair, onDeleteRepair, onGoToVehicle }) {
  const [editTarget, setEditTarget] = useState(null);
  const [deleteTarget, setDeleteTarget] = useState(null);
  const [plateFilter, setPlateFilter] = useState("all");
  const [monthFilter, setMonthFilter] = useState("all");
  const [statusFilter, setStatusFilter] = useState("all");
  const [query, setQuery] = useState("");
  const [sortBy, setSortBy] = useState("date_desc");

  const monthOptions = useMemo(() => {
    const set = new Set();
    repairs.forEach((r) => {
      const d = new Date(r.date);
      set.add(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`);
    });
    return Array.from(set).sort().reverse();
  }, [repairs]);

  const filtered = useMemo(() => {
    let list = [...repairs];
    if (plateFilter !== "all") list = list.filter((r) => r.plate === plateFilter);
    if (statusFilter !== "all") list = list.filter((r) => r.status === statusFilter);
    if (monthFilter !== "all") {
      list = list.filter((r) => {
        const d = new Date(r.date);
        const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
        return key === monthFilter;
      });
    }
    if (query.trim()) {
      const q = query.trim().toLowerCase();
      list = list.filter((r) => r.description.toLowerCase().includes(q) || r.type.toLowerCase().includes(q) || r.plate.toLowerCase().includes(q) || (r.garage || "").toLowerCase().includes(q) || (r.pr_number || "").toLowerCase().includes(q) || (r.po_number || "").toLowerCase().includes(q) || refLabel("PR", r.pr_number).toLowerCase().includes(q) || refLabel("PO", r.po_number).toLowerCase().includes(q));
    }
    if (sortBy === "date_desc") list.sort((a, b) => new Date(b.date) - new Date(a.date));
    else if (sortBy === "date_asc") list.sort((a, b) => new Date(a.date) - new Date(b.date));
    else if (sortBy === "cost_desc") list.sort((a, b) => b.cost - a.cost);
    else if (sortBy === "cost_asc") list.sort((a, b) => a.cost - b.cost);
    return list;
  }, [repairs, plateFilter, monthFilter, statusFilter, query, sortBy]);

  const totalFiltered = sumCost(filtered);

  const [exporting, setExporting] = useState(false);
  async function handleExport() {
    setExporting(true);
    try {
      const rows = filtered.map((r) => ({
        "ทะเบียน": r.plate,
        "วันที่": fmtDate(r.date),
        "ประเภท": r.type,
        "รายละเอียด": r.description,
        "เลขที่ PR": r.pr_number ? refLabel("PR", r.pr_number) : "",
        "เลขที่ PO": r.po_number ? refLabel("PO", r.po_number) : "",
        "อู่/ศูนย์บริการ": r.garage,
        "สถานะ": r.status,
        "ค่าใช้จ่าย (บาท)": isPendingCost(r.cost) ? "รอประเมินราคา" : Number(r.cost || 0),
      }));
      await exportToExcel(rows, `ประวัติการซ่อม-${todayISO()}.xlsx`, "ประวัติการซ่อม");
    } finally {
      setExporting(false);
    }
  }

  const monthLabel = (key) => {
    const [y, m] = key.split("-");
    return `${THAI_MONTHS[Number(m) - 1]} ${Number(y) + 543}`;
  };

  return (
    <div>
      <div className="flex items-center justify-between flex-wrap gap-3 mb-4">
        <SectionTitle icon={ClipboardList} sub="รวมประวัติการซ่อมของรถทุกคัน กรองดูตามรถ / เดือน / คำค้นหาได้">
          ประวัติการซ่อมทั้งหมด
        </SectionTitle>
        <button onClick={handleExport} disabled={exporting || filtered.length === 0} className="flex items-center gap-2 rounded-lg px-4 py-2" style={{ background: "#16A34A", color: "#FFFFFF", fontSize: 13, fontWeight: 700, height: 38, opacity: exporting || filtered.length === 0 ? 0.6 : 1 }}>
          <Download size={16} />{exporting ? "กำลังสร้างไฟล์..." : "ดาวน์โหลด Excel"}
        </button>
      </div>

      <div className="flex items-center gap-2 mb-4 flex-wrap">
        <select value={plateFilter} onChange={(e) => setPlateFilter(e.target.value)} style={{ ...inputStyle, width: "auto", minWidth: 140 }}>
          <option value="all">ทุกคัน</option>
          {vehicles.map((v) => <option key={v.id} value={v.id}>{v.id} · {v.brand} {v.model}</option>)}
        </select>
        <select value={monthFilter} onChange={(e) => setMonthFilter(e.target.value)} style={{ ...inputStyle, width: "auto", minWidth: 140 }}>
          <option value="all">ทุกเดือน</option>
          {monthOptions.map((m) => <option key={m} value={m}>{monthLabel(m)}</option>)}
        </select>
        <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} style={{ ...inputStyle, width: "auto", minWidth: 150 }}>
          <option value="all">ทุกสถานะ</option>
          <option value="เสร็จสิ้น">เสร็จสิ้นแล้ว</option>
          <option value="กำลังดำเนินการ">กำลังดำเนินการ</option>
        </select>
        <select value={sortBy} onChange={(e) => setSortBy(e.target.value)} style={{ ...inputStyle, width: "auto", minWidth: 160 }}>
          <option value="date_desc">วันที่ล่าสุดก่อน</option>
          <option value="date_asc">วันที่เก่าสุดก่อน</option>
          <option value="cost_desc">ค่าใช้จ่ายมากไปน้อย</option>
          <option value="cost_asc">ค่าใช้จ่ายน้อยไปมาก</option>
        </select>
        <div className="flex items-center gap-2 rounded-lg px-3 py-2" style={{ background: "var(--surface-2)", border: "1px solid var(--border)", minWidth: 200, flex: 1 }}>
          <Search size={15} style={{ color: "var(--text-muted)" }} />
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="ค้นหา เช่น ยาง, เบรก, ชื่ออู่" style={{ background: "transparent", outline: "none", color: "var(--text)", fontSize: 13, width: "100%" }} />
        </div>
      </div>

      <div className="flex items-center justify-between mb-3 rounded-lg px-4 py-2.5" style={{ background: "rgba(14,143,160,0.08)", border: "1px solid #0E8FA045" }}>
        <span style={{ fontSize: 13, color: "var(--text)" }}>
          พบ {filtered.length} รายการ
          <span style={{ color: "var(--text-muted)" }}> · เสร็จสิ้น {filtered.filter((r) => r.status === "เสร็จสิ้น").length} · กำลังดำเนินการ {filtered.filter((r) => r.status === "กำลังดำเนินการ").length}</span>
        </span>
        <span style={{ fontSize: 13, color: "var(--accent-frost)", fontWeight: 700 }}>รวม {fmtMoney(totalFiltered)} บาท</span>
      </div>

      <Card style={{ overflow: "hidden" }}>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 1060 }}>
            <thead>
              <tr style={{ background: "var(--surface-2)" }}>
                {["ทะเบียน", "วันที่", "ประเภท", "รายละเอียด", "เลขที่ PR", "เลขที่ PO", "อู่/ศูนย์บริการ", "สถานะ", "ค่าใช้จ่าย", "จัดการ"].map((h) => (
                  <th key={h} style={{ textAlign: "left", padding: "10px 14px", fontSize: 12, color: "var(--text-muted)", fontWeight: 600 }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map((r) => (
                <tr
                  key={r.id}
                  onClick={() => setEditTarget(r)}
                  title="คลิกเพื่อแก้ไขรายการซ่อมนี้"
                  style={{ borderTop: "1px solid var(--border)", cursor: "pointer", transition: "background .12s" }}
                  onMouseEnter={(e) => (e.currentTarget.style.background = "rgba(14,143,160,0.07)")}
                  onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
                >
                  <td style={{ padding: "10px 14px" }}><PlateBadge plate={r.plate} /></td>
                  <td style={{ padding: "10px 14px", fontSize: 12, color: "var(--text-muted)", fontFamily: "'JetBrains Mono', monospace" }}>{fmtDate(r.date)}</td>
                  <td style={{ padding: "10px 14px" }}>
                    <span className="rounded-full px-2 py-1 text-xs font-medium" style={{ background: "rgba(14,143,160,0.12)", color: "var(--accent-frost)", whiteSpace: "nowrap" }}>{r.type}</span>
                  </td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: "var(--text)" }}>{r.description}</td>
                  <td style={{ padding: "10px 14px", fontSize: 12, color: "var(--text-muted)", fontFamily: "'JetBrains Mono', monospace" }}>{refLabel("PR", r.pr_number)}</td>
                  <td style={{ padding: "10px 14px", fontSize: 12, color: "var(--text-muted)", fontFamily: "'JetBrains Mono', monospace" }}>{refLabel("PO", r.po_number)}</td>
                  <td style={{ padding: "10px 14px", fontSize: 12, color: "var(--text-muted)" }}>{r.garage}</td>
                  <td style={{ padding: "10px 14px", fontSize: 12, color: r.status === "เสร็จสิ้น" ? "#16A34A" : "#D97706" }}>{r.status}</td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: isPendingCost(r.cost) ? "#D97706" : "var(--text)", fontWeight: 600 }}>{fmtCost(r.cost)}</td>
                  <td style={{ padding: "10px 14px" }} onClick={(e) => e.stopPropagation()}>
                    <div className="flex items-center gap-1">
                      <button onClick={() => setEditTarget(r)} title="แก้ไขรายการซ่อม" style={iconBtnStyle}><Pencil size={15} /></button>
                      <button onClick={() => onGoToVehicle(r.plate)} title="ไปที่หน้ารถคันนี้" style={iconBtnStyle}><Truck size={15} /></button>
                      <button onClick={() => setDeleteTarget(r)} title="ลบ" style={{ ...iconBtnStyle, color: "#DC2626" }}><Trash2 size={15} /></button>
                    </div>
                  </td>
                </tr>
              ))}
              {filtered.length === 0 && (
                <tr><td colSpan={10} style={{ padding: "24px 14px", textAlign: "center", color: "var(--text-muted)", fontSize: 13 }}>ไม่พบรายการซ่อมที่ตรงกับตัวกรอง</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </Card>

      {editTarget && (
        <RepairFormModal
          plate={editTarget.plate}
          initial={editTarget}
          onClose={() => setEditTarget(null)}
          onSave={async (data) => { await onUpdateRepair(data); }}
        />
      )}

      {deleteTarget && (
        <ConfirmDelete
          label={`ต้องการลบรายการซ่อม "${deleteTarget.description}" ของรถ ${deleteTarget.plate} ใช่หรือไม่?`}
          onCancel={() => setDeleteTarget(null)}
          onConfirm={async () => { await onDeleteRepair(deleteTarget.id); setDeleteTarget(null); }}
        />
      )}
    </div>
  );
}


/* ---------------------------------------------------------------
   FUEL FORM (add + edit)
---------------------------------------------------------------- */

function FuelFormModal({ initial, vehicles, fuelLogs = [], onClose, onSave }) {
  const isEdit = !!initial;

  // ค่าเริ่มต้นจากการเติมครั้งล่าสุด (ราคา/ปั๊ม) เพื่อไม่ต้องกรอกซ้ำ
  const lastLog = useMemo(() => {
    if (!fuelLogs.length) return null;
    return [...fuelLogs].sort((a, b) => new Date(b.date) - new Date(a.date))[0];
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const recentPlates = useMemo(() => {
    const seen = [];
    [...fuelLogs].sort((a, b) => new Date(b.date) - new Date(a.date)).forEach((f) => {
      if (!seen.includes(f.plate) && vehicles.some((v) => v.id === f.plate)) seen.push(f.plate);
    });
    return seen.slice(0, 6);
  }, [fuelLogs, vehicles]);

  const initStation = initial ? (initial.station || "") : (lastLog && lastLog.station ? lastLog.station : (STATION_OPTIONS[0] || ""));
  const [form, setForm] = useState(() => initial ? {
    ...initial,
    total_cost: String(initial.total_cost ?? ""),
    price_per_liter: String(initial.price_per_liter ?? ""),
    odometer: initial.odometer === null || initial.odometer === undefined ? "" : String(initial.odometer),
    station: initStation, driver: initial.driver || "", note: initial.note || "",
  } : {
    plate: "", date: todayISO(), total_cost: "",
    price_per_liter: lastLog ? String(lastLog.price_per_liter) : "",
    odometer: "", station: initStation, driver: "", note: "",
  });
  const [stationCustom, setStationCustom] = useState(() => !!initStation && !STATION_OPTIONS.includes(initStation));
  const [showMore, setShowMore] = useState(() => !!(initial && (initial.odometer || initial.note)));
  const [plateQuery, setPlateQuery] = useState("");
  const [savedMsg, setSavedMsg] = useState("");
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);
  const plateInputRef = useRef(null);
  const amountRef = useRef(null);
  const update = (key, val) => setForm((f) => ({ ...f, [key]: val }));

  const selectedVehicle = vehicles.find((v) => v.id === form.plate);
  const yesterdayISO = new Date(Date.now() - 86400000).toISOString().slice(0, 10);

  const norm = (t) => String(t || "").toLowerCase().replace(/[^0-9a-zก-๙]/g, "");
  const matches = useMemo(() => {
    const q = norm(plateQuery);
    if (!q) return [];
    return vehicles.filter((v) => norm(v.id).includes(q) || norm(`${v.brand}${v.model}`).includes(q)).slice(0, 8);
  }, [plateQuery, vehicles]);

  const totalCostNum = Number(form.total_cost) || 0;
  const priceNum = Number(form.price_per_liter) || 0;
  const litersNum = priceNum > 0 ? Math.round((totalCostNum / priceNum) * 100) / 100 : 0;

  function pickPlate(v) {
    setForm((f) => ({ ...f, plate: v.id, driver: v.driver && v.driver !== "-" ? v.driver : "" }));
    setPlateQuery("");
    setError("");
    setTimeout(() => { if (amountRef.current) amountRef.current.focus(); }, 60);
  }

  async function save(addNext) {
    if (!form.plate) return setError("กรุณาเลือกรถก่อน (พิมพ์เลขทะเบียนค้นหาได้)");
    if (totalCostNum <= 0) return setError("กรุณากรอกจำนวนเงิน");
    if (priceNum <= 0) return setError("กรุณากรอกราคาน้ำมันต่อลิตร");
    setSaving(true);
    setError("");
    try {
      await onSave({
        id: isEdit ? initial.id : undefined,
        plate: form.plate, date: form.date,
        liters: litersNum, price_per_liter: priceNum, total_cost: totalCostNum,
        odometer: form.odometer === "" ? null : Number(form.odometer),
        station: form.station.trim(), driver: form.driver.trim(), note: form.note.trim(),
      });
      if (addNext && !isEdit) {
        setSavedMsg(`บันทึกแล้ว: ${form.plate} · ${litersNum.toLocaleString("th-TH", { maximumFractionDigits: 2 })} ลิตร · ${fmtMoney(totalCostNum)} บาท`);
        setForm((f) => ({ ...f, plate: "", total_cost: "", odometer: "", note: "", driver: "" }));
        setPlateQuery("");
        setTimeout(() => { if (plateInputRef.current) plateInputRef.current.focus(); }, 60);
      } else {
        onClose();
      }
    } catch (err) {
      setError(err.message || "บันทึกไม่สำเร็จ กรุณาลองใหม่อีกครั้ง");
    } finally {
      setSaving(false);
    }
  }

  return (
    <ModalShell title={isEdit ? "แก้ไขรายการเติมน้ำมัน" : "บันทึกการเติมน้ำมัน"} onClose={onClose} width={620}>
      <form onSubmit={(e) => { e.preventDefault(); save(!isEdit); }} className="flex flex-col gap-4">

        {savedMsg && (
          <div className="flex items-center gap-2 rounded-lg px-3 py-2" style={{ background: "rgba(22,163,74,0.12)", border: "1px solid #16A34A55", color: "#16A34A", fontSize: 13, fontWeight: 600 }}>
            <CheckCircle2 size={16} />{savedMsg}
          </div>
        )}

        {/* 1. เลือกรถ */}
        <div className="flex flex-col gap-2">
          <StepLabel n={1}>เติมให้รถคันไหน</StepLabel>
          {selectedVehicle ? (
            <div className="flex items-center justify-between rounded-lg px-3 py-2" style={{ background: "var(--surface-2)", border: "1px solid var(--accent-frost)" }}>
              <div className="flex items-center gap-3">
                <PlateBadge plate={selectedVehicle.id} />
                <span style={{ fontSize: 13, color: "var(--text-muted)" }}>{selectedVehicle.brand} {selectedVehicle.model}{form.driver ? ` · ${form.driver}` : ""}</span>
              </div>
              <button type="button" onClick={() => { update("plate", ""); setTimeout(() => { if (plateInputRef.current) plateInputRef.current.focus(); }, 60); }} style={{ fontSize: 12, color: "var(--accent-frost)", fontWeight: 600, background: "transparent", border: "none", cursor: "pointer" }}>
                เปลี่ยนรถ
              </button>
            </div>
          ) : (
            <>
              <div className="flex items-center gap-2" style={{ ...inputStyle, padding: "10px 12px" }}>
                <Search size={15} style={{ color: "var(--text-muted)" }} />
                <input
                  ref={plateInputRef} autoFocus={!isEdit}
                  value={plateQuery} onChange={(e) => setPlateQuery(e.target.value)}
                  onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); if (matches[0]) pickPlate(matches[0]); } }}
                  placeholder="พิมพ์เลขทะเบียน เช่น 9474"
                  style={{ background: "transparent", border: "none", outline: "none", color: "var(--text)", fontSize: 15, width: "100%" }}
                />
              </div>
              {matches.length > 0 && (
                <div className="flex flex-col gap-1" style={{ maxHeight: 220, overflowY: "auto" }}>
                  {matches.map((v) => (
                    <button key={v.id} type="button" onClick={() => pickPlate(v)} className="flex items-center gap-3 rounded-lg px-3 py-2" style={{ background: "var(--surface-2)", border: "1px solid var(--border)", textAlign: "left", cursor: "pointer" }}>
                      <PlateBadge plate={v.id} />
                      <span style={{ fontSize: 12, color: "var(--text-muted)" }}>{v.brand} {v.model}{v.driver && v.driver !== "-" ? ` · ${v.driver}` : ""}</span>
                    </button>
                  ))}
                </div>
              )}
              {!plateQuery && recentPlates.length > 0 && (
                <div className="flex flex-wrap items-center gap-2">
                  <span style={{ fontSize: 11, color: "var(--text-muted)" }}>ใช้บ่อย:</span>
                  {recentPlates.map((pl) => (
                    <TapChip key={pl} onClick={() => pickPlate(vehicles.find((v) => v.id === pl))} style={{ minHeight: 34, padding: "4px 12px", fontFamily: "'JetBrains Mono', monospace" }}>{pl}</TapChip>
                  ))}
                </div>
              )}
              {plateQuery && matches.length === 0 && <span style={{ fontSize: 12, color: "var(--text-muted)" }}>ไม่พบทะเบียนที่ตรงกัน</span>}
            </>
          )}
        </div>

        {/* 2. จำนวนเงิน + ราคา */}
        <div className="flex flex-col gap-2">
          <StepLabel n={2}>เติมเท่าไหร่</StepLabel>
          <div className="grid grid-cols-5 gap-3">
            <div className="flex items-center gap-2" style={{ ...inputStyle, gridColumn: "span 3", padding: "6px 12px" }}>
              <input
                ref={amountRef} type="number" inputMode="decimal" min="0" step="0.01"
                value={form.total_cost} onChange={(e) => update("total_cost", e.target.value)}
                placeholder="จำนวนเงิน"
                style={{ background: "transparent", border: "none", outline: "none", color: "var(--text)", fontSize: 22, fontWeight: 700, width: "100%" }}
              />
              <span style={{ fontSize: 13, color: "var(--text-muted)" }}>บาท</span>
            </div>
            <div className="flex items-center gap-1" style={{ ...inputStyle, gridColumn: "span 2", padding: "6px 10px" }}>
              <input
                type="number" inputMode="decimal" min="0" step="0.01"
                value={form.price_per_liter} onChange={(e) => update("price_per_liter", e.target.value)}
                placeholder="ราคา/ลิตร"
                style={{ background: "transparent", border: "none", outline: "none", color: "var(--text)", fontSize: 16, fontWeight: 600, width: "100%" }}
              />
              <span style={{ fontSize: 11, color: "var(--text-muted)", whiteSpace: "nowrap" }}>บ./ล.</span>
            </div>
          </div>
          <div className="flex flex-wrap gap-2">
            {[500, 1000, 1500, 2000, 3000, 5000].map((amt) => (
              <TapChip key={amt} active={Number(form.total_cost) === amt} onClick={() => update("total_cost", String(amt))} style={{ minHeight: 34, padding: "4px 12px" }}>{fmtMoney(amt)}</TapChip>
            ))}
          </div>
          <div className="flex items-center justify-between rounded-lg px-4 py-3" style={{ background: "rgba(14,143,160,0.08)", border: "1px solid #0E8FA055" }}>
            <span style={{ fontSize: 13, color: "var(--text-muted)" }}>จำนวนลิตร (คำนวณให้อัตโนมัติ)</span>
            <span style={{ fontSize: 20, fontWeight: 700, color: "var(--accent-frost)" }}>{litersNum.toLocaleString("th-TH", { maximumFractionDigits: 2 })} ลิตร</span>
          </div>
        </div>

        {/* 3. วันที่ + ปั๊ม */}
        <div className="flex flex-col gap-2">
          <StepLabel n={3}>วันที่และปั๊ม</StepLabel>
          <div className="flex flex-wrap items-center gap-2">
            <input style={{ ...inputStyle, width: "auto", minWidth: 160, padding: "9px 10px" }} type="date" value={form.date} onChange={(e) => update("date", e.target.value)} />
            <TapChip active={form.date === todayISO()} onClick={() => update("date", todayISO())}>วันนี้</TapChip>
            <TapChip active={form.date === yesterdayISO} onClick={() => update("date", yesterdayISO)}>เมื่อวาน</TapChip>
          </div>
          <div className="flex flex-wrap gap-2">
            {STATION_OPTIONS.map((st) => (
              <TapChip key={st} active={!stationCustom && form.station === st} onClick={() => { setStationCustom(false); update("station", st); }}>{st}</TapChip>
            ))}
            <TapChip active={stationCustom} onClick={() => { setStationCustom(true); if (STATION_OPTIONS.includes(form.station)) update("station", ""); }}>อื่นๆ</TapChip>
          </div>
          {stationCustom && (
            <input style={{ ...inputStyle, padding: "10px 12px" }} placeholder="พิมพ์ชื่อปั๊ม/สถานที่เติม" value={form.station} onChange={(e) => update("station", e.target.value)} />
          )}
        </div>

        {/* รายละเอียดเพิ่มเติม (ไม่บังคับ) */}
        <div className="flex flex-col gap-2">
          <button type="button" onClick={() => setShowMore((v) => !v)} className="flex items-center gap-1" style={{ fontSize: 13, color: "var(--accent-frost)", fontWeight: 600, background: "transparent", border: "none", cursor: "pointer", alignSelf: "flex-start", padding: 0 }}>
            <ChevronRight size={14} style={{ transform: showMore ? "rotate(90deg)" : "none", transition: "transform .15s" }} />
            เพิ่มเลขไมล์ / คนขับ / หมายเหตุ (ไม่บังคับ)
          </button>
          {showMore && (
            <div className="grid grid-cols-3 gap-3">
              <Field label="เลขไมล์"><input style={inputStyle} type="number" inputMode="numeric" placeholder="เช่น 182450" value={form.odometer} onChange={(e) => update("odometer", e.target.value)} /></Field>
              <Field label="คนขับที่เติม"><input style={inputStyle} placeholder="ชื่อ-นามสกุล" value={form.driver} onChange={(e) => update("driver", e.target.value)} /></Field>
              <Field label="หมายเหตุ"><input style={inputStyle} placeholder="ถ้ามี" value={form.note} onChange={(e) => update("note", e.target.value)} /></Field>
            </div>
          )}
        </div>

        {error && <div style={{ fontSize: 12, color: "#DC2626", background: "rgba(220,38,38,0.1)", border: "1px solid #DC262655", borderRadius: 8, padding: "8px 10px" }}>{error}</div>}

        <div className="flex items-center justify-end gap-2 flex-wrap" style={{ position: "sticky", bottom: 0, background: "var(--surface-solid, var(--surface))", paddingTop: 12, marginTop: 2, borderTop: "1px solid var(--border)" }}>
          <button type="button" onClick={onClose} style={{ ...inputStyle, width: "auto", padding: "10px 16px", cursor: "pointer" }}>{savedMsg ? "เสร็จแล้ว ปิด" : "ยกเลิก"}</button>
          {!isEdit && (
            <button type="button" disabled={saving} onClick={() => save(false)} style={{ ...inputStyle, width: "auto", padding: "10px 16px", cursor: "pointer", color: "var(--accent-frost)", fontWeight: 600, opacity: saving ? 0.6 : 1 }}>
              บันทึกแล้วปิด
            </button>
          )}
          <button type="submit" disabled={saving} style={{ background: "var(--accent-frost)", color: "#FFFFFF", border: "none", borderRadius: 8, padding: "10px 22px", fontSize: 14, fontWeight: 700, cursor: "pointer", opacity: saving ? 0.6 : 1 }}>
            {saving ? "กำลังบันทึก..." : isEdit ? "บันทึก" : "บันทึก + เพิ่มคันต่อไป"}
          </button>
        </div>
      </form>
    </ModalShell>
  );
}

/* ---------------------------------------------------------------
   นำเข้ารายงานเติมน้ำมันจากไฟล์ Excel (Transaction Report ของบัตรเติมน้ำมัน PTT)
---------------------------------------------------------------- */

const normText = (t) => String(t ?? "").replace(/\s+/g, " ").trim().toLowerCase();
const normPlate = (t) => String(t ?? "").toLowerCase().replace(/[^0-9a-zก-๙]/g, "");

// ชื่อปั๊มในรายงานเป็นรหัสย่อ เช่น "PTTST.D PULPERM OIL SERVI" -> แปลงเป็นชื่อที่ใช้ในระบบ
function mapStationName(name, location) {
  const n = String(name || "").trim();
  const loc = String(location || "").trim();
  if (/ลาดหลุมแก้ว/.test(loc) || /pulperm/i.test(n)) return STATION_OPTIONS[0] || n;
  return n || loc;
}

function toISODate(v, XLSX) {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") {
    const d = XLSX.SSF.parse_date_code(v);
    return d ? `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}` : null;
  }
  const t = String(v).trim();
  let m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;
  m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) {
    const y = Number(m[3]) > 2400 ? Number(m[3]) - 543 : Number(m[3]);
    return `${y}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  }
  return null;
}

async function parseFuelReport(file) {
  const XLSX = await import("xlsx");
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
  return parseFuelRows(rows, XLSX);
}

// แปลงตาราง (แถว x คอลัมน์) ของรายงานเป็นรายการเติมน้ำมัน
function parseFuelRows(rows, XLSX) {
  const hIdx = rows.findIndex((r) => r.some((c) => normText(c) === "plate no."));
  if (hIdx < 0) throw new Error("ไม่พบหัวตาราง \"Plate No.\" ในไฟล์ ใช่รายงาน Transaction Report ของบัตรเติมน้ำมันหรือไม่?");

  const heads = rows[hIdx].map(normText);
  const col = (pred) => heads.findIndex(pred);
  const idx = {
    plate: col((h) => h === "plate no."),
    date: col((h) => h === "date"),
    time: col((h) => h === "time"),
    ref: col((h) => h.startsWith("tax invoice")),
    qty: col((h) => h.startsWith("quantity")),
    price: col((h) => h === "price"),
    amount: col((h) => h.startsWith("amount")),
    odo: col((h) => h === "odometer"),
    station: col((h) => h === "station name"),
    location: col((h) => h === "location"),
  };
  const labels = { plate: "Plate No.", date: "Date", qty: "Quantity", price: "Price", amount: "Amount (THB)" };
  const missing = Object.keys(labels).filter((k) => idx[k] < 0).map((k) => labels[k]);
  if (missing.length) throw new Error(`ไฟล์ขาดคอลัมน์ที่จำเป็น: ${missing.join(", ")}`);

  const num = (v) => {
    if (typeof v === "number") return v;
    const n = Number(String(v ?? "").replace(/,/g, "").trim());
    return Number.isFinite(n) ? n : NaN;
  };
  const get = (r, k) => (idx[k] >= 0 ? r[idx[k]] : "");

  const out = [];
  rows.slice(hIdx + 1).forEach((r, i) => {
    const plateRaw = String(get(r, "plate") ?? "").trim();
    if (!plateRaw) return;
    const date = toISODate(get(r, "date"), XLSX);
    const liters = num(get(r, "qty"));
    const amount = num(get(r, "amount"));
    let price = num(get(r, "price"));
    if (!(price > 0) && liters > 0 && amount > 0) price = Math.round((amount / liters) * 100) / 100;
    const time = String(get(r, "time") ?? "").trim();
    let ref = typeof get(r, "ref") === "number" ? "" : String(get(r, "ref") ?? "").trim();
    // ไม่มีเลขใบกำกับ: สร้างรหัสจากทะเบียน+วันเวลา+ยอด เพื่อให้นำเข้าซ้ำแล้วยังตรวจจับได้
    if (!ref && date) ref = `auto:${normPlate(plateRaw)}:${date}:${time}:${amount}`;
    const odo = num(get(r, "odo"));
    out.push({
      key: `${i}-${ref}`,
      line: hIdx + 2 + i,
      plateRaw, date, liters, price, amount, ref,
      odometer: odo > 0 ? odo : null, // ในรายงานบางรายการเป็น 0 = ไม่ได้กรอกเลขไมล์
      station: mapStationName(get(r, "station"), get(r, "location")),
      invalid: !date || !(liters > 0) || !(amount > 0),
    });
  });
  return out;
}

const IMPORT_STATUS = {
  new: { label: "ใหม่", color: "#16A34A" },
  dup: { label: "มีในระบบแล้ว (ข้าม)", color: "#64748B" },
  maybe: { label: "อาจซ้ำกับที่คีย์เอง", color: "#D97706" },
  no_plate: { label: "ไม่พบทะเบียนในระบบ", color: "#DC2626" },
  invalid: { label: "ข้อมูลไม่ครบ", color: "#DC2626" },
};

function FuelImportModal({ vehicles, fuelLogs, onClose, onImport }) {
  const [fileName, setFileName] = useState("");
  const [parsed, setParsed] = useState(null); // รายการดิบจากไฟล์
  const [selected, setSelected] = useState({});
  const [parsing, setParsing] = useState(false);
  const [importing, setImporting] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);

  const plateMap = useMemo(() => {
    const m = {};
    vehicles.forEach((v) => { m[normPlate(v.id)] = v; });
    return m;
  }, [vehicles]);

  // จัดสถานะแต่ละแถวเทียบกับข้อมูลที่มีอยู่ในระบบ
  const rows = useMemo(() => {
    if (!parsed) return [];
    const refs = new Set(fuelLogs.map((f) => f.ref_no).filter(Boolean));
    const manual = new Set(fuelLogs.filter((f) => !f.ref_no).map((f) => `${f.plate}|${f.date}|${Math.round(Number(f.total_cost))}`));
    return parsed.map((r) => {
      const vehicle = plateMap[normPlate(r.plateRaw)] || null;
      let status = "new";
      if (r.invalid) status = "invalid";
      else if (!vehicle) status = "no_plate";
      else if (r.ref && refs.has(r.ref)) status = "dup";
      else if (manual.has(`${vehicle.id}|${r.date}|${Math.round(r.amount)}`)) status = "maybe";
      return { ...r, vehicle, status };
    });
  }, [parsed, fuelLogs, plateMap]);

  async function handleFile(file) {
    if (!file) return;
    setError(""); setResult(null); setParsing(true); setFileName(file.name);
    try {
      const list = await parseFuelReport(file);
      if (!list.length) throw new Error("ไม่พบรายการเติมน้ำมันในไฟล์");
      setParsed(list);
    } catch (e) {
      setParsed(null);
      setError(e.message || "อ่านไฟล์ไม่สำเร็จ");
    } finally {
      setParsing(false);
    }
  }

  // ตั้งค่าเริ่มต้น: ติ๊กเฉพาะรายการ "ใหม่"
  useEffect(() => {
    const sel = {};
    rows.forEach((r) => { if (r.status === "new") sel[r.key] = true; });
    setSelected(sel);
  }, [parsed]); // eslint-disable-line react-hooks/exhaustive-deps

  const counts = rows.reduce((c, r) => { c[r.status] = (c[r.status] || 0) + 1; return c; }, {});
  const chosen = rows.filter((r) => selected[r.key] && (r.status === "new" || r.status === "maybe"));
  const chosenLiters = chosen.reduce((s, r) => s + r.liters, 0);
  const chosenAmount = chosen.reduce((s, r) => s + r.amount, 0);
  const unmatched = Object.entries(rows.filter((r) => r.status === "no_plate").reduce((m, r) => { m[r.plateRaw] = (m[r.plateRaw] || 0) + 1; return m; }, {}));

  async function doImport() {
    setImporting(true); setError("");
    try {
      const payload = chosen.map((r) => ({
        plate: r.vehicle.id, date: r.date,
        liters: r.liters, price_per_liter: r.price, total_cost: r.amount,
        odometer: r.odometer, station: r.station,
        driver: r.vehicle.driver && r.vehicle.driver !== "-" ? r.vehicle.driver : "",
        note: "", ref_no: r.ref,
      }));
      const inserted = await onImport(payload);
      setResult({ inserted, skipped: payload.length - inserted });
      setParsed(null);
    } catch (e) {
      const msg = String(e.message || "");
      setError(/ref_no/.test(msg) ? "ฐานข้อมูลยังไม่พร้อมรับการนำเข้า กรุณารัน migration_v9.sql ใน Supabase ก่อน" : (msg || "นำเข้าไม่สำเร็จ") + " (รายการที่นำเข้าไปแล้วบางส่วนจะไม่ซ้ำ ถ้าอัปโหลดไฟล์เดิมอีกครั้ง)");
    } finally {
      setImporting(false);
    }
  }

  return (
    <ModalShell title="นำเข้าการเติมน้ำมันจากไฟล์ Excel" onClose={onClose} width={1000}>
      <div className="flex flex-col gap-4">

        {result ? (
          <div className="flex flex-col gap-3">
            <div className="flex items-center gap-2 rounded-lg px-4 py-3" style={{ background: "rgba(22,163,74,0.12)", border: "1px solid #16A34A55", color: "#16A34A", fontWeight: 700 }}>
              <CheckCircle2 size={18} />นำเข้าสำเร็จ {result.inserted} รายการ{result.skipped > 0 ? ` (ข้ามที่ซ้ำ ${result.skipped} รายการ)` : ""}
            </div>
            <div className="flex justify-end gap-2">
              <button onClick={() => { setResult(null); setFileName(""); }} style={{ ...inputStyle, width: "auto", padding: "10px 16px", cursor: "pointer" }}>นำเข้าไฟล์อื่น</button>
              <button onClick={onClose} style={{ background: "var(--accent-frost)", color: "#FFFFFF", border: "none", borderRadius: 8, padding: "10px 22px", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>เสร็จสิ้น</button>
            </div>
          </div>
        ) : (
          <>
            <label className="flex flex-col items-center justify-center gap-2 rounded-xl" style={{ border: "2px dashed var(--border)", background: "var(--surface-2)", padding: "22px 16px", cursor: "pointer", textAlign: "center" }}>
              <Upload size={26} style={{ color: "var(--accent-frost)" }} />
              <span style={{ fontSize: 14, fontWeight: 700, color: "var(--text)" }}>{parsing ? "กำลังอ่านไฟล์..." : fileName ? `ไฟล์: ${fileName} (คลิกเพื่อเลือกไฟล์อื่น)` : "คลิกเลือกไฟล์ Excel รายงานการเติมน้ำมัน"}</span>
              <span style={{ fontSize: 12, color: "var(--text-muted)" }}>รองรับรายงาน Transaction Report ของบัตรเติมน้ำมัน (.xlsx) ระบบจะแสดงตัวอย่างให้ตรวจก่อน ยังไม่บันทึกจนกว่าจะกดยืนยัน</span>
              <input type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={(e) => { handleFile(e.target.files && e.target.files[0]); e.target.value = ""; }} />
            </label>

            {error && <div style={{ fontSize: 13, color: "#DC2626", background: "rgba(220,38,38,0.1)", border: "1px solid #DC262655", borderRadius: 8, padding: "10px 12px" }}>{error}</div>}

            {parsed && (
              <>
                <div className="flex flex-wrap items-center gap-2">
                  <span style={{ fontSize: 13, color: "var(--text)", fontWeight: 700 }}>พบ {rows.length} รายการ:</span>
                  {Object.keys(IMPORT_STATUS).filter((k) => counts[k]).map((k) => (
                    <span key={k} className="inline-flex items-center rounded-full px-3 py-1" style={{ background: `${IMPORT_STATUS[k].color}1E`, color: IMPORT_STATUS[k].color, border: `1px solid ${IMPORT_STATUS[k].color}55`, fontSize: 12, fontWeight: 600 }}>
                      {IMPORT_STATUS[k].label} {counts[k]}
                    </span>
                  ))}
                </div>

                {unmatched.length > 0 && (
                  <div style={{ fontSize: 12, color: "#DC2626", background: "rgba(220,38,38,0.08)", border: "1px solid #DC262640", borderRadius: 8, padding: "8px 12px" }}>
                    ทะเบียนที่ไม่พบในระบบ: {unmatched.map(([p, n]) => `${p} (${n} รายการ)`).join(", ")} — ให้เพิ่มรถคันนี้ในแท็บ "รถทั้งหมด" ก่อน แล้วอัปโหลดไฟล์เดิมซ้ำได้เลย (รายการที่นำเข้าไปแล้วจะถูกข้ามให้เอง)
                  </div>
                )}

                <div style={{ maxHeight: 340, overflow: "auto", border: "1px solid var(--border)", borderRadius: 10 }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 820 }}>
                    <thead style={{ position: "sticky", top: 0, background: "var(--surface-2)" }}>
                      <tr>
                        <th style={{ padding: "8px 10px", width: 36 }}>
                          <input type="checkbox"
                            checked={rows.some((r) => r.status === "new") && rows.filter((r) => r.status === "new").every((r) => selected[r.key])}
                            onChange={(e) => { const sel = { ...selected }; rows.forEach((r) => { if (r.status === "new") sel[r.key] = e.target.checked; }); setSelected(sel); }}
                            title="เลือก/ยกเลิกทุกรายการใหม่"
                          />
                        </th>
                        {["วันที่", "ทะเบียน", "ลิตร", "ราคา/ลิตร", "ยอดเงิน", "ปั๊ม", "สถานะ"].map((h) => (
                          <th key={h} style={{ textAlign: "left", padding: "8px 10px", fontSize: 12, color: "var(--text-muted)", fontWeight: 600 }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {rows.map((r) => {
                        const st = IMPORT_STATUS[r.status];
                        const canPick = r.status === "new" || r.status === "maybe";
                        return (
                          <tr key={r.key} style={{ borderTop: "1px solid var(--border)", opacity: canPick ? 1 : 0.6 }}>
                            <td style={{ padding: "6px 10px" }}>
                              <input type="checkbox" disabled={!canPick} checked={!!selected[r.key] && canPick} onChange={(e) => setSelected((s) => ({ ...s, [r.key]: e.target.checked }))} />
                            </td>
                            <td style={{ padding: "6px 10px", fontSize: 12, color: "var(--text-muted)", fontFamily: "'JetBrains Mono', monospace" }}>{r.date ? fmtDate(r.date) : "-"}</td>
                            <td style={{ padding: "6px 10px" }}>
                              {r.vehicle ? <PlateBadge plate={r.vehicle.id} /> : <span style={{ fontSize: 13, color: "#DC2626", fontFamily: "'JetBrains Mono', monospace" }}>{r.plateRaw}</span>}
                            </td>
                            <td style={{ padding: "6px 10px", fontSize: 13, color: "var(--text)" }}>{Number.isFinite(r.liters) ? r.liters.toLocaleString("th-TH", { maximumFractionDigits: 2 }) : "-"}</td>
                            <td style={{ padding: "6px 10px", fontSize: 13, color: "var(--text-muted)" }}>{Number.isFinite(r.price) ? r.price.toFixed(2) : "-"}</td>
                            <td style={{ padding: "6px 10px", fontSize: 13, color: "var(--text)", fontWeight: 600 }}>{Number.isFinite(r.amount) ? fmtMoney(r.amount) : "-"}</td>
                            <td style={{ padding: "6px 10px", fontSize: 12, color: "var(--text-muted)" }}>{r.station || "-"}</td>
                            <td style={{ padding: "6px 10px" }}>
                              <span style={{ fontSize: 11, fontWeight: 600, color: st.color }}>{st.label}</span>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>

                <div className="flex items-center justify-between flex-wrap gap-3" style={{ position: "sticky", bottom: 0, background: "var(--surface-solid, var(--surface))", paddingTop: 10, borderTop: "1px solid var(--border)" }}>
                  <span style={{ fontSize: 13, color: "var(--text-muted)" }}>
                    เลือกนำเข้า <b style={{ color: "var(--text)" }}>{chosen.length}</b> รายการ · <b style={{ color: "var(--text)" }}>{chosenLiters.toLocaleString("th-TH", { maximumFractionDigits: 2 })}</b> ลิตร · <b style={{ color: "var(--accent-frost)" }}>{fmtMoney(Math.round(chosenAmount))}</b> บาท
                  </span>
                  <div className="flex items-center gap-2">
                    <button onClick={onClose} style={{ ...inputStyle, width: "auto", padding: "10px 16px", cursor: "pointer" }}>ยกเลิก</button>
                    <button onClick={doImport} disabled={importing || chosen.length === 0} style={{ background: "var(--accent-frost)", color: "#FFFFFF", border: "none", borderRadius: 8, padding: "10px 22px", fontSize: 14, fontWeight: 700, cursor: "pointer", opacity: importing || chosen.length === 0 ? 0.5 : 1 }}>
                      {importing ? "กำลังนำเข้า..." : `นำเข้า ${chosen.length} รายการ`}
                    </button>
                  </div>
                </div>
              </>
            )}
          </>
        )}
      </div>
    </ModalShell>
  );
}

/* ---------------------------------------------------------------
   FUEL VIEW (บันทึก + สรุปการเติมน้ำมัน)
---------------------------------------------------------------- */

// คืนค่าวันจันทร์ของสัปดาห์นั้น (ใช้จัดกลุ่มรายสัปดาห์)
function startOfWeek(dateStr) {
  const d = new Date(dateStr);
  const day = (d.getDay() + 6) % 7; // จันทร์ = 0
  d.setDate(d.getDate() - day);
  return d.toISOString().slice(0, 10);
}

function FuelView({ vehicles, fuelLogs, onAdd, onUpdate, onDelete, onImport, dark }) {
  const [showImport, setShowImport] = useState(false);
  const [plateFilter, setPlateFilter] = useState("all");
  const [monthFilter, setMonthFilter] = useState("all");
  const [groupBy, setGroupBy] = useState("month"); // day | week | month | plate
  const [formState, setFormState] = useState(null);
  const [deleteTarget, setDeleteTarget] = useState(null);

  const monthOptions = useMemo(() => {
    const set = new Set();
    fuelLogs.forEach((f) => {
      const d = new Date(f.date);
      set.add(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`);
    });
    return Array.from(set).sort().reverse();
  }, [fuelLogs]);

  const monthLabel = (key) => {
    const [y, m] = key.split("-");
    return `${THAI_MONTHS[Number(m) - 1]} ${Number(y) + 543}`;
  };

  const filtered = useMemo(() => {
    let list = [...fuelLogs];
    if (plateFilter !== "all") list = list.filter((f) => f.plate === plateFilter);
    if (monthFilter !== "all") {
      list = list.filter((f) => {
        const d = new Date(f.date);
        return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}` === monthFilter;
      });
    }
    return list.sort((a, b) => new Date(b.date) - new Date(a.date));
  }, [fuelLogs, plateFilter, monthFilter]);

  const totalLiters = filtered.reduce((s, f) => s + Number(f.liters || 0), 0);
  const totalCost = filtered.reduce((s, f) => s + Number(f.total_cost || 0), 0);
  const avgPrice = totalLiters > 0 ? totalCost / totalLiters : 0;

  // จัดกลุ่มตามที่เลือก
  const grouped = useMemo(() => {
    const map = {};
    filtered.forEach((f) => {
      let key, label;
      const d = new Date(f.date);
      if (groupBy === "day") { key = f.date; label = fmtDate(f.date); }
      else if (groupBy === "week") { key = startOfWeek(f.date); label = `สัปดาห์ของ ${fmtDate(key)}`; }
      else if (groupBy === "month") { key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`; label = monthLabel(key); }
      else { key = f.plate; label = f.plate; }
      if (!map[key]) map[key] = { key, label, liters: 0, cost: 0, count: 0 };
      map[key].liters += Number(f.liters || 0);
      map[key].cost += Number(f.total_cost || 0);
      map[key].count += 1;
    });
    const arr = Object.values(map);
    if (groupBy === "plate") arr.sort((a, b) => b.cost - a.cost);
    else arr.sort((a, b) => (a.key < b.key ? 1 : -1));
    return arr;
  }, [filtered, groupBy]);

  const chartData = useMemo(() => [...grouped].reverse().slice(-12).map((g) => ({ name: g.label.replace("สัปดาห์ของ ", ""), ลิตร: Math.round(g.liters), บาท: Math.round(g.cost) })), [grouped]);

  return (
    <div>
      <div className="flex items-center justify-between flex-wrap gap-3 mb-4">
        <SectionTitle icon={Fuel} sub="บันทึกการเติมน้ำมันและดูสรุปยอดรายวัน / รายสัปดาห์ / รายเดือน / รายคัน">
          การเติมน้ำมัน
        </SectionTitle>
        <div className="flex items-center gap-2 flex-wrap">
          <button onClick={() => setShowImport(true)} className="flex items-center gap-2 rounded-lg px-4 py-2" style={{ background: "var(--surface-2)", border: "1px solid var(--accent-frost)", color: "var(--accent-frost)", fontSize: 13, fontWeight: 700, height: 38 }}>
            <Upload size={16} />นำเข้าจากไฟล์ Excel
          </button>
          <button onClick={() => setFormState("add")} className="flex items-center gap-2 rounded-lg px-4 py-2" style={{ background: "var(--accent-frost)", color: "#FFFFFF", fontSize: 13, fontWeight: 700, height: 38 }}>
            <Plus size={16} />บันทึกการเติมน้ำมัน
          </button>
        </div>
      </div>

      <div className="flex items-center gap-2 mb-4 flex-wrap">
        <select value={plateFilter} onChange={(e) => setPlateFilter(e.target.value)} style={{ ...inputStyle, width: "auto", minWidth: 150 }}>
          <option value="all">ทุกคัน</option>
          {vehicles.map((v) => <option key={v.id} value={v.id}>{v.id} · {v.brand} {v.model}</option>)}
        </select>
        <select value={monthFilter} onChange={(e) => setMonthFilter(e.target.value)} style={{ ...inputStyle, width: "auto", minWidth: 150 }}>
          <option value="all">ทุกเดือน</option>
          {monthOptions.map((m) => <option key={m} value={m}>{monthLabel(m)}</option>)}
        </select>
        <div className="flex items-center gap-1 rounded-lg p-1" style={{ background: "var(--surface-2)", border: "1px solid var(--border)" }}>
          {[["day", "รายวัน"], ["week", "รายสัปดาห์"], ["month", "รายเดือน"], ["plate", "รายคัน"]].map(([k, label]) => (
            <button key={k} onClick={() => setGroupBy(k)} className="rounded-md px-3 py-1.5" style={{ fontSize: 12, fontWeight: 600, background: groupBy === k ? "var(--accent-frost)" : "transparent", color: groupBy === k ? "#FFFFFF" : "var(--text-muted)" }}>
              {label}
            </button>
          ))}
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-5">
        <Card style={{ padding: 18 }}>
          <div className="flex items-center gap-2 mb-2"><Fuel size={17} style={{ color: "var(--accent-frost)" }} /><span style={{ fontSize: 12, color: "var(--text-muted)" }}>ปริมาณน้ำมันรวม</span></div>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 27, fontWeight: 700, color: "var(--text)" }}>{totalLiters.toLocaleString("th-TH", { maximumFractionDigits: 2 })} <span style={{ fontSize: 14, color: "var(--text-muted)" }}>ลิตร</span></div>
        </Card>
        <Card style={{ padding: 18 }}>
          <div className="flex items-center gap-2 mb-2"><ClipboardList size={17} style={{ color: "#16A34A" }} /><span style={{ fontSize: 12, color: "var(--text-muted)" }}>ยอดเงินรวม</span></div>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 27, fontWeight: 700, color: "#16A34A" }}>{fmtMoney(Math.round(totalCost))} <span style={{ fontSize: 14, color: "var(--text-muted)" }}>บาท</span></div>
        </Card>
        <Card style={{ padding: 18 }}>
          <div className="flex items-center gap-2 mb-2"><Gauge size={17} style={{ color: "#D97706" }} /><span style={{ fontSize: 12, color: "var(--text-muted)" }}>ราคาเฉลี่ยต่อลิตร</span></div>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 27, fontWeight: 700, color: "#D97706" }}>{avgPrice.toFixed(2)} <span style={{ fontSize: 14, color: "var(--text-muted)" }}>บาท/ลิตร</span></div>
        </Card>
      </div>

      {chartData.length > 0 && (
        <Card style={{ padding: 20, marginBottom: 20 }}>
          <div style={{ fontSize: 13, color: "var(--text-muted)", marginBottom: 12, fontWeight: 600 }}>กราฟสรุปการเติมน้ำมัน</div>
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={chartData}>
              <CartesianGrid strokeDasharray="3 3" stroke={dark ? "#2C3B4E" : "#E2E8F0"} vertical={false} />
              <XAxis dataKey="name" tick={{ fill: dark ? "#93A4B8" : "#64748B", fontSize: 11 }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fill: dark ? "#93A4B8" : "#64748B", fontSize: 11 }} axisLine={false} tickLine={false} />
              <Tooltip contentStyle={{ background: dark ? "#1F2A3B" : "#FFFFFF", border: `1px solid ${dark ? "#334155" : "#E2E8F0"}`, borderRadius: 8, fontSize: 12 }} cursor={{ fill: dark ? "rgba(255,255,255,0.05)" : "rgba(15,23,42,0.05)" }} />
              <Bar dataKey="บาท" radius={[6, 6, 0, 0]} fill="#0E8FA0" barSize={26} />
            </BarChart>
          </ResponsiveContainer>
        </Card>
      )}

      <Card style={{ overflow: "hidden", marginBottom: 20 }}>
        <div style={{ padding: "12px 16px", borderBottom: "1px solid var(--border)", fontSize: 13, fontWeight: 700, color: "var(--text)" }}>
          สรุปยอด{{ day: "รายวัน", week: "รายสัปดาห์", month: "รายเดือน", plate: "รายคัน" }[groupBy]}
        </div>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 620 }}>
            <thead>
              <tr style={{ background: "var(--surface-2)" }}>
                {["ช่วง / ทะเบียน", "จำนวนครั้ง", "ลิตรรวม", "ยอดเงินรวม", "ราคาเฉลี่ย/ลิตร"].map((h) => (
                  <th key={h} style={{ textAlign: "left", padding: "10px 14px", fontSize: 12, color: "var(--text-muted)", fontWeight: 600 }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {grouped.map((g) => (
                <tr key={g.key} style={{ borderTop: "1px solid var(--border)" }}>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: "var(--text)" }}>
                    {groupBy === "plate" ? <PlateBadge plate={g.label} /> : g.label}
                  </td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: "var(--text-muted)" }}>{g.count} ครั้ง</td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: "var(--text)" }}>{g.liters.toLocaleString("th-TH", { maximumFractionDigits: 2 })} ลิตร</td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: "var(--accent-frost)", fontWeight: 700 }}>{fmtMoney(Math.round(g.cost))} บ.</td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: "var(--text-muted)" }}>{(g.liters > 0 ? g.cost / g.liters : 0).toFixed(2)} บ./ล.</td>
                </tr>
              ))}
              {grouped.length === 0 && <tr><td colSpan={5} style={{ padding: "24px 14px", textAlign: "center", color: "var(--text-muted)", fontSize: 13 }}>ยังไม่มีข้อมูลการเติมน้ำมัน</td></tr>}
            </tbody>
          </table>
        </div>
      </Card>

      <Card style={{ overflow: "hidden" }}>
        <div style={{ padding: "12px 16px", borderBottom: "1px solid var(--border)", fontSize: 13, fontWeight: 700, color: "var(--text)" }}>
          รายการเติมน้ำมันทั้งหมด ({filtered.length})
        </div>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 900 }}>
            <thead>
              <tr style={{ background: "var(--surface-2)" }}>
                {["ทะเบียน", "วันที่", "ลิตร", "ราคา/ลิตร", "ยอดเงิน", "เลขไมล์", "ปั๊ม", "คนขับ", "จัดการ"].map((h) => (
                  <th key={h} style={{ textAlign: "left", padding: "10px 14px", fontSize: 12, color: "var(--text-muted)", fontWeight: 600 }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map((f) => (
                <tr key={f.id} style={{ borderTop: "1px solid var(--border)" }}>
                  <td style={{ padding: "10px 14px" }}><PlateBadge plate={f.plate} /></td>
                  <td style={{ padding: "10px 14px", fontSize: 12, color: "var(--text-muted)", fontFamily: "'JetBrains Mono', monospace" }}>{fmtDate(f.date)}</td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: "var(--text)" }}>{Number(f.liters).toLocaleString("th-TH", { maximumFractionDigits: 2 })}</td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: "var(--text-muted)" }}>{Number(f.price_per_liter).toFixed(2)}</td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: "var(--text)", fontWeight: 600 }}>{fmtMoney(Math.round(f.total_cost))} บ.</td>
                  <td style={{ padding: "10px 14px", fontSize: 12, color: "var(--text-muted)" }}>{f.odometer ? Number(f.odometer).toLocaleString("th-TH") : "-"}</td>
                  <td style={{ padding: "10px 14px", fontSize: 12, color: "var(--text-muted)" }}>{f.station || "-"}</td>
                  <td style={{ padding: "10px 14px", fontSize: 12, color: "var(--text-muted)" }}>{f.driver || "-"}</td>
                  <td style={{ padding: "10px 14px" }}>
                    <div className="flex items-center gap-1">
                      <button onClick={() => setFormState(f)} title="แก้ไข" style={iconBtnStyle}><Pencil size={15} /></button>
                      <button onClick={() => setDeleteTarget(f)} title="ลบ" style={{ ...iconBtnStyle, color: "#DC2626" }}><Trash2 size={15} /></button>
                    </div>
                  </td>
                </tr>
              ))}
              {filtered.length === 0 && <tr><td colSpan={9} style={{ padding: "24px 14px", textAlign: "center", color: "var(--text-muted)", fontSize: 13 }}>ยังไม่มีรายการเติมน้ำมัน</td></tr>}
            </tbody>
          </table>
        </div>
      </Card>

      {showImport && (
        <FuelImportModal vehicles={vehicles} fuelLogs={fuelLogs} onClose={() => setShowImport(false)} onImport={onImport} />
      )}

      {formState && (
        <FuelFormModal
          initial={formState === "add" ? null : formState}
          vehicles={vehicles}
          fuelLogs={fuelLogs}
          onClose={() => setFormState(null)}
          onSave={async (data) => { if (formState === "add") await onAdd(data); else await onUpdate(data.id, data); }}
        />
      )}

      {deleteTarget && (
        <ConfirmDelete
          label={`ต้องการลบรายการเติมน้ำมันของ ${deleteTarget.plate} วันที่ ${fmtDate(deleteTarget.date)} ใช่หรือไม่?`}
          onCancel={() => setDeleteTarget(null)}
          onConfirm={async () => { await onDelete(deleteTarget.id); setDeleteTarget(null); }}
        />
      )}
    </div>
  );
}


/* ---------------------------------------------------------------
   READINESS VIEW (รถพร้อมใช้งาน / ไม่พร้อมใช้งาน + สินค้าค้างในตู้)
---------------------------------------------------------------- */

function ReadinessView({ vehicles, onGoToVehicle }) {
  const ready = vehicles.filter((v) => v.status === "ready");
  const notReady = vehicles.filter((v) => v.status === "not_ready");
  const withCargo = vehicles.filter((v) => v.has_cargo);

  return (
    <div>
      <SectionTitle icon={Box} sub="ดูภาพรวมว่ารถคันไหนพร้อมใช้งาน คันไหนติดปัญหาอะไร และคันไหนยังมีสินค้า/ตะกร้าค้างอยู่ในตู้">
        ความพร้อมใช้งานของรถ
      </SectionTitle>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
        <Card style={{ padding: 18 }}>
          <div className="flex items-center gap-2 mb-1"><CheckCircle2 size={17} style={{ color: "#16A34A" }} /><span style={{ fontSize: 12, color: "var(--text-muted)" }}>พร้อมใช้งาน</span></div>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 27, fontWeight: 700, color: "#16A34A" }}>{ready.length} <span style={{ fontSize: 14, color: "var(--text-muted)" }}>คัน</span></div>
        </Card>
        <Card style={{ padding: 18 }}>
          <div className="flex items-center gap-2 mb-1"><AlertTriangle size={17} style={{ color: "#DC2626" }} /><span style={{ fontSize: 12, color: "var(--text-muted)" }}>ไม่พร้อมใช้งาน</span></div>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 27, fontWeight: 700, color: "#DC2626" }}>{notReady.length} <span style={{ fontSize: 14, color: "var(--text-muted)" }}>คัน</span></div>
        </Card>
        <Card style={{ padding: 18 }}>
          <div className="flex items-center gap-2 mb-1"><Package size={17} style={{ color: "#D97706" }} /><span style={{ fontSize: 12, color: "var(--text-muted)" }}>มีสินค้า/ตะกร้าค้างในตู้</span></div>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 27, fontWeight: 700, color: "#D97706" }}>{withCargo.length} <span style={{ fontSize: 14, color: "var(--text-muted)" }}>คัน</span></div>
        </Card>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <Card style={{ padding: 20 }}>
          <div className="flex items-center gap-2 mb-4"><CheckCircle2 size={16} style={{ color: "#16A34A" }} /><span style={{ fontSize: 14, fontWeight: 700, color: "var(--text)" }}>รถที่พร้อมใช้งาน ({ready.length})</span></div>
          <div className="flex flex-wrap gap-2">
            {ready.length === 0 && <p style={{ color: "var(--text-muted)", fontSize: 13 }}>ไม่มีรถที่พร้อมใช้งานตอนนี้</p>}
            {ready.map((v) => (
              <div key={v.id} onClick={() => onGoToVehicle(v.id)} className="flex items-center gap-2 rounded-lg px-2 py-2" style={{ background: "var(--surface-2)", cursor: "pointer", border: "1px solid var(--border)" }}>
                <PlateBadge plate={v.id} />
                {v.has_cargo && (
                  <span title={v.cargo_note} className="inline-flex items-center gap-1 rounded-full px-2 py-1" style={{ background: "rgba(217,119,6,0.12)", color: "#D97706", fontSize: 10, fontWeight: 600 }}>
                    <Package size={11} />มีของค้าง
                  </span>
                )}
              </div>
            ))}
          </div>
        </Card>

        <Card style={{ padding: 20 }}>
          <div className="flex items-center gap-2 mb-4"><AlertTriangle size={16} style={{ color: "#DC2626" }} /><span style={{ fontSize: 14, fontWeight: 700, color: "var(--text)" }}>รถที่ไม่พร้อมใช้งาน ({notReady.length})</span></div>
          <div className="flex flex-col gap-2">
            {notReady.length === 0 && <p style={{ color: "var(--text-muted)", fontSize: 13 }}>ไม่มีรถที่ติดปัญหาตอนนี้</p>}
            {notReady.map((v) => (
              <div key={v.id} onClick={() => onGoToVehicle(v.id)} className="rounded-lg px-3 py-2.5" style={{ background: "var(--surface-2)", cursor: "pointer", borderLeft: "3px solid #DC2626" }}>
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-3">
                    <PlateBadge plate={v.id} />
                    <span style={{ fontSize: 12, color: "var(--text-muted)" }}>{v.brand} {v.model}</span>
                  </div>
                  <ChevronRight size={15} style={{ color: "var(--text-muted)" }} />
                </div>
                <div style={{ fontSize: 12, color: "#DC2626", marginTop: 6 }}>ปัญหา: {v.reason || "ไม่ได้ระบุ"}</div>
                {v.has_cargo ? (
                  <div className="flex items-center gap-1 mt-1" style={{ fontSize: 12, color: "#D97706" }}>
                    <Package size={12} />มีสินค้า/ตะกร้าค้างในตู้: {v.cargo_note || "ไม่ได้ระบุรายละเอียด"}
                  </div>
                ) : (
                  <div className="flex items-center gap-1 mt-1" style={{ fontSize: 12, color: "var(--text-muted)" }}>
                    <Box size={12} />ตู้ว่าง ไม่มีสินค้าค้าง
                  </div>
                )}
              </div>
            ))}
          </div>
        </Card>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   PURCHASE REQUESTS (ใบขอซื้อ PR) - ไม่ต้องระบุทะเบียนก็ได้
---------------------------------------------------------------- */

// เลือกรถแบบพิมพ์ค้นหา (ไม่บังคับ: ไม่เลือก = ซื้อทั่วไป)
function PlatePicker({ vehicles, value, onChange }) {
  const [query, setQuery] = useState("");
  const selected = vehicles.find((v) => v.id === value);
  const matches = useMemo(() => {
    const q = normPlate(query);
    if (!q) return [];
    return vehicles.filter((v) => normPlate(v.id).includes(q) || normPlate(`${v.brand}${v.model}`).includes(q)).slice(0, 6);
  }, [query, vehicles]);

  if (selected) {
    return (
      <div className="flex items-center justify-between rounded-lg px-3 py-2" style={{ background: "var(--surface-2)", border: "1px solid var(--accent-frost)" }}>
        <div className="flex items-center gap-3">
          <PlateBadge plate={selected.id} />
          <span style={{ fontSize: 13, color: "var(--text-muted)" }}>{selected.brand} {selected.model}</span>
        </div>
        <button type="button" onClick={() => onChange("")} style={{ fontSize: 12, color: "var(--accent-frost)", fontWeight: 600, background: "transparent", border: "none", cursor: "pointer" }}>
          เอาออก / เปลี่ยน
        </button>
      </div>
    );
  }
  return (
    <div className="flex flex-col gap-2">
      <div className="flex items-center gap-2" style={{ ...inputStyle, padding: "10px 12px" }}>
        <Search size={15} style={{ color: "var(--text-muted)" }} />
        <input
          value={query} onChange={(e) => setQuery(e.target.value)}
          onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); if (matches[0]) { onChange(matches[0].id); setQuery(""); } } }}
          placeholder="พิมพ์เลขทะเบียนถ้าเกี่ยวกับรถคันใดคันหนึ่ง (ไม่ต้องก็ได้)"
          style={{ background: "transparent", border: "none", outline: "none", color: "var(--text)", fontSize: 14, width: "100%" }}
        />
      </div>
      {matches.length > 0 && (
        <div className="flex flex-col gap-1">
          {matches.map((v) => (
            <button key={v.id} type="button" onClick={() => { onChange(v.id); setQuery(""); }} className="flex items-center gap-3 rounded-lg px-3 py-2" style={{ background: "var(--surface-2)", border: "1px solid var(--border)", textAlign: "left", cursor: "pointer" }}>
              <PlateBadge plate={v.id} />
              <span style={{ fontSize: 12, color: "var(--text-muted)" }}>{v.brand} {v.model}</span>
            </button>
          ))}
        </div>
      )}
      {!query && <span style={{ fontSize: 11, color: "var(--text-muted)" }}>ไม่ระบุทะเบียน = ซื้อทั่วไป ไม่เกี่ยวกับรถคันใดคันหนึ่ง</span>}
    </div>
  );
}

function PRFormModal({ initial, vehicles, onClose, onSave }) {
  const isEdit = !!initial;
  const [form, setForm] = useState(() => initial ? {
    ...initial,
    pr_number: refNumber("PR", initial.pr_number), description: initial.description || "",
    category: initial.category || "อื่นๆ", plate: initial.plate || "", vendor: initial.vendor || "",
    estimated_cost: initial.estimated_cost === null || initial.estimated_cost === undefined ? "" : String(initial.estimated_cost),
    needed_by: initial.needed_by || "", po_number: refNumber("PO", initial.po_number),
    requester: initial.requester || "", note: initial.note || "",
  } : {
    pr_number: "", request_date: todayISO(), description: "", category: PR_CATEGORIES[0], plate: "",
    vendor: "", estimated_cost: "", needed_by: "", status: "รออนุมัติ", po_number: "", requester: "", note: "",
  });
  const [showMore, setShowMore] = useState(() => !!(initial && (initial.requester || initial.note)));
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);
  const update = (key, val) => setForm((f) => ({ ...f, [key]: val }));
  const descMissing = !!error && !form.description.trim();

  async function handleSubmit(e) {
    e.preventDefault();
    if (!form.description.trim()) return setError("กรุณากรอกว่าขอซื้ออะไร");
    if (!form.request_date) return setError("กรุณาเลือกวันที่เปิด PR");
    setSaving(true);
    setError("");
    try {
      await onSave({
        id: isEdit ? initial.id : undefined,
        pr_number: form.pr_number.trim(), request_date: form.request_date,
        description: form.description.trim(), category: form.category,
        plate: form.plate || null, vendor: form.vendor.trim(),
        estimated_cost: form.estimated_cost === "" ? null : Number(form.estimated_cost),
        needed_by: form.needed_by || null, status: form.status,
        po_number: form.po_number.trim(), requester: form.requester.trim(), note: form.note.trim(),
      });
      onClose();
    } catch (err) {
      setError(err.message || "บันทึกไม่สำเร็จ กรุณาลองใหม่อีกครั้ง");
    } finally {
      setSaving(false);
    }
  }

  return (
    <ModalShell title={isEdit ? "แก้ไขใบขอซื้อ (PR)" : "เปิดใบขอซื้อ (PR) ใหม่"} onClose={onClose} width={640}>
      <form onSubmit={handleSubmit} className="flex flex-col gap-4">

        <div className="flex flex-col gap-2">
          <StepLabel n={1}>ขอซื้ออะไร</StepLabel>
          <div className="flex flex-wrap gap-2">
            {PR_CATEGORIES.map((c) => (
              <TapChip key={c} active={form.category === c} onClick={() => update("category", c)}>{c}</TapChip>
            ))}
          </div>
          <textarea
            rows={3} value={form.description} onChange={(e) => update("description", e.target.value)}
            placeholder="รายการที่ขอซื้อ พิมพ์ได้หลายรายการ เช่น ผ้าเบรกหน้า 2 ชุด, น้ำมันเบรก 4 ขวด"
            style={{ ...inputStyle, padding: "11px 12px", fontSize: 14, resize: "vertical", borderColor: descMissing ? "#DC2626" : "var(--border)" }}
          />
        </div>

        <div className="flex flex-col gap-2">
          <StepLabel n={2}>เกี่ยวกับรถคันไหน (ไม่บังคับ)</StepLabel>
          <PlatePicker vehicles={vehicles} value={form.plate} onChange={(v) => update("plate", v)} />
        </div>

        <div className="flex flex-col gap-2">
          <StepLabel n={3}>เลขที่เอกสารและราคา</StepLabel>
          <div className="grid grid-cols-2 gap-3">
            <Field label="เลขที่ PR"><RefInput prefix="PR" value={form.pr_number} onChange={(v) => update("pr_number", v)} /></Field>
            <Field label="เลขที่ PO"><RefInput prefix="PO" value={form.po_number} onChange={(v) => update("po_number", v)} /></Field>
          </div>
          <div className="grid grid-cols-2 gap-3">
            <Field label="วันที่เปิด PR *">
              <input style={inputStyle} type="date" value={form.request_date} onChange={(e) => update("request_date", e.target.value)} />
            </Field>
            <Field label="ต้องการใช้ภายในวันที่ (ไม่บังคับ)">
              <input style={inputStyle} type="date" value={form.needed_by} onChange={(e) => update("needed_by", e.target.value)} />
            </Field>
          </div>
          <div className="grid grid-cols-2 gap-3">
            <Field label="ราคาประมาณการ (บาท) เว้นว่าง = ยังไม่ทราบราคา">
              <input style={inputStyle} type="number" inputMode="decimal" min="0" step="0.01" placeholder="เช่น 8500" value={form.estimated_cost} onChange={(e) => update("estimated_cost", e.target.value)} />
            </Field>
            <Field label="ผู้ขาย / ร้านค้า"><input style={inputStyle} placeholder="ถ้าทราบ" value={form.vendor} onChange={(e) => update("vendor", e.target.value)} /></Field>
          </div>
        </div>

        <div className="flex flex-col gap-2">
          <StepLabel n={4}>สถานะ</StepLabel>
          <div className="flex flex-wrap gap-2">
            {PR_STATUSES.map((st) => (
              <TapChip key={st.key} color={st.color} active={form.status === st.key} onClick={() => update("status", st.key)}>{st.key}</TapChip>
            ))}
          </div>
        </div>

        <div className="flex flex-col gap-2">
          <button type="button" onClick={() => setShowMore((v) => !v)} className="flex items-center gap-1" style={{ fontSize: 13, color: "var(--accent-frost)", fontWeight: 600, background: "transparent", border: "none", cursor: "pointer", alignSelf: "flex-start", padding: 0 }}>
            <ChevronRight size={14} style={{ transform: showMore ? "rotate(90deg)" : "none", transition: "transform .15s" }} />
            เพิ่มผู้ขอซื้อ / หมายเหตุ (ไม่บังคับ)
          </button>
          {showMore && (
            <div className="grid grid-cols-2 gap-3">
              <Field label="ผู้ขอซื้อ"><input style={inputStyle} placeholder="ชื่อ-นามสกุล" value={form.requester} onChange={(e) => update("requester", e.target.value)} /></Field>
              <Field label="หมายเหตุ"><input style={inputStyle} placeholder="ถ้ามี" value={form.note} onChange={(e) => update("note", e.target.value)} /></Field>
            </div>
          )}
        </div>

        {error && <div style={{ fontSize: 12, color: "#DC2626", background: "rgba(220,38,38,0.1)", border: "1px solid #DC262655", borderRadius: 8, padding: "8px 10px" }}>{error}</div>}

        <div className="flex items-center justify-end gap-2" style={{ position: "sticky", bottom: 0, background: "var(--surface-solid, var(--surface))", paddingTop: 12, marginTop: 2, borderTop: "1px solid var(--border)" }}>
          <button type="button" onClick={onClose} style={{ ...inputStyle, width: "auto", padding: "10px 18px", cursor: "pointer" }}>ยกเลิก</button>
          <button type="submit" disabled={saving} style={{ background: "var(--accent-frost)", color: "#FFFFFF", border: "none", borderRadius: 8, padding: "10px 26px", fontSize: 14, fontWeight: 700, cursor: "pointer", opacity: saving ? 0.6 : 1 }}>
            {saving ? "กำลังบันทึก..." : "บันทึก"}
          </button>
        </div>
      </form>
    </ModalShell>
  );
}

const SOURCE_META = {
  pr: { label: "สั่งซื้อ", color: "#6366F1" },
  repair: { label: "งานซ่อม", color: "#0E8FA0" },
};
const REPAIR_STATUS_COLORS = { "กำลังดำเนินการ": "#D97706", "เสร็จสิ้น": "#16A34A" };

// หน้ารวม PR / PO ทั้งหมด: ใบขอซื้อ + งานซ่อมที่ใส่เลข PR/PO ไว้
function PurchaseRequestsView({ vehicles, prs, repairs, tableError, onAdd, onUpdate, onDelete, onUpdateRepair, onGoToVehicle }) {
  const [sourceFilter, setSourceFilter] = useState("all");
  const [statusFilter, setStatusFilter] = useState("all");
  const [categoryFilter, setCategoryFilter] = useState("all");
  const [query, setQuery] = useState("");
  const [formState, setFormState] = useState(null);   // แก้ไข/เปิด PR (สั่งซื้อ)
  const [repairEdit, setRepairEdit] = useState(null); // แก้ไขงานซ่อม
  const [deleteTarget, setDeleteTarget] = useState(null);
  const [actionError, setActionError] = useState("");

  const statusColor = (k) => (PR_STATUSES.find((s) => s.key === k) || {}).color || REPAIR_STATUS_COLORS[k] || "#64748B";
  const allStatuses = [...PR_STATUSES.map((s) => s.key), "กำลังดำเนินการ", "เสร็จสิ้น"];
  const allCategories = Array.from(new Set([...PR_CATEGORIES, ...REPAIR_TYPES]));

  // รวมทั้ง 2 แหล่งเป็นรายการเดียวกัน
  const items = useMemo(() => {
    const fromPR = prs.map((p) => ({
      uid: `pr-${p.id}`, source: "pr", raw: p, date: p.request_date, pr: p.pr_number, po: p.po_number,
      category: p.category || "อื่นๆ", description: p.description, plate: p.plate, party: p.vendor || "",
      amount: p.estimated_cost, status: p.status, neededBy: p.needed_by,
      closed: p.status === "รับของแล้ว" || p.status === "ยกเลิก",
      extra: [p.requester, p.note],
    }));
    const fromRepair = repairs.filter((r) => r.pr_number || r.po_number).map((r) => ({
      uid: `rp-${r.id}`, source: "repair", raw: r, date: r.date, pr: r.pr_number, po: r.po_number,
      category: r.type, description: r.description, plate: r.plate, party: r.garage && r.garage !== "-" ? r.garage : "",
      amount: r.cost, status: r.status, neededBy: null,
      closed: r.status === "เสร็จสิ้น",
      extra: [],
    }));
    return [...fromPR, ...fromRepair];
  }, [prs, repairs]);

  const prCount = items.filter((i) => i.source === "pr").length;
  const repairCount = items.length - prCount;

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    return items
      .filter((i) => sourceFilter === "all" || i.source === sourceFilter)
      .filter((i) => statusFilter === "all" || i.status === statusFilter)
      .filter((i) => categoryFilter === "all" || i.category === categoryFilter)
      .filter((i) => !q || [refLabel("PR", i.pr), refLabel("PO", i.po), i.pr, i.po, i.description, i.plate, i.party, i.category, ...i.extra]
        .some((t) => String(t || "").toLowerCase().includes(q)))
      .sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));
  }, [items, sourceFilter, statusFilter, categoryFilter, query]);

  const open = items.filter((i) => !i.closed);
  const openAmount = open.reduce((s, i) => s + (isPendingCost(i.amount) ? 0 : Number(i.amount || 0)), 0);
  const openNoPrice = open.filter((i) => isPendingCost(i.amount)).length;

  async function changeStatus(item, status) {
    setActionError("");
    try {
      if (item.source === "pr") await onUpdate({ ...item.raw, status });
      else await onUpdateRepair({ ...item.raw, status });
    } catch (e) {
      setActionError(e.message || "เปลี่ยนสถานะไม่สำเร็จ");
    }
  }

  async function handleExport() {
    const rows = filtered.map((i) => ({
      "ประเภท": SOURCE_META[i.source].label,
      "วันที่": fmtDate(i.date),
      "เลขที่ PR": i.pr ? refLabel("PR", i.pr) : "",
      "เลขที่ PO": i.po ? refLabel("PO", i.po) : "",
      "หมวด/ประเภท": i.category || "",
      "รายการ": i.description,
      "ทะเบียน": i.plate || "",
      "ผู้ขาย/อู่": i.party || "",
      "ยอดเงิน (บาท)": isPendingCost(i.amount) ? "ยังไม่ทราบราคา" : Number(i.amount),
      "ต้องการใช้ภายใน": i.neededBy ? fmtDate(i.neededBy) : "",
      "สถานะ": i.status,
    }));
    await exportToExcel(rows, `PR-PO-รวม-${todayISO()}.xlsx`, "PR PO รวม");
  }

  const kpis = [
    { label: "PR / PO ทั้งหมด", value: items.length, color: "var(--accent-frost)", icon: ClipboardList },
    { label: "ใบขอซื้อ (สั่งซื้อ)", value: prCount, color: SOURCE_META.pr.color, icon: ShoppingCart },
    { label: "งานซ่อมที่เปิด PR/PO", value: repairCount, color: SOURCE_META.repair.color, icon: Wrench },
  ];

  return (
    <div>
      <div className="flex items-center justify-between flex-wrap gap-3 mb-4">
        <SectionTitle icon={ShoppingCart} sub="รวมใบขอซื้อและงานซ่อมที่เปิด PR/PO ไว้ในหน้าเดียว ติดตามสถานะได้ทั้งหมด ใบขอซื้อไม่ต้องระบุทะเบียนก็ได้">
          PR / PO ทั้งหมด
        </SectionTitle>
        <div className="flex items-center gap-2 flex-wrap">
          <button onClick={handleExport} disabled={filtered.length === 0} className="flex items-center gap-2 rounded-lg px-4 py-2" style={{ background: "#16A34A", color: "#FFFFFF", fontSize: 13, fontWeight: 700, height: 38, opacity: filtered.length === 0 ? 0.6 : 1 }}>
            <Download size={16} />Excel
          </button>
          <button onClick={() => setFormState("add")} className="flex items-center gap-2 rounded-lg px-4 py-2" style={{ background: "var(--accent-frost)", color: "#FFFFFF", fontSize: 13, fontWeight: 700, height: 38 }}>
            <Plus size={16} />เปิด PR ใหม่
          </button>
        </div>
      </div>

      {tableError && (
        <div style={{ marginBottom: 14, fontSize: 13, color: "#D97706", background: "rgba(217,119,6,0.1)", border: "1px solid #D9770655", borderRadius: 8, padding: "10px 14px" }}>
          ส่วนใบขอซื้อยังใช้ไม่ได้: ฐานข้อมูลยังไม่มีตารางใบขอซื้อ กรุณารัน <b>migration_v10.sql</b> ใน Supabase ก่อน ({tableError}) ส่วนงานซ่อมที่เปิด PR/PO ยังแสดงได้ตามปกติ
        </div>
      )}
      {actionError && <div style={{ marginBottom: 14, fontSize: 12, color: "#DC2626", background: "rgba(220,38,38,0.1)", border: "1px solid #DC262655", borderRadius: 8, padding: "8px 12px" }}>{actionError}</div>}

      <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-5">
        {kpis.map((k) => (
          <Card key={k.label} style={{ padding: 16 }}>
            <div className="flex items-center gap-2 mb-1"><k.icon size={16} style={{ color: k.color }} /><span style={{ fontSize: 12, color: "var(--text-muted)" }}>{k.label}</span></div>
            <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 26, fontWeight: 700, color: k.color }}>{k.value}</div>
          </Card>
        ))}
        <Card style={{ padding: 16 }}>
          <div className="flex items-center gap-2 mb-1"><Hourglass size={16} style={{ color: "#D97706" }} /><span style={{ fontSize: 12, color: "var(--text-muted)" }}>ยังเปิดอยู่ {open.length} รายการ · ยอดรวม</span></div>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 26, fontWeight: 700, color: "#D97706" }}>{fmtMoney(Math.round(openAmount))} <span style={{ fontSize: 13, color: "var(--text-muted)" }}>บาท</span></div>
          {openNoPrice > 0 && <div style={{ fontSize: 11, color: "#D97706", marginTop: 2 }}>ไม่รวม {openNoPrice} รายการที่ยังไม่ทราบราคา</div>}
        </Card>
      </div>

      <div className="flex flex-wrap items-center gap-2 mb-3">
        <TapChip active={sourceFilter === "all"} onClick={() => setSourceFilter("all")} style={{ minHeight: 34 }}>ทั้งหมด {items.length}</TapChip>
        <TapChip color={SOURCE_META.pr.color} active={sourceFilter === "pr"} onClick={() => setSourceFilter("pr")} style={{ minHeight: 34 }}>สั่งซื้อ {prCount}</TapChip>
        <TapChip color={SOURCE_META.repair.color} active={sourceFilter === "repair"} onClick={() => setSourceFilter("repair")} style={{ minHeight: 34 }}>งานซ่อม {repairCount}</TapChip>
      </div>
      <div className="flex flex-wrap items-center gap-2 mb-4">
        <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} style={{ ...inputStyle, width: "auto", minWidth: 150 }}>
          <option value="all">ทุกสถานะ</option>
          {allStatuses.map((st) => <option key={st} value={st}>{st}</option>)}
        </select>
        <select value={categoryFilter} onChange={(e) => setCategoryFilter(e.target.value)} style={{ ...inputStyle, width: "auto", minWidth: 150 }}>
          <option value="all">ทุกหมวด/ประเภท</option>
          {allCategories.map((c) => <option key={c} value={c}>{c}</option>)}
        </select>
        <div className="flex items-center gap-2 rounded-lg px-3 py-2" style={{ background: "var(--surface-2)", border: "1px solid var(--border)", minWidth: 240, flex: 1, maxWidth: 420 }}>
          <Search size={15} style={{ color: "var(--text-muted)" }} />
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="ค้นหา เลข PR / PO / รายการ / ทะเบียน / ผู้ขาย" style={{ background: "transparent", outline: "none", color: "var(--text)", fontSize: 13, width: "100%" }} />
        </div>
      </div>

      <Card style={{ overflow: "hidden" }}>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 1240 }}>
            <thead>
              <tr style={{ background: "var(--surface-2)" }}>
                {["ประเภท", "วันที่", "เลขที่ PR", "เลขที่ PO", "รายการ", "ทะเบียน", "ผู้ขาย / อู่", "ยอดเงิน", "สถานะ", "จัดการ"].map((h) => (
                  <th key={h} style={{ textAlign: "left", padding: "10px 14px", fontSize: 12, color: "var(--text-muted)", fontWeight: 600 }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map((i) => {
                const meta = SOURCE_META[i.source];
                const late = i.neededBy && !i.closed && i.neededBy < todayISO();
                const statusOptions = i.source === "pr" ? PR_STATUSES.map((s) => s.key) : ["กำลังดำเนินการ", "เสร็จสิ้น"];
                return (
                  <tr
                    key={i.uid}
                    onClick={() => (i.source === "pr" ? setFormState(i.raw) : setRepairEdit(i.raw))}
                    title="คลิกเพื่อแก้ไข"
                    style={{ borderTop: "1px solid var(--border)", cursor: "pointer", transition: "background .12s", opacity: i.status === "ยกเลิก" ? 0.55 : 1 }}
                    onMouseEnter={(e) => (e.currentTarget.style.background = "rgba(14,143,160,0.07)")}
                    onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
                  >
                    <td style={{ padding: "10px 14px" }}>
                      <span className="inline-flex items-center rounded-full px-2 py-1" style={{ background: `${meta.color}1E`, color: meta.color, border: `1px solid ${meta.color}55`, fontSize: 11, fontWeight: 700, whiteSpace: "nowrap" }}>{meta.label}</span>
                    </td>
                    <td style={{ padding: "10px 14px", fontSize: 12, color: "var(--text-muted)", fontFamily: "'JetBrains Mono', monospace", whiteSpace: "nowrap" }}>
                      {fmtDate(i.date)}
                      {i.neededBy && (
                        <div style={{ fontSize: 11, color: late ? "#DC2626" : "var(--text-muted)", fontWeight: late ? 700 : 400 }}>ต้องการ {fmtDate(i.neededBy)}{late ? " (เลยกำหนด)" : ""}</div>
                      )}
                    </td>
                    <td style={{ padding: "10px 14px", fontSize: 12, color: "var(--text)", fontFamily: "'JetBrains Mono', monospace", whiteSpace: "nowrap" }}>{refLabel("PR", i.pr)}</td>
                    <td style={{ padding: "10px 14px", fontSize: 12, color: "var(--text)", fontFamily: "'JetBrains Mono', monospace", whiteSpace: "nowrap" }}>{refLabel("PO", i.po)}</td>
                    <td style={{ padding: "10px 14px", maxWidth: 340 }}>
                      <span className="inline-flex items-center rounded-full px-2 py-0.5" style={{ background: "rgba(14,143,160,0.12)", color: "var(--accent-frost)", fontSize: 11, fontWeight: 600, marginBottom: 3 }}>{i.category}</span>
                      <div style={{ fontSize: 13, color: "var(--text)", whiteSpace: "pre-wrap" }}>{i.description}</div>
                    </td>
                    <td style={{ padding: "10px 14px" }}>{i.plate ? <PlateBadge plate={i.plate} /> : <span style={{ fontSize: 12, color: "var(--text-muted)" }}>ทั่วไป</span>}</td>
                    <td style={{ padding: "10px 14px", fontSize: 12, color: "var(--text-muted)" }}>{i.party || "-"}</td>
                    <td style={{ padding: "10px 14px", fontSize: 13, fontWeight: 600, color: isPendingCost(i.amount) ? "#D97706" : "var(--text)", whiteSpace: "nowrap" }}>
                      {isPendingCost(i.amount) ? "ยังไม่ทราบราคา" : `${fmtMoney(i.amount)} บ.`}
                    </td>
                    <td style={{ padding: "10px 14px" }} onClick={(e) => e.stopPropagation()}>
                      <select
                        value={i.status} onChange={(e) => changeStatus(i, e.target.value)} title="เปลี่ยนสถานะได้ทันที"
                        style={{ background: `${statusColor(i.status)}1E`, color: statusColor(i.status), border: `1px solid ${statusColor(i.status)}66`, borderRadius: 999, padding: "5px 10px", fontSize: 12, fontWeight: 700, cursor: "pointer", outline: "none" }}
                      >
                        {statusOptions.map((st) => <option key={st} value={st} style={{ color: "#1E293B", background: "#FFFFFF" }}>{st}</option>)}
                      </select>
                    </td>
                    <td style={{ padding: "10px 14px" }} onClick={(e) => e.stopPropagation()}>
                      <div className="flex items-center gap-1">
                        <button onClick={() => (i.source === "pr" ? setFormState(i.raw) : setRepairEdit(i.raw))} title="แก้ไข" style={iconBtnStyle}><Pencil size={15} /></button>
                        {i.plate && <button onClick={() => onGoToVehicle(i.plate)} title="ไปที่หน้ารถคันนี้" style={iconBtnStyle}><Truck size={15} /></button>}
                        {i.source === "pr" && <button onClick={() => setDeleteTarget(i.raw)} title="ลบใบขอซื้อ" style={{ ...iconBtnStyle, color: "#DC2626" }}><Trash2 size={15} /></button>}
                      </div>
                    </td>
                  </tr>
                );
              })}
              {filtered.length === 0 && (
                <tr><td colSpan={10} style={{ padding: "28px 14px", textAlign: "center", color: "var(--text-muted)", fontSize: 13 }}>
                  {items.length === 0 ? "ยังไม่มี PR/PO กด \"เปิด PR ใหม่\" เพื่อเริ่มบันทึก หรือใส่เลข PR/PO ในรายการซ่อมแล้วจะมาแสดงที่นี่" : "ไม่พบรายการที่ตรงกับตัวกรอง"}
                </td></tr>
              )}
            </tbody>
          </table>
        </div>
      </Card>

      {formState && (
        <PRFormModal
          initial={formState === "add" ? null : formState}
          vehicles={vehicles}
          onClose={() => setFormState(null)}
          onSave={async (data) => { if (formState === "add") await onAdd(data); else await onUpdate(data); }}
        />
      )}

      {repairEdit && (
        <RepairFormModal
          plate={repairEdit.plate}
          initial={repairEdit}
          onClose={() => setRepairEdit(null)}
          onSave={async (data) => { await onUpdateRepair(data); }}
        />
      )}

      {deleteTarget && (
        <ConfirmDelete
          label={`ต้องการลบใบขอซื้อ ${deleteTarget.pr_number ? refLabel("PR", deleteTarget.pr_number) : ""} "${String(deleteTarget.description).slice(0, 60)}" ใช่หรือไม่?`}
          onCancel={() => setDeleteTarget(null)}
          onConfirm={async () => { try { await onDelete(deleteTarget.id); } catch (e) { setActionError(e.message || "ลบไม่สำเร็จ"); } setDeleteTarget(null); }}
        />
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   DRIVERS VIEW
---------------------------------------------------------------- */

function DriversView({ drivers, vehiclePlates, onAdd, onUpdate, onDelete }) {
  const [formState, setFormState] = useState(null);
  const [deleteTarget, setDeleteTarget] = useState(null);
  const [viewer, setViewer] = useState(null); // { url, title }

  const urlMap = useSignedUrls(
    drivers.flatMap((d) => [d.photo_path, d.license_photo_path, d.company_license_photo_path])
  );

  return (
    <div>
      <div className="flex items-center justify-between flex-wrap gap-3 mb-4">
        <SectionTitle icon={User} sub="ข้อมูลพนักงานขับรถ ใบขับขี่ และรถที่รับผิดชอบ">พนักงานขับรถ</SectionTitle>
        <button onClick={() => setFormState("add")} className="flex items-center gap-2 rounded-lg px-4 py-2" style={{ background: "var(--accent-frost)", color: "#FFFFFF", fontSize: 13, fontWeight: 700, height: 38 }}>
          <Plus size={16} />เพิ่มพนักงาน
        </button>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
        {drivers.map((d) => {
          const expDays = daysBetween(TODAY, new Date(d.expiry));
          const expSoon = expDays <= 60;
          return (
            <Card key={d.id} style={{ padding: 18 }}>
              <div className="flex items-start justify-between mb-4">
                <div className="flex items-center gap-3">
                  {d.photo_path && urlMap[d.photo_path] ? (
                    <img
                      src={urlMap[d.photo_path]} alt={d.name}
                      onClick={() => setViewer({ url: urlMap[d.photo_path], title: `รูปหน้า ${d.name}` })}
                      style={{ width: 52, height: 52, borderRadius: 999, objectFit: "cover", cursor: "pointer", border: "2px solid var(--border)" }}
                    />
                  ) : (
                    <div className="flex items-center justify-center rounded-full" style={{ width: 52, height: 52, background: "rgba(14,143,160,0.15)", color: "var(--accent-frost)" }}><User size={22} /></div>
                  )}
                  <div><div style={{ fontSize: 14, fontWeight: 600, color: "var(--text)" }}>{d.name}</div><div style={{ fontSize: 12, color: "var(--text-muted)" }}>{d.years} ปีประสบการณ์</div></div>
                </div>
                <div className="flex items-center gap-1">
                  <button onClick={() => setFormState(d)} title="แก้ไข" style={iconBtnStyle}><Pencil size={14} /></button>
                  <button onClick={() => setDeleteTarget(d)} title="ลบ" style={{ ...iconBtnStyle, color: "#DC2626" }}><Trash2 size={14} /></button>
                </div>
              </div>
              <div className="flex flex-col gap-2 mb-4">
                <div className="flex items-center gap-2" style={{ fontSize: 13, color: "var(--text-muted)" }}><CreditCard size={14} /><span>ใบขับขี่ {d.license_type} · {d.license}</span></div>
                <div className="flex items-center gap-2" style={{ fontSize: 13, color: "var(--text-muted)" }}><Phone size={14} /><span>{d.phone || "-"}</span></div>
              </div>
              <div className="flex items-center justify-between rounded-lg px-3 py-2 mb-3" style={{ background: "var(--surface-2)" }}>
                <span style={{ fontSize: 12, color: "var(--text-muted)" }}>วันหมดอายุใบขับขี่</span>
                <span style={{ fontSize: 12, fontWeight: 600, color: expSoon ? "#D97706" : "var(--text)", fontFamily: "'JetBrains Mono', monospace" }}>{fmtDate(d.expiry)}</span>
              </div>
              {(d.company_license_no || d.company_license_expiry) && (
                <div className="flex items-center justify-between rounded-lg px-3 py-2 mb-3" style={{ background: "var(--surface-2)" }}>
                  <span style={{ fontSize: 12, color: "var(--text-muted)" }}>ใบขับขี่เครือ {d.company_license_no ? `· ${d.company_license_no}` : ""}</span>
                  {d.company_license_expiry && (
                    <span style={{ fontSize: 12, fontWeight: 600, color: daysBetween(TODAY, new Date(d.company_license_expiry)) <= 60 ? "#D97706" : "var(--text)", fontFamily: "'JetBrains Mono', monospace" }}>{fmtDate(d.company_license_expiry)}</span>
                  )}
                </div>
              )}
              <div className="flex items-center gap-2 flex-wrap mb-3">
                {d.license_photo_path && urlMap[d.license_photo_path] && (
                  <button onClick={() => setViewer({ url: urlMap[d.license_photo_path], title: `ใบขับขี่ ${d.name}` })} className="inline-flex items-center gap-1 rounded-full px-3 py-1" style={{ background: "rgba(14,143,160,0.12)", color: "var(--accent-frost)", fontSize: 12, fontWeight: 600 }}>
                    <Camera size={12} />ดูรูปใบขับขี่
                  </button>
                )}
                {d.company_license_photo_path && urlMap[d.company_license_photo_path] && (
                  <button onClick={() => setViewer({ url: urlMap[d.company_license_photo_path], title: `ใบขับขี่ของเครือ ${d.name}` })} className="inline-flex items-center gap-1 rounded-full px-3 py-1" style={{ background: "rgba(14,143,160,0.12)", color: "var(--accent-frost)", fontSize: 12, fontWeight: 600 }}>
                    <Camera size={12} />ดูใบขับขี่เครือ
                  </button>
                )}
                {!d.license_photo_path && !d.company_license_photo_path && (
                  <span style={{ fontSize: 11, color: "var(--text-muted)" }}>ยังไม่ได้แนบรูปใบขับขี่</span>
                )}
              </div>
              <div className="flex items-center justify-between">
                <span style={{ fontSize: 12, color: "var(--text-muted)" }}>รถที่รับผิดชอบ</span>
                {d.vehicle === "สำรอง" ? <span style={{ fontSize: 12, color: "var(--text-muted)" }}>พนักงานสำรอง</span> : <PlateBadge plate={d.vehicle} />}
              </div>
            </Card>
          );
        })}
        {drivers.length === 0 && <p style={{ fontSize: 13, color: "var(--text-muted)" }}>ยังไม่มีข้อมูลพนักงานขับรถ</p>}
      </div>

      {formState && (
        <DriverFormModal
          initial={formState === "add" ? null : formState}
          vehiclePlates={vehiclePlates}
          urlMap={urlMap}
          onClose={() => setFormState(null)}
          onSave={async (data) => { if (formState === "add") await onAdd(data); else await onUpdate(data.id, data); }}
        />
      )}

      {deleteTarget && (
        <ConfirmDelete
          label={`ต้องการลบข้อมูลพนักงาน "${deleteTarget.name}" ใช่หรือไม่?`}
          onCancel={() => setDeleteTarget(null)}
          onConfirm={async () => { await onDelete(deleteTarget.id); setDeleteTarget(null); }}
        />
      )}

      {viewer && <ImageViewer url={viewer.url} title={viewer.title} onClose={() => setViewer(null)} />}
    </div>
  );
}

/* ---------------------------------------------------------------
   ตกแต่งพื้นหลังเว็บ (ใช้ร่วมกันทุกคน เก็บใน Supabase)
---------------------------------------------------------------- */

const BG_BUCKET = "app-assets";

function hexA(hex, a) {
  const n = parseInt(String(hex).replace("#", ""), 16);
  return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${a})`;
}

async function uploadBgImage(file) {
  const blob = await compressImage(file, 1920, 0.8);
  const rand = Math.random().toString(36).slice(2, 8);
  const path = `background/${Date.now()}-${rand}.jpg`;
  const { error } = await supabase.storage.from(BG_BUCKET).upload(path, blob, { contentType: "image/jpeg", upsert: false, cacheControl: "31536000" });
  if (error) throw error;
  return path;
}

async function removeBgImage(path) {
  if (!path) return;
  try { await supabase.storage.from(BG_BUCKET).remove([path]); } catch (e) { /* ลบไฟล์เก่าไม่สำเร็จไม่เป็นไร */ }
}

function BackgroundModal({ current, currentUrl, theme, onClose, onSaved }) {
  const [file, setFile] = useState(null);
  const [preview, setPreview] = useState(null);
  const [removeImg, setRemoveImg] = useState(false);
  const [dim, setDim] = useState(current && typeof current.dim === "number" ? current.dim : 50);
  const [card, setCard] = useState(current && typeof current.card === "number" ? current.card : 92);
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    if (!file) { setPreview(null); return undefined; }
    const u = URL.createObjectURL(file);
    setPreview(u);
    return () => URL.revokeObjectURL(u);
  }, [file]);

  const shownUrl = preview || (!removeImg ? currentUrl : "");

  async function save() {
    setSaving(true);
    setError("");
    let uploaded = null;
    try {
      let path = removeImg ? null : (current && current.path) || null;
      if (file) { uploaded = await uploadBgImage(file); path = uploaded; }
      const { error: e } = await supabase.from("app_settings").upsert(
        { key: "background", value: { path, dim, card }, updated_at: new Date().toISOString() },
        { onConflict: "key" }
      );
      if (e) throw e;
      const old = current && current.path;
      if (old && old !== path) await removeBgImage(old);
      await onSaved();
      onClose();
    } catch (err) {
      if (uploaded) await removeBgImage(uploaded);
      const msg = String((err && err.message) || "");
      setError(/app_settings|schema cache|bucket|relation/i.test(msg)
        ? "ฐานข้อมูลยังไม่พร้อม กรุณารัน migration_v11.sql ใน Supabase ก่อน"
        : (msg || "บันทึกไม่สำเร็จ กรุณาลองใหม่อีกครั้ง"));
    } finally {
      setSaving(false);
    }
  }

  return (
    <ModalShell title="ตกแต่งพื้นหลังเว็บ" onClose={onClose} width={600}>
      <div className="flex flex-col gap-4">
        {/* ตัวอย่าง */}
        <div style={{ position: "relative", height: 170, borderRadius: 12, overflow: "hidden", border: "1px solid var(--border)", background: "var(--surface-2)" }}>
          {shownUrl ? (
            <>
              <div style={{ position: "absolute", inset: 0, backgroundImage: `url(${shownUrl})`, backgroundSize: "cover", backgroundPosition: "center" }} />
              <div style={{ position: "absolute", inset: 0, background: theme.bg, opacity: dim / 100 }} />
              <div style={{ position: "absolute", left: 22, right: 22, top: 26, bottom: 22, borderRadius: 12, background: hexA(theme.surface, card / 100), border: `1px solid ${theme.border}`, padding: 12 }}>
                <div style={{ fontSize: 12, fontWeight: 700, color: theme.text }}>ตัวอย่างการ์ดข้อมูล</div>
                <div style={{ fontSize: 11, color: theme.textMuted, marginTop: 4 }}>พื้นหลังที่ปรับแล้วจะเป็นแบบนี้ ลองเลื่อนแถบด้านล่างดู</div>
              </div>
            </>
          ) : (
            <div className="flex items-center justify-center" style={{ height: "100%", fontSize: 13, color: "var(--text-muted)" }}>ยังไม่มีรูปพื้นหลัง (ใช้สีพื้นเรียบตามเดิม)</div>
          )}
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <label className="rounded-lg px-4 py-2" style={{ background: "var(--accent-frost)", color: "#FFFFFF", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
            {shownUrl ? "เปลี่ยนรูป" : "เลือกรูปพื้นหลัง"}
            <input type="file" accept="image/*" style={{ display: "none" }} onChange={(e) => { const f = e.target.files && e.target.files[0]; if (f) { setFile(f); setRemoveImg(false); } e.target.value = ""; }} />
          </label>
          {shownUrl && (
            <button type="button" onClick={() => { setFile(null); setRemoveImg(true); }} className="rounded-lg px-4 py-2" style={{ background: "transparent", border: "1px solid #DC262655", color: "#DC2626", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
              ลบรูปพื้นหลัง
            </button>
          )}
        </div>

        <div className="flex flex-col gap-1">
          <div className="flex items-center justify-between"><span style={{ fontSize: 12, color: "var(--text-muted)", fontWeight: 600 }}>ความเข้มของฉากหลัง (ยิ่งมากยิ่งอ่านตัวหนังสือง่าย)</span><b style={{ fontSize: 12, color: "var(--text)" }}>{dim}%</b></div>
          <input type="range" min="0" max="90" step="5" value={dim} onChange={(e) => setDim(Number(e.target.value))} style={{ width: "100%", accentColor: "#0E8FA0" }} />
        </div>
        <div className="flex flex-col gap-1">
          <div className="flex items-center justify-between"><span style={{ fontSize: 12, color: "var(--text-muted)", fontWeight: 600 }}>ความทึบของการ์ด (น้อย = เห็นรูปทะลุการ์ดมากขึ้น)</span><b style={{ fontSize: 12, color: "var(--text)" }}>{card}%</b></div>
          <input type="range" min="60" max="100" step="2" value={card} onChange={(e) => setCard(Number(e.target.value))} style={{ width: "100%", accentColor: "#0E8FA0" }} />
        </div>

        <div style={{ fontSize: 11, color: "var(--text-muted)" }}>รูปจะถูกย่อขนาดอัตโนมัติ และ<b>มีผลกับผู้ใช้ทุกคน</b>ที่เปิดเว็บนี้ เหมาะกับรูปแนวนอน เช่น ภาพรถบรรทุก ลานจอด หรือโลโก้บริษัททำเป็นลวดลาย</div>

        {error && <div style={{ fontSize: 12, color: "#DC2626", background: "rgba(220,38,38,0.1)", border: "1px solid #DC262655", borderRadius: 8, padding: "8px 10px" }}>{error}</div>}

        <div className="flex items-center justify-end gap-2" style={{ position: "sticky", bottom: 0, background: "var(--surface-solid, var(--surface))", paddingTop: 12, borderTop: "1px solid var(--border)" }}>
          <button type="button" onClick={onClose} style={{ ...inputStyle, width: "auto", padding: "10px 18px", cursor: "pointer" }}>ยกเลิก</button>
          <button type="button" onClick={save} disabled={saving} style={{ background: "var(--accent-frost)", color: "#FFFFFF", border: "none", borderRadius: 8, padding: "10px 26px", fontSize: 14, fontWeight: 700, cursor: "pointer", opacity: saving ? 0.6 : 1 }}>
            {saving ? "กำลังบันทึก..." : "บันทึก"}
          </button>
        </div>
      </div>
    </ModalShell>
  );
}

/* ---------------------------------------------------------------
   APP SHELL
---------------------------------------------------------------- */

const TABS = [
  { key: "dashboard", label: "แดชบอร์ด", icon: Gauge },
  { key: "readiness", label: "ความพร้อมใช้งาน", icon: Box },
  { key: "vehicles", label: "รถทั้งหมด", icon: Truck },
  { key: "maintenance", label: "บำรุงรักษา", icon: Settings2 },
  { key: "repairlog", label: "ประวัติการซ่อม", icon: ClipboardList },
  { key: "pr", label: "PR / PO", icon: ShoppingCart },
  { key: "fuel", label: "เติมน้ำมัน", icon: Fuel },
  { key: "drivers", label: "พนักงานขับรถ", icon: User },
];

export default function FleetApp({ user }) {
  const [tab, setTab] = useState("dashboard");
  const [dark, setDark] = useState(false);

  useEffect(() => {
    try {
      const saved = window.localStorage.getItem("cpram_theme");
      if (saved === "dark") setDark(true);
    } catch (e) {}
  }, []);

  function toggleTheme() {
    setDark((d) => {
      const next = !d;
      try { window.localStorage.setItem("cpram_theme", next ? "dark" : "light"); } catch (e) {}
      return next;
    });
  }

  const theme = dark ? THEMES.dark : THEMES.light;
  const bgUrl = bgSetting && bgSetting.path ? supabase.storage.from(BG_BUCKET).getPublicUrl(bgSetting.path).data.publicUrl : "";
  const bgDim = bgSetting && typeof bgSetting.dim === "number" ? bgSetting.dim : 50;
  const bgCard = bgSetting && typeof bgSetting.card === "number" ? bgSetting.card : 92;
  const [focusPlate, setFocusPlate] = useState(null);

  function handleTabClick(key) {
    setFocusPlate(null);
    setTab(key);
  }
  function goToVehicle(plate) {
    setFocusPlate(plate);
    setTab("vehicles");
  }
  const [vehicles, setVehicles] = useState([]);
  const [repairs, setRepairs] = useState([]);
  const [drivers, setDrivers] = useState([]);
  const [fuelLogs, setFuelLogs] = useState([]);
  const [prs, setPrs] = useState([]);
  const [bgSetting, setBgSetting] = useState(null);
  const [showBg, setShowBg] = useState(false);
  const [prError, setPrError] = useState("");
  const [loading, setLoading] = useState(true);
  const [errorMsg, setErrorMsg] = useState("");

  useEffect(() => { loadAll(true); }, []);
  useEffect(() => { loadBg(); }, []);

  // อ่านค่าพื้นหลัง (ถ้ายังไม่ได้รัน SQL หรือยังไม่เคยตั้งค่า ก็แค่ไม่แสดงรูป ไม่กระทบส่วนอื่น)
  async function loadBg() {
    try {
      const { data, error } = await supabase.from("app_settings").select("value").eq("key", "background").maybeSingle();
      if (!error && data && data.value) setBgSetting(data.value);
      else if (!error) setBgSetting(null);
    } catch (e) { /* ignore */ }
  }

  // initial=true แสดงหน้าโหลดเต็มจอเฉพาะครั้งแรก ส่วนหลังบันทึก/แก้ไข รีเฟรชข้อมูลเงียบๆ
  async function loadAll(initial = false) {
    if (initial) setLoading(true);
    try {
      let prFail = "";
      const [v, r, d, fl, pr] = await Promise.all([
        fetchVehicles(), fetchRepairs(), fetchDrivers(), fetchFuelLogs(),
        fetchPurchaseRequests().catch((e) => { prFail = e.message || "ไม่พบตาราง purchase_requests"; return []; }),
      ]);
      setVehicles(v);
      setRepairs(r);
      setDrivers(d);
      setFuelLogs(fl);
      setPrs(pr);
      setPrError(prFail);
      setErrorMsg("");
    } catch (e) {
      setErrorMsg("โหลดข้อมูลไม่สำเร็จ: " + (e.message || "unknown error"));
    } finally {
      setLoading(false);
    }
  }

  const computedVehicles = useMemo(() => vehicles.map(computeVehicle), [vehicles]);

  async function handleAddVehicle(v) {
    const { oldId, ...clean } = v;
    const { error } = await supabase.from("vehicles").insert(clean);
    if (error) throw error;
    await loadAll();
  }
  async function handleUpdateVehicle(oldId, updates) {
    const { oldId: _drop, ...clean } = updates;
    const { error } = await supabase.from("vehicles").update(clean).eq("id", oldId);
    if (error) throw error;
    await loadAll();
  }
  async function handleDeleteVehicle(plate) {
    const { error } = await supabase.from("vehicles").delete().eq("id", plate);
    if (error) throw error;
    await supabase.from("drivers").update({ vehicle: "สำรอง" }).eq("vehicle", plate);
    await loadAll();
  }

  async function handleAddRepair(r) {
    const { id, ...clean } = r;
    const { error } = await supabase.from("repairs").insert(clean);
    if (error) throw error;
    await loadAll();
  }
  async function handleUpdateRepair(r) {
    const { id, ...clean } = r;
    const { error } = await supabase.from("repairs").update(clean).eq("id", id);
    if (error) throw error;
    await loadAll();
  }
  async function handleDeleteRepair(id) {
    const { error } = await supabase.from("repairs").delete().eq("id", id);
    if (error) throw error;
    await loadAll();
  }

  async function handleAddDriver(d) {
    const { id, ...clean } = d;
    const { error } = await supabase.from("drivers").insert(clean);
    if (error) throw error;
    await loadAll();
  }
  async function handleUpdateDriver(id, updates) {
    const { id: _drop, ...clean } = updates;
    const { error } = await supabase.from("drivers").update(clean).eq("id", id);
    if (error) throw error;
    await loadAll();
  }
  async function handleDeleteDriver(id) {
    const target = drivers.find((d) => d.id === id);
    const { error } = await supabase.from("drivers").delete().eq("id", id);
    if (error) throw error;
    if (target) await removeDriverImages([target.photo_path, target.license_photo_path, target.company_license_photo_path]);
    await loadAll();
  }

  async function handleAddPR(data) {
    const { id, ...clean } = data;
    const { error } = await supabase.from("purchase_requests").insert(clean);
    if (error) throw error;
    await loadAll();
  }
  async function handleUpdatePR(data) {
    const { id, created_at, ...clean } = data;
    const { error } = await supabase.from("purchase_requests").update(clean).eq("id", id);
    if (error) throw error;
    await loadAll();
  }
  async function handleDeletePR(id) {
    const { error } = await supabase.from("purchase_requests").delete().eq("id", id);
    if (error) throw error;
    await loadAll();
  }

  // นำเข้าหลายรายการจากไฟล์ Excel (ข้ามรายการที่เลขใบกำกับซ้ำ) คืนจำนวนที่เพิ่มจริง
  async function handleImportFuel(rows) {
    let inserted = 0;
    for (let i = 0; i < rows.length; i += 100) {
      const chunk = rows.slice(i, i + 100);
      const { data, error } = await supabase.from("fuel_logs").upsert(chunk, { onConflict: "ref_no", ignoreDuplicates: true }).select("id");
      if (error) throw error;
      inserted += data ? data.length : 0;
    }
    await loadAll();
    return inserted;
  }

  async function handleAddFuel(f) {
    const { id, ...clean } = f;
    const { error } = await supabase.from("fuel_logs").insert(clean);
    if (error) throw error;
    await loadAll();
  }
  async function handleUpdateFuel(id, updates) {
    const { id: _drop, ...clean } = updates;
    const { error } = await supabase.from("fuel_logs").update(clean).eq("id", id);
    if (error) throw error;
    await loadAll();
  }
  async function handleDeleteFuel(id) {
    const { error } = await supabase.from("fuel_logs").delete().eq("id", id);
    if (error) throw error;
    await loadAll();
  }

  async function handleSignOut() {
    await supabase.auth.signOut();
    window.location.href = "/login";
  }

  return (
    <div style={{
      "--bg": theme.bg, "--surface": bgUrl ? hexA(theme.surface, bgCard / 100) : theme.surface, "--surface-solid": theme.surface,
      "--surface-2": theme.surface2, "--border": theme.border, "--card-blur": bgUrl ? "blur(8px)" : "none",
      "--text": theme.text, "--text-muted": theme.textMuted, "--accent-frost": theme.accent,
      background: "var(--bg)", minHeight: "100vh", fontFamily: "'Inter', sans-serif", transition: "background .2s ease",
      position: "relative", isolation: "isolate",
    }}>
      {bgUrl && (
        <>
          <div aria-hidden="true" style={{ position: "fixed", inset: 0, zIndex: -1, backgroundImage: `url(${bgUrl})`, backgroundSize: "cover", backgroundPosition: "center", pointerEvents: "none" }} />
          <div aria-hidden="true" style={{ position: "fixed", inset: 0, zIndex: -1, background: theme.bg, opacity: bgDim / 100, pointerEvents: "none" }} />
        </>
      )}
      <header style={{ borderBottom: "1px solid var(--border)", background: theme.headerBg, position: "sticky", top: 0, zIndex: 10, backdropFilter: "blur(6px)" }}>
        <div className="mx-auto px-5 py-4 flex items-center justify-between flex-wrap gap-3" style={{ maxWidth: 1720 }}>
          <div className="flex items-center gap-3">
            <div className="flex items-center justify-center rounded-lg" style={{ width: 42, height: 42, background: "#FFFFFF", boxShadow: "0 2px 8px rgba(15,23,42,0.12)" }}>
              <span style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 800, fontSize: 15, fontStyle: "italic", letterSpacing: -0.5 }}>
                <span style={{ color: "#E4432E" }}>cp</span><span style={{ color: "#4C9A3D" }}>ram</span>
              </span>
            </div>
            <div>
              <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 17, color: "var(--text)", letterSpacing: 0.3 }}>CPRAM TRANSPORT</div>
              <div style={{ fontSize: 11, color: "var(--text-muted)" }}>ระบบบริหารจัดการรถขนส่งควบคุมอุณหภูมิ</div>
            </div>
          </div>
          <nav className="flex items-center gap-1 rounded-lg p-1 flex-wrap" style={{ background: "var(--surface-2)" }}>
            {TABS.map((t) => {
              const active = tab === t.key;
              return (
                <button key={t.key} onClick={() => handleTabClick(t.key)} className="flex items-center gap-2 rounded-md px-3 py-2 text-sm font-medium" style={{ color: active ? "#FFFFFF" : "var(--text-muted)", background: active ? "var(--accent-frost)" : "transparent", transition: "all .15s" }}>
                  <t.icon size={15} /><span className="hidden sm:inline">{t.label}</span>
                </button>
              );
            })}
          </nav>
          <div className="flex items-center gap-3">
            <span style={{ fontSize: 12, color: "var(--text-muted)" }} className="hidden md:inline">{user?.email}</span>
            <button onClick={() => setShowBg(true)} title="ตกแต่งพื้นหลังเว็บ" className="flex items-center justify-center rounded-lg" style={{ width: 36, height: 36, background: "var(--surface-2)", border: "1px solid var(--border)", color: "var(--text-muted)" }}>
              <Palette size={16} />
            </button>
            <button onClick={toggleTheme} title={dark ? "สลับเป็นโหมดสว่าง" : "สลับเป็นโหมดมืด"} className="flex items-center justify-center rounded-lg" style={{ width: 36, height: 36, background: "var(--surface-2)", border: "1px solid var(--border)", color: "var(--text-muted)" }}>
              {dark ? <Sun size={16} /> : <Moon size={16} />}
            </button>
            <button onClick={handleSignOut} className="flex items-center gap-1 rounded-lg px-3 py-2" style={{ background: "var(--surface-2)", border: "1px solid var(--border)", color: "var(--text-muted)", fontSize: 12 }}>
              <LogOut size={14} />ออกจากระบบ
            </button>
          </div>
        </div>
      </header>

      <main className="mx-auto px-5 py-6" style={{ maxWidth: 1720 }}>
        {errorMsg && (
          <div style={{ marginBottom: 16, fontSize: 12, color: "#D97706", background: "rgba(217,119,6,0.1)", border: "1px solid #D9770655", borderRadius: 8, padding: "8px 12px" }}>{errorMsg}</div>
        )}
        {loading ? (
          <div style={{ color: "var(--text-muted)", fontSize: 14, padding: "40px 0", textAlign: "center" }}>กำลังโหลดข้อมูล...</div>
        ) : (
          <>
            {tab === "dashboard" && <Dashboard vehicles={computedVehicles} repairs={repairs} onGoToVehicle={goToVehicle} dark={dark} />}
            {tab === "readiness" && <ReadinessView vehicles={computedVehicles} onGoToVehicle={goToVehicle} />}
            {tab === "vehicles" && (
              <VehiclesView
                vehicles={computedVehicles} repairs={repairs}
                onAdd={handleAddVehicle} onUpdate={handleUpdateVehicle} onDelete={handleDeleteVehicle}
                onAddRepair={handleAddRepair} onUpdateRepair={handleUpdateRepair} onDeleteRepair={handleDeleteRepair}
                initialPlate={focusPlate}
              />
            )}
            {tab === "maintenance" && <MaintenanceView vehicles={computedVehicles} />}
            {tab === "repairlog" && <RepairsLogView vehicles={computedVehicles} repairs={repairs} onUpdateRepair={handleUpdateRepair} onDeleteRepair={handleDeleteRepair} onGoToVehicle={goToVehicle} />}
            {tab === "pr" && <PurchaseRequestsView vehicles={vehicles} prs={prs} repairs={repairs} tableError={prError} onAdd={handleAddPR} onUpdate={handleUpdatePR} onDelete={handleDeletePR} onUpdateRepair={handleUpdateRepair} onGoToVehicle={goToVehicle} />}
            {tab === "fuel" && <FuelView vehicles={vehicles} fuelLogs={fuelLogs} onAdd={handleAddFuel} onUpdate={handleUpdateFuel} onDelete={handleDeleteFuel} onImport={handleImportFuel} dark={dark} />}
            {tab === "drivers" && (
              <DriversView drivers={drivers} vehiclePlates={vehicles.map((v) => v.id)} onAdd={handleAddDriver} onUpdate={handleUpdateDriver} onDelete={handleDeleteDriver} />
            )}
          </>
        )}
      </main>

      <footer className="mx-auto px-5 py-6" style={{ maxWidth: 1720, color: "var(--text-muted)", fontSize: 12 }}>
        ข้อมูลบันทึกลงฐานข้อมูลจริง ใช้งานพร้อมกันได้หลายคน · CPRAM Fleet Management
      </footer>

      {showBg && (
        <BackgroundModal current={bgSetting} currentUrl={bgUrl} theme={theme} onClose={() => setShowBg(false)} onSaved={loadBg} />
      )}
    </div>
  );
}
