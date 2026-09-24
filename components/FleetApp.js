import { useEffect, useMemo, useState } from "react";
import {
  BarChart, Bar, XAxis, YAxis, ResponsiveContainer, Tooltip, CartesianGrid, Cell,
} from "recharts";
import {
  Truck, Snowflake, Droplet, Sun, Wrench, Calendar, AlertTriangle, CheckCircle2,
  User, Phone, Search, X, Clock3, CreditCard, ChevronRight, ClipboardList, Gauge,
  Plus, Trash2, Pencil, LogOut, Gauge as GaugeIcon, Settings2, FileText, Lock, Hourglass, Fuel, Package, Box, Moon, Download,
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
const DRIVER_LICENSE_TYPES = ["ท.2", "ท.3", "ท.4"];
const THAI_MONTHS = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];

/* ---------------------------------------------------------------
   HELPERS
---------------------------------------------------------------- */

function daysBetween(a, b) { return Math.round((b - a) / (1000 * 60 * 60 * 24)); }
function addDays(dateStr, n) { const d = new Date(dateStr); d.setDate(d.getDate() + n); return d; }
function fmtDate(d) { if (!d) return "-"; const date = typeof d === "string" ? new Date(d) : d; return `${date.getDate()} ${THAI_MONTHS[date.getMonth()]} ${date.getFullYear() + 543}`; }
function fmtMoney(n) { return Number(n || 0).toLocaleString("th-TH"); }

// ส่งออกข้อมูลเป็นไฟล์ Excel (.xlsx) - โหลดไลบรารีเฉพาะตอนกดใช้งานจริง
async function exportToExcel(rows, filename, sheetName) {
  const XLSX = await import("xlsx");
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, filename);
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
  return <div style={{ background: "var(--surface)", border: "1px solid var(--border)", borderRadius: 14, boxShadow: "0 1px 2px rgba(15,23,42,0.04), 0 2px 10px rgba(15,23,42,0.04)", ...style }}>{children}</div>;
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

const inputStyle = { background: "var(--surface-2)", border: "1px solid var(--border)", borderRadius: 8, padding: "8px 10px", color: "var(--text)", fontSize: 13, outline: "none", width: "100%" };
const iconBtnStyle = { color: "var(--text-muted)", padding: 4, borderRadius: 6, cursor: "pointer" };

function ModalShell({ title, onClose, children, width = 560 }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(8,10,14,0.7)", zIndex: 50, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ maxWidth: width, width: "100%", maxHeight: "90vh", overflowY: "auto" }}>
        <Card style={{ padding: 22 }}>
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
        <Card style={{ padding: 20 }}>
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

function RepairFormModal({ plate, initial, onClose, onSave }) {
  const isEdit = !!initial;
  const [form, setForm] = useState(() => initial ? {
    ...initial,
    cost: initial.cost === null || initial.cost === undefined ? "" : String(initial.cost),
    pendingEstimate: initial.cost === null || initial.cost === undefined,
    pr_number: initial.pr_number || "",
    po_number: initial.po_number || "",
  } : {
    date: todayISO(), type: REPAIR_TYPES[0], description: "", cost: "", garage: "", status: "เสร็จสิ้น", pendingEstimate: false, pr_number: "", po_number: "",
  });
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);
  const update = (key, val) => setForm((f) => ({ ...f, [key]: val }));

  async function handleSubmit(e) {
    e.preventDefault();
    if (!form.date) return setError("กรุณาเลือกวันที่ซ่อม");
    if (!form.description.trim()) return setError("กรุณากรอกรายละเอียดงานซ่อม");
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
    <ModalShell title={isEdit ? "แก้ไขรายการซ่อม" : "เพิ่มรายการซ่อม"} onClose={onClose}>
      <div style={{ marginBottom: 12 }}><PlateBadge plate={plate} /></div>
      <form onSubmit={handleSubmit} className="flex flex-col gap-3">
        <div className="grid grid-cols-2 gap-3">
          <Field label="วันที่ซ่อม *"><input style={inputStyle} type="date" value={form.date} onChange={(e) => update("date", e.target.value)} /></Field>
          <Field label="ประเภทงานซ่อม">
            <select style={inputStyle} value={form.type} onChange={(e) => update("type", e.target.value)}>
              {REPAIR_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
            </select>
          </Field>
        </div>
        <Field label="รายละเอียดงานซ่อม (ทำอะไรไปบ้าง) *"><input style={inputStyle} placeholder="เช่น เปลี่ยนผ้าเบรกหน้า" value={form.description} onChange={(e) => update("description", e.target.value)} /></Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="ค่าใช้จ่าย (บาท)">
            <input
              style={{ ...inputStyle, opacity: form.pendingEstimate ? 0.45 : 1 }}
              type="number"
              value={form.pendingEstimate ? "" : form.cost}
              disabled={form.pendingEstimate}
              placeholder={form.pendingEstimate ? "รอประเมินราคา" : ""}
              onChange={(e) => update("cost", e.target.value)}
            />
          </Field>
          <Field label="สถานะงาน">
            <select style={inputStyle} value={form.status} onChange={(e) => update("status", e.target.value)}>
              <option value="เสร็จสิ้น">เสร็จสิ้น</option>
              <option value="กำลังดำเนินการ">กำลังดำเนินการ</option>
            </select>
          </Field>
        </div>
        <label className="flex items-center gap-2 rounded-lg px-3 py-2.5" style={{ background: form.pendingEstimate ? "rgba(217,119,6,0.1)" : "var(--surface-2)", border: `1px solid ${form.pendingEstimate ? "#D9770655" : "var(--border)"}`, cursor: "pointer" }}>
          <input type="checkbox" checked={form.pendingEstimate} onChange={(e) => update("pendingEstimate", e.target.checked)} style={{ width: 15, height: 15, accentColor: "#D97706", cursor: "pointer" }} />
          <Hourglass size={14} style={{ color: form.pendingEstimate ? "#D97706" : "var(--text-muted)" }} />
          <span style={{ fontSize: 13, color: form.pendingEstimate ? "#D97706" : "var(--text-muted)", fontWeight: 600 }}>ยังไม่ทราบค่าใช้จ่าย (รอประเมินราคา)</span>
        </label>
        <div className="grid grid-cols-2 gap-3">
          <Field label="อู่ / ศูนย์บริการ (ทำที่ไหน)"><input style={inputStyle} placeholder="เช่น อู่กลาง CPRAM" value={form.garage} onChange={(e) => update("garage", e.target.value)} /></Field>
          <Field label="เลขที่ PR (ถ้ามี)"><input style={inputStyle} placeholder="เช่น PR-2569-00123" value={form.pr_number} onChange={(e) => update("pr_number", e.target.value)} /></Field>
        </div>
        <Field label="เลขที่ PO (ถ้ามี)"><input style={inputStyle} placeholder="เช่น PO-2569-00456" value={form.po_number} onChange={(e) => update("po_number", e.target.value)} /></Field>
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
   DRIVER FORM (add + edit)
---------------------------------------------------------------- */

function DriverFormModal({ initial, onClose, onSave, vehiclePlates }) {
  const isEdit = !!initial;
  const [form, setForm] = useState(() => initial ? { ...initial, years: String(initial.years) } : {
    name: "", license: "", license_type: "ท.2", expiry: "2027-01-01", phone: "", vehicle: "สำรอง", years: "1",
  });
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);
  const update = (key, val) => setForm((f) => ({ ...f, [key]: val }));

  async function handleSubmit(e) {
    e.preventDefault();
    if (!form.name.trim()) return setError("กรุณากรอกชื่อ-นามสกุล");
    if (!form.license.trim()) return setError("กรุณากรอกเลขที่ใบขับขี่");
    if (!form.expiry) return setError("กรุณาเลือกวันหมดอายุใบขับขี่");
    setSaving(true);
    setError("");
    try {
      await onSave({
        id: isEdit ? initial.id : undefined,
        name: form.name.trim(), license: form.license.trim(), license_type: form.license_type,
        expiry: form.expiry, phone: form.phone.trim(), vehicle: form.vehicle, years: Number(form.years) || 0,
      });
      onClose();
    } catch (err) {
      setError(err.message || "บันทึกไม่สำเร็จ กรุณาลองใหม่อีกครั้ง");
    } finally {
      setSaving(false);
    }
  }

  return (
    <ModalShell title={isEdit ? "แก้ไขข้อมูลพนักงานขับรถ" : "เพิ่มพนักงานขับรถ"} onClose={onClose}>
      <form onSubmit={handleSubmit} className="flex flex-col gap-3">
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
                    "เลขที่ PR": r.pr_number || "", "เลขที่ PO": r.po_number || "",
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
                        <div><div style={{ fontSize: 13, color: "var(--text)" }}>{r.description}</div><div style={{ fontSize: 11, color: "var(--text-muted)" }}>{r.garage} · {r.status}{r.pr_number ? ` · PR ${r.pr_number}` : ""}{r.po_number ? ` · PO ${r.po_number}` : ""}</div></div>
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

function RepairsLogView({ vehicles, repairs }) {
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
      list = list.filter((r) => r.description.toLowerCase().includes(q) || r.type.toLowerCase().includes(q) || r.plate.toLowerCase().includes(q) || (r.garage || "").toLowerCase().includes(q) || (r.pr_number || "").toLowerCase().includes(q) || (r.po_number || "").toLowerCase().includes(q));
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
        "เลขที่ PR": r.pr_number || "",
        "เลขที่ PO": r.po_number || "",
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
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 920 }}>
            <thead>
              <tr style={{ background: "var(--surface-2)" }}>
                {["ทะเบียน", "วันที่", "ประเภท", "รายละเอียด", "เลขที่ PR", "เลขที่ PO", "อู่/ศูนย์บริการ", "สถานะ", "ค่าใช้จ่าย"].map((h) => (
                  <th key={h} style={{ textAlign: "left", padding: "10px 14px", fontSize: 12, color: "var(--text-muted)", fontWeight: 600 }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map((r) => (
                <tr key={r.id} style={{ borderTop: "1px solid var(--border)" }}>
                  <td style={{ padding: "10px 14px" }}><PlateBadge plate={r.plate} /></td>
                  <td style={{ padding: "10px 14px", fontSize: 12, color: "var(--text-muted)", fontFamily: "'JetBrains Mono', monospace" }}>{fmtDate(r.date)}</td>
                  <td style={{ padding: "10px 14px" }}>
                    <span className="rounded-full px-2 py-1 text-xs font-medium" style={{ background: "rgba(14,143,160,0.12)", color: "var(--accent-frost)", whiteSpace: "nowrap" }}>{r.type}</span>
                  </td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: "var(--text)" }}>{r.description}</td>
                  <td style={{ padding: "10px 14px", fontSize: 12, color: "var(--text-muted)", fontFamily: "'JetBrains Mono', monospace" }}>{r.pr_number || "-"}</td>
                  <td style={{ padding: "10px 14px", fontSize: 12, color: "var(--text-muted)", fontFamily: "'JetBrains Mono', monospace" }}>{r.po_number || "-"}</td>
                  <td style={{ padding: "10px 14px", fontSize: 12, color: "var(--text-muted)" }}>{r.garage}</td>
                  <td style={{ padding: "10px 14px", fontSize: 12, color: r.status === "เสร็จสิ้น" ? "#16A34A" : "#D97706" }}>{r.status}</td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: isPendingCost(r.cost) ? "#D97706" : "var(--text)", fontWeight: 600 }}>{fmtCost(r.cost)}</td>
                </tr>
              ))}
              {filtered.length === 0 && (
                <tr><td colSpan={9} style={{ padding: "24px 14px", textAlign: "center", color: "var(--text-muted)", fontSize: 13 }}>ไม่พบรายการซ่อมที่ตรงกับตัวกรอง</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}


/* ---------------------------------------------------------------
   FUEL FORM (add + edit)
---------------------------------------------------------------- */

function FuelFormModal({ initial, vehicles, onClose, onSave }) {
  const isEdit = !!initial;
  const [form, setForm] = useState(() => initial ? {
    ...initial,
    liters: String(initial.liters ?? ""),
    price_per_liter: String(initial.price_per_liter ?? ""),
    odometer: initial.odometer === null || initial.odometer === undefined ? "" : String(initial.odometer),
    station: initial.station || "",
    driver: initial.driver || "",
    note: initial.note || "",
  } : {
    plate: vehicles[0]?.id || "", date: todayISO(), liters: "", price_per_liter: "",
    odometer: "", station: "", driver: "", note: "",
  });
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);
  const update = (key, val) => setForm((f) => ({ ...f, [key]: val }));

  const litersNum = Number(form.liters) || 0;
  const priceNum = Number(form.price_per_liter) || 0;
  const totalCost = Math.round(litersNum * priceNum * 100) / 100;

  async function handleSubmit(e) {
    e.preventDefault();
    if (!form.plate) return setError("กรุณาเลือกทะเบียนรถ");
    if (!form.date) return setError("กรุณาเลือกวันที่เติม");
    if (litersNum <= 0) return setError("กรุณากรอกจำนวนลิตรให้ถูกต้อง");
    if (priceNum <= 0) return setError("กรุณากรอกราคาต่อลิตรให้ถูกต้อง");
    setSaving(true);
    setError("");
    try {
      await onSave({
        id: isEdit ? initial.id : undefined,
        plate: form.plate, date: form.date,
        liters: litersNum, price_per_liter: priceNum, total_cost: totalCost,
        odometer: form.odometer === "" ? null : Number(form.odometer),
        station: form.station.trim(), driver: form.driver.trim(), note: form.note.trim(),
      });
      onClose();
    } catch (err) {
      setError(err.message || "บันทึกไม่สำเร็จ กรุณาลองใหม่อีกครั้ง");
    } finally {
      setSaving(false);
    }
  }

  return (
    <ModalShell title={isEdit ? "แก้ไขรายการเติมน้ำมัน" : "บันทึกการเติมน้ำมัน"} onClose={onClose}>
      <form onSubmit={handleSubmit} className="flex flex-col gap-3">
        <div className="grid grid-cols-2 gap-3">
          <Field label="ทะเบียนรถ *">
            <select style={inputStyle} value={form.plate} onChange={(e) => update("plate", e.target.value)}>
              {vehicles.map((v) => <option key={v.id} value={v.id}>{v.id} · {v.brand} {v.model}</option>)}
            </select>
          </Field>
          <Field label="วันที่เติม *"><input style={inputStyle} type="date" value={form.date} onChange={(e) => update("date", e.target.value)} /></Field>
        </div>

        <div className="grid grid-cols-2 gap-3">
          <Field label="จำนวนลิตร *"><input style={inputStyle} type="number" step="0.01" placeholder="เช่น 120.50" value={form.liters} onChange={(e) => update("liters", e.target.value)} /></Field>
          <Field label="ราคาต่อลิตร (บาท) *"><input style={inputStyle} type="number" step="0.01" placeholder="เช่น 31.94" value={form.price_per_liter} onChange={(e) => update("price_per_liter", e.target.value)} /></Field>
        </div>

        <div className="flex items-center justify-between rounded-lg px-4 py-3" style={{ background: "rgba(14,143,160,0.08)", border: "1px solid #0E8FA055" }}>
          <span style={{ fontSize: 13, color: "var(--text-muted)" }}>ยอดเงินรวม (คำนวณอัตโนมัติ)</span>
          <span style={{ fontSize: 17, fontWeight: 700, color: "var(--accent-frost)" }}>{fmtMoney(totalCost)} บาท</span>
        </div>

        <div className="grid grid-cols-2 gap-3">
          <Field label="เลขไมล์ขณะเติม (ถ้ามี)"><input style={inputStyle} type="number" placeholder="เช่น 182450" value={form.odometer} onChange={(e) => update("odometer", e.target.value)} /></Field>
          <Field label="ปั๊ม / สถานที่เติม"><input style={inputStyle} placeholder="เช่น ปตท. สาขาลาดกระบัง" value={form.station} onChange={(e) => update("station", e.target.value)} /></Field>
        </div>

        <div className="grid grid-cols-2 gap-3">
          <Field label="คนขับที่เติม"><input style={inputStyle} placeholder="ชื่อ-นามสกุล" value={form.driver} onChange={(e) => update("driver", e.target.value)} /></Field>
          <Field label="หมายเหตุ"><input style={inputStyle} placeholder="ถ้ามี" value={form.note} onChange={(e) => update("note", e.target.value)} /></Field>
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
   FUEL VIEW (บันทึก + สรุปการเติมน้ำมัน)
---------------------------------------------------------------- */

// คืนค่าวันจันทร์ของสัปดาห์นั้น (ใช้จัดกลุ่มรายสัปดาห์)
function startOfWeek(dateStr) {
  const d = new Date(dateStr);
  const day = (d.getDay() + 6) % 7; // จันทร์ = 0
  d.setDate(d.getDate() - day);
  return d.toISOString().slice(0, 10);
}

function FuelView({ vehicles, fuelLogs, onAdd, onUpdate, onDelete, dark }) {
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
        <button onClick={() => setFormState("add")} className="flex items-center gap-2 rounded-lg px-4 py-2" style={{ background: "var(--accent-frost)", color: "#FFFFFF", fontSize: 13, fontWeight: 700, height: 38 }}>
          <Plus size={16} />บันทึกการเติมน้ำมัน
        </button>
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

      {formState && (
        <FuelFormModal
          initial={formState === "add" ? null : formState}
          vehicles={vehicles}
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
   DRIVERS VIEW
---------------------------------------------------------------- */

function DriversView({ drivers, vehiclePlates, onAdd, onUpdate, onDelete }) {
  const [formState, setFormState] = useState(null);
  const [deleteTarget, setDeleteTarget] = useState(null);

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
                  <div className="flex items-center justify-center rounded-full" style={{ width: 42, height: 42, background: "rgba(14,143,160,0.15)", color: "var(--accent-frost)" }}><User size={20} /></div>
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
    </div>
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
  const [loading, setLoading] = useState(true);
  const [errorMsg, setErrorMsg] = useState("");

  useEffect(() => { loadAll(); }, []);

  async function loadAll() {
    setLoading(true);
    try {
      const [v, r, d, fl] = await Promise.all([fetchVehicles(), fetchRepairs(), fetchDrivers(), fetchFuelLogs()]);
      setVehicles(v);
      setRepairs(r);
      setDrivers(d);
      setFuelLogs(fl);
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
    const { error } = await supabase.from("drivers").delete().eq("id", id);
    if (error) throw error;
    await loadAll();
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
      "--bg": theme.bg, "--surface": theme.surface, "--surface-2": theme.surface2, "--border": theme.border,
      "--text": theme.text, "--text-muted": theme.textMuted, "--accent-frost": theme.accent,
      background: "var(--bg)", minHeight: "100vh", fontFamily: "'Inter', sans-serif", transition: "background .2s ease",
    }}>
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
            {tab === "repairlog" && <RepairsLogView vehicles={computedVehicles} repairs={repairs} />}
            {tab === "fuel" && <FuelView vehicles={vehicles} fuelLogs={fuelLogs} onAdd={handleAddFuel} onUpdate={handleUpdateFuel} onDelete={handleDeleteFuel} dark={dark} />}
            {tab === "drivers" && (
              <DriversView drivers={drivers} vehiclePlates={vehicles.map((v) => v.id)} onAdd={handleAddDriver} onUpdate={handleUpdateDriver} onDelete={handleDeleteDriver} />
            )}
          </>
        )}
      </main>

      <footer className="mx-auto px-5 py-6" style={{ maxWidth: 1720, color: "var(--text-muted)", fontSize: 12 }}>
        ข้อมูลบันทึกลงฐานข้อมูลจริง ใช้งานพร้อมกันได้หลายคน · CPRAM Fleet Management
      </footer>
    </div>
  );
} <div className="grid grid-cols-2 gap-3">
          <Field label="วันหมดอายุใบขับขี่ *"><input style={inputStyle} type="date" value={form.expiry} onChange={(e) => update("expiry", e.target.value)} /></Field>
          <Field label="เบอร์โทรศัพท์"><input style={inputStyle} placeholder="081-XXX-XXXX" value={form.phone} onChange={(e) => update("phone", e.target.value)} /></Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="รถประจำตำแหน่ง">
            <select style={inputStyle} value={form.vehicle} onChange={(e) => update("vehicle", e.target.value)}>
              <option value="สำรอง">รถสำรอง / ไม่ระบุ</option>
              {vehiclePlates.map((p) => <option key={p} value={p}>{p}</option>)}
            </select>
          </Field>
          <Field label="อายุงาน (ปี)"><input style={inputStyle} type="number" value={form.years} onChange={(e) => update("years", e.target.value)} /></Field>
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
   FUEL LOG FORM MODAL
---------------------------------------------------------------- */

function FuelFormModal({ plate, onClose, onSave }) {
  const [form, setForm] = useState({
    date: todayISO(),
    odometer: "",
    liters: "",
    cost: "",
    station: "",
    driver: "",
  });
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);
  const update = (key, val) => setForm((f) => ({ ...f, [key]: val }));

  async function handleSubmit(e) {
    e.preventDefault();
    if (!form.date) return setError("กรุณาเลือกวันที่เติมน้ำมัน");
    if (!form.liters || Number(form.liters) <= 0) return setError("กรุณากรอกจำนวนลิตร");
    if (!form.cost || Number(form.cost) <= 0) return setError("กรุณากรอกจำนวนเงิน");

    setSaving(true);
    setError("");
    try {
      await onSave({
        plate,
        date: form.date,
        odometer: Number(form.odometer) || 0,
        liters: Number(form.liters) || 0,
        cost: Number(form.cost) || 0,
        station: form.station.trim() || "-",
        driver: form.driver.trim() || "-",
      });
      onClose();
    } catch (err) {
      setError(err.message || "บันทึกไม่สำเร็จ กรุณาลองใหม่อีกครั้ง");
    } finally {
      setSaving(false);
    }
  }

  return (
    <ModalShell title="บันทึกการเติมน้ำมัน" onClose={onClose}>
      <div style={{ marginBottom: 12 }}><PlateBadge plate={plate} /></div>
      <form onSubmit={handleSubmit} className="flex flex-col gap-3">
        <div className="grid grid-cols-2 gap-3">
          <Field label="วันที่เติม *"><input style={inputStyle} type="date" value={form.date} onChange={(e) => update("date", e.target.value)} /></Field>
          <Field label="เลขไมล์ปัจจุบัน (กม.)"><input style={inputStyle} type="number" placeholder="เช่น 125000" value={form.odometer} onChange={(e) => update("odometer", e.target.value)} /></Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="จำนวนลิตร *"><input style={inputStyle} type="number" step="0.01" placeholder="0.00" value={form.liters} onChange={(e) => update("liters", e.target.value)} /></Field>
          <Field label="จำนวนเงิน (บาท) *"><input style={inputStyle} type="number" placeholder="0.00" value={form.cost} onChange={(e) => update("cost", e.target.value)} /></Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="ปั๊มน้ำมัน / สถานที่เติม"><input style={inputStyle} placeholder="เช่น ปั๊ม PTT สาขา..." value={form.station} onChange={(e) => update("station", e.target.value)} /></Field>
          <Field label="ผู้บันทึก / คนขับ"><input style={inputStyle} placeholder="ชื่อคนขับ" value={form.driver} onChange={(e) => update("driver", e.target.value)} /></Field>
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
   VEHICLE DETAIL MODAL
---------------------------------------------------------------- */

function VehicleDetailModal({ vehicle, repairs, fuelLogs, onClose, onEdit, onDeleteRepair }) {
  const [tab, setTab] = useState("repairs"); // repairs | fuel | specs
  const v = computeVehicle(vehicle);
  const vRepairs = repairs.filter((r) => r.plate === v.id);
  const vFuel = fuelLogs.filter((f) => f.plate === v.id);
  const totalRepairCost = sumCost(vRepairs);
  const totalFuelCost = vFuel.reduce((s, f) => s + (Number(f.cost) || 0), 0);

  return (
    <ModalShell title={`รายละเอียดรถทะเบียน ${v.id}`} onClose={onClose} width={760}>
      <div className="flex items-center justify-between pb-4 border-b mb-4" style={{ borderColor: "var(--border)" }}>
        <div className="flex items-center gap-3">
          <PlateBadge plate={v.id} size="lg" />
          <div>
            <div style={{ fontSize: 16, fontWeight: 700, color: "var(--text)" }}>{v.brand} {v.model} ({v.year})</div>
            <div className="flex items-center gap-2 mt-1">
              <TempChip temp={v.temp} />
              <StatusChip status={v.status} reason={v.reason} />
            </div>
          </div>
        </div>
        <button onClick={() => { onClose(); onEdit(v); }} className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg" style={{ background: "var(--surface-2)", border: "1px solid var(--border)", color: "var(--text)", fontSize: 12, fontWeight: 600 }}>
          <Pencil size={14} />แก้ไขข้อมูลรถ
        </button>
      </div>

      {/* KPI Cards */}
      <div className="grid grid-cols-3 gap-3 mb-5">
        <Card style={{ padding: 12, background: "var(--surface-2)" }}>
          <div style={{ fontSize: 11, color: "var(--text-muted)" }}>ค่าซ่อมบำรุงรวม</div>
          <div style={{ fontSize: 16, fontWeight: 700, color: "var(--accent-frost)", marginTop: 2 }}>{fmtMoney(totalRepairCost)} บ.</div>
        </Card>
        <Card style={{ padding: 12, background: "var(--surface-2)" }}>
          <div style={{ fontSize: 11, color: "var(--text-muted)" }}>ค่าน้ำมันรวม</div>
          <div style={{ fontSize: 16, fontWeight: 700, color: "#E0A85B", marginTop: 2 }}>{fmtMoney(totalFuelCost)} บ.</div>
        </Card>
        <Card style={{ padding: 12, background: "var(--surface-2)" }}>
          <div style={{ fontSize: 11, color: "var(--text-muted)" }}>คนขับประจำ</div>
          <div style={{ fontSize: 14, fontWeight: 700, color: "var(--text)", marginTop: 2 }}>{v.driver || "-"}</div>
        </Card>
      </div>

      {/* Sub Tabs */}
      <div className="flex gap-2 border-b mb-4" style={{ borderColor: "var(--border)" }}>
        {[
          { id: "repairs", label: `ประวัติการซ่อม (${vRepairs.length})`, icon: Wrench },
          { id: "fuel", label: `ประวัติการเติมน้ำมัน (${vFuel.length})`, icon: Fuel },
          { id: "specs", label: "กำหนดการ & สเปค", icon: ClipboardList },
        ].map((t) => {
          const Icon = t.icon;
          const active = tab === t.id;
          return (
            <button key={t.id} onClick={() => setTab(t.id)} className="flex items-center gap-1.5 py-2 px-3 text-xs font-semibold border-b-2" style={{ borderBottomColor: active ? "var(--accent-frost)" : "transparent", color: active ? "var(--accent-frost)" : "var(--text-muted)" }}>
              <Icon size={14} />{t.label}
            </button>
          );
        })}
      </div>

      {/* Tab Contents */}
      {tab === "repairs" && (
        <div className="flex flex-col gap-2 max-h-80 overflow-y-auto">
          {vRepairs.length === 0 ? (
            <div className="text-center py-8 text-xs style-muted">ไม่มีประวัติการซ่อมบำรุง</div>
          ) : (
            vRepairs.map((r) => (
              <div key={r.id} className="flex items-center justify-between p-3 rounded-lg" style={{ background: "var(--surface-2)", border: "1px solid var(--border)" }}>
                <div>
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-bold" style={{ color: "var(--text)" }}>{fmtDate(r.date)}</span>
                    <span className="text-xs px-2 py-0.5 rounded" style={{ background: "var(--surface)", border: "1px solid var(--border)", color: "var(--text-muted)" }}>{r.type}</span>
                  </div>
                  <p className="text-xs mt-1" style={{ color: "var(--text)" }}>{r.description}</p>
                  <div className="flex gap-3 text-xs style-muted mt-1">
                    <span>อู่: {r.garage}</span>
                    {r.pr_number && <span>PR: {r.pr_number}</span>}
                    {r.po_number && <span>PO: {r.po_number}</span>}
                  </div>
                </div>
                <div className="flex items-center gap-3">
                  <span className="text-xs font-bold" style={{ color: r.cost ? "var(--text)" : "#D97706" }}>{fmtCost(r.cost)}</span>
                  <button onClick={() => onDeleteRepair(r.id)} style={iconBtnStyle} title="ลบรายการ"><Trash2 size={14} /></button>
                </div>
              </div>
            ))
          )}
        </div>
      )}

      {tab === "fuel" && (
        <div className="flex flex-col gap-2 max-h-80 overflow-y-auto">
          {vFuel.length === 0 ? (
            <div className="text-center py-8 text-xs style-muted">ไม่มีประวัติการเติมน้ำมัน</div>
          ) : (
            vFuel.map((f, idx) => (
              <div key={idx} className="flex items-center justify-between p-3 rounded-lg" style={{ background: "var(--surface-2)", border: "1px solid var(--border)" }}>
                <div>
                  <div className="text-xs font-bold" style={{ color: "var(--text)" }}>{fmtDate(f.date)}</div>
                  <div className="text-xs style-muted mt-1">
                    {f.odometer ? `เลขไมล์ ${fmtMoney(f.odometer)} กม.` : "ไม่ระบุเลขไมล์"} | {f.station}
                  </div>
                </div>
                <div className="text-right">
                  <div className="text-xs font-bold" style={{ color: "var(--text)" }}>{fmtMoney(f.cost)} บ.</div>
                  <div className="text-xs style-muted">{f.liters} ลิตร</div>
                </div>
              </div>
            ))
          )}
        </div>
      )}

      {tab === "specs" && (
        <div className="grid grid-cols-2 gap-4 text-xs">
          <div className="flex flex-col gap-2">
            <SubHeading>กำหนดการดูแลรักษา</SubHeading>
            <div className="flex justify-between py-1 border-b" style={{ borderColor: "var(--border)" }}><span>PM เครื่องยนต์:</span><DueChip status={v.pmStatus} daysLeft={v.daysLeft} /></div>
            <div className="flex justify-between py-1 border-b" style={{ borderColor: "var(--border)" }}><span>PM ตู้เย็น:</span><DueChip status={v.coolingPm.status} daysLeft={v.coolingPm.daysLeft} /></div>
            <div className="flex justify-between py-1 border-b" style={{ borderColor: "var(--border)" }}><span>คาลิเบรทตู้เย็น:</span><DueChip status={v.calibration.status} daysLeft={v.calibration.daysLeft} /></div>
            <div className="flex justify-between py-1 border-b" style={{ borderColor: "var(--border)" }}><span>ภาษี/พ.ร.บ.:</span><DueChip status={v.tax.status} daysLeft={v.tax.daysLeft} /></div>
          </div>
          <div className="flex flex-col gap-2">
            <SubHeading>ข้อมูลตัวรถ / ตู้บรรทุก</SubHeading>
            <div className="flex justify-between py-1 border-b" style={{ borderColor: "var(--border)" }}><span>ประเภทตู้:</span><TempChip temp={v.temp} /></div>
            <div className="flex justify-between py-1 border-b" style={{ borderColor: "var(--border)" }}><span>แบรนด์ตู้เย็น:</span><span>{v.cooling_brand || "-"}</span></div>
            <div className="flex justify-between py-1 border-b" style={{ borderColor: "var(--border)" }}><span>จำนวนล้อ / ความยาว:</span><span>{v.wheels} ล้อ / {v.length_m} ม.</span></div>
            <div className="flex justify-between py-1 border-b" style={{ borderColor: "var(--border)" }}><span>ถังน้ำมัน:</span><span>{v.fuel_tank_liters} ลิตร</span></div>
            <div className="flex justify-between py-1 border-b" style={{ borderColor: "var(--border)" }}><span>ลิฟท์ท้าย (Tail Lift):</span><span>{v.has_tail_lift ? "มี" : "ไม่มี"}</span></div>
          </div>
        </div>
      )}
    </ModalShell>
  );
}

/* ---------------------------------------------------------------
   DASHBOARD VIEW
---------------------------------------------------------------- */

function DashboardView({ vehicles, repairs, fuelLogs }) {
  const readyCount = vehicles.filter((v) => v.status === "ready").length;
  const notReadyCount = vehicles.length - readyCount;
  
  const overdueVehicles = vehicles.filter((v) => {
    const cv = computeVehicle(v);
    return cv.pmStatus === "overdue" || cv.coolingPm.status === "overdue" || cv.tax.status === "overdue";
  });

  const totalRepairCost = sumCost(repairs);
  const totalFuelCost = fuelLogs.reduce((s, f) => s + (Number(f.cost) || 0), 0);

  // กราฟสรุปค่าใช้จ่ายรายเดือน
  const monthlyCostData = useMemo(() => {
    const months = Array(12).fill(0).map((_, i) => ({ month: THAI_MONTHS[i], repair: 0, fuel: 0 }));
    repairs.forEach((r) => {
      if (r.date && !isPendingCost(r.cost)) {
        const m = new Date(r.date).getMonth();
        months[m].repair += Number(r.cost || 0);
      }
    });
    fuelLogs.forEach((f) => {
      if (f.date) {
        const m = new Date(f.date).getMonth();
        months[m].fuel += Number(f.cost || 0);
      }
    });
    return months;
  }, [repairs, fuelLogs]);

  return (
    <div className="flex flex-col gap-5">
      {/* KPI Cards Header */}
      <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
        <Card style={{ padding: 16 }}>
          <div className="flex items-center justify-between text-xs style-muted"><span>รถในฟลีตทั้งหมด</span><Truck size={16} /></div>
          <div className="text-2xl font-bold mt-2" style={{ color: "var(--text)" }}>{vehicles.length} <span className="text-xs font-normal">คัน</span></div>
          <div className="flex items-center gap-2 text-xs mt-2">
            <span style={{ color: "#16A34A" }}>พร้อมใช้งาน: {readyCount}</span>
            <span style={{ color: "#DC2626" }}>ซ่อมบำรุง: {notReadyCount}</span>
          </div>
        </Card>

        <Card style={{ padding: 16 }}>
          <div className="flex items-center justify-between text-xs style-muted"><span>ค่าซ่อมบำรุงรวม</span><Wrench size={16} /></div>
          <div className="text-2xl font-bold mt-2" style={{ color: "var(--accent-frost)" }}>{fmtMoney(totalRepairCost)} <span className="text-xs font-normal">บาท</span></div>
          <div className="text-xs style-muted mt-2">จากทั้งหมด {repairs.length} รายการ</div>
        </Card>

        <Card style={{ padding: 16 }}>
          <div className="flex items-center justify-between text-xs style-muted"><span>ค่าน้ำมันเชื้อเพลิงรวม</span><Fuel size={16} /></div>
          <div className="text-2xl font-bold mt-2" style={{ color: "#E0A85B" }}>{fmtMoney(totalFuelCost)} <span className="text-xs font-normal">บาท</span></div>
          <div className="text-xs style-muted mt-2">จากบันทึก {fuelLogs.length} ครั้ง</div>
        </Card>

        <Card style={{ padding: 16 }}>
          <div className="flex items-center justify-between text-xs style-muted"><span>ต้องดูแลเร่งด่วน (Overdue)</span><AlertTriangle size={16} style={{ color: "#DC2626" }} /></div>
          <div className="text-2xl font-bold mt-2" style={{ color: "#DC2626" }}>{overdueVehicles.length} <span className="text-xs font-normal">คัน</span></div>
          <div className="text-xs style-muted mt-2">เกินกำหนด PM หรือภาษี</div>
        </Card>
      </div>

      {/* Monthly Expense Chart */}
      <Card style={{ padding: 20 }}>
        <SectionTitle icon={BarChart} sub="เปรียบเทียบค่าใช้จ่ายซ่อมบำรุงและค่าน้ำมันรายเดือน">แนวโน้มค่าใช้จ่ายประจำปี</SectionTitle>
        <div style={{ width: "100%", height: 260 }}>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={monthlyCostData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="var(--border)" />
              <XAxis dataKey="month" tick={{ fill: "var(--text-muted)", fontSize: 12 }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fill: "var(--text-muted)", fontSize: 11 }} axisLine={false} tickLine={false} />
              <Tooltip formatter={(v) => [`${fmtMoney(v)} บาท`, ""]} contentStyle={{ background: "var(--surface-2)", borderColor: "var(--border)", borderRadius: 8, color: "var(--text)", fontSize: 12 }} />
              <Bar dataKey="repair" name="ค่าซ่อมบำรุง" fill="#0E8FA0" radius={[4, 4, 0, 0]} />
              <Bar dataKey="fuel" name="ค่าน้ำมัน" fill="#E0A85B" radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </Card>
    </div>
  );
}

/* ---------------------------------------------------------------
   MAIN DASHBOARD APP LAYOUT
---------------------------------------------------------------- */

export default function FleetDashboard() {
  const [theme, setTheme] = useState("light");
  const [activeTab, setActiveTab] = useState("vehicles"); // vehicles | repairs | drivers | fuel | overview
  
  const [vehicles, setVehicles] = useState([]);
  const [repairs, setRepairs] = useState([]);
  const [fuelLogs, setFuelLogs] = useState([]);
  const [drivers, setDrivers] = useState([]);
  
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");

  // Modals state
  const [vModal, setVModal] = useState({ open: false, initial: null });
  const [rModal, setRModal] = useState({ open: false, plate: "", initial: null });
  const [dModal, setDModal] = useState({ open: false, initial: null });
  const [fModal, setFModal] = useState({ open: false, plate: "" });
  const [detailModal, setDetailModal] = useState({ open: false, vehicle: null });
  const [deleteConfirm, setDeleteConfirm] = useState({ open: false, label: "", onConfirm: null, pwd: false });

  // Load Data
  const loadAllData = async () => {
    setLoading(true);
    try {
      const [vData, rData, fData, dData] = await Promise.all([
        fetchVehicles(), fetchRepairs(), fetchFuelLogs(), fetchDrivers()
      ]);
      setVehicles(vData || []);
      setRepairs(rData || []);
      setFuelLogs(fData || []);
      setDrivers(dData || []);
    } catch (err) {
      console.error("Error loading fleet data:", err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadAllData();
  }, []);

  // Sync css variables based on theme
  useEffect(() => {
    const t = THEMES[theme];
    Object.keys(t).forEach((k) => {
      document.documentElement.style.setProperty(`--${k.replace(/[A-Z]/g, (m) => `-${m.toLowerCase()}`)}`, t[k]);
    });
  }, [theme]);

  // Vehicle Save / Delete
  async function handleSaveVehicle(data) {
    if (data.oldId && data.oldId !== data.id) {
      await supabase.from("vehicles").delete().eq("id", data.oldId);
    }
    const { oldId, ...payload } = data;
    const { error } = await supabase.from("vehicles").upsert(payload);
    if (error) throw error;
    await loadAllData();
  }

  async function handleDeleteVehicle(id) {
    setDeleteConfirm({
      open: true,
      label: `คุณแน่ใจหรือไม่ว่าต้องการลบรถทะเบียน ${id}?`,
      pwd: true,
      onConfirm: async () => {
        await supabase.from("vehicles").delete().eq("id", id);
        setDeleteConfirm({ open: false });
        await loadAllData();
      },
    });
  }

  // Repair Save / Delete
  async function handleSaveRepair(data) {
    const { error } = await supabase.from("repairs").upsert(data);
    if (error) throw error;
    await loadAllData();
  }

  async function handleDeleteRepair(id) {
    setDeleteConfirm({
      open: true,
      label: "คุณแน่ใจหรือไม่ว่าต้องการลบรายการซ่อมนี้?",
      pwd: false,
      onConfirm: async () => {
        await supabase.from("repairs").delete().eq("id", id);
        setDeleteConfirm({ open: false });
        await loadAllData();
      },
    });
  }

  // Fuel Save
  async function handleSaveFuel(data) {
    const { error } = await supabase.from("fuel_logs").insert(data);
    if (error) throw error;
    await loadAllData();
  }

  // Driver Save
  async function handleSaveDriver(data) {
    const { error } = await supabase.from("drivers").upsert(data);
    if (error) throw error;
    await loadAllData();
  }

  // Filtered List
  const filteredVehicles = useMemo(() => {
    return vehicles.filter((v) => v.id.toLowerCase().includes(search.toLowerCase()) || v.brand.toLowerCase().includes(search.toLowerCase()) || v.driver.toLowerCase().includes(search.toLowerCase()));
  }, [vehicles, search]);

  return (
    <div style={{ background: "var(--bg)", minHeight: "100vh", color: "var(--text)", fontFamily: "'Inter', sans-serif" }}>
      {/* Header */}
      <header style={{ background: "var(--header-bg)", backdropFilter: "blur(8px)", borderBottom: "1px solid var(--border)", position: "sticky", top: 0, zIndex: 40, padding: "12px 24px" }} className="flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div style={{ background: "var(--accent-frost)", padding: 8, borderRadius: 10, color: "#fff" }}><Truck size={22} /></div>
          <div>
            <h1 style={{ fontSize: 18, fontWeight: 800, fontFamily: "'Space Grotesk', sans-serif", letterSpacing: -0.5 }}>COLD CHAIN FLEET</h1>
            <p style={{ fontSize: 11, color: "var(--text-muted)" }}>ระบบบริหารจัดการฟลีตรถขนส่งสินค้าแช่เย็น</p>
          </div>
        </div>

        <div className="flex items-center gap-3">
          <button onClick={() => setTheme(theme === "light" ? "dark" : "light")} style={{ background: "var(--surface-2)", border: "1px solid var(--border)", padding: 8, borderRadius: 8, color: "var(--text)", cursor: "pointer" }}>
            {theme === "light" ? <Moon size={16} /> : <Sun size={16} />}
          </button>
          <button onClick={() => setVModal({ open: true, initial: null })} className="flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-bold" style={{ background: "var(--accent-frost)", color: "#fff" }}>
            <Plus size={16} /> เพิ่มรถคันใหม่
          </button>
        </div>
      </header>

      {/* Main Content Layout */}
      <main className="max-w-7xl mx-auto p-6">
        {/* Navigation Tabs & Search */}
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-6">
          <div className="flex gap-2 border-b md:border-b-0 pb-2 md:pb-0" style={{ borderColor: "var(--border)" }}>
            {[
              { id: "overview", label: "ภาพรวมระบบ", icon: GaugeIcon },
              { id: "vehicles", label: `รถในฟลีต (${vehicles.length})`, icon: Truck },
              { id: "repairs", label: "ประวัติการซ่อม", icon: Wrench },
              { id: "fuel", label: "บันทึกค่าน้ำมัน", icon: Fuel },
              { id: "drivers", label: `คนขับรถ (${drivers.length})`, icon: User },
            ].map((t) => {
              const Icon = t.icon;
              const active = activeTab === t.id;
              return (
                <button key={t.id} onClick={() => setActiveTab(t.id)} className="flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-bold transition-all" style={{ background: active ? "var(--accent-frost)" : "var(--surface)", color: active ? "#fff" : "var(--text-muted)", border: `1px solid ${active ? "var(--accent-frost)" : "var(--border)"}` }}>
                  <Icon size={15} />{t.label}
                </button>
              );
            })}
          </div>

          <div className="flex items-center gap-2">
            <div className="relative">
              <Search size={15} className="absolute left-3 top-2.5" style={{ color: "var(--text-muted)" }} />
              <input style={{ ...inputStyle, paddingLeft: 32, width: 220 }} placeholder="ค้นหา ทะเบียน, ยี่ห้อ, คนขับ..." value={search} onChange={(e) => setSearch(e.target.value)} />
            </div>
            <button onClick={() => exportToExcel(vehicles, "fleet_vehicles.xlsx", "Vehicles")} className="flex items-center gap-1 px-3 py-2 rounded-lg text-xs font-semibold" style={{ background: "var(--surface-2)", border: "1px solid var(--border)", color: "var(--text)" }}>
              <Download size={14} /> Export
            </button>
          </div>
        </div>

        {/* Tab Views */}
        {loading ? (
          <div className="text-center py-20 text-sm style-muted">กำลังโหลดข้อมูลฟลีตรถ...</div>
        ) : (
          <>
            {activeTab === "overview" && <DashboardView vehicles={vehicles} repairs={repairs} fuelLogs={fuelLogs} />}

            {activeTab === "vehicles" && (
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                {filteredVehicles.map((vehicle) => {
                  const v = computeVehicle(vehicle);
                  return (
                    <Card key={v.id} style={{ padding: 18 }}>
                      <div className="flex items-start justify-between mb-3">
                        <PlateBadge plate={v.id} size="lg" />
                        <div className="flex items-center gap-1">
                          <button onClick={() => setDetailModal({ open: true, vehicle: v })} style={iconBtnStyle} title="ดูรายละเอียด"><FileText size={16} /></button>
                          <button onClick={() => setVModal({ open: true, initial: v })} style={iconBtnStyle} title="แก้ไข"><Pencil size={16} /></button>
                          <button onClick={() => handleDeleteVehicle(v.id)} style={iconBtnStyle} title="ลบ"><Trash2 size={16} /></button>
                        </div>
                      </div>

                      <div className="flex items-center justify-between text-xs mb-3">
                        <span style={{ fontWeight: 700, color: "var(--text)" }}>{v.brand} {v.model} ({v.year})</span>
                        <StatusChip status={v.status} reason={v.reason} />
                      </div>

                      <div className="flex items-center gap-2 mb-4">
                        <TempChip temp={v.temp} />
                        <span className="text-xs style-muted">คนขับ: {v.driver || "-"}</span>
                      </div>

                      <div className="flex flex-col gap-1.5 pt-3 border-t text-xs" style={{ borderColor: "var(--border)" }}>
                        <div className="flex justify-between"><span>PM เครื่องยนต์:</span><DueChip status={v.pmStatus} daysLeft={v.daysLeft} /></div>
                        <div className="flex justify-between"><span>PM ตู้เย็น:</span><DueChip status={v.coolingPm.status} daysLeft={v.coolingPm.daysLeft} /></div>
                        <div className="flex justify-between"><span>ภาษี/พ.ร.บ.:</span><DueChip status={v.tax.status} daysLeft={v.tax.daysLeft} /></div>
                      </div>

                      <div className="flex gap-2 mt-4 pt-2">
                        <button onClick={() => setRModal({ open: true, plate: v.id, initial: null })} className="flex-1 py-1.5 rounded text-xs font-semibold" style={{ background: "var(--surface-2)", border: "1px solid var(--border)", color: "var(--text)" }}>+ บันทึกซ่อม</button>
                        <button onClick={() => setFModal({ open: true, plate: v.id })} className="flex-1 py-1.5 rounded text-xs font-semibold" style={{ background: "var(--surface-2)", border: "1px solid var(--border)", color: "var(--text)" }}>+ เติมน้ำมัน</button>
                      </div>
                    </Card>
                  );
                })}
              </div>
            )}

            {activeTab === "repairs" && (
              <Card style={{ padding: 18 }}>
                <SectionTitle icon={Wrench}>ประวัติรายการแจ้งซ่อมบำรุงทั้งหมด</SectionTitle>
                <div className="overflow-x-auto">
                  <table className="w-full text-xs text-left" style={{ borderCollapse: "collapse" }}>
                    <thead>
                      <tr style={{ borderBottom: "1px solid var(--border)", color: "var(--text-muted)" }}>
                        <th className="py-2.5 px-3">วันที่</th>
                        <th className="py-2.5 px-3">ทะเบียน</th>
                        <th className="py-2.5 px-3">ประเภทงาน</th>
                        <th className="py-2.5 px-3">รายละเอียด</th>
                        <th className="py-2.5 px-3">อู่/ศูนย์บริการ</th>
                        <th className="py-2.5 px-3">ค่าใช้จ่าย</th>
                        <th className="py-2.5 px-3 text-right">จัดการ</th>
                      </tr>
                    </thead>
                    <tbody>
                      {repairs.map((r) => (
                        <tr key={r.id} style={{ borderBottom: "1px solid var(--border)" }}>
                          <td className="py-3 px-3 font-semibold">{fmtDate(r.date)}</td>
                          <td className="py-3 px-3"><PlateBadge plate={r.plate} /></td>
                          <td className="py-3 px-3">{r.type}</td>
                          <td className="py-3 px-3">{r.description}</td>
                          <td className="py-3 px-3">{r.garage}</td>
                          <td className="py-3 px-3 font-bold" style={{ color: r.cost ? "var(--text)" : "#D97706" }}>{fmtCost(r.cost)}</td>
                          <td className="py-3 px-3 text-right">
                            <button onClick={() => handleDeleteRepair(r.id)} style={iconBtnStyle}><Trash2 size={14} /></button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </Card>
            )}

            {activeTab === "drivers" && (
              <div>
                <div className="flex justify-end mb-4">
                  <button onClick={() => setDModal({ open: true, initial: null })} className="flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-bold" style={{ background: "var(--accent-frost)", color: "#fff" }}>
                    <Plus size={15} /> เพิ่มคนขับรถ
                  </button>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  {drivers.map((d) => (
                    <Card key={d.id} style={{ padding: 16 }}>
                      <div className="flex items-center gap-3 mb-3">
                        <div style={{ background: "var(--surface-2)", padding: 10, borderRadius: "50%", color: "var(--accent-frost)" }}><User size={20} /></div>
                        <div>
                          <div style={{ fontWeight: 700, fontSize: 14 }}>{d.name}</div>
                          <div style={{ fontSize: 11, color: "var(--text-muted)" }}>โทร: {d.phone || "-"}</div>
                        </div>
                      </div>
                      <div className="flex flex-col gap-1 text-xs pt-2 border-t" style={{ borderColor: "var(--border)" }}>
                        <div className="flex justify-between"><span>ใบขับขี่ประเภท:</span><span className="font-bold">{d.license_type}</span></div>
                        <div className="flex justify-between"><span>วันหมดอายุ:</span><span>{fmtDate(d.expiry)}</span></div>
                        <div className="flex justify-between"><span>ประจำรถ:</span><span className="font-bold">{d.vehicle}</span></div>
                      </div>
                    </Card>
                  ))}
                </div>
              </div>
            )}
          </>
        )}
      </main>

      {/* MODALS */}
      {vModal.open && <VehicleFormModal initial={vModal.initial} existingPlates={vehicles.map((v) => v.id)} onClose={() => setVModal({ open: false, initial: null })} onSave={handleSaveVehicle} />}
      {rModal.open && <RepairFormModal plate={rModal.plate} initial={rModal.initial} onClose={() => setRModal({ open: false, plate: "", initial: null })} onSave={handleSaveRepair} />}
      {fModal.open && <FuelFormModal plate={fModal.plate} onClose={() => setFModal({ open: false, plate: "" })} onSave={handleSaveFuel} />}
      {dModal.open && <DriverFormModal initial={dModal.initial} vehiclePlates={vehicles.map((v) => v.id)} onClose={() => setDModal({ open: false, initial: null })} onSave={handleSaveDriver} />}
      {detailModal.open && <VehicleDetailModal vehicle={detailModal.vehicle} repairs={repairs} fuelLogs={fuelLogs} onClose={() => setDetailModal({ open: false, vehicle: null })} onEdit={(v) => setVModal({ open: true, initial: v })} onDeleteRepair={handleDeleteRepair} />}
      {deleteConfirm.open && <ConfirmDelete label={deleteConfirm.label} requirePassword={deleteConfirm.pwd} onConfirm={deleteConfirm.onConfirm} onCancel={() => setDeleteConfirm({ open: false })} />}
    </div>
  );
}
