import { useEffect, useMemo, useState } from "react";
import {
  BarChart, Bar, XAxis, YAxis, ResponsiveContainer, Tooltip, CartesianGrid,
} from "recharts";
import {
  Truck, Snowflake, Droplet, Sun, Wrench, AlertTriangle, CheckCircle2,
  User, Search, X, Clock3, Pencil, Trash2, Fuel, Package,
  Download, Plus, Moon, Gauge as GaugeIcon, FileText, Lock, Hourglass, ClipboardList
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
const DELETE_PASSWORD = "500412";

/* ---------------------------------------------------------------
   HELPERS
---------------------------------------------------------------- */

function daysBetween(a, b) { return Math.round((b - a) / (1000 * 60 * 60 * 24)); }
function addDays(dateStr, n) { const d = new Date(dateStr); d.setDate(d.getDate() + n); return d; }
function fmtDate(d) { if (!d) return "-"; const date = typeof d === "string" ? new Date(d) : d; return `${date.getDate()} ${THAI_MONTHS[date.getMonth()]} ${date.getFullYear() + 543}`; }
function fmtMoney(n) { return Number(n || 0).toLocaleString("th-TH"); }

async function exportToExcel(rows, filename, sheetName) {
  const XLSX = await import("xlsx");
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, filename);
}

function isPendingCost(c) { return c === null || c === undefined; }
function fmtCost(c) { return isPendingCost(c) ? "รอประเมินราคา" : `${fmtMoney(c)} บ.`; }
function sumCost(list) { return list.reduce((s, r) => s + (isPendingCost(r.cost) ? 0 : Number(r.cost || 0)), 0); }
function todayISO() { return new Date().toISOString().slice(0, 10); }

function computeDue(lastDateStr, intervalDays) {
  if (!lastDateStr) return { next: null, daysLeft: null, status: "unknown" };
  const next = addDays(lastDateStr, intervalDays || 0);
  const daysLeft = daysBetween(TODAY, next);
  let status = "ok";
  if (daysLeft < 0) status = "overdue";
  else if (daysLeft <= 7) status = "soon";
  return { next, daysLeft, status };
}

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
  return data || [];
}
async function fetchRepairs() {
  const { data, error } = await supabase.from("repairs").select("*").order("date", { ascending: false });
  if (error) throw error;
  return data || [];
}
async function fetchFuelLogs() {
  const { data, error } = await supabase.from("fuel_logs").select("*").order("date", { ascending: false });
  if (error) throw error;
  return data || [];
}
async function fetchDrivers() {
  const { data, error } = await supabase.from("drivers").select("*").order("name");
  if (error) throw error;
  return data || [];
}

/* ---------------------------------------------------------------
   UI COMPONENTS
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
  return <div style={{ background: "var(--surface)", border: "1px solid var(--border)", borderRadius: 14, boxShadow: "0 1px 2px rgba(15,23,42,0.04)", ...style }}>{children}</div>;
}

function SectionTitle({ icon: Icon, children, sub }) {
  return (
    <div className="mb-4">
      <div className="flex items-center gap-2">
        <Icon size={18} style={{ color: "var(--accent)" }} />
        <h2 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 19, fontWeight: 700, color: "var(--text)" }}>{children}</h2>
      </div>
      {sub && <p style={{ color: "var(--text-muted)", fontSize: 13, marginTop: 4 }}>{sub}</p>}
    </div>
  );
}

function SubHeading({ children }) {
  return (
    <div style={{ fontSize: 12, fontWeight: 700, color: "var(--accent)", textTransform: "uppercase", letterSpacing: 0.5, margin: "4px 0 2px" }}>
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

const inputStyle = { background: "var(--surface2)", border: "1px solid var(--border)", borderRadius: 8, padding: "8px 10px", color: "var(--text)", fontSize: 13, outline: "none", width: "100%" };
const iconBtnStyle = { color: "var(--text-muted)", padding: 4, borderRadius: 6, cursor: "pointer", background: "transparent", border: "none" };

function ModalShell({ title, onClose, children, width = 560 }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(8,10,14,0.7)", zIndex: 50, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ maxWidth: width, width: "100%", maxHeight: "90vh", overflowY: "auto" }}>
        <Card style={{ padding: 22 }}>
          <div className="flex items-center justify-between mb-5">
            <span style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 17, fontWeight: 700, color: "var(--text)" }}>{title}</span>
            <button onClick={onClose} style={{ color: "var(--text-muted)", background: "none", border: "none", cursor: "pointer" }}><X size={18} /></button>
          </div>
          {children}
        </Card>
      </div>
    </div>
  );
}

function ConfirmDelete({ label, onConfirm, onCancel, requirePassword = false }) {
  const [pwd, setPwd] = useState("");
  const [err, setErr] = useState("");

  function handleConfirm() {
    if (requirePassword && pwd !== DELETE_PASSWORD) {
      setErr("รหัสผ่านไม่ถูกต้อง");
      return;
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
                <input style={inputStyle} type="password" value={pwd} autoFocus placeholder="กรอกรหัสผ่าน" onChange={(e) => { setPwd(e.target.value); setErr(""); }} onKeyDown={(e) => { if (e.key === "Enter") handleConfirm(); }} />
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
   FORMS & MODALS
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
      setError(err.message || "บันทึกไม่สำเร็จ");
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

        <label className="flex items-center gap-2 rounded-lg px-3 py-2.5" style={{ background: form.has_cargo ? "rgba(217,119,6,0.1)" : "var(--surface2)", border: `1px solid ${form.has_cargo ? "#D9770655" : "var(--border)"}`, cursor: "pointer" }}>
          <input type="checkbox" checked={form.has_cargo} onChange={(e) => update("has_cargo", e.target.checked)} style={{ width: 15, height: 15 }} />
          <Package size={14} style={{ color: form.has_cargo ? "#D97706" : "var(--text-muted)" }} />
          <span style={{ fontSize: 13, color: form.has_cargo ? "#D97706" : "var(--text-muted)", fontWeight: 600 }}>มีสินค้า/ตะกร้าค้างอยู่ในตู้</span>
        </label>
        {form.has_cargo && (
          <Field label="รายละเอียดสินค้าที่ค้าง"><input style={inputStyle} value={form.cargo_note} onChange={(e) => update("cargo_note", e.target.value)} /></Field>
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
          <Field label="ยี่ห้อเครื่องทำความเย็น"><input style={inputStyle} value={form.cooling_brand} onChange={(e) => update("cooling_brand", e.target.value)} /></Field>
        </div>
        <div className="grid grid-cols-3 gap-3">
          <Field label="จำนวนล้อ"><input style={inputStyle} type="number" value={form.wheels} onChange={(e) => update("wheels", e.target.value)} /></Field>
          <Field label="ความยาว (ม.)"><input style={inputStyle} type="number" step="0.1" value={form.length_m} onChange={(e) => update("length_m", e.target.value)} /></Field>
          <Field label="ถังน้ำมัน (ลิตร)"><input style={inputStyle} type="number" value={form.fuel_tank_liters} onChange={(e) => update("fuel_tank_liters", e.target.value)} /></Field>
        </div>

        <SubHeading>กำหนดการบำรุงรักษา</SubHeading>
        <div className="grid grid-cols-2 gap-3">
          <Field label="PM เครื่องยนต์ล่าสุด"><input style={inputStyle} type="date" value={form.last_pm} onChange={(e) => update("last_pm", e.target.value)} /></Field>
          <Field label="รอบ PM เครื่องยนต์ (วัน)"><input style={inputStyle} type="number" value={form.pm_interval} onChange={(e) => update("pm_interval", e.target.value)} /></Field>
        </div>

        {error && <div style={{ fontSize: 12, color: "#DC2626", background: "rgba(220,38,38,0.1)", borderRadius: 8, padding: "8px 10px" }}>{error}</div>}
        <div className="flex items-center justify-end gap-2 mt-2">
          <button type="button" onClick={onClose} style={{ ...inputStyle, width: "auto", cursor: "pointer" }}>ยกเลิก</button>
          <button type="submit" disabled={saving} style={{ background: "var(--accent)", color: "#FFFFFF", border: "none", borderRadius: 8, padding: "8px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
            {saving ? "กำลังบันทึก..." : "บันทึก"}
          </button>
        </div>
      </form>
    </ModalShell>
  );
}

function RepairFormModal({ plate, initial, onClose, onSave }) {
  const isEdit = !!initial;
  const [form, setForm] = useState(() => initial ? {
    ...initial,
    cost: isPendingCost(initial.cost) ? "" : String(initial.cost),
    pendingEstimate: isPendingCost(initial.cost),
  } : {
    date: todayISO(), type: REPAIR_TYPES[0], description: "", cost: "", garage: "", status: "เสร็จสิ้น", pendingEstimate: false, pr_number: "", po_number: "",
  });
  const [error, setError] = useState("");
  const update = (key, val) => setForm((f) => ({ ...f, [key]: val }));

  async function handleSubmit(e) {
    e.preventDefault();
    if (!form.description.trim()) return setError("กรุณากรอกรายละเอียดงานซ่อม");
    try {
      await onSave({
        id: isEdit ? initial.id : undefined,
        plate, date: form.date, type: form.type, description: form.description.trim(),
        cost: form.pendingEstimate ? null : (Number(form.cost) || 0),
        garage: form.garage.trim() || "-", status: form.status,
        pr_number: form.pr_number.trim(), po_number: form.po_number.trim(),
      });
      onClose();
    } catch (err) {
      setError(err.message || "บันทึกไม่สำเร็จ");
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
        <Field label="รายละเอียดงานซ่อม *"><input style={inputStyle} value={form.description} onChange={(e) => update("description", e.target.value)} /></Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="ค่าใช้จ่าย (บาท)">
            <input style={{ ...inputStyle, opacity: form.pendingEstimate ? 0.5 : 1 }} type="number" disabled={form.pendingEstimate} value={form.pendingEstimate ? "" : form.cost} onChange={(e) => update("cost", e.target.value)} />
          </Field>
          <Field label="อู่/ศูนย์บริการ"><input style={inputStyle} value={form.garage} onChange={(e) => update("garage", e.target.value)} /></Field>
        </div>
        <label className="flex items-center gap-2 rounded-lg px-3 py-2" style={{ background: "var(--surface2)", border: "1px solid var(--border)", cursor: "pointer" }}>
          <input type="checkbox" checked={form.pendingEstimate} onChange={(e) => update("pendingEstimate", e.target.checked)} />
          <Hourglass size={14} style={{ color: "#D97706" }} />
          <span style={{ fontSize: 13, color: "var(--text)" }}>ยังไม่ทราบค่าใช้จ่าย (รอประเมินราคา)</span>
        </label>
        {error && <div style={{ fontSize: 12, color: "#DC2626" }}>{error}</div>}
        <div className="flex justify-end gap-2 mt-2">
          <button type="button" onClick={onClose} style={{ ...inputStyle, width: "auto" }}>ยกเลิก</button>
          <button type="submit" style={{ background: "var(--accent)", color: "#fff", border: "none", borderRadius: 8, padding: "8px 18px", fontWeight: 700 }}>บันทึก</button>
        </div>
      </form>
    </ModalShell>
  );
}

function DriverFormModal({ initial, onClose, onSave, vehiclePlates }) {
  const isEdit = !!initial;
  const [form, setForm] = useState(() => initial ? { ...initial, years: String(initial.years) } : {
    name: "", license: "", license_type: "ท.2", expiry: todayISO(), phone: "", vehicle: "สำรอง", years: "1",
  });
  const [error, setError] = useState("");
  const update = (key, val) => setForm((f) => ({ ...f, [key]: val }));

  async function handleSubmit(e) {
    e.preventDefault();
    if (!form.name.trim() || !form.license.trim()) return setError("กรุณากรอกข้อมูลสำคัญให้ครบถ้วน");
    try {
      await onSave({
        id: isEdit ? initial.id : undefined,
        name: form.name.trim(), license: form.license.trim(), license_type: form.license_type,
        expiry: form.expiry, phone: form.phone.trim(), vehicle: form.vehicle, years: Number(form.years) || 0,
      });
      onClose();
    } catch (err) {
      setError(err.message || "บันทึกไม่สำเร็จ");
    }
  }

  return (
    <ModalShell title={isEdit ? "แก้ไขพนักงานขับรถ" : "เพิ่มพนักงานขับรถ"} onClose={onClose}>
      <form onSubmit={handleSubmit} className="flex flex-col gap-3">
        <Field label="ชื่อ-นามสกุล *"><input style={inputStyle} value={form.name} onChange={(e) => update("name", e.target.value)} /></Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="เลขที่ใบขับขี่ *"><input style={inputStyle} value={form.license} onChange={(e) => update("license", e.target.value)} /></Field>
          <Field label="ประเภทใบขับขี่">
            <select style={inputStyle} value={form.license_type} onChange={(e) => update("license_type", e.target.value)}>
              {DRIVER_LICENSE_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
            </select>
          </Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="วันหมดอายุใบขับขี่ *"><input style={inputStyle} type="date" value={form.expiry} onChange={(e) => update("expiry", e.target.value)} /></Field>
          <Field label="เบอร์โทรศัพท์"><input style={inputStyle} value={form.phone} onChange={(e) => update("phone", e.target.value)} /></Field>
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
        {error && <div style={{ fontSize: 12, color: "#DC2626" }}>{error}</div>}
        <div className="flex justify-end gap-2 mt-2">
          <button type="button" onClick={onClose} style={{ ...inputStyle, width: "auto" }}>ยกเลิก</button>
          <button type="submit" style={{ background: "var(--accent)", color: "#fff", border: "none", borderRadius: 8, padding: "8px 18px", fontWeight: 700 }}>บันทึก</button>
        </div>
      </form>
    </ModalShell>
  );
}

function FuelFormModal({ plate, onClose, onSave }) {
  const [form, setForm] = useState({ date: todayISO(), odometer: "", liters: "", cost: "", station: "", driver: "" });
  const [error, setError] = useState("");
  const update = (key, val) => setForm((f) => ({ ...f, [key]: val }));

  async function handleSubmit(e) {
    e.preventDefault();
    if (!form.liters || !form.cost) return setError("กรุณากรอกจำนวนลิตรและราคาให้ถูกต้อง");
    try {
      await onSave({
        plate, date: form.date, odometer: Number(form.odometer) || 0,
        liters: Number(form.liters) || 0, cost: Number(form.cost) || 0,
        station: form.station.trim() || "-", driver: form.driver.trim() || "-",
      });
      onClose();
    } catch (err) {
      setError(err.message || "บันทึกไม่สำเร็จ");
    }
  }

  return (
    <ModalShell title="บันทึกการเติมน้ำมัน" onClose={onClose}>
      <div style={{ marginBottom: 12 }}><PlateBadge plate={plate} /></div>
      <form onSubmit={handleSubmit} className="flex flex-col gap-3">
        <div className="grid grid-cols-2 gap-3">
          <Field label="วันที่เติม *"><input style={inputStyle} type="date" value={form.date} onChange={(e) => update("date", e.target.value)} /></Field>
          <Field label="เลขไมล์ปัจจุบัน (กม.)"><input style={inputStyle} type="number" value={form.odometer} onChange={(e) => update("odometer", e.target.value)} /></Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="จำนวนลิตร *"><input style={inputStyle} type="number" step="0.01" value={form.liters} onChange={(e) => update("liters", e.target.value)} /></Field>
          <Field label="จำนวนเงิน (บาท) *"><input style={inputStyle} type="number" value={form.cost} onChange={(e) => update("cost", e.target.value)} /></Field>
        </div>
        {error && <div style={{ fontSize: 12, color: "#DC2626" }}>{error}</div>}
        <div className="flex justify-end gap-2 mt-2">
          <button type="button" onClick={onClose} style={{ ...inputStyle, width: "auto" }}>ยกเลิก</button>
          <button type="submit" style={{ background: "var(--accent)", color: "#fff", border: "none", borderRadius: 8, padding: "8px 18px", fontWeight: 700 }}>บันทึก</button>
        </div>
      </form>
    </ModalShell>
  );
}

function VehicleDetailModal({ vehicle, repairs, fuelLogs, onClose, onEdit, onDeleteRepair }) {
  const [tab, setTab] = useState("repairs");
  const v = computeVehicle(vehicle);
  const vRepairs = repairs.filter((r) => r.plate === v.id);
  const vFuel = fuelLogs.filter((f) => f.plate === v.id);

  return (
    <ModalShell title={`รายละเอียดรถ ${v.id}`} onClose={onClose} width={760}>
      <div className="flex items-center justify-between pb-4 border-b mb-4" style={{ borderColor: "var(--border)" }}>
        <div className="flex items-center gap-3">
          <PlateBadge plate={v.id} size="lg" />
          <div>
            <div style={{ fontSize: 16, fontWeight: 700, color: "var(--text)" }}>{v.brand} {v.model}</div>
            <div className="flex items-center gap-2 mt-1">
              <TempChip temp={v.temp} />
              <StatusChip status={v.status} reason={v.reason} />
            </div>
          </div>
        </div>
        <button onClick={() => { onClose(); onEdit(v); }} className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg" style={{ background: "var(--surface2)", border: "1px solid var(--border)", color: "var(--text)", fontSize: 12, fontWeight: 600 }}>
          <Pencil size={14} />แก้ไขข้อมูล
        </button>
      </div>

      <div className="flex gap-2 border-b mb-4" style={{ borderColor: "var(--border)" }}>
        {[
          { id: "repairs", label: `ประวัติการซ่อม (${vRepairs.length})`, icon: Wrench },
          { id: "fuel", label: `ประวัติค่าน้ำมัน (${vFuel.length})`, icon: Fuel },
          { id: "specs", label: "กำหนดการ & สเปค", icon: ClipboardList },
        ].map((t) => (
          <button key={t.id} onClick={() => setTab(t.id)} className="flex items-center gap-1.5 py-2 px-3 text-xs font-semibold border-b-2" style={{ borderBottomColor: tab === t.id ? "var(--accent)" : "transparent", color: tab === t.id ? "var(--accent)" : "var(--text-muted)", background: "none", borderLeft: "none", borderRight: "none", borderTop: "none" }}>
            <t.icon size={14} />{t.label}
          </button>
        ))}
      </div>

      {tab === "repairs" && (
        <div className="flex flex-col gap-2 max-h-80 overflow-y-auto">
          {vRepairs.length === 0 ? <div className="text-center py-8 text-xs" style={{ color: "var(--text-muted)" }}>ไม่มีประวัติการซ่อม</div> : vRepairs.map((r) => (
            <div key={r.id} className="flex items-center justify-between p-3 rounded-lg" style={{ background: "var(--surface2)", border: "1px solid var(--border)" }}>
              <div>
                <div className="text-xs font-bold">{fmtDate(r.date)} - {r.type}</div>
                <p className="text-xs mt-0.5">{r.description}</p>
              </div>
              <div className="flex items-center gap-3">
                <span className="text-xs font-bold">{fmtCost(r.cost)}</span>
                <button onClick={() => onDeleteRepair(r.id)} style={iconBtnStyle}><Trash2 size={14} /></button>
              </div>
            </div>
          ))}
        </div>
      )}

      {tab === "fuel" && (
        <div className="flex flex-col gap-2 max-h-80 overflow-y-auto">
          {vFuel.length === 0 ? <div className="text-center py-8 text-xs" style={{ color: "var(--text-muted)" }}>ไม่มีประวัติการเติมน้ำมัน</div> : vFuel.map((f, idx) => (
            <div key={idx} className="flex items-center justify-between p-3 rounded-lg" style={{ background: "var(--surface2)", border: "1px solid var(--border)" }}>
              <div>
                <div className="text-xs font-bold">{fmtDate(f.date)}</div>
                <div className="text-xs" style={{ color: "var(--text-muted)" }}>{f.odometer ? `เลขไมล์ ${fmtMoney(f.odometer)} กม.` : "ไม่ระบุเลขไมล์"}</div>
              </div>
              <div className="text-right">
                <div className="text-xs font-bold">{fmtMoney(f.cost)} บ.</div>
                <div className="text-xs" style={{ color: "var(--text-muted)" }}>{f.liters} ลิตร</div>
              </div>
            </div>
          ))}
        </div>
      )}

      {tab === "specs" && (
        <div className="grid grid-cols-2 gap-4 text-xs">
          <div className="flex flex-col gap-2">
            <SubHeading>กำหนดการดูแลรักษา</SubHeading>
            <div className="flex justify-between py-1 border-b" style={{ borderColor: "var(--border)" }}><span>PM เครื่องยนต์:</span><DueChip status={v.pmStatus} daysLeft={v.daysLeft} /></div>
            <div className="flex justify-between py-1 border-b" style={{ borderColor: "var(--border)" }}><span>PM ตู้เย็น:</span><DueChip status={v.coolingPm.status} daysLeft={v.coolingPm.daysLeft} /></div>
          </div>
          <div className="flex flex-col gap-2">
            <SubHeading>สเปคตัวรถ</SubHeading>
            <div className="flex justify-between py-1 border-b" style={{ borderColor: "var(--border)" }}><span>ประเภทตู้:</span><TempChip temp={v.temp} /></div>
            <div className="flex justify-between py-1 border-b" style={{ borderColor: "var(--border)" }}><span>แบรนด์ตู้เย็น:</span><span>{v.cooling_brand || "-"}</span></div>
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
  const overdueVehicles = vehicles.filter((v) => {
    const cv = computeVehicle(v);
    return cv.pmStatus === "overdue" || cv.coolingPm.status === "overdue" || cv.tax.status === "overdue";
  });

  const totalRepairCost = sumCost(repairs);
  const totalFuelCost = fuelLogs.reduce((s, f) => s + (Number(f.cost) || 0), 0);

  const monthlyCostData = useMemo(() => {
    const months = Array(12).fill(0).map((_, i) => ({ month: THAI_MONTHS[i], repair: 0, fuel: 0 }));
    repairs.forEach((r) => {
      if (r.date && !isPendingCost(r.cost)) {
        months[new Date(r.date).getMonth()].repair += Number(r.cost || 0);
      }
    });
    fuelLogs.forEach((f) => {
      if (f.date) {
        months[new Date(f.date).getMonth()].fuel += Number(f.cost || 0);
      }
    });
    return months;
  }, [repairs, fuelLogs]);

  return (
    <div className="flex flex-col gap-5">
      <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
        <Card style={{ padding: 16 }}>
          <div className="flex items-center justify-between text-xs" style={{ color: "var(--text-muted)" }}><span>รถในฟลีตทั้งหมด</span><Truck size={16} /></div>
          <div className="text-2xl font-bold mt-2" style={{ color: "var(--text)" }}>{vehicles.length} <span className="text-xs font-normal">คัน</span></div>
          <div className="flex items-center gap-2 text-xs mt-2">
            <span style={{ color: "#16A34A" }}>พร้อม: {readyCount}</span>
            <span style={{ color: "#DC2626" }}>ไม่พร้อม: {vehicles.length - readyCount}</span>
          </div>
        </Card>
        <Card style={{ padding: 16 }}>
          <div className="flex items-center justify-between text-xs" style={{ color: "var(--text-muted)" }}><span>ค่าซ่อมบำรุงรวม</span><Wrench size={16} /></div>
          <div className="text-2xl font-bold mt-2" style={{ color: "var(--accent)" }}>{fmtMoney(totalRepairCost)} <span className="text-xs font-normal">บาท</span></div>
        </Card>
        <Card style={{ padding: 16 }}>
          <div className="flex items-center justify-between text-xs" style={{ color: "var(--text-muted)" }}><span>ค่าน้ำมันเชื้อเพลิงรวม</span><Fuel size={16} /></div>
          <div className="text-2xl font-bold mt-2" style={{ color: "#E0A85B" }}>{fmtMoney(totalFuelCost)} <span className="text-xs font-normal">บาท</span></div>
        </Card>
        <Card style={{ padding: 16 }}>
          <div className="flex items-center justify-between text-xs" style={{ color: "var(--text-muted)" }}><span>ต้องดูแลเร่งด่วน</span><AlertTriangle size={16} style={{ color: "#DC2626" }} /></div>
          <div className="text-2xl font-bold mt-2" style={{ color: "#DC2626" }}>{overdueVehicles.length} <span className="text-xs font-normal">คัน</span></div>
        </Card>
      </div>

      <Card style={{ padding: 20 }}>
        <SectionTitle icon={BarChart} sub="เปรียบเทียบค่าใช้จ่ายซ่อมบำรุงและค่าน้ำมันรายเดือน">แนวโน้มค่าใช้จ่ายประจำปี</SectionTitle>
        <div style={{ width: "100%", height: 260 }}>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={monthlyCostData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="var(--border)" />
              <XAxis dataKey="month" tick={{ fill: "var(--text-muted)", fontSize: 12 }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fill: "var(--text-muted)", fontSize: 11 }} axisLine={false} tickLine={false} />
              <Tooltip formatter={(v) => [`${fmtMoney(v)} บาท`, ""]} contentStyle={{ background: "var(--surface2)", borderColor: "var(--border)", borderRadius: 8, color: "var(--text)", fontSize: 12 }} />
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
   MAIN DASHBOARD APP
---------------------------------------------------------------- */

export default function FleetDashboard() {
  const [theme, setTheme] = useState("light");
  const [activeTab, setActiveTab] = useState("vehicles");
  
  const [vehicles, setVehicles] = useState([]);
  const [repairs, setRepairs] = useState([]);
  const [fuelLogs, setFuelLogs] = useState([]);
  const [drivers, setDrivers] = useState([]);
  
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");

  const [vModal, setVModal] = useState({ open: false, initial: null });
  const [rModal, setRModal] = useState({ open: false, plate: "", initial: null });
  const [dModal, setDModal] = useState({ open: false, initial: null });
  const [fModal, setFModal] = useState({ open: false, plate: "" });
  const [detailModal, setDetailModal] = useState({ open: false, vehicle: null });
  const [deleteConfirm, setDeleteConfirm] = useState({ open: false, label: "", onConfirm: null, pwd: false });

  const loadAllData = async () => {
    setLoading(true);
    try {
      const [vData, rData, fData, dData] = await Promise.all([
        fetchVehicles(), fetchRepairs(), fetchFuelLogs(), fetchDrivers()
      ]);
      setVehicles(vData);
      setRepairs(rData);
      setFuelLogs(fData);
      setDrivers(dData);
    } catch (err) {
      console.error("Error loading fleet data:", err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { loadAllData(); }, []);

  useEffect(() => {
    const t = THEMES[theme];
    Object.keys(t).forEach((k) => {
      const cssVar = `--${k.replace(/[A-Z]/g, (m) => `-${m.toLowerCase()}`)}`;
      document.documentElement.style.setProperty(cssVar, t[k]);
    });
  }, [theme]);

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

  async function handleSaveFuel(data) {
    const { error } = await supabase.from("fuel_logs").insert(data);
    if (error) throw error;
    await loadAllData();
  }

  async function handleSaveDriver(data) {
    const { error } = await supabase.from("drivers").upsert(data);
    if (error) throw error;
    await loadAllData();
  }

  const filteredVehicles = useMemo(() => {
    return vehicles.filter((v) => 
      (v.id || "").toLowerCase().includes(search.toLowerCase()) || 
      (v.brand || "").toLowerCase().includes(search.toLowerCase()) || 
      (v.driver || "").toLowerCase().includes(search.toLowerCase())
    );
  }, [vehicles, search]);

  return (
    <div style={{ background: "var(--bg)", minHeight: "100vh", color: "var(--text)", fontFamily: "'Inter', sans-serif" }}>
      <header style={{ background: "var(--header-bg)", backdropFilter: "blur(8px)", borderBottom: "1px solid var(--border)", position: "sticky", top: 0, zIndex: 40, padding: "12px 24px" }} className="flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div style={{ background: "var(--accent)", padding: 8, borderRadius: 10, color: "#fff" }}><Truck size={22} /></div>
          <div>
            <h1 style={{ fontSize: 18, fontWeight: 800, fontFamily: "'Space Grotesk', sans-serif" }}>COLD CHAIN FLEET</h1>
            <p style={{ fontSize: 11, color: "var(--text-muted)" }}>ระบบบริหารจัดการฟลีตรถขนส่งสินค้าแช่เย็น</p>
          </div>
        </div>

        <div className="flex items-center gap-3">
          <button onClick={() => setTheme(theme === "light" ? "dark" : "light")} style={{ background: "var(--surface2)", border: "1px solid var(--border)", padding: 8, borderRadius: 8, color: "var(--text)", cursor: "pointer" }}>
            {theme === "light" ? <Moon size={16} /> : <Sun size={16} />}
          </button>
          <button onClick={() => setVModal({ open: true, initial: null })} className="flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-bold" style={{ background: "var(--accent)", color: "#fff", border: "none", cursor: "pointer" }}>
            <Plus size={16} /> เพิ่มรถคันใหม่
          </button>
        </div>
      </header>

      <main className="max-w-7xl mx-auto p-6">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-6">
          <div className="flex gap-2 border-b md:border-b-0 pb-2 md:pb-0" style={{ borderColor: "var(--border)" }}>
            {[
              { id: "overview", label: "ภาพรวมระบบ", icon: GaugeIcon },
              { id: "vehicles", label: `รถในฟลีต (${vehicles.length})`, icon: Truck },
              { id: "repairs", label: "ประวัติการซ่อม", icon: Wrench },
              { id: "drivers", label: `คนขับรถ (${drivers.length})`, icon: User },
            ].map((t) => (
              <button key={t.id} onClick={() => setActiveTab(t.id)} className="flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-bold transition-all" style={{ background: activeTab === t.id ? "var(--accent)" : "var(--surface)", color: activeTab === t.id ? "#fff" : "var(--text-muted)", border: `1px solid ${activeTab === t.id ? "var(--accent)" : "var(--border)"}`, cursor: "pointer" }}>
                <t.icon size={15} />{t.label}
              </button>
            ))}
          </div>

          <div className="flex items-center gap-2">
            <div className="relative">
              <Search size={15} className="absolute left-3 top-2.5" style={{ color: "var(--text-muted)" }} />
              <input style={{ ...inputStyle, paddingLeft: 32, width: 220 }} placeholder="ค้นหา ทะเบียน, ยี่ห้อ..." value={search} onChange={(e) => setSearch(e.target.value)} />
            </div>
            <button onClick={() => exportToExcel(vehicles, "fleet_vehicles.xlsx", "Vehicles")} className="flex items-center gap-1 px-3 py-2 rounded-lg text-xs font-semibold" style={{ background: "var(--surface2)", border: "1px solid var(--border)", color: "var(--text)", cursor: "pointer" }}>
              <Download size={14} /> Export
            </button>
          </div>
        </div>

        {loading ? (
          <div className="text-center py-20 text-sm" style={{ color: "var(--text-muted)" }}>กำลังโหลดข้อมูลฟลีตรถ...</div>
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
                          <button onClick={() => setDetailModal({ open: true, vehicle: v })} style={iconBtnStyle}><FileText size={16} /></button>
                          <button onClick={() => setVModal({ open: true, initial: v })} style={iconBtnStyle}><Pencil size={16} /></button>
                          <button onClick={() => handleDeleteVehicle(v.id)} style={iconBtnStyle}><Trash2 size={16} /></button>
                        </div>
                      </div>

                      <div className="flex items-center justify-between text-xs mb-3">
                        <span style={{ fontWeight: 700, color: "var(--text)" }}>{v.brand} {v.model} ({v.year})</span>
                        <StatusChip status={v.status} reason={v.reason} />
                      </div>

                      <div className="flex items-center gap-2 mb-4">
                        <TempChip temp={v.temp} />
                        <span className="text-xs" style={{ color: "var(--text-muted)" }}>คนขับ: {v.driver || "-"}</span>
                      </div>

                      <div className="flex flex-col gap-1.5 pt-3 border-t text-xs" style={{ borderColor: "var(--border)" }}>
                        <div className="flex justify-between"><span>PM เครื่องยนต์:</span><DueChip status={v.pmStatus} daysLeft={v.daysLeft} /></div>
                        <div className="flex justify-between"><span>PM ตู้เย็น:</span><DueChip status={v.coolingPm.status} daysLeft={v.coolingPm.daysLeft} /></div>
                        <div className="flex justify-between"><span>ภาษี/พ.ร.บ.:</span><DueChip status={v.tax.status} daysLeft={v.tax.daysLeft} /></div>
                      </div>

                      <div className="flex gap-2 mt-4 pt-2">
                        <button onClick={() => setRModal({ open: true, plate: v.id, initial: null })} className="flex-1 py-1.5 rounded text-xs font-semibold" style={{ background: "var(--surface2)", border: "1px solid var(--border)", color: "var(--text)", cursor: "pointer" }}>+ บันทึกซ่อม</button>
                        <button onClick={() => setFModal({ open: true, plate: v.id })} className="flex-1 py-1.5 rounded text-xs font-semibold" style={{ background: "var(--surface2)", border: "1px solid var(--border)", color: "var(--text)", cursor: "pointer" }}>+ เติมน้ำมัน</button>
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
                  <button onClick={() => setDModal({ open: true, initial: null })} className="flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-bold" style={{ background: "var(--accent)", color: "#fff", border: "none", cursor: "pointer" }}>
                    <Plus size={15} /> เพิ่มคนขับรถ
                  </button>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  {drivers.map((d) => (
                    <Card key={d.id} style={{ padding: 16 }}>
                      <div className="flex items-center gap-3 mb-3">
                        <div style={{ background: "var(--surface2)", padding: 10, borderRadius: "50%", color: "var(--accent)" }}><User size={20} /></div>
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
