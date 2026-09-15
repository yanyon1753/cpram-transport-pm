import { useEffect, useMemo, useState } from "react";
import {
  BarChart, Bar, XAxis, YAxis, ResponsiveContainer, Tooltip, CartesianGrid, Cell,
} from "recharts";
import {
  Truck, Snowflake, Droplet, Sun, Wrench, Calendar, AlertTriangle, CheckCircle2,
  User, Phone, Search, X, Clock3, CreditCard, ChevronRight, ClipboardList, Gauge,
  Plus, Trash2, Pencil, LogOut, Settings2, FileText, LayoutDashboard, Car, ShieldAlert
} from "lucide-react";
import { supabase } from "../lib/supabaseClient";

/* ---------------------------------------------------------------
   CONFIG & CONSTANTS
---------------------------------------------------------------- */

const TODAY = new Date(new Date().toDateString());

const TEMP_META = {
  freezer: { label: "ห้องเย็น -18°C", icon: Snowflake, bg: "#E0F2FE", color: "#0284C7", border: "#BAE6FD" },
  chiller: { label: "แช่เย็น 0-4°C", icon: Droplet, bg: "#EFF6FF", color: "#2563EB", border: "#BFDBFE" },
  ambient: { label: "อุณหภูมิห้อง", icon: Sun, bg: "#FEF3C7", color: "#D97706", border: "#FDE68A" },
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
async function fetchDrivers() {
  const { data, error } = await supabase.from("drivers").select("*").order("name");
  if (error) throw error;
  return data || [];
}

/* ---------------------------------------------------------------
   REUSABLE UI COMPONENTS (Clean Light Theme)
---------------------------------------------------------------- */

function PlateBadge({ plate, size = "md" }) {
  const parts = (plate || "").split("-");
  const prefix = parts[0] || plate || "-";
  const num = parts[1] || "";
  const isLg = size === "lg";
  return (
    <div className={`inline-flex flex-col items-center justify-center bg-slate-50 border-2 border-slate-800 rounded-lg shadow-sm ${isLg ? 'px-3 py-1.5 min-w-[100px]' : 'px-2 py-1 min-w-[80px]'}`}>
      <span className={`font-mono font-bold tracking-wider text-slate-900 leading-none ${isLg ? 'text-lg' : 'text-sm'}`}>
        {num ? `${prefix}-${num}` : prefix}
      </span>
      <span className={`font-sans font-semibold text-slate-500 leading-tight ${isLg ? 'text-[10px]' : 'text-[8px]'}`}>
        กรุงเทพมหานคร
      </span>
    </div>
  );
}

function TempChip({ temp }) {
  const meta = TEMP_META[temp] || TEMP_META.ambient;
  const Icon = meta.icon;
  return (
    <span 
      className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-semibold"
      style={{ backgroundColor: meta.bg, color: meta.color, border: `1px solid ${meta.border}` }}
    >
      <Icon size={13} />
      {meta.label}
    </span>
  );
}

function StatusChip({ status, reason }) {
  const ready = status === "ready";
  return (
    <span 
      className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-semibold ${
        ready 
          ? 'bg-emerald-50 text-emerald-700 border border-emerald-200' 
          : 'bg-rose-50 text-rose-700 border border-rose-200'
      }`}
      title={reason || ""}
    >
      {ready ? <CheckCircle2 size={13} /> : <AlertTriangle size={13} />}
      {ready ? "พร้อมใช้งาน" : "ไม่พร้อมใช้งาน"}
    </span>
  );
}

function DueChip({ status, daysLeft }) {
  if (status === "unknown" || daysLeft === null || daysLeft === undefined) {
    return (
      <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-medium bg-slate-100 text-slate-500 border border-slate-200">
        ไม่มีข้อมูล
      </span>
    );
  }
  const conf = {
    overdue: { bg: "bg-rose-50", text: "text-rose-700", border: "border-rose-200", label: `เกินกำหนด ${Math.abs(daysLeft)} วัน` },
    soon: { bg: "bg-amber-50", text: "text-amber-700", border: "border-amber-200", label: `ใกล้ครบกำหนด (${daysLeft} วัน)` },
    ok: { bg: "bg-slate-50", text: "text-slate-600", border: "border-slate-200", label: `อีก ${daysLeft} วัน` },
  }[status];

  return (
    <span className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-medium border ${conf.bg} ${conf.text} ${conf.border}`}>
      <Clock3 size={13} />
      {conf.label}
    </span>
  );
}

function Card({ children, className = "" }) {
  return (
    <div className={`bg-white border border-slate-200/80 rounded-2xl shadow-sm hover:shadow-md transition-shadow duration-200 ${className}`}>
      {children}
    </div>
  );
}

function SectionTitle({ icon: Icon, children, sub, action }) {
  return (
    <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-6">
      <div>
        <div className="flex items-center gap-2.5">
          {Icon && <Icon className="w-6 h-6 text-sky-600" />}
          <h1 className="text-xl font-bold text-slate-800 tracking-tight">{children}</h1>
        </div>
        {sub && <p className="text-sm text-slate-500 mt-1">{sub}</p>}
      </div>
      {action && <div>{action}</div>}
    </div>
  );
}

function SubHeading({ children }) {
  return (
    <div className="text-xs font-bold text-sky-700 uppercase tracking-wider mb-2 mt-4 pb-1 border-b border-slate-100">
      {children}
    </div>
  );
}

function Field({ label, children }) {
  return (
    <label className="flex flex-col gap-1.5">
      <span className="text-xs font-semibold text-slate-600">{label}</span>
      {children}
    </label>
  );
}

const inputClass = "w-full bg-slate-50 border border-slate-300 focus:border-sky-500 focus:bg-white focus:ring-2 focus:ring-sky-100 rounded-lg px-3 py-2 text-sm text-slate-800 outline-none transition-all";

function ModalShell({ title, onClose, children, width = 560 }) {
  return (
    <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center p-4 animate-fadeIn" onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ maxWidth: width }} className="w-full max-h-[90vh] overflow-y-auto">
        <Card className="p-6">
          <div className="flex items-center justify-between pb-4 mb-4 border-b border-slate-100">
            <h3 className="text-lg font-bold text-slate-800">{title}</h3>
            <button onClick={onClose} className="p-1 rounded-lg text-slate-400 hover:text-slate-600 hover:bg-slate-100 transition-colors">
              <X size={20} />
            </button>
          </div>
          {children}
        </Card>
      </div>
    </div>
  );
}

function ConfirmDelete({ label, onConfirm, onCancel }) {
  return (
    <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center p-4" onClick={onCancel}>
      <div onClick={(e) => e.stopPropagation()} className="max-w-md w-full">
        <Card className="p-6">
          <div className="flex items-center gap-3 text-rose-600 mb-3">
            <AlertTriangle size={24} />
            <h3 className="text-lg font-bold text-slate-800">ยืนยันการลบข้อมูล</h3>
          </div>
          <p className="text-sm text-slate-600 mb-6">{label}</p>
          <div className="flex items-center justify-end gap-3">
            <button onClick={onCancel} className="px-4 py-2 rounded-lg border border-slate-300 text-slate-700 text-sm font-semibold hover:bg-slate-50 transition-colors">
              ยกเลิก
            </button>
            <button onClick={onConfirm} className="px-4 py-2 rounded-lg bg-rose-600 hover:bg-rose-700 text-white text-sm font-semibold transition-colors shadow-sm">
              ยืนยันการลบ
            </button>
          </div>
        </Card>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   MODALS (Vehicle, Repair, Driver)
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
    tax_expiry: initial.tax_expiry || todayISO(),
    cooling_brand: initial.cooling_brand || "",
  } : {
    id: "", brand: "", model: "", year: "2024", temp: "chiller",
    status: "ready", reason: "", last_pm: todayISO(), pm_interval: "90", driver: "",
    cooling_brand: "", wheels: "6", length_m: "6", fuel_tank_liters: "150",
    tax_expiry: todayISO(), tire_changed: todayISO(), battery_changed: todayISO(),
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
          <Field label="ทะเบียนรถ *"><input className={inputClass} placeholder="เช่น 70-1234" value={form.id} onChange={(e) => update("id", e.target.value)} /></Field>
          <Field label="คนขับประจำ"><input className={inputClass} placeholder="ชื่อ-นามสกุล" value={form.driver} onChange={(e) => update("driver", e.target.value)} /></Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="ยี่ห้อ *"><input className={inputClass} placeholder="เช่น Isuzu" value={form.brand} onChange={(e) => update("brand", e.target.value)} /></Field>
          <Field label="รุ่น *"><input className={inputClass} placeholder="เช่น FRR90" value={form.model} onChange={(e) => update("model", e.target.value)} /></Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="ปีรถ"><input className={inputClass} type="number" value={form.year} onChange={(e) => update("year", e.target.value)} /></Field>
          <Field label="สถานะรถ">
            <select className={inputClass} value={form.status} onChange={(e) => update("status", e.target.value)}>
              <option value="ready">พร้อมใช้งาน</option>
              <option value="not_ready">ไม่พร้อมใช้งาน</option>
            </select>
          </Field>
        </div>
        {form.status === "not_ready" && (
          <Field label="เหตุผลที่ไม่พร้อมใช้งาน"><input className={inputClass} placeholder="เช่น รอซ่อมเครื่องยนต์" value={form.reason} onChange={(e) => update("reason", e.target.value)} /></Field>
        )}

        <SubHeading>สเปครถ / ระบบทำความเย็น</SubHeading>
        <div className="grid grid-cols-2 gap-3">
          <Field label="ประเภทตู้บรรทุก">
            <select className={inputClass} value={form.temp} onChange={(e) => update("temp", e.target.value)}>
              <option value="freezer">ห้องเย็น -18°C</option>
              <option value="chiller">แช่เย็น 0-4°C</option>
              <option value="ambient">อุณหภูมิห้อง</option>
            </select>
          </Field>
          <Field label="ยี่ห้อเครื่องทำความเย็น"><input className={inputClass} placeholder="เช่น Thermo King" value={form.cooling_brand} onChange={(e) => update("cooling_brand", e.target.value)} /></Field>
        </div>
        <div className="grid grid-cols-3 gap-3">
          <Field label="จำนวนล้อ"><input className={inputClass} type="number" value={form.wheels} onChange={(e) => update("wheels", e.target.value)} /></Field>
          <Field label="ความยาว (เมตร)"><input className={inputClass} type="number" step="0.1" value={form.length_m} onChange={(e) => update("length_m", e.target.value)} /></Field>
          <Field label="ถังน้ำมัน (ลิตร)"><input className={inputClass} type="number" value={form.fuel_tank_liters} onChange={(e) => update("fuel_tank_liters", e.target.value)} /></Field>
        </div>

        <SubHeading>กำหนดการบำรุงรักษา</SubHeading>
        <div className="grid grid-cols-2 gap-3">
          <Field label="PM เครื่องยนต์ล่าสุด *"><input className={inputClass} type="date" value={form.last_pm} onChange={(e) => update("last_pm", e.target.value)} /></Field>
          <Field label="รอบ PM เครื่องยนต์ (วัน)"><input className={inputClass} type="number" value={form.pm_interval} onChange={(e) => update("pm_interval", e.target.value)} /></Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="PM ตู้เย็นล่าสุด"><input className={inputClass} type="date" value={form.cooling_pm_last} onChange={(e) => update("cooling_pm_last", e.target.value)} /></Field>
          <Field label="รอบ PM ตู้เย็น (วัน)"><input className={inputClass} type="number" value={form.cooling_pm_interval} onChange={(e) => update("cooling_pm_interval", e.target.value)} /></Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="คาลิเบรทตู้เย็นล่าสุด"><input className={inputClass} type="date" value={form.calibration_last} onChange={(e) => update("calibration_last", e.target.value)} /></Field>
          <Field label="รอบคาลิเบรท (วัน)"><input className={inputClass} type="number" value={form.calibration_interval} onChange={(e) => update("calibration_interval", e.target.value)} /></Field>
        </div>
        <Field label="วันหมดอายุภาษี/พ.ร.บ."><input className={inputClass} type="date" value={form.tax_expiry} onChange={(e) => update("tax_expiry", e.target.value)} /></Field>

        {error && <div className="text-xs text-rose-600 bg-rose-50 border border-rose-200 rounded-lg p-2.5 mt-2">{error}</div>}
        
        <div className="flex items-center justify-end gap-3 mt-4 pt-3 border-t border-slate-100">
          <button type="button" onClick={onClose} className="px-4 py-2 rounded-lg border border-slate-300 text-slate-700 text-sm font-semibold hover:bg-slate-50 transition-colors">ยกเลิก</button>
          <button type="submit" disabled={saving} className="px-5 py-2 rounded-lg bg-sky-600 hover:bg-sky-700 text-white text-sm font-semibold transition-colors shadow-sm disabled:opacity-50">
            {saving ? "กำลังบันทึก..." : "บันทึก"}
          </button>
        </div>
      </form>
    </ModalShell>
  );
}

function RepairFormModal({ plate, initial, onClose, onSave }) {
  const isEdit = !!initial;
  const [form, setForm] = useState(() => initial ? { ...initial, cost: String(initial.cost) } : {
    date: todayISO(), type: REPAIR_TYPES[0], description: "", cost: "", garage: "", status: "เสร็จสิ้น",
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
        cost: Number(form.cost) || 0, garage: form.garage.trim() || "-", status: form.status,
      });
      onClose();
    } catch (err) {
      setError(err.message || "บันทึกไม่สำเร็จ");
    } finally {
      setSaving(false);
    }
  }

  return (
    <ModalShell title={isEdit ? "แก้ไขรายการซ่อม" : "เพิ่มรายการซ่อม"} onClose={onClose}>
      <div className="mb-4"><PlateBadge plate={plate} /></div>
      <form onSubmit={handleSubmit} className="flex flex-col gap-3">
        <div className="grid grid-cols-2 gap-3">
          <Field label="วันที่ซ่อม *"><input className={inputClass} type="date" value={form.date} onChange={(e) => update("date", e.target.value)} /></Field>
          <Field label="ประเภทงานซ่อม">
            <select className={inputClass} value={form.type} onChange={(e) => update("type", e.target.value)}>
              {REPAIR_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
            </select>
          </Field>
        </div>
        <Field label="รายละเอียดงานซ่อม *"><input className={inputClass} placeholder="เช่น เปลี่ยนผ้าเบรกหน้า" value={form.description} onChange={(e) => update("description", e.target.value)} /></Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="ค่าใช้จ่าย (บาท)"><input className={inputClass} type="number" value={form.cost} onChange={(e) => update("cost", e.target.value)} /></Field>
          <Field label="สถานะงาน">
            <select className={inputClass} value={form.status} onChange={(e) => update("status", e.target.value)}>
              <option value="เสร็จสิ้น">เสร็จสิ้น</option>
              <option value="กำลังดำเนินการ">กำลังดำเนินการ</option>
            </select>
          </Field>
        </div>
        <Field label="อู่ / ศูนย์บริการ"><input className={inputClass} placeholder="เช่น อู่กลาง CPRAM" value={form.garage} onChange={(e) => update("garage", e.target.value)} /></Field>
        
        {error && <div className="text-xs text-rose-600 bg-rose-50 border border-rose-200 rounded-lg p-2.5 mt-2">{error}</div>}

        <div className="flex items-center justify-end gap-3 mt-4 pt-3 border-t border-slate-100">
          <button type="button" onClick={onClose} className="px-4 py-2 rounded-lg border border-slate-300 text-slate-700 text-sm font-semibold hover:bg-slate-50 transition-colors">ยกเลิก</button>
          <button type="submit" disabled={saving} className="px-5 py-2 rounded-lg bg-sky-600 hover:bg-sky-700 text-white text-sm font-semibold transition-colors shadow-sm disabled:opacity-50">
            {saving ? "กำลังบันทึก..." : "บันทึก"}
          </button>
        </div>
      </form>
    </ModalShell>
  );
}

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
      setError(err.message || "บันทึกไม่สำเร็จ");
    } finally {
      setSaving(false);
    }
  }

  return (
    <ModalShell title={isEdit ? "แก้ไขข้อมูลพนักงานขับรถ" : "เพิ่มพนักงานขับรถ"} onClose={onClose}>
      <form onSubmit={handleSubmit} className="flex flex-col gap-3">
        <Field label="ชื่อ-นามสกุล *"><input className={inputClass} value={form.name} onChange={(e) => update("name", e.target.value)} /></Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="เลขที่ใบขับขี่ *"><input className={inputClass} placeholder="เช่น 31-045678-9" value={form.license} onChange={(e) => update("license", e.target.value)} /></Field>
          <Field label="ประเภทใบขับขี่">
            <select className={inputClass} value={form.license_type} onChange={(e) => update("license_type", e.target.value)}>
              {DRIVER_LICENSE_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
            </select>
          </Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="วันหมดอายุใบขับขี่ *"><input className={inputClass} type="date" value={form.expiry} onChange={(e) => update("expiry", e.target.value)} /></Field>
          <Field label="เบอร์โทรศัพท์"><input className={inputClass} placeholder="08x-xxx-xxxx" value={form.phone} onChange={(e) => update("phone", e.target.value)} /></Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="รถที่รับผิดชอบ">
            <select className={inputClass} value={form.vehicle} onChange={(e) => update("vehicle", e.target.value)}>
              <option value="สำรอง">พนักงานสำรอง (ไม่ประจำรถ)</option>
              {vehiclePlates.map((p) => <option key={p} value={p}>{p}</option>)}
            </select>
          </Field>
          <Field label="ประสบการณ์ (ปี)"><input className={inputClass} type="number" value={form.years} onChange={(e) => update("years", e.target.value)} /></Field>
        </div>

        {error && <div className="text-xs text-rose-600 bg-rose-50 border border-rose-200 rounded-lg p-2.5 mt-2">{error}</div>}

        <div className="flex items-center justify-end gap-3 mt-4 pt-3 border-t border-slate-100">
          <button type="button" onClick={onClose} className="px-4 py-2 rounded-lg border border-slate-300 text-slate-700 text-sm font-semibold hover:bg-slate-50 transition-colors">ยกเลิก</button>
          <button type="submit" disabled={saving} className="px-5 py-2 rounded-lg bg-sky-600 hover:bg-sky-700 text-white text-sm font-semibold transition-colors shadow-sm disabled:opacity-50">
            {saving ? "กำลังบันทึก..." : "บันทึก"}
          </button>
        </div>
      </form>
    </ModalShell>
  );
}

/* ---------------------------------------------------------------
   MAIN DASHBOARD VIEW
---------------------------------------------------------------- */

function Dashboard({ vehicles, repairs }) {
  const readyCount = vehicles.filter((v) => v.status === "ready").length;
  const notReadyCount = vehicles.length - readyCount;

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

  const recentRepairs = [...repairs].sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 5);

  const statusData = [
    { name: "พร้อมใช้งาน", value: readyCount, color: "#10B981" },
    { name: "ไม่พร้อมใช้งาน", value: notReadyCount, color: "#F43F5E" },
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

  const totalSpent = repairs.reduce((s, r) => s + Number(r.cost || 0), 0);

  const kpis = [
    { label: "รถทั้งหมด", value: vehicles.length, unit: "คัน", icon: Truck, color: "text-sky-600", bg: "bg-sky-50" },
    { label: "พร้อมใช้งาน", value: readyCount, unit: "คัน", icon: CheckCircle2, color: "text-emerald-600", bg: "bg-emerald-50" },
    { label: "ไม่พร้อมใช้งาน", value: notReadyCount, unit: "คัน", icon: AlertTriangle, color: "text-rose-600", bg: "bg-rose-50" },
    { label: "ต้องดำเนินการ", value: alerts.length, unit: "รายการ", icon: Clock3, color: "text-amber-600", bg: "bg-amber-50" },
  ];

  return (
    <div className="space-y-6">
      <SectionTitle icon={Gauge} sub="ภาพรวมสถานะฟลีตรถบรรทุกควบคุมอุณหภูมิ ณ วันนี้">
        แดชบอร์ดภาพรวม
      </SectionTitle>

      {/* KPI Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        {kpis.map((k) => (
          <Card key={k.label} className="p-5 flex items-center justify-between">
            <div>
              <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide">{k.label}</p>
              <div className="flex items-baseline gap-1.5 mt-2">
                <span className="text-3xl font-extrabold text-slate-800 tracking-tight">{k.value}</span>
                <span className="text-xs font-medium text-slate-400">{k.unit}</span>
              </div>
            </div>
            <div className={`p-3 rounded-2xl ${k.bg}`}>
              <k.icon className={`w-6 h-6 ${k.color}`} />
            </div>
          </Card>
        ))}
      </div>

      {/* Charts Section */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <Card className="p-5">
          <h3 className="text-sm font-bold text-slate-700 mb-4">สถานะความพร้อมของรถ</h3>
          <ResponsiveContainer width="100%" height={160}>
            <BarChart data={statusData} layout="vertical" margin={{ left: -20, right: 20 }}>
              <XAxis type="number" hide />
              <YAxis type="category" dataKey="name" width={110} tick={{ fill: "#64748B", fontSize: 12 }} axisLine={false} tickLine={false} />
              <Tooltip contentStyle={{ backgroundColor: "#FFFFFF", borderRadius: 8, borderColor: "#E2E8F0", boxShadow: "0 4px 6px -1px rgb(0 0 0 / 0.1)" }} cursor={{ fill: "rgba(0,0,0,0.02)" }} />
              <Bar dataKey="value" radius={[0, 6, 6, 0]} barSize={24}>
                {statusData.map((d, i) => <Cell key={i} fill={d.color} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
          <div className="mt-4 pt-3 border-t border-slate-100 flex justify-between items-center text-xs">
            <span className="text-slate-500">ค่าใช้จ่ายซ่อมบำรุงสะสม:</span>
            <span className="font-bold text-sky-700 text-sm">{fmtMoney(totalSpent)} ฿</span>
          </div>
        </Card>

        <Card className="p-5 lg:col-span-2">
          <h3 className="text-sm font-bold text-slate-700 mb-4">ค่าใช้จ่ายซ่อมบำรุงรายเดือน (บาท)</h3>
          <ResponsiveContainer width="100%" height={180}>
            <BarChart data={monthlyCost}>
              <CartesianGrid strokeDasharray="3 3" stroke="#F1F5F9" vertical={false} />
              <XAxis dataKey="month" tick={{ fill: "#64748B", fontSize: 11 }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fill: "#64748B", fontSize: 11 }} axisLine={false} tickLine={false} />
              <Tooltip formatter={(v) => [`${fmtMoney(v)} บาท`, "ค่าใช้จ่าย"]} contentStyle={{ backgroundColor: "#FFFFFF", borderRadius: 8, borderColor: "#E2E8F0" }} />
              <Bar dataKey="cost" fill="#0284C7" radius={[4, 4, 0, 0]} barSize={32} />
            </BarChart>
          </ResponsiveContainer>
        </Card>
      </div>

      {/* Action Alerts & Recent Activity */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Alerts */}
        <Card className="p-5">
          <div className="flex items-center justify-between mb-4">
            <h3 className="text-sm font-bold text-slate-800 flex items-center gap-2">
              <ShieldAlert size={18} className="text-amber-500" />
              กำหนดการที่ต้องดำเนินการ ({alerts.length})
            </h3>
          </div>
          <div className="space-y-2.5 max-h-[300px] overflow-y-auto pr-1">
            {alerts.length === 0 ? (
              <p className="text-xs text-slate-400 text-center py-8">ไม่มีรายการแจ้งเตือนในขณะนี้</p>
            ) : (
              alerts.map((a, idx) => (
                <div key={idx} className="flex items-center justify-between p-3 rounded-xl bg-slate-50 hover:bg-slate-100/80 transition-colors border border-slate-100">
                  <div className="flex items-center gap-3">
                    <PlateBadge plate={a.plate} />
                    <div>
                      <p className="text-xs font-bold text-slate-800">{a.label}</p>
                      <p className="text-[11px] text-slate-500">{a.brand} {a.model}</p>
                    </div>
                  </div>
                  <DueChip status={a.status} daysLeft={a.daysLeft} />
                </div>
              ))
            )}
          </div>
        </Card>

        {/* Recent Repairs */}
        <Card className="p-5">
          <div className="flex items-center justify-between mb-4">
            <h3 className="text-sm font-bold text-slate-800 flex items-center gap-2">
              <Wrench size={18} className="text-sky-600" />
              ประวัติการซ่อมบำรุงล่าสุด
            </h3>
          </div>
          <div className="space-y-2.5 max-h-[300px] overflow-y-auto pr-1">
            {recentRepairs.length === 0 ? (
              <p className="text-xs text-slate-400 text-center py-8">ไม่มีประวัติการซ่อมบำรุง</p>
            ) : (
              recentRepairs.map((r) => (
                <div key={r.id} className="flex items-center justify-between p-3 rounded-xl bg-slate-50 hover:bg-slate-100/80 transition-colors border border-slate-100">
                  <div className="flex items-center gap-3">
                    <PlateBadge plate={r.plate} />
                    <div>
                      <p className="text-xs font-bold text-slate-800">{r.description}</p>
                      <p className="text-[11px] text-slate-500">{fmtDate(r.date)} • {r.garage}</p>
                    </div>
                  </div>
                  <span className="text-xs font-bold text-slate-700">{fmtMoney(r.cost)} ฿</span>
                </div>
              ))
            )}
          </div>
        </Card>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   VEHICLES LIST VIEW
---------------------------------------------------------------- */

function VehiclesList({ vehicles, onAdd, onEdit, onDelete, onAddRepair }) {
  const [search, setSearch] = useState("");
  const [filterTemp, setFilterTemp] = useState("all");

  const filtered = vehicles.filter((v) => {
    const matchSearch = (v.id + v.brand + v.model + v.driver).toLowerCase().includes(search.toLowerCase());
    const matchTemp = filterTemp === "all" || v.temp === filterTemp;
    return matchSearch && matchTemp;
  });

  return (
    <div className="space-y-6">
      <SectionTitle 
        icon={Truck} 
        sub="จัดการข้อมูลรถบรรทุก กำหนดการ PM และวันหมดอายุเอกสาร"
        action={
          <button onClick={onAdd} className="inline-flex items-center gap-2 px-4 py-2 bg-sky-600 hover:bg-sky-700 text-white font-semibold text-sm rounded-xl transition-all shadow-sm">
            <Plus size={16} /> เพิ่มรถคันใหม่
          </button>
        }
      >
        รายการรถบรรทุกทั้งหมด ({vehicles.length})
      </SectionTitle>

      {/* Filters */}
      <Card className="p-4 flex flex-col sm:flex-row gap-3 items-center justify-between">
        <div className="relative w-full sm:w-72">
          <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
          <input 
            className="w-full bg-slate-50 border border-slate-200 rounded-lg pl-9 pr-3 py-1.5 text-xs text-slate-800 outline-none focus:border-sky-500" 
            placeholder="ค้นหา ทะเบียน, ยี่ห้อ, คนขับ..." 
            value={search} 
            onChange={(e) => setSearch(e.target.value)} 
          />
        </div>
        <div className="flex items-center gap-2 w-full sm:w-auto overflow-x-auto">
          {["all", "freezer", "chiller", "ambient"].map((t) => (
            <button
              key={t}
              onClick={() => setFilterTemp(t)}
              className={`px-3 py-1.5 rounded-lg text-xs font-semibold whitespace-nowrap transition-colors ${
                filterTemp === t 
                  ? 'bg-slate-800 text-white' 
                  : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
              }`}
            >
              {t === "all" ? "ทั้งหมด" : TEMP_META[t].label}
            </button>
          ))}
        </div>
      </Card>

      {/* Table Card */}
      <Card className="overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs text-slate-600">
            <thead className="bg-slate-50 border-b border-slate-200 text-slate-500 uppercase tracking-wider font-semibold">
              <tr>
                <th className="p-4">ทะเบียน / รถ</th>
                <th className="p-4">ประเภทตู้</th>
                <th className="p-4">สถานะ</th>
                <th className="p-4">คนขับ</th>
                <th className="p-4">PM เครื่องยนต์</th>
                <th className="p-4">PM ตู้เย็น</th>
                <th className="p-4">ภาษี/พ.ร.บ.</th>
                <th className="p-4 text-right">จัดการ</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {filtered.map((v) => (
                <tr key={v.id} className="hover:bg-slate-50/80 transition-colors">
                  <td className="p-4">
                    <div className="flex items-center gap-3">
                      <PlateBadge plate={v.id} />
                      <div>
                        <p className="font-bold text-slate-800">{v.brand} {v.model}</p>
                        <p className="text-[11px] text-slate-400">ปี {v.year}</p>
                      </div>
                    </div>
                  </td>
                  <td className="p-4"><TempChip temp={v.temp} /></td>
                  <td className="p-4"><StatusChip status={v.status} reason={v.reason} /></td>
                  <td className="p-4 font-medium text-slate-700">{v.driver}</td>
                  <td className="p-4"><DueChip status={v.pm.status} daysLeft={v.pm.daysLeft} /></td>
                  <td className="p-4"><DueChip status={v.coolingPm.status} daysLeft={v.coolingPm.daysLeft} /></td>
                  <td className="p-4"><DueChip status={v.tax.status} daysLeft={v.tax.daysLeft} /></td>
                  <td className="p-4 text-right">
                    <div className="flex items-center justify-end gap-1">
                      <button onClick={() => onAddRepair(v.id)} title="บันทึกการซ่อม" className="p-1.5 rounded-lg text-slate-400 hover:text-sky-600 hover:bg-sky-50 transition-colors">
                        <Wrench size={15} />
                      </button>
                      <button onClick={() => onEdit(v)} title="แก้ไข" className="p-1.5 rounded-lg text-slate-400 hover:text-amber-600 hover:bg-amber-50 transition-colors">
                        <Pencil size={15} />
                      </button>
                      <button onClick={() => onDelete(v.id)} title="ลบ" className="p-1.5 rounded-lg text-slate-400 hover:text-rose-600 hover:bg-rose-50 transition-colors">
                        <Trash2 size={15} />
                      </button>
                    </div>
                  </td>
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
   REPAIRS LIST VIEW
---------------------------------------------------------------- */

function RepairsList({ repairs, vehicles, onAdd, onEdit, onDelete }) {
  const [search, setSearch] = useState("");

  const filtered = repairs.filter((r) => 
    (r.plate + r.description + r.garage + r.type).toLowerCase().includes(search.toLowerCase())
  );

  return (
    <div className="space-y-6">
      <SectionTitle 
        icon={Wrench} 
        sub="ประวัติค่าใช้จ่ายและรายการซ่อมบำรุงทั้งหมดในระบบ"
        action={
          <button onClick={() => onAdd()} className="inline-flex items-center gap-2 px-4 py-2 bg-sky-600 hover:bg-sky-700 text-white font-semibold text-sm rounded-xl transition-all shadow-sm">
            <Plus size={16} /> บันทึกงานซ่อมใหม่
          </button>
        }
      >
        ประวัติการซ่อมบำรุง ({repairs.length})
      </SectionTitle>

      <Card className="p-4 flex gap-3 items-center">
        <Search size={16} className="text-slate-400" />
        <input 
          className="w-full bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 text-xs text-slate-800 outline-none focus:border-sky-500" 
          placeholder="ค้นหาทะเบียน, รายละเอียด, ประเภทงาน, อู่ซ่อม..." 
          value={search} 
          onChange={(e) => setSearch(e.target.value)} 
        />
      </Card>

      <Card className="overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs text-slate-600">
            <thead className="bg-slate-50 border-b border-slate-200 text-slate-500 uppercase tracking-wider font-semibold">
              <tr>
                <th className="p-4">วันที่</th>
                <th className="p-4">ทะเบียนรถ</th>
                <th className="p-4">ประเภทงาน</th>
                <th className="p-4">รายละเอียด</th>
                <th className="p-4">ศูนย์บริการ / อู่</th>
                <th className="p-4">ค่าใช้จ่าย</th>
                <th className="p-4">สถานะ</th>
                <th className="p-4 text-right">จัดการ</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {filtered.map((r) => (
                <tr key={r.id} className="hover:bg-slate-50/80 transition-colors">
                  <td className="p-4 whitespace-nowrap font-medium text-slate-700">{fmtDate(r.date)}</td>
                  <td className="p-4"><PlateBadge plate={r.plate} /></td>
                  <td className="p-4"><span className="px-2 py-1 rounded bg-slate-100 font-semibold text-slate-700">{r.type}</span></td>
                  <td className="p-4 font-semibold text-slate-800">{r.description}</td>
                  <td className="p-4 text-slate-500">{r.garage}</td>
                  <td className="p-4 font-bold text-sky-700">{fmtMoney(r.cost)} ฿</td>
                  <td className="p-4">
                    <span className={`px-2 py-0.5 rounded-full text-[11px] font-bold ${
                      r.status === 'เสร็จสิ้น' ? 'bg-emerald-50 text-emerald-600' : 'bg-amber-50 text-amber-600'
                    }`}>
                      {r.status}
                    </span>
                  </td>
                  <td className="p-4 text-right">
                    <div className="flex items-center justify-end gap-1">
                      <button onClick={() => onEdit(r)} title="แก้ไข" className="p-1.5 rounded-lg text-slate-400 hover:text-amber-600 hover:bg-amber-50 transition-colors">
                        <Pencil size={15} />
                      </button>
                      <button onClick={() => onDelete(r.id)} title="ลบ" className="p-1.5 rounded-lg text-slate-400 hover:text-rose-600 hover:bg-rose-50 transition-colors">
                        <Trash2 size={15} />
                      </button>
                    </div>
                  </td>
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
   DRIVERS LIST VIEW
---------------------------------------------------------------- */

function DriversList({ drivers, vehicles, onAdd, onEdit, onDelete }) {
  const [search, setSearch] = useState("");

  const filtered = drivers.filter((d) => 
    (d.name + d.license + d.vehicle + d.phone).toLowerCase().includes(search.toLowerCase())
  );

  return (
    <div className="space-y-6">
      <SectionTitle 
        icon={User} 
        sub="จัดการรายชื่อพนักงานขับรถ ใบอนุญาต และการมอบหมายรถประจำ"
        action={
          <button onClick={onAdd} className="inline-flex items-center gap-2 px-4 py-2 bg-sky-600 hover:bg-sky-700 text-white font-semibold text-sm rounded-xl transition-all shadow-sm">
            <Plus size={16} /> เพิ่มพนักงานใหม่
          </button>
        }
      >
        พนักงานขับรถ ({drivers.length})
      </SectionTitle>

      <Card className="p-4 flex gap-3 items-center">
        <Search size={16} className="text-slate-400" />
        <input 
          className="w-full bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 text-xs text-slate-800 outline-none focus:border-sky-500" 
          placeholder="ค้นหาชื่อ, เลขใบขับขี่, เบอร์โทร..." 
          value={search} 
          onChange={(e) => setSearch(e.target.value)} 
        />
      </Card>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
        {filtered.map((d) => (
          <Card key={d.id} className="p-5 flex flex-col justify-between">
            <div>
              <div className="flex items-start justify-between gap-2 mb-3">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 rounded-full bg-sky-50 border border-sky-100 flex items-center justify-center text-sky-600 font-bold">
                    <User size={20} />
                  </div>
                  <div>
                    <h4 className="font-bold text-slate-800 text-sm">{d.name}</h4>
                    <p className="text-xs text-slate-400">ประสบการณ์ {d.years} ปี</p>
                  </div>
                </div>
                <span className="px-2 py-0.5 rounded text-[11px] font-bold bg-slate-100 text-slate-600 border border-slate-200">
                  ใบอนุญาต {d.license_type}
                </span>
              </div>

              <div className="space-y-2 my-4 pt-3 border-t border-slate-100 text-xs">
                <div className="flex justify-between text-slate-600">
                  <span className="text-slate-400">เลขที่ใบขับขี่:</span>
                  <span className="font-mono font-semibold">{d.license}</span>
                </div>
                <div className="flex justify-between text-slate-600">
                  <span className="text-slate-400">วันหมดอายุ:</span>
                  <span>{fmtDate(d.expiry)}</span>
                </div>
                <div className="flex justify-between text-slate-600">
                  <span className="text-slate-400">เบอร์โทรศัพท์:</span>
                  <span className="font-semibold text-slate-700">{d.phone || "-"}</span>
                </div>
                <div className="flex justify-between items-center text-slate-600 pt-1">
                  <span className="text-slate-400">ประจำรถ:</span>
                  {d.vehicle === "สำรอง" ? (
                    <span className="px-2 py-0.5 rounded bg-amber-50 text-amber-700 font-semibold text-[11px]">พนักงานสำรอง</span>
                  ) : (
                    <PlateBadge plate={d.vehicle} />
                  )}
                </div>
              </div>
            </div>

            <div className="flex items-center justify-end gap-2 pt-3 border-t border-slate-100">
              <button onClick={() => onEdit(d)} className="px-3 py-1.5 rounded-lg border border-slate-200 text-slate-600 text-xs font-semibold hover:bg-slate-50 transition-colors">
                แก้ไข
              </button>
              <button onClick={() => onDelete(d.id)} className="px-3 py-1.5 rounded-lg border border-rose-200 text-rose-600 text-xs font-semibold hover:bg-rose-50 transition-colors">
                ลบ
              </button>
            </div>
          </Card>
        ))}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   MAIN APPLICATION CONTAINER (Layout & Navigation)
---------------------------------------------------------------- */

export default function FleetApp() {
  const [tab, setTab] = useState("dashboard");
  const [vehicles, setVehicles] = useState([]);
  const [repairs, setRepairs] = useState([]);
  const [drivers, setDrivers] = useState([]);
  const [loading, setLoading] = useState(true);

  // Modal States
  const [vehicleModal, setVehicleModal] = useState({ open: false, initial: null });
  const [repairModal, setRepairModal] = useState({ open: false, plate: "", initial: null });
  const [driverModal, setDriverModal] = useState({ open: false, initial: null });
  const [deleteConfirm, setDeleteConfirm] = useState({ open: false, label: "", onConfirm: null });

  const loadData = async () => {
    setLoading(true);
    try {
      const [vData, rData, dData] = await Promise.all([
        fetchVehicles(),
        fetchRepairs(),
        fetchDrivers(),
      ]);
      setVehicles(vData.map(computeVehicle));
      setRepairs(rData);
      setDrivers(dData);
    } catch (err) {
      console.error("Error loading data:", err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { loadData(); }, []);

  // Handlers for Vehicle CRUD
  const handleSaveVehicle = async (data) => {
    if (data.oldId && data.oldId !== data.id) {
      await supabase.from("vehicles").delete().eq("id", data.oldId);
    }
    const payload = { ...data };
    delete payload.oldId;
    const { error } = await supabase.from("vehicles").upsert(payload);
    if (error) throw error;
    await loadData();
  };

  const handleDeleteVehicle = (id) => {
    setDeleteConfirm({
      open: true,
      label: `คุณต้องการลบรถทะเบียน ${id} และข้อมูลที่เกี่ยวข้องใช่หรือไม่?`,
      onConfirm: async () => {
        await supabase.from("vehicles").delete().eq("id", id);
        setDeleteConfirm({ open: false });
        await loadData();
      }
    });
  };

  // Handlers for Repair CRUD
  const handleSaveRepair = async (data) => {
    const { error } = await supabase.from("repairs").upsert(data);
    if (error) throw error;
    await loadData();
  };

  const handleDeleteRepair = (id) => {
    setDeleteConfirm({
      open: true,
      label: "คุณต้องการลบรายการซ่อมนี้ใช่หรือไม่?",
      onConfirm: async () => {
        await supabase.from("repairs").delete().eq("id", id);
        setDeleteConfirm({ open: false });
        await loadData();
      }
    });
  };

  // Handlers for Driver CRUD
  const handleSaveDriver = async (data) => {
    const { error } = await supabase.from("drivers").upsert(data);
    if (error) throw error;
    await loadData();
  };

  const handleDeleteDriver = (id) => {
    setDeleteConfirm({
      open: true,
      label: "คุณต้องการลบข้อมูลพนักงานขับรถคนนี้ใช่หรือไม่?",
      onConfirm: async () => {
        await supabase.from("drivers").delete().eq("id", id);
        setDeleteConfirm({ open: false });
        await loadData();
      }
    });
  };

  const navItems = [
    { id: "dashboard", label: "แดชบอร์ด", icon: LayoutDashboard },
    { id: "vehicles", label: "รายการรถบรรทุก", icon: Truck },
    { id: "repairs", label: "การซ่อมบำรุง", icon: Wrench },
    { id: "drivers", label: "พนักงานขับรถ", icon: User },
  ];

  return (
    <div className="min-h-screen bg-slate-100 font-sans text-slate-800 antialiased flex flex-col md:flex-row">
      {/* Sidebar Navigation */}
      <aside className="w-full md:w-64 bg-slate-900 text-slate-300 flex-shrink-0 p-4 flex md:flex-col justify-between border-r border-slate-800">
        <div>
          <div className="flex items-center gap-3 px-3 py-4 mb-6 border-b border-slate-800">
            <div className="p-2 rounded-xl bg-sky-500 text-white">
              <Truck size={22} />
            </div>
            <div>
              <h2 className="font-extrabold text-white tracking-tight text-base leading-none">CPRAM Fleet</h2>
              <span className="text-[10px] text-slate-400 font-medium">Temperature Control</span>
            </div>
          </div>

          <nav className="space-y-1">
            {navItems.map((item) => {
              const Icon = item.icon;
              const active = tab === item.id;
              return (
                <button
                  key={item.id}
                  onClick={() => setTab(item.id)}
                  className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl font-semibold text-xs transition-all ${
                    active 
                      ? 'bg-sky-600 text-white shadow-sm' 
                      : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800/60'
                  }`}
                >
                  <Icon size={18} />
                  {item.label}
                </button>
              );
            })}
          </nav>
        </div>

        <div className="hidden md:block pt-4 border-t border-slate-800">
          <div className="px-3 py-2 text-[11px] text-slate-500">
            <p>ระบบจัดการฟลีตรถบรรทุก</p>
            <p className="mt-0.5">v2.0 Clean Light Mode</p>
          </div>
        </div>
      </aside>

      {/* Main Content Area */}
      <main className="flex-1 p-4 md:p-8 max-w-7xl mx-auto w-full overflow-y-auto">
        {loading ? (
          <div className="flex items-center justify-center min-h-[400px]">
            <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-sky-600"></div>
          </div>
        ) : (
          <>
            {tab === "dashboard" && <Dashboard vehicles={vehicles} repairs={repairs} />}
            {tab === "vehicles" && (
              <VehiclesList 
                vehicles={vehicles} 
                onAdd={() => setVehicleModal({ open: true, initial: null })}
                onEdit={(v) => setVehicleModal({ open: true, initial: v })}
                onDelete={handleDeleteVehicle}
                onAddRepair={(plate) => setRepairModal({ open: true, plate, initial: null })}
              />
            )}
            {tab === "repairs" && (
              <RepairsList 
                repairs={repairs} 
                vehicles={vehicles}
                onAdd={() => setRepairModal({ open: true, plate: vehicles[0]?.id || "", initial: null })}
                onEdit={(r) => setRepairModal({ open: true, plate: r.plate, initial: r })}
                onDelete={handleDeleteRepair}
              />
            )}
            {tab === "drivers" && (
              <DriversList 
                drivers={drivers} 
                vehicles={vehicles}
                onAdd={() => setDriverModal({ open: true, initial: null })}
                onEdit={(d) => setDriverModal({ open: true, initial: d })}
                onDelete={handleDeleteDriver}
              />
            )}
          </>
        )}
      </main>

      {/* Render Modals */}
      {vehicleModal.open && (
        <VehicleFormModal 
          initial={vehicleModal.initial} 
          existingPlates={vehicles.map(v => v.id)}
          onClose={() => setVehicleModal({ open: false, initial: null })}
          onSave={handleSaveVehicle}
        />
      )}

      {repairModal.open && (
        <RepairFormModal 
          plate={repairModal.plate} 
          initial={repairModal.initial}
          onClose={() => setRepairModal({ open: false, plate: "", initial: null })}
          onSave={handleSaveRepair}
        />
      )}

      {driverModal.open && (
        <DriverFormModal 
          initial={driverModal.initial} 
          vehiclePlates={vehicles.map(v => v.id)}
          onClose={() => setDriverModal({ open: false, initial: null })}
          onSave={handleSaveDriver}
        />
      )}

      {deleteConfirm.open && (
        <ConfirmDelete 
          label={deleteConfirm.label} 
          onConfirm={deleteConfirm.onConfirm} 
          onCancel={() => setDeleteConfirm({ open: false })}
        />
      )}
    </div>
  );
}
