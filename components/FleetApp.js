import { useEffect, useMemo, useState } from "react";
import {
  BarChart, Bar, XAxis, YAxis, ResponsiveContainer, Tooltip, CartesianGrid, Cell,
} from "recharts";
import {
  Truck, Snowflake, Droplet, Sun, Wrench, Calendar, AlertTriangle, CheckCircle2,
  User, Phone, Search, X, Clock3, CreditCard, ChevronRight, ClipboardList, Gauge,
  Plus, Trash2, Pencil, LogOut,
} from "lucide-react";
import { supabase } from "../lib/supabaseClient";

/* ---------------------------------------------------------------
   CONFIG
---------------------------------------------------------------- */

const TODAY = new Date(new Date().toDateString());

const TEMP_META = {
  freezer: { label: "ห้องเย็น -18°C", icon: Snowflake, color: "#45B8C8" },
  chiller: { label: "แช่เย็น 0-4°C", icon: Droplet, color: "#5B9FE0" },
  ambient: { label: "อุณหภูมิห้อง", icon: Sun, color: "#E0A85B" },
};
const REPAIR_TYPES = ["เครื่องยนต์", "ระบบทำความเย็น", "ระบบไฟฟ้า", "เบรก", "ยาง", "ตัวถัง", "อื่นๆ"];
const DRIVER_LICENSE_TYPES = ["ท.2", "ท.3", "ท.4"];
const THAI_MONTHS = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];

/* ---------------------------------------------------------------
   HELPERS
---------------------------------------------------------------- */

function daysBetween(a, b) { return Math.round((b - a) / (1000 * 60 * 60 * 24)); }
function addDays(dateStr, n) { const d = new Date(dateStr); d.setDate(d.getDate() + n); return d; }
function fmtDate(d) { const date = typeof d === "string" ? new Date(d) : d; return `${date.getDate()} ${THAI_MONTHS[date.getMonth()]} ${date.getFullYear() + 543}`; }
function fmtMoney(n) { return Number(n || 0).toLocaleString("th-TH"); }
function computeVehicle(v) {
  const nextPM = addDays(v.last_pm, v.pm_interval);
  const daysLeft = daysBetween(TODAY, nextPM);
  let pmStatus = "ok";
  if (daysLeft < 0) pmStatus = "overdue";
  else if (daysLeft <= 7) pmStatus = "soon";
  return { ...v, nextPM, daysLeft, pmStatus };
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
    <span className="inline-flex items-center gap-1 rounded-full px-2 py-1 text-xs font-semibold" style={{ background: ready ? "rgba(95,190,132,0.15)" : "rgba(228,88,79,0.15)", color: ready ? "#5FBE84" : "#E4584F", border: `1px solid ${ready ? "#5FBE8455" : "#E4584F55"}` }} title={reason || ""}>
      {ready ? <CheckCircle2 size={13} /> : <AlertTriangle size={13} />}
      {ready ? "พร้อมใช้งาน" : "ไม่พร้อมใช้งาน"}
    </span>
  );
}

function PMChip({ pmStatus, daysLeft }) {
  const conf = {
    overdue: { color: "#E4584F", label: `เกินกำหนด ${Math.abs(daysLeft)} วัน` },
    soon: { color: "#F0A94E", label: `ใกล้ครบกำหนด (${daysLeft} วัน)` },
    ok: { color: "#8FA0B3", label: `อีก ${daysLeft} วัน` },
  }[pmStatus];
  return (
    <span className="inline-flex items-center gap-1 rounded-full px-2 py-1 text-xs font-semibold" style={{ background: `${conf.color}1E`, color: conf.color, border: `1px solid ${conf.color}55` }}>
      <Clock3 size={13} />{conf.label}
    </span>
  );
}

function Card({ children, style }) {
  return <div style={{ background: "var(--surface)", border: "1px solid var(--border)", borderRadius: 14, ...style }}>{children}</div>;
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

function ModalShell({ title, onClose, children, width = 520 }) {
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

function ConfirmDelete({ label, onConfirm, onCancel }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(8,10,14,0.7)", zIndex: 60, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }} onClick={onCancel}>
      <div onClick={(e) => e.stopPropagation()} style={{ maxWidth: 380, width: "100%" }}>
        <Card style={{ padding: 20 }}>
          <div className="flex items-center gap-2 mb-3"><AlertTriangle size={18} style={{ color: "#E4584F" }} /><span style={{ fontSize: 15, fontWeight: 700, color: "var(--text)" }}>ยืนยันการลบ</span></div>
          <p style={{ fontSize: 13, color: "var(--text-muted)", marginBottom: 18 }}>{label}</p>
          <div className="flex items-center justify-end gap-2">
            <button onClick={onCancel} style={{ ...inputStyle, width: "auto", padding: "8px 16px", cursor: "pointer" }}>ยกเลิก</button>
            <button onClick={onConfirm} style={{ background: "#E4584F", color: "#fff", border: "none", borderRadius: 8, padding: "8px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>ลบ</button>
          </div>
        </Card>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   VEHICLE FORM (add + edit)
---------------------------------------------------------------- */

function VehicleFormModal({ initial, onClose, onSave, existingPlates }) {
  const isEdit = !!initial;
  const [form, setForm] = useState(() => initial ? { ...initial, year: String(initial.year || ""), pm_interval: String(initial.pm_interval) } : {
    id: "", brand: "", model: "", year: "2024", temp: "chiller",
    status: "ready", reason: "", last_pm: new Date().toISOString().slice(0, 10), pm_interval: "90", driver: "",
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
    if (!form.last_pm) return setError("กรุณาเลือกวันที่ PM ล่าสุด");

    setSaving(true);
    setError("");
    try {
      await onSave({
        oldId: isEdit ? initial.id : null,
        id: plate, brand: form.brand.trim(), model: form.model.trim(),
        year: Number(form.year) || new Date().getFullYear(), temp: form.temp, status: form.status,
        reason: form.status === "not_ready" ? form.reason.trim() : "",
        last_pm: form.last_pm, pm_interval: Number(form.pm_interval) || 90, driver: form.driver.trim() || "-",
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
          <Field label="ประเภทตู้บรรทุก">
            <select style={inputStyle} value={form.temp} onChange={(e) => update("temp", e.target.value)}>
              <option value="freezer">ห้องเย็น -18°C</option>
              <option value="chiller">แช่เย็น 0-4°C</option>
              <option value="ambient">อุณหภูมิห้อง</option>
            </select>
          </Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="สถานะรถ">
            <select style={inputStyle} value={form.status} onChange={(e) => update("status", e.target.value)}>
              <option value="ready">พร้อมใช้งาน</option>
              <option value="not_ready">ไม่พร้อมใช้งาน</option>
            </select>
          </Field>
          {form.status === "not_ready" && (
            <Field label="เหตุผลที่ไม่พร้อมใช้งาน"><input style={inputStyle} placeholder="เช่น รอซ่อมเครื่องยนต์" value={form.reason} onChange={(e) => update("reason", e.target.value)} /></Field>
          )}
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="วันที่ PM ล่าสุด *"><input style={inputStyle} type="date" value={form.last_pm} onChange={(e) => update("last_pm", e.target.value)} /></Field>
          <Field label="รอบ PM (วัน)"><input style={inputStyle} type="number" value={form.pm_interval} onChange={(e) => update("pm_interval", e.target.value)} /></Field>
        </div>
        {error && <div style={{ fontSize: 12, color: "#E4584F", background: "rgba(228,88,79,0.1)", border: "1px solid #E4584F55", borderRadius: 8, padding: "8px 10px" }}>{error}</div>}
        <div className="flex items-center justify-end gap-2 mt-2">
          <button type="button" onClick={onClose} style={{ ...inputStyle, width: "auto", padding: "8px 16px", cursor: "pointer" }}>ยกเลิก</button>
          <button type="submit" disabled={saving} style={{ background: "var(--accent-frost)", color: "#0F1620", border: "none", borderRadius: 8, padding: "8px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer", opacity: saving ? 0.6 : 1 }}>
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
  const [form, setForm] = useState(() => initial ? { ...initial, cost: String(initial.cost) } : {
    date: new Date().toISOString().slice(0, 10), type: REPAIR_TYPES[0], description: "", cost: "", garage: "", status: "เสร็จสิ้น",
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
        <Field label="รายละเอียดงานซ่อม *"><input style={inputStyle} placeholder="เช่น เปลี่ยนผ้าเบรกหน้า" value={form.description} onChange={(e) => update("description", e.target.value)} /></Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="ค่าใช้จ่าย (บาท)"><input style={inputStyle} type="number" value={form.cost} onChange={(e) => update("cost", e.target.value)} /></Field>
          <Field label="สถานะงาน">
            <select style={inputStyle} value={form.status} onChange={(e) => update("status", e.target.value)}>
              <option value="เสร็จสิ้น">เสร็จสิ้น</option>
              <option value="กำลังดำเนินการ">กำลังดำเนินการ</option>
            </select>
          </Field>
        </div>
        <Field label="อู่ / ศูนย์บริการ"><input style={inputStyle} placeholder="เช่น อู่กลาง CPRAM" value={form.garage} onChange={(e) => update("garage", e.target.value)} /></Field>
        {error && <div style={{ fontSize: 12, color: "#E4584F", background: "rgba(228,88,79,0.1)", border: "1px solid #E4584F55", borderRadius: 8, padding: "8px 10px" }}>{error}</div>}
        <div className="flex items-center justify-end gap-2 mt-2">
          <button type="button" onClick={onClose} style={{ ...inputStyle, width: "auto", padding: "8px 16px", cursor: "pointer" }}>ยกเลิก</button>
          <button type="submit" disabled={saving} style={{ background: "var(--accent-frost)", color: "#0F1620", border: "none", borderRadius: 8, padding: "8px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer", opacity: saving ? 0.6 : 1 }}>
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
        {error && <div style={{ fontSize: 12, color: "#E4584F", background: "rgba(228,88,79,0.1)", border: "1px solid #E4584F55", borderRadius: 8, padding: "8px 10px" }}>{error}</div>}
        <div className="flex items-center justify-end gap-2 mt-2">
          <button type="button" onClick={onClose} style={{ ...inputStyle, width: "auto", padding: "8px 16px", cursor: "pointer" }}>ยกเลิก</button>
          <button type="submit" disabled={saving} style={{ background: "var(--accent-frost)", color: "#0F1620", border: "none", borderRadius: 8, padding: "8px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer", opacity: saving ? 0.6 : 1 }}>
            {saving ? "กำลังบันทึก..." : "บันทึก"}
          </button>
        </div>
      </form>
    </ModalShell>
  );
}

/* ---------------------------------------------------------------
   VIEWS
---------------------------------------------------------------- */

function Dashboard({ vehicles, repairs }) {
  const readyCount = vehicles.filter((v) => v.status === "ready").length;
  const notReadyCount = vehicles.length - readyCount;
  const pmSoon = vehicles.filter((v) => v.pmStatus !== "ok").sort((a, b) => a.daysLeft - b.daysLeft);
  const recentRepairs = [...repairs].sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 6);

  const statusData = [
    { name: "พร้อมใช้งาน", value: readyCount, color: "#5FBE84" },
    { name: "ไม่พร้อมใช้งาน", value: notReadyCount, color: "#E4584F" },
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

  const kpis = [
    { label: "รถทั้งหมด", value: vehicles.length, icon: Truck, color: "#45B8C8" },
    { label: "พร้อมใช้งาน", value: readyCount, icon: CheckCircle2, color: "#5FBE84" },
    { label: "ไม่พร้อมใช้งาน", value: notReadyCount, icon: AlertTriangle, color: "#E4584F" },
    { label: "ใกล้/เกินกำหนด PM", value: pmSoon.length, icon: Clock3, color: "#F0A94E" },
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
          <ResponsiveContainer width="100%" height={180}>
            <BarChart data={statusData} layout="vertical" margin={{ left: 0, right: 20 }}>
              <XAxis type="number" hide />
              <YAxis type="category" dataKey="name" width={110} tick={{ fill: "#8FA0B3", fontSize: 12 }} axisLine={false} tickLine={false} />
              <Tooltip contentStyle={{ background: "#1A222C", border: "1px solid #2B3644", borderRadius: 8, fontSize: 12 }} labelStyle={{ color: "#E8EDF3" }} cursor={{ fill: "rgba(255,255,255,0.03)" }} />
              <Bar dataKey="value" radius={[0, 6, 6, 0]} barSize={28}>{statusData.map((d, i) => <Cell key={i} fill={d.color} />)}</Bar>
            </BarChart>
          </ResponsiveContainer>
        </Card>

        <Card style={{ padding: 20, gridColumn: "span 2 / span 2" }}>
          <div style={{ fontSize: 13, color: "var(--text-muted)", marginBottom: 12, fontWeight: 600 }}>ค่าใช้จ่ายซ่อมบำรุงรายเดือน (บาท)</div>
          <ResponsiveContainer width="100%" height={180}>
            <BarChart data={monthlyCost}>
              <CartesianGrid strokeDasharray="3 3" stroke="#2B3644" vertical={false} />
              <XAxis dataKey="month" tick={{ fill: "#8FA0B3", fontSize: 12 }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fill: "#8FA0B3", fontSize: 11 }} axisLine={false} tickLine={false} />
              <Tooltip formatter={(v) => `${fmtMoney(v)} บาท`} contentStyle={{ background: "#1A222C", border: "1px solid #2B3644", borderRadius: 8, fontSize: 12 }} labelStyle={{ color: "#E8EDF3" }} cursor={{ fill: "rgba(255,255,255,0.03)" }} />
              <Bar dataKey="cost" radius={[6, 6, 0, 0]} fill="#45B8C8" barSize={28} />
            </BarChart>
          </ResponsiveContainer>
        </Card>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <Card style={{ padding: 20 }}>
          <div className="flex items-center gap-2 mb-3"><AlertTriangle size={16} style={{ color: "#F0A94E" }} /><span style={{ fontSize: 14, fontWeight: 600, color: "var(--text)" }}>รถที่ใกล้/เกินกำหนด PM</span></div>
          <div className="flex flex-col gap-2">
            {pmSoon.length === 0 && <p style={{ color: "var(--text-muted)", fontSize: 13 }}>ไม่มีรถที่ใกล้ครบกำหนด PM</p>}
            {pmSoon.map((v) => (
              <div key={v.id} className="flex items-center justify-between rounded-lg px-3 py-2" style={{ background: "var(--surface-2)" }}>
                <div className="flex items-center gap-3"><PlateBadge plate={v.id} /><span style={{ fontSize: 12, color: "var(--text-muted)" }}>{v.brand} {v.model}</span></div>
                <PMChip pmStatus={v.pmStatus} daysLeft={v.daysLeft} />
              </div>
            ))}
          </div>
        </Card>

        <Card style={{ padding: 20 }}>
          <div className="flex items-center gap-2 mb-3"><Wrench size={16} style={{ color: "#45B8C8" }} /><span style={{ fontSize: 14, fontWeight: 600, color: "var(--text)" }}>ประวัติการซ่อมล่าสุด</span></div>
          <div className="flex flex-col gap-2">
            {recentRepairs.length === 0 && <p style={{ color: "var(--text-muted)", fontSize: 13 }}>ยังไม่มีประวัติการซ่อม</p>}
            {recentRepairs.map((r) => (
              <div key={r.id} className="flex items-center justify-between rounded-lg px-3 py-2" style={{ background: "var(--surface-2)" }}>
                <div><div style={{ fontSize: 13, color: "var(--text)", fontWeight: 500 }}>{r.plate} · {r.type}</div><div style={{ fontSize: 12, color: "var(--text-muted)" }}>{r.description}</div></div>
                <div style={{ textAlign: "right" }}>
                  <div style={{ fontSize: 12, color: "var(--text-muted)", fontFamily: "'JetBrains Mono', monospace" }}>{fmtDate(r.date)}</div>
                  <div style={{ fontSize: 12, color: "var(--accent-frost)", fontWeight: 600 }}>{fmtMoney(r.cost)} บ.</div>
                </div>
              </div>
            ))}
          </div>
        </Card>
      </div>
    </div>
  );
}

function VehiclesView({ vehicles, repairs, onAdd, onUpdate, onDelete, onAddRepair, onUpdateRepair, onDeleteRepair }) {
  const [query, setQuery] = useState("");
  const [selected, setSelected] = useState(null);
  const [formState, setFormState] = useState(null);
  const [deleteTarget, setDeleteTarget] = useState(null);
  const [repairForm, setRepairForm] = useState(null);
  const [deleteRepairTarget, setDeleteRepairTarget] = useState(null);

  const filtered = vehicles.filter((v) => v.id.includes(query) || v.brand.toLowerCase().includes(query.toLowerCase()) || v.model.toLowerCase().includes(query.toLowerCase()));
  const selVehicle = vehicles.find((v) => v.id === selected);
  const selRepairs = selVehicle ? repairs.filter((r) => r.plate === selVehicle.id).sort((a, b) => new Date(b.date) - new Date(a.date)) : [];

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
        <SectionTitle icon={Truck} sub="รายละเอียดรถบรรทุกทุกคัน แก้ไข เพิ่ม ลบ พร้อมประวัติการซ่อมรายทะเบียน">รถทั้งหมด</SectionTitle>
        <button onClick={() => setFormState("add")} className="flex items-center gap-2 rounded-lg px-4 py-2" style={{ background: "var(--accent-frost)", color: "#0F1620", fontSize: 13, fontWeight: 700, height: 38 }}>
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
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 820 }}>
            <thead>
              <tr style={{ background: "var(--surface-2)" }}>
                {["ทะเบียน", "ยี่ห้อ / รุ่น", "ประเภทตู้", "สถานะรถ", "กำหนด PM ถัดไป", "คนขับ", "จัดการ", ""].map((h) => (
                  <th key={h} style={{ textAlign: "left", padding: "10px 14px", fontSize: 12, color: "var(--text-muted)", fontWeight: 600 }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map((v) => (
                <tr key={v.id} style={{ borderTop: "1px solid var(--border)", background: selected === v.id ? "rgba(69,184,200,0.08)" : "transparent" }}>
                  <td style={{ padding: "10px 14px", cursor: "pointer" }} onClick={() => setSelected(v.id === selected ? null : v.id)}><PlateBadge plate={v.id} /></td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: "var(--text)", cursor: "pointer" }} onClick={() => setSelected(v.id === selected ? null : v.id)}>{v.brand} {v.model} <span style={{ color: "var(--text-muted)" }}>'{String(v.year).slice(2)}</span></td>
                  <td style={{ padding: "10px 14px", cursor: "pointer" }} onClick={() => setSelected(v.id === selected ? null : v.id)}><TempChip temp={v.temp} /></td>
                  <td style={{ padding: "10px 14px", cursor: "pointer" }} onClick={() => setSelected(v.id === selected ? null : v.id)}><StatusChip status={v.status} reason={v.reason} /></td>
                  <td style={{ padding: "10px 14px", cursor: "pointer" }} onClick={() => setSelected(v.id === selected ? null : v.id)}><PMChip pmStatus={v.pmStatus} daysLeft={v.daysLeft} /></td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: "var(--text-muted)", cursor: "pointer" }} onClick={() => setSelected(v.id === selected ? null : v.id)}>{v.driver}</td>
                  <td style={{ padding: "10px 14px" }}>
                    <div className="flex items-center gap-1">
                      <button onClick={() => setFormState(v)} title="แก้ไข" style={iconBtnStyle}><Pencil size={15} /></button>
                      <button onClick={() => setDeleteTarget(v)} title="ลบ" style={{ ...iconBtnStyle, color: "#E4584F" }}><Trash2 size={15} /></button>
                    </div>
                  </td>
                  <td style={{ padding: "10px 14px", cursor: "pointer" }} onClick={() => setSelected(v.id === selected ? null : v.id)}>
                    <ChevronRight size={16} style={{ color: "var(--text-muted)", transform: selected === v.id ? "rotate(90deg)" : "none", transition: "transform .15s" }} />
                  </td>
                </tr>
              ))}
              {filtered.length === 0 && (
                <tr><td colSpan={8} style={{ padding: "24px 14px", textAlign: "center", color: "var(--text-muted)", fontSize: 13 }}>ไม่พบรถที่ตรงกับคำค้นหา</td></tr>
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
            <div className="flex items-center gap-2 rounded-lg px-3 py-2 mb-5" style={{ background: "rgba(228,88,79,0.1)", border: "1px solid #E4584F55" }}>
              <AlertTriangle size={14} style={{ color: "#E4584F" }} /><span style={{ fontSize: 13, color: "#E4584F" }}>{selVehicle.reason}</span>
            </div>
          )}

          <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-6">
            <div className="rounded-lg px-3 py-3" style={{ background: "var(--surface-2)" }}><div style={{ fontSize: 11, color: "var(--text-muted)" }}>PM ล่าสุด</div><div style={{ fontSize: 13, color: "var(--text)", fontFamily: "'JetBrains Mono', monospace", marginTop: 2 }}>{fmtDate(selVehicle.last_pm)}</div></div>
            <div className="rounded-lg px-3 py-3" style={{ background: "var(--surface-2)" }}><div style={{ fontSize: 11, color: "var(--text-muted)" }}>PM ถัดไป</div><div style={{ fontSize: 13, color: "var(--text)", fontFamily: "'JetBrains Mono', monospace", marginTop: 2 }}>{fmtDate(selVehicle.nextPM)}</div></div>
            <div className="rounded-lg px-3 py-3" style={{ background: "var(--surface-2)" }}><div style={{ fontSize: 11, color: "var(--text-muted)" }}>รอบ PM</div><div style={{ fontSize: 13, color: "var(--text)", marginTop: 2 }}>ทุก {selVehicle.pm_interval} วัน</div></div>
            <div className="rounded-lg px-3 py-3" style={{ background: "var(--surface-2)" }}><div style={{ fontSize: 11, color: "var(--text-muted)" }}>คนขับประจำ</div><div style={{ fontSize: 13, color: "var(--text)", marginTop: 2 }}>{selVehicle.driver}</div></div>
          </div>

          <div className="flex items-center justify-between mb-3">
            <div className="flex items-center gap-2"><ClipboardList size={15} style={{ color: "var(--accent-frost)" }} /><span style={{ fontSize: 14, fontWeight: 600, color: "var(--text)" }}>ประวัติการซ่อม ({selRepairs.length} รายการ)</span></div>
            <button onClick={() => setRepairForm("add")} className="flex items-center gap-1 rounded-lg px-3 py-1.5" style={{ background: "var(--surface-2)", border: "1px solid var(--border)", color: "var(--accent-frost)", fontSize: 12, fontWeight: 600 }}>
              <Plus size={14} />เพิ่มรายการซ่อม
            </button>
          </div>

          {Object.keys(groupedByMonth).length === 0 && <p style={{ fontSize: 13, color: "var(--text-muted)" }}>ยังไม่มีประวัติการซ่อมสำหรับคันนี้</p>}

          <div className="flex flex-col gap-5">
            {Object.entries(groupedByMonth).map(([month, items]) => (
              <div key={month}>
                <div style={{ fontSize: 12, fontWeight: 700, color: "var(--accent-frost)", marginBottom: 8, textTransform: "uppercase", letterSpacing: 0.5 }}>{month} · รวม {fmtMoney(items.reduce((s, i) => s + i.cost, 0))} บาท</div>
                <div className="flex flex-col gap-2">
                  {items.map((r) => (
                    <div key={r.id} className="flex items-center justify-between rounded-lg px-3 py-2" style={{ background: "var(--surface-2)" }}>
                      <div className="flex items-center gap-3">
                        <span className="rounded-full px-2 py-1 text-xs font-medium" style={{ background: "rgba(69,184,200,0.12)", color: "var(--accent-frost)", whiteSpace: "nowrap" }}>{r.type}</span>
                        <div><div style={{ fontSize: 13, color: "var(--text)" }}>{r.description}</div><div style={{ fontSize: 11, color: "var(--text-muted)" }}>{r.garage} · {r.status}</div></div>
                      </div>
                      <div className="flex items-center gap-3">
                        <div style={{ textAlign: "right" }}>
                          <div style={{ fontSize: 12, color: "var(--text-muted)", fontFamily: "'JetBrains Mono', monospace" }}>{fmtDate(r.date)}</div>
                          <div style={{ fontSize: 13, color: "var(--text)", fontWeight: 600 }}>{fmtMoney(r.cost)} บ.</div>
                        </div>
                        <div className="flex items-center gap-1">
                          <button onClick={() => setRepairForm(r)} title="แก้ไข" style={iconBtnStyle}><Pencil size={14} /></button>
                          <button onClick={() => setDeleteRepairTarget(r)} title="ลบ" style={{ ...iconBtnStyle, color: "#E4584F" }}><Trash2 size={14} /></button>
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

function PMScheduleView({ vehicles }) {
  const sorted = [...vehicles].sort((a, b) => a.daysLeft - b.daysLeft);
  return (
    <div>
      <SectionTitle icon={Calendar} sub="กำหนดการบำรุงรักษาตามระยะ (PM) ของรถทุกคัน เรียงตามวันที่ใกล้ครบกำหนดที่สุด">กำหนดการ PM</SectionTitle>
      <div className="flex items-center gap-4 mb-4" style={{ fontSize: 12, color: "var(--text-muted)" }}>
        <span className="flex items-center gap-1"><span style={{ width: 8, height: 8, borderRadius: 99, background: "#E4584F", display: "inline-block" }} /> เกินกำหนด</span>
        <span className="flex items-center gap-1"><span style={{ width: 8, height: 8, borderRadius: 99, background: "#F0A94E", display: "inline-block" }} /> ใกล้ครบกำหนด (≤ 7 วัน)</span>
        <span className="flex items-center gap-1"><span style={{ width: 8, height: 8, borderRadius: 99, background: "#8FA0B3", display: "inline-block" }} /> ปกติ</span>
      </div>
      <Card style={{ overflow: "hidden" }}>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 780 }}>
            <thead>
              <tr style={{ background: "var(--surface-2)" }}>
                {["ทะเบียน", "ยี่ห้อ / รุ่น", "PM ล่าสุด", "รอบ PM", "PM ถัดไป", "สถานะ"].map((h) => (
                  <th key={h} style={{ textAlign: "left", padding: "10px 14px", fontSize: 12, color: "var(--text-muted)", fontWeight: 600 }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {sorted.map((v) => (
                <tr key={v.id} style={{ borderTop: "1px solid var(--border)" }}>
                  <td style={{ padding: "10px 14px" }}><PlateBadge plate={v.id} /></td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: "var(--text)" }}>{v.brand} {v.model}</td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: "var(--text-muted)", fontFamily: "'JetBrains Mono', monospace" }}>{fmtDate(v.last_pm)}</td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: "var(--text-muted)" }}>ทุก {v.pm_interval} วัน</td>
                  <td style={{ padding: "10px 14px", fontSize: 13, color: "var(--text)", fontFamily: "'JetBrains Mono', monospace" }}>{fmtDate(v.nextPM)}</td>
                  <td style={{ padding: "10px 14px" }}><PMChip pmStatus={v.pmStatus} daysLeft={v.daysLeft} /></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

function DriversView({ drivers, vehiclePlates, onAdd, onUpdate, onDelete }) {
  const [formState, setFormState] = useState(null);
  const [deleteTarget, setDeleteTarget] = useState(null);

  return (
    <div>
      <div className="flex items-center justify-between flex-wrap gap-3 mb-4">
        <SectionTitle icon={User} sub="ข้อมูลพนักงานขับรถ ใบขับขี่ และรถที่รับผิดชอบ">พนักงานขับรถ</SectionTitle>
        <button onClick={() => setFormState("add")} className="flex items-center gap-2 rounded-lg px-4 py-2" style={{ background: "var(--accent-frost)", color: "#0F1620", fontSize: 13, fontWeight: 700, height: 38 }}>
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
                  <div className="flex items-center justify-center rounded-full" style={{ width: 42, height: 42, background: "rgba(69,184,200,0.15)", color: "var(--accent-frost)" }}><User size={20} /></div>
                  <div><div style={{ fontSize: 14, fontWeight: 600, color: "var(--text)" }}>{d.name}</div><div style={{ fontSize: 12, color: "var(--text-muted)" }}>{d.years} ปีประสบการณ์</div></div>
                </div>
                <div className="flex items-center gap-1">
                  <button onClick={() => setFormState(d)} title="แก้ไข" style={iconBtnStyle}><Pencil size={14} /></button>
                  <button onClick={() => setDeleteTarget(d)} title="ลบ" style={{ ...iconBtnStyle, color: "#E4584F" }}><Trash2 size={14} /></button>
                </div>
              </div>
              <div className="flex flex-col gap-2 mb-4">
                <div className="flex items-center gap-2" style={{ fontSize: 13, color: "var(--text-muted)" }}><CreditCard size={14} /><span>ใบขับขี่ {d.license_type} · {d.license}</span></div>
                <div className="flex items-center gap-2" style={{ fontSize: 13, color: "var(--text-muted)" }}><Phone size={14} /><span>{d.phone || "-"}</span></div>
              </div>
              <div className="flex items-center justify-between rounded-lg px-3 py-2 mb-3" style={{ background: "var(--surface-2)" }}>
                <span style={{ fontSize: 12, color: "var(--text-muted)" }}>วันหมดอายุใบขับขี่</span>
                <span style={{ fontSize: 12, fontWeight: 600, color: expSoon ? "#F0A94E" : "var(--text)", fontFamily: "'JetBrains Mono', monospace" }}>{fmtDate(d.expiry)}</span>
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
  { key: "vehicles", label: "รถทั้งหมด", icon: Truck },
  { key: "pm", label: "กำหนดการ PM", icon: Calendar },
  { key: "drivers", label: "พนักงานขับรถ", icon: User },
];

export default function FleetApp({ user }) {
  const [tab, setTab] = useState("dashboard");
  const [vehicles, setVehicles] = useState([]);
  const [repairs, setRepairs] = useState([]);
  const [drivers, setDrivers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [errorMsg, setErrorMsg] = useState("");

  useEffect(() => { loadAll(); }, []);

  async function loadAll() {
    setLoading(true);
    try {
      const [v, r, d] = await Promise.all([fetchVehicles(), fetchRepairs(), fetchDrivers()]);
      setVehicles(v);
      setRepairs(r);
      setDrivers(d);
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

  async function handleSignOut() {
    await supabase.auth.signOut();
    window.location.href = "/login";
  }

  return (
    <div style={{
      "--bg": "#10151C", "--surface": "#1A222C", "--surface-2": "#212B37", "--border": "#2B3644",
      "--text": "#E8EDF3", "--text-muted": "#8FA0B3", "--accent-frost": "#45B8C8",
      background: "var(--bg)", minHeight: "100vh", fontFamily: "'Inter', sans-serif",
    }}>
      <header style={{ borderBottom: "1px solid var(--border)", background: "rgba(16,21,28,0.9)", position: "sticky", top: 0, zIndex: 10, backdropFilter: "blur(6px)" }}>
        <div className="max-w-7xl mx-auto px-5 py-4 flex items-center justify-between flex-wrap gap-3">
          <div className="flex items-center gap-3">
            <div className="flex items-center justify-center rounded-lg" style={{ width: 38, height: 38, background: "rgba(69,184,200,0.15)" }}><Snowflake size={20} style={{ color: "var(--accent-frost)" }} /></div>
            <div>
              <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 17, color: "var(--text)", letterSpacing: 0.3 }}>CPRAM FLEET</div>
              <div style={{ fontSize: 11, color: "var(--text-muted)" }}>ระบบบริหารจัดการรถขนส่งควบคุมอุณหภูมิ</div>
            </div>
          </div>
          <nav className="flex items-center gap-1 rounded-lg p-1" style={{ background: "var(--surface-2)" }}>
            {TABS.map((t) => {
              const active = tab === t.key;
              return (
                <button key={t.key} onClick={() => setTab(t.key)} className="flex items-center gap-2 rounded-md px-3 py-2 text-sm font-medium" style={{ color: active ? "#0F1620" : "var(--text-muted)", background: active ? "var(--accent-frost)" : "transparent", transition: "all .15s" }}>
                  <t.icon size={15} /><span className="hidden sm:inline">{t.label}</span>
                </button>
              );
            })}
          </nav>
          <div className="flex items-center gap-3">
            <span style={{ fontSize: 12, color: "var(--text-muted)" }} className="hidden md:inline">{user?.email}</span>
            <button onClick={handleSignOut} className="flex items-center gap-1 rounded-lg px-3 py-2" style={{ background: "var(--surface-2)", border: "1px solid var(--border)", color: "var(--text-muted)", fontSize: 12 }}>
              <LogOut size={14} />ออกจากระบบ
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-5 py-6">
        {errorMsg && (
          <div style={{ marginBottom: 16, fontSize: 12, color: "#F0A94E", background: "rgba(240,169,78,0.1)", border: "1px solid #F0A94E55", borderRadius: 8, padding: "8px 12px" }}>{errorMsg}</div>
        )}
        {loading ? (
          <div style={{ color: "var(--text-muted)", fontSize: 14, padding: "40px 0", textAlign: "center" }}>กำลังโหลดข้อมูล...</div>
        ) : (
          <>
            {tab === "dashboard" && <Dashboard vehicles={computedVehicles} repairs={repairs} />}
            {tab === "vehicles" && (
              <VehiclesView
                vehicles={computedVehicles} repairs={repairs}
                onAdd={handleAddVehicle} onUpdate={handleUpdateVehicle} onDelete={handleDeleteVehicle}
                onAddRepair={handleAddRepair} onUpdateRepair={handleUpdateRepair} onDeleteRepair={handleDeleteRepair}
              />
            )}
            {tab === "pm" && <PMScheduleView vehicles={computedVehicles} />}
            {tab === "drivers" && (
              <DriversView drivers={drivers} vehiclePlates={vehicles.map((v) => v.id)} onAdd={handleAddDriver} onUpdate={handleUpdateDriver} onDelete={handleDeleteDriver} />
            )}
          </>
        )}
      </main>

      <footer className="max-w-7xl mx-auto px-5 py-6" style={{ color: "var(--text-muted)", fontSize: 12 }}>
        ข้อมูลบันทึกลงฐานข้อมูลจริง ใช้งานพร้อมกันได้หลายคน · CPRAM Fleet Management
      </footer>
    </div>
  );
}
