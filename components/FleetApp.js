import { useEffect, useMemo, useState } from "react";
import {
  BarChart, Bar, XAxis, YAxis, ResponsiveContainer, Tooltip, CartesianGrid, Cell,
  AreaChart, Area
} from "recharts";
import {
  Truck, Snowflake, Droplet, Sun, Wrench, Calendar, AlertTriangle, CheckCircle2,
  User, Phone, Search, X, Clock3, CreditCard, ChevronRight, ClipboardList, Gauge,
  Plus, Trash2, Pencil, LogOut, Settings2, FileText, Zap, ShieldCheck, Users, Building2
} from "lucide-react";
import { supabase } from "../lib/supabaseClient";

/* ---------------------------------------------------------------
   CONFIG & CONSTANTS
---------------------------------------------------------------- */
const TODAY = new Date(new Date().toDateString());

const TEMP_META = {
  freezer: { label: "ห้องเย็น -18°C", icon: Snowflake, color: "#38BDF8" },
  chiller: { label: "แช่เย็น 0-4°C", icon: Droplet, color: "#2DD4BF" },
  ambient: { label: "อุณหภูมิห้อง", icon: Sun, color: "#F59E0B" },
};

const REPAIR_TYPES = ["เครื่องยนต์", "ระบบทำความเย็น", "ระบบไฟฟ้า", "เบรก", "ยาง", "แบตเตอรี่", "ตัวถัง", "อื่นๆ"];
const DRIVER_LICENSE_TYPES = ["ท.2", "ท.3", "ท.4"];
const THAI_MONTHS = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];

/* ---------------------------------------------------------------
   HELPERS
---------------------------------------------------------------- */
function daysBetween(a, b) { return Math.round((b - a) / (1000 * 60 * 60 * 24)); }
function addDays(dateStr, n) { const d = new Date(dateStr); d.setDate(d.getDate() + n); return d; }
function fmtDate(d) { 
  if (!d) return "-"; 
  const date = typeof d === "string" ? new Date(d) : d; 
  return `${date.getDate()} ${THAI_MONTHS[date.getMonth()]} ${date.getFullYear() + 543}`; 
}
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

/* ---------------------------------------------------------------
   UI COMPONENTS
---------------------------------------------------------------- */
function PlateBadge({ plate, size = "md" }) {
  const parts = (plate || "").split("-");
  const prefix = parts[0] || plate || "-";
  const num = parts[1] || "";
  const big = size === "lg";
  return (
    <div style={{
      background: "linear-gradient(180deg, #F8FAFC 0%, #E2E8F0 100%)",
      border: "2px solid #0F172A",
      borderRadius: 8,
      padding: big ? "4px 12px" : "2px 8px",
      display: "inline-flex",
      flexDirection: "column",
      alignItems: "center",
      lineHeight: 1.1,
      minWidth: big ? 100 : 80,
      boxShadow: "0 2px 4px rgba(0,0,0,0.3)"
    }}>
      <span style={{ fontFamily: "'JetBrains Mono', monospace", fontWeight: 800, fontSize: big ? 18 : 14, color: "#0F172A", letterSpacing: 1 }}>
        {num ? `${prefix}-${num}` : prefix}
      </span>
      <span style={{ fontSize: big ? 9 : 7, color: "#334155", fontWeight: 700 }}>กรุงเทพมหานคร</span>
    </div>
  );
}

function TempChip({ temp }) {
  const meta = TEMP_META[temp] || TEMP_META.ambient;
  const Icon = meta.icon;
  return (
    <span className="inline-flex items-center gap-1.5 rounded-full px-2.5 py-1 text-xs font-semibold" 
          style={{ background: `${meta.color}15`, color: meta.color, border: `1px solid ${meta.color}33` }}>
      <Icon size={12} />{meta.label}
    </span>
  );
}

function StatusChip({ status, reason }) {
  const ready = status === "ready";
  return (
    <span className="inline-flex items-center gap-1.5 rounded-full px-2.5 py-1 text-xs font-semibold" 
          style={{ 
            background: ready ? "rgba(16,185,129,0.12)" : "rgba(239,68,68,0.12)", 
            color: ready ? "#10B981" : "#EF4444", 
            border: `1px solid ${ready ? "#10B98133" : "#EF444433"}` 
          }} 
          title={reason || ""}>
      {ready ? <CheckCircle2 size={13} /> : <AlertTriangle size={13} />}
      {ready ? "พร้อมใช้งาน" : "ไม่พร้อมใช้งาน"}
    </span>
  );
}

function Card({ children, style, className = "" }) {
  return (
    <div className={`p-5 rounded-2xl transition-all ${className}`} style={{ 
      background: "#141B26", 
      border: "1px solid #2A364F",
      boxShadow: "0 4px 20px -2px rgba(0, 0, 0, 0.5)",
      ...style 
    }}>
      {children}
    </div>
  );
}

/* ---------------------------------------------------------------
   MAIN DASHBOARD COMPONENT
---------------------------------------------------------------- */
export default function ColdChainDashboard() {
  const [activeTab, setActiveTab] = useState("dashboard");
  const [vehicles, setVehicles] = useState([]);
  const [repairs, setRepairs] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    // Mock Data หรือโหลดจาก Supabase
    setLoading(false);
  }, []);

  const menuItems = [
    { id: "dashboard", label: "ภาพรวมฟลีตรถ", icon: Gauge },
    { id: "vehicles", label: "รายการรถบรรทุก", icon: Truck },
    { id: "repairs", label: "การซ่อมบำรุง (11 Steps)", icon: Wrench },
    { id: "cd5_energy", label: "Project CD5 (ประหยัดพลังงาน)", icon: Zap },
    { id: "tsm_compliance", label: "มาตรฐาน TSM & Safety", icon: ShieldCheck },
    { id: "vendors", label: "ผู้รับเหมาขนส่ง (Suppliers)", icon: Building2 },
    { id: "drivers", label: "พนักงานขับรถ", icon: Users },
  ];

  return (
    <div className="min-h-screen text-slate-100 font-sans" style={{ background: "#0B0F17" }}>
      {/* Header Bar */}
      <header className="border-b border-slate-800 bg-slate-900/60 backdrop-blur-md sticky top-0 z-40 px-6 py-3.5 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="p-2 bg-sky-500/10 border border-sky-500/30 rounded-xl text-sky-400">
            <Snowflake size={22} />
          </div>
          <div>
            <h1 className="text-lg font-bold text-white tracking-wide">COLD CHAIN FLEET MANAGEMENT</h1>
            <p className="text-xs text-slate-400">ระบบบริหารฟลีตรถขนส่งสินค้าควบคุมอุณหภูมิ</p>
          </div>
        </div>

        <div className="flex items-center gap-4">
          <div className="hidden md:flex items-center gap-2 bg-slate-800/80 px-3 py-1.5 rounded-lg border border-slate-700/60 text-xs text-slate-300">
            <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse"></span>
            สถานะระบบ: ปกติ (Live Sync)
          </div>
        </div>
      </header>

      <div className="flex">
        {/* Sidebar Navigation */}
        <aside className="w-64 min-h-[calc(100vh-65px)] border-r border-slate-800 p-4 space-y-1 hidden lg:block">
          <div className="text-[11px] font-bold text-slate-500 uppercase tracking-wider px-3 mb-2">เมนูหลัก</div>
          {menuItems.map((item) => {
            const Icon = item.icon;
            const active = activeTab === item.id;
            return (
              <button
                key={item.id}
                onClick={() => setActiveTab(item.id)}
                className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-sm font-medium transition-all ${
                  active 
                    ? "bg-sky-500/15 text-sky-400 border border-sky-500/30" 
                    : "text-slate-400 hover:bg-slate-800/50 hover:text-slate-200"
                }`}
              >
                <Icon size={18} />
                {item.label}
              </button>
            );
          })}
        </aside>

        {/* Main Content Area */}
        <main className="flex-1 p-6 max-w-7xl mx-auto space-y-6">
          {/* Top KPI Cards */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <Card className="border-l-4 border-l-sky-500">
              <div className="flex items-center justify-between">
                <span className="text-xs font-semibold text-slate-400">จำนวนรถทั้งหมด</span>
                <Truck className="text-sky-400" size={20} />
              </div>
              <div className="text-3xl font-extrabold text-white mt-2">97 <span className="text-xs font-normal text-slate-400">คัน</span></div>
              <div className="text-xs text-slate-400 mt-1">พร้อมใช้งาน 92 คัน | รอซ่อม 5 คัน</div>
            </Card>

            <Card className="border-l-4 border-l-emerald-500">
              <div className="flex items-center justify-between">
                <span className="text-xs font-semibold text-slate-400">Project CD5 (ประหยัดพลังงาน)</span>
                <Zap className="text-emerald-400" size={20} />
              </div>
              <div className="text-3xl font-extrabold text-emerald-400 mt-2">38,400 <span className="text-xs font-normal text-slate-400">ลิตร/ปี</span></div>
              <div className="text-xs text-emerald-500/80 mt-1">ลดการใช้น้ำมันดีเซลขณะจอดคิว</div>
            </Card>

            <Card className="border-l-4 border-l-amber-500">
              <div className="flex items-center justify-between">
                <span className="text-xs font-semibold text-slate-400">รายการแจ้งเตือนด่วน</span>
                <Clock3 className="text-amber-400" size={20} />
              </div>
              <div className="text-3xl font-extrabold text-amber-400 mt-2">4 <span className="text-xs font-normal text-slate-400">รายการ</span></div>
              <div className="text-xs text-slate-400 mt-1">PM เครื่องยนต์ 2 | คาลิเบรท 2</div>
            </Card>

            <Card className="border-l-4 border-l-purple-500">
              <div className="flex items-center justify-between">
                <span className="text-xs font-semibold text-slate-400">ความพร้อม TSM Standard</span>
                <ShieldCheck className="text-purple-400" size={20} />
              </div>
              <div className="text-3xl font-extrabold text-purple-400 mt-2">98.5%</div>
              <div className="text-xs text-slate-400 mt-1">ผ่านเกณฑ์มาตรฐานขนส่งปลอดภัย</div>
            </Card>
          </div>

          {/* Quick Vehicle Status Overview Table */}
          <Card>
            <div className="flex items-center justify-between mb-4">
              <div>
                <h3 className="text-base font-bold text-white flex items-center gap-2">
                  <Truck size={18} className="text-sky-400" />
                  สถานะการบำรุงรักษารถควบคุมอุณหภูมิ
                </h3>
                <p className="text-xs text-slate-400 mt-0.5">ติดตามกำหนดการ PM เครื่องยนต์ ตู้เย็น และการต่อภาษี</p>
              </div>
              <button className="flex items-center gap-2 bg-sky-500 hover:bg-sky-400 text-slate-950 font-bold text-xs px-3.5 py-2 rounded-xl transition-all">
                <Plus size={15} /> เพิ่มรถคันใหม่
              </button>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse text-sm">
                <thead>
                  <tr className="border-b border-slate-800 text-slate-400 text-xs uppercase tracking-wider">
                    <th className="py-3 px-4">ทะเบียนรถ</th>
                    <th className="py-3 px-4">ยี่ห้อ/รุ่น</th>
                    <th className="py-3 px-4">ระบบความเย็น</th>
                    <th className="py-3 px-4">คนขับประจำ</th>
                    <th className="py-3 px-4">สถานะ</th>
                    <th className="py-3 px-4 text-right">การจัดการ</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800/60 text-slate-300">
                  <tr className="hover:bg-slate-800/30 transition-all">
                    <td className="py-3 px-4"><PlateBadge plate="70-1234" /></td>
                    <td className="py-3 px-4"><span className="font-semibold text-white">Isuzu FRR90</span> <span className="text-xs text-slate-400">(6 ล้อ)</span></td>
                    <td className="py-3 px-4"><TempChip temp="freezer" /></td>
                    <td className="py-3 px-4">สมชาย ใจดี</td>
                    <td className="py-3 px-4"><StatusChip status="ready" /></td>
                    <td className="py-3 px-4 text-right">
                      <button className="p-1.5 text-slate-400 hover:text-sky-400 transition-all"><Pencil size={16} /></button>
                      <button className="p-1.5 text-slate-400 hover:text-red-400 transition-all"><Trash2 size={16} /></button>
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </Card>
        </main>
      </div>
    </div>
  );
}
