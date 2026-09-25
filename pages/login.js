import { useEffect, useState } from "react";
import { useRouter } from "next/router";
import { supabase } from "../lib/supabaseClient";

const inputStyle = {
  background: "var(--surface-2)",
  border: "1px solid var(--border)",
  borderRadius: 8,
  padding: "10px 12px",
  color: "var(--text)",
  fontSize: 14,
  outline: "none",
  width: "100%",
};

function LogoMark({ size = 22 }) {
  return (
    <div className="flex items-center justify-center rounded-lg" style={{ width: 40, height: 40, background: "#FFFFFF", boxShadow: "0 2px 8px rgba(15,23,42,0.12)" }}>
      <span style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 800, fontSize: size * 0.68, fontStyle: "italic", letterSpacing: -0.5 }}>
        <span style={{ color: "#E4432E" }}>cp</span><span style={{ color: "#4C9A3D" }}>ram</span>
      </span>
    </div>
  );
}

export default function Login() {
  const router = useRouter();
  const [mode, setMode] = useState("signin");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => {
      if (data.session) router.replace("/");
    });
  }, [router]);

  async function handleSubmit(e) {
    e.preventDefault();
    setError("");
    setInfo("");
    setLoading(true);
    try {
      if (mode === "signin") {
        const { error } = await supabase.auth.signInWithPassword({ email, password });
        if (error) throw error;
        router.replace("/");
      } else {
        const { error } = await supabase.auth.signUp({ email, password });
        if (error) throw error;
        setInfo("สมัครสำเร็จ! ถ้าเปิดใช้การยืนยันอีเมลไว้ กรุณาเช็คอีเมลเพื่อยืนยันก่อนเข้าสู่ระบบ");
      }
    } catch (err) {
      setError(err.message || "เกิดข้อผิดพลาด กรุณาลองใหม่อีกครั้ง");
    } finally {
      setLoading(false);
    }
  }

  return (
    <div style={{
      "--bg": "#F4F6F9", "--surface": "#FFFFFF", "--surface-2": "#F1F4F8", "--border": "#E2E8F0",
      "--text": "#1E293B", "--text-muted": "#64748B", "--accent-frost": "#0E8FA0",
      minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "var(--bg)", padding: 16, fontFamily: "'Inter', sans-serif",
    }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@600;700;800&family=Inter:wght@400;500;600;700&display=swap');`}</style>
      <div style={{ width: "100%", maxWidth: 380 }}>
        <div className="flex items-center gap-3 mb-8 justify-center">
          <LogoMark />
          <div>
            <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 18, color: "var(--text)" }}>ข้อมูลรถบริษัท 
(ยานยนต์-ขนส่ง)</div>
            <div style={{ fontSize: 11, color: "var(--text-muted)" }}>ระบบบริหารรถโอนย้ายสินค้าหน่วยงานยานยนต์-ขนส่ง</div>
          </div>
        </div>

        <div style={{ background: "var(--surface)", border: "1px solid var(--border)", borderRadius: 14, padding: 24 }}>
          <div className="flex rounded-lg p-1 mb-5" style={{ background: "var(--surface-2)" }}>
            <button
              onClick={() => setMode("signin")}
              style={{ flex: 1, padding: "8px 0", borderRadius: 8, fontSize: 13, fontWeight: 600, background: mode === "signin" ? "var(--accent-frost)" : "transparent", color: mode === "signin" ? "#FFFFFF" : "var(--text-muted)" }}
            >
              เข้าสู่ระบบ
            </button>
            <button
              onClick={() => setMode("signup")}
              style={{ flex: 1, padding: "8px 0", borderRadius: 8, fontSize: 13, fontWeight: 600, background: mode === "signup" ? "var(--accent-frost)" : "transparent", color: mode === "signup" ? "#FFFFFF" : "var(--text-muted)" }}
            >
              สมัครใช้งานใหม่
            </button>
          </div>

          <form onSubmit={handleSubmit} className="flex flex-col gap-3">
            <label className="flex flex-col gap-1">
              <span style={{ fontSize: 12, color: "var(--text-muted)", fontWeight: 600 }}>อีเมล</span>
              <input style={inputStyle} type="email" required value={email} onChange={(e) => setEmail(e.target.value)} placeholder="name@cpram.co.th" />
            </label>
            <label className="flex flex-col gap-1">
              <span style={{ fontSize: 12, color: "var(--text-muted)", fontWeight: 600 }}>รหัสผ่าน</span>
              <input style={inputStyle} type="password" required minLength={6} value={password} onChange={(e) => setPassword(e.target.value)} placeholder="อย่างน้อย 6 ตัวอักษร" />
            </label>

            {error && (
              <div style={{ fontSize: 12, color: "#DC2626", background: "rgba(220,38,38,0.1)", border: "1px solid #DC262655", borderRadius: 8, padding: "8px 10px" }}>
                {error}
              </div>
            )}
            {info && (
              <div style={{ fontSize: 12, color: "#16A34A", background: "rgba(22,163,74,0.1)", border: "1px solid #16A34A55", borderRadius: 8, padding: "8px 10px" }}>
                {info}
              </div>
            )}

            <button
              type="submit"
              disabled={loading}
              style={{ background: "var(--accent-frost)", color: "#FFFFFF", border: "none", borderRadius: 8, padding: "10px 0", fontSize: 14, fontWeight: 700, cursor: "pointer", marginTop: 6, opacity: loading ? 0.6 : 1 }}
            >
              {loading ? "กำลังดำเนินการ..." : mode === "signin" ? "เข้าสู่ระบบ" : "สมัครใช้งาน"}
            </button>
          </form>
        </div>

        <p style={{ fontSize: 11, color: "var(--text-muted)", textAlign: "center", marginTop: 16 }}>
          ผู้ใช้ทุกคนที่เข้าสู่ระบบมีสิทธิ์แก้ไขข้อมูลได้เท่ากัน
        </p>
      </div>
    </div>
  );
}
