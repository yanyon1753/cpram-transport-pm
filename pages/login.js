import { useEffect, useState } from "react";
import { useRouter } from "next/router";
import { Snowflake } from "lucide-react";
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
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "var(--bg)", padding: 16 }}>
      <div style={{ width: "100%", maxWidth: 380 }}>
        <div className="flex items-center gap-3 mb-8 justify-center">
          <div className="flex items-center justify-center rounded-lg" style={{ width: 40, height: 40, background: "rgba(69,184,200,0.15)" }}>
            <Snowflake size={22} style={{ color: "var(--accent-frost)" }} />
          </div>
          <div>
            <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 18, color: "var(--text)" }}>CPRAM FLEET</div>
            <div style={{ fontSize: 11, color: "var(--text-muted)" }}>ระบบบริหารจัดการรถขนส่งควบคุมอุณหภูมิ</div>
          </div>
        </div>

        <div style={{ background: "var(--surface)", border: "1px solid var(--border)", borderRadius: 14, padding: 24 }}>
          <div className="flex rounded-lg p-1 mb-5" style={{ background: "var(--surface-2)" }}>
            <button
              onClick={() => setMode("signin")}
              style={{ flex: 1, padding: "8px 0", borderRadius: 8, fontSize: 13, fontWeight: 600, background: mode === "signin" ? "var(--accent-frost)" : "transparent", color: mode === "signin" ? "#0F1620" : "var(--text-muted)" }}
            >
              เข้าสู่ระบบ
            </button>
            <button
              onClick={() => setMode("signup")}
              style={{ flex: 1, padding: "8px 0", borderRadius: 8, fontSize: 13, fontWeight: 600, background: mode === "signup" ? "var(--accent-frost)" : "transparent", color: mode === "signup" ? "#0F1620" : "var(--text-muted)" }}
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
              <div style={{ fontSize: 12, color: "#E4584F", background: "rgba(228,88,79,0.1)", border: "1px solid #E4584F55", borderRadius: 8, padding: "8px 10px" }}>
                {error}
              </div>
            )}
            {info && (
              <div style={{ fontSize: 12, color: "#5FBE84", background: "rgba(95,190,132,0.1)", border: "1px solid #5FBE8455", borderRadius: 8, padding: "8px 10px" }}>
                {info}
              </div>
            )}

            <button
              type="submit"
              disabled={loading}
              style={{ background: "var(--accent-frost)", color: "#0F1620", border: "none", borderRadius: 8, padding: "10px 0", fontSize: 14, fontWeight: 700, cursor: "pointer", marginTop: 6, opacity: loading ? 0.6 : 1 }}
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
