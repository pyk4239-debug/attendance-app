import { useState, useEffect } from "react";
import { db } from "./firebase";
import {
  doc, onSnapshot, setDoc, getDoc, collection,
  getDocs, writeBatch
} from "firebase/firestore";

// 기존 서비스워커 완전 제거
if ("serviceWorker" in navigator) {
  navigator.serviceWorker.getRegistrations().then(registrations => {
    registrations.forEach(r => r.unregister());
  });
}

// ── 테마 (화이트모드) ──────────────────────────────────────────
const T = {
  bg: "#f5f6fa", card: "#ffffff", border: "#e8eaf0",
  text: "#1a1a2e", sub: "#6b7280", muted: "#9ca3af",
  headerBg: "#1a1a2e", headerText: "#ffffff",
  adminHeader: "#1e3a5f",
  green: "#16a34a", greenBg: "#dcfce7", greenText: "#15803d",
  blue: "#2563eb", blueBg: "#dbeafe", blueText: "#1d4ed8",
  yellow: "#d97706", yellowBg: "#fef3c7", yellowText: "#b45309",
  red: "#dc2626", redBg: "#fee2e2", redText: "#b91c1c",
  purple: "#7c3aed", purpleBg: "#ede9fe", purpleText: "#6d28d9",
  orange: "#ea580c", orangeBg: "#ffedd5", orangeText: "#c2410c",
};

// ── Firebase 컬렉션 키 ─────────────────────────────────────────
const COL_USERS    = "users";
const COL_RECORDS  = "records";
const COL_LEAVES   = "leaves";
const DOC_SETTINGS = "app/settings";

// ── 초기 데이터 ────────────────────────────────────────────────
const DEFAULT_USERS = [
  { id: "admin", name: "관리자", pin: "000000", role: "admin" },
  { id: "u1", name: "팀원1", pin: "111111", role: "member" },
  { id: "u2", name: "팀원2", pin: "222222", role: "member" },
  { id: "u3", name: "팀원3", pin: "333333", role: "member" },
  { id: "u4", name: "팀원4", pin: "444444", role: "member" },
];
const DEFAULT_SETTINGS = {
  workStart: "09:00", workEnd: "18:00",
  officeLat: null, officeLng: null, officeRadius: 200
};
const MASTER_CODE = "att2026!"; // 관리자 PIN 분실 시 비상 코드

// ── 유틸 ──────────────────────────────────────────────────────
function getToday() { return new Date().toISOString().slice(0, 10); }
function formatTime(iso) {
  if (!iso) return "--:--";
  return new Date(iso).toLocaleTimeString("ko-KR", { hour: "2-digit", minute: "2-digit", hour12: false });
}
function formatDate(d) {
  return new Date(d).toLocaleDateString("ko-KR", { month: "long", day: "numeric", weekday: "short" });
}
function isWeekend(d) { const w = new Date(d).getDay(); return w === 0 || w === 6; }
function isLate(iso, ws) {
  if (!iso || !ws) return false;
  const d = new Date(iso); const [h, m] = ws.split(":").map(Number);
  return (d.getHours() * 60 + d.getMinutes()) > (h * 60 + m);
}
function isEarlyOut(iso, we) {
  if (!iso || !we) return false;
  const d = new Date(iso); const [h, m] = we.split(":").map(Number);
  return (h * 60 + m) > (d.getHours() * 60 + d.getMinutes());
}
function calcLateMin(inI, ws) {
  if (!inI || !ws) return 0;
  const d = new Date(inI); const [h, m] = ws.split(":").map(Number);
  return Math.max(0, (d.getHours() * 60 + d.getMinutes()) - (h * 60 + m));
}
function calcEarlyOutMin(outI, we) {
  if (!outI || !we) return 0;
  const d = new Date(outI); const [h, m] = we.split(":").map(Number);
  return Math.max(0, (h * 60 + m) - (d.getHours() * 60 + d.getMinutes()));
}
function calcOvertimeMin(outI, we) {
  if (!outI || !we) return 0;
  const d = new Date(outI); const [h, m] = we.split(":").map(Number);
  return Math.max(0, (d.getHours() * 60 + d.getMinutes()) - (h * 60 + m));
}
// 조기출근 잔업 계산 (출근 기준 이전 출근분)
function calcEarlyInOvertimeMin(inI, ws) {
  if (!inI || !ws) return 0;
  const d = new Date(inI); const [h, m] = ws.split(":").map(Number);
  return Math.max(0, (h * 60 + m) - (d.getHours() * 60 + d.getMinutes()));
}
// 전체 잔업 = 조기출근 잔업 + 퇴근 후 잔업
function calcTotalOvertimeMin(inI, outI, ws, we) {
  return calcEarlyInOvertimeMin(inI, ws) + calcOvertimeMin(outI, we);
}
function fmtMinutes(min) {
  if (min == null || min === 0) return "-";
  const h = Math.floor(min / 60), m = Math.round(min % 60);
  return m === 0 ? `${h}시간` : `${h}시간 ${m}분`;
}
function roundTo30(min) { return Math.round(min / 30) * 30; }
function setTimeOnDate(date, time) {
  const d = new Date(`${date}T00:00:00`);
  const [h, m] = time.split(":").map(Number);
  d.setHours(h, m, 0, 0); return d.toISOString();
}
function calcMonthStats(days, settings) {
  return days.reduce((acc, [date, rec]) => {
    if (rec.in) {
      acc.days++;
      const lm = calcLateMin(rec.in, settings.workStart);
      const em = calcEarlyOutMin(rec.out, settings.workEnd);
      const om = calcTotalOvertimeMin(rec.in, rec.out, settings.workStart, settings.workEnd);
      if (lm > 0) { acc.late++; acc.lateMin += lm; }
      if (em > 0) { acc.early++; acc.earlyMin += em; }
      if (om >= 30) { acc.ot++; acc.otMin += roundTo30(om); }
      if (isWeekend(date)) acc.holiday++;
    }
    return acc;
  }, { days: 0, late: 0, lateMin: 0, early: 0, earlyMin: 0, ot: 0, otMin: 0, holiday: 0 });
}
function downloadCSV(filename, rows) {
  const csv = rows.map(r => r.map(c => '"' + (String(c || "").replace(/"/g, '""')) + '"').join(",")).join("\n");
  const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}
const LEAVE_TYPES = ["연차", "반차(오전)", "반차(오후)", "시간연차"];
const monthOptions = Array.from({ length: 12 }, (_, i) => {
  const d = new Date(new Date().getFullYear(), new Date().getMonth() - i, 1);
  return d.toISOString().slice(0, 7);
});
const monthLabel = m => { const [y, mo] = m.split("-"); return `${y}년 ${parseInt(mo)}월`; };

// ── GPS 유틸 ──────────────────────────────────────────────────
function getGPS() {
  return new Promise((resolve, reject) => {
    if (!navigator.geolocation) { reject(new Error("GPS 미지원")); return; }
    navigator.geolocation.getCurrentPosition(
      p => resolve({ lat: p.coords.latitude, lng: p.coords.longitude, acc: Math.round(p.coords.accuracy) }),
      e => reject(e),
      { timeout: 8000, maximumAge: 0, enableHighAccuracy: true }
    );
  });
}
function calcDistance(lat1, lng1, lat2, lng2) {
  if (lat1 == null || lng1 == null || lat2 == null || lng2 == null) return null;
  const R = 6371000;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLng = (lat2 - lng1) * Math.PI / 180;
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) * Math.sin(dLng / 2) ** 2;
  return Math.round(R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
}
function gpsStatusLabel(gps, settings) {
  if (!gps || gps.lat == null || gps.lng == null) return null;
  const officeLat = settings?.officeLat ?? null;
  const officeLng = settings?.officeLng ?? null;
  if (officeLat == null || officeLng == null) return { label: "위치기록", color: "gray" };
  const dist = calcDistance(gps.lat, gps.lng, Number(officeLat), Number(officeLng));
  if (dist == null) return { label: "위치기록", color: "gray" };
  const radius = Number(settings.officeRadius) || 200;
  if (dist <= radius) return { label: `회사 내 (${dist}m)`, color: "green" };
  return { label: `회사 외 (${dist}m)`, color: "red" };
}

// ── Firebase CRUD ──────────────────────────────────────────────
async function fbSaveRecord(userId, date, rec) {
  await setDoc(doc(db, COL_RECORDS, `${userId}_${date}`), { userId, date, ...rec });
}
async function fbSaveLeave(userId, date, leaveData) {
  if (leaveData) {
    await setDoc(doc(db, COL_LEAVES, `${userId}_${date}`), { userId, date, ...leaveData });
  } else {
    await setDoc(doc(db, COL_LEAVES, `${userId}_${date}`), { userId, date, deleted: true });
  }
}
async function fbSaveUsers(users) {
  const batch = writeBatch(db);
  users.forEach(u => batch.set(doc(db, COL_USERS, u.id), u));
  await batch.commit();
}
async function fbSaveSettings(settings) {
  await setDoc(doc(db, "app", "settings"), settings);
}

// ── Badge ──────────────────────────────────────────────────────
function Badge({ label, color }) {
  const map = {
    green: [T.greenBg, T.greenText], blue: [T.blueBg, T.blueText],
    yellow: [T.yellowBg, T.yellowText], red: [T.redBg, T.redText],
    purple: [T.purpleBg, T.purpleText], orange: [T.orangeBg, T.orangeText],
    gray: ["#f3f4f6", "#6b7280"]
  };
  const [bg, tc] = map[color] || map.gray;
  return <span style={{ fontSize: 10, background: bg, color: tc, padding: "2px 8px", borderRadius: 8, fontWeight: 700, whiteSpace: "nowrap" }}>{label}</span>;
}

function StatBox({ label, value, color }) {
  return (
    <div style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 12, padding: "10px 0", textAlign: "center", boxShadow: "0 1px 3px #0000000a" }}>
      <div style={{ fontSize: 10, color: T.muted, marginBottom: 3, fontWeight: 500 }}>{label}</div>
      <div style={{ fontSize: 14, fontWeight: 800, color: color || T.text }}>{value}</div>
    </div>
  );
}

function Btn({ children, onClick, variant = "default", disabled, fullWidth = true }) {
  const V = {
    default: { background: "#f3f4f6", color: T.text },
    green: { background: "#16a34a", color: "#fff" },
    blue: { background: "#2563eb", color: "#fff" },
    red: { background: "#dc2626", color: "#fff" },
    admin: { background: "#1e3a5f", color: "#fff" },
    primary: { background: "#2563eb", color: "#fff" },
    orange: { background: "#ea580c", color: "#fff" },
    ghost: { background: "#fff", color: T.text, border: `1px solid ${T.border}` }
  };
  const s = V[variant] || V.default;
  return (
    <button onClick={onClick} disabled={disabled}
      style={{ ...s, border: s.border || "none", borderRadius: 12, padding: "13px 0", fontSize: 14, fontWeight: 700, cursor: disabled ? "not-allowed" : "pointer", opacity: disabled ? 0.4 : 1, width: fullWidth ? "100%" : "auto", fontFamily: "inherit" }}>
      {children}
    </button>
  );
}

// ── 앱 로딩 래퍼 (Firebase 데이터 초기화) ─────────────────────
function AppLoader() {
  const [ready, setReady] = useState(false);
  const [users, setUsers] = useState([]);
  const [settings, setSettings] = useState(DEFAULT_SETTINGS);
  const [records, setRecords] = useState({});
  const [leaves, setLeaves] = useState({});

  useEffect(() => {
    let unsubUsers, unsubSettings, unsubRecords, unsubLeaves;

    // 유저 실시간 구독
    unsubUsers = onSnapshot(collection(db, COL_USERS), snap => {
      if (snap.empty) {
        // 최초 실행 시 기본 유저 등록
        fbSaveUsers(DEFAULT_USERS);
        setUsers(DEFAULT_USERS);
      } else {
        setUsers(snap.docs.map(d => d.data()));
      }
    });

    // 설정 실시간 구독
    unsubSettings = onSnapshot(doc(db, "app", "settings"), snap => {
      if (snap.exists()) setSettings(snap.data());
      else fbSaveSettings(DEFAULT_SETTINGS);
    });

    // 출퇴근 기록 실시간 구독
    unsubRecords = onSnapshot(collection(db, COL_RECORDS), snap => {
      const r = {};
      snap.docs.forEach(d => {
        const data = d.data();
        if (!r[data.userId]) r[data.userId] = {};
        const { userId, date, ...rest } = data;
        r[data.userId][data.date] = rest;
      });
      setRecords(r);
    });

    // 연차 기록 실시간 구독
    unsubLeaves = onSnapshot(collection(db, COL_LEAVES), snap => {
      const l = {};
      snap.docs.forEach(d => {
        const data = d.data();
        if (data.deleted) return;
        if (!l[data.userId]) l[data.userId] = {};
        const { userId, date, ...rest } = data;
        l[data.userId][data.date] = rest;
      });
      setLeaves(l);
      setReady(true);
    });

    return () => {
      unsubUsers?.(); unsubSettings?.(); unsubRecords?.(); unsubLeaves?.();
    };
  }, []);

  if (!ready) return (
    <div style={{ minHeight: "100vh", background: T.bg, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Noto Sans KR',sans-serif" }}>
      <div style={{ textAlign: "center" }}>
        <div style={{ width: 48, height: 48, border: `4px solid ${T.border}`, borderTop: `4px solid ${T.adminHeader}`, borderRadius: "50%", animation: "spin 1s linear infinite", margin: "0 auto 16px" }} />
        <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
        <div style={{ fontSize: 14, color: T.muted }}>연결 중...</div>
      </div>
    </div>
  );

  return <App users={users} settings={settings} records={records} leaves={leaves}
    onSaveUsers={fbSaveUsers} onSaveSettings={fbSaveSettings}
    onSaveRecord={fbSaveRecord} onSaveLeave={fbSaveLeave} />;
}

// ── 로그인 ─────────────────────────────────────────────────────
function LoginScreen({ users, onLogin, onUpdateUsers }) {
  const [name, setName] = useState(""), [pin, setPin] = useState(""), [err, setErr] = useState("");
  const [tab, setTab] = useState("login");
  const [newPin, setNewPin] = useState(""), [newPin2, setNewPin2] = useState(""), [setupErr, setSetupErr] = useState("");

  const login = () => {
    const u = users.find(u => u.name === name.trim() && u.pin === pin);
    if (u) onLogin(u);
    else { setErr("이름 또는 PIN이 맞지 않아요"); setTimeout(() => setErr(""), 2000); }
  };

  const changePin = async () => {
    const u = users.find(u => u.name === name.trim());
    if (!u || u.role === "admin") { setSetupErr("등록된 팀원 이름이 아니에요"); return; }
    if (!pin) { setSetupErr("현재 PIN을 입력해주세요"); return; }
    if (u.pin !== pin) { setSetupErr("현재 PIN이 맞지 않아요"); return; }
    if (!/^\d{6}$/.test(newPin)) { setSetupErr("새 PIN은 숫자 6자리여야 해요"); return; }
    if (newPin !== newPin2) { setSetupErr("새 PIN이 일치하지 않아요"); return; }
    const updated = users.map(x => x.id === u.id ? { ...x, pin: newPin } : x);
    await onUpdateUsers(updated);
    setSetupErr(""); setPin(""); setNewPin(""); setNewPin2("");
    setTab("login"); setErr("PIN이 변경됐어요 ✓");
    setTimeout(() => setErr(""), 2500);
  };

  const iStyle = { width: "100%", padding: "14px 16px", borderRadius: 12, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 15, fontWeight: 600, boxSizing: "border-box", fontFamily: "inherit", outline: "none" };

  return (
    <div style={{ minHeight: "100vh", background: T.bg, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: 24, fontFamily: "'Noto Sans KR',sans-serif" }}>
      <div style={{ background: T.headerBg, borderRadius: 20, padding: "20px 28px", marginBottom: 28, textAlign: "center" }}>
        <div style={{ color: "#ffffff60", fontSize: 11, letterSpacing: 4, marginBottom: 4 }}>ATTENDANCE</div>
        <div style={{ color: "#fff", fontSize: 22, fontWeight: 800, letterSpacing: -1 }}>출퇴근 관리</div>
      </div>
      <div style={{ display: "flex", background: T.card, borderRadius: 12, padding: 4, marginBottom: 24, border: `1px solid ${T.border}`, width: "100%", maxWidth: 300 }}>
        {[["login", "로그인"], ["setup", "PIN 변경"]].map(([key, label]) => (
          <button key={key} onClick={() => { setTab(key); setErr(""); setSetupErr(""); setPin(""); setNewPin(""); setNewPin2(""); }}
            style={{ flex: 1, padding: "9px 0", border: "none", borderRadius: 10, fontSize: 13, fontWeight: 700, cursor: "pointer", background: tab === key ? T.headerBg : "transparent", color: tab === key ? "#fff" : T.muted, fontFamily: "inherit" }}>
            {label}
          </button>
        ))}
      </div>
      <div style={{ width: "100%", maxWidth: 300 }}>
        {tab === "login" && <>
          <div style={{ marginBottom: 14 }}>
            <div style={{ fontSize: 13, color: T.sub, marginBottom: 6, fontWeight: 600 }}>이름</div>
            <input value={name} onChange={e => setName(e.target.value)} placeholder="이름 입력" onKeyDown={e => e.key === "Enter" && login()} style={iStyle} />
          </div>
          <div style={{ marginBottom: 20 }}>
            <div style={{ fontSize: 13, color: T.sub, marginBottom: 6, fontWeight: 600 }}>PIN</div>
            <input type="password" inputMode="numeric" maxLength={6} value={pin} onChange={e => setPin(e.target.value.replace(/\D/g, "").slice(0, 6))} placeholder="숫자 6자리" onKeyDown={e => e.key === "Enter" && login()} style={{ ...iStyle, letterSpacing: 6, fontSize: 20 }} />
          </div>
          {err && <div style={{ color: err.includes("✓") ? T.green : T.red, fontSize: 13, marginBottom: 12, textAlign: "center", fontWeight: 600 }}>{err}</div>}
          <Btn variant="admin" onClick={login}>로그인</Btn>
        </>}
        {tab === "setup" && <>
          <div style={{ background: T.blueBg, borderRadius: 10, padding: "10px 14px", marginBottom: 16, fontSize: 12, color: T.blueText, lineHeight: 1.6 }}>
            현재 PIN 확인 후 새 PIN으로 변경할 수 있어요.<br />본인만 변경 가능합니다.
          </div>
          <div style={{ marginBottom: 12 }}>
            <div style={{ fontSize: 13, color: T.sub, marginBottom: 6, fontWeight: 600 }}>이름</div>
            <input value={name} onChange={e => setName(e.target.value)} placeholder="본인 이름" style={iStyle} />
          </div>
          <div style={{ marginBottom: 12 }}>
            <div style={{ fontSize: 13, color: T.sub, marginBottom: 6, fontWeight: 600 }}>현재 PIN</div>
            <input type="password" inputMode="numeric" maxLength={6} value={pin} onChange={e => setPin(e.target.value.replace(/\D/g, "").slice(0, 6))} placeholder="현재 PIN 6자리" style={{ ...iStyle, letterSpacing: 6, fontSize: 18 }} />
          </div>
          <div style={{ marginBottom: 12 }}>
            <div style={{ fontSize: 13, color: T.sub, marginBottom: 6, fontWeight: 600 }}>새 PIN</div>
            <input type="password" inputMode="numeric" maxLength={6} value={newPin} onChange={e => setNewPin(e.target.value.replace(/\D/g, "").slice(0, 6))} placeholder="새 PIN 6자리" style={{ ...iStyle, letterSpacing: 6, fontSize: 18 }} />
          </div>
          <div style={{ marginBottom: 20 }}>
            <div style={{ fontSize: 13, color: T.sub, marginBottom: 6, fontWeight: 600 }}>새 PIN 확인</div>
            <input type="password" inputMode="numeric" maxLength={6} value={newPin2} onChange={e => setNewPin2(e.target.value.replace(/\D/g, "").slice(0, 6))} placeholder="새 PIN 재입력" style={{ ...iStyle, letterSpacing: 6, fontSize: 18 }} />
          </div>
          {setupErr && <div style={{ color: T.red, fontSize: 13, marginBottom: 12, textAlign: "center", fontWeight: 600 }}>{setupErr}</div>}
          <Btn variant="green" onClick={changePin}>PIN 변경</Btn>
        </>}
      </div>
    </div>
  );
}

// ── 팀원 화면 ──────────────────────────────────────────────────
function MemberScreen({ user, settings, records, leaves, onSaveRecord, onLogout }) {
  const [now, setNow] = useState(new Date());
  const [flash, setFlash] = useState(null);
  useEffect(() => { const t = setInterval(() => setNow(new Date()), 1000); return () => clearInterval(t); }, []);

  const today = now.toISOString().slice(0, 10); // now 기반 → 자정 넘으면 자동 갱신
  const todayRec = records[user.id]?.[today] || {};
  const hasIn = !!todayRec.in, hasOut = !!todayRec.out;
  const outings = todayRec.outing || [];
  const isOutside = outings.length > 0 && !outings[outings.length - 1].in;

  const punch = async type => {
    const iso = new Date().toISOString();
    let gps = null;
    try { gps = await getGPS(); } catch (e) { }
    let newRec = { ...todayRec };
    if (type === "in") newRec = { ...newRec, in: iso, inGps: gps };
    else if (type === "out") newRec = { ...newRec, out: iso, outGps: gps };
    else if (type === "outing_out") newRec = { ...newRec, outing: [...outings, { out: iso, outGps: gps }] };
    else if (type === "outing_in") newRec = { ...newRec, outing: outings.map((o, i) => i === outings.length - 1 ? { ...o, in: iso, inGps: gps } : o) };
    await onSaveRecord(user.id, today, newRec);
    const msgs = { in: "출근 완료! 👍", out: "퇴근 완료! 수고하셨어요 🙌", outing_out: "외출 처리됐어요 🚶", outing_in: "복귀 완료! 💪" };
    setFlash(msgs[type]); setTimeout(() => setFlash(null), 2500);
  };

  const thisMonth = now.toISOString().slice(0, 7);
  const monthDays = Object.entries(records[user.id] || {}).filter(([d]) => d.startsWith(thisMonth)).sort(([a], [b]) => b.localeCompare(a));
  const ms = calcMonthStats(monthDays, settings);
  const monthLeaves = Object.entries(leaves[user.id] || {}).filter(([d]) => d.startsWith(thisMonth));
  const annualCount = monthLeaves.filter(([, l]) => l.type === "연차").length;
  const lateToday = isLate(todayRec.in, settings.workStart);
  const earlyToday = isEarlyOut(todayRec.out, settings.workEnd);

  return (
    <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'Noto Sans KR',sans-serif", color: T.text }}>
      <div style={{ background: T.headerBg, padding: "18px 20px 24px" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <div>
            <div style={{ fontSize: 11, color: "#ffffff40", letterSpacing: 3 }}>ATTENDANCE</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>{user.name}님</div>
          </div>
          <button onClick={onLogout} style={{ background: "#ffffff18", border: "none", color: "#fff", padding: "7px 14px", borderRadius: 20, fontSize: 12, cursor: "pointer", fontWeight: 600 }}>로그아웃</button>
        </div>
        <div style={{ textAlign: "center" }}>
          <div style={{ fontSize: 44, fontWeight: 800, color: "#fff", letterSpacing: -2, fontVariantNumeric: "tabular-nums" }}>
            {now.toLocaleTimeString("ko-KR", { hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: false })}
          </div>
          <div style={{ fontSize: 13, color: "#ffffff60", marginTop: 4 }}>
            {now.toLocaleDateString("ko-KR", { year: "numeric", month: "long", day: "numeric", weekday: "long" })}
          </div>
          <div style={{ fontSize: 11, color: "#ffffff35", marginTop: 5 }}>출근 {settings.workStart} · 퇴근 {settings.workEnd}</div>
        </div>
      </div>

      <div style={{ padding: 16 }}>
        {/* 오늘 현황 */}
        <div style={{ background: T.card, borderRadius: 16, padding: "16px 20px", marginBottom: 14, boxShadow: "0 2px 8px #0000000d", border: `1px solid ${T.border}` }}>
          <div style={{ display: "flex", marginBottom: 10 }}>
            {[
              ["출근", hasIn ? formatTime(todayRec.in) : "--:--", hasIn ? (lateToday ? T.yellow : T.green) : T.muted],
              ["퇴근", hasOut ? formatTime(todayRec.out) : "--:--", hasOut ? (earlyToday ? T.orange : T.blue) : T.muted],
              ["상태", !hasIn ? "대기중" : !hasOut ? "근무중" : "완료", !hasIn ? T.muted : !hasOut ? T.green : T.blue]
            ].map(([label, val, color], i) => (
              <div key={i} style={{ flex: 1, textAlign: "center", borderLeft: i > 0 ? `1px solid ${T.border}` : "none" }}>
                <div style={{ fontSize: 11, color: T.muted, marginBottom: 4, fontWeight: 500 }}>{label}</div>
                <div style={{ fontSize: i === 2 ? 13 : 17, fontWeight: 800, color }}>{val}</div>
              </div>
            ))}
          </div>
          <div style={{ display: "flex", gap: 5, flexWrap: "wrap" }}>
            {lateToday && hasIn && <Badge label="지각" color="yellow" />}
            {earlyToday && hasOut && <Badge label="조퇴" color="orange" />}
            {isOutside && <Badge label="외출중" color="blue" />}
            {outings.filter(o => o.in).length > 0 && <Badge label={`외출 ${outings.filter(o => o.in).length}회`} color="gray" />}
            {isWeekend(today) && hasIn && <Badge label="휴일근무" color="red" />}
            {todayRec.inGps && (() => { const s = gpsStatusLabel(todayRec.inGps, settings); return s ? <Badge label={`출근 ${s.label}`} color={s.color} /> : null; })()}
            {todayRec.outGps && (() => { const s = gpsStatusLabel(todayRec.outGps, settings); return s ? <Badge label={`퇴근 ${s.label}`} color={s.color} /> : null; })()}
          </div>
        </div>

        {/* 버튼 */}
        {flash ? (
          <div style={{ textAlign: "center", padding: "16px 0", fontSize: 16, fontWeight: 700, color: T.green }}>{flash}</div>
        ) : (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 10 }}>
            <Btn variant="green" onClick={() => punch("in")} disabled={hasIn}>{hasIn ? "✓ 출근완료" : "출근"}</Btn>
            <Btn variant="blue" onClick={() => punch("out")} disabled={!hasIn || hasOut}>{hasOut ? "✓ 퇴근완료" : "퇴근"}</Btn>
          </div>
        )}
        {hasIn && !hasOut && (
          <div style={{ marginBottom: 14 }}>
            {!isOutside
              ? <Btn variant="orange" onClick={() => punch("outing_out")}>🚶 외출</Btn>
              : <Btn variant="primary" onClick={() => punch("outing_in")}>🏃 복귀</Btn>}
          </div>
        )}

        {/* 이번달 현황 */}
        <div style={{ background: T.card, borderRadius: 16, padding: "14px 16px", marginBottom: 16, border: `1px solid ${T.border}`, boxShadow: "0 1px 4px #0000000a" }}>
          <div style={{ fontSize: 12, color: T.muted, marginBottom: 10, fontWeight: 600 }}>이번달 현황</div>
          {[
            [["출근", ms.days + "일", T.green], ["지각", ms.late + "회", T.yellow], ["지각시간", fmtMinutes(ms.lateMin), T.yellow]],
            [["휴일", ms.holiday + "일", T.red], ["잔업", ms.ot + "일", T.purple], ["잔업시간", fmtMinutes(ms.otMin), T.purple]],
            [["연차", annualCount + "일", "#7c3aed"], ["조퇴", ms.early + "회", T.orange], ["조퇴시간", fmtMinutes(ms.earlyMin), T.orange]],
          ].map((row, ri) => (
            <div key={ri} style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 6, marginBottom: ri < 2 ? 6 : 0 }}>
              {row.map(([l, v, c]) => <StatBox key={l} label={l} value={v} color={c} />)}
            </div>
          ))}
        </div>

        {/* 이번달 기록 */}
        <div style={{ fontSize: 13, color: T.muted, marginBottom: 10, fontWeight: 600 }}>이번달 기록</div>
        {monthDays.length === 0
          ? <div style={{ textAlign: "center", color: T.muted, padding: 30, fontSize: 14, background: T.card, borderRadius: 12, border: `1px solid ${T.border}` }}>기록 없음</div>
          : monthDays.map(([date, rec]) => {
            const lm = calcLateMin(rec.in, settings.workStart);
            const em = calcEarlyOutMin(rec.out, settings.workEnd);
            const om = calcTotalOvertimeMin(rec.in, rec.out, settings.workStart, settings.workEnd);
            const late = lm > 0, early = em > 0, ot = om >= 30;
            const weekend = isWeekend(date), leave = leaves[user.id]?.[date];
            const isNormal = !late && !early && !ot && !weekend && rec.in && rec.out;
            return (
              <div key={date} style={{ background: T.card, borderRadius: 12, padding: "12px 14px", marginBottom: 8, border: `1px solid ${T.border}` }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: rec.in ? 6 : 0 }}>
                  <div style={{ fontSize: 13, color: weekend ? T.red : T.sub, minWidth: 76, fontWeight: 600 }}>{formatDate(date)}</div>
                  <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
                    {weekend && rec.in && <Badge label="휴일" color="red" />}
                    {late && <Badge label={`지각 ${fmtMinutes(lm)}`} color="yellow" />}
                    {early && <Badge label={`조퇴 ${fmtMinutes(em)}`} color="orange" />}
                    {(rec.outing || []).length > 0 && <Badge label="외출" color="blue" />}
                    {ot && <Badge label={`잔업 ${fmtMinutes(roundTo30(om))}`} color="purple" />}
                    {leave && <Badge label={leave.type} color="purple" />}
                    {isNormal && <Badge label="정상" color="green" />}
                  </div>
                </div>
                {rec.in && (
                  <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    <span style={{ fontSize: 13, fontWeight: 700, color: late ? T.yellow : T.green }}>{formatTime(rec.in)}</span>
                    <span style={{ fontSize: 11, color: T.muted }}>→</span>
                    <span style={{ fontSize: 13, fontWeight: 700, color: early ? T.orange : T.blue }}>{formatTime(rec.out)}</span>
                  </div>
                )}
                {rec.note && <div style={{ fontSize: 11, color: T.muted, marginTop: 4 }}>📝 {rec.note}</div>}
              </div>
            );
          })
        }
      </div>
    </div>
  );
}

// ── 출퇴근 수정 모달 ───────────────────────────────────────────
function EditRecordModal({ user, date, rec, settings, userLeaves, onSave, onClose }) {
  const [inTime, setInTime] = useState(rec?.in ? formatTime(rec.in) : "");
  const [outTime, setOutTime] = useState(rec?.out ? formatTime(rec.out) : "");
  const [outings, setOutings] = useState(rec?.outing || []);
  const [leaveType, setLeaveType] = useState(userLeaves?.[date]?.type || "");
  const [leaveHours, setLeaveHours] = useState(userLeaves?.[date]?.hours || 1);
  const [note, setNote] = useState(rec?.note || "");
  const times = [];
  for (let h = 0; h < 24; h++) for (let m of [0, 30]) times.push(`${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`);
  const iStyle = { width: "100%", padding: "11px 14px", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 14, fontWeight: 600, boxSizing: "border-box", fontFamily: "inherit" };
  const inIso = inTime ? setTimeOnDate(date, inTime) : null;
  const outIso = outTime ? setTimeOnDate(date, outTime) : null;
  const lm = calcLateMin(inIso, settings.workStart);
  const em = calcEarlyOutMin(outIso, settings.workEnd);
  const om = calcOvertimeMin(outIso, settings.workEnd);

  const save = async () => {
    const nr = { ...rec, outing: outings, note };
    if (inTime) nr.in = setTimeOnDate(date, inTime); else delete nr.in;
    if (outTime) nr.out = setTimeOnDate(date, outTime); else delete nr.out;
    // GPS 데이터는 수정 시 유지
    await onSave(date, nr, leaveType ? { type: leaveType, hours: leaveHours } : null);
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "#00000066", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, padding: 16 }}>
      <div style={{ background: T.card, borderRadius: 20, padding: 22, width: "100%", maxWidth: 340, maxHeight: "90vh", overflowY: "auto", boxShadow: "0 20px 60px #00000030" }}>
        <div style={{ fontWeight: 800, fontSize: 17, color: T.text, marginBottom: 4 }}>{user.name} — 기록 수정</div>
        <div style={{ fontSize: 12, color: T.muted, marginBottom: 18 }}>{formatDate(date)}</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          {[["출근", inTime, setInTime, T.green], ["퇴근", outTime, setOutTime, T.blue]].map(([label, val, setter, color]) => (
            <div key={label}>
              <div style={{ fontSize: 12, color, marginBottom: 5, fontWeight: 700 }}>{label}</div>
              <select value={val} onChange={e => setter(e.target.value)} style={{ ...iStyle, color: val ? color : T.muted }}>
                <option value="">-</option>
                {times.map(t => <option key={t} value={t}>{t}</option>)}
              </select>
            </div>
          ))}
        </div>
        {inTime && outTime && (
          <div style={{ background: T.bg, borderRadius: 10, padding: "9px 12px", marginBottom: 12, display: "flex", gap: 6, flexWrap: "wrap" }}>
            {!lm && !em && om < 30 && !isWeekend(date) && <Badge label="정상" color="green" />}
            {lm > 0 && <Badge label={`지각 ${fmtMinutes(lm)}`} color="yellow" />}
            {em > 0 && <Badge label={`조퇴 ${fmtMinutes(em)}`} color="orange" />}
            {om >= 30 && <Badge label={`잔업 ${fmtMinutes(roundTo30(om))}`} color="purple" />}
            {isWeekend(date) && inTime && <Badge label="휴일근무" color="red" />}
          </div>
        )}
        {outings.map((o, i) => (
          <div key={i} style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 6 }}>
            <select value={o.out ? formatTime(o.out) : ""} onChange={e => { const n = [...outings]; n[i] = { ...n[i], out: e.target.value ? setTimeOnDate(date, e.target.value) : null }; setOutings(n); }} style={{ ...iStyle, flex: 1, fontSize: 12 }}>
              <option value="">외출</option>{times.map(t => <option key={t} value={t}>{t}</option>)}
            </select>
            <span style={{ color: T.muted, fontSize: 12 }}>→</span>
            <select value={o.in ? formatTime(o.in) : ""} onChange={e => { const n = [...outings]; n[i] = { ...n[i], in: e.target.value ? setTimeOnDate(date, e.target.value) : null }; setOutings(n); }} style={{ ...iStyle, flex: 1, fontSize: 12 }}>
              <option value="">복귀</option>{times.map(t => <option key={t} value={t}>{t}</option>)}
            </select>
            <button onClick={() => setOutings(outings.filter((_, ii) => ii !== i))} style={{ background: T.redBg, border: "none", color: T.red, borderRadius: 8, padding: "8px 10px", cursor: "pointer", fontWeight: 700 }}>✕</button>
          </div>
        ))}
        <button onClick={() => setOutings([...outings, { out: null, in: null }])} style={{ width: "100%", padding: "8px 0", borderRadius: 10, border: `1px dashed ${T.border}`, background: "none", color: T.muted, fontSize: 12, cursor: "pointer", marginBottom: 12, fontWeight: 600 }}>+ 외출 추가</button>
        <div style={{ marginBottom: 12 }}>
          <div style={{ fontSize: 12, color: T.purple, marginBottom: 5, fontWeight: 700 }}>연차 / 반차</div>
          <select value={leaveType} onChange={e => setLeaveType(e.target.value)} style={iStyle}>
            <option value="">해당없음</option>
            {LEAVE_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
          </select>
          {leaveType === "시간연차" && <select value={leaveHours} onChange={e => setLeaveHours(Number(e.target.value))} style={{ ...iStyle, marginTop: 6 }}>{[1, 2, 3, 4].map(h => <option key={h} value={h}>{h}시간</option>)}</select>}
        </div>
        <div style={{ marginBottom: 18 }}>
          <div style={{ fontSize: 12, color: T.muted, marginBottom: 5, fontWeight: 600 }}>메모</div>
          <input value={note} onChange={e => setNote(e.target.value)} placeholder="특이사항" style={iStyle} />
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          <Btn variant="ghost" onClick={onClose}>취소</Btn>
          <Btn variant="admin" onClick={save}>저장</Btn>
        </div>
      </div>
    </div>
  );
}

// ── 월별 탭 ────────────────────────────────────────────────────
function MonthTab({ records, leaves, members, settings, onSaveRecord, onSaveLeave }) {
  const [selectedMonth, setSelectedMonth] = useState(new Date().toISOString().slice(0, 7));
  const [drillUser, setDrillUser] = useState(null);
  const [editTarget, setEditTarget] = useState(null);

  const handleSaveRecord = async (date, newRec, leaveData) => {
    await onSaveRecord(editTarget.user.id, date, newRec);
    await onSaveLeave(editTarget.user.id, date, leaveData);
    setEditTarget(null);
  };

  if (drillUser) {
    const days = Object.entries(records[drillUser.id] || {}).filter(([d]) => d.startsWith(selectedMonth)).sort(([a], [b]) => a.localeCompare(b));
    const userLeaves = leaves[drillUser.id] || {};
    const mLeaves = Object.entries(userLeaves).filter(([d]) => d.startsWith(selectedMonth));
    const ms = calcMonthStats(days, settings);
    ms.annual = mLeaves.filter(([, l]) => l.type === "연차").length;

    const handleDownload = () => {
      const header = ["날짜", "요일", "출근", "퇴근", "지각", "지각시간", "조퇴", "조퇴시간", "잔업", "잔업시간", "외출", "연차/반차", "메모"];
      const rows = days.map(([date, rec]) => {
        const dow = new Date(date).toLocaleDateString("ko-KR", { weekday: "short" });
        const lm = calcLateMin(rec.in, settings.workStart), em = calcEarlyOutMin(rec.out, settings.workEnd), om = calcTotalOvertimeMin(rec.in, rec.out, settings.workStart, settings.workEnd);
        const leave = userLeaves[date];
        return [date, dow, formatTime(rec.in), formatTime(rec.out), lm > 0 ? "O" : "", lm > 0 ? fmtMinutes(lm) : "", em > 0 ? "O" : "", em > 0 ? fmtMinutes(em) : "", om >= 30 ? "O" : "", om >= 30 ? fmtMinutes(roundTo30(om)) : "", (rec.outing || []).length > 0 ? `${(rec.outing || []).length}회` : "", leave ? leave.type : "", rec.note || ""];
      });
      downloadCSV(`${drillUser.name}_${monthLabel(selectedMonth)}_근태.csv`, [header, ...rows]);
    };

    return (
      <div>
        <button onClick={() => setDrillUser(null)} style={{ background: "none", border: "none", color: T.muted, cursor: "pointer", marginBottom: 14, fontSize: 13, padding: 0, fontWeight: 600 }}>← 전체 목록</button>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 16 }}>
          <div style={{ width: 42, height: 42, borderRadius: "50%", background: T.headerBg, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 17, fontWeight: 800, color: "#fff" }}>{drillUser.name[0]}</div>
          <div style={{ flex: 1 }}>
            <div style={{ fontWeight: 800, fontSize: 17, color: T.text }}>{drillUser.name}</div>
            <div style={{ fontSize: 12, color: T.muted }}>{monthLabel(selectedMonth)}</div>
          </div>
          <button onClick={handleDownload} style={{ background: T.green, border: "none", color: "#fff", borderRadius: 10, padding: "8px 14px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>⬇ CSV</button>
        </div>
        {[
          [["출근", ms.days + "일", T.green], ["지각", ms.late + "회", T.yellow], ["지각시간", fmtMinutes(ms.lateMin), T.yellow]],
          [["휴일", ms.holiday + "일", T.red], ["잔업", ms.ot + "일", T.purple], ["잔업시간", fmtMinutes(ms.otMin), T.purple]],
          [["연차", ms.annual + "일", "#7c3aed"], ["조퇴", ms.early + "회", T.orange], ["조퇴시간", fmtMinutes(ms.earlyMin), T.orange]],
        ].map((row, ri) => (
          <div key={ri} style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 6, marginBottom: ri < 2 ? 6 : 12 }}>
            {row.map(([l, v, c]) => <StatBox key={l} label={l} value={v} color={c} />)}
          </div>
        ))}
        <div style={{ fontSize: 12, color: T.muted, marginBottom: 10, fontWeight: 600 }}>날짜별 상세</div>
        {days.length === 0
          ? <div style={{ textAlign: "center", color: T.muted, padding: 24, fontSize: 14, background: T.card, borderRadius: 12, border: `1px solid ${T.border}` }}>기록 없음</div>
          : days.map(([date, rec]) => {
            const lm = calcLateMin(rec.in, settings.workStart), em = calcEarlyOutMin(rec.out, settings.workEnd), om = calcTotalOvertimeMin(rec.in, rec.out, settings.workStart, settings.workEnd);
            const late = lm > 0, early = em > 0, ot = om >= 30, weekend = isWeekend(date), leave = userLeaves[date];
            const [, , dd] = date.split("-"), dow = new Date(date).toLocaleDateString("ko-KR", { weekday: "short" });
            const isNormal = !late && !early && !ot && !weekend && rec.in && rec.out;
            return (
              <div key={date} style={{ background: T.card, borderRadius: 12, padding: "12px 14px", marginBottom: 8, border: `1px solid ${T.border}` }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <div style={{ minWidth: 36, textAlign: "center" }}>
                    <div style={{ fontSize: 17, fontWeight: 800, color: weekend ? T.red : T.text }}>{parseInt(dd)}</div>
                    <div style={{ fontSize: 10, color: T.muted }}>{dow}</div>
                  </div>
                  <div style={{ width: 1, height: 34, background: T.border }} />
                  <div style={{ flex: 1 }}>
                    <div style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 4 }}>
                      <span style={{ fontSize: 13, fontWeight: 700, color: late ? T.yellow : T.green }}>{formatTime(rec.in)}</span>
                      <span style={{ fontSize: 11, color: T.muted }}>→</span>
                      <span style={{ fontSize: 13, fontWeight: 700, color: early ? T.orange : T.blue }}>{formatTime(rec.out)}</span>
                      {isNormal && <span style={{ marginLeft: "auto", fontSize: 11, color: T.muted }}>정상</span>}
                    </div>
                    <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
                      {late && <Badge label={`지각 ${fmtMinutes(lm)}`} color="yellow" />}
                      {early && <Badge label={`조퇴 ${fmtMinutes(em)}`} color="orange" />}
                      {ot && <Badge label={`잔업 ${fmtMinutes(roundTo30(om))}`} color="purple" />}
                      {weekend && rec.in && <Badge label="휴일" color="red" />}
                      {(rec.outing || []).length > 0 && <Badge label="외출" color="blue" />}
                      {leave && <Badge label={leave.type} color="purple" />}
                      {rec.note && <Badge label={`📝 ${rec.note}`} color="gray" />}
                    </div>
                  </div>
                  <button onClick={() => setEditTarget({ user: drillUser, date })} style={{ background: T.bg, border: `1px solid ${T.border}`, color: T.text, borderRadius: 8, padding: "5px 10px", fontSize: 11, cursor: "pointer", fontWeight: 600, flexShrink: 0 }}>수정</button>
                </div>
              </div>
            );
          })
        }
        {editTarget && (
          <EditRecordModal user={editTarget.user} date={editTarget.date}
            rec={records[editTarget.user.id]?.[editTarget.date] || {}}
            settings={settings} userLeaves={leaves[editTarget.user.id] || {}}
            onSave={handleSaveRecord} onClose={() => setEditTarget(null)} />
        )}
      </div>
    );
  }

  return (
    <div>
      <select value={selectedMonth} onChange={e => { setSelectedMonth(e.target.value); setDrillUser(null); }}
        style={{ width: "100%", padding: "11px 16px", borderRadius: 12, border: `1px solid ${T.border}`, background: T.card, color: T.text, fontSize: 15, fontWeight: 700, marginBottom: 16, boxSizing: "border-box" }}>
        {monthOptions.map(m => <option key={m} value={m}>{monthLabel(m)}</option>)}
      </select>
      {members.map(u => {
        const days = Object.entries(records[u.id] || {}).filter(([d]) => d.startsWith(selectedMonth));
        const ms = calcMonthStats(days, settings);
        return (
          <div key={u.id} style={{ background: T.card, borderRadius: 16, padding: "14px 16px", marginBottom: 12, border: `1px solid ${T.border}`, boxShadow: "0 1px 4px #0000000a" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 12 }}>
              <div style={{ width: 38, height: 38, borderRadius: "50%", background: T.headerBg, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 15, fontWeight: 800, color: "#fff" }}>{u.name[0]}</div>
              <span style={{ fontWeight: 700, fontSize: 15, flex: 1, color: T.text }}>{u.name}</span>
              <button onClick={() => setDrillUser(u)} style={{ background: T.headerBg, border: "none", color: "#fff", borderRadius: 10, padding: "7px 14px", fontSize: 12, cursor: "pointer", fontWeight: 700 }}>상세 →</button>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(5,1fr)", gap: 6 }}>
              {[["출근", ms.days + "일", T.green], ["지각", ms.late + "회", T.yellow], ["조퇴", ms.early + "회", T.orange], ["잔업", ms.ot + "일", T.purple], ["휴일", ms.holiday + "일", T.red]].map(([l, v, c]) => (
                <StatBox key={l} label={l} value={v} color={c} />
              ))}
            </div>
          </div>
        );
      })}
    </div>
  );
}

// ── 관리자 계정 모달 ───────────────────────────────────────────
function AdminAccountModal({ users, onUpdateUsers, onClose }) {
  const admin = users.find(u => u.role === "admin") || {};
  const [tab, setTab] = useState("info");
  const [name, setName] = useState(admin.name || "");
  const [curPin, setCurPin] = useState(""), [newPin, setNewPin] = useState(""), [newPin2, setNewPin2] = useState("");
  const [masterCode, setMasterCode] = useState("");
  const [err, setErr] = useState(""), [ok, setOk] = useState("");

  const iStyle = { width: "100%", padding: "12px 14px", borderRadius: 12, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 15, fontWeight: 600, boxSizing: "border-box", fontFamily: "inherit", marginBottom: 12, outline: "none" };

  const saveName = async () => {
    if (!name.trim()) { setErr("이름을 입력해주세요"); return; }
    await onUpdateUsers(users.map(u => u.role === "admin" ? { ...u, name: name.trim() } : u));
    setErr(""); setOk("이름이 변경됐어요 ✓"); setTimeout(() => setOk(""), 2000);
  };
  const savePin = async () => {
    if (admin.pin !== curPin) { setErr("현재 PIN이 맞지 않아요"); return; }
    if (!/^\d{6}$/.test(newPin)) { setErr("새 PIN은 숫자 6자리여야 해요"); return; }
    if (newPin !== newPin2) { setErr("새 PIN이 일치하지 않아요"); return; }
    await onUpdateUsers(users.map(u => u.role === "admin" ? { ...u, pin: newPin } : u));
    setCurPin(""); setNewPin(""); setNewPin2("");
    setErr(""); setOk("PIN이 변경됐어요 ✓"); setTimeout(() => setOk(""), 2000);
  };
  const resetByMaster = async () => {
    if (masterCode !== MASTER_CODE) { setErr("마스터코드가 맞지 않아요"); return; }
    if (!/^\d{6}$/.test(newPin)) { setErr("새 PIN은 숫자 6자리여야 해요"); return; }
    if (newPin !== newPin2) { setErr("새 PIN이 일치하지 않아요"); return; }
    await onUpdateUsers(users.map(u => u.role === "admin" ? { ...u, pin: newPin } : u));
    setMasterCode(""); setNewPin(""); setNewPin2("");
    setErr(""); setOk("PIN이 초기화됐어요 ✓"); setTimeout(() => setOk(""), 2000);
  };

  const tabBtn = (key, label) => (
    <button onClick={() => { setTab(key); setErr(""); setOk(""); }}
      style={{ flex: 1, padding: "8px 0", border: "none", borderRadius: 9, fontSize: 12, fontWeight: 700, cursor: "pointer", background: tab === key ? T.headerBg : "transparent", color: tab === key ? "#fff" : T.muted, fontFamily: "inherit" }}>
      {label}
    </button>
  );

  return (
    <div style={{ position: "fixed", inset: 0, background: "#00000066", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, padding: 20 }}>
      <div style={{ background: T.card, borderRadius: 20, padding: 22, width: "100%", maxWidth: 320, boxShadow: "0 20px 60px #00000020" }}>
        <div style={{ fontWeight: 800, fontSize: 17, color: T.text, marginBottom: 16 }}>관리자 계정</div>
        <div style={{ display: "flex", background: T.bg, borderRadius: 12, padding: 3, marginBottom: 20, border: `1px solid ${T.border}` }}>
          {tabBtn("info", "이름변경")}{tabBtn("pin", "PIN변경")}{tabBtn("lost", "PIN분실")}
        </div>
        {tab === "info" && <>
          <div style={{ fontSize: 13, color: T.sub, marginBottom: 6, fontWeight: 600 }}>관리자 이름</div>
          <input value={name} onChange={e => setName(e.target.value)} placeholder="이름 입력" style={iStyle} />
          {err && <div style={{ color: T.red, fontSize: 13, marginBottom: 8, fontWeight: 600 }}>{err}</div>}
          {ok && <div style={{ color: T.green, fontSize: 13, marginBottom: 8, fontWeight: 600 }}>{ok}</div>}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
            <Btn variant="ghost" onClick={onClose}>닫기</Btn>
            <Btn variant="admin" onClick={saveName}>저장</Btn>
          </div>
        </>}
        {tab === "pin" && <>
          <div style={{ fontSize: 13, color: T.sub, marginBottom: 6, fontWeight: 600 }}>현재 PIN</div>
          <input type="password" inputMode="numeric" maxLength={6} value={curPin} onChange={e => setCurPin(e.target.value.replace(/\D/g, "").slice(0, 6))} placeholder="현재 PIN 6자리" style={{ ...iStyle, letterSpacing: 6, fontSize: 18 }} />
          <div style={{ fontSize: 13, color: T.sub, marginBottom: 6, fontWeight: 600 }}>새 PIN</div>
          <input type="password" inputMode="numeric" maxLength={6} value={newPin} onChange={e => setNewPin(e.target.value.replace(/\D/g, "").slice(0, 6))} placeholder="새 PIN 6자리" style={{ ...iStyle, letterSpacing: 6, fontSize: 18 }} />
          <div style={{ fontSize: 13, color: T.sub, marginBottom: 6, fontWeight: 600 }}>새 PIN 확인</div>
          <input type="password" inputMode="numeric" maxLength={6} value={newPin2} onChange={e => setNewPin2(e.target.value.replace(/\D/g, "").slice(0, 6))} placeholder="새 PIN 재입력" style={{ ...iStyle, letterSpacing: 6, fontSize: 18 }} />
          {err && <div style={{ color: T.red, fontSize: 13, marginBottom: 8, fontWeight: 600 }}>{err}</div>}
          {ok && <div style={{ color: T.green, fontSize: 13, marginBottom: 8, fontWeight: 600 }}>{ok}</div>}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
            <Btn variant="ghost" onClick={onClose}>닫기</Btn>
            <Btn variant="admin" onClick={savePin}>변경</Btn>
          </div>
        </>}
        {tab === "lost" && <>
          <div style={{ background: T.yellowBg, borderRadius: 10, padding: "10px 14px", marginBottom: 16, fontSize: 12, color: T.yellowText, lineHeight: 1.6 }}>
            PIN을 잊어버린 경우 마스터코드로 초기화할 수 있어요.
          </div>
          <div style={{ fontSize: 13, color: T.sub, marginBottom: 6, fontWeight: 600 }}>마스터코드</div>
          <input type="password" value={masterCode} onChange={e => setMasterCode(e.target.value)} placeholder="마스터코드 입력" style={iStyle} />
          <div style={{ fontSize: 13, color: T.sub, marginBottom: 6, fontWeight: 600 }}>새 PIN</div>
          <input type="password" inputMode="numeric" maxLength={6} value={newPin} onChange={e => setNewPin(e.target.value.replace(/\D/g, "").slice(0, 6))} placeholder="새 PIN 6자리" style={{ ...iStyle, letterSpacing: 6, fontSize: 18 }} />
          <div style={{ fontSize: 13, color: T.sub, marginBottom: 6, fontWeight: 600 }}>새 PIN 확인</div>
          <input type="password" inputMode="numeric" maxLength={6} value={newPin2} onChange={e => setNewPin2(e.target.value.replace(/\D/g, "").slice(0, 6))} placeholder="새 PIN 재입력" style={{ ...iStyle, letterSpacing: 6, fontSize: 18 }} />
          {err && <div style={{ color: T.red, fontSize: 13, marginBottom: 8, fontWeight: 600 }}>{err}</div>}
          {ok && <div style={{ color: T.green, fontSize: 13, marginBottom: 8, fontWeight: 600 }}>{ok}</div>}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
            <Btn variant="ghost" onClick={onClose}>닫기</Btn>
            <Btn variant="red" onClick={resetByMaster}>초기화</Btn>
          </div>
        </>}
      </div>
    </div>
  );
}

// ── 설정 모달 ──────────────────────────────────────────────────
function SettingsModal({ settings, onSave, onClose }) {
  const [s, setS] = useState({ ...settings });
  const [gpsMsg, setGpsMsg] = useState("");
  const iStyle = { width: "100%", padding: "12px 14px", borderRadius: 12, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 18, fontWeight: 700, boxSizing: "border-box" };

  const registerOffice = async () => {
    setGpsMsg("위치 가져오는 중...");
    try {
      const gps = await getGPS();
      setS(p => ({ ...p, officeLat: gps.lat, officeLng: gps.lng }));
      setGpsMsg(`✓ 등록됨 (정확도 ±${gps.acc}m)`);
    } catch (e) {
      setGpsMsg("GPS 오류 — 위치 권한을 허용해주세요");
    }
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "#00000066", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, padding: 24 }}>
      <div style={{ background: T.card, borderRadius: 20, padding: 26, width: "100%", maxWidth: 320, boxShadow: "0 20px 60px #00000020" }}>
        <div style={{ fontWeight: 800, fontSize: 17, color: T.text, marginBottom: 20 }}>근무 설정</div>
        <div style={{ marginBottom: 16 }}>
          <div style={{ fontSize: 13, color: T.green, marginBottom: 6, fontWeight: 700 }}>출근 기준</div>
          <input type="time" value={s.workStart} onChange={e => setS(p => ({ ...p, workStart: e.target.value }))} style={{ ...iStyle, borderColor: T.green + "44" }} />
        </div>
        <div style={{ marginBottom: 20 }}>
          <div style={{ fontSize: 13, color: T.blue, marginBottom: 6, fontWeight: 700 }}>퇴근 기준</div>
          <input type="time" value={s.workEnd} onChange={e => setS(p => ({ ...p, workEnd: e.target.value }))} style={{ ...iStyle, borderColor: T.blue + "44" }} />
        </div>
        <div style={{ background: T.bg, border: `1px solid ${T.border}`, borderRadius: 14, padding: "14px 16px", marginBottom: 20 }}>
          <div style={{ fontSize: 13, color: T.text, fontWeight: 700, marginBottom: 4 }}>📍 회사 위치</div>
          <div style={{ fontSize: 11, color: T.muted, marginBottom: 10, lineHeight: 1.5 }}>
            회사에서 이 버튼을 눌러 위치를 등록하세요.<br />등록 후 팀원 출퇴근 위치를 확인할 수 있어요.
          </div>
          {s.officeLat && s.officeLng && (
            <div style={{ fontSize: 11, color: T.green, marginBottom: 8, fontWeight: 600 }}>
              ✓ 위치 등록됨 ({s.officeLat.toFixed(5)}, {s.officeLng.toFixed(5)})
            </div>
          )}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
            <button onClick={registerOffice} style={{ padding: "10px 0", borderRadius: 10, border: "none", background: T.adminHeader, color: "#fff", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>현재 위치 등록</button>
            <button onClick={() => { setS(p => ({ ...p, officeLat: null, officeLng: null })); setGpsMsg("위치가 삭제됐어요"); }} style={{ padding: "10px 0", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: T.muted, fontSize: 12, fontWeight: 600, cursor: "pointer" }}>위치 삭제</button>
          </div>
          {gpsMsg && <div style={{ fontSize: 11, color: gpsMsg.includes("✓") ? T.green : T.red, fontWeight: 600, marginBottom: 8 }}>{gpsMsg}</div>}
          <div style={{ fontSize: 11, color: T.muted, marginBottom: 4 }}>허용 반경</div>
          <select value={s.officeRadius || 200} onChange={e => setS(p => ({ ...p, officeRadius: Number(e.target.value) }))}
            style={{ width: "100%", padding: "8px 12px", borderRadius: 8, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 13, fontWeight: 600 }}>
            {[50, 100, 150, 200, 300, 500].map(r => <option key={r} value={r}>{r}m</option>)}
          </select>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          <Btn variant="ghost" onClick={onClose}>취소</Btn>
          <Btn variant="admin" onClick={() => onSave(s)}>저장</Btn>
        </div>
      </div>
    </div>
  );
}

// ── 팀원 관리 모달 ─────────────────────────────────────────────
function UserManageModal({ users, onSave, onClose }) {
  const [list, setList] = useState(users.map(u => ({ ...u })));
  const [mode, setMode] = useState("list");
  const [editing, setEditing] = useState(null);
  const [delTarget, setDelTarget] = useState(null);
  const [err, setErr] = useState("");
  const members = list.filter(u => u.role === "member");
  const iStyle = { width: "100%", padding: "12px 14px", borderRadius: 12, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 15, boxSizing: "border-box", marginBottom: 12, fontFamily: "inherit" };

  const applyAdd = () => {
    if (!editing.name.trim()) { setErr("이름을 입력해주세요"); return; }
    if (!/^\d{6}$/.test(editing.newPin)) { setErr("PIN은 숫자 6자리"); return; }
    if (editing.newPin !== editing.newPin2) { setErr("PIN이 일치하지 않아요"); return; }
    const u = { ...editing, pin: editing.newPin }; delete u.newPin; delete u.newPin2;
    setList(l => [...l, u]); setMode("list"); setErr("");
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "#00000066", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, padding: 20 }}>
      <div style={{ background: T.card, borderRadius: 20, padding: 22, width: "100%", maxWidth: 320, maxHeight: "88vh", overflowY: "auto", boxShadow: "0 20px 60px #00000020" }}>
        {mode === "list" && <>
          <div style={{ fontWeight: 800, fontSize: 17, color: T.text, marginBottom: 6 }}>팀원 관리</div>
          <div style={{ fontSize: 12, color: T.muted, marginBottom: 16 }}>팀원은 로그인 화면에서 직접 PIN을 변경할 수 있어요</div>
          {members.map(u => (
            <div key={u.id} style={{ display: "flex", alignItems: "center", gap: 10, background: T.bg, borderRadius: 12, padding: "12px 14px", marginBottom: 10, border: `1px solid ${T.border}` }}>
              <div style={{ width: 36, height: 36, borderRadius: "50%", background: T.headerBg, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14, fontWeight: 800, color: "#fff" }}>{u.name[0]}</div>
              <div style={{ flex: 1 }}>
                <div style={{ fontWeight: 700, color: T.text }}>{u.name}</div>
                <div style={{ fontSize: 11, color: T.muted }}>PIN 본인 관리</div>
              </div>
              <button onClick={() => { setEditing({ ...u }); setMode("edit"); }} style={{ background: T.card, border: `1px solid ${T.border}`, color: T.text, borderRadius: 8, padding: "5px 10px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>이름수정</button>
              <button onClick={() => { setDelTarget(u); setMode("delete"); }} style={{ background: T.redBg, border: "none", color: T.red, borderRadius: 8, padding: "5px 10px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>삭제</button>
            </div>
          ))}
          <button onClick={() => { setEditing({ id: "u" + Date.now(), name: "", newPin: "", newPin2: "", role: "member" }); setMode("add"); }} style={{ width: "100%", padding: "11px 0", borderRadius: 12, border: `2px dashed ${T.border}`, background: "none", color: T.muted, fontSize: 13, cursor: "pointer", marginBottom: 14, fontWeight: 600 }}>+ 팀원 추가</button>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
            <Btn variant="ghost" onClick={onClose}>취소</Btn>
            <Btn variant="admin" onClick={() => onSave(list)}>저장</Btn>
          </div>
        </>}
        {mode === "edit" && <>
          <button onClick={() => { setMode("list"); setErr(""); }} style={{ background: "none", border: "none", color: T.muted, cursor: "pointer", marginBottom: 14, fontSize: 13, fontWeight: 600 }}>← 뒤로</button>
          <div style={{ fontWeight: 800, fontSize: 16, color: T.text, marginBottom: 18 }}>이름 수정</div>
          <input value={editing.name} onChange={e => setEditing(p => ({ ...p, name: e.target.value }))} placeholder="이름" style={iStyle} />
          {err && <div style={{ color: T.red, fontSize: 13, marginBottom: 8, fontWeight: 600 }}>{err}</div>}
          <Btn variant="primary" onClick={() => {
            if (!editing.name.trim()) { setErr("이름을 입력해주세요"); return; }
            setList(l => l.map(x => x.id === editing.id ? { ...x, name: editing.name.trim() } : x)); setMode("list"); setErr("");
          }}>적용</Btn>
        </>}
        {mode === "add" && <>
          <button onClick={() => { setMode("list"); setErr(""); }} style={{ background: "none", border: "none", color: T.muted, cursor: "pointer", marginBottom: 14, fontSize: 13, fontWeight: 600 }}>← 뒤로</button>
          <div style={{ fontWeight: 800, fontSize: 16, color: T.text, marginBottom: 18 }}>팀원 추가</div>
          <input value={editing.name} onChange={e => setEditing(p => ({ ...p, name: e.target.value }))} placeholder="이름" style={iStyle} />
          <input type="password" inputMode="numeric" maxLength={6} value={editing.newPin} onChange={e => setEditing(p => ({ ...p, newPin: e.target.value.replace(/\D/g, "").slice(0, 6) }))} placeholder="PIN 6자리" style={{ ...iStyle, letterSpacing: 4 }} />
          <input type="password" inputMode="numeric" maxLength={6} value={editing.newPin2} onChange={e => setEditing(p => ({ ...p, newPin2: e.target.value.replace(/\D/g, "").slice(0, 6) }))} placeholder="PIN 확인" style={{ ...iStyle, letterSpacing: 4 }} />
          {err && <div style={{ color: T.red, fontSize: 13, marginBottom: 8, fontWeight: 600 }}>{err}</div>}
          <Btn variant="green" onClick={applyAdd}>추가</Btn>
        </>}
        {mode === "delete" && <>
          <div style={{ textAlign: "center", padding: "20px 0" }}>
            <div style={{ fontSize: 36, marginBottom: 14 }}>⚠️</div>
            <div style={{ fontWeight: 800, fontSize: 17, color: T.text, marginBottom: 10 }}>{delTarget?.name} 삭제</div>
            <div style={{ fontSize: 13, color: T.muted, lineHeight: 1.7 }}>출퇴근 기록도 함께 삭제됩니다.<br />계속할까요?</div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
            <Btn variant="ghost" onClick={() => { setMode("list"); setDelTarget(null); }}>취소</Btn>
            <Btn variant="red" onClick={() => { setList(l => l.filter(u => u.id !== delTarget.id)); setMode("list"); setDelTarget(null); }}>삭제</Btn>
          </div>
        </>}
      </div>
    </div>
  );
}

// ── 관리자 화면 ────────────────────────────────────────────────
function AdminScreen({ users, settings, records, leaves, onSaveRecord, onSaveLeave, onSaveUsers, onSaveSettings, onLogout }) {
  const [tab, setTab] = useState("today");
  const [showAccount, setShowAccount] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [showUsers, setShowUsers] = useState(false);
  const [editTarget, setEditTarget] = useState(null);
  const [now, setNow] = useState(new Date());
  useEffect(() => { const t = setInterval(() => setNow(new Date()), 60000); return () => clearInterval(t); }, []);
  const today = now.toISOString().slice(0, 10);
  const members = users.filter(u => u.role === "member");

  const handleSaveRecord = async (date, newRec, leaveData) => {
    await onSaveRecord(editTarget.user.id, date, newRec);
    await onSaveLeave(editTarget.user.id, date, leaveData);
    setEditTarget(null);
  };

  return (
    <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'Noto Sans KR',sans-serif", color: T.text }}>
      <div style={{ background: T.adminHeader, padding: "16px 16px 0" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
          <div>
            <div style={{ fontSize: 11, color: "#ffffff40", letterSpacing: 3 }}>ADMIN</div>
            <div style={{ fontSize: 20, fontWeight: 800, color: "#fff" }}>근태 현황</div>
          </div>
          <button onClick={onLogout} style={{ background: "#ffffff18", border: "none", color: "#fff", padding: "7px 14px", borderRadius: 18, fontSize: 12, cursor: "pointer", fontWeight: 600, whiteSpace: "nowrap" }}>로그아웃</button>
        </div>
        <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
          {[["👥 팀원", () => setShowUsers(true)], ["👤 계정", () => setShowAccount(true)], ["⚙ 설정", () => setShowSettings(true)]].map(([label, fn]) => (
            <button key={label} onClick={fn} style={{ background: "#ffffff14", border: "1px solid #ffffff20", color: "#fff", padding: "7px 14px", borderRadius: 20, fontSize: 12, cursor: "pointer", fontWeight: 600, whiteSpace: "nowrap" }}>{label}</button>
          ))}
        </div>
        <div style={{ fontSize: 11, color: "#ffffff40", marginBottom: 12 }}>출근 <strong style={{ color: "#4ade80" }}>{settings.workStart}</strong> · 퇴근 <strong style={{ color: "#60a5fa" }}>{settings.workEnd}</strong></div>
        <div style={{ display: "flex", borderBottom: "1px solid #ffffff18" }}>
          {[["today", "오늘"], ["month", "월별"]].map(([key, label]) => (
            <button key={key} onClick={() => setTab(key)} style={{ padding: "9px 20px", border: "none", background: "none", color: tab === key ? "#fff" : "#ffffff40", fontWeight: tab === key ? 800 : 400, fontSize: 14, cursor: "pointer", borderBottom: tab === key ? "2px solid #60a5fa" : "2px solid transparent", fontFamily: "inherit", whiteSpace: "nowrap" }}>
              {label}
            </button>
          ))}
        </div>
      </div>

      <div style={{ padding: 16 }}>
        {tab === "today" && <>
          <div style={{ fontSize: 13, color: T.muted, marginBottom: 12, fontWeight: 600 }}>
            {now.toLocaleDateString("ko-KR", { month: "long", day: "numeric", weekday: "long" })}
            {isWeekend(today) && <span style={{ marginLeft: 8 }}><Badge label="휴일" color="red" /></span>}
          </div>
          {members.map(u => {
            const rec = records[u.id]?.[today] || {};
            const status = !rec.in ? "미출근" : !rec.out ? "근무중" : "퇴근";
            const sColor = { 미출근: "gray", 근무중: "green", 퇴근: "blue" }[status];
            const lm = calcLateMin(rec.in, settings.workStart), em = calcEarlyOutMin(rec.out, settings.workEnd), om = calcTotalOvertimeMin(rec.in, rec.out, settings.workStart, settings.workEnd);
            const outings = rec.outing || [], isOut = outings.length > 0 && !outings[outings.length - 1].in;
            return (
              <div key={u.id} style={{ background: T.card, borderRadius: 16, padding: "14px 16px", marginBottom: 10, border: `1px solid ${T.border}`, boxShadow: "0 1px 4px #0000000a" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 10 }}>
                  <div style={{ width: 40, height: 40, borderRadius: "50%", background: T.adminHeader, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 15, fontWeight: 800, color: "#fff", flexShrink: 0 }}>{u.name[0]}</div>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 700, fontSize: 15, color: T.text }}>{u.name}</div>
                    <div style={{ fontSize: 12, color: T.muted, marginTop: 2 }}>
                      {rec.in && <span style={{ color: lm > 0 ? T.yellow : T.green, fontWeight: 600 }}>{formatTime(rec.in)}</span>}
                      {rec.in && rec.out && <span style={{ color: T.muted }}> → </span>}
                      {rec.out && <span style={{ color: em > 0 ? T.orange : T.blue, fontWeight: 600 }}>{formatTime(rec.out)}</span>}
                    </div>
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 6 }}>
                    <Badge label={status} color={sColor} />
                    <button onClick={() => setEditTarget({ user: u, date: today })} style={{ background: T.bg, border: `1px solid ${T.border}`, color: T.text, borderRadius: 8, padding: "4px 10px", fontSize: 11, cursor: "pointer", fontWeight: 600 }}>수정</button>
                  </div>
                </div>
                <div style={{ display: "flex", gap: 5, flexWrap: "wrap" }}>
                  {lm > 0 && rec.in && <Badge label={`지각 ${fmtMinutes(lm)}`} color="yellow" />}
                  {em > 0 && rec.out && <Badge label={`조퇴 ${fmtMinutes(em)}`} color="orange" />}
                  {isOut && <Badge label="외출중" color="blue" />}
                  {outings.filter(o => o.in).length > 0 && <Badge label={`외출 ${outings.filter(o => o.in).length}회`} color="gray" />}
                  {om >= 30 && <Badge label={`잔업 ${fmtMinutes(roundTo30(om))}`} color="purple" />}
                  {isWeekend(today) && rec.in && <Badge label="휴일근무" color="red" />}
                  {rec.note && <Badge label={`📝 ${rec.note}`} color="gray" />}
                  {rec.inGps && (() => { const s = gpsStatusLabel(rec.inGps, settings); return s ? <Badge label={`출근 ${s.label}`} color={s.color} /> : null; })()}
                  {rec.outGps && (() => { const s = gpsStatusLabel(rec.outGps, settings); return s ? <Badge label={`퇴근 ${s.label}`} color={s.color} /> : null; })()}
                </div>
              </div>
            );
          })}
        </>}
        {tab === "month" && <MonthTab records={records} leaves={leaves} members={members} settings={settings} onSaveRecord={onSaveRecord} onSaveLeave={onSaveLeave} />}
      </div>

      {showAccount && <AdminAccountModal users={users} onUpdateUsers={onSaveUsers} onClose={() => setShowAccount(false)} />}
      {editTarget && <EditRecordModal user={editTarget.user} date={editTarget.date} rec={records[editTarget.user.id]?.[editTarget.date] || {}} settings={settings} userLeaves={leaves[editTarget.user.id] || {}} onSave={handleSaveRecord} onClose={() => setEditTarget(null)} />}
      {showSettings && <SettingsModal settings={settings} onSave={async s => { await onSaveSettings(s); setShowSettings(false); }} onClose={() => setShowSettings(false)} />}
      {showUsers && <UserManageModal users={users} onSave={async u => { await onSaveUsers(u); setShowUsers(false); }} onClose={() => setShowUsers(false)} />}
    </div>
  );
}

// ── 메인 App ───────────────────────────────────────────────────
function App({ users, settings, records, leaves, onSaveUsers, onSaveSettings, onSaveRecord, onSaveLeave }) {
  const [user, setUser] = useState(null);
  if (!user) return <LoginScreen users={users} onLogin={setUser} onUpdateUsers={onSaveUsers} />;
  if (user.role === "admin") return (
    <AdminScreen users={users} settings={settings} records={records} leaves={leaves}
      onSaveRecord={onSaveRecord} onSaveLeave={onSaveLeave}
      onSaveUsers={onSaveUsers} onSaveSettings={onSaveSettings}
      onLogout={() => setUser(null)} />
  );
  return <MemberScreen user={user} settings={settings} records={records} leaves={leaves} onSaveRecord={onSaveRecord} onLogout={() => setUser(null)} />;
}

export default AppLoader;
