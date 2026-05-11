import { useState, useEffect } from "react";
import { db, storage } from "./firebase";
import {
  doc, onSnapshot, setDoc, getDoc, collection,
  getDocs, writeBatch, addDoc, deleteDoc, query, orderBy
} from "firebase/firestore";
import { ref, uploadBytes, getDownloadURL, deleteObject } from "firebase/storage";

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
const COL_NOTICES  = "notices";
const COL_BOARD    = "board";
const COL_PAYSLIPS = "payslips";
const COL_ANNUAL   = "annual";
const COL_LEAVE_REQ = "leave_requests";
const COL_MEMBER_INFO = "member_info"; // 팀원 임금 기초 데이터
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
  officeLat: null, officeLng: null, officeRadius: 200,
  holidays: []
};
const MASTER_CODE = "att2026!"; // 관리자 PIN 분실 시 비상 코드

// ── 유틸 ──────────────────────────────────────────────────────
function getToday() {
  const now = new Date();
  const kst = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  return kst.toISOString().slice(0, 10);
}
function formatTime(iso) {
  if (!iso) return "--:--";
  return new Date(iso).toLocaleTimeString("ko-KR", { hour: "2-digit", minute: "2-digit", hour12: false });
}
function formatDate(d) {
  return new Date(d).toLocaleDateString("ko-KR", { month: "long", day: "numeric", weekday: "short" });
}
function isWeekend(d) { const w = new Date(d).getDay(); return w === 0 || w === 6; }
function isHoliday(d, holidays) {
  return isWeekend(d) || (holidays || []).includes(d);
}
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
function calcMonthStats(days, settings, userLeaves, leaveRequests, userId, month) {
  const stats = days.reduce((acc, [date, rec]) => {
    if (rec.in) {
      acc.days++;
      const lm = calcLateMin(rec.in, settings.workStart);
      const em = calcEarlyOutMin(rec.out, settings.workEnd);
      const om = calcTotalOvertimeMin(rec.in, rec.out, settings.workStart, settings.workEnd);
      const finalLate = rec.lateConfirm !== undefined && rec.lateConfirm !== null ? rec.lateConfirm : lm > 0;
      const finalEarly = rec.earlyConfirm !== undefined && rec.earlyConfirm !== null ? rec.earlyConfirm : em > 0;
      const finalOt = rec.overtimeConfirm !== undefined && rec.overtimeConfirm !== null ? rec.overtimeConfirm : om >= 30;
      if (finalLate) { acc.late++; acc.lateMin += lm; }
      if (finalEarly) { acc.early++; acc.earlyMin += em; }
      if (finalOt) { acc.ot++; acc.otMin += roundTo30(om); }
      if (isHoliday(date, settings.holidays)) acc.holiday++;
    }
    return acc;
  }, { days: 0, late: 0, lateMin: 0, early: 0, earlyMin: 0, ot: 0, otMin: 0, holiday: 0, annualDays: 0 });

  // leaves 기반 연차/반차 집계
  if (userLeaves) {
    Object.entries(userLeaves)
      .filter(([date]) => !month || date.startsWith(month))
      .forEach(([date, l]) => {
        const dayRec = days.find(([d]) => d === date);
        const hasClockIn = !!(dayRec && dayRec[1]?.in); // 실제 출근 기록 있는지
        if (l.type === "연차") {
          stats.annualDays++;
          if (!hasClockIn) stats.days++; // 출근 찍은 기록 없는 연차일만 출근일수 추가
        } else if (l.type?.includes("반차")) {
          stats.annualDays += 0.5;
          if (!hasClockIn) stats.days++; // 반차도 출근 기록 없으면 출근 1일로
        }
      });
  }
  return stats;
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
  const kst = new Date(new Date().getTime() + 9 * 60 * 60 * 1000);
  const d = new Date(kst.getFullYear(), kst.getMonth() - i, 1);
  return new Date(d.getTime() + 9 * 60 * 60 * 1000).toISOString().slice(0, 7);
});
const monthLabel = m => { const [y, mo] = m.split("-"); return `${y}년 ${parseInt(mo)}월`; };

// ── GPS 유틸 ──────────────────────────────────────────────────
function getGPS() {
  return new Promise((resolve, reject) => {
    if (!navigator.geolocation) { reject(new Error("GPS 미지원")); return; }
    let attempts = 0;
    const tryGet = () => {
      navigator.geolocation.getCurrentPosition(
        p => {
          const lat = p.coords.latitude;
          const lng = p.coords.longitude;
          // 0,0 이거나 유효하지 않으면 재시도 (최대 3회)
          if ((Math.abs(lat) < 0.001 && Math.abs(lng) < 0.001) && attempts < 3) {
            attempts++;
            setTimeout(tryGet, 1000);
          } else {
            resolve({ lat, lng, acc: Math.round(p.coords.accuracy) });
          }
        },
        e => reject(e),
        { timeout: 10000, maximumAge: 0, enableHighAccuracy: true }
      );
    };
    tryGet();
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
  const lat = Number(gps.lat);
  const lng = Number(gps.lng);
  if (isNaN(lat) || isNaN(lng)) return null;
  if (Math.abs(lat) < 0.001 && Math.abs(lng) < 0.001) return null;
  if (!settings) return { label: "위치기록", color: "gray" };
  const officeLat = Number(settings.officeLat);
  const officeLng = Number(settings.officeLng);
  if (!settings.officeLat || !settings.officeLng || isNaN(officeLat) || isNaN(officeLng)) return { label: "위치기록", color: "gray" };
  const dist = calcDistance(lat, lng, officeLat, officeLng);
  if (dist == null || isNaN(dist) || dist > 100000) return { label: "위치기록", color: "gray" }; // 100km 이상이면 오류로 판단
  const radius = Number(settings.officeRadius) || 200;
  if (dist <= radius) return { label: `회사 내 (${dist}m)`, color: "green" };
  return { label: `회사 외 (${dist}m)`, color: "red" };
}
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
async function fbSaveUsers(newUsers, allUsers) {
  const batch = writeBatch(db);
  // 새 목록 저장
  newUsers.forEach(u => batch.set(doc(db, COL_USERS, u.id), u));
  // 삭제된 팀원 Firebase에서 제거
  if (allUsers) {
    const newIds = newUsers.map(u => u.id);
    allUsers.filter(u => !newIds.includes(u.id)).forEach(u => {
      batch.delete(doc(db, COL_USERS, u.id));
    });
  }
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
  const [notices, setNotices] = useState([]);
  const [board, setBoard] = useState([]);
  const [payslips, setPayslips] = useState([]);
  const [annual, setAnnual] = useState({});
  const [leaveRequests, setLeaveRequests] = useState([]);
  const [memberInfo, setMemberInfo] = useState({});

  useEffect(() => {
    let unsubs = [];

    // 유저
    unsubs.push(onSnapshot(collection(db, COL_USERS), snap => {
      if (snap.empty) { fbSaveUsers(DEFAULT_USERS); setUsers(DEFAULT_USERS); }
      else {
        const all = snap.docs.map(d => d.data());
        const admin = all.filter(u => u.role === "admin");
        const members = all.filter(u => u.role === "member").sort((a,b) => (a.createdAt||"").localeCompare(b.createdAt||""));
        setUsers([...admin, ...members]);
      }
    }));

    // 설정
    unsubs.push(onSnapshot(doc(db, "app", "settings"), snap => {
      if (snap.exists()) {
        const s = snap.data();
        if (s.officeLat != null) s.officeLat = Number(s.officeLat);
        if (s.officeLng != null) s.officeLng = Number(s.officeLng);
        if (s.officeRadius != null) s.officeRadius = Number(s.officeRadius);
        setSettings(s);
      } else fbSaveSettings(DEFAULT_SETTINGS);
    }));

    // 출퇴근 기록
    unsubs.push(onSnapshot(collection(db, COL_RECORDS), snap => {
      const r = {};
      snap.docs.forEach(d => {
        const data = d.data();
        if (!r[data.userId]) r[data.userId] = {};
        const { userId, date, ...rest } = data;
        r[data.userId][data.date] = rest;
      });
      setRecords(r);
    }));

    // 연차 기록
    unsubs.push(onSnapshot(collection(db, COL_LEAVES), snap => {
      const l = {};
      snap.docs.forEach(d => {
        const data = d.data();
        if (data.deleted === true) return; // deleted:true 면 무시
        if (!data.type) return; // type 없으면 무시 (잘못된 데이터)
        if (!l[data.userId]) l[data.userId] = {};
        const { userId, date, ...rest } = data;
        l[data.userId][data.date] = rest;
      });
      setLeaves(l);
    }));

    // 공지사항
    unsubs.push(onSnapshot(query(collection(db, COL_NOTICES), orderBy("createdAt", "desc")), snap => {
      setNotices(snap.docs.map(d => ({ id: d.id, ...d.data() })));
    }));

    // 자유게시판
    unsubs.push(onSnapshot(query(collection(db, COL_BOARD), orderBy("createdAt", "desc")), snap => {
      setBoard(snap.docs.map(d => ({ id: d.id, ...d.data() })));
    }));

    // 급여명세서
    unsubs.push(onSnapshot(query(collection(db, COL_PAYSLIPS), orderBy("createdAt", "desc")), snap => {
      setPayslips(snap.docs.map(d => ({ id: d.id, ...d.data() })));
    }));

    // 연차 현황
    unsubs.push(onSnapshot(collection(db, COL_ANNUAL), snap => {
      const a = {};
      snap.docs.forEach(d => { a[d.id] = d.data(); });
      setAnnual(a);
    }));

    // 연차 신청
    unsubs.push(onSnapshot(query(collection(db, COL_LEAVE_REQ), orderBy("createdAt", "desc")), snap => {
      setLeaveRequests(snap.docs.map(d => ({ id: d.id, ...d.data() })));
    }));

    // 팀원 기초 데이터
    unsubs.push(onSnapshot(collection(db, COL_MEMBER_INFO), snap => {
      const m = {};
      snap.docs.forEach(d => { m[d.id] = d.data(); });
      setMemberInfo(m);
    }));

    // 모든 구독 완료 후 ready
    setTimeout(() => setReady(true), 500);

    return () => unsubs.forEach(u => u());
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
    notices={notices} board={board} payslips={payslips} annual={annual} leaveRequests={leaveRequests} memberInfo={memberInfo}
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

  const today = new Date(now.getTime() + 9 * 60 * 60 * 1000).toISOString().slice(0, 10);
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

  const thisMonth = new Date(now.getTime() + 9 * 60 * 60 * 1000).toISOString().slice(0, 7);
  const monthDays = Object.entries(records[user.id] || {}).filter(([d]) => d.startsWith(thisMonth)).sort(([a], [b]) => b.localeCompare(a));
  const monthLeaves = Object.entries(leaves[user.id] || {}).filter(([d]) => d.startsWith(thisMonth));
  const monthLeavesObj = Object.fromEntries(monthLeaves); // 월 필터된 연차만
  const ms = calcMonthStats(monthDays, settings, monthLeavesObj, null, user.id, thisMonth);
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
          <div style={{ fontSize: 10, color: "#ffffff30", marginTop: 2 }}>GPS: {String(settings.officeLat)} | today: {today}</div>
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
            [["연차", ms.annualDays > 0 ? ms.annualDays + "일" : "0일", "#7c3aed"], ["조퇴", ms.early + "회", T.orange], ["조퇴시간", fmtMinutes(ms.earlyMin), T.orange]],
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
            const weekend = isHoliday(date, settings.holidays), leave = leaves[user.id]?.[date];
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

  // 관리자 확정값 (null이면 자동, true/false면 수동 확정)
  const [lateConfirm, setLateConfirm] = useState(rec?.lateConfirm ?? null);
  const [earlyConfirm, setEarlyConfirm] = useState(rec?.earlyConfirm ?? null);
  const [overtimeConfirm, setOvertimeConfirm] = useState(rec?.overtimeConfirm ?? null);

  const iStyle = { width: "100%", padding: "11px 14px", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 14, fontWeight: 600, boxSizing: "border-box", fontFamily: "inherit" };
  const inIso = inTime ? setTimeOnDate(date, inTime) : null;
  const outIso = outTime ? setTimeOnDate(date, outTime) : null;
  const autoLate = calcLateMin(inIso, settings.workStart) > 0;
  const autoEarly = calcEarlyOutMin(outIso, settings.workEnd) > 0;
  const autoOvertime = calcTotalOvertimeMin(inIso, outIso, settings.workStart, settings.workEnd) >= 30;
  const lm = calcLateMin(inIso, settings.workStart);
  const em = calcEarlyOutMin(outIso, settings.workEnd);
  const om = calcTotalOvertimeMin(inIso, outIso, settings.workStart, settings.workEnd);

  // 최종 확정값 (관리자가 수동 설정했으면 그걸, 아니면 자동)
  const finalLate = lateConfirm !== null ? lateConfirm : autoLate;
  const finalEarly = earlyConfirm !== null ? earlyConfirm : autoEarly;
  const finalOvertime = overtimeConfirm !== null ? overtimeConfirm : autoOvertime;

  const ToggleBtn = ({ label, auto, confirmed, onToggle, color }) => {
    const isOn = confirmed !== null ? confirmed : auto;
    return (
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "10px 14px", background: T.bg, borderRadius: 10, marginBottom: 8, border: `1px solid ${T.border}` }}>
        <div>
          <div style={{ fontSize: 13, fontWeight: 700, color: isOn ? color : T.muted }}>{label}</div>
          <div style={{ fontSize: 11, color: T.muted, marginTop: 2 }}>
            자동감지: {auto ? "해당" : "없음"} {confirmed !== null ? "· 수동확정" : ""}
          </div>
        </div>
        <div style={{ display: "flex", gap: 6 }}>
          <button onClick={() => onToggle(true)}
            style={{ padding: "6px 12px", borderRadius: 8, border: "none", fontSize: 12, fontWeight: 700, cursor: "pointer", background: isOn ? color : T.border, color: isOn ? "#fff" : T.muted }}>
            인정
          </button>
          <button onClick={() => onToggle(false)}
            style={{ padding: "6px 12px", borderRadius: 8, border: "none", fontSize: 12, fontWeight: 700, cursor: "pointer", background: !isOn ? "#6b7280" : T.border, color: !isOn ? "#fff" : T.muted }}>
            제외
          </button>
          {confirmed !== null && (
            <button onClick={() => onToggle(null)}
              style={{ padding: "6px 10px", borderRadius: 8, border: `1px solid ${T.border}`, fontSize: 11, cursor: "pointer", background: "#fff", color: T.muted }}>
              자동
            </button>
          )}
        </div>
      </div>
    );
  };

  const save = async () => {
    const nr = { ...rec, outing: outings, note, lateConfirm, earlyConfirm, overtimeConfirm };
    if (inTime) nr.in = setTimeOnDate(date, inTime); else delete nr.in;
    if (outTime) nr.out = setTimeOnDate(date, outTime); else delete nr.out;
    await onSave(date, nr, leaveType ? { type: leaveType, hours: leaveHours } : null);
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "#00000066", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, padding: 16 }}>
      <div style={{ background: T.card, borderRadius: 20, padding: 22, width: "100%", maxWidth: 340, maxHeight: "90vh", overflowY: "auto", boxShadow: "0 20px 60px #00000030" }}>
        <div style={{ fontWeight: 800, fontSize: 17, color: T.text, marginBottom: 4 }}>{user.name} — 기록 수정</div>
        <div style={{ fontSize: 12, color: T.muted, marginBottom: 18 }}>{formatDate(date)}</div>

        {/* 실제 출퇴근 시간 */}
        <div style={{ fontSize: 12, color: T.sub, marginBottom: 8, fontWeight: 700 }}>실제 출퇴근 시간</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 16 }}>
          {[["출근", inTime, setInTime, T.green], ["퇴근", outTime, setOutTime, T.blue]].map(([label, val, setter, color]) => (
            <div key={label}>
              <div style={{ fontSize: 12, color, marginBottom: 5, fontWeight: 700 }}>{label}</div>
              <input type="time" value={val} onChange={e => setter(e.target.value)}
                style={{ ...iStyle, color: val ? color : T.muted, fontSize: 16 }} />
            </div>
          ))}
        </div>

        {/* 관리자 확정 */}
        {inTime && outTime && (
          <>
            <div style={{ fontSize: 12, color: T.sub, marginBottom: 8, fontWeight: 700 }}>급여 반영 확정</div>
            <ToggleBtn label={`지각 ${lm > 0 ? fmtMinutes(lm) : ""}`} auto={autoLate} confirmed={lateConfirm} onToggle={v => setLateConfirm(v)} color={T.yellow} />
            <ToggleBtn label={`조퇴 ${em > 0 ? fmtMinutes(em) : ""}`} auto={autoEarly} confirmed={earlyConfirm} onToggle={v => setEarlyConfirm(v)} color={T.orange} />
            <ToggleBtn label={`잔업 ${om >= 30 ? fmtMinutes(roundTo30(om)) : ""}`} auto={autoOvertime} confirmed={overtimeConfirm} onToggle={v => setOvertimeConfirm(v)} color={T.purple} />
            <div style={{ fontSize: 11, color: T.muted, marginBottom: 12, padding: "6px 10px", background: T.bg, borderRadius: 8 }}>
              💡 "제외" 선택 시 실제 시간은 보존되고 급여 계산에서만 제외돼요
            </div>
          </>
        )}

        {/* 외출 */}
        {outings.map((o, i) => (
          <div key={i} style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 6 }}>
            <input type="time" value={o.out ? formatTime(o.out) : ""}
              onChange={e => { const n=[...outings]; n[i]={...n[i],out:e.target.value?setTimeOnDate(date,e.target.value):null}; setOutings(n); }}
              placeholder="외출" style={{ ...iStyle, flex: 1, fontSize: 13 }} />
            <span style={{ color: T.muted, fontSize: 12 }}>→</span>
            <input type="time" value={o.in ? formatTime(o.in) : ""}
              onChange={e => { const n=[...outings]; n[i]={...n[i],in:e.target.value?setTimeOnDate(date,e.target.value):null}; setOutings(n); }}
              placeholder="복귀" style={{ ...iStyle, flex: 1, fontSize: 13 }} />
            <button onClick={() => setOutings(outings.filter((_, ii) => ii !== i))}
              style={{ background: T.redBg, border: "none", color: T.red, borderRadius: 8, padding: "8px 10px", cursor: "pointer", fontWeight: 700 }}>✕</button>
          </div>
        ))}
        <button onClick={() => setOutings([...outings, { out: null, in: null }])} style={{ width: "100%", padding: "8px 0", borderRadius: 10, border: `1px dashed ${T.border}`, background: "none", color: T.muted, fontSize: 12, cursor: "pointer", marginBottom: 12, fontWeight: 600 }}>+ 외출 추가</button>

        {/* 연차 */}
        <div style={{ marginBottom: 12 }}>
          <div style={{ fontSize: 12, color: T.purple, marginBottom: 5, fontWeight: 700 }}>연차 / 반차</div>
          <select value={leaveType} onChange={e => setLeaveType(e.target.value)} style={iStyle}>
            <option value="">해당없음</option>
            {LEAVE_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
          </select>
          {leaveType === "시간연차" && <select value={leaveHours} onChange={e => setLeaveHours(Number(e.target.value))} style={{ ...iStyle, marginTop: 6 }}>{[1, 2, 3, 4].map(h => <option key={h} value={h}>{h}시간</option>)}</select>}
        </div>

        {/* 메모 */}
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
function MonthTab({ records, leaves, members, settings, leaveRequests, onSaveRecord, onSaveLeave }) {
  const kstNow = new Date(new Date().getTime() + 9 * 60 * 60 * 1000);
  const currentMonth = kstNow.toISOString().slice(0, 7);
  const [selectedMonth, setSelectedMonth] = useState(currentMonth);
  const [drillUser, setDrillUser] = useState(null);
  const [editTarget, setEditTarget] = useState(null);

  const prevMonth = () => {
    const [y, m] = selectedMonth.split("-").map(Number);
    const d = new Date(y, m - 2, 1);
    setSelectedMonth(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`);
    setDrillUser(null);
  };
  const nextMonth = () => {
    const [y, m] = selectedMonth.split("-").map(Number);
    const d = new Date(y, m, 1);
    const next = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    if (next <= currentMonth) { setSelectedMonth(next); setDrillUser(null); }
  };
  const isCurrentMonth = selectedMonth === currentMonth;

  const handleSaveRecord = async (date, newRec, leaveData) => {
    await onSaveRecord(editTarget.user.id, date, newRec);
    await onSaveLeave(editTarget.user.id, date, leaveData);
    setEditTarget(null);
  };

  if (drillUser) {
    const days = Object.entries(records[drillUser.id] || {}).filter(([d]) => d.startsWith(selectedMonth)).sort(([a], [b]) => a.localeCompare(b));
    const userLeaves = leaves[drillUser.id] || {};
    const mLeaves = Object.entries(userLeaves).filter(([d]) => d.startsWith(selectedMonth));
    const mLeavesObj = Object.fromEntries(mLeaves); // 월 필터된 연차만
    const ms = calcMonthStats(days, settings, mLeavesObj, leaveRequests, drillUser.id, selectedMonth);

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
        </div>
        {[
          [["출근", ms.days + "일", T.green], ["지각", ms.late + "회", T.yellow], ["지각시간", fmtMinutes(ms.lateMin), T.yellow]],
          [["휴일", ms.holiday + "일", T.red], ["잔업", ms.ot + "일", T.purple], ["잔업시간", fmtMinutes(ms.otMin), T.purple]],
          [["연차", ms.annualDays > 0 ? ms.annualDays + "일" : "0일", "#7c3aed"], ["조퇴", ms.early + "회", T.orange], ["조퇴시간", fmtMinutes(ms.earlyMin), T.orange]],
        ].map((row, ri) => (
          <div key={ri} style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 6, marginBottom: ri < 2 ? 6 : 12 }}>
            {row.map(([l, v, c]) => <StatBox key={l} label={l} value={v} color={c} />)}
          </div>
        ))}
        <div style={{ fontSize: 12, color: T.muted, marginBottom: 10, fontWeight: 600 }}>날짜별 상세</div>
        {(() => {
          // 출근 기록 + 연차만 있는 날짜 합치기
          const leaveDates = Object.entries(userLeaves)
            .filter(([date]) => date.startsWith(selectedMonth) && !days.find(([d]) => d === date))
            .map(([date, l]) => [date, { in: null, out: null, leaveOnly: true }]);
          const allDays = [...days, ...leaveDates].sort(([a], [b]) => a.localeCompare(b));
          if (allDays.length === 0) return <div style={{ textAlign: "center", color: T.muted, padding: 24, fontSize: 14, background: T.card, borderRadius: 12, border: `1px solid ${T.border}` }}>기록 없음</div>;
          return allDays.map(([date, rec]) => {
            const lm = calcLateMin(rec.in, settings.workStart), em = calcEarlyOutMin(rec.out, settings.workEnd), om = calcTotalOvertimeMin(rec.in, rec.out, settings.workStart, settings.workEnd);
            const late = lm > 0, early = em > 0, ot = om >= 30, weekend = isHoliday(date, settings.holidays), leave = userLeaves[date];
            const [, , dd] = date.split("-"), dow = new Date(date).toLocaleDateString("ko-KR", { weekday: "short" });
            const isNormal = !late && !early && !ot && !weekend && rec.in && rec.out && !leave;
            return (
              <div key={date} style={{ background: T.card, borderRadius: 12, padding: "12px 14px", marginBottom: 8, border: `1px solid ${leave && !rec.in ? T.purpleBg : T.border}` }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <div style={{ minWidth: 36, textAlign: "center" }}>
                    <div style={{ fontSize: 17, fontWeight: 800, color: weekend ? T.red : T.text }}>{parseInt(dd)}</div>
                    <div style={{ fontSize: 10, color: T.muted }}>{dow}</div>
                  </div>
                  <div style={{ width: 1, height: 34, background: T.border }} />
                  <div style={{ flex: 1 }}>
                    {rec.in ? (
                      <div style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 4 }}>
                        <span style={{ fontSize: 13, fontWeight: 700, color: late ? T.yellow : T.green }}>{formatTime(rec.in)}</span>
                        <span style={{ fontSize: 11, color: T.muted }}>→</span>
                        <span style={{ fontSize: 13, fontWeight: 700, color: early ? T.orange : T.blue }}>{formatTime(rec.out)}</span>
                        {isNormal && <span style={{ marginLeft: "auto", fontSize: 11, color: T.muted }}>정상</span>}
                      </div>
                    ) : (
                      <div style={{ fontSize: 13, color: T.purple, fontWeight: 700, marginBottom: 4 }}>출근 기록 없음</div>
                    )}
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
          });
        })()}
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
      {/* 월 선택 - 버튼 방식 */}
      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 12 }}>
        <button onClick={prevMonth} style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 10, padding: "10px 16px", fontSize: 16, cursor: "pointer", fontWeight: 700, color: T.text }}>‹</button>
        <div style={{ flex: 1, textAlign: "center", fontSize: 16, fontWeight: 800, color: T.text }}>{monthLabel(selectedMonth)}</div>
        <button onClick={nextMonth} style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 10, padding: "10px 16px", fontSize: 16, cursor: "pointer", fontWeight: 700, color: isCurrentMonth ? T.muted : T.text, opacity: isCurrentMonth ? 0.3 : 1 }}>›</button>
      </div>
      {/* 전체 직원 CSV 다운로드 */}
      <button onClick={() => {
        const header = ["이름", "날짜", "요일", "출근", "퇴근", "지각", "지각시간", "조퇴", "조퇴시간", "잔업", "잔업시간", "외출", "연차/반차", "메모"];
        const rows = [];
        members.forEach(u => {
          const days = Object.entries(records[u.id] || {}).filter(([d]) => d.startsWith(selectedMonth)).sort(([a],[b])=>a.localeCompare(b));
          const userLeaves = leaves[u.id] || {};
          days.forEach(([date, rec]) => {
            const dow = new Date(date).toLocaleDateString("ko-KR", { weekday: "short" });
            const lm = calcLateMin(rec.in, settings.workStart);
            const em = calcEarlyOutMin(rec.out, settings.workEnd);
            const om = calcTotalOvertimeMin(rec.in, rec.out, settings.workStart, settings.workEnd);
            const leave = userLeaves[date];
            const finalLate = rec.lateConfirm !== undefined && rec.lateConfirm !== null ? rec.lateConfirm : lm > 0;
            const finalEarly = rec.earlyConfirm !== undefined && rec.earlyConfirm !== null ? rec.earlyConfirm : em > 0;
            const finalOt = rec.overtimeConfirm !== undefined && rec.overtimeConfirm !== null ? rec.overtimeConfirm : om >= 30;
            rows.push([u.name, date, dow, formatTime(rec.in), formatTime(rec.out),
              finalLate?"O":"", finalLate?fmtMinutes(lm):"",
              finalEarly?"O":"", finalEarly?fmtMinutes(em):"",
              finalOt?"O":"", finalOt?fmtMinutes(roundTo30(om)):"",
              (rec.outing||[]).length>0?`${(rec.outing||[]).length}회`:"",
              leave?leave.type:"", rec.note||""]);
          });
        });
        downloadCSV(`전체직원_${monthLabel(selectedMonth)}_근태.csv`, [header, ...rows]);
      }} style={{ width: "100%", padding: "11px 0", borderRadius: 12, border: "none", background: T.green, color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer", marginBottom: 14 }}>
        ⬇ 전체 직원 CSV 다운로드
      </button>
      {members.map(u => {
        const days = Object.entries(records[u.id] || {}).filter(([d]) => d.startsWith(selectedMonth));
        const uLeavesObj = Object.fromEntries(Object.entries(leaves[u.id] || {}).filter(([d]) => d.startsWith(selectedMonth)));
        const ms = calcMonthStats(days, settings, uLeavesObj, leaveRequests, u.id, selectedMonth);
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
  const [s, setS] = useState({ ...settings, holidays: settings.holidays || [] });
  const [gpsMsg, setGpsMsg] = useState("");
  const [newHoliday, setNewHoliday] = useState("");
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

  const addHoliday = () => {
    if (!newHoliday) return;
    if (s.holidays.includes(newHoliday)) { setNewHoliday(""); return; }
    setS(p => ({ ...p, holidays: [...p.holidays, newHoliday].sort() }));
    setNewHoliday("");
  };

  const removeHoliday = (date) => {
    setS(p => ({ ...p, holidays: p.holidays.filter(d => d !== date) }));
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "#00000066", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, padding: 16 }}>
      <div style={{ background: T.card, borderRadius: 20, padding: 22, width: "100%", maxWidth: 340, maxHeight: "90vh", overflowY: "auto", boxShadow: "0 20px 60px #00000020" }}>
        <div style={{ fontWeight: 800, fontSize: 17, color: T.text, marginBottom: 20 }}>근무 설정</div>

        <div style={{ marginBottom: 16 }}>
          <div style={{ fontSize: 13, color: T.green, marginBottom: 6, fontWeight: 700 }}>출근 기준</div>
          <input type="time" value={s.workStart} onChange={e => setS(p => ({ ...p, workStart: e.target.value }))} style={{ ...iStyle, borderColor: T.green + "44" }} />
        </div>
        <div style={{ marginBottom: 20 }}>
          <div style={{ fontSize: 13, color: T.blue, marginBottom: 6, fontWeight: 700 }}>퇴근 기준</div>
          <input type="time" value={s.workEnd} onChange={e => setS(p => ({ ...p, workEnd: e.target.value }))} style={{ ...iStyle, borderColor: T.blue + "44" }} />
        </div>

        {/* 공휴일 관리 */}
        <div style={{ background: T.bg, border: `1px solid ${T.border}`, borderRadius: 14, padding: "14px 16px", marginBottom: 20 }}>
          <div style={{ fontSize: 13, color: T.text, fontWeight: 700, marginBottom: 4 }}>🗓 공휴일 지정</div>
          <div style={{ fontSize: 11, color: T.muted, marginBottom: 12, lineHeight: 1.5 }}>
            토/일 외 공휴일을 직접 등록하세요.<br />등록된 날은 휴일근무로 자동 처리돼요.
          </div>
          <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
            <input type="date" value={newHoliday} onChange={e => setNewHoliday(e.target.value)}
              style={{ flex: 1, padding: "10px 12px", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 14, fontWeight: 600, boxSizing: "border-box" }} />
            <button onClick={addHoliday}
              style={{ background: T.adminHeader, border: "none", color: "#fff", borderRadius: 10, padding: "10px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer", whiteSpace: "nowrap" }}>추가</button>
          </div>
          {s.holidays.length === 0 ? (
            <div style={{ fontSize: 12, color: T.muted, textAlign: "center", padding: "8px 0" }}>등록된 공휴일 없음</div>
          ) : s.holidays.map(date => (
            <div key={date} style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 0", borderTop: `1px solid ${T.border}` }}>
              <span style={{ flex: 1, fontSize: 13, fontWeight: 600, color: T.text }}>{formatDate(date)}</span>
              <button onClick={() => removeHoliday(date)}
                style={{ background: T.redBg, border: "none", color: T.red, borderRadius: 8, padding: "4px 10px", fontSize: 12, cursor: "pointer", fontWeight: 700 }}>삭제</button>
            </div>
          ))}
        </div>

        {/* 회사 위치 */}
        <div style={{ background: T.bg, border: `1px solid ${T.border}`, borderRadius: 14, padding: "14px 16px", marginBottom: 20 }}>
          <div style={{ fontSize: 13, color: T.text, fontWeight: 700, marginBottom: 4 }}>📍 회사 위치</div>
          <div style={{ fontSize: 11, color: T.muted, marginBottom: 10, lineHeight: 1.5 }}>
            회사에서 이 버튼을 눌러 위치를 등록하세요.
          </div>
          {s.officeLat && s.officeLng && (
            <div style={{ fontSize: 11, color: T.green, marginBottom: 8, fontWeight: 600 }}>
              ✓ 위치 등록됨 ({Number(s.officeLat).toFixed(5)}, {Number(s.officeLng).toFixed(5)})
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
    const u = { ...editing, pin: editing.newPin, createdAt: new Date().toISOString() };
    delete u.newPin; delete u.newPin2;
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
// ── 관리자 대문 ────────────────────────────────────────────────
function AdminHome({ user, onLogout, onSection }) {
  const now = new Date();
  const kst = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  const dateStr = kst.toLocaleDateString("ko-KR", { year: "numeric", month: "long", day: "numeric", weekday: "long" });

  const sections = [
    { key: "attendance", icon: "📋", label: "근태", desc: "출퇴근 현황 · 월별 기록", color: "#2563eb" },
    { key: "wage",       icon: "💰", label: "임금", desc: "급여 계산 · 임금대장",   color: "#16a34a" },
    { key: "members",    icon: "👥", label: "팀원", desc: "직원 정보 · 기초 데이터", color: "#7c3aed" },
    { key: "general",    icon: "⚙",  label: "일반", desc: "설정 · 공지 · 게시판",   color: "#ea580c" },
  ];

  return (
    <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'Noto Sans KR',sans-serif" }}>
      {/* 헤더 */}
      <div style={{ background: T.adminHeader, padding: "24px 20px 28px" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 20 }}>
          <div>
            <div style={{ fontSize: 11, color: "#ffffff40", letterSpacing: 3, marginBottom: 4 }}>ADMIN</div>
            <div style={{ fontSize: 22, fontWeight: 800, color: "#fff" }}>{user.name}님</div>
            <div style={{ fontSize: 12, color: "#ffffff50", marginTop: 4 }}>{dateStr}</div>
          </div>
          <button onClick={onLogout} style={{ background: "#ffffff18", border: "none", color: "#fff", padding: "8px 16px", borderRadius: 20, fontSize: 12, cursor: "pointer", fontWeight: 600 }}>로그아웃</button>
        </div>
      </div>

      {/* 섹션 버튼 */}
      <div style={{ padding: 20, marginTop: -14 }}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
          {sections.map(s => (
            <button key={s.key} onClick={() => onSection(s.key)}
              style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 20, padding: "24px 16px", cursor: "pointer", textAlign: "left", boxShadow: "0 2px 12px #0000000d", transition: "transform .1s" }}
              onMouseDown={e => e.currentTarget.style.transform = "scale(0.97)"}
              onMouseUp={e => e.currentTarget.style.transform = "scale(1)"}
              onTouchStart={e => e.currentTarget.style.transform = "scale(0.97)"}
              onTouchEnd={e => e.currentTarget.style.transform = "scale(1)"}>
              <div style={{ fontSize: 32, marginBottom: 10 }}>{s.icon}</div>
              <div style={{ fontSize: 18, fontWeight: 800, color: s.color, marginBottom: 4 }}>{s.label}</div>
              <div style={{ fontSize: 12, color: T.muted, lineHeight: 1.5 }}>{s.desc}</div>
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}

// ── 관리자 근태 섹션 ───────────────────────────────────────────
function AdminAttendance({ users, settings, records, leaves, leaveRequests, onSaveRecord, onSaveLeave, onSaveSettings, onBack }) {
  const [tab, setTab] = useState("today");
  const [showSettings, setShowSettings] = useState(false);
  const [editTarget, setEditTarget] = useState(null);
  const [now, setNow] = useState(new Date());
  useEffect(() => { const t = setInterval(() => setNow(new Date()), 60000); return () => clearInterval(t); }, []);
  const today = new Date(now.getTime() + 9 * 60 * 60 * 1000).toISOString().slice(0, 10);
  const members = users.filter(u => u.role === "member");

  const handleSaveRecord = async (date, newRec, leaveData) => {
    await onSaveRecord(editTarget.user.id, date, newRec);
    await onSaveLeave(editTarget.user.id, date, leaveData);
    setEditTarget(null);
  };

  return (
    <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'Noto Sans KR',sans-serif" }}>
      <div style={{ background: T.adminHeader, padding: "16px 16px 0" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <button onClick={onBack} style={{ background: "#ffffff18", border: "none", color: "#fff", fontSize: 18, cursor: "pointer", padding: "8px 14px", borderRadius: 12, fontWeight: 700 }}>‹</button>
            <div>
              <div style={{ fontSize: 11, color: "#ffffff40", letterSpacing: 3 }}>ADMIN</div>
              <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>📋 근태</div>
            </div>
          </div>
          <button onClick={() => setShowSettings(true)} style={{ background: "#ffffff18", border: "none", color: "#fff", padding: "7px 14px", borderRadius: 18, fontSize: 12, cursor: "pointer", fontWeight: 600 }}>⚙ 설정</button>
        </div>
        <div style={{ fontSize: 11, color: "#ffffff40", marginBottom: 12 }}>
          출근 <strong style={{ color: "#4ade80" }}>{settings.workStart}</strong> · 퇴근 <strong style={{ color: "#60a5fa" }}>{settings.workEnd}</strong>
        </div>
        <div style={{ display: "flex", borderBottom: "1px solid #ffffff18" }}>
          {[["today", "오늘"], ["month", "월별"]].map(([key, label]) => (
            <button key={key} onClick={() => setTab(key)}
              style={{ padding: "9px 20px", border: "none", background: "none", color: tab === key ? "#fff" : "#ffffff40", fontWeight: tab === key ? 800 : 400, fontSize: 14, cursor: "pointer", borderBottom: tab === key ? "2px solid #60a5fa" : "2px solid transparent", fontFamily: "inherit" }}>
              {label}
            </button>
          ))}
        </div>
      </div>

      <div style={{ padding: 16 }}>
        {tab === "today" && <>
          <div style={{ fontSize: 13, color: T.muted, marginBottom: 12, fontWeight: 600 }}>
            {now.toLocaleDateString("ko-KR", { month: "long", day: "numeric", weekday: "long" })}
            {isHoliday(today, settings.holidays) && <span style={{ marginLeft: 8 }}><Badge label="휴일" color="red" /></span>}
          </div>
          {members.map(u => {
            const rec = records[u.id]?.[today] || {};
            const status = !rec.in ? "미출근" : !rec.out ? "근무중" : "퇴근";
            const sColor = { 미출근: "gray", 근무중: "green", 퇴근: "blue" }[status];
            const lm = calcLateMin(rec.in, settings.workStart), em = calcEarlyOutMin(rec.out, settings.workEnd), om = calcTotalOvertimeMin(rec.in, rec.out, settings.workStart, settings.workEnd);
            const outings = rec.outing || [], isOut = outings.length > 0 && !outings[outings.length - 1].in;
            const finalLate = rec.lateConfirm !== undefined && rec.lateConfirm !== null ? rec.lateConfirm : lm > 0;
            const finalEarly = rec.earlyConfirm !== undefined && rec.earlyConfirm !== null ? rec.earlyConfirm : em > 0;
            return (
              <div key={u.id} style={{ background: T.card, borderRadius: 16, padding: "14px 16px", marginBottom: 10, border: `1px solid ${T.border}`, boxShadow: "0 1px 4px #0000000a" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 10 }}>
                  <div style={{ width: 40, height: 40, borderRadius: "50%", background: T.adminHeader, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 15, fontWeight: 800, color: "#fff", flexShrink: 0 }}>{u.name[0]}</div>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 700, fontSize: 15, color: T.text }}>{u.name}</div>
                    <div style={{ fontSize: 12, color: T.muted, marginTop: 2 }}>
                      {rec.in && <span style={{ color: finalLate ? T.yellow : T.green, fontWeight: 600 }}>{formatTime(rec.in)}</span>}
                      {rec.in && rec.out && <span style={{ color: T.muted }}> → </span>}
                      {rec.out && <span style={{ color: finalEarly ? T.orange : T.blue, fontWeight: 600 }}>{formatTime(rec.out)}</span>}
                    </div>
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 6 }}>
                    <Badge label={status} color={sColor} />
                    <button onClick={() => setEditTarget({ user: u, date: today })} style={{ background: T.bg, border: `1px solid ${T.border}`, color: T.text, borderRadius: 8, padding: "4px 10px", fontSize: 11, cursor: "pointer", fontWeight: 600 }}>수정</button>
                  </div>
                </div>
                <div style={{ display: "flex", gap: 5, flexWrap: "wrap" }}>
                  {finalLate && rec.in && <Badge label={`지각 ${fmtMinutes(lm)}`} color="yellow" />}
                  {finalEarly && rec.out && <Badge label={`조퇴 ${fmtMinutes(em)}`} color="orange" />}
                  {isOut && <Badge label="외출중" color="blue" />}
                  {outings.filter(o => o.in).length > 0 && <Badge label={`외출 ${outings.filter(o => o.in).length}회`} color="gray" />}
                  {om >= 30 && <Badge label={`잔업 ${fmtMinutes(roundTo30(om))}`} color="purple" />}
                  {isHoliday(today, settings.holidays) && rec.in && <Badge label="휴일근무" color="red" />}
                  {rec.note && <Badge label={`📝 ${rec.note}`} color="gray" />}
                  {rec.inGps && (() => { const s = gpsStatusLabel(rec.inGps, settings); return s ? <Badge label={`출근 ${s.label}`} color={s.color} /> : null; })()}
                  {rec.outGps && (() => { const s = gpsStatusLabel(rec.outGps, settings); return s ? <Badge label={`퇴근 ${s.label}`} color={s.color} /> : null; })()}
                </div>
              </div>
            );
          })}
        </>}
        {tab === "month" && <MonthTab records={records} leaves={leaves} members={members} settings={settings} leaveRequests={leaveRequests} onSaveRecord={onSaveRecord} onSaveLeave={onSaveLeave} />}
      </div>

      {editTarget && <EditRecordModal user={editTarget.user} date={editTarget.date} rec={records[editTarget.user.id]?.[editTarget.date] || {}} settings={settings} userLeaves={leaves[editTarget.user.id] || {}} onSave={handleSaveRecord} onClose={() => setEditTarget(null)} />}
      {showSettings && <SettingsModal settings={settings} onSave={async s => { await onSaveSettings(s); setShowSettings(false); }} onClose={() => setShowSettings(false)} />}
    </div>
  );
}

// ── 관리자 임금 섹션 (추후 개발) ──────────────────────────────
function AdminWage({ users, records, settings, onBack }) {
  const members = users.filter(u => u.role === "member");
  return (
    <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'Noto Sans KR',sans-serif" }}>
      <div style={{ background: "#16a34a", padding: "16px 16px 20px" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <button onClick={onBack} style={{ background: "#ffffff18", border: "none", color: "#fff", fontSize: 18, cursor: "pointer", padding: "8px 14px", borderRadius: 12, fontWeight: 700 }}>‹</button>
          <div>
            <div style={{ fontSize: 11, color: "#ffffff40", letterSpacing: 3 }}>ADMIN</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>💰 임금</div>
          </div>
        </div>
      </div>
      <div style={{ padding: 40, textAlign: "center" }}>
        <div style={{ fontSize: 48, marginBottom: 16 }}>🚧</div>
        <div style={{ fontSize: 18, fontWeight: 800, color: T.text, marginBottom: 8 }}>개발 중</div>
        <div style={{ fontSize: 14, color: T.muted }}>임금 계산 기능을 준비 중이에요</div>
      </div>
    </div>
  );
}

// ── 팀원 기초 데이터 편집 모달 ────────────────────────────────
function MemberInfoModal({ user, info, onSave, onClose }) {
  const BANKS = ["국민","신한","우리","하나","농협","기업","카카오","토스","새마을","우체국","SC제일","씨티","광주","전북","제주","경남","부산","대구","수협"];
  const [d, setD] = useState({
    empNo: "", ssn: "", joinDate: "", bank: "", account: "",
    hourlyWage: "", employType: "정규직", weeklyHours: 40, insurance: true,
    ...info
  });
  const iStyle = { width: "100%", padding: "11px 14px", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 14, fontWeight: 600, boxSizing: "border-box", fontFamily: "inherit" };

  return (
    <div style={{ position: "fixed", inset: 0, background: "#00000066", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, padding: 16 }}>
      <div style={{ background: T.card, borderRadius: 20, padding: 22, width: "100%", maxWidth: 340, maxHeight: "90vh", overflowY: "auto", boxShadow: "0 20px 60px #00000030" }}>
        <div style={{ fontWeight: 800, fontSize: 17, color: T.text, marginBottom: 4 }}>{user.name} — 기초 데이터</div>
        <div style={{ fontSize: 12, color: T.muted, marginBottom: 18 }}>임금 계산에 사용되는 정보예요</div>

        {[
          ["사번", "empNo", "text", "사번 입력"],
          ["주민등록번호", "ssn", "text", "000000-0000000"],
          ["입사일", "joinDate", "date", ""],
        ].map(([label, key, type, placeholder]) => (
          <div key={key} style={{ marginBottom: 12 }}>
            <div style={{ fontSize: 12, color: T.sub, marginBottom: 5, fontWeight: 600 }}>{label}</div>
            <input type={type} value={d[key]} onChange={e => setD(p => ({ ...p, [key]: e.target.value }))}
              placeholder={placeholder} style={iStyle} />
          </div>
        ))}

        <div style={{ marginBottom: 12 }}>
          <div style={{ fontSize: 12, color: T.sub, marginBottom: 5, fontWeight: 600 }}>은행</div>
          <select value={d.bank} onChange={e => setD(p => ({ ...p, bank: e.target.value }))} style={iStyle}>
            <option value="">선택</option>
            {BANKS.map(b => <option key={b} value={b}>{b}은행</option>)}
          </select>
        </div>

        <div style={{ marginBottom: 12 }}>
          <div style={{ fontSize: 12, color: T.sub, marginBottom: 5, fontWeight: 600 }}>계좌번호</div>
          <input value={d.account} onChange={e => setD(p => ({ ...p, account: e.target.value }))}
            placeholder="계좌번호 입력" style={iStyle} />
        </div>

        <div style={{ marginBottom: 12 }}>
          <div style={{ fontSize: 12, color: T.sub, marginBottom: 5, fontWeight: 600 }}>시급 (원)</div>
          <input type="number" value={d.hourlyWage} onChange={e => setD(p => ({ ...p, hourlyWage: e.target.value }))}
            placeholder="시급 입력" style={iStyle} />
        </div>

        <div style={{ marginBottom: 12 }}>
          <div style={{ fontSize: 12, color: T.sub, marginBottom: 5, fontWeight: 600 }}>고용형태</div>
          <select value={d.employType} onChange={e => setD(p => ({ ...p, employType: e.target.value }))} style={iStyle}>
            {["정규직", "계약직", "파트타임"].map(t => <option key={t} value={t}>{t}</option>)}
          </select>
        </div>

        <div style={{ marginBottom: 12 }}>
          <div style={{ fontSize: 12, color: T.sub, marginBottom: 5, fontWeight: 600 }}>소정근로시간 (주)</div>
          <select value={d.weeklyHours} onChange={e => setD(p => ({ ...p, weeklyHours: Number(e.target.value) }))} style={iStyle}>
            {[40, 35, 30, 25, 20, 15].map(h => <option key={h} value={h}>주 {h}시간</option>)}
          </select>
        </div>

        <div style={{ marginBottom: 20 }}>
          <div style={{ fontSize: 12, color: T.sub, marginBottom: 8, fontWeight: 600 }}>4대보험</div>
          <div style={{ display: "flex", gap: 10 }}>
            {[["적용", true], ["미적용", false]].map(([label, val]) => (
              <button key={label} onClick={() => setD(p => ({ ...p, insurance: val }))}
                style={{ flex: 1, padding: "10px 0", borderRadius: 10, border: "none", fontSize: 13, fontWeight: 700, cursor: "pointer", background: d.insurance === val ? T.adminHeader : T.bg, color: d.insurance === val ? "#fff" : T.muted }}>
                {label}
              </button>
            ))}
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          <Btn variant="ghost" onClick={onClose}>취소</Btn>
          <Btn variant="admin" onClick={() => onSave(d)}>저장</Btn>
        </div>
      </div>
    </div>
  );
}

// ── 관리자 팀원 섹션 ───────────────────────────────────────────
function AdminMembers({ users, annual, leaveRequests, memberInfo = {}, onSaveUsers, onBack }) {
  const [showUserModal, setShowUserModal] = useState(false);
  const [showAccount, setShowAccount] = useState(false);
  const [editInfo, setEditInfo] = useState(null);
  const [editAnnual, setEditAnnual] = useState(null);
  const [delConfirm, setDelConfirm] = useState(null);
  const members = users.filter(u => u.role === "member");

  const saveInfo = async (userId, data) => {
    await setDoc(doc(db, COL_MEMBER_INFO, userId), data);
    setEditInfo(null);
  };

  const saveAnnual = async () => {
    await setDoc(doc(db, COL_ANNUAL, editAnnual.userId), {
      total: Number(editAnnual.total), used: Number(editAnnual.used)
    });
    setEditAnnual(null);
  };

  return (
    <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'Noto Sans KR',sans-serif" }}>
      <div style={{ background: "#7c3aed", padding: "16px 16px 20px" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <button onClick={onBack} style={{ background: "#ffffff18", border: "none", color: "#fff", fontSize: 18, cursor: "pointer", padding: "8px 14px", borderRadius: 12, fontWeight: 700 }}>‹</button>
            <div>
              <div style={{ fontSize: 11, color: "#ffffff40", letterSpacing: 3 }}>ADMIN</div>
              <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>👥 팀원</div>
            </div>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={() => setShowAccount(true)} style={{ background: "#ffffff18", border: "none", color: "#fff", padding: "7px 14px", borderRadius: 18, fontSize: 12, cursor: "pointer", fontWeight: 600 }}>👤 내 계정</button>
            <button onClick={() => setShowUserModal(true)} style={{ background: "#ffffff18", border: "none", color: "#fff", padding: "7px 14px", borderRadius: 18, fontSize: 12, cursor: "pointer", fontWeight: 600 }}>+ 관리</button>
          </div>
        </div>
      </div>

      <div style={{ padding: 16 }}>
        {members.map(u => {
          const a = annual[u.id] || { total: 0, used: 0 };
          const remain = (a.total || 0) - (a.used || 0);
          const pending = leaveRequests.filter(r => r.userId === u.id && r.status === "대기").length;
          const info = memberInfo[u.id] || {};

          return (
            <div key={u.id} style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 14, border: `1px solid ${T.border}`, boxShadow: "0 1px 4px #0000000a" }}>
              {/* 이름 + 버튼 */}
              <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 12 }}>
                <div style={{ width: 44, height: 44, borderRadius: "50%", background: "#7c3aed", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 17, fontWeight: 800, color: "#fff" }}>{u.name[0]}</div>
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 800, fontSize: 16, color: T.text }}>{u.name}</div>
                  <div style={{ fontSize: 12, color: T.muted }}>
                    {info.empNo && `사번 ${info.empNo} · `}{info.employType || ""}
                  </div>
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 6, alignItems: "flex-end" }}>
                  {pending > 0 && <Badge label={`연차신청 ${pending}건`} color="yellow" />}
                  <button onClick={() => setEditInfo({ user: u })}
                    style={{ background: T.bg, border: `1px solid ${T.border}`, color: T.text, borderRadius: 8, padding: "5px 12px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>기초데이터</button>
                </div>
              </div>

              {/* 기초 데이터 요약 */}
              {info.joinDate || info.hourlyWage || info.bank ? (
                <div style={{ background: T.bg, borderRadius: 10, padding: "10px 14px", marginBottom: 10, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
                  {info.joinDate && <div style={{ fontSize: 11 }}><span style={{ color: T.muted }}>입사일 </span><span style={{ fontWeight: 600, color: T.text }}>{info.joinDate}</span></div>}
                  {info.hourlyWage && <div style={{ fontSize: 11 }}><span style={{ color: T.muted }}>시급 </span><span style={{ fontWeight: 600, color: T.text }}>{Number(info.hourlyWage).toLocaleString()}원</span></div>}
                  {info.bank && <div style={{ fontSize: 11 }}><span style={{ color: T.muted }}>은행 </span><span style={{ fontWeight: 600, color: T.text }}>{info.bank}은행</span></div>}
                  {info.account && <div style={{ fontSize: 11 }}><span style={{ color: T.muted }}>계좌 </span><span style={{ fontWeight: 600, color: T.text }}>{info.account}</span></div>}
                  {info.weeklyHours && <div style={{ fontSize: 11 }}><span style={{ color: T.muted }}>소정 </span><span style={{ fontWeight: 600, color: T.text }}>주 {info.weeklyHours}시간</span></div>}
                  {info.insurance !== undefined && <div style={{ fontSize: 11 }}><span style={{ color: T.muted }}>4대보험 </span><span style={{ fontWeight: 600, color: info.insurance ? T.green : T.red }}>{info.insurance ? "적용" : "미적용"}</span></div>}
                </div>
              ) : (
                <div style={{ background: T.bg, borderRadius: 10, padding: "10px 14px", marginBottom: 10, fontSize: 12, color: T.muted, textAlign: "center" }}>
                  기초 데이터 미입력 — 우측 버튼을 눌러 입력해주세요
                </div>
              )}

              {/* 연차 현황 */}
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 6 }}>
                <div style={{ fontSize: 11, color: T.muted, fontWeight: 600 }}>연차 현황</div>
                <button onClick={() => setEditAnnual({ userId: u.id, total: a.total||0, used: a.used||0 })}
                  style={{ background: T.bg, border: `1px solid ${T.border}`, color: T.text, borderRadius: 8, padding: "3px 10px", fontSize: 11, cursor: "pointer", fontWeight: 600 }}>수정</button>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 6 }}>
                {[["총 연차", a.total || 0, T.blue], ["사용", a.used || 0, T.orange], ["잔여", remain, T.green]].map(([l, v, c]) => (
                  <StatBox key={l} label={l} value={v + "일"} color={c} />
                ))}
              </div>
            </div>
          );
        })}

        {/* 연차 신청 목록 */}
        <div style={{ fontSize: 13, color: T.muted, margin: "16px 0 10px", fontWeight: 600 }}>연차 신청 목록</div>
        {leaveRequests.length === 0
          ? <div style={{ textAlign: "center", color: T.muted, padding: 24, background: T.card, borderRadius: 12, border: `1px solid ${T.border}` }}>신청 없음</div>
          : leaveRequests.map(r => {
            const statusColor = { "대기": "yellow", "승인": "green", "반려": "red" };
            return (
              <div key={r.id} style={{ background: T.card, borderRadius: 12, padding: "12px 14px", marginBottom: 8, border: `1px solid ${T.border}` }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                  <div style={{ fontWeight: 700, fontSize: 13, color: T.text, flex: 1 }}>{r.userName} · {r.date} · {r.type}</div>
                  <Badge label={r.status} color={statusColor[r.status] || "gray"} />
                </div>
                {r.note && <div style={{ fontSize: 11, color: T.muted, marginBottom: 8 }}>📝 {r.note}</div>}
                <div style={{ display: "flex", gap: 8 }}>
                  <button onClick={async () => {
                    await setDoc(doc(db, COL_LEAVE_REQ, r.id), { status: "승인" }, { merge: true });
                    // leaves에 완전히 새로 저장 (deleted 필드 없이)
                    const leaveData = { userId: r.userId, date: r.date, type: r.type };
                    if (r.hours) leaveData.hours = r.hours;
                    await setDoc(doc(db, COL_LEAVES, `${r.userId}_${r.date}`), leaveData);
                  }} style={{ flex: 1, padding: "7px 0", borderRadius: 8, border: "none", background: T.greenBg, color: T.green, fontSize: 12, fontWeight: 700, cursor: "pointer" }}>승인</button>
                  <button onClick={async () => {
                    await setDoc(doc(db, COL_LEAVE_REQ, r.id), { status: "반려" }, { merge: true });
                    // 반려 시 leaves에서 완전 삭제
                    try { await deleteDoc(doc(db, COL_LEAVES, `${r.userId}_${r.date}`)); } catch(e) {}
                  }} style={{ flex: 1, padding: "7px 0", borderRadius: 8, border: "none", background: T.redBg, color: T.red, fontSize: 12, fontWeight: 700, cursor: "pointer" }}>반려</button>
                  <button onClick={() => setDelConfirm(r)}
                    style={{ padding: "7px 12px", borderRadius: 8, border: `1px solid ${T.border}`, background: "#fff", color: T.muted, fontSize: 12, cursor: "pointer" }}>삭제</button>
                </div>
              </div>
            );
          })
        }
      </div>

    {/* 삭제 경고 모달 */}
    {delConfirm && (
      <div style={{ position: "fixed", inset: 0, background: "#00000066", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 300, padding: 24 }}>
        <div style={{ background: T.card, borderRadius: 20, padding: 26, width: "100%", maxWidth: 300, boxShadow: "0 20px 60px #00000020" }}>
          <div style={{ textAlign: "center", marginBottom: 16 }}>
            <div style={{ fontSize: 36, marginBottom: 12 }}>⚠️</div>
            <div style={{ fontWeight: 800, fontSize: 16, color: T.text, marginBottom: 8 }}>연차 신청 삭제</div>
            <div style={{ fontSize: 13, color: T.muted, lineHeight: 1.7 }}>
              <strong style={{ color: T.text }}>{delConfirm.userName}</strong>님의<br />
              <strong style={{ color: T.text }}>{delConfirm.date} · {delConfirm.type}</strong><br />
              신청을 삭제할까요?<br />
              <span style={{ color: T.red, fontSize: 12 }}>팀원에게 별도 안내가 필요해요!</span>
            </div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
            <Btn variant="ghost" onClick={() => setDelConfirm(null)}>취소</Btn>
            <Btn variant="red" onClick={async () => { await deleteDoc(doc(db, COL_LEAVE_REQ, delConfirm.id)); setDelConfirm(null); }}>삭제</Btn>
          </div>
        </div>
      </div>
    )}
      {editInfo && (
        <MemberInfoModal user={editInfo.user} info={memberInfo[editInfo.user.id] || {}}
          onSave={data => saveInfo(editInfo.user.id, data)} onClose={() => setEditInfo(null)} />
      )}
      {/* 연차 수정 모달 */}
      {editAnnual && (
        <div style={{ position: "fixed", inset: 0, background: "#00000066", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, padding: 24 }}>
          <div style={{ background: T.card, borderRadius: 20, padding: 26, width: "100%", maxWidth: 300, boxShadow: "0 20px 60px #00000020" }}>
            <div style={{ fontWeight: 800, fontSize: 17, color: T.text, marginBottom: 20 }}>연차 수정</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 24 }}>
              <div>
                <div style={{ fontSize: 12, color: T.muted, marginBottom: 6, fontWeight: 600 }}>총 연차</div>
                <input type="number" value={editAnnual.total} onChange={e => setEditAnnual(p => ({ ...p, total: e.target.value }))}
                  style={{ width: "100%", padding: "12px 0", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 22, fontWeight: 800, boxSizing: "border-box", textAlign: "center" }} />
              </div>
              <div>
                <div style={{ fontSize: 12, color: T.muted, marginBottom: 6, fontWeight: 600 }}>사용</div>
                <input type="number" value={editAnnual.used} onChange={e => setEditAnnual(p => ({ ...p, used: e.target.value }))}
                  style={{ width: "100%", padding: "12px 0", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 22, fontWeight: 800, boxSizing: "border-box", textAlign: "center" }} />
              </div>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              <Btn variant="ghost" onClick={() => setEditAnnual(null)}>취소</Btn>
              <Btn variant="admin" onClick={saveAnnual}>저장</Btn>
            </div>
          </div>
        </div>
      )}
      {showUserModal && <UserManageModal users={users} onSave={async u => { await fbSaveUsers(u, users); setShowUserModal(false); }} onClose={() => setShowUserModal(false)} />}
      {showAccount && <AdminAccountModal users={users} onUpdateUsers={onSaveUsers} onClose={() => setShowAccount(false)} />}
    </div>
  );
}

// ── 관리자 일반 섹션 ───────────────────────────────────────────
function AdminGeneral({ user, users, settings, notices, board, payslips, onSaveSettings, onSaveUsers, onBack }) {
  const [subMenu, setSubMenu] = useState(null);

  const menus = [
    { key: "settings", icon: "⚙", label: "근무 설정", desc: "출퇴근 기준 · GPS · 공휴일" },
    { key: "notice",   icon: "📢", label: "공지사항",  desc: `전체 ${notices.length}건` },
    { key: "board",    icon: "💬", label: "자유게시판", desc: `전체 ${board.length}건` },
    { key: "payslip",  icon: "💰", label: "급여명세서", desc: `전체 ${payslips.length}건` },
  ];

  if (subMenu === "settings") return <SettingsModal settings={settings} onSave={async s => { await onSaveSettings(s); setSubMenu(null); }} onClose={() => setSubMenu(null)} />;

  return (
    <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'Noto Sans KR',sans-serif" }}>
      <div style={{ background: "#ea580c", padding: "16px 16px 20px" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <button onClick={onBack} style={{ background: "#ffffff18", border: "none", color: "#fff", fontSize: 18, cursor: "pointer", padding: "8px 14px", borderRadius: 12, fontWeight: 700 }}>‹</button>
          <div>
            <div style={{ fontSize: 11, color: "#ffffff40", letterSpacing: 3 }}>ADMIN</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>⚙ 일반</div>
          </div>
        </div>
      </div>
      <div style={{ padding: 16 }}>
        {menus.map(m => (
          <button key={m.key} onClick={() => setSubMenu(m.key)}
            style={{ width: "100%", background: T.card, border: `1px solid ${T.border}`, borderRadius: 14, padding: "16px", marginBottom: 10, cursor: "pointer", textAlign: "left", display: "flex", alignItems: "center", gap: 14, boxShadow: "0 1px 4px #0000000a" }}>
            <div style={{ fontSize: 28 }}>{m.icon}</div>
            <div>
              <div style={{ fontWeight: 700, fontSize: 15, color: T.text }}>{m.label}</div>
              <div style={{ fontSize: 12, color: T.muted, marginTop: 2 }}>{m.desc}</div>
            </div>
            <div style={{ marginLeft: "auto", color: T.muted, fontSize: 18 }}>›</div>
          </button>
        ))}
      </div>
      {/* 서브메뉴 */}
      {subMenu === "notice" && (
        <div style={{ position: "fixed", inset: 0, background: T.bg, zIndex: 100, overflowY: "auto" }}>
          <div style={{ background: T.adminHeader, padding: "16px 16px 14px" }}>
            <button onClick={() => setSubMenu(null)} style={{ background: "none", border: "none", color: "#fff", fontSize: 20, cursor: "pointer" }}>‹</button>
          </div>
          <NoticeScreen user={user} users={users} notices={notices} />
        </div>
      )}
      {subMenu === "board" && (
        <div style={{ position: "fixed", inset: 0, background: T.bg, zIndex: 100, overflowY: "auto" }}>
          <div style={{ background: T.adminHeader, padding: "16px 16px 14px" }}>
            <button onClick={() => setSubMenu(null)} style={{ background: "none", border: "none", color: "#fff", fontSize: 20, cursor: "pointer" }}>‹</button>
          </div>
          <BoardScreen user={user} board={board} />
        </div>
      )}
      {subMenu === "payslip" && (
        <div style={{ position: "fixed", inset: 0, background: T.bg, zIndex: 100, overflowY: "auto" }}>
          <div style={{ background: T.adminHeader, padding: "16px 16px 14px" }}>
            <button onClick={() => setSubMenu(null)} style={{ background: "none", border: "none", color: "#fff", fontSize: 20, cursor: "pointer" }}>‹</button>
          </div>
          <PayslipScreen user={user} users={users} payslips={payslips} />
        </div>
      )}
    </div>
  );
}

// ── 관리자 화면 (라우터) ───────────────────────────────────────
function AdminScreen({ user, users, settings, records, leaves, notices, board, payslips, annual, leaveRequests, memberInfo, onSaveRecord, onSaveLeave, onSaveUsers, onSaveSettings, onLogout }) {
  const [section, setSection] = useState(null);

  if (!section) return <AdminHome user={user} onLogout={onLogout} onSection={setSection} />;
  if (section === "attendance") return <AdminAttendance users={users} settings={settings} records={records} leaves={leaves} leaveRequests={leaveRequests} onSaveRecord={onSaveRecord} onSaveLeave={onSaveLeave} onSaveSettings={onSaveSettings} onBack={() => setSection(null)} />;
  if (section === "wage") return <AdminWage users={users} records={records} settings={settings} onBack={() => setSection(null)} />;
  if (section === "members") return <AdminMembers users={users} annual={annual} leaveRequests={leaveRequests} memberInfo={memberInfo} onSaveUsers={onSaveUsers} onBack={() => setSection(null)} />;
  if (section === "general") return <AdminGeneral user={user} users={users} settings={settings} notices={notices} board={board} payslips={payslips} onSaveSettings={onSaveSettings} onSaveUsers={onSaveUsers} onBack={() => setSection(null)} />;
  return null;
}

// ── 공지사항 ────────────────────────────────────────────────────
function NoticeScreen({ user, users, notices }) {
  const isAdmin = user.role === "admin";
  const members = users.filter(u => u.role === "member");
  const [showWrite, setShowWrite] = useState(false);
  const [editTarget, setEditTarget] = useState(null);
  const [title, setTitle] = useState("");
  const [content, setContent] = useState("");
  const [recipient, setRecipient] = useState("all"); // "all" or userId
  const [file, setFile] = useState(null);
  const [uploading, setUploading] = useState(false);
  const [expanded, setExpanded] = useState(null);

  // 내가 볼 수 있는 공지 필터
  const visibleNotices = notices.filter(n =>
    n.recipient === "all" || n.recipient === user.id || user.role === "admin"
  );

  const resetForm = () => { setTitle(""); setContent(""); setRecipient("all"); setFile(null); setShowWrite(false); setEditTarget(null); };

  const submit = async () => {
    if (!title.trim() || !content.trim()) return;
    setUploading(true);
    let fileUrl = null, fileName = null;
    if (file) {
      const path = `notices/${Date.now()}_${file.name}`;
      const storageRef = ref(storage, path);
      await uploadBytes(storageRef, file);
      fileUrl = await getDownloadURL(storageRef);
      fileName = file.name;
    }
    const data = { title: title.trim(), content: content.trim(), recipient, author: user.name, createdAt: new Date().toISOString() };
    if (fileUrl) { data.fileUrl = fileUrl; data.fileName = fileName; }
    await addDoc(collection(db, COL_NOTICES), data);
    resetForm(); setUploading(false);
  };

  const update = async () => {
    if (!title.trim() || !content.trim()) return;
    setUploading(true);
    let fileUrl = editTarget.fileUrl || null, fileName = editTarget.fileName || null;
    if (file) {
      const path = `notices/${Date.now()}_${file.name}`;
      const storageRef = ref(storage, path);
      await uploadBytes(storageRef, file);
      fileUrl = await getDownloadURL(storageRef);
      fileName = file.name;
    }
    const data = { title: title.trim(), content: content.trim(), recipient, fileUrl, fileName };
    await setDoc(doc(db, COL_NOTICES, editTarget.id), data, { merge: true });
    resetForm(); setUploading(false);
  };

  const del = async (n) => {
    if (n.fileUrl) { try { await deleteObject(ref(storage, `notices/${n.fileName}`)); } catch {} }
    await deleteDoc(doc(db, COL_NOTICES, n.id));
  };

  const openEdit = (n) => {
    setEditTarget(n); setTitle(n.title); setContent(n.content);
    setRecipient(n.recipient || "all"); setFile(null); setShowWrite(true);
  };

  const iStyle = { width: "100%", padding: "12px 14px", borderRadius: 12, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 14, boxSizing: "border-box", fontFamily: "inherit", marginBottom: 10 };

  const recipientLabel = (r) => {
    if (r === "all" || !r) return null;
    const m = members.find(u => u.id === r);
    return m ? <Badge label={`${m.name}에게`} color="blue" /> : null;
  };

  return (
    <div style={{ padding: 16 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <div style={{ fontSize: 16, fontWeight: 800, color: T.text }}>📢 공지사항</div>
        {isAdmin && <button onClick={() => { resetForm(); setShowWrite(true); }} style={{ background: T.adminHeader, border: "none", color: "#fff", borderRadius: 10, padding: "8px 14px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>+ 작성</button>}
      </div>

      {showWrite && (
        <div style={{ background: T.card, borderRadius: 14, padding: 16, marginBottom: 16, border: `1px solid ${T.border}` }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: T.text, marginBottom: 10 }}>{editTarget ? "공지 수정" : "새 공지"}</div>
          {/* 수신인 */}
          <div style={{ marginBottom: 10 }}>
            <div style={{ fontSize: 12, color: T.muted, marginBottom: 6, fontWeight: 600 }}>수신인</div>
            <select value={recipient} onChange={e => setRecipient(e.target.value)} style={{ ...iStyle, marginBottom: 0 }}>
              <option value="all">모두</option>
              {members.map(u => <option key={u.id} value={u.id}>{u.name}</option>)}
            </select>
          </div>
          <input value={title} onChange={e => setTitle(e.target.value)} placeholder="제목" style={iStyle} />
          <textarea value={content} onChange={e => setContent(e.target.value)} placeholder="내용" rows={4}
            style={{ ...iStyle, resize: "none", lineHeight: 1.6 }} />
          {/* 첨부파일 */}
          <label style={{ display: "block", padding: "10px 0", borderRadius: 10, border: `2px dashed ${T.border}`, textAlign: "center", fontSize: 12, color: file ? T.green : T.muted, fontWeight: 600, cursor: "pointer", marginBottom: 10 }}>
            {file ? `✓ ${file.name}` : editTarget?.fileName ? `📎 ${editTarget.fileName} (변경하려면 선택)` : "📎 파일 첨부 (선택)"}
            <input type="file" onChange={e => setFile(e.target.files[0])} style={{ display: "none" }} />
          </label>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
            <Btn variant="ghost" onClick={resetForm}>취소</Btn>
            <Btn variant="admin" onClick={editTarget ? update : submit} disabled={uploading}>{uploading ? "처리중..." : editTarget ? "수정" : "등록"}</Btn>
          </div>
        </div>
      )}

      {visibleNotices.length === 0
        ? <div style={{ textAlign: "center", color: T.muted, padding: 40 }}>공지사항이 없어요</div>
        : visibleNotices.map(n => (
          <div key={n.id} style={{ background: T.card, borderRadius: 14, padding: "14px 16px", marginBottom: 10, border: `1px solid ${T.border}` }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }} onClick={() => setExpanded(expanded === n.id ? null : n.id)}>
              <div style={{ flex: 1 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 3 }}>
                  <div style={{ fontWeight: 700, fontSize: 14, color: T.text }}>{n.title}</div>
                  {recipientLabel(n.recipient)}
                  {n.fileName && <Badge label="📎" color="gray" />}
                </div>
                <div style={{ fontSize: 11, color: T.muted }}>{n.author} · {n.createdAt?.slice(0,10)}</div>
              </div>
              <span style={{ color: T.muted, fontSize: 14 }}>{expanded === n.id ? "▲" : "▼"}</span>
            </div>
            {expanded === n.id && (
              <div style={{ marginTop: 12, paddingTop: 12, borderTop: `1px solid ${T.border}` }}>
                <div style={{ fontSize: 14, color: T.text, lineHeight: 1.7, whiteSpace: "pre-wrap", marginBottom: 10 }}>{n.content}</div>
                {n.fileUrl && (
                  <a href={n.fileUrl} target="_blank" rel="noopener noreferrer"
                    style={{ display: "inline-block", background: T.blueBg, color: T.blue, borderRadius: 8, padding: "6px 14px", fontSize: 12, fontWeight: 700, textDecoration: "none", marginBottom: 10 }}>
                    📎 {n.fileName} 다운로드
                  </a>
                )}
                {isAdmin && (
                  <div style={{ display: "flex", gap: 8 }}>
                    <button onClick={() => openEdit(n)} style={{ background: T.bg, border: `1px solid ${T.border}`, color: T.text, borderRadius: 8, padding: "5px 12px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>수정</button>
                    <button onClick={() => del(n)} style={{ background: T.redBg, border: "none", color: T.red, borderRadius: 8, padding: "5px 12px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>삭제</button>
                  </div>
                )}
              </div>
            )}
          </div>
        ))
      }
    </div>
  );
}

// ── 자유게시판 ──────────────────────────────────────────────────
function BoardScreen({ user, board }) {
  const isAdmin = user.role === "admin";
  const [showWrite, setShowWrite] = useState(false);
  const [title, setTitle] = useState(""), [content, setContent] = useState("");
  const [expanded, setExpanded] = useState(null);

  const submit = async () => {
    if (!title.trim() || !content.trim()) return;
    await addDoc(collection(db, COL_BOARD), {
      title: title.trim(), content: content.trim(),
      author: user.name, userId: user.id, createdAt: new Date().toISOString()
    });
    setTitle(""); setContent(""); setShowWrite(false);
  };

  const del = async (id) => { await deleteDoc(doc(db, COL_BOARD, id)); };

  const iStyle = { width: "100%", padding: "12px 14px", borderRadius: 12, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 14, boxSizing: "border-box", fontFamily: "inherit", marginBottom: 10 };

  return (
    <div style={{ padding: 16 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <div style={{ fontSize: 16, fontWeight: 800, color: T.text }}>💬 자유게시판</div>
        <button onClick={() => setShowWrite(!showWrite)} style={{ background: T.adminHeader, border: "none", color: "#fff", borderRadius: 10, padding: "8px 14px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>+ 글쓰기</button>
      </div>

      {showWrite && (
        <div style={{ background: T.card, borderRadius: 14, padding: 16, marginBottom: 16, border: `1px solid ${T.border}` }}>
          <input value={title} onChange={e => setTitle(e.target.value)} placeholder="제목" style={iStyle} />
          <textarea value={content} onChange={e => setContent(e.target.value)} placeholder="내용" rows={4}
            style={{ ...iStyle, resize: "none", lineHeight: 1.6 }} />
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
            <Btn variant="ghost" onClick={() => setShowWrite(false)}>취소</Btn>
            <Btn variant="primary" onClick={submit}>등록</Btn>
          </div>
        </div>
      )}

      {board.length === 0
        ? <div style={{ textAlign: "center", color: T.muted, padding: 40 }}>게시글이 없어요</div>
        : board.map(b => (
          <div key={b.id} style={{ background: T.card, borderRadius: 14, padding: "14px 16px", marginBottom: 10, border: `1px solid ${T.border}` }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }} onClick={() => setExpanded(expanded === b.id ? null : b.id)}>
              <div style={{ flex: 1 }}>
                <div style={{ fontWeight: 700, fontSize: 14, color: T.text }}>{b.title}</div>
                <div style={{ fontSize: 11, color: T.muted, marginTop: 3 }}>{b.author} · {b.createdAt?.slice(0,10)}</div>
              </div>
              <span style={{ color: T.muted, fontSize: 14 }}>{expanded === b.id ? "▲" : "▼"}</span>
            </div>
            {expanded === b.id && (
              <div style={{ marginTop: 12, paddingTop: 12, borderTop: `1px solid ${T.border}` }}>
                <div style={{ fontSize: 14, color: T.text, lineHeight: 1.7, whiteSpace: "pre-wrap" }}>{b.content}</div>
                {(isAdmin || b.userId === user.id) && (
                  <button onClick={() => del(b.id)} style={{ marginTop: 10, background: T.redBg, border: "none", color: T.red, borderRadius: 8, padding: "5px 12px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>삭제</button>
                )}
              </div>
            )}
          </div>
        ))
      }
    </div>
  );
}

// ── 급여명세서 ──────────────────────────────────────────────────
function PayslipScreen({ user, users, payslips }) {
  const isAdmin = user.role === "admin";
  const [uploading, setUploading] = useState(false);
  const [selUser, setSelUser] = useState("");
  const [selMonth, setSelMonth] = useState(new Date(new Date().getTime() + 9*60*60*1000).toISOString().slice(0,7));
  const [file, setFile] = useState(null);
  const [msg, setMsg] = useState("");

  const members = users.filter(u => u.role === "member");
  const myPayslips = isAdmin ? payslips : payslips.filter(p => p.userId === user.id);

  const upload = async () => {
    if (!selUser || !file) { setMsg("팀원과 파일을 선택해주세요"); return; }
    setUploading(true);
    try {
      const ext = file.name.split(".").pop();
      const path = `payslips/${selUser}_${selMonth}.${ext}`;
      const storageRef = ref(storage, path);
      await uploadBytes(storageRef, file);
      const url = await getDownloadURL(storageRef);
      await addDoc(collection(db, COL_PAYSLIPS), {
        userId: selUser, month: selMonth, url, fileName: file.name,
        uploadedBy: user.name, createdAt: new Date().toISOString()
      });
      setMsg("업로드 완료! ✓"); setFile(null);
      setTimeout(() => setMsg(""), 2000);
    } catch (e) { setMsg("업로드 실패: " + e.message); }
    setUploading(false);
  };

  const del = async (p) => {
    try { await deleteObject(ref(storage, `payslips/${p.userId}_${p.month}.${p.fileName.split(".").pop()}`)); } catch {}
    await deleteDoc(doc(db, COL_PAYSLIPS, p.id));
  };

  const kstMonths = Array.from({length:12},(_,i)=>{
    const d=new Date(new Date(new Date().getTime()+9*60*60*1000).getFullYear(),new Date(new Date().getTime()+9*60*60*1000).getMonth()-i,1);
    return new Date(d.getTime()+9*60*60*1000).toISOString().slice(0,7);
  });

  return (
    <div style={{ padding: 16 }}>
      <div style={{ fontSize: 16, fontWeight: 800, color: T.text, marginBottom: 16 }}>💰 급여명세서</div>

      {isAdmin && (
        <div style={{ background: T.card, borderRadius: 14, padding: 16, marginBottom: 16, border: `1px solid ${T.border}` }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: T.text, marginBottom: 12 }}>명세서 업로드</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 10 }}>
            <select value={selUser} onChange={e => setSelUser(e.target.value)}
              style={{ padding: "10px 12px", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: selUser ? T.text : T.muted, fontSize: 13, fontWeight: 600 }}>
              <option value="">팀원 선택</option>
              {members.map(u => <option key={u.id} value={u.id}>{u.name}</option>)}
            </select>
            <select value={selMonth} onChange={e => setSelMonth(e.target.value)}
              style={{ padding: "10px 12px", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 13, fontWeight: 600 }}>
              {kstMonths.map(m => <option key={m} value={m}>{monthLabel(m)}</option>)}
            </select>
          </div>
          <label style={{ display: "block", padding: "12px 0", borderRadius: 10, border: `2px dashed ${T.border}`, textAlign: "center", fontSize: 13, color: file ? T.green : T.muted, fontWeight: 600, cursor: "pointer", marginBottom: 10 }}>
            {file ? `✓ ${file.name}` : "📎 파일 선택 (PNG / PDF)"}
            <input type="file" accept="image/*,.pdf" onChange={e => setFile(e.target.files[0])} style={{ display: "none" }} />
          </label>
          {msg && <div style={{ fontSize: 12, color: msg.includes("✓") ? T.green : T.red, marginBottom: 8, fontWeight: 600 }}>{msg}</div>}
          <Btn variant="admin" onClick={upload} disabled={uploading}>{uploading ? "업로드 중..." : "업로드"}</Btn>
        </div>
      )}

      {myPayslips.length === 0
        ? <div style={{ textAlign: "center", color: T.muted, padding: 40 }}>등록된 명세서가 없어요</div>
        : myPayslips.map(p => {
          const member = users.find(u => u.id === p.userId);
          return (
            <div key={p.id} style={{ background: T.card, borderRadius: 14, padding: "14px 16px", marginBottom: 10, border: `1px solid ${T.border}`, display: "flex", alignItems: "center", gap: 12 }}>
              <div style={{ flex: 1 }}>
                <div style={{ fontWeight: 700, fontSize: 14, color: T.text }}>
                  {isAdmin && `${member?.name} · `}{monthLabel(p.month)}
                </div>
                <div style={{ fontSize: 11, color: T.muted, marginTop: 3 }}>{p.createdAt?.slice(0,10)} 업로드</div>
              </div>
              <a href={p.url} target="_blank" rel="noopener noreferrer"
                style={{ background: T.blueBg, color: T.blue, borderRadius: 8, padding: "7px 14px", fontSize: 12, fontWeight: 700, textDecoration: "none" }}>보기</a>
              {isAdmin && <button onClick={() => del(p)} style={{ background: T.redBg, border: "none", color: T.red, borderRadius: 8, padding: "7px 10px", fontSize: 12, cursor: "pointer", fontWeight: 700 }}>삭제</button>}
            </div>
          );
        })
      }
    </div>
  );
}

// ── 연차 현황 ────────────────────────────────────────────────────
function AnnualScreen({ user, users, annual, leaveRequests }) {
  const isAdmin = user.role === "admin";
  const members = users.filter(u => u.role === "member");
  const [editUser, setEditUser] = useState(null);
  const [total, setTotal] = useState(15);
  const [used, setUsed] = useState(0);
  const [showReqForm, setShowReqForm] = useState(false);
  const [reqDate, setReqDate] = useState("");
  const [reqType, setReqType] = useState("연차");
  const [reqNote, setReqNote] = useState("");
  const [reqMsg, setReqMsg] = useState("");

  const myAnnual = annual[user.id] || { total: 0, used: 0 };
  const myRemain = (myAnnual.total || 0) - (myAnnual.used || 0);
  const myRequests = leaveRequests.filter(r => r.userId === user.id);

  const saveAnnual = async (uid) => {
    await setDoc(doc(db, COL_ANNUAL, uid), { total, used });
    setEditUser(null);
  };

  const submitRequest = async () => {
    if (!reqDate) { setReqMsg("날짜를 선택해주세요"); return; }
    await addDoc(collection(db, COL_LEAVE_REQ), {
      userId: user.id, userName: user.name,
      date: reqDate, type: reqType, note: reqNote,
      status: "대기", createdAt: new Date().toISOString()
    });
    setReqMsg("신청 완료! ✓"); setReqDate(""); setReqNote("");
    setTimeout(() => { setReqMsg(""); setShowReqForm(false); }, 2000);
  };

  const updateReqStatus = async (id, status) => {
    await setDoc(doc(db, COL_LEAVE_REQ, id), { status }, { merge: true });
  };

  const [delConfirm, setDelConfirm] = useState(null);
  const delReq = async (id) => { await deleteDoc(doc(db, COL_LEAVE_REQ, id)); setDelConfirm(null); };

  const statusColor = { "대기": "yellow", "승인": "green", "반려": "red" };

  return (
    <div style={{ padding: 16 }}>
      <div style={{ fontSize: 16, fontWeight: 800, color: T.text, marginBottom: 16 }}>📅 연차</div>

      {!isAdmin && (
        <>
          {/* 내 연차 현황 */}
          <div style={{ background: T.card, borderRadius: 16, padding: "16px 20px", marginBottom: 16, border: `1px solid ${T.border}` }}>
            <div style={{ fontSize: 13, color: T.muted, marginBottom: 12, fontWeight: 600 }}>내 연차 현황</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 8 }}>
              {[["총 연차", myAnnual.total||0+"일", T.blue], ["사용", myAnnual.used||0+"일", T.orange], ["잔여", myRemain+"일", T.green]].map(([l,v,c])=>(
                <StatBox key={l} label={l} value={v} color={c}/>
              ))}
            </div>
          </div>

          {/* 연차 신청 */}
          <div style={{ background: T.card, borderRadius: 14, padding: 16, marginBottom: 16, border: `1px solid ${T.border}` }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: showReqForm ? 14 : 0 }}>
              <div style={{ fontSize: 13, fontWeight: 700, color: T.text }}>연차 신청</div>
              <button onClick={() => setShowReqForm(!showReqForm)} style={{ background: T.adminHeader, border: "none", color: "#fff", borderRadius: 10, padding: "7px 14px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>+ 신청</button>
            </div>
            {showReqForm && (
              <>
                <input type="date" value={reqDate} onChange={e => setReqDate(e.target.value)}
                  style={{ width: "100%", padding: "11px 14px", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 14, fontWeight: 600, boxSizing: "border-box", marginBottom: 10 }} />
                <select value={reqType} onChange={e => setReqType(e.target.value)}
                  style={{ width: "100%", padding: "11px 14px", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 14, fontWeight: 600, boxSizing: "border-box", marginBottom: 10 }}>
                  {LEAVE_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
                </select>
                <input value={reqNote} onChange={e => setReqNote(e.target.value)} placeholder="사유 (선택)"
                  style={{ width: "100%", padding: "11px 14px", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 14, boxSizing: "border-box", marginBottom: 10, fontFamily: "inherit" }} />
                {reqMsg && <div style={{ fontSize: 12, color: reqMsg.includes("✓") ? T.green : T.red, marginBottom: 8, fontWeight: 600 }}>{reqMsg}</div>}
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                  <Btn variant="ghost" onClick={() => setShowReqForm(false)}>취소</Btn>
                  <Btn variant="green" onClick={submitRequest}>신청</Btn>
                </div>
              </>
            )}
          </div>

          {/* 내 신청 내역 */}
          <div style={{ fontSize: 13, color: T.muted, marginBottom: 10, fontWeight: 600 }}>신청 내역</div>
          {myRequests.length === 0
            ? <div style={{ textAlign: "center", color: T.muted, padding: 24, background: T.card, borderRadius: 12, border: `1px solid ${T.border}` }}>신청 내역 없음</div>
            : myRequests.map(r => (
              <div key={r.id} style={{ background: T.card, borderRadius: 12, padding: "12px 14px", marginBottom: 8, border: `1px solid ${T.border}`, display: "flex", alignItems: "center", gap: 10 }}>
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 700, fontSize: 13, color: T.text }}>{r.date} · {r.type}</div>
                  {r.note && <div style={{ fontSize: 11, color: T.muted, marginTop: 2 }}>{r.note}</div>}
                </div>
                <Badge label={r.status} color={statusColor[r.status] || "gray"} />
              </div>
            ))
          }
        </>
      )}

      {isAdmin && (
        <>
          {/* 팀원별 연차 관리 */}
          <div style={{ fontSize: 13, color: T.muted, marginBottom: 10, fontWeight: 600 }}>팀원별 연차 현황</div>
          {members.map(u => {
            const a = annual[u.id] || { total: 0, used: 0 };
            const remain = (a.total||0) - (a.used||0);
            return (
              <div key={u.id} style={{ background: T.card, borderRadius: 14, padding: "14px 16px", marginBottom: 10, border: `1px solid ${T.border}` }}>
                <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 10 }}>
                  <div style={{ width: 36, height: 36, borderRadius: "50%", background: T.headerBg, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14, fontWeight: 800, color: "#fff" }}>{u.name[0]}</div>
                  <span style={{ fontWeight: 700, flex: 1, color: T.text }}>{u.name}</span>
                  <button onClick={() => { setEditUser(u.id); setTotal(a.total||15); setUsed(a.used||0); }}
                    style={{ background: T.bg, border: `1px solid ${T.border}`, color: T.text, borderRadius: 8, padding: "5px 12px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>수정</button>
                </div>
                {editUser === u.id ? (
                  <div>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 10 }}>
                      <div>
                        <div style={{ fontSize: 11, color: T.muted, marginBottom: 4 }}>총 연차</div>
                        <input type="number" value={total} onChange={e => setTotal(Number(e.target.value))}
                          style={{ width: "100%", padding: "10px 12px", borderRadius: 8, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 16, fontWeight: 700, boxSizing: "border-box", textAlign: "center" }} />
                      </div>
                      <div>
                        <div style={{ fontSize: 11, color: T.muted, marginBottom: 4 }}>사용</div>
                        <input type="number" value={used} onChange={e => setUsed(Number(e.target.value))}
                          style={{ width: "100%", padding: "10px 12px", borderRadius: 8, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 16, fontWeight: 700, boxSizing: "border-box", textAlign: "center" }} />
                      </div>
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                      <Btn variant="ghost" onClick={() => setEditUser(null)}>취소</Btn>
                      <Btn variant="admin" onClick={() => saveAnnual(u.id)}>저장</Btn>
                    </div>
                  </div>
                ) : (
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 6 }}>
                    {[["총 연차", a.total||0, T.blue], ["사용", a.used||0, T.orange], ["잔여", remain, T.green]].map(([l,v,c])=>(
                      <StatBox key={l} label={l} value={v+"일"} color={c}/>
                    ))}
                  </div>
                )}
              </div>
            );
          })}

          {/* 연차 신청 목록 */}
          <div style={{ fontSize: 13, color: T.muted, margin: "16px 0 10px", fontWeight: 600 }}>연차 신청 목록</div>
          {leaveRequests.length === 0
            ? <div style={{ textAlign: "center", color: T.muted, padding: 24, background: T.card, borderRadius: 12, border: `1px solid ${T.border}` }}>신청 없음</div>
            : leaveRequests.map(r => (
              <div key={r.id} style={{ background: T.card, borderRadius: 12, padding: "12px 14px", marginBottom: 8, border: `1px solid ${T.border}` }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                  <div style={{ fontWeight: 700, fontSize: 13, color: T.text, flex: 1 }}>{r.userName} · {r.date} · {r.type}</div>
                  <Badge label={r.status} color={statusColor[r.status]||"gray"} />
                </div>
                {r.note && <div style={{ fontSize: 11, color: T.muted, marginBottom: 8 }}>📝 {r.note}</div>}
                <div style={{ display: "flex", gap: 8 }}>
                  <button onClick={() => updateReqStatus(r.id, "승인")} style={{ flex: 1, padding: "7px 0", borderRadius: 8, border: "none", background: T.greenBg, color: T.green, fontSize: 12, fontWeight: 700, cursor: "pointer" }}>승인</button>
                  <button onClick={() => updateReqStatus(r.id, "반려")} style={{ flex: 1, padding: "7px 0", borderRadius: 8, border: "none", background: T.redBg, color: T.red, fontSize: 12, fontWeight: 700, cursor: "pointer" }}>반려</button>
                  <button onClick={() => setDelConfirm(r)} style={{ padding: "7px 12px", borderRadius: 8, border: `1px solid ${T.border}`, background: "#fff", color: T.muted, fontSize: 12, cursor: "pointer" }}>삭제</button>
                </div>
              </div>
            ))
          }
        </>
      )}

      {/* 삭제 경고 모달 */}
      {delConfirm && (
        <div style={{ position: "fixed", inset: 0, background: "#00000066", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 300, padding: 24 }}>
          <div style={{ background: T.card, borderRadius: 20, padding: 26, width: "100%", maxWidth: 300, boxShadow: "0 20px 60px #00000020" }}>
            <div style={{ textAlign: "center", marginBottom: 16 }}>
              <div style={{ fontSize: 36, marginBottom: 12 }}>⚠️</div>
              <div style={{ fontWeight: 800, fontSize: 16, color: T.text, marginBottom: 8 }}>연차 신청 삭제</div>
              <div style={{ fontSize: 13, color: T.muted, lineHeight: 1.7 }}>
                <strong style={{ color: T.text }}>{delConfirm.userName}</strong>님이 신청한<br />
                <strong style={{ color: T.text }}>{delConfirm.date} · {delConfirm.type}</strong><br />
                을 삭제할까요?<br />
                <span style={{ color: T.red, fontSize: 12 }}>팀원에게 별도 안내가 필요해요!</span>
              </div>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              <Btn variant="ghost" onClick={() => setDelConfirm(null)}>취소</Btn>
              <Btn variant="red" onClick={() => delReq(delConfirm.id)}>삭제</Btn>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ── 하단 탭바 ────────────────────────────────────────────────────
function TabBar({ tab, setTab, isAdmin, leaveRequests }) {
  const pendingCount = leaveRequests.filter(r => r.status === "대기").length;
  const tabs = isAdmin
    ? [["att","🏠","출퇴근"],["notice","📢","공지"],["board","💬","게시판"],["payslip","💰","명세서"],["annual","📅","연차"]]
    : [["att","🏠","출퇴근"],["notice","📢","공지"],["board","💬","게시판"],["payslip","💰","명세서"],["annual","📅","연차"]];
  return (
    <div style={{ position: "fixed", bottom: 0, left: 0, right: 0, background: T.card, borderTop: `1px solid ${T.border}`, display: "flex", zIndex: 50, paddingBottom: "env(safe-area-inset-bottom)" }}>
      {tabs.map(([key, icon, label]) => (
        <button key={key} onClick={() => setTab(key)}
          style={{ flex: 1, padding: "10px 0 8px", border: "none", background: "none", cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", gap: 2, position: "relative" }}>
          <span style={{ fontSize: 20 }}>{icon}</span>
          <span style={{ fontSize: 10, fontWeight: tab===key?800:500, color: tab===key?T.adminHeader:T.muted }}>{label}</span>
          {key === "annual" && pendingCount > 0 && isAdmin && (
            <div style={{ position: "absolute", top: 6, right: "calc(50% - 16px)", background: T.red, color: "#fff", borderRadius: "50%", width: 16, height: 16, fontSize: 10, fontWeight: 800, display: "flex", alignItems: "center", justifyContent: "center" }}>{pendingCount}</div>
          )}
          {tab === key && <div style={{ position: "absolute", bottom: 0, left: "20%", right: "20%", height: 2, background: T.adminHeader, borderRadius: 2 }} />}
        </button>
      ))}
    </div>
  );
}

// ── 메인 App ───────────────────────────────────────────────────
function App({ users, settings, records, leaves, notices, board, payslips, annual, leaveRequests, memberInfo, onSaveUsers, onSaveSettings, onSaveRecord, onSaveLeave }) {
  const [user, setUser] = useState(null);
  const [tab, setTab] = useState("att");

  if (!user) return <LoginScreen users={users} onLogin={setUser} onUpdateUsers={onSaveUsers} />;

  const isAdmin = user.role === "admin";

  // 관리자는 대문+섹션 구조 (탭바 없음)
  if (isAdmin) return (
    <AdminScreen user={user} users={users} settings={settings} records={records} leaves={leaves}
      notices={notices} board={board} payslips={payslips} annual={annual} leaveRequests={leaveRequests} memberInfo={memberInfo}
      onSaveRecord={onSaveRecord} onSaveLeave={onSaveLeave}
      onSaveUsers={onSaveUsers} onSaveSettings={onSaveSettings}
      onLogout={() => { setUser(null); setTab("att"); }} />
  );

  // 팀원은 탭바 구조
  return (
    <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'Noto Sans KR',sans-serif", paddingBottom: 70 }}>
      {tab === "att" && (
        <MemberScreen user={user} settings={settings} records={records} leaves={leaves}
          onSaveRecord={onSaveRecord} onLogout={() => { setUser(null); setTab("att"); }} />
      )}
      {tab === "notice" && (
        <>
          <div style={{ background: T.headerBg, padding: "18px 16px 14px" }}>
            <div style={{ fontSize: 11, color: "#ffffff50", letterSpacing: 3 }}>ATTENDANCE</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>공지사항</div>
          </div>
          <NoticeScreen user={user} users={users} notices={notices} />
        </>
      )}
      {tab === "board" && (
        <>
          <div style={{ background: T.headerBg, padding: "18px 16px 14px" }}>
            <div style={{ fontSize: 11, color: "#ffffff50", letterSpacing: 3 }}>ATTENDANCE</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>자유게시판</div>
          </div>
          <BoardScreen user={user} board={board} />
        </>
      )}
      {tab === "payslip" && (
        <>
          <div style={{ background: T.headerBg, padding: "18px 16px 14px" }}>
            <div style={{ fontSize: 11, color: "#ffffff50", letterSpacing: 3 }}>ATTENDANCE</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>급여명세서</div>
          </div>
          <PayslipScreen user={user} users={users} payslips={payslips} />
        </>
      )}
      {tab === "annual" && (
        <>
          <div style={{ background: T.headerBg, padding: "18px 16px 14px" }}>
            <div style={{ fontSize: 11, color: "#ffffff50", letterSpacing: 3 }}>ATTENDANCE</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>연차</div>
          </div>
          <AnnualScreen user={user} users={users} annual={annual} leaveRequests={leaveRequests} />
        </>
      )}
      <TabBar tab={tab} setTab={setTab} isAdmin={isAdmin} leaveRequests={leaveRequests} />
    </div>
  );
}

export default AppLoader;
