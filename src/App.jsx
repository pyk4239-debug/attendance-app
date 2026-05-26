import { useState, useEffect } from "react";
import { db, storage } from "./firebase";
import { jsPDF } from "jspdf";
import html2canvas from "html2canvas";
import {
  doc, onSnapshot, setDoc, getDoc, collection,
  getDocs, writeBatch, addDoc, deleteDoc, query, orderBy
} from "firebase/firestore";
import { ref, uploadBytes, getDownloadURL, deleteObject } from "firebase/storage";

// 기존 서비스워커 완전 제거


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
// v3.1 - 연차수당 일할계산
const COL_USERS    = "users";
const COL_RECORDS  = "records";
const COL_LEAVES   = "leaves";
const COL_NOTICES  = "notices";
const COL_BOARD    = "board";
const COL_PAYSLIPS = "payslips";
const COL_ANNUAL   = "annual";
const COL_LEAVE_REQ = "leave_requests";
const COL_MEMBER_INFO = "member_info";
const COL_READS = "reads"; // 읽음 기록
const COL_REMINDERS = "reminders";
const DOC_SETTINGS = "app/settings";

// ── 초기 데이터 ────────────────────────────────────────────────
const DEFAULT_USERS = [
  { id: "admin", name: "관리자", pin: "000000", role: "admin" },
  { id: "u1", name: "팀원1", pin: "111111", role: "member" },
  { id: "u2", name: "팀원2", pin: "222222", role: "member" },
  { id: "u3", name: "팀원3", pin: "333333", role: "member" },
  { id: "u4", name: "팀원4", pin: "444444", role: "member" },
];

// ── OneSignal 푸시 알림 ──────────────────────────────────────────
async function sendPush({ title, message, targetUserId = null }) {
  try {
    await fetch("/api/notify", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ title, message, targetUserId }),
    });
  } catch(e) { console.error("Push 발송 실패:", e); }
}


// ── 플로팅 뒤로가기 버튼 ─────────────────────────────────────────
function FloatBack({ onClick }) {
  return (
    <button onClick={onClick} style={{
      position: "fixed", bottom: 80, right: 16, zIndex: 200,
      height: 40, padding: "0 16px", borderRadius: 20,
      background: "rgba(255,255,255,0.7)",
      backdropFilter: "blur(10px)",
      WebkitBackdropFilter: "blur(10px)",
      border: "1px solid rgba(255,255,255,0.9)",
      color: "#1a1a2e", fontSize: 13, fontWeight: 700, cursor: "pointer",
      boxShadow: "0 4px 20px rgba(0,0,0,0.15)",
      display: "flex", alignItems: "center", gap: 4
    }}>‹ 이전</button>
  );
}

const DEFAULT_SETTINGS = {
  workStart: "09:00", workEnd: "18:00",
  lunchStart: "12:00", lunchEnd: "13:00",
  officeLat: null, officeLng: null, officeRadius: 200,
  holidays: [],
  monthlyHours: 209,      // 월 소정근로시간 (주40시간 기준)
  // 4대보험 요율 (% 단위)
  ratePension: 4.75,
  rateHealth: 3.595,
  rateEmployment: 0.9,
  rateLongCare: 13.14,
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
  return isWeekend(d) || (holidays || []).some(h => (typeof h === "string" ? h : h.date) === d);
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

// 반차 고려 지각/조퇴 기준 계산
function calcLateMinWithLeave(inI, ws, leave, settings) {
  if (!inI || !ws) return 0;
  // 오전 반차: 출근 기준 = 점심 종료 시간 (오전에 쉬고 점심 후 출근)
  if (leave?.type === "반차(오전)") {
    const lunchEnd = settings?.lunchEnd || "13:00";
    return calcLateMin(inI, lunchEnd);
  }
  return calcLateMin(inI, ws);
}
function calcEarlyOutMinWithLeave(outI, we, inI, ws, leave, settings) {
  if (!outI) return 0;
  // 오후 반차: 퇴근 기준 = 출근시간 + 4시간 (오후에 쉬고 점심 전 퇴근)
  if (leave?.type === "반차(오후)" && inI) {
    const inD = new Date(inI);
    const halfEnd = new Date(inD.getTime() + 4 * 60 * 60 * 1000); // 출근 + 4시간
    const outD = new Date(outI);
    return Math.max(0, (halfEnd.getHours() * 60 + halfEnd.getMinutes()) - (outD.getHours() * 60 + outD.getMinutes()));
  }
  if (!we) return 0;
  return calcEarlyOutMin(outI, we);
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
      const leave = userLeaves?.[date];
      const lm = calcLateMinWithLeave(rec.in, settings.workStart, leave, settings);
      const em = calcEarlyOutMinWithLeave(rec.out, settings.workEnd, rec.in, settings.workStart, leave, settings);
      const om = calcTotalOvertimeMin(rec.in, rec.out, settings.workStart, settings.workEnd);
      const finalLate = rec.lateConfirm !== undefined && rec.lateConfirm !== null ? rec.lateConfirm : lm > 0;
      const finalEarly = rec.earlyConfirm !== undefined && rec.earlyConfirm !== null ? rec.earlyConfirm : em > 0;
      const finalOt = rec.overtimeConfirm !== undefined && rec.overtimeConfirm !== null ? rec.overtimeConfirm : om >= 30;
      if (finalLate) { acc.late++; acc.lateMin += lm; }
      if (finalEarly) { acc.early++; acc.earlyMin += em; }
      if (finalOt) { acc.ot++; acc.otMin += roundTo30(om); }
      if (isHoliday(date, settings.holidays)) acc.holiday++;
      // 외출시간 합산
      (rec.outing || []).forEach(o => {
        if (o.out && o.in) {
          const outMin = Math.round((new Date(o.in) - new Date(o.out)) / 60000);
          if (outMin > 0) acc.outingMin += outMin;
        }
      });
    }
    return acc;
  }, { days: 0, late: 0, lateMin: 0, early: 0, earlyMin: 0, ot: 0, otMin: 0, holiday: 0, annualDays: 0, outingMin: 0 });

  // leaves 기반 연차/반차 집계
  if (userLeaves) {
    Object.entries(userLeaves)
      .filter(([date]) => !month || date.startsWith(month))
      .forEach(([date, l]) => {
        const dayRec = days.find(([d]) => d === date);
        const hasClockIn = !!(dayRec && dayRec[1]?.in);
        if (l.type === "연차") {
          stats.annualDays++;
          if (!hasClockIn) stats.days++;
        } else if (l.type?.includes("반차")) {
          stats.annualDays += 0.5;
          if (!hasClockIn) stats.days++;
        } else if (l.type === "시간연차") {
          stats.annualDays += (l.hours || 1) / 8;
        }
      });
  }

  // 휴무일 + 결근 계산
  if (month) {
    const [y, m] = month.split("-").map(Number);
    const daysInMonth = new Date(y, m, 0).getDate();
    // KST 오늘
    const kstToday = new Date(new Date().getTime() + 9 * 60 * 60 * 1000).toISOString().slice(0, 10);
    let offDays = 0;
    const absentDates = [];
    for (let i = 1; i <= daysInMonth; i++) {
      const dateStr = `${month}-${String(i).padStart(2,"0")}`;
      const dow = new Date(y, m - 1, i).getDay();
      const isOff = dow === 0 || dow === 6 || isHoliday(dateStr, settings.holidays || []);
      if (isOff) { offDays++; continue; } // 공휴일/주말은 미래도 포함
      if (dateStr > kstToday) continue; // 소정근로일 중 미래 날짜는 결근 계산 제외
      // 소정근로일 중 출근기록도 연차도 없으면 결근
      const hasRecord = days.some(([d, rec]) => d === dateStr && rec.in);
      const hasLeave = !!(userLeaves?.[dateStr]);
      if (!hasRecord && !hasLeave) absentDates.push(dateStr);
    }
    // 결근 있는 주 수
    const getWeek = (dateStr) => {
      const d = new Date(dateStr);
      const day = d.getDay() || 7;
      d.setDate(d.getDate() + 4 - day);
      const yearStart = new Date(d.getFullYear(), 0, 1);
      return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    };
    const absentWeeks = new Set(absentDates.map(d => `${d.slice(0,4)}_${getWeek(d)}`));
    stats.offDays = offDays;
    stats.totalDays = daysInMonth;
    stats.absentDays = absentDates.length;
    stats.absentWeeks = absentWeeks.size;
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
  const a1 = Number(lat1), o1 = Number(lng1), a2 = Number(lat2), o2 = Number(lng2);
  if (isNaN(a1) || isNaN(o1) || isNaN(a2) || isNaN(o2)) return null;
  const R = 6371000;
  const dLat = (a2 - a1) * Math.PI / 180;
  const dLng = (o2 - o1) * Math.PI / 180;
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(a1 * Math.PI / 180) * Math.cos(a2 * Math.PI / 180) * Math.sin(dLng / 2) ** 2;
  return Math.round(R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
}
function gpsStatusLabel(gps, settings) {
  if (!gps || gps.lat == null || gps.lng == null) return null;
  const lat = Number(gps.lat);
  const lng = Number(gps.lng);
  if (isNaN(lat) || isNaN(lng)) return null;
  if (Math.abs(lat) < 0.001 && Math.abs(lng) < 0.001) return null;
  if (!settings) return { label: "설정없음", color: "gray" };
  const officeLat = Number(settings.officeLat);
  const officeLng = Number(settings.officeLng);
  if (!settings.officeLat || !settings.officeLng || isNaN(officeLat) || isNaN(officeLng)) return { label: `회사위치미등록`, color: "gray" };
  const dist = calcDistance(lat, lng, officeLat, officeLng);
  if (dist == null || isNaN(dist)) return { label: "계산오류", color: "gray" };
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
  const [reads, setReads] = useState({});
  const [reminders, setReminders] = useState([]);

  useEffect(() => {
    let unsubs = [];

    // 유저
    unsubs.push(onSnapshot(collection(db, COL_USERS), snap => {
      if (snap.empty) return; // 빈 snapshot이면 절대 건드리지 않음
      const all = snap.docs.map(d => d.data());
      const admin = all.filter(u => u.role === "admin");
      const members = all.filter(u => u.role === "member").sort((a,b) => (a.createdAt||"").localeCompare(b.createdAt||""));
      setUsers([...admin, ...members]);
    }));

    // 설정
    unsubs.push(onSnapshot(doc(db, "app", "settings"), snap => {
      if (snap.exists()) {
        const s = snap.data();
        if (s.officeLat != null) s.officeLat = Number(s.officeLat);
        if (s.officeLng != null) s.officeLng = Number(s.officeLng);
        if (s.officeRadius != null) s.officeRadius = Number(s.officeRadius);
        setSettings(s);
      }
      // 문서 없어도 절대 덮어쓰지 않음 (네트워크 일시 오류 대비)
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

    // reads 구독
    unsubs.push(onSnapshot(collection(db, COL_READS), snap => {
      const r = {};
      snap.docs.forEach(d => { r[d.id] = d.data(); });
      setReads(r);
      setReady(true);
    }));

    // 리마인더 구독
    unsubs.push(onSnapshot(query(collection(db, COL_REMINDERS), orderBy("createdAt", "desc")), snap => {
      setReminders(snap.docs.map(d => ({ id: d.id, ...d.data() })));
    }));

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
    notices={notices} board={board} payslips={payslips} annual={annual} leaveRequests={leaveRequests} memberInfo={memberInfo} reads={reads}
    reminders={reminders}
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
    if (u) {
      // OneSignal 태그 설정 + 알림 권한 요청
      try {
        const tagValue = u.role === "admin" ? "admin" : u.id;
        if (window.OneSignalDeferred) {
          window.OneSignalDeferred.push(async (OneSignal) => {
            // 알림 권한 요청 (사용자 인터랙션 후 실행)
            const permission = await OneSignal.Notifications.permission;
            if (!permission) {
              await OneSignal.Notifications.requestPermission();
            }
            await OneSignal.User.addTag("userId", tagValue);
            console.log("OneSignal 태그 설정:", tagValue);
          });
        }
      } catch(e) { console.log("OneSignal 태그 오류:", e); }
      onLogin(u);
    }
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
    if (type === "in") await sendPush({ title: "🏢 출근", message: `${user.name}님이 출근했습니다.`, targetUserId: "admin" });
    if (type === "out") await sendPush({ title: "🏠 퇴근", message: `${user.name}님이 퇴근했습니다.`, targetUserId: "admin" });
    setFlash(msgs[type]); setTimeout(() => setFlash(null), 2500);
  };

  const thisMonth = new Date(now.getTime() + 9 * 60 * 60 * 1000).toISOString().slice(0, 7);
  const [selectedMonth, setSelectedMonth] = useState(thisMonth);
  const isCurrentMonth = selectedMonth === thisMonth;
  const prevMonth = () => {
    const [y, m] = selectedMonth.split("-").map(Number);
    const d = new Date(y, m - 2, 1);
    setSelectedMonth(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`);
  };
  const nextMonth = () => {
    const [y, m] = selectedMonth.split("-").map(Number);
    const d = new Date(y, m, 1);
    const next = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    if (next <= thisMonth) setSelectedMonth(next);
  };
  const monthDays = Object.entries(records[user.id] || {}).filter(([d]) => d.startsWith(selectedMonth)).sort(([a], [b]) => b.localeCompare(a));
  const monthLeaves = Object.entries(leaves[user.id] || {}).filter(([d]) => d.startsWith(selectedMonth));
  const monthLeavesObj = Object.fromEntries(monthLeaves);
  const ms = calcMonthStats(monthDays, settings, monthLeavesObj, null, user.id, selectedMonth);
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

        {/* 버튼 — 당월에만 표시 */}
        {isCurrentMonth && (flash ? (
          <div style={{ textAlign: "center", padding: "16px 0", fontSize: 16, fontWeight: 700, color: T.green }}>{flash}</div>
        ) : (
          <>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 10 }}>
              <Btn variant="green" onClick={() => punch("in")} disabled={hasIn}>{hasIn ? "✓ 출근완료" : "출근"}</Btn>
              <Btn variant="blue" onClick={() => punch("out")} disabled={!hasIn || hasOut}>{hasOut ? "✓ 퇴근완료" : "퇴근"}</Btn>
            </div>
            {hasIn && !hasOut && (
              <div style={{ marginBottom: 14 }}>
                {!isOutside
                  ? <Btn variant="orange" onClick={() => punch("outing_out")}>🚶 외출</Btn>
                  : <Btn variant="primary" onClick={() => punch("outing_in")}>🏃 복귀</Btn>}
              </div>
            )}
          </>
        ))}

        {/* 월 선택 */}
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12, background: T.card, borderRadius: 12, padding: "10px 14px", border: `1px solid ${T.border}` }}>
          <button onClick={prevMonth} style={{ background: "none", border: "none", fontSize: 18, cursor: "pointer", color: T.text, padding: "0 6px" }}>◀</button>
          <div style={{ fontWeight: 800, fontSize: 15, color: T.text }}>{monthLabel(selectedMonth)}</div>
          <button onClick={nextMonth} style={{ background: "none", border: "none", fontSize: 18, cursor: "pointer", color: isCurrentMonth ? T.muted : T.text, padding: "0 6px" }} disabled={isCurrentMonth}>▶</button>
        </div>

        {/* 이번달 현황 */}
        <div style={{ background: T.card, borderRadius: 16, padding: "14px 16px", marginBottom: 16, border: `1px solid ${T.border}`, boxShadow: "0 1px 4px #0000000a" }}>
          <div style={{ fontSize: 12, color: T.muted, marginBottom: 10, fontWeight: 600 }}>{monthLabel(selectedMonth)} 현황</div>
          {/* 근무 현황 */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 6, marginBottom: 6 }}>
            {[["출근", ms.days + "일", T.green], ["결근", (ms.absentDays||0) + "일", T.red], ["연차", ms.annualDays > 0 ? ms.annualDays + "일" : "0일", "#7c3aed"]].map(([l,v,c]) => <StatBox key={l} label={l} value={v} color={c} />)}
          </div>
          {/* 추가근무 */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(2,1fr)", gap: 6, marginBottom: 6 }}>
            {[["잔업", fmtMinutes(ms.otMin), T.purple], ["휴일근무", ms.holiday + "일", T.red]].map(([l,v,c]) => <StatBox key={l} label={l} value={v} color={c} />)}
          </div>
          {/* 차감시간 통합 */}
          {(() => {
            const totalDeductMin = ms.lateMin + ms.earlyMin + (ms.outingMin||0);
            const parts = [];
            if (ms.late > 0) parts.push(`지각 ${ms.late}회 ${fmtMinutes(ms.lateMin)}`);
            if (ms.early > 0) parts.push(`조퇴 ${ms.early}회 ${fmtMinutes(ms.earlyMin)}`);
            if ((ms.outingMin||0) > 0) parts.push(`외출 ${fmtMinutes(ms.outingMin)}`);
            return (
              <div style={{ background: totalDeductMin > 0 ? T.orangeBg : T.bg, borderRadius: 10, padding: "10px 14px", marginBottom: 6, border: `1px solid ${totalDeductMin > 0 ? T.orange : T.border}` }}>
                <div style={{ fontSize: 10, color: T.muted, marginBottom: 3, fontWeight: 500 }}>차감시간 (지각·조퇴·외출)</div>
                <div style={{ fontSize: 15, fontWeight: 800, color: totalDeductMin > 0 ? T.orange : T.muted }}>
                  {totalDeductMin > 0 ? fmtMinutes(totalDeductMin) : "-"}
                </div>
                {parts.length > 0 && <div style={{ fontSize: 11, color: T.orange, marginTop: 3 }}>{parts.join(" · ")}</div>}
              </div>
            );
          })()}
          {/* 전체/휴무일 한 줄 */}
          <div style={{ fontSize: 11, color: T.muted, textAlign: "right", paddingRight: 2 }}>
            {monthLabel(selectedMonth)} · 전체 {ms.totalDays||0}일 · 휴무 {ms.offDays||0}일 · 근무 {(ms.totalDays||0) - (ms.offDays||0)}일
          </div>
        </div>

        {/* 이번달 기록 */}
        <div style={{ fontSize: 13, color: T.muted, marginBottom: 10, fontWeight: 600 }}>{monthLabel(selectedMonth)} 기록</div>
        {monthDays.length === 0
          ? <div style={{ textAlign: "center", color: T.muted, padding: 30, fontSize: 14, background: T.card, borderRadius: 12, border: `1px solid ${T.border}` }}>기록 없음</div>
          : monthDays.map(([date, rec]) => {
            const leave = leaves[user.id]?.[date];
            const lm = calcLateMinWithLeave(rec.in, settings.workStart, leave, settings);
            const em = calcEarlyOutMinWithLeave(rec.out, settings.workEnd, rec.in, settings.workStart, leave, settings);
            const om = calcTotalOvertimeMin(rec.in, rec.out, settings.workStart, settings.workEnd);
            const late = lm > 0, early = em > 0, ot = om >= 30;
            const weekend = isHoliday(date, settings.holidays);
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
    if (inTime) nr.in = setTimeOnDate(date, inTime);
    else { delete nr.in; delete nr.inGps; } // 출근 시간 지우면 GPS도 삭제
    if (outTime) nr.out = setTimeOnDate(date, outTime);
    else { delete nr.out; delete nr.outGps; } // 퇴근 시간 지우면 GPS도 삭제
    await onSave(date, nr, leaveType ? { type: leaveType, hours: leaveHours } : null);
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "#00000066", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, padding: 16 }}>
      <div style={{ background: T.card, borderRadius: 20, padding: 22, width: "100%", maxWidth: 340, maxHeight: "90vh", overflowY: "scroll", WebkitOverflowScrolling: "touch", boxShadow: "0 20px 60px #00000030" }}>
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
        <FloatBack onClick={() => setDrillUser(null)} />
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 16 }}>
          <div style={{ width: 42, height: 42, borderRadius: "50%", background: T.headerBg, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 17, fontWeight: 800, color: "#fff" }}>{drillUser.name[0]}</div>
          <div style={{ flex: 1 }}>
            <div style={{ fontWeight: 800, fontSize: 17, color: T.text }}>{drillUser.name}</div>
            <div style={{ fontSize: 12, color: T.muted }}>{monthLabel(selectedMonth)}</div>
          </div>
        </div>
        {/* 근무 현황 */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 6, marginBottom: 6 }}>
          {[["출근", ms.days + "일", T.green], ["결근", (ms.absentDays||0) + "일", T.red], ["연차", ms.annualDays > 0 ? ms.annualDays + "일" : "0일", "#7c3aed"]].map(([l,v,c]) => <StatBox key={l} label={l} value={v} color={c} />)}
        </div>
        {/* 추가근무 */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(2,1fr)", gap: 6, marginBottom: 6 }}>
          {[["잔업", fmtMinutes(ms.otMin), T.purple], ["휴일근무", ms.holiday + "일", T.red]].map(([l,v,c]) => <StatBox key={l} label={l} value={v} color={c} />)}
        </div>
        {/* 차감시간 통합 */}
        {(() => {
          const totalDeductMin = ms.lateMin + ms.earlyMin + (ms.outingMin||0);
          const parts = [];
          if (ms.late > 0) parts.push(`지각 ${ms.late}회 ${fmtMinutes(ms.lateMin)}`);
          if (ms.early > 0) parts.push(`조퇴 ${ms.early}회 ${fmtMinutes(ms.earlyMin)}`);
          if ((ms.outingMin||0) > 0) parts.push(`외출 ${fmtMinutes(ms.outingMin)}`);
          return (
            <div style={{ background: totalDeductMin > 0 ? T.orangeBg : T.bg, borderRadius: 10, padding: "10px 14px", marginBottom: 6, border: `1px solid ${totalDeductMin > 0 ? T.orange : T.border}` }}>
              <div style={{ fontSize: 10, color: T.muted, marginBottom: 3, fontWeight: 500 }}>차감시간 (지각·조퇴·외출)</div>
              <div style={{ fontSize: 15, fontWeight: 800, color: totalDeductMin > 0 ? T.orange : T.muted }}>
                {totalDeductMin > 0 ? fmtMinutes(totalDeductMin) : "-"}
              </div>
              {parts.length > 0 && <div style={{ fontSize: 11, color: T.orange, marginTop: 3 }}>{parts.join(" · ")}</div>}
            </div>
          );
        })()}
        {/* 전체/휴무일 한 줄 */}
        <div style={{ fontSize: 11, color: T.muted, textAlign: "right", paddingRight: 2, marginBottom: 12 }}>
          {monthLabel(selectedMonth)} · 전체 {ms.totalDays||0}일 · 휴무 {ms.offDays||0}일 · 근무 {(ms.totalDays||0) - (ms.offDays||0)}일
        </div>
        <div style={{ fontSize: 12, color: T.muted, marginBottom: 10, fontWeight: 600 }}>날짜별 상세</div>
        {(() => {
          // 출근 기록 + 연차만 있는 날짜 합치기
          const leaveDates = Object.entries(userLeaves)
            .filter(([date]) => date.startsWith(selectedMonth) && !days.find(([d]) => d === date))
            .map(([date, l]) => [date, { in: null, out: null, leaveOnly: true }]);
          const allDays = [...days, ...leaveDates].sort(([a], [b]) => b.localeCompare(a));
          if (allDays.length === 0) return <div style={{ textAlign: "center", color: T.muted, padding: 24, fontSize: 14, background: T.card, borderRadius: 12, border: `1px solid ${T.border}` }}>기록 없음</div>;
          return allDays.map(([date, rec]) => {
            const leave = userLeaves[date];
            const lm = calcLateMinWithLeave(rec.in, settings.workStart, leave, settings);
            const em = calcEarlyOutMinWithLeave(rec.out, settings.workEnd, rec.in, settings.workStart, leave, settings);
            const om = calcTotalOvertimeMin(rec.in, rec.out, settings.workStart, settings.workEnd);
            const late = lm > 0, early = em > 0, ot = om >= 30, weekend = isHoliday(date, settings.holidays);
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

  // holidays: [{date, memo}] 형태로 저장 — 하위호환: string도 처리
  const normalizeHolidays = (arr) => (arr || []).map(h =>
    typeof h === "string" ? { date: h, memo: "" } : h
  );
  const holidayList = normalizeHolidays(s.holidays);

  const addHoliday = () => {
    if (!newHoliday) return;
    if (holidayList.some(h => h.date === newHoliday)) { setNewHoliday(""); setNewMemo(""); return; }
    const updated = [...holidayList, { date: newHoliday, memo: newMemo.trim() }]
      .sort((a, b) => a.date.localeCompare(b.date));
    setS(p => ({ ...p, holidays: updated }));
    setNewHoliday(""); setNewMemo("");
  };

  const removeHoliday = (date) => {
    setS(p => ({ ...p, holidays: holidayList.filter(h => h.date !== date) }));
  };

  const [newMemo, setNewMemo] = useState("");
  const [viewYear, setViewYear] = useState(String(new Date().getFullYear()));
  const years = [...new Set(holidayList.map(h => h.date.slice(0, 4)))].sort();
  if (!years.includes(viewYear) && years.length > 0) { /* viewYear는 존재하는 연도 중 최신으로 */ }
  const filteredHolidays = holidayList.filter(h => h.date.startsWith(viewYear));

  return (
    <div style={{ position: "fixed", inset: 0, background: "#00000066", display: "flex", alignItems: "flex-start", justifyContent: "center", zIndex: 200, padding: "20px 16px", overflowY: "auto" }}>
      <div style={{ background: T.card, borderRadius: 20, padding: 22, width: "100%", maxWidth: 340, margin: "auto", boxShadow: "0 20px 60px #00000020" }}>
        <div style={{ fontWeight: 800, fontSize: 17, color: T.text, marginBottom: 20 }}>근무 설정</div>

        <div style={{ marginBottom: 16 }}>
          <div style={{ fontSize: 13, color: T.green, marginBottom: 6, fontWeight: 700 }}>출근 기준</div>
          <input type="time" value={s.workStart} onChange={e => setS(p => ({ ...p, workStart: e.target.value }))} style={{ ...iStyle, borderColor: T.green + "44" }} />
        </div>
        <div style={{ marginBottom: 16 }}>
          <div style={{ fontSize: 13, color: T.blue, marginBottom: 6, fontWeight: 700 }}>퇴근 기준</div>
          <input type="time" value={s.workEnd} onChange={e => setS(p => ({ ...p, workEnd: e.target.value }))} style={{ ...iStyle, borderColor: T.blue + "44" }} />
        </div>
        <div style={{ marginBottom: 16 }}>
          <div style={{ fontSize: 13, color: T.text, marginBottom: 3, fontWeight: 700 }}>⏱ 월 소정근로시간</div>
          <div style={{ fontSize: 11, color: T.muted, marginBottom: 6 }}>주40시간=209h · 주36시간=188h · 법 개정 시 변경</div>
          <input type="number" value={s.monthlyHours ?? 209}
            onChange={e => setS(p => ({ ...p, monthlyHours: Number(e.target.value) }))}
            style={{ ...iStyle }} placeholder="209" />
        </div>
        {/* 점심시간 */}
        <div style={{ background: T.bg, border: `1px solid ${T.border}`, borderRadius: 14, padding: "14px 16px", marginBottom: 16 }}>
          <div style={{ fontSize: 13, color: T.text, fontWeight: 700, marginBottom: 4 }}>🍱 점심시간</div>
          <div style={{ fontSize: 11, color: T.muted, marginBottom: 12 }}>반차 기준 계산에 사용돼요</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
            <div>
              <div style={{ fontSize: 12, color: T.muted, marginBottom: 5, fontWeight: 600 }}>시작</div>
              <input type="time" value={s.lunchStart || "12:00"} onChange={e => setS(p => ({ ...p, lunchStart: e.target.value }))}
                style={{ ...iStyle, fontSize: 16 }} />
            </div>
            <div>
              <div style={{ fontSize: 12, color: T.muted, marginBottom: 5, fontWeight: 600 }}>종료</div>
              <input type="time" value={s.lunchEnd || "13:00"} onChange={e => setS(p => ({ ...p, lunchEnd: e.target.value }))}
                style={{ ...iStyle, fontSize: 16 }} />
            </div>
          </div>
        </div>

        {/* 공휴일 관리 */}
        <div style={{ background: T.bg, border: `1px solid ${T.border}`, borderRadius: 14, padding: "14px 16px", marginBottom: 20 }}>
          <div style={{ fontSize: 13, color: T.text, fontWeight: 700, marginBottom: 4 }}>🗓 공휴일 지정</div>
          <div style={{ fontSize: 11, color: T.muted, marginBottom: 12, lineHeight: 1.5 }}>
            토/일 외 공휴일을 등록하세요. 등록된 날은 휴일근무로 자동 처리돼요.
          </div>
          {/* 입력 */}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 6 }}>
            <input type="date" value={newHoliday} onChange={e => setNewHoliday(e.target.value)}
              style={{ padding: "10px 10px", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 13, fontWeight: 600, boxSizing: "border-box" }} />
            <input value={newMemo} onChange={e => setNewMemo(e.target.value)} placeholder="공휴일 이름"
              style={{ padding: "10px 10px", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 13, fontWeight: 600, boxSizing: "border-box" }} />
          </div>
          <button onClick={addHoliday}
            style={{ width: "100%", background: T.adminHeader, border: "none", color: "#fff", borderRadius: 10, padding: "10px 0", fontSize: 13, fontWeight: 700, cursor: "pointer", marginBottom: 14 }}>+ 추가</button>
          {/* 연도 탭 */}
          {holidayList.length === 0 ? (
            <div style={{ fontSize: 12, color: T.muted, textAlign: "center", padding: "8px 0" }}>등록된 공휴일 없음</div>
          ) : <>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
              <button onClick={() => { const idx = years.indexOf(viewYear); if (idx > 0) setViewYear(years[idx-1]); }}
                style={{ background: "none", border: `1px solid ${T.border}`, borderRadius: 8, padding: "4px 10px", fontSize: 14, cursor: "pointer", color: T.text, fontWeight: 700 }}>‹</button>
              <span style={{ fontSize: 14, fontWeight: 800, color: T.text }}>{viewYear}년 ({filteredHolidays.length}일)</span>
              <button onClick={() => { const idx = years.indexOf(viewYear); if (idx < years.length - 1) setViewYear(years[idx+1]); }}
                style={{ background: "none", border: `1px solid ${T.border}`, borderRadius: 8, padding: "4px 10px", fontSize: 14, cursor: "pointer", color: T.text, fontWeight: 700 }}>›</button>
            </div>
            {filteredHolidays.length === 0 ? (
              <div style={{ fontSize: 12, color: T.muted, textAlign: "center", padding: "8px 0" }}>{viewYear}년 등록 없음</div>
            ) : filteredHolidays.map(h => (
              <div key={h.date} style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 0", borderTop: `1px solid ${T.border}` }}>
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 13, fontWeight: 700, color: T.text }}>{formatDate(h.date)}</div>
                  {h.memo && <div style={{ fontSize: 11, color: T.muted, marginTop: 2 }}>{h.memo}</div>}
                </div>
                <button onClick={() => removeHoliday(h.date)}
                  style={{ background: T.redBg, border: "none", color: T.red, borderRadius: 8, padding: "4px 10px", fontSize: 12, cursor: "pointer", fontWeight: 700, flexShrink: 0 }}>삭제</button>
              </div>
            ))}
          </>}
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

        {/* 4대보험 요율 */}
        <div style={{ background: T.bg, border: `1px solid ${T.border}`, borderRadius: 14, padding: "14px 16px", marginBottom: 20 }}>
          <div style={{ fontSize: 13, color: T.text, fontWeight: 700, marginBottom: 4 }}>🏥 4대보험 요율 (%)</div>
          <div style={{ fontSize: 11, color: T.muted, marginBottom: 12 }}>매년 변경 시 업데이트해주세요</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
            {[
              ["국민연금", "ratePension", 4.75],
              ["건강보험", "rateHealth", 3.595],
              ["고용보험", "rateEmployment", 0.9],
              ["장기요양*", "rateLongCare", 13.14],
            ].map(([label, key, defaultVal]) => (
              <div key={key}>
                <div style={{ fontSize: 11, color: T.muted, marginBottom: 4, fontWeight: 600 }}>{label}</div>
                <input type="number" step="0.001"
                  value={s[key] !== undefined && s[key] !== null ? s[key] : defaultVal}
                  onChange={e => setS(p => ({ ...p, [key]: Number(e.target.value) }))}
                  style={{ width: "100%", padding: "9px 12px", borderRadius: 8, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 14, fontWeight: 600, boxSizing: "border-box" }} />
              </div>
            ))}
          </div>
          <div style={{ fontSize: 10, color: T.muted, marginTop: 8 }}>* 장기요양은 건강보험료 대비 %</div>
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
      <div style={{ background: T.card, borderRadius: 20, padding: 22, width: "100%", maxWidth: 320, maxHeight: "88vh", overflowY: "scroll", WebkitOverflowScrolling: "touch", boxShadow: "0 20px 60px #00000020" }}>
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
// ── 퇴직소득세 상수 (법 개정 시 여기만 수정) ───────────────────
const TENURE_DEDUCTION_BRACKETS = [
  { upTo: 5,  perYear: 1000000 },   // 5년 이하: 100만 × 연수
  { upTo: 10, base: 5000000,  perYear: 2000000 }, // 5~10년
  { upTo: 20, base: 15000000, perYear: 2500000 }, // 10~20년
  { upTo: Infinity, base: 40000000, perYear: 3000000 }, // 20년 초과
];
const CONVERTED_WAGE_DEDUCTION_BRACKETS = [
  { upTo: 8000000,   rate: 1.0,  base: 0 },
  { upTo: 70000000,  rate: 0.6,  base: 8000000 },
  { upTo: 100000000, rate: 0.55, base: 45200000 },
  { upTo: 300000000, rate: 0.45, base: 62450000 },
  { upTo: Infinity,  rate: 0.35, base: 152450000 },
];
const INCOME_TAX_BRACKETS = [
  { upTo: 14000000,   rate: 0.06, base: 0 },
  { upTo: 50000000,   rate: 0.15, base: 840000 },
  { upTo: 88000000,   rate: 0.24, base: 6240000 },
  { upTo: 150000000,  rate: 0.35, base: 15360000 },
  { upTo: 300000000,  rate: 0.38, base: 19400000 },
  { upTo: 500000000,  rate: 0.40, base: 25400000 },
  { upTo: Infinity,   rate: 0.42, base: 35400000 },
];

function calcTenureDeduction(years) {
  const y = Math.ceil(years);
  if (y <= 5) return y * 1000000;
  if (y <= 10) return 5000000 + (y - 5) * 2000000;
  if (y <= 20) return 15000000 + (y - 10) * 2500000;
  return 40000000 + (y - 20) * 3000000;
}
function calcConvertedWageDeduction(amount) {
  if (amount <= 8000000) return amount;
  if (amount <= 70000000) return 8000000 + (amount - 8000000) * 0.6;
  if (amount <= 100000000) return 45200000 + (amount - 70000000) * 0.55;
  if (amount <= 300000000) return 62450000 + (amount - 100000000) * 0.45;
  return 152450000 + (amount - 300000000) * 0.35;
}
function calcIncomeTax(taxBase) {
  if (taxBase <= 0) return 0;
  for (const b of INCOME_TAX_BRACKETS) {
    if (taxBase <= b.upTo) return Math.floor(taxBase * b.rate - b.base);
  }
  return 0;
}

// ── 퇴직금 계산 ────────────────────────────────────────────────
function AdminSeverance({ users, memberInfo, annual, onBack }) {
  const members = users.filter(u => u.role === "member");

;
  const [selUser, setSelUser] = useState(members[0]?.id || "");
  const [retireDate, setRetireDate] = useState("");
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);

  const user = members.find(u => u.id === selUser);
  const info = memberInfo[selUser] || {};
  const joinDate = info.joinDate || "";
  const hourlyWage = Number(info.hourlyWage || 0);
  const annualData = annual[selUser] || { total: 0, used: 0 };

  const calc = async () => {
    if (!selUser || !retireDate || !joinDate) return;
    setLoading(true);

    const retire = new Date(retireDate);
    const join = new Date(joinDate);
    const workDays = Math.floor((retire - join) / (1000 * 60 * 60 * 24));
    const workYears = workDays / 365;
    const tenureYears = Math.ceil(workYears); // 근속연수 (올림)

    // ① 지급 요건 체크
    const eligible1year = workDays >= 365;
    const weeklyHours = Number(info.weeklyHours || 40);
    const eligible15h = weeklyHours >= 15;
    const eligible = eligible1year && eligible15h;

    // wages 로드
    const snap = await getDocs(collection(db, "wages"));
    const all = snap.docs.map(d => d.data()).filter(w => w.userId === selUser);
    all.sort((a, b) => b.yearMonth.localeCompare(a.yearMonth));

    // 퇴직일 기준 직전 3개월
    const d3ago = new Date(retire); d3ago.setMonth(d3ago.getMonth() - 3);
    const recentWages = all.filter(w => {
      const wDate = new Date(w.yearMonth + "-01");
      return wDate >= d3ago && wDate < retire;
    }).slice(0, 3);

    const calDays3 = Math.floor((retire - d3ago) / (1000 * 60 * 60 * 24));

    // ② 3개월 임금총액 (상여 제외)
    const pay3 = recentWages.reduce((s, w) => {
      const income = Number(w.totalIncome || 0);
      const bonus = Number(w.bonus || 0);
      return s + income - bonus;
    }, 0);

    // ③ 연간 상여금 → 3개월 환산
    const annualBonus = all
      .filter(w => {
        const wDate = new Date(w.yearMonth + "-01");
        const oneYearAgo = new Date(retire); oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);
        return wDate >= oneYearAgo && wDate < retire;
      })
      .reduce((s, w) => s + Number(w.bonus || 0), 0);
    const bonusFor3 = Math.round(annualBonus / 12 * 3);

    // ④ 3개월 임금총액 = 기본임금 + 상여환산
    const totalPay3 = pay3 + bonusFor3;
    const avgDailyWage = calDays3 > 0 ? Math.round(totalPay3 / calDays3) : 0;

    // ⑤ 퇴직금 = 평균임금 × 30 × (근속일수/365)
    const severancePay = Math.round(avgDailyWage * 30 * workYears);

    // ⑥ 잔여연차수당 - 당해연도 일할계산 (최신 annual 데이터 직접 참조)
    const latestAnnual = annual[selUser] || { total: 0, used: 0 };
    const totalAnnual = latestAnnual.total || 0;
    const usedAnnual = latestAnnual.used || 0;
    const yearStart = new Date(retire.getFullYear(), 0, 1);
    const workedDaysThisYear = Math.floor((retire - yearStart) / (1000 * 60 * 60 * 24));
    // 당해연도 발생 연차 = 총연차 × (당해연도 근무일수/365), 소수점 1자리
    const earnedThisYear = Math.floor(totalAnnual * workedDaysThisYear / 365 * 10) / 10;
    const annualRemain = Math.max(0, Math.round((earnedThisYear - usedAnnual) * 10) / 10);
    const annualAllowance = Math.round(hourlyWage * 8 * annualRemain);

    // ⑦ 퇴직소득세 계산 (8단계)
    const step1 = severancePay; // 퇴직급여
    const step2 = calcTenureDeduction(tenureYears); // 근속연수공제
    const step3 = Math.max(0, step1 - step2); // 과세표준
    const step4 = Math.round(step3 * 12 / tenureYears); // 환산급여
    const step5 = Math.round(calcConvertedWageDeduction(step4)); // 환산급여공제
    const step6 = Math.max(0, step4 - step5); // 환산과세표준
    const step7 = calcIncomeTax(step6); // 환산산출세액
    const retirementTax = Math.round(step7 * tenureYears / 12); // 퇴직소득세
    const localTax = Math.floor(retirementTax * 0.1 / 10) * 10; // 지방소득세(주민세)

    const totalDeduct = retirementTax + localTax;
    const netPay = severancePay + annualAllowance - totalDeduct;

    setResult({
      // 요건
      eligible, eligible1year, eligible15h, workDays, workYears, tenureYears, weeklyHours,
      // 평균임금
      recentWages, pay3, annualBonus, bonusFor3, totalPay3, calDays3, avgDailyWage,
      // 퇴직금
      severancePay,
      // 연차
      annualRemain, annualAllowance, earnedThisYear, usedAnnual, totalAnnual, workedDaysThisYear,
      // 세금 단계
      step1, step2, step3, step4, step5, step6, step7,
      retirementTax, localTax, totalDeduct, netPay,
    });
    setLoading(false);
  };

  const iStyle = { width: "100%", padding: "12px 14px", borderRadius: 12, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 15, fontWeight: 600, boxSizing: "border-box", fontFamily: "inherit" };
  const Row = ({ label, value, calc, bold, color }) => (
    <div style={{ padding: "5px 0", borderBottom: `1px solid ${T.border}` }}>
      <div style={{ display: "flex", justifyContent: "space-between" }}>
        <span style={{ fontSize: 12, color: bold ? T.text : T.muted, fontWeight: bold ? 700 : 400 }}>{label}</span>
        <span style={{ fontSize: 13, fontWeight: bold ? 800 : 600, color: color || T.text }}>{Number(value||0).toLocaleString()}원</span>
      </div>
      {calc && <div style={{ fontSize: 10, color: T.muted, marginTop: 2 }}>{calc}</div>}
    </div>
  );

  return (
    <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'Noto Sans KR',sans-serif" }}>
      <div style={{ background: "#0891b2", padding: "16px 16px 20px" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <button onClick={onBack} style={{ background: "#ffffff18", border: "none", color: "#fff", fontSize: 18, cursor: "pointer", padding: "8px 14px", borderRadius: 12, fontWeight: 700 }}>‹</button>
          <div>
            <div style={{ fontSize: 11, color: "#ffffff40", letterSpacing: 3 }}>ADMIN</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>💼 퇴직금 계산</div>
          </div>
        </div>
      </div>

      <div style={{ padding: 16 }}>
        {/* 팀원 선택 */}
        <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 12, border: `1px solid ${T.border}` }}>
          <div style={{ fontSize: 13, color: T.muted, fontWeight: 700, marginBottom: 8 }}>팀원 선택</div>
          <select value={selUser} onChange={e => { setSelUser(e.target.value); setResult(null); }} style={iStyle}>
            {members.map(u => <option key={u.id} value={u.id}>{u.name}</option>)}
          </select>
          {user && (
            <div style={{ marginTop: 10, display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 6 }}>
              {[["입사일", joinDate||"미입력"], ["시급", Number(info.hourlyWage||0).toLocaleString()+"원"], ["주소정", (info.weeklyHours||40)+"시간"]].map(([l,v]) => (
                <div key={l} style={{ background: T.bg, borderRadius: 8, padding: "8px 10px", textAlign: "center" }}>
                  <div style={{ fontSize: 10, color: T.muted }}>{l}</div>
                  <div style={{ fontSize: 12, fontWeight: 700, color: T.text }}>{v}</div>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* 퇴직일 */}
        <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 12, border: `1px solid ${T.border}` }}>
          <div style={{ fontSize: 13, color: T.muted, fontWeight: 700, marginBottom: 8 }}>퇴직일</div>
          <input type="date" value={retireDate} onChange={e => { setRetireDate(e.target.value); setResult(null); }} style={iStyle} />
          {joinDate && retireDate && (() => {
            const days = Math.floor((new Date(retireDate) - new Date(joinDate)) / (1000*60*60*24));
            return <div style={{ fontSize: 12, color: T.muted, marginTop: 8, textAlign: "center" }}>
              근속 <strong style={{ color: T.text }}>{days}일 ({(days/365).toFixed(1)}년)</strong>
            </div>;
          })()}
        </div>

        <button onClick={calc} disabled={!selUser || !retireDate || !joinDate || loading}
          style={{ width: "100%", padding: "14px 0", borderRadius: 14, border: "none", background: "#0891b2", color: "#fff", fontSize: 15, fontWeight: 800, cursor: "pointer", marginBottom: 16, opacity: (!selUser || !retireDate || !joinDate) ? 0.5 : 1 }}>
          {loading ? "계산중..." : "퇴직금 계산"}
        </button>

        {result && (
          <div>
            {/* 지급 요건 */}
            <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 12, border: `2px solid ${result.eligible ? "#16a34a" : T.red}` }}>
              <div style={{ fontSize: 13, fontWeight: 800, color: result.eligible ? "#16a34a" : T.red, marginBottom: 10 }}>
                {result.eligible ? "✅ 퇴직금 지급 대상" : "❌ 퇴직금 지급 불가"}
              </div>
              {[
                [result.eligible1year, `1년 이상 근무`, `근속 ${result.workDays}일 (${result.workYears.toFixed(1)}년) — ${result.eligible1year ? "충족" : "미충족 (1년 미만)"}`],
                [result.eligible15h, `주 15시간 이상`, `소정근로시간 주 ${result.weeklyHours}시간 — ${result.eligible15h ? "충족" : "미충족 (15시간 미만)"}`],
              ].map(([ok, label, desc]) => (
                <div key={label} style={{ display: "flex", gap: 8, marginBottom: 6, alignItems: "flex-start" }}>
                  <span style={{ fontSize: 14 }}>{ok ? "✅" : "❌"}</span>
                  <div>
                    <div style={{ fontSize: 13, fontWeight: 700, color: ok ? T.text : T.red }}>{label}</div>
                    <div style={{ fontSize: 11, color: T.muted }}>{desc}</div>
                  </div>
                </div>
              ))}
            </div>

            {result.eligible && <>
              {/* 평균임금 계산 */}
              <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 12, border: `1px solid ${T.border}` }}>
                <div style={{ fontSize: 13, color: T.muted, fontWeight: 700, marginBottom: 10 }}>📊 평균임금 계산 (퇴직 전 3개월)</div>
                {result.recentWages.length === 0
                  ? <div style={{ fontSize: 12, color: T.orange }}>⚠ 급여 확정 데이터 없음</div>
                  : result.recentWages.map(w => (
                    <div key={w.yearMonth} style={{ display: "flex", justifyContent: "space-between", padding: "4px 0", borderBottom: `1px solid ${T.border}` }}>
                      <span style={{ fontSize: 12, color: T.muted }}>{monthLabel(w.yearMonth)} (상여 {Number(w.bonus||0).toLocaleString()}원 제외)</span>
                      <span style={{ fontSize: 12, fontWeight: 600 }}>{(Number(w.totalIncome||0)-Number(w.bonus||0)).toLocaleString()}원</span>
                    </div>
                  ))
                }
                <div style={{ marginTop: 8 }}>
                  {[
                    ["기본임금 3개월 합계", result.pay3],
                    ["연간 상여금", result.annualBonus, `(연간 합계 ÷ 12 × 3 = ${result.bonusFor3.toLocaleString()}원)`],
                    ["3개월 임금총액", result.totalPay3],
                  ].map(([l,v,c]) => (
                    <div key={l} style={{ display: "flex", justifyContent: "space-between", padding: "4px 0" }}>
                      <div><span style={{ fontSize: 12, color: T.muted }}>{l}</span>{c && <span style={{ fontSize: 10, color: T.muted }}> {c}</span>}</div>
                      <span style={{ fontSize: 12, fontWeight: 700 }}>{Number(v).toLocaleString()}원</span>
                    </div>
                  ))}
                  <div style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", borderTop: `1px solid ${T.border}`, marginTop: 4 }}>
                    <span style={{ fontSize: 13, fontWeight: 700 }}>1일 평균임금</span>
                    <span style={{ fontSize: 14, fontWeight: 800, color: "#0891b2" }}>{result.avgDailyWage.toLocaleString()}원</span>
                  </div>
                  <div style={{ fontSize: 10, color: T.muted }}>= 임금총액 {result.totalPay3.toLocaleString()} ÷ 달력 {result.calDays3}일</div>
                </div>
              </div>

              {/* 퇴직금 계산 */}
              <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 12, border: `1px solid ${T.border}` }}>
                <div style={{ fontSize: 13, color: T.muted, fontWeight: 700, marginBottom: 10 }}>💰 퇴직금</div>
                <Row label="퇴직금" value={result.severancePay}
                  calc={`평균임금 ${result.avgDailyWage.toLocaleString()} × 30 × (${result.workDays}일 ÷ 365)`} bold />
                <Row label="잔여연차수당" value={result.annualAllowance}
                  calc={`총연차 ${result.totalAnnual}일 × ${result.workedDaysThisYear}일/365 = ${result.earnedThisYear}일 발생, 사용 ${result.usedAnnual}일, 잔여 ${result.annualRemain}일 × 시급×8`} />
              </div>

              {/* 퇴직소득세 계산식 */}
              <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 12, border: `1px solid ${T.border}` }}>
                <div style={{ fontSize: 13, color: T.muted, fontWeight: 700, marginBottom: 10 }}>🧮 퇴직소득세 계산식</div>
                {[
                  [`① 퇴직급여`, result.step1, ``],
                  [`② 근속연수공제 (${result.tenureYears}년)`, result.step2, `근속연수별 공제액`],
                  [`③ 과세표준 (①-②)`, result.step3, `${result.step1.toLocaleString()} - ${result.step2.toLocaleString()}`],
                  [`④ 환산급여 (③×12÷근속연수)`, result.step4, `${result.step3.toLocaleString()} × 12 ÷ ${result.tenureYears}`],
                  [`⑤ 환산급여공제`, result.step5, `구간별 공제율 적용`],
                  [`⑥ 환산과세표준 (④-⑤)`, result.step6, `${result.step4.toLocaleString()} - ${result.step5.toLocaleString()}`],
                  [`⑦ 환산산출세액 (세율적용)`, result.step7, `소득세율 구간 적용`],
                  [`⑧ 퇴직소득세 (⑦×근속÷12)`, result.retirementTax, `${result.step7.toLocaleString()} × ${result.tenureYears} ÷ 12`],
                ].map(([l,v,c]) => (
                  <div key={l} style={{ padding: "5px 0", borderBottom: `1px solid ${T.border}` }}>
                    <div style={{ display: "flex", justifyContent: "space-between" }}>
                      <span style={{ fontSize: 12, color: T.text }}>{l}</span>
                      <span style={{ fontSize: 12, fontWeight: 700 }}>{Number(v||0).toLocaleString()}원</span>
                    </div>
                    {c && <div style={{ fontSize: 10, color: T.muted }}>{c}</div>}
                  </div>
                ))}
                <div style={{ padding: "5px 0" }}>
                  <div style={{ display: "flex", justifyContent: "space-between" }}>
                    <span style={{ fontSize: 12, color: T.text }}>지방소득세 (퇴직소득세×10%)</span>
                    <span style={{ fontSize: 12, fontWeight: 700 }}>{result.localTax.toLocaleString()}원</span>
                  </div>
                </div>
              </div>

              {/* 최종 지급 */}
              <div style={{ background: "#ecfeff", borderRadius: 16, padding: 20, border: "2px solid #0891b2", marginBottom: 20 }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
                  <span style={{ fontSize: 13, color: T.muted }}>퇴직금</span>
                  <span style={{ fontSize: 13, fontWeight: 700 }}>{result.severancePay.toLocaleString()}원</span>
                </div>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
                  <span style={{ fontSize: 13, color: T.muted }}>잔여연차수당</span>
                  <span style={{ fontSize: 13, fontWeight: 700 }}>+{result.annualAllowance.toLocaleString()}원</span>
                </div>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
                  <span style={{ fontSize: 13, color: T.muted }}>퇴직소득세</span>
                  <span style={{ fontSize: 13, fontWeight: 700, color: T.red }}>-{result.retirementTax.toLocaleString()}원</span>
                </div>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 12 }}>
                  <span style={{ fontSize: 13, color: T.muted }}>지방소득세</span>
                  <span style={{ fontSize: 13, fontWeight: 700, color: T.red }}>-{result.localTax.toLocaleString()}원</span>
                </div>
                <div style={{ borderTop: "1px solid #0891b244", paddingTop: 12, textAlign: "center" }}>
                  <div style={{ fontSize: 13, color: "#0891b2", fontWeight: 700, marginBottom: 6 }}>실 지급액</div>
                  <div style={{ fontSize: 30, fontWeight: 800, color: "#0891b2" }}>{result.netPay.toLocaleString()}원</div>
                  <div style={{ fontSize: 11, color: T.muted, marginTop: 6 }}>근속 {result.workDays}일 ({result.workYears.toFixed(1)}년)</div>
                </div>
              </div>
            </>}
          </div>
        )}
      </div>
    </div>
  );
}

// ── 관리자 대문 ────────────────────────────────────────────────
function AdminHome({ user, onLogout, onSection, leaveRequests = [], board = [], reads = {} }) {
  const now = new Date();
  const kst = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  const dateStr = kst.toLocaleDateString("ko-KR", { year: "numeric", month: "long", day: "numeric", weekday: "long" });

  const sections = [
    { key: "attendance", icon: "📋", label: "근태",   desc: "출퇴근 현황 · 월별 기록", color: "#2563eb" },
    { key: "wage",       icon: "💰", label: "임금",   desc: "급여 계산 · 임금대장",   color: "#16a34a" },
    { key: "members",    icon: "👥", label: "팀원",   desc: "직원 정보 · 기초 데이터", color: "#7c3aed" },
    { key: "annual",     icon: "📅", label: "연차",   desc: "연차 현황 · 신청 승인",   color: "#0284c7", badge: leaveRequests.filter(r => r.status === "대기").length },
    { key: "notice",     icon: "📢", label: "공지",   desc: "공지사항 작성 · 관리",   color: "#ea580c" },
    { key: "board",      icon: "💬", label: "게시판", desc: "자유게시판",              color: "#0891b2",
      badge: board.filter(b => !reads[`${user.id}_board_${b.id}`]).length },
    { key: "settings",   icon: "⚙",  label: "설정",   desc: "근무시간 · GPS · 공휴일", color: "#6b7280" },
    { key: "reminder",   icon: "🔔", label: "리마인더", desc: "반복 알림 · 일정 관리",  color: "#7c3aed" },
    { key: "severance",  icon: "💼", label: "퇴직금", desc: "퇴직금 계산",            color: "#b45309" },
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
      <div style={{ padding: "10px 16px 24px" }}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          {sections.map(s => (
            <button key={s.key} onClick={() => onSection(s.key)}
              style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 16, padding: "14px 14px", cursor: "pointer", textAlign: "left", boxShadow: "0 2px 8px #0000000d", transition: "transform .1s", position: "relative", minHeight: 100 }}
              onMouseDown={e => e.currentTarget.style.transform = "scale(0.97)"}
              onMouseUp={e => e.currentTarget.style.transform = "scale(1)"}
              onTouchStart={e => e.currentTarget.style.transform = "scale(0.97)"}
              onTouchEnd={e => e.currentTarget.style.transform = "scale(1)"}>
              {s.badge > 0 && (
                <div style={{ position: "absolute", top: 8, right: 8, background: "#ef4444", color: "#fff", borderRadius: 10, padding: "1px 6px", fontSize: 10, fontWeight: 800 }}>{s.badge}</div>
              )}
              <div style={{ fontSize: 26, marginBottom: 6 }}>{s.icon}</div>
              <div style={{ fontSize: 15, fontWeight: 800, color: s.color, marginBottom: 2 }}>{s.label}</div>
              <div style={{ fontSize: 11, color: T.muted, lineHeight: 1.4 }}>{s.desc}</div>
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
            const todayLeave = (leaves[u.id] || {})[today];
            const lm = calcLateMinWithLeave(rec.in, settings.workStart, todayLeave, settings);
            const em = calcEarlyOutMinWithLeave(rec.out, settings.workEnd, rec.in, settings.workStart, todayLeave, settings);
            const om = calcTotalOvertimeMin(rec.in, rec.out, settings.workStart, settings.workEnd);
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
                  {todayLeave && !todayLeave.deleted && <Badge label={todayLeave.type || "연차"} color="purple" />}
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
// ── 지급일 계산 ─────────────────────────────────────────────
function getPayDate(yearMonth, holidays = []) {
  const [y, m] = yearMonth.split("-").map(Number);
  // 당월 급여 → 익월 15일 (KST 기준)
  let d = new Date(y, m, 15); // JS month는 0-based, m은 1-based이므로 m = 익월
  while (d.getDay() === 0 || d.getDay() === 6 || holidays.includes(
    `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`
  )) {
    d.setDate(d.getDate() + 1);
  }
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}

// ── 급여 계산 모달 ───────────────────────────────────────────
function WageModal({ user, info, monthStats, yearMonth, existing, holidays, annualData, settings, onClose, onSave }) {
  const monthlyHours = Number(settings?.monthlyHours ?? 209);
  const hourlyWage = Number(info?.hourlyWage || 0);
  const dailyWage = Math.round(hourlyWage * 8);
  const monthlyBase = Math.round(hourlyWage * monthlyHours);
  const otPay = Math.round(hourlyWage * 1.5 * (monthStats.otMin / 60));
  const holidayPay = Math.round(dailyWage * 1.5 * (monthStats.holiday || 0));
  const deductMin = (monthStats.lateMin || 0) + (monthStats.earlyMin || 0) + (monthStats.outingMin || 0);
  const deductPay = Math.round(hourlyWage * (deductMin / 60));

  // settings에서 요율 읽기 (없으면 기본값)
  const ratePension    = Number(settings?.ratePension    ?? 4.75);
  const rateHealth     = Number(settings?.rateHealth     ?? 3.595);
  const rateEmployment = Number(settings?.rateEmployment ?? 0.9);
  const rateLongCare   = Number(settings?.rateLongCare   ?? 13.14);

  const pensionBase = Number(info?.pensionBase || 0);
  const insuranceBase = Number(info?.insuranceBase || 0);
  const incomeTax = Number(info?.incomeTax || 0);
  const residentTax    = Math.floor(incomeTax * 0.1 / 10) * 10;
  const nationalPension = Math.floor(pensionBase * (ratePension / 100) / 10) * 10;
  const health          = Math.floor(insuranceBase * (rateHealth / 100) / 10) * 10;
  const employment      = Math.floor(insuranceBase * (rateEmployment / 100) / 10) * 10;
  const longCare        = Math.floor(health * (rateLongCare / 100) / 10) * 10;

  // 결근 공제 = 결근일수 × 일급 + 결근있는주수 × 일급(주휴)
  const absentDays = monthStats.absentDays || 0;
  const absentWeeks = monthStats.absentWeeks || 0;
  const absentPay = Math.round(dailyWage * (absentDays + absentWeeks));
  const isDecember = yearMonth?.slice(5, 7) === "12";
  const annualRemain = Math.max(0, (annualData?.total || 0) - (annualData?.used || 0));
  const autoAnnualPay = isDecember ? Math.round(dailyWage * annualRemain) : 0;

  const [form, setForm] = useState({
    bonus: existing?.bonus || 0,
    annualPay: existing?.annualPay !== undefined ? existing.annualPay : autoAnnualPay,
    carryOver: existing?.carryOver || 0,
    otherIncome: existing?.otherIncome || 0,
    otherDeduct: existing?.otherDeduct || 0,
    memo: existing?.memo || "",
  });

  const totalIncome = monthlyBase + otPay + holidayPay +
    Number(form.bonus) + Number(form.annualPay) + Number(form.carryOver) + Number(form.otherIncome);
  const totalDeduct = incomeTax + residentTax + nationalPension + health + employment + longCare +
    deductPay + absentPay + Number(form.otherDeduct);
  const netPay = totalIncome - totalDeduct;

  const iStyle = { width: "100%", padding: "9px 12px", borderRadius: 8, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 14, fontWeight: 600, boxSizing: "border-box", textAlign: "right", fontFamily: "inherit" };

  const Row = ({ label, value, color, bold, sub }) => (
    <div style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", borderBottom: `1px solid ${T.border}` }}>
      <span style={{ fontSize: sub ? 11 : 13, color: sub ? T.muted : T.text }}>{label}</span>
      <span style={{ fontSize: 13, fontWeight: bold ? 800 : 600, color: color || T.text }}>{Number(value || 0).toLocaleString()}원</span>
    </div>
  );

  return (
    <div style={{ position: "fixed", inset: 0, background: "#00000066", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, padding: 12 }}>
      <div style={{ background: T.card, borderRadius: 20, width: "100%", maxWidth: 360, maxHeight: "92vh", overflowY: "scroll", WebkitOverflowScrolling: "touch", boxShadow: "0 20px 60px #00000030" }}>
        <div style={{ background: "#16a34a", padding: "16px 20px", borderRadius: "20px 20px 0 0" }}>
          <div style={{ fontWeight: 800, fontSize: 17, color: "#fff" }}>{user.name} — {monthLabel(yearMonth)}</div>
          <div style={{ fontSize: 12, color: "#ffffff80", marginTop: 4 }}>지급일 {getPayDate(yearMonth, holidays)} · {info?.bank ? `${info.bank}은행 ${info.account}` : "계좌 미등록"}</div>
        </div>

        <div style={{ padding: "16px 20px" }}>
          {/* 근태 요약 */}
          <div style={{ background: T.bg, borderRadius: 12, padding: "12px 14px", marginBottom: 14 }}>
            <div style={{ fontSize: 12, color: T.muted, fontWeight: 700, marginBottom: 8 }}>근태 내역</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 6 }}>
              {[
                ["출근", monthStats.days + "일"],
                ["연장", fmtMinutes(monthStats.otMin)],
                ["휴일근무", monthStats.holiday + "일"],
                ["지각", fmtMinutes(monthStats.lateMin)],
                ["조퇴", fmtMinutes(monthStats.earlyMin)],
                ["외출", fmtMinutes(monthStats.outingMin || 0)],
                ["결근", (monthStats.absentDays||0) + "일"],
                ["휴무일", (monthStats.offDays||0) + "일"],
                ["전체", (monthStats.totalDays||0) + "일"],
              ].map(([l, v]) => (
                <div key={l} style={{ textAlign: "center" }}>
                  <div style={{ fontSize: 10, color: T.muted }}>{l}</div>
                  <div style={{ fontSize: 13, fontWeight: 700, color: T.text }}>{v}</div>
                </div>
              ))}
            </div>
          </div>

          {/* 소득 */}
          <div style={{ marginBottom: 14 }}>
            <div style={{ fontSize: 12, color: "#16a34a", fontWeight: 700, marginBottom: 8 }}>소득 내역</div>
            <Row label={`기본급 (${hourlyWage.toLocaleString()}×${monthlyHours})`} value={monthlyBase} />
            {otPay > 0 && <Row label={`연장수당 (×1.5, ${fmtMinutes(monthStats.otMin)})`} value={otPay} />}
            {holidayPay > 0 && <Row label={`휴일수당 (×1.5, ${monthStats.holiday}일)`} value={holidayPay} />}
            {[["상여금", "bonus"], ["이월분", "carryOver"], ["기타", "otherIncome"]].map(([label, key]) => (
              <div key={key} style={{ padding: "5px 0", borderBottom: `1px solid ${T.border}` }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
                  <span style={{ fontSize: 13, color: T.text, whiteSpace: "nowrap" }}>{label}</span>
                  <input type="number" value={form[key]} onChange={e => setForm(p => ({...p, [key]: e.target.value}))} style={{ ...iStyle, width: 130 }} placeholder="0" />
                </div>
              </div>
            ))}
            {/* 12월 연차수당 */}
            {/* 연차수당 - 항상 표시, 12월만 활성화 */}
            <div style={{ padding: "5px 0", borderBottom: `1px solid ${T.border}` }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
                <div>
                  <span style={{ fontSize: 13, color: T.text, whiteSpace: "nowrap" }}>연차수당</span>
                  <div style={{ fontSize: 10, color: T.muted }}>
                    {isDecember ? `잔여 ${annualRemain}일 × 일급 ${dailyWage.toLocaleString()} (수정가능)` : `잔여 ${annualRemain}일 · 12월 급여에 반영`}
                  </div>
                </div>
                <input type="number" value={form.annualPay} onChange={e => setForm(p => ({...p, annualPay: e.target.value}))}
                  disabled={!isDecember}
                  style={{ ...iStyle, width: 130, background: isDecember ? "#fff" : T.bg, color: isDecember ? T.text : T.muted, cursor: isDecember ? "auto" : "not-allowed" }} placeholder="0" />
              </div>
            </div>
            <Row label="소득 합계" value={totalIncome} bold color={T.blue} />
          </div>

          {/* 공제 */}
          <div style={{ marginBottom: 14 }}>
            <div style={{ fontSize: 12, color: T.red, fontWeight: 700, marginBottom: 8 }}>공제 내역</div>
            {deductPay > 0 && <Row label={`지각/조퇴/외출 차감 (${fmtMinutes(deductMin)})`} value={deductPay} color={T.orange} />}
            {absentPay > 0 && <Row label={`결근공제 (${absentDays}일×일급 + ${absentWeeks}주×주휴)`} value={absentPay} color={T.red} />}
            <Row label={`소득세`} value={incomeTax} sub />
            <Row label={`주민세 (소득세×10%)`} value={residentTax} sub />
            <Row label={`국민연금 (${pensionBase.toLocaleString()}×${ratePension}%)`} value={nationalPension} sub />
            <Row label={`건강보험 (${insuranceBase.toLocaleString()}×${rateHealth}%)`} value={health} sub />
            <Row label={`고용보험 (${insuranceBase.toLocaleString()}×${rateEmployment}%)`} value={employment} sub />
            <Row label={`장기요양 (건강보험×${rateLongCare}%)`} value={longCare} sub />
            <div style={{ padding: "5px 0", borderBottom: `1px solid ${T.border}` }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
                <span style={{ fontSize: 13, color: T.text, whiteSpace: "nowrap" }}>기타 공제</span>
                <input type="number" value={form.otherDeduct} onChange={e => setForm(p => ({...p, otherDeduct: e.target.value}))} style={{ ...iStyle, width: 130 }} placeholder="0" />
              </div>
            </div>
            <Row label="공제 합계" value={totalDeduct} bold color={T.red} />
          </div>

          {/* 실지급액 */}
          <div style={{ background: "#16a34a18", borderRadius: 12, padding: "14px 16px", marginBottom: 14, textAlign: "center" }}>
            <div style={{ fontSize: 13, color: T.muted, marginBottom: 4 }}>실 지급액</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: "#16a34a" }}>{netPay.toLocaleString()}원</div>
          </div>

          {/* 메모 */}
          <div style={{ marginBottom: 16 }}>
            <div style={{ fontSize: 12, color: T.muted, fontWeight: 700, marginBottom: 6 }}>메모</div>
            <textarea value={form.memo} onChange={e => setForm(p => ({...p, memo: e.target.value}))} rows={2}
              style={{ ...iStyle, textAlign: "left", resize: "none", lineHeight: 1.6 }} placeholder="특이사항 등" />
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
            <Btn variant="ghost" onClick={onClose}>취소</Btn>
            <Btn variant="green" onClick={() => onSave({
              userId: user.id, userName: user.name, yearMonth,
              hourlyWage, monthlyHours, monthlyBase, otPay, holidayPay, deductPay, deductMin,
              absentDays, absentWeeks, absentPay,
              ...form,
              annualRemain: isDecember ? annualRemain : 0,
              incomeTax, residentTax, nationalPension, health, employment, longCare,
              pensionBase, insuranceBase,
              ratePension, rateHealth, rateEmployment, rateLongCare,
              totalIncome, totalDeduct, netPay,
              monthStats, payDate: getPayDate(yearMonth, holidays),
              createdAt: new Date().toISOString()
            })}>저장 + 확정</Btn>
          </div>
        </div>
      </div>
    </div>
  );
}

// ── 관리자 임금 섹션 ──────────────────────────────────────────
function AdminWage({ users, records, leaves, settings, memberInfo, annual, leaveRequests, payslips, reads, onBack }) {
  const members = users.filter(u => u.role === "member");
  const kst = new Date(new Date().getTime() + 9 * 60 * 60 * 1000);
  const [selectedMonth, setSelectedMonth] = useState(kst.toISOString().slice(0, 7));
  const [wageModal, setWageModal] = useState(null);
  const [savedWages, setSavedWages] = useState({});
  const [tab, setTab] = useState("calc"); // calc | ledger | payslip
  const [sending, setSending] = useState(null); // userId

  // 저장된 급여 로드
  useEffect(() => {
    const loadWages = async () => {
      const snap = await getDocs(collection(db, "wages"));
      const w = {};
      snap.docs.forEach(d => { w[d.id] = d.data(); });
      setSavedWages(w);
    };
    loadWages();
  }, [selectedMonth]);

  const prevMonth = () => {
    const [y, m] = selectedMonth.split("-").map(Number);
    const d = new Date(y, m - 2, 1);
    setSelectedMonth(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`);
  };
  const nextMonth = () => {
    const [y, m] = selectedMonth.split("-").map(Number);
    const kstNow = new Date(new Date().getTime() + 9 * 60 * 60 * 1000);
    const cur = `${kstNow.getFullYear()}-${String(kstNow.getMonth() + 1).padStart(2, "0")}`;
    if (selectedMonth >= cur) return;
    const d = new Date(y, m, 1);
    setSelectedMonth(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`);
  };

  const getMonthStats = (userId) => {
    const days = Object.entries(records[userId] || {}).filter(([d]) => d.startsWith(selectedMonth));
    const uLeaves = Object.fromEntries(Object.entries(leaves[userId] || {}).filter(([d]) => d.startsWith(selectedMonth)));
    const ms = calcMonthStats(days, settings, uLeaves, leaveRequests, userId, selectedMonth);
    // 외출 시간 합산
    let outingMin = 0;
    days.forEach(([, rec]) => {
      (rec.outing || []).forEach(o => {
        if (o.out && o.in) outingMin += (new Date(o.in) - new Date(o.out)) / 60000;
      });
    });
    ms.outingMin = Math.round(outingMin);
    return ms;
  };

  const saveWage = async (data) => {
    const key = `${data.userId}_${data.yearMonth}`;
    await setDoc(doc(db, "wages", key), data);
    setSavedWages(p => ({ ...p, [key]: data }));
    setWageModal(null);
  };

  // 명세서 팀원에게 전송
  const sendPayslip = async (userId, userName, saved) => {
    setSending(userId);
    try {
      // wages 데이터를 payslips 컬렉션에 저장 (팀원 명세서 탭에 표시)
      const key = `${userId}_${selectedMonth}`;
      await setDoc(doc(db, COL_PAYSLIPS, key), {
        userId, month: selectedMonth,
        fileName: `${monthLabel(selectedMonth)} 급여명세서`,
        wageData: saved,
        isAuto: true, // 자동 생성 명세서
        uploadedBy: "관리자",
        createdAt: new Date().toISOString()
      });
      // 공지 발송
      await addDoc(collection(db, COL_NOTICES), {
        title: `💰 ${monthLabel(selectedMonth)} 급여명세서 발급`,
        content: `${monthLabel(selectedMonth)} 급여명세서가 발급되었습니다.\n명세서 탭에서 확인해주세요.\n\n실 지급액: ${saved.netPay?.toLocaleString()}원\n지급일: ${saved.payDate}`,
        recipient: userId, author: "관리자",
        createdAt: new Date().toISOString(), auto: true
      });
      await sendPush({ title: `💰 급여명세서 발급`, message: `${monthLabel(selectedMonth)} 급여명세서가 발급되었습니다. 명세서 탭에서 확인해주세요.`, targetUserId: userId });
      alert(`${userName}님께 명세서가 전송됐어요!`);
    } catch(e) { alert("전송 실패: " + e.message); }
    setSending(null);
  };

  // 임금대장 CSV
  const downloadLedger = () => {
    const header = ["이름", "지급일", "출근", "연장", "휴일", "기본급", "연장수당", "휴일수당", "상여금", "이월분", "기타소득", "소득합계", "소득세", "주민세", "국민연금", "건강보험", "고용보험", "장기요양", "차감", "기타공제", "공제합계", "실지급액"];
    const rows = members.map(u => {
      const s = savedWages[`${u.id}_${selectedMonth}`];
      if (!s) return [u.name, getPayDate(selectedMonth), ...Array(20).fill("-")];
      return [u.name, s.payDate, s.monthStats?.days||0, fmtMinutes(s.monthStats?.otMin||0), s.monthStats?.holiday||0,
        s.monthlyBase||0, s.otPay||0, s.holidayPay||0, s.bonus||0, s.carryOver||0, s.otherIncome||0, s.totalIncome||0,
        s.incomeTax||0, s.residentTax||0, s.nationalPension||0, s.health||0, s.employment||0, s.longCare||0,
        s.deductPay||0, s.otherDeduct||0, s.totalDeduct||0, s.netPay||0];
    });
    downloadCSV(`임금대장_${monthLabel(selectedMonth)}.csv`, [header, ...rows]);
  };

  const isCurrentMonth = selectedMonth >= new Date(new Date().getTime() + 9 * 60 * 60 * 1000).toISOString().slice(0, 7);

  return (
    <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'Noto Sans KR',sans-serif" }}>
      <div style={{ background: "#16a34a", padding: "16px 16px 0" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
          <button onClick={onBack} style={{ background: "#ffffff18", border: "none", color: "#fff", fontSize: 18, cursor: "pointer", padding: "8px 14px", borderRadius: 12, fontWeight: 700 }}>‹</button>
          <div>
            <div style={{ fontSize: 11, color: "#ffffff40", letterSpacing: 3 }}>ADMIN</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>💰 임금</div>
          </div>
        </div>
        <div style={{ display: "flex" }}>
          {[["calc","급여 계산"], ["ledger","임금대장"], ["payslip","명세서"]].map(([key, label]) => (
            <button key={key} onClick={() => setTab(key)}
              style={{ padding: "10px 24px", border: "none", background: "none", color: tab===key?"#fff":"#ffffff60", fontWeight: tab===key?800:500, fontSize: 14, cursor: "pointer", borderBottom: tab===key?"3px solid #fff":"3px solid transparent", fontFamily: "inherit" }}>
              {label}
            </button>
          ))}
        </div>
      </div>

      <div style={{ padding: 16 }}>
        {/* 월 선택 */}
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 12 }}>
          <button onClick={prevMonth} style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 10, padding: "10px 16px", fontSize: 16, cursor: "pointer", fontWeight: 700, color: T.text }}>‹</button>
          <div style={{ flex: 1, textAlign: "center", fontSize: 16, fontWeight: 800, color: T.text }}>{monthLabel(selectedMonth)}</div>
          <button onClick={nextMonth} style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 10, padding: "10px 16px", fontSize: 16, cursor: "pointer", fontWeight: 700, color: isCurrentMonth ? T.muted : T.text, opacity: isCurrentMonth ? 0.3 : 1 }}>›</button>
        </div>
        <div style={{ fontSize: 12, color: T.muted, marginBottom: 14, textAlign: "center" }}>
          지급일 <strong style={{ color: T.text }}>{getPayDate(selectedMonth)}</strong>
        </div>

        {/* 급여 계산 탭 */}
        {tab === "calc" && members.map(u => {
          const info = memberInfo[u.id] || {};
          const ms = getMonthStats(u.id);
          const key = `${u.id}_${selectedMonth}`;
          const saved = savedWages[key];
          const hourlyWage = Number(info.hourlyWage || 0);
          return (
            <div key={u.id} style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 12, border: `1px solid ${saved ? "#16a34a44" : T.border}`, boxShadow: "0 1px 4px #0000000a" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 12 }}>
                <div style={{ width: 44, height: 44, borderRadius: "50%", background: "#16a34a", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 17, fontWeight: 800, color: "#fff" }}>{u.name[0]}</div>
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 800, fontSize: 16, color: T.text }}>{u.name}</div>
                  <div style={{ fontSize: 12, color: T.muted }}>{hourlyWage ? `시급 ${hourlyWage.toLocaleString()}원` : "⚠ 시급 미입력"}</div>
                </div>
                <Badge label={saved ? "확정" : "미확정"} color={saved ? "green" : "gray"} />
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 6, marginBottom: 12 }}>
                {[["출근", ms.days+"일", T.green], ["연장", fmtMinutes(ms.otMin), T.purple], ["휴일", ms.holiday+"일", T.red]].map(([l,v,c]) => (
                  <StatBox key={l} label={l} value={v} color={c} />
                ))}
              </div>
              {saved && (
                <div style={{ background: T.bg, borderRadius: 10, padding: "10px 14px", marginBottom: 10 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                    <span style={{ fontSize: 12, color: T.muted }}>소득 합계</span>
                    <span style={{ fontSize: 13, fontWeight: 700, color: T.blue }}>{saved.totalIncome?.toLocaleString()}원</span>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                    <span style={{ fontSize: 12, color: T.muted }}>공제 합계</span>
                    <span style={{ fontSize: 13, fontWeight: 700, color: T.red }}>{saved.totalDeduct?.toLocaleString()}원</span>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between" }}>
                    <span style={{ fontSize: 13, fontWeight: 700, color: T.text }}>실 지급액</span>
                    <span style={{ fontSize: 15, fontWeight: 800, color: "#16a34a" }}>{saved.netPay?.toLocaleString()}원</span>
                  </div>
                </div>
              )}
              {!hourlyWage && !saved && (
                <div style={{ fontSize: 12, color: T.orange, marginBottom: 10, padding: "8px 12px", background: "#fff7ed", borderRadius: 8 }}>
                  팀원 섹션에서 시급을 먼저 입력해주세요
                </div>
              )}
              <div style={{ display: "grid", gridTemplateColumns: saved ? "1fr 1fr" : "1fr", gap: 8 }}>
                <button onClick={() => setWageModal({ user: u })}
                  style={{ padding: "10px 0", borderRadius: 10, border: "none", background: saved ? "#16a34a22" : "#16a34a", color: saved ? "#16a34a" : "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                  {saved ? "✏ 수정" : "급여 계산"}
                </button>
                {saved && (
                  <button onClick={() => sendPayslip(u.id, u.name, saved)} disabled={sending === u.id}
                    style={{ padding: "10px 0", borderRadius: 10, border: "none", background: "#2563eb", color: "#fff", fontSize: 13, fontWeight: 700, cursor: sending === u.id ? "not-allowed" : "pointer", opacity: sending === u.id ? 0.6 : 1 }}>
                    {sending === u.id ? "전송중..." : "📤 명세서 전송"}
                  </button>
                )}
              </div>
            </div>
          );
        })}

        {/* 임금대장 탭 */}
        {tab === "ledger" && (
          <div>
            <button onClick={downloadLedger}
              style={{ width: "100%", padding: "11px 0", borderRadius: 12, border: "none", background: T.green, color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer", marginBottom: 14 }}>
              ⬇ 임금대장 CSV 다운로드
            </button>
            {members.map(u => {
              const s = savedWages[`${u.id}_${selectedMonth}`];
              return (
                <div key={u.id} style={{ background: T.card, borderRadius: 14, padding: "14px 16px", marginBottom: 10, border: `1px solid ${s ? "#16a34a44" : T.border}` }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: s ? 10 : 0 }}>
                    <div style={{ width: 36, height: 36, borderRadius: "50%", background: "#16a34a", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14, fontWeight: 800, color: "#fff" }}>{u.name[0]}</div>
                    <div style={{ flex: 1, fontWeight: 800, fontSize: 15, color: T.text }}>{u.name}</div>
                    <Badge label={s ? "확정" : "미확정"} color={s ? "green" : "gray"} />
                  </div>
                  {s ? (
                    <div>
                      {/* 소득 */}
                      <div style={{ fontSize: 11, color: T.blue, fontWeight: 700, marginBottom: 4 }}>소득</div>
                      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 2, marginBottom: 8 }}>
                        {[["기본급", s.monthlyBase], ["연장수당", s.otPay], ["휴일수당", s.holidayPay],
                          ["상여금", s.bonus], ["연차수당", s.annualPay], ["이월분", s.carryOver], ["기타", s.otherIncome],
                        ].filter(([,v]) => Number(v) > 0).map(([l,v]) => (
                          <div key={l} style={{ display: "flex", justifyContent: "space-between", padding: "3px 6px", background: "#eff6ff", borderRadius: 4 }}>
                            <span style={{ fontSize: 11, color: T.muted }}>{l}</span>
                            <span style={{ fontSize: 11, fontWeight: 600, color: T.blue }}>{Number(v).toLocaleString()}</span>
                          </div>
                        ))}
                        <div style={{ gridColumn: "span 2", display: "flex", justifyContent: "space-between", padding: "4px 6px", background: "#dbeafe", borderRadius: 4 }}>
                          <span style={{ fontSize: 12, fontWeight: 700, color: T.blue }}>소득 합계</span>
                          <span style={{ fontSize: 12, fontWeight: 800, color: T.blue }}>{Number(s.totalIncome||0).toLocaleString()}</span>
                        </div>
                      </div>
                      {/* 공제 */}
                      <div style={{ fontSize: 11, color: T.red, fontWeight: 700, marginBottom: 4 }}>공제</div>
                      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 2, marginBottom: 8 }}>
                        {[["소득세", s.incomeTax], ["주민세", s.residentTax], ["국민연금", s.nationalPension],
                          ["건강보험", s.health], ["고용보험", s.employment], ["장기요양", s.longCare],
                          ["지각/조퇴차감", s.deductPay], ["결근공제", s.absentPay], ["기타공제", s.otherDeduct],
                        ].filter(([,v]) => Number(v) > 0).map(([l,v]) => (
                          <div key={l} style={{ display: "flex", justifyContent: "space-between", padding: "3px 6px", background: "#fff1f2", borderRadius: 4 }}>
                            <span style={{ fontSize: 11, color: T.muted }}>{l}</span>
                            <span style={{ fontSize: 11, fontWeight: 600, color: T.red }}>{Number(v).toLocaleString()}</span>
                          </div>
                        ))}
                        <div style={{ gridColumn: "span 2", display: "flex", justifyContent: "space-between", padding: "4px 6px", background: "#ffe4e6", borderRadius: 4 }}>
                          <span style={{ fontSize: 12, fontWeight: 700, color: T.red }}>공제 합계</span>
                          <span style={{ fontSize: 12, fontWeight: 800, color: T.red }}>{Number(s.totalDeduct||0).toLocaleString()}</span>
                        </div>
                      </div>
                      {/* 실지급액 */}
                      <div style={{ display: "flex", justifyContent: "space-between", padding: "8px 10px", background: "#f0fdf4", borderRadius: 8, border: "1px solid #16a34a44" }}>
                        <span style={{ fontSize: 13, fontWeight: 700, color: T.text }}>실 지급액</span>
                        <span style={{ fontSize: 16, fontWeight: 800, color: "#16a34a" }}>{Number(s.netPay||0).toLocaleString()}원</span>
                      </div>
                    </div>
                  ) : (
                    <div style={{ fontSize: 12, color: T.muted, marginTop: 8 }}>급여 미확정</div>
                  )}
                </div>
              );
            })}
          </div>
        )}
        {tab === "payslip" && (
          <PayslipScreen user={users.find(u => u.role === "admin") || users[0]} users={users} payslips={payslips || []} reads={reads || {}} />
        )}
      </div>

      {wageModal && (
        <WageModal
          user={wageModal.user}
          info={memberInfo[wageModal.user.id] || {}}
          monthStats={getMonthStats(wageModal.user.id)}
          yearMonth={selectedMonth}
          existing={savedWages[`${wageModal.user.id}_${selectedMonth}`]}
          holidays={settings.holidays || []}
          annualData={annual[wageModal.user.id]}
          settings={settings}
          onClose={() => setWageModal(null)}
          onSave={saveWage}
        />
      )}
    </div>
  );
}

// ── 팀원 기초 데이터 편집 모달 ────────────────────────────────
function MemberInfoModal({ user, info, onSave, onClose }) {
  const BANKS = ["국민","신한","우리","하나","농협","기업","카카오","토스","새마을","우체국","SC제일","씨티","광주","전북","제주","경남","부산","대구","수협"];
  const [d, setD] = useState({
    empNo: "", ssn: "", joinDate: "", bank: "", account: "",
    hourlyWage: "", employType: "정규직", weeklyHours: 40, insurance: true,
    incomeTax: "",        // 소득세 (월 고정)
    pensionBase: "",      // 국민연금 기준소득월액
    insuranceBase: "",    // 건강/고용보험 보수월액
    ...info
  });
  const iStyle = { width: "100%", padding: "11px 14px", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 14, fontWeight: 600, boxSizing: "border-box", fontFamily: "inherit" };
  const numField = (label, key, placeholder, hint) => (
    <div style={{ marginBottom: 12 }}>
      <div style={{ fontSize: 12, color: T.sub, marginBottom: 3, fontWeight: 600 }}>{label}</div>
      {hint && <div style={{ fontSize: 11, color: T.muted, marginBottom: 5 }}>{hint}</div>}
      <input type="number" value={d[key]} onChange={e => setD(p => ({ ...p, [key]: e.target.value }))}
        placeholder={placeholder} style={iStyle} />
    </div>
  );

  return (
    <div style={{ position: "fixed", inset: 0, background: "#00000066", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, padding: 16 }}>
      <div style={{ background: T.card, borderRadius: 20, padding: 22, width: "100%", maxWidth: 340, maxHeight: "90vh", overflowY: "scroll", WebkitOverflowScrolling: "touch", boxShadow: "0 20px 60px #00000030" }}>
        <div style={{ fontWeight: 800, fontSize: 17, color: T.text, marginBottom: 4 }}>{user.name} — 기초 데이터</div>
        <div style={{ fontSize: 12, color: T.muted, marginBottom: 18 }}>임금 계산에 사용되는 정보예요</div>

        {/* 기본 정보 */}
        <div style={{ fontSize: 12, color: T.adminHeader, fontWeight: 700, marginBottom: 10 }}>📋 기본 정보</div>
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
          <div style={{ fontSize: 12, color: T.sub, marginBottom: 5, fontWeight: 600 }}>고용형태</div>
          <select value={d.employType} onChange={e => setD(p => ({ ...p, employType: e.target.value }))} style={iStyle}>
            {["정규직", "계약직", "파트타임"].map(t => <option key={t} value={t}>{t}</option>)}
          </select>
        </div>
        <div style={{ marginBottom: 12 }}>
          <div style={{ fontSize: 12, color: T.sub, marginBottom: 5, fontWeight: 600 }}>소정근로시간 (주, 시간)</div>
          <input type="number" value={d.weeklyHours} onChange={e => setD(p => ({ ...p, weeklyHours: Number(e.target.value) }))}
            placeholder="40" style={iStyle} />
        </div>
        <div style={{ marginBottom: 16 }}>
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

        {/* 임금 정보 */}
        <div style={{ fontSize: 12, color: "#16a34a", fontWeight: 700, marginBottom: 10, marginTop: 4 }}>💰 임금 정보</div>
        {numField("시급 (원)", "hourlyWage", "시급 입력")}
        {numField("소득세 (월 고정, 원)", "incomeTax", "0", "간이세액표 기준 수동 입력")}

        {/* 보험 과세표준 */}
        <div style={{ fontSize: 12, color: "#16a34a", fontWeight: 700, marginBottom: 10, marginTop: 4 }}>🏥 보험 과세표준</div>
        {numField("국민연금 기준소득월액 (원)", "pensionBase", "0", "× 4.5% = 국민연금")}
        {numField("건강/고용보험 보수월액 (원)", "insuranceBase", "0", "건강 ×3.595% / 고용 ×0.9%")}

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 8 }}>
          <Btn variant="ghost" onClick={onClose}>취소</Btn>
          <Btn variant="admin" onClick={() => onSave(d)}>저장</Btn>
        </div>
      </div>
    </div>
  );
}

// ── 관리자 팀원 섹션 ───────────────────────────────────────────
// ── 연차 신청 아이템 컴포넌트 ─────────────────────────────────
function LeaveRequestItem({ r, statusColor, setDelConfirm }) {
  const [processing, setProcessing] = useState(null);
  const [done, setDone] = useState(null);

  const sendNotice = async (status) => {
    const title = `📅 연차 신청 ${status} 안내`;
    const content = `${r.date} ${r.type} 신청이 ${status}되었습니다.${r.note ? `\n신청 사유: ${r.note}` : ""}`;
    await addDoc(collection(db, COL_NOTICES), {
      title, content, recipient: r.userId,
      author: "관리자", createdAt: new Date().toISOString(), auto: true
    });
  };

  const updateAnnualUsed = async (delta) => {
    const annualRef = doc(db, COL_ANNUAL, r.userId);
    const snap = await getDoc(annualRef);
    const current = snap.exists() ? snap.data() : { total: 0, used: 0 };
    const newUsed = Math.max(0, Number(current.used || 0) + delta);
    await setDoc(annualRef, { ...current, used: newUsed });
  };

  const handleApprove = async () => {
    setProcessing("승인");
    const leaveData = { userId: r.userId, date: r.date, type: r.type };
    if (r.hours) leaveData.hours = r.hours;
    const delta = r.type === "시간연차" ? (r.hours || 1) / 8 : r.type?.includes("반차") ? 0.5 : 1;
    await Promise.all([
      setDoc(doc(db, COL_LEAVE_REQ, r.id), { status: "승인" }, { merge: true }),
      setDoc(doc(db, COL_LEAVES, `${r.userId}_${r.date}`), leaveData),
      updateAnnualUsed(delta),
      sendNotice("승인"),
    ]);
    setProcessing(null);
    setDone("승인");
    setTimeout(() => setDone(null), 1500);
  };

  const handleReject = async () => {
    setProcessing("반려");
    const delta = r.type === "시간연차" ? -((r.hours || 1) / 8) : r.type?.includes("반차") ? -0.5 : -1;
    await Promise.all([
      setDoc(doc(db, COL_LEAVE_REQ, r.id), { status: "반려" }, { merge: true }),
      setDoc(doc(db, COL_LEAVES, `${r.userId}_${r.date}`), { userId: r.userId, date: r.date, deleted: true }),
      sendNotice("반려"),
      ...(r.status === "승인" ? [updateAnnualUsed(delta)] : []),
    ]);
    setProcessing(null);
    setDone("반려");
    setTimeout(() => setDone(null), 1500);
  };

  return (
    <div style={{ background: T.card, borderRadius: 12, padding: "12px 14px", marginBottom: 8, border: `1px solid ${done ? (done === "승인" ? T.green : T.red) : T.border}`, transition: "border 0.3s" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
        <div style={{ fontWeight: 700, fontSize: 13, color: T.text, flex: 1 }}>{r.userName} · {r.date} · {r.type}{r.type === "시간연차" && r.hours ? ` (${r.hours}시간)` : ""}</div>
        <Badge label={done || r.status} color={statusColor[done || r.status] || "gray"} />
      </div>
      {r.note && <div style={{ fontSize: 11, color: T.muted, marginBottom: 8 }}>📝 {r.note}</div>}
      {done ? (
        <div style={{ textAlign: "center", padding: "8px 0", fontSize: 13, fontWeight: 700, color: done === "승인" ? T.green : T.red }}>
          {done === "승인" ? "✓ 승인 완료!" : "✓ 반려 완료!"}
        </div>
      ) : r.status === "승인" ? (
        // 이미 승인된 건 - 승인완료 표시 + 삭제만 가능
        <div style={{ display: "flex", gap: 8 }}>
          <div style={{ flex: 1, padding: "9px 0", borderRadius: 8, background: T.greenBg, color: T.green, fontSize: 12, fontWeight: 700, textAlign: "center" }}>✓ 승인완료</div>
          <button onClick={() => setDelConfirm(r)}
            style={{ padding: "9px 12px", borderRadius: 8, border: `1px solid ${T.border}`, background: "#fff", color: T.muted, fontSize: 12, cursor: "pointer" }}>삭제</button>
        </div>
      ) : (
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={handleApprove} disabled={!!processing}
            style={{ flex: 1, padding: "9px 0", borderRadius: 8, border: "none", background: processing === "승인" ? T.green : T.greenBg, color: processing === "승인" ? "#fff" : T.green, fontSize: 12, fontWeight: 700, cursor: processing ? "not-allowed" : "pointer", transition: "all 0.15s" }}>
            {processing === "승인" ? "처리중..." : "승인"}
          </button>
          <button onClick={handleReject} disabled={!!processing}
            style={{ flex: 1, padding: "9px 0", borderRadius: 8, border: "none", background: processing === "반려" ? T.red : T.redBg, color: processing === "반려" ? "#fff" : T.red, fontSize: 12, fontWeight: 700, cursor: processing ? "not-allowed" : "pointer", transition: "all 0.15s" }}>
            {processing === "반려" ? "처리중..." : "반려"}
          </button>
          <button onClick={() => setDelConfirm(r)}
            style={{ padding: "9px 12px", borderRadius: 8, border: `1px solid ${T.border}`, background: "#fff", color: T.muted, fontSize: 12, cursor: "pointer" }}>삭제</button>
        </div>
      )}
    </div>
  );
}

function AdminMembers({ users, annual, leaveRequests, memberInfo = {}, onSaveUsers, onBack }) {
  const [showUserModal, setShowUserModal] = useState(false);
  const [showAccount, setShowAccount] = useState(false);
  const [editInfo, setEditInfo] = useState(null);
  const members = users.filter(u => u.role === "member");

  const saveInfo = async (userId, data) => {
    await setDoc(doc(db, COL_MEMBER_INFO, userId), data);
    setEditInfo(null);
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

            </div>
          );
        })}

      </div>


      {editInfo && (
        <MemberInfoModal user={editInfo.user} info={memberInfo[editInfo.user.id] || {}}
          onSave={data => saveInfo(editInfo.user.id, data)} onClose={() => setEditInfo(null)} />
      )}

      {showUserModal && <UserManageModal users={users} onSave={async u => { await fbSaveUsers(u, users); setShowUserModal(false); }} onClose={() => setShowUserModal(false)} />}
      {showAccount && <AdminAccountModal users={users} onUpdateUsers={onSaveUsers} onClose={() => setShowAccount(false)} />}
    </div>
  );
}

// ── 관리자 일반 섹션 ───────────────────────────────────────────

// ── 리마인더 ────────────────────────────────────────────────────
function AdminReminder({ reminders = [], users = [] }) {
  const EMPTY = { title: "", time: "09:00", repeat: "daily", monthDay: 1, weekDay: 1, target: "admin", sendBeforeHoliday: false };
  const [form, setForm] = useState(EMPTY);
  const [editId, setEditId] = useState(null); // null=추가, id=수정
  const [adding, setAdding] = useState(false);
  const [loading, setLoading] = useState(false);

  const DOW = ["일","월","화","수","목","금","토"];

  const openEdit = (r) => {
    setForm({ title: r.title, time: r.time, repeat: r.repeat, monthDay: r.monthDay || 1, weekDay: r.weekDay || 1, target: r.target || "admin", sendBeforeHoliday: r.sendBeforeHoliday || false });
    setEditId(r.id);
    setAdding(false); // 상단 추가 폼 닫기
  };

  const saveReminder = async () => {
    if (!form.title.trim()) return;
    setLoading(true);
    try {
      const id = editId || Date.now().toString();
      const existing = editId ? reminders.find(r => r.id === editId) : null;
      await setDoc(doc(db, COL_REMINDERS, id), {
        ...form,
        title: form.title.trim(),
        createdAt: existing?.createdAt || new Date().toISOString(),
        active: existing?.active ?? true,
      });
      setForm(EMPTY);
      setEditId(null);
      setAdding(false);
    } catch(e) { console.error(e); }
    setLoading(false);
  };

  const cancelForm = () => { setForm(EMPTY); setEditId(null); setAdding(false); };

  const toggleActive = async (r) => {
    await setDoc(doc(db, COL_REMINDERS, r.id), { ...r, active: !r.active });
  };

  const deleteReminder = async (id) => {
    await deleteDoc(doc(db, COL_REMINDERS, id));
  };

  const repeatDesc = (r) => {
    if (r.repeat === "daily") return "매일";
    if (r.repeat === "weekly") return `매주 ${DOW[r.weekDay]}요일`;
    if (r.repeat === "monthly") return `매월 ${r.monthDay}일`;
    return "";
  };

  return (
    <div style={{ padding: 16 }}>
      {/* 추가 버튼 */}
      {!adding && (
        <button onClick={() => setAdding(true)}
          style={{ width: "100%", background: "#7c3aed", border: "none", color: "#fff", borderRadius: 12, padding: "12px 0", fontSize: 14, fontWeight: 700, cursor: "pointer", marginBottom: 16 }}>
          + 리마인더 추가
        </button>
      )}

      {/* 추가 폼 */}
      {adding && (
        <div style={{ background: T.card, border: `1px solid #7c3aed44`, borderRadius: 14, padding: 16, marginBottom: 16 }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: "#7c3aed", marginBottom: 12 }}>{editId ? "✏ 리마인더 수정" : "새 리마인더"}</div>

          <input value={form.title} onChange={e => setForm(p => ({...p, title: e.target.value}))}
            placeholder="제목 (예: 급여일, 4대보험 납부)"
            style={{ width: "100%", padding: "10px 12px", borderRadius: 10, border: `1px solid ${T.border}`, fontSize: 13, fontWeight: 600, marginBottom: 10, boxSizing: "border-box", background: "#fff", color: T.text }} />

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 10 }}>
            <div>
              <div style={{ fontSize: 11, color: T.muted, marginBottom: 4 }}>알림 시간</div>
              <input type="time" value={form.time} onChange={e => setForm(p => ({...p, time: e.target.value}))}
                style={{ width: "100%", padding: "10px 10px", borderRadius: 10, border: `1px solid ${T.border}`, fontSize: 13, fontWeight: 700, boxSizing: "border-box", background: "#fff", color: T.text }} />
            </div>
            <div>
              <div style={{ fontSize: 11, color: T.muted, marginBottom: 4 }}>반복</div>
              <select value={form.repeat} onChange={e => setForm(p => ({...p, repeat: e.target.value}))}
                style={{ width: "100%", padding: "10px 8px", borderRadius: 10, border: `1px solid ${T.border}`, fontSize: 13, fontWeight: 700, boxSizing: "border-box", background: "#fff", color: T.text }}>
                <option value="daily">매일</option>
                <option value="weekly">매주</option>
                <option value="monthly">매월</option>
              </select>
            </div>
          </div>

          {form.repeat === "weekly" && (
            <div style={{ marginBottom: 10 }}>
              <div style={{ fontSize: 11, color: T.muted, marginBottom: 6 }}>요일 선택</div>
              <div style={{ display: "flex", gap: 6 }}>
                {DOW.map((d, i) => (
                  <button key={i} onClick={() => setForm(p => ({...p, weekDay: i}))}
                    style={{ flex: 1, padding: "8px 0", borderRadius: 8, border: `1px solid ${form.weekDay === i ? "#7c3aed" : T.border}`,
                      background: form.weekDay === i ? "#7c3aed" : "#fff", color: form.weekDay === i ? "#fff" : T.text,
                      fontSize: 12, fontWeight: 700, cursor: "pointer" }}>{d}</button>
                ))}
              </div>
            </div>
          )}

          {form.repeat === "monthly" && (
            <div style={{ marginBottom: 10 }}>
              <div style={{ fontSize: 11, color: T.muted, marginBottom: 4 }}>날짜</div>
              <select value={form.monthDay} onChange={e => setForm(p => ({...p, monthDay: Number(e.target.value)}))}
                style={{ width: "100%", padding: "10px 8px", borderRadius: 10, border: `1px solid ${T.border}`, fontSize: 13, fontWeight: 700, background: "#fff", color: T.text }}>
                {Array.from({length: 31}, (_, i) => i+1).map(d => <option key={d} value={d}>{d}일</option>)}
              </select>
            </div>
          )}

          <div style={{ marginBottom: 12 }}>
            <div style={{ fontSize: 11, color: T.muted, marginBottom: 4 }}>수신 대상</div>
            <select value={form.target} onChange={e => setForm(p => ({...p, target: e.target.value}))}
              style={{ width: "100%", padding: "10px 8px", borderRadius: 10, border: `1px solid ${T.border}`, fontSize: 13, fontWeight: 700, background: "#fff", color: T.text }}>
              <option value="admin">관리자만</option>
              <option value="all">전체</option>
            </select>
          </div>

          <div onClick={() => setForm(p => ({...p, sendBeforeHoliday: !p.sendBeforeHoliday}))}
            style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12, cursor: "pointer", padding: "8px 10px", borderRadius: 10, background: form.sendBeforeHoliday ? "#ede9fe" : T.bg, border: `1px solid ${form.sendBeforeHoliday ? "#7c3aed" : T.border}` }}>
            <div style={{ width: 18, height: 18, borderRadius: 4, border: `2px solid ${form.sendBeforeHoliday ? "#7c3aed" : T.border}`, background: form.sendBeforeHoliday ? "#7c3aed" : "#fff", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
              {form.sendBeforeHoliday && <span style={{ color: "#fff", fontSize: 12, fontWeight: 900 }}>✓</span>}
            </div>
            <div>
              <div style={{ fontSize: 12, fontWeight: 700, color: form.sendBeforeHoliday ? "#7c3aed" : T.text }}>공휴일 전날 대신 발송</div>
              <div style={{ fontSize: 10, color: T.muted }}>해당일이 공휴일이면 직전 평일에 발송</div>
            </div>
          </div>

          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={cancelForm}
              style={{ flex: 1, background: T.bg, border: `1px solid ${T.border}`, color: T.muted, borderRadius: 10, padding: "10px 0", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>취소</button>
            <button onClick={saveReminder} disabled={loading}
              style={{ flex: 2, background: "#7c3aed", border: "none", color: "#fff", borderRadius: 10, padding: "10px 0", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
              {loading ? "저장 중..." : "저장"}
            </button>
          </div>
        </div>
      )}

      {/* 리마인더 목록 */}
      {reminders.length === 0 ? (
        <div style={{ textAlign: "center", color: T.muted, fontSize: 13, padding: "32px 0" }}>등록된 리마인더 없음</div>
      ) : reminders.map(r => (
        <div key={r.id} style={{ background: T.card, border: `1px solid ${editId === r.id ? "#7c3aed" : r.active ? "#7c3aed44" : T.border}`, borderRadius: 14, marginBottom: 10, overflow: "hidden", opacity: r.active ? 1 : 0.6 }}>
          {/* 요약 행 */}
          <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "12px 14px" }}>
            <div style={{ flex: 1 }}>
              <div style={{ fontSize: 14, fontWeight: 800, color: r.active ? T.text : T.muted }}>{r.title}</div>
              <div style={{ fontSize: 11, color: "#7c3aed", fontWeight: 600, marginTop: 2 }}>
                {repeatDesc(r)} · {r.time} · {r.target === "all" ? "전체" : "관리자"}{r.sendBeforeHoliday ? " · 전날발송" : ""}
              </div>
            </div>
            <div style={{ display: "flex", gap: 5 }}>
              <button onClick={() => toggleActive(r)}
                style={{ background: r.active ? "#dcfce7" : T.bg, border: `1px solid ${r.active ? "#16a34a" : T.border}`, color: r.active ? "#16a34a" : T.muted, borderRadius: 8, padding: "4px 8px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>
                {r.active ? "ON" : "OFF"}
              </button>
              <button onClick={() => editId === r.id ? cancelForm() : openEdit(r)}
                style={{ background: editId === r.id ? "#ede9fe" : T.bg, border: `1px solid ${editId === r.id ? "#7c3aed" : T.border}`, color: editId === r.id ? "#7c3aed" : T.muted, borderRadius: 8, padding: "4px 8px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>
                {editId === r.id ? "닫기" : "수정"}
              </button>
              <button onClick={() => deleteReminder(r.id)}
                style={{ background: T.redBg, border: "none", color: T.red, borderRadius: 8, padding: "4px 8px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>삭제</button>
            </div>
          </div>
          {/* 인라인 수정 폼 */}
          {editId === r.id && (
            <div style={{ borderTop: `1px solid #7c3aed33`, padding: "12px 14px", background: "#faf5ff" }}>
              <input value={form.title} onChange={e => setForm(p => ({...p, title: e.target.value}))}
                placeholder="제목"
                style={{ width: "100%", padding: "9px 10px", borderRadius: 10, border: `1px solid ${T.border}`, fontSize: 13, fontWeight: 600, marginBottom: 8, boxSizing: "border-box", background: "#fff", color: T.text }} />
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
                <input type="time" value={form.time} onChange={e => setForm(p => ({...p, time: e.target.value}))}
                  style={{ padding: "9px 10px", borderRadius: 10, border: `1px solid ${T.border}`, fontSize: 13, fontWeight: 700, background: "#fff", color: T.text, width: "100%", boxSizing: "border-box" }} />
                <select value={form.repeat} onChange={e => setForm(p => ({...p, repeat: e.target.value}))}
                  style={{ padding: "9px 8px", borderRadius: 10, border: `1px solid ${T.border}`, fontSize: 13, fontWeight: 700, background: "#fff", color: T.text, width: "100%", boxSizing: "border-box" }}>
                  <option value="daily">매일</option>
                  <option value="weekly">매주</option>
                  <option value="monthly">매월</option>
                </select>
              </div>
              {form.repeat === "weekly" && (
                <div style={{ display: "flex", gap: 4, marginBottom: 8 }}>
                  {["일","월","화","수","목","금","토"].map((d, i) => (
                    <button key={i} onClick={() => setForm(p => ({...p, weekDay: i}))}
                      style={{ flex: 1, padding: "6px 0", borderRadius: 8, border: `1px solid ${form.weekDay === i ? "#7c3aed" : T.border}`,
                        background: form.weekDay === i ? "#7c3aed" : "#fff", color: form.weekDay === i ? "#fff" : T.text,
                        fontSize: 11, fontWeight: 700, cursor: "pointer" }}>{d}</button>
                  ))}
                </div>
              )}
              {form.repeat === "monthly" && (
                <select value={form.monthDay} onChange={e => setForm(p => ({...p, monthDay: Number(e.target.value)}))}
                  style={{ width: "100%", padding: "9px 8px", borderRadius: 10, border: `1px solid ${T.border}`, fontSize: 13, fontWeight: 700, background: "#fff", color: T.text, marginBottom: 8, boxSizing: "border-box" }}>
                  {Array.from({length: 31}, (_, i) => i+1).map(d => <option key={d} value={d}>{d}일</option>)}
                </select>
              )}
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
                <select value={form.target} onChange={e => setForm(p => ({...p, target: e.target.value}))}
                  style={{ padding: "9px 8px", borderRadius: 10, border: `1px solid ${T.border}`, fontSize: 13, fontWeight: 700, background: "#fff", color: T.text, boxSizing: "border-box" }}>
                  <option value="admin">관리자만</option>
                  <option value="all">전체</option>
                </select>
                <button onClick={saveReminder} disabled={loading}
                  style={{ background: "#7c3aed", border: "none", color: "#fff", borderRadius: 10, fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                  {loading ? "저장 중..." : "저장"}
                </button>
              </div>
              <div onClick={() => setForm(p => ({...p, sendBeforeHoliday: !p.sendBeforeHoliday}))}
                style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4, cursor: "pointer", padding: "7px 10px", borderRadius: 10, background: form.sendBeforeHoliday ? "#ede9fe" : "#fff", border: `1px solid ${form.sendBeforeHoliday ? "#7c3aed" : T.border}` }}>
                <div style={{ width: 16, height: 16, borderRadius: 4, border: `2px solid ${form.sendBeforeHoliday ? "#7c3aed" : T.border}`, background: form.sendBeforeHoliday ? "#7c3aed" : "#fff", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                  {form.sendBeforeHoliday && <span style={{ color: "#fff", fontSize: 10, fontWeight: 900 }}>✓</span>}
                </div>
                <div style={{ fontSize: 11, fontWeight: 700, color: form.sendBeforeHoliday ? "#7c3aed" : T.muted }}>공휴일 전날 대신 발송</div>
              </div>
            </div>
          )}
        </div>
      ))}
    </div>
  );
}

// ── 관리자 섹션 래퍼 (공지/게시판용) ─────────────────────────────
function AdminSectionWrap({ title, color, onBack, children }) {
  return (
    <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'Noto Sans KR',sans-serif" }}>
      <div style={{ background: color || T.adminHeader, padding: "16px 16px 14px" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <button onClick={onBack} style={{ background: "#ffffff18", border: "none", color: "#fff", fontSize: 18, cursor: "pointer", padding: "8px 14px", borderRadius: 12, fontWeight: 700 }}>‹</button>
          <div>
            <div style={{ fontSize: 11, color: "#ffffff40", letterSpacing: 3 }}>ADMIN</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>{title}</div>
          </div>
        </div>
      </div>
      {children}
      <FloatBack onClick={onBack} />
    </div>
  );
}

// ── 관리자 화면 (라우터) ───────────────────────────────────────
function AdminScreen({ user, users, settings, records, leaves, notices, board, payslips, annual, leaveRequests, memberInfo, reads, reminders = [], onSaveRecord, onSaveLeave, onSaveUsers, onSaveSettings, onLogout }) {
  const [section, setSection] = useState(null);


  const back = () => { setSection(null); window.scrollTo(0,0); };
  if (!section) return <AdminHome user={user} onLogout={onLogout} onSection={s => { setSection(s); window.scrollTo(0,0); }} leaveRequests={leaveRequests} board={board} reads={reads} />;
  if (section === "attendance") return <><AdminAttendance users={users} settings={settings} records={records} leaves={leaves} leaveRequests={leaveRequests} onSaveRecord={onSaveRecord} onSaveLeave={onSaveLeave} onSaveSettings={onSaveSettings} onBack={back} /><FloatBack onClick={back} /></>;
  if (section === "wage") return <><AdminWage users={users} records={records} leaves={leaves} settings={settings} memberInfo={memberInfo} annual={annual} leaveRequests={leaveRequests} payslips={payslips} reads={reads} onBack={back} /><FloatBack onClick={back} /></>;
  if (section === "members") return <><AdminMembers users={users} annual={annual} leaveRequests={leaveRequests} memberInfo={memberInfo} onSaveUsers={onSaveUsers} onBack={back} /><FloatBack onClick={back} /></>;
  if (section === "annual") return <><AnnualScreen user={user} users={users} annual={annual} leaveRequests={leaveRequests} onBack={back} /><FloatBack onClick={back} /></>;
  if (section === "severance") return <><AdminSeverance users={users} memberInfo={memberInfo} annual={annual} onBack={back} /><FloatBack onClick={back} /></>;
  if (section === "notice") return <AdminSectionWrap title="📢 공지사항" color="#ea580c" onBack={back}><NoticeScreen user={user} users={users} notices={notices} reads={reads} /></AdminSectionWrap>;
  if (section === "board") return <AdminSectionWrap title="💬 게시판" color="#0891b2" onBack={back}><BoardScreen user={user} board={board} reads={reads} /></AdminSectionWrap>;
  if (section === "settings") return <><SettingsModal settings={settings} onSave={async s => { await onSaveSettings(s); back(); }} onClose={back} /></>;
  if (section === "reminder") return <AdminSectionWrap title="🔔 리마인더" color="#7c3aed" onBack={back}><AdminReminder reminders={reminders} users={users} /></AdminSectionWrap>;
  return null;
}

// ── 공지사항 ────────────────────────────────────────────────────
function NoticeScreen({ user, users, notices, reads }) {
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

  const markRead = async (id) => {
    const key = `${user.id}_notice_${id}`;
    if (!reads[key]) await setDoc(doc(db, COL_READS, key), { userId: user.id, type: "notice", docId: id, readAt: new Date().toISOString() });
  };

  const toggleExpanded = (id) => {
    setExpanded(expanded === id ? null : id);
    if (expanded !== id) markRead(id);
  };

  const isUnread = (n) => !isAdmin && !reads[`${user.id}_notice_${n.id}`];

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
    const pushTarget = recipient === "all" ? null : recipient;
    await sendPush({ title: `📢 공지: ${title.trim()}`, message: content.trim(), targetUserId: pushTarget });
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
    setRecipient(n.recipient || "all"); setFile(null);
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
          <div key={n.id} style={{ background: T.card, borderRadius: 14, padding: "14px 16px", marginBottom: 10, border: `1px solid ${isUnread(n) ? T.blue : T.border}` }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }} onClick={() => toggleExpanded(n.id)}>
              <div style={{ flex: 1 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 3 }}>
                  {isUnread(n) && <div style={{ width: 8, height: 8, borderRadius: "50%", background: T.blue, flexShrink: 0 }} />}
                  <div style={{ fontWeight: isUnread(n) ? 800 : 700, fontSize: 14, color: T.text }}>{n.title}</div>
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
                  <div>
                    {editTarget?.id === n.id ? (
                      <div style={{ marginTop: 10 }}>
                        <input value={title} onChange={e => setTitle(e.target.value)}
                          style={{ width: "100%", padding: "10px 14px", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 14, boxSizing: "border-box", marginBottom: 8 }} placeholder="제목" />
                        <textarea value={content} onChange={e => setContent(e.target.value)}
                          rows={4} style={{ width: "100%", padding: "10px 14px", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 14, boxSizing: "border-box", resize: "none", lineHeight: 1.6, fontFamily: "inherit", marginBottom: 8 }} placeholder="내용" />
                        <select value={recipient} onChange={e => setRecipient(e.target.value)}
                          style={{ width: "100%", padding: "10px 14px", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 14, boxSizing: "border-box", marginBottom: 8 }}>
                          <option value="all">전체</option>
                          {members.map(m => <option key={m.id} value={m.id}>{m.name}</option>)}
                        </select>
                        <div style={{ display: "flex", gap: 8 }}>
                          <Btn variant="ghost" onClick={() => setEditTarget(null)}>취소</Btn>
                          <Btn variant="admin" onClick={update}>수정 완료</Btn>
                        </div>
                      </div>
                    ) : (
                      <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
                        <button onClick={() => openEdit(n)} style={{ background: T.bg, border: `1px solid ${T.border}`, color: T.text, borderRadius: 8, padding: "5px 12px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>수정</button>
                        <button onClick={() => del(n)} style={{ background: T.redBg, border: "none", color: T.red, borderRadius: 8, padding: "5px 12px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>삭제</button>
                      </div>
                    )}
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
function BoardScreen({ user, board, reads }) {
  const isAdmin = user.role === "admin";
  const [showWrite, setShowWrite] = useState(false);
  const [title, setTitle] = useState(""), [content, setContent] = useState("");
  const [expanded, setExpanded] = useState(null);
  const [editTarget, setEditTarget] = useState(null);
  const [editTitle, setEditTitle] = useState(""), [editContent, setEditContent] = useState("");

  const submit = async () => {
    if (!title.trim() || !content.trim()) return;
    await addDoc(collection(db, COL_BOARD), {
      title: title.trim(), content: content.trim(),
      author: user.name, userId: user.id, createdAt: new Date().toISOString()
    });
    await sendPush({ title: `💬 게시판: ${title.trim()}`, message: `${user.name}: ${content.trim()}` });
    setTitle(""); setContent(""); setShowWrite(false);
  };

  const del = async (id) => { await deleteDoc(doc(db, COL_BOARD, id)); };

  const startEdit = (b) => { setEditTarget(b.id); setEditTitle(b.title); setEditContent(b.content); };
  const cancelEdit = () => { setEditTarget(null); setEditTitle(""); setEditContent(""); };
  const saveEdit = async (b) => {
    if (!editTitle.trim() || !editContent.trim()) return;
    await setDoc(doc(db, COL_BOARD, b.id), { ...b, title: editTitle.trim(), content: editContent.trim() });
    cancelEdit();
  };

  const markRead = async (id) => {
    const key = `${user.id}_board_${id}`;
    if (!reads[key]) await setDoc(doc(db, COL_READS, key), { userId: user.id, type: "board", docId: id, readAt: new Date().toISOString() });
  };

  const toggleExpanded = (id) => {
    setExpanded(expanded === id ? null : id);
    if (expanded !== id) markRead(id);
  };

  const isUnread = (b) => !reads[`${user.id}_board_${b.id}`] && b.userId !== user.id;

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
          <div key={b.id} style={{ background: T.card, borderRadius: 14, padding: "14px 16px", marginBottom: 10, border: `1px solid ${isUnread(b) ? T.blue : T.border}` }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }} onClick={() => toggleExpanded(b.id)}>
              <div style={{ flex: 1 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 3 }}>
                  {isUnread(b) && <div style={{ width: 8, height: 8, borderRadius: "50%", background: T.blue, flexShrink: 0 }} />}
                  <div style={{ fontWeight: isUnread(b) ? 800 : 700, fontSize: 14, color: T.text }}>{b.title}</div>
                </div>
                <div style={{ fontSize: 11, color: T.muted }}>{b.author} · {b.createdAt?.slice(0,10)}</div>
              </div>
              <span style={{ color: T.muted, fontSize: 14 }}>{expanded === b.id ? "▲" : "▼"}</span>
            </div>
            {expanded === b.id && (
              <div style={{ marginTop: 12, paddingTop: 12, borderTop: `1px solid ${T.border}` }}>
                <div style={{ fontSize: 14, color: T.text, lineHeight: 1.7, whiteSpace: "pre-wrap" }}>{b.content}</div>
                {(isAdmin || b.userId === user.id) && (
                  editTarget === b.id ? (
                    <div style={{ marginTop: 10 }}>
                      <input value={editTitle} onChange={e => setEditTitle(e.target.value)}
                        style={{ width: "100%", padding: "10px 14px", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 14, boxSizing: "border-box", marginBottom: 8 }} />
                      <textarea value={editContent} onChange={e => setEditContent(e.target.value)} rows={4}
                        style={{ width: "100%", padding: "10px 14px", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 14, boxSizing: "border-box", resize: "none", fontFamily: "inherit", marginBottom: 8 }} />
                      <div style={{ display: "flex", gap: 8 }}>
                        <Btn variant="ghost" onClick={cancelEdit}>취소</Btn>
                        <Btn variant="primary" onClick={() => saveEdit(b)}>수정 완료</Btn>
                      </div>
                    </div>
                  ) : (
                    <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
                      {b.userId === user.id && <button onClick={() => startEdit(b)} style={{ background: T.bg, border: `1px solid ${T.border}`, color: T.text, borderRadius: 8, padding: "5px 12px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>수정</button>}
                      <button onClick={() => del(b.id)} style={{ background: T.redBg, border: "none", color: T.red, borderRadius: 8, padding: "5px 12px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>삭제</button>
                    </div>
                  )
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
function PayslipScreen({ user, users, payslips, reads }) {
  const isAdmin = user.role === "admin";
  const [uploading, setUploading] = useState(false);
  const [selUser, setSelUser] = useState("");
  const [selMonth, setSelMonth] = useState(new Date(new Date().getTime() + 9*60*60*1000).toISOString().slice(0,7));
  const [file, setFile] = useState(null);
  const [msg, setMsg] = useState("");
  const [expanded, setExpanded] = useState({});

  const members = users.filter(u => u.role === "member");
  const myPayslips = isAdmin ? payslips : payslips.filter(p => p.userId === user.id);

  const [pdfLoading, setPdfLoading] = useState(null);

  const downloadPDF = async (p) => {
    const w = p.wageData;
    if (!w) return;
    const member = users.find(u => u.id === p.userId);
    const name = member?.name || "팀원";
    setPdfLoading(p.id);
    try {
      const el = document.getElementById(`payslip-content-${p.id}`);
      if (!el) { alert("명세서 영역 없음"); setPdfLoading(null); return; }
      const canvas = await html2canvas(el, { scale: 2, useCORS: true, backgroundColor: "#ffffff" });
      const imgData = canvas.toDataURL("image/png");
      const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
      const pageW = pdf.internal.pageSize.getWidth();
      const pageH = pdf.internal.pageSize.getHeight();
      const margin = 10;
      const maxW = pageW - margin * 2;
      const maxH = pageH - margin * 2;
      const ratio = canvas.width / canvas.height;
      let imgW = maxW;
      let imgH = imgW / ratio;
      if (imgH > maxH) { imgH = maxH; imgW = imgH * ratio; }
      const x = (pageW - imgW) / 2;
      const y = (pageH - imgH) / 2;
      pdf.addImage(imgData, "PNG", x, y, imgW, imgH);
      pdf.save(`${name}_${monthLabel(p.month)}_급여명세서.pdf`);
    } catch(e) {
      alert("PDF 생성 실패: " + e.message);
    }
    setPdfLoading(null);
  };

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
          const unread = !isAdmin && !reads?.[`${user.id}_payslip_${p.id}`];
          const handleRead = async () => {
            const key = `${user.id}_payslip_${p.id}`;
            if (!reads?.[key]) await setDoc(doc(db, COL_READS, key), { userId: user.id, type: "payslip", docId: p.id, readAt: new Date().toISOString() });
          };
          const w = p.wageData;
          const isOpen = !!expanded[p.id];
          return (
            <div key={p.id} style={{ background: T.card, borderRadius: 14, marginBottom: 10, border: `1px solid ${unread ? T.blue : T.border}`, overflow: "hidden" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "14px 16px", cursor: "pointer" }}
                onClick={() => { setExpanded(e => ({...e, [p.id]: !e[p.id]})); handleRead(); }}>
                <div style={{ flex: 1 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    {unread && <div style={{ width: 8, height: 8, borderRadius: "50%", background: T.blue, flexShrink: 0 }} />}
                    <div style={{ fontWeight: unread ? 800 : 700, fontSize: 14, color: T.text }}>
                      {isAdmin && `${member?.name} · `}{monthLabel(p.month)}
                      {p.isAuto && <span style={{ fontSize: 10, background: "#16a34a22", color: "#16a34a", borderRadius: 6, padding: "2px 6px", marginLeft: 6, fontWeight: 700 }}>자동생성</span>}
                    </div>
                  </div>
                  <div style={{ fontSize: 11, color: T.muted, marginTop: 2 }}>
                    {p.createdAt?.slice(0,10)} 발급 · 지급일 {w?.payDate || "-"}
                    {w && <span style={{ marginLeft: 8, fontWeight: 700, color: "#16a34a" }}>실지급 {Number(w.netPay||0).toLocaleString()}원</span>}
                  </div>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                  {p.url && <a href={p.url} target="_blank" rel="noopener noreferrer" onClick={e => e.stopPropagation()}
                    style={{ background: T.blueBg, color: T.blue, borderRadius: 8, padding: "5px 10px", fontSize: 11, fontWeight: 700, textDecoration: "none" }}>보기</a>}

                  {isAdmin && <button onClick={e => { e.stopPropagation(); deleteDoc(doc(db, COL_PAYSLIPS, p.id)); }}
                    style={{ background: T.redBg, border: "none", color: T.red, borderRadius: 8, padding: "5px 8px", fontSize: 11, cursor: "pointer", fontWeight: 700 }}>삭제</button>}
                  <span style={{ color: T.muted, fontSize: 14 }}>{isOpen ? "▲" : "▼"}</span>
                </div>
              </div>
              {isOpen && w && (
                <div id={`payslip-content-${p.id}`} style={{ background: T.bg, padding: "12px 16px", borderTop: `1px solid ${T.border}` }}>
                  {/* 근태 내역 */}
                  <div style={{ fontSize: 11, color: T.muted, fontWeight: 700, marginBottom: 6 }}>근태 내역</div>
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 4, marginBottom: 10 }}>
                    {[
                      ["출근일수", (w.monthStats?.days||0)+"일"],
                      ["연장시간", fmtMinutes(w.monthStats?.otMin||0)],
                      ["휴일근무", (w.monthStats?.holiday||0)+"일"],
                      ["지각", fmtMinutes(w.monthStats?.lateMin||0)],
                      ["조퇴", fmtMinutes(w.monthStats?.earlyMin||0)],
                      ["외출", fmtMinutes(w.monthStats?.outingMin||0)],
                    ].map(([l,v]) => (
                      <div key={l} style={{ textAlign: "center", background: T.card, borderRadius: 6, padding: "4px 0" }}>
                        <div style={{ fontSize: 9, color: T.muted }}>{l}</div>
                        <div style={{ fontSize: 11, fontWeight: 700, color: T.text }}>{v}</div>
                      </div>
                    ))}
                  </div>
                  {/* 소득 내역 */}
                  <div style={{ fontSize: 11, color: T.blue, fontWeight: 700, marginBottom: 6 }}>소득 내역</div>
                  {[
                    ["기본급", w.monthlyBase, `시급 ${Number(w.hourlyWage||0).toLocaleString()} × ${w.monthlyHours||209}`],
                    ["연장수당", w.otPay, `시급 × 연장${fmtMinutes(w.monthStats?.otMin||0)} × 1.5`],
                    ["휴일수당", w.holidayPay, `일급 × 휴일${w.monthStats?.holiday||0}일 × 1.5`],
                    ["상여금", w.bonus, ""],
                    ["이월분", w.carryOver, ""],
                    ["기타", w.otherIncome, ""],
                  ].filter(([,v]) => Number(v) > 0).map(([l,v,c]) => (
                    <div key={l} style={{ display: "flex", justifyContent: "space-between", padding: "3px 0", borderBottom: `1px solid ${T.border}` }}>
                      <div>
                        <span style={{ fontSize: 12, color: T.text }}>{l}</span>
                        {c && <span style={{ fontSize: 10, color: T.muted, marginLeft: 6 }}>{c}</span>}
                      </div>
                      <span style={{ fontSize: 12, fontWeight: 700, color: T.blue }}>{Number(v).toLocaleString()}</span>
                    </div>
                  ))}
                  <div style={{ display: "flex", justifyContent: "space-between", padding: "4px 0", marginBottom: 8 }}>
                    <span style={{ fontSize: 12, fontWeight: 800 }}>소득 합계</span>
                    <span style={{ fontSize: 13, fontWeight: 800, color: T.blue }}>{Number(w.totalIncome||0).toLocaleString()}</span>
                  </div>
                  {/* 공제 내역 */}
                  <div style={{ fontSize: 11, color: T.red, fontWeight: 700, marginBottom: 6 }}>공제 내역</div>
                  {[
                    ["소득세", w.incomeTax, "간이세액표 기준"],
                    ["주민세", w.residentTax, `소득세 × 10%`],
                    ["국민연금", w.nationalPension, `기준소득 ${Number(w.pensionBase||0).toLocaleString()} × ${w.ratePension||4.75}%`],
                    ["건강보험", w.health, `보수월액 ${Number(w.insuranceBase||0).toLocaleString()} × ${w.rateHealth||3.595}%`],
                    ["고용보험", w.employment, `보수월액 ${Number(w.insuranceBase||0).toLocaleString()} × ${w.rateEmployment||0.9}%`],
                    ["장기요양", w.longCare, `건강보험 × ${w.rateLongCare||13.14}%`],
                    ["지각/조퇴차감", w.deductPay, `${fmtMinutes(w.deductMin||0)} × 시급`],
                    ["결근공제", w.absentPay, `결근 ${w.absentDays||0}일 + ${w.absentWeeks||0}주 주휴`],
                    ["기타공제", w.otherDeduct, ""],
                  ].filter(([,v]) => Number(v) > 0).map(([l,v,c]) => (
                    <div key={l} style={{ display: "flex", justifyContent: "space-between", padding: "3px 0", borderBottom: `1px solid ${T.border}` }}>
                      <div>
                        <span style={{ fontSize: 12, color: T.text }}>{l}</span>
                        {c && <span style={{ fontSize: 10, color: T.muted, marginLeft: 6 }}>{c}</span>}
                      </div>
                      <span style={{ fontSize: 12, fontWeight: 700, color: T.red }}>{Number(v).toLocaleString()}</span>
                    </div>
                  ))}
                  <div style={{ display: "flex", justifyContent: "space-between", padding: "4px 0", marginBottom: 8 }}>
                    <span style={{ fontSize: 12, fontWeight: 800 }}>공제 합계</span>
                    <span style={{ fontSize: 13, fontWeight: 800, color: T.red }}>{Number(w.totalDeduct||0).toLocaleString()}</span>
                  </div>
                  {/* 실지급액 */}
                  <div style={{ display: "flex", justifyContent: "space-between", borderTop: `2px solid #16a34a44`, paddingTop: 6 }}>
                    <span style={{ fontSize: 13, fontWeight: 800 }}>실 지급액</span>
                    <span style={{ fontSize: 16, fontWeight: 800, color: "#16a34a" }}>{Number(w.netPay||0).toLocaleString()}원</span>
                  </div>
                  {w.memo && <div style={{ fontSize: 11, color: T.muted, marginTop: 6 }}>📝 {w.memo}</div>}
                  <button onClick={() => downloadPDF(p)} disabled={pdfLoading === p.id}
                    style={{ width: "100%", marginTop: 12, background: "#16a34a", border: "none", color: "#fff", borderRadius: 10, padding: "11px 0", fontSize: 13, fontWeight: 700, cursor: "pointer", opacity: pdfLoading === p.id ? 0.6 : 1 }}>{pdfLoading === p.id ? "생성 중..." : "⬇ PDF 다운로드"}</button>
                </div>
              )}
            </div>
          );
        })
      }
    </div>
  );
}

// ── 연차 현황 ────────────────────────────────────────────────────
function AnnualScreen({ user, users, annual, leaveRequests, onBack }) {
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
  const [reqHours, setReqHours] = useState(1);

  const myAnnual = annual[user.id] || { total: 0, used: 0 };
  const myRemain = (myAnnual.total || 0) - (myAnnual.used || 0);
  const myRequests = leaveRequests.filter(r => r.userId === user.id);

  const thisYear = new Date().getFullYear().toString();
  const [selectedYear, setSelectedYear] = useState(thisYear);
  const years = [...new Set(leaveRequests.map(r => r.date?.slice(0,4)).filter(Boolean))].sort((a,b) => b-a);
  if (!years.includes(thisYear)) years.unshift(thisYear);



  const saveAnnual = async (uid) => {
    await setDoc(doc(db, COL_ANNUAL, uid), { total, used });
    setEditUser(null);
  };

  const submitRequest = async () => {
    if (!reqDate) { setReqMsg("날짜를 선택해주세요"); return; }
    await addDoc(collection(db, COL_LEAVE_REQ), {
      userId: user.id, userName: user.name,
      date: reqDate, type: reqType, note: reqNote,
      ...(reqType === "시간연차" ? { hours: reqHours } : {}),
      status: "대기", createdAt: new Date().toISOString()
    });
    await sendPush({ title: "📅 연차 신청", message: `${user.name}님이 ${reqDate} ${reqType}을 신청했습니다.`, targetUserId: "admin" });
    setReqMsg("신청 완료! ✓"); setReqDate(""); setReqNote("");
    setTimeout(() => { setReqMsg(""); setShowReqForm(false); }, 2000);
  };

  const [delConfirm, setDelConfirm] = useState(null);
  const delReq = async (r) => {
    const tasks = [deleteDoc(doc(db, COL_LEAVE_REQ, r.id))];
    if (r.status === "승인") {
      tasks.push(setDoc(doc(db, COL_LEAVES, `${r.userId}_${r.date}`), { userId: r.userId, date: r.date, deleted: true }));
      const annualRef = doc(db, COL_ANNUAL, r.userId);
      const snap = await getDoc(annualRef);
      const current = snap.exists() ? snap.data() : { total: 0, used: 0 };
      const delta = r.type === "시간연차" ? -((r.hours || 1) / 8) : r.type?.includes("반차") ? -0.5 : -1;
      const newUsed = Math.max(0, Number(current.used || 0) + delta);
      tasks.push(setDoc(annualRef, { ...current, used: newUsed }));
    }
    await Promise.all(tasks);
    setDelConfirm(null);
  };

  const statusColor = { "대기": "yellow", "승인": "green", "반려": "red" };

  return (
    <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'Noto Sans KR',sans-serif" }}>
      {isAdmin && onBack && (
        <div style={{ background: "#0284c7", padding: "16px 16px 20px" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <button onClick={onBack} style={{ background: "#ffffff18", border: "none", color: "#fff", fontSize: 18, cursor: "pointer", padding: "8px 14px", borderRadius: 12, fontWeight: 700 }}>‹</button>
            <div>
              <div style={{ fontSize: 11, color: "#ffffff40", letterSpacing: 3 }}>ADMIN</div>
              <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>📅 연차</div>
            </div>
          </div>
        </div>
      )}
      <div style={{ padding: 16 }}>
      {!isAdmin && <div style={{ fontSize: 16, fontWeight: 800, color: T.text, marginBottom: 16 }}>📅 연차</div>}

      {!isAdmin && (
        <>
          {/* 내 연차 현황 */}
          <div style={{ background: T.card, borderRadius: 16, padding: "16px 20px", marginBottom: 16, border: `1px solid ${T.border}` }}>
            <div style={{ fontSize: 13, color: T.muted, marginBottom: 12, fontWeight: 600 }}>내 연차 현황</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 8 }}>
              {[["총 연차", (myAnnual.total||0)+"일", T.blue], ["사용", (myAnnual.used||0)+"일", T.orange], ["잔여", myRemain+"일", T.green]].map(([l,v,c])=>(
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
                <div style={{ fontSize: 12, color: T.muted, fontWeight: 600, marginBottom: 5 }}>📅 신청 날짜</div>
                <input type="date" value={reqDate} onChange={e => setReqDate(e.target.value)}
                  style={{ width: "100%", padding: "11px 14px", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: reqDate ? T.text : T.muted, fontSize: 14, fontWeight: 600, boxSizing: "border-box", marginBottom: 10 }} />
                <div style={{ fontSize: 12, color: T.muted, fontWeight: 600, marginBottom: 5 }}>🗂 연차 종류</div>
                <select value={reqType} onChange={e => setReqType(e.target.value)}
                  style={{ width: "100%", padding: "11px 14px", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 14, fontWeight: 600, boxSizing: "border-box", marginBottom: 10 }}>
                  {LEAVE_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
                </select>
                {reqType === "시간연차" && (
                  <>
                    <div style={{ fontSize: 12, color: T.muted, fontWeight: 600, marginBottom: 5 }}>⏱ 시간 선택</div>
                    <select value={reqHours} onChange={e => setReqHours(Number(e.target.value))}
                      style={{ width: "100%", padding: "11px 14px", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 14, fontWeight: 600, boxSizing: "border-box", marginBottom: 10 }}>
                      {[1,2,3].map(h => <option key={h} value={h}>{h}시간</option>)}
                    </select>
                  </>
                )}
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
          <div style={{ fontSize: 13, color: T.muted, marginBottom: 8, fontWeight: 600 }}>신청 내역</div>
          <div style={{ display: "flex", gap: 6, marginBottom: 12, flexWrap: "wrap" }}>
            {years.map(y => (
              <button key={y} onClick={() => setSelectedYear(y)}
                style={{ padding: "5px 14px", borderRadius: 20, border: `1px solid ${selectedYear === y ? '#0284c7' : T.border}`, fontWeight: 700, fontSize: 13, cursor: "pointer",
                  background: selectedYear === y ? '#0284c7' : T.card, color: selectedYear === y ? "#fff" : T.muted }}>
                {y}년
              </button>
            ))}
          </div>
          {myRequests.filter(r => r.date?.startsWith(selectedYear)).length === 0
            ? <div style={{ textAlign: "center", color: T.muted, padding: 24, background: T.card, borderRadius: 12, border: `1px solid ${T.border}` }}>{selectedYear}년 신청 내역 없음</div>
            : myRequests.filter(r => r.date?.startsWith(selectedYear)).map(r => (
              <div key={r.id} style={{ background: T.card, borderRadius: 12, padding: "12px 14px", marginBottom: 8, border: `1px solid ${T.border}`, display: "flex", alignItems: "center", gap: 10 }}>
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 700, fontSize: 13, color: T.text }}>{r.date} · {r.type}{r.type === "시간연차" && r.hours ? ` (${r.hours}시간)` : ""}</div>
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
          <div style={{ fontSize: 13, color: T.muted, margin: "16px 0 8px", fontWeight: 600 }}>연차 신청 목록</div>
          <div style={{ display: "flex", gap: 6, marginBottom: 12, flexWrap: "wrap" }}>
            {years.map(y => (
              <button key={y} onClick={() => setSelectedYear(y)}
                style={{ padding: "5px 14px", borderRadius: 20, border: "none", fontWeight: 700, fontSize: 13, cursor: "pointer",
                  background: selectedYear === y ? '#0284c7' : T.card,
                  color: selectedYear === y ? "#fff" : T.muted,
                  border: `1px solid ${selectedYear === y ? '#0284c7' : T.border}` }}>
                {y}년
              </button>
            ))}
          </div>
          {leaveRequests.filter(r => r.date?.startsWith(selectedYear)).length === 0
            ? <div style={{ textAlign: "center", color: T.muted, padding: 24, background: T.card, borderRadius: 12, border: `1px solid ${T.border}` }}>{selectedYear}년 신청 없음</div>
            : leaveRequests.filter(r => r.date?.startsWith(selectedYear)).map(r => {
              const sc = { "대기": "yellow", "승인": "green", "반려": "red" };
              return <LeaveRequestItem key={r.id} r={r} statusColor={sc} setDelConfirm={setDelConfirm} />;
            })
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
                <span style={{ color: T.blue, fontSize: 12 }}>팀원에게 자동으로 공지가 발송돼요</span>
              </div>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              <Btn variant="ghost" onClick={() => setDelConfirm(null)}>취소</Btn>
              <Btn variant="red" onClick={async () => {
                await delReq(delConfirm);
                await addDoc(collection(db, COL_NOTICES), {
                  title: "📅 연차 신청 삭제 안내",
                  content: `${delConfirm.date} ${delConfirm.type} 신청이 삭제되었습니다.\n문의사항은 관리자에게 연락해주세요.`,
                  recipient: delConfirm.userId, author: "관리자",
                  createdAt: new Date().toISOString(), auto: true
                });
              }}>삭제 + 공지</Btn>
            </div>
          </div>
        </div>
      )}
      </div>
    </div>
  );
}

// ── 하단 탭바 ────────────────────────────────────────────────────
function TabBar({ tab, setTab, isAdmin, leaveRequests, notices, board, payslips, user, reads }) {
  const pendingCount = leaveRequests.filter(r => r.status === "대기").length;

  const unreadCount = (items, type) => {
    if (isAdmin || !user || !reads) return 0;
    return items.filter(item => !reads[`${user.id}_${type}_${item.id}`] &&
      (type !== "notice" || item.recipient === "all" || item.recipient === user.id) &&
      (type !== "board" || item.userId !== user.id)
    ).length;
  };

  const unreadNotice = unreadCount(notices, "notice");
  const unreadBoard = unreadCount(board, "board");
  const unreadPayslip = unreadCount(payslips.filter(p => p.userId === user?.id), "payslip");

  const tabs = [
    ["att", "🏠", "출퇴근", 0],
    ["notice", "📢", "공지", unreadNotice],
    ["board", "💬", "게시판", unreadBoard],
    ["payslip", "💰", "명세서", unreadPayslip],
    ["annual", "📅", "연차", isAdmin ? pendingCount : 0],
  ];

  return (
    <div style={{ position: "fixed", bottom: 0, left: 0, right: 0, background: T.card, borderTop: `1px solid ${T.border}`, display: "flex", zIndex: 50, paddingBottom: "env(safe-area-inset-bottom)" }}>
      {tabs.map(([key, icon, label, badge]) => (
        <button key={key} onClick={() => setTab(key)}
          style={{ flex: 1, padding: "10px 0 8px", border: "none", background: "none", cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", gap: 2, position: "relative" }}>
          <span style={{ fontSize: 20 }}>{icon}</span>
          <span style={{ fontSize: 10, fontWeight: tab===key?800:500, color: tab===key?T.adminHeader:T.muted }}>{label}</span>
          {badge > 0 && (
            <div style={{ position: "absolute", top: 6, right: "calc(50% - 16px)", background: T.red, color: "#fff", borderRadius: "50%", width: 16, height: 16, fontSize: 10, fontWeight: 800, display: "flex", alignItems: "center", justifyContent: "center" }}>{badge}</div>
          )}
          {tab === key && <div style={{ position: "absolute", bottom: 0, left: "20%", right: "20%", height: 2, background: T.adminHeader, borderRadius: 2 }} />}
        </button>
      ))}
    </div>
  );
}

// ── 메인 App ───────────────────────────────────────────────────
function App({ users, settings, records, leaves, notices, board, payslips, annual, leaveRequests, memberInfo, reads, reminders = [], onSaveUsers, onSaveSettings, onSaveRecord, onSaveLeave }) {
  const [user, setUser] = useState(null);
  const [tab, setTab] = useState("att");

  if (!user) return <LoginScreen users={users} onLogin={setUser} onUpdateUsers={onSaveUsers} />;

  const isAdmin = user.role === "admin";

  // 관리자는 대문+섹션 구조 (탭바 없음)
  if (isAdmin) return (
    <AdminScreen user={user} users={users} settings={settings} records={records} leaves={leaves}
      notices={notices} board={board} payslips={payslips} annual={annual} leaveRequests={leaveRequests} memberInfo={memberInfo} reads={reads}
      reminders={reminders}
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
      {tab !== "att" && <FloatBack onClick={() => setTab("att")} />}
      {tab === "notice" && (
        <>
          <div style={{ background: T.headerBg, padding: "18px 16px 14px" }}>
            <div style={{ fontSize: 11, color: "#ffffff50", letterSpacing: 3 }}>ATTENDANCE</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>공지사항</div>
          </div>
          <NoticeScreen user={user} users={users} notices={notices} reads={reads} />
        </>
      )}
      {tab === "board" && (
        <>
          <div style={{ background: T.headerBg, padding: "18px 16px 14px" }}>
            <div style={{ fontSize: 11, color: "#ffffff50", letterSpacing: 3 }}>ATTENDANCE</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>자유게시판</div>
          </div>
          <BoardScreen user={user} board={board} reads={reads} />
        </>
      )}
      {tab === "payslip" && (
        <>
          <div style={{ background: T.headerBg, padding: "18px 16px 14px" }}>
            <div style={{ fontSize: 11, color: "#ffffff50", letterSpacing: 3 }}>ATTENDANCE</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>급여명세서</div>
          </div>
          <PayslipScreen user={user} users={users} payslips={payslips} reads={reads} />
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
      <TabBar tab={tab} setTab={t => { setTab(t); window.scrollTo(0, 0); }} isAdmin={isAdmin} leaveRequests={leaveRequests} notices={notices} board={board} payslips={payslips} user={user} reads={reads} />
    </div>
  );
}

export default AppLoader;
