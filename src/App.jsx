import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
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
const COL_EVENTS   = "schedule_events";
const COL_CONTRACTS = "contracts"; // 하위호환 유지
const COL_DOCS = "contracts";      // 문서함 (동일 컬렉션 사용)
const COL_VAULT = "vault";

// 문서 종류
const DOC_TYPES = [
  { key: "contract",   label: "근로계약서", icon: "📋" },
  { key: "agreement",  label: "동의서",     icon: "✍️" },
  { key: "confirm",    label: "확인서",     icon: "✅" },
  { key: "other",      label: "기타 문서",  icon: "📄" },
];

// 동의서 템플릿
const AGREEMENT_TEMPLATES = [
  { key: "overtime", label: "연장·휴일근로 포괄 동의서",
    content: `본인은 업무상 필요에 의한 연장근로 및 휴일근로에 대해 사전에 포괄적으로 동의하며, 해당 수당은 근로계약서에 명시된 임금 구성 항목에 따라 별도 지급받음을 확인합니다.\n\n단, 연장근로는 1주 12시간을 초과할 수 없으며, 본 동의는 근로자의 자유의사에 의한 것으로 언제든지 철회할 수 있습니다.` },
  { key: "privacy", label: "개인정보 수집·이용 동의서",
    content: `본인은 아래와 같이 개인정보를 수집·이용하는 것에 동의합니다.\n\n1. 수집 항목: 성명, 주민등록번호, 주소, 연락처, 계좌정보\n2. 수집 목적: 근로계약 체결, 급여 지급, 4대보험 처리\n3. 보유 기간: 퇴직 후 3년\n4. 동의를 거부할 권리가 있으나, 거부 시 근로계약 체결이 어려울 수 있습니다.` },
  { key: "custom", label: "직접 입력", content: "" },
];

// 확인서 템플릿
const CONFIRM_TEMPLATES = [
  { key: "wage", label: "임금변경 확인서",
    content: `본인은 아래와 같이 임금이 변경됨을 확인하고 동의합니다.\n\n변경 시급: 원\n적용 시작일: 년  월  일\n\n위 변경 내용은 근로기준법 및 최저임금법에 따른 것으로, 기존 근로계약서의 해당 조항을 본 확인서로 갈음합니다.` },
  { key: "safety", label: "안전보건교육 이수 확인서",
    content: `본인은 아래 안전보건교육을 이수하였음을 확인합니다.\n\n교육 일시: 년  월  일\n교육 내용: 산업안전보건법에 따른 정기 안전보건교육\n교육 시간: 시간\n\n위 교육을 성실히 이수하였으며 그 내용을 숙지하였습니다.` },
  { key: "harassment", label: "직장 내 성희롱 예방교육 이수 확인서",
    content: `본인은 아래 직장 내 성희롱 예방교육을 이수하였음을 확인합니다.\n\n교육 일시: 년  월  일\n교육 내용: 남녀고용평등과 일·가정 양립 지원에 관한 법률에 따른 직장 내 성희롱 예방교육\n교육 시간: 시간\n\n위 교육을 성실히 이수하였으며 그 내용을 숙지하였습니다.` },
  { key: "disability", label: "직장 내 장애인 인식개선 교육 이수 확인서",
    content: `본인은 아래 직장 내 장애인 인식개선 교육을 이수하였음을 확인합니다.\n\n교육 일시: 년  월  일\n교육 내용: 장애인고용촉진 및 직업재활법에 따른 직장 내 장애인 인식개선 교육\n교육 시간: 시간\n\n위 교육을 성실히 이수하였으며 그 내용을 숙지하였습니다.` },
  { key: "privacy", label: "개인정보 보호교육 이수 확인서",
    content: `본인은 아래 개인정보 보호교육을 이수하였음을 확인합니다.\n\n교육 일시: 년  월  일\n교육 내용: 개인정보 보호법에 따른 개인정보 보호교육\n교육 시간: 시간\n\n위 교육을 성실히 이수하였으며 그 내용을 숙지하였습니다.` },
  { key: "rule", label: "취업규칙 교부 확인서",
    content: `본인은 취업규칙을 교부받았으며 그 내용을 충분히 읽고 이해하였음을 확인합니다.\n\n교부 일시: 년  월  일` },
  { key: "custom", label: "직접 입력", content: "" },
];

const DOC_TYPE_MAP = Object.fromEntries(DOC_TYPES.map(d => [d.key, d]));

// 동의서/확인서 기본 템플릿
const DOC_TEMPLATES = {
  agreement: {
    title: "연장·휴일근로 포괄 동의서",
    content: `본인은 업무상 필요에 의한 연장근로 및 휴일근로에 대해 사전에 포괄적으로 동의합니다.\n\n해당 수당은 근로계약서에 명시된 임금 구성 항목에 따라 별도 지급됩니다.\n\n단, 1주 연장근로는 12시간을 초과하지 않으며, 본 동의는 자유의사에 의한 것으로 언제든지 철회할 수 있습니다.`,
  },
  confirm: {
    title: "임금 변경 확인서",
    content: `위 당사자는 아래와 같이 임금 변경 사항을 확인합니다.\n\n변경 내용:\n- 변경 시급: \n- 적용 일자: \n- 사유: 최저임금 변동\n\n이 확인서는 기존 근로계약서의 임금 조항을 갈음합니다.`,
  },
  other: {
    title: "",
    content: "",
  },
}
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
  // 오후 반차: 퇴근 기준 = 점심 시작 시간 (lunchStart)
  if (leave?.type === "반차(오후)") {
    const lunchStart = settings?.lunchStart || "11:30";
    return calcEarlyOutMin(outI, lunchStart);
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
  newUsers.forEach(u => batch.set(doc(db, COL_USERS, u.id), u));
  if (allUsers) {
    const newIds = newUsers.map(u => u.id);
    allUsers.filter(u => !newIds.includes(u.id)).forEach(u => {
      batch.delete(doc(db, COL_USERS, u.id));
    });
  }
  await batch.commit();
}

async function fbRetireUser(user) {
  await setDoc(doc(db, COL_USERS, user.id), { ...user, status: "retired", retiredAt: new Date().toISOString() });
}

async function fbRestoreUser(user) {
  await setDoc(doc(db, COL_USERS, user.id), { ...user, status: "active", retiredAt: null });
}

async function fbDeleteUserCompletely(userId) {
  await deleteDoc(doc(db, COL_USERS, userId));
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
    ghost: { background: "#fff", color: T.text, border: `1px solid ${T.border}` },
    yellow: { background: "#d97706", color: "#fff" }
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
  const [scheduleEvents, setScheduleEvents] = useState([]);
  const [contracts, setContracts] = useState([]);
  const [vault, setVault] = useState([]);
  const [educations, setEducations] = useState([]);

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

    // 일정 이벤트 구독
    unsubs.push(onSnapshot(query(collection(db, COL_EVENTS), orderBy("date", "asc")), snap => {
      setScheduleEvents(snap.docs.map(d => ({ id: d.id, ...d.data() })));
    }));

    // 근로계약서 구독
    unsubs.push(onSnapshot(query(collection(db, COL_CONTRACTS), orderBy("createdAt", "desc")), snap => {
      setContracts(snap.docs.map(d => ({ id: d.id, ...d.data() })));
    }));

    // 보관함 구독
    unsubs.push(onSnapshot(query(collection(db, COL_VAULT), orderBy("createdAt", "desc")), snap => {
      setVault(snap.docs.map(d => ({ id: d.id, ...d.data() })));
    }));
    // 교육 구독
    unsubs.push(onSnapshot(query(collection(db, "education"), orderBy("createdAt", "desc")),
      snap => { setEducations(snap.docs.map(d => ({ id: d.id, ...d.data() }))); },
      () => { setEducations([]); }
    ));
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
    reminders={reminders} scheduleEvents={scheduleEvents} contracts={contracts} vault={vault}
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
    if (u && u.status === "retired") { setErr("퇴직 처리된 계정입니다. 관리자에게 문의하세요."); return; }
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
            // 알림 클릭 시 앱이 열려있으면 화면 유지
            OneSignal.Notifications.addEventListener("click", (e) => {
              e.preventDefault();
            });
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
function MemberScreen({ user, settings, records, leaves, onSaveRecord, onLogout, scheduleEvents = [] }) {
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
    // 출근 알림은 Firebase Functions onCheckIn에서 발송
    if (type === "out") await sendPush({ title: "🏠 퇴근", message: `${user.name}님이 퇴근했습니다.`, targetUserId: "admin" });
    setFlash(msgs[type]); setTimeout(() => setFlash(null), 2500);
  };

  const thisMonth = new Date(now.getTime() + 9 * 60 * 60 * 1000).toISOString().slice(0, 7);
  const [selectedMonth, setSelectedMonth] = useState(thisMonth);
  const [calView, setCalView] = useState("list"); // "list" | "calendar" | "schedule"
  const [scrollToDate, setScrollToDate] = useState(null);
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
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
          <div style={{ fontSize: 13, color: T.muted, fontWeight: 600 }}>{monthLabel(selectedMonth)} 기록</div>
          <div style={{ display: "flex", gap: 4 }}>
            {[["일자별", "list"], ["월별", "calendar"]].map(([label, key]) => (
              <button key={key} onClick={() => setCalView(key)}
                style={{ padding: "4px 12px", borderRadius: 16, border: `1px solid ${calView === key ? T.headerBg : T.border}`, fontSize: 11, fontWeight: 700, cursor: "pointer",
                  background: calView === key ? T.headerBg : T.card, color: calView === key ? "#fff" : T.muted }}>
                {label}
              </button>
            ))}
          </div>
        </div>
        {calView === "calendar" && (
          <CalendarView
            monthDays={monthDays}
            leaves={leaves[user.id] || {}}
            selectedMonth={selectedMonth}
            settings={settings}
            onSelectDate={(date) => {
              setCalView("list");
              setScrollToDate(date);
              setTimeout(() => {
                const el = document.getElementById(`rec-${date}`);
                if (el) el.scrollIntoView({ behavior: "smooth", block: "center" });
              }, 100);
              setTimeout(() => setScrollToDate(null), 1500);
            }}
          />
        )}
        {calView === "list" && (monthDays.length === 0
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
              <div key={date} id={`rec-${date}`} style={{ background: scrollToDate === date ? "#eff6ff" : T.card, borderRadius: 12, padding: "12px 14px", marginBottom: 8, border: `1px solid ${scrollToDate === date ? "#2563eb" : T.border}`, transition: "all 0.3s" }}>
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
        )}
      </div>
    </div>
  );
}
// ── 팀원 개인 일정 캘린더 ──────────────────────────────────────────
function MemberScheduleCalendar({ settings = {}, scheduleEvents = [], userId }) {
  const kstNow = new Date(new Date().getTime() + 9 * 60 * 60 * 1000);
  const currentMonth = kstNow.toISOString().slice(0, 10).slice(0, 7);
  const [selectedMonth, setSelectedMonth] = useState(currentMonth);
  const [selDate, setSelDate] = useState(null);
  const [editMode, setEditMode] = useState(null); // null | "add" | "edit"
  const [editTarget, setEditTarget] = useState(null);
  const [evTitle, setEvTitle] = useState("");
  const [evColor, setEvColor] = useState("#0891b2");
  const [evNote, setEvNote] = useState("");
  const [saving, setSaving] = useState(false);

  const [y, m] = selectedMonth.split("-").map(Number);
  const holidays = settings.holidays || [];
  const pad = (n) => String(n).padStart(2, "0");
  const getDateStr = (d) => `${y}-${pad(m)}-${pad(d)}`;
  const firstDay = new Date(y, m - 1, 1).getDay();
  const daysInMonth = new Date(y, m, 0).getDate();
  const today = kstNow.toISOString().slice(0, 10);

  const prevMonth = () => {
    const d = new Date(y, m - 2, 1);
    setSelectedMonth(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`);
    setSelDate(null); setEditMode(null);
  };
  const nextMonth = () => {
    const d = new Date(y, m, 1);
    setSelectedMonth(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`);
    setSelDate(null); setEditMode(null);
  };

  // 날짜별 이벤트 맵 — 본인 일정 + 공휴일만
  const eventMap = {};
  holidays.forEach(h => {
    const date = typeof h === "string" ? h : h.date;
    const memo = typeof h === "object" ? h.memo : null;
    if (date?.startsWith(selectedMonth)) {
      if (!eventMap[date]) eventMap[date] = [];
      eventMap[date].push({ type: "holiday", label: memo || "공휴일", color: "#dc2626", id: null });
    }
  });
  scheduleEvents
    .filter(ev => {
      const t = ev.target || ev.userId;
      return t === "all" || t === userId;
    })
    .forEach(ev => {
      if (ev.date?.startsWith(selectedMonth)) {
        if (!eventMap[ev.date]) eventMap[ev.date] = [];
        eventMap[ev.date].push({ type: "event", label: ev.title, color: ev.color || "#0891b2", id: ev.id, note: ev.note, target: ev.target || ev.userId, isAdminPost: ev.isAdminPost || false });
      }
    });

  // 주간 그리드
  const weeks = [];
  let days = [];
  for (let i = 0; i < firstDay; i++) days.push(null);
  for (let d = 1; d <= daysInMonth; d++) {
    days.push(d);
    if (days.length === 7) { weeks.push(days); days = []; }
  }
  while (days.length < 7) days.push(null);
  if (days.some(d => d)) weeks.push(days);

  const selEvents = selDate ? (eventMap[selDate] || []) : [];

  const openAdd = () => { setEvTitle(""); setEvColor("#0891b2"); setEvNote(""); setEditTarget(null); setEditMode("add"); };
  const openEdit = (ev) => { setEvTitle(ev.label); setEvColor(ev.color || "#0891b2"); setEvNote(ev.note || ""); setEditTarget(ev); setEditMode("edit"); };
  const cancelEdit = () => { setEditMode(null); setEditTarget(null); };

  const saveEvent = async () => {
    if (!evTitle.trim() || !selDate) return;
    setSaving(true);
    try {
      if (editMode === "add") {
        await addDoc(collection(db, COL_EVENTS), {
          date: selDate, title: evTitle.trim(), color: evColor, note: evNote.trim(),
          userId, target: userId, isAdminPost: false, createdAt: new Date().toISOString()
        });
      } else if (editMode === "edit" && editTarget?.id) {
        await setDoc(doc(db, COL_EVENTS, editTarget.id), {
          date: selDate, title: evTitle.trim(), color: evColor, note: evNote.trim(),
          userId, target: userId, isAdminPost: false, createdAt: editTarget.createdAt || new Date().toISOString()
        });
      }
      cancelEdit();
    } catch(e) { console.error(e); }
    setSaving(false);
  };

  const deleteEvent = async (id) => { if (id) await deleteDoc(doc(db, COL_EVENTS, id)); };

  const EVENT_COLORS = [
    ["#0891b2","청록"],["#7c3aed","보라"],["#16a34a","초록"],
    ["#d97706","주황"],["#dc2626","빨강"],["#1d4ed8","파랑"],
  ];

  return (
    <div style={{ paddingBottom: 140 }}>
      {/* 월 이동 */}
      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 14, marginTop: 4 }}>
        <button onClick={prevMonth}
          style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 10, padding: "8px 13px", fontSize: 15, cursor: "pointer", fontWeight: 700, color: T.text }}>‹</button>
        <div style={{ flex: 1, textAlign: "center", fontSize: 15, fontWeight: 800, color: T.text }}>{y}년 {m}월</div>
        <button onClick={nextMonth}
          style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 10, padding: "8px 13px", fontSize: 15, cursor: "pointer", fontWeight: 700, color: T.text }}>›</button>
      </div>

      {/* 캘린더 그리드 */}
      <div style={{ background: T.card, borderRadius: 14, border: `1px solid ${T.border}`, overflow: "hidden", marginBottom: 12 }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(7,1fr)" }}>
          {["일","월","화","수","목","금","토"].map((d, i) => (
            <div key={d} style={{ textAlign: "center", padding: "7px 0", fontSize: 10, fontWeight: 700,
              color: i === 0 ? "#dc2626" : i === 6 ? "#2563eb" : T.muted,
              borderBottom: `1px solid ${T.border}` }}>{d}</div>
          ))}
        </div>
        {weeks.map((week, wi) => (
          <div key={wi} style={{ display: "grid", gridTemplateColumns: "repeat(7,1fr)" }}>
            {week.map((d, di) => {
              const dateStr = d ? getDateStr(d) : null;
              const events = d ? (eventMap[dateStr] || []) : [];
              const isToday = dateStr === today;
              const isSelected = !!d && dateStr === selDate;
              const isHol = d ? isHoliday(dateStr, holidays) : false;
              const hasHol = events.some(e => e.type === "holiday");
              let dateColor = T.muted;
              if (di === 6) dateColor = "#2563eb";
              else if (di === 0 || isHol || hasHol) dateColor = "#dc2626";
              return (
                <div key={di}
                  onClick={() => { if (!d) return; setSelDate(selDate === dateStr ? null : dateStr); setEditMode(null); setEditTarget(null); }}
                  style={{
                    height: 80, boxSizing: "border-box", padding: "4px 3px 3px",
                    borderBottom: wi < weeks.length - 1 ? `1px solid ${T.border}` : "none",
                    borderRight: di < 6 ? `1px solid ${T.border}` : "none",
                    display: "flex", flexDirection: "column", alignItems: "stretch", gap: 1,
                    cursor: d ? "pointer" : "default",
                    background: isSelected ? "#e0f2fe" : isToday ? "#f0f9ff" : "transparent",
                    overflow: "hidden", transition: "background 0.15s"
                  }}>
                  {d && <>
                    <div style={{
                      fontSize: 12, fontWeight: isToday ? 900 : 600,
                      color: di === 6 ? "#2563eb" : (di === 0 || isHol || hasHol) ? "#dc2626" : isToday ? T.headerBg : dateColor,
                      width: 20, height: 20, borderRadius: "50%",
                      display: "flex", alignItems: "center", justifyContent: "center",
                      background: isToday ? "#dbeafe" : "transparent",
                      flexShrink: 0, alignSelf: "center"
                    }}>{d}</div>
                    <CalEventLabels events={events} />
                  </>}
                </div>
              );
            })}
          </div>
        ))}
      </div>

      {/* 범례 */}
      <div style={{ display: "flex", gap: 10, marginBottom: 14, justifyContent: "flex-end" }}>
        {[["#dc2626","공휴일"],["#0891b2","내 일정"]].map(([color,label]) => (
          <div key={label} style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 10, fontWeight: 700 }}>
            <div style={{ width: 7, height: 7, borderRadius: "50%", background: color }} />
            <span style={{ color: T.muted }}>{label}</span>
          </div>
        ))}
      </div>

      {/* 선택 날짜 패널 */}
      {selDate && (
        <div style={{ background: T.card, borderRadius: 14, border: `1px solid ${T.border}`, overflow: "hidden", marginBottom: 12 }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between",
            padding: "12px 14px", borderBottom: `1px solid ${T.border}`, background: "#f0f9ff" }}>
            <div style={{ fontSize: 14, fontWeight: 800, color: T.text }}>{formatDate(selDate)}</div>
            <button onClick={openAdd}
              style={{ background: T.headerBg, border: "none", color: "#fff", borderRadius: 10,
                padding: "6px 14px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
              + 일정 추가
            </button>
          </div>
          <div style={{ padding: "8px 14px" }}>
            {selEvents.length === 0 && editMode !== "add" ? (
              <div style={{ fontSize: 13, color: T.muted, textAlign: "center", padding: "16px 0" }}>등록된 일정 없음</div>
            ) : selEvents.map((ev, i) => (
              <div key={i}>
                {editMode === "edit" && editTarget?.id === ev.id ? (
                  <EventForm evTitle={evTitle} setEvTitle={setEvTitle} evColor={evColor} setEvColor={setEvColor}
                    evNote={evNote} setEvNote={setEvNote} EVENT_COLORS={EVENT_COLORS}
                    saving={saving} onSave={saveEvent} onCancel={cancelEdit} label="수정" />
                ) : (
                  <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "9px 0",
                    borderBottom: i < selEvents.length - 1 ? `1px solid ${T.border}` : "none" }}>
                    <div style={{ width: 8, height: 8, borderRadius: "50%", background: ev.color, flexShrink: 0 }} />
                    <div style={{ flex: 1 }}>
                      <div style={{ fontSize: 13, fontWeight: 700, color: ev.type === "holiday" ? "#dc2626" : T.text }}>{ev.label}</div>
                      {ev.note && <div style={{ fontSize: 11, color: T.muted, marginTop: 2 }}>{ev.note}</div>}
                    </div>
                    {ev.type === "holiday" && <span style={{ fontSize: 10, color: "#dc2626", fontWeight: 700 }}>공휴일</span>}
                    {ev.type === "event" && ev.target === "all" && (
                      <span style={{ fontSize: 10, color: "#7c3aed", fontWeight: 700, background: "#ede9fe", borderRadius: 6, padding: "2px 7px" }}>관리자 일정</span>
                    )}
                    {ev.type === "event" && ev.target === userId && ev.isAdminPost && (
                      <span style={{ fontSize: 10, color: "#0891b2", fontWeight: 700, background: "#e0f2fe", borderRadius: 6, padding: "2px 7px" }}>개인 일정</span>
                    )}
                    {ev.type === "event" && ((ev.target === userId && !ev.isAdminPost) || (!ev.target && !ev.isAdminPost)) && (
                      // 본인 일정 — 수정/삭제
                      <div style={{ display: "flex", gap: 4 }}>
                        <button onClick={() => openEdit(ev)}
                          style={{ background: T.bg, border: `1px solid ${T.border}`, color: T.muted,
                            borderRadius: 7, padding: "3px 9px", fontSize: 11, cursor: "pointer", fontWeight: 600 }}>수정</button>
                        <button onClick={() => { if (window.confirm("일정을 삭제할까요?")) deleteEvent(ev.id); }}
                          style={{ background: T.redBg, border: "none", color: T.red,
                            borderRadius: 7, padding: "3px 9px", fontSize: 11, cursor: "pointer", fontWeight: 700 }}>삭제</button>
                      </div>
                    )}
                  </div>
                )}
              </div>
            ))}
            {editMode === "add" && (
              <div style={{ borderTop: selEvents.length > 0 ? `1px solid ${T.border}` : "none", paddingTop: selEvents.length > 0 ? 10 : 0 }}>
                <EventForm evTitle={evTitle} setEvTitle={setEvTitle} evColor={evColor} setEvColor={setEvColor}
                  evNote={evNote} setEvNote={setEvNote} EVENT_COLORS={EVENT_COLORS}
                  saving={saving} onSave={saveEvent} onCancel={cancelEdit} label="추가" />
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

// ── 캘린더 이벤트 라벨 — 우선순위 정렬 후 칸 여유만큼 표시 ────────
const LINE_H = 14;
const TOTAL_LINES = 3; // 칸 높이(80px) - 날짜(22px) - 패딩(8px) / 줄높이(14px)

function CalEventLabels({ events }) {
  const ORDER = { holiday: 0, reminder: 1, event: 2 };
  const sorted = [...events].sort((a, b) => (ORDER[a.type] ?? 9) - (ORDER[b.type] ?? 9));

  const overflow = sorted.length > TOTAL_LINES;
  const showEvents = overflow ? sorted.slice(0, TOTAL_LINES - 1) : sorted;

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 1, overflow: "hidden", flex: 1 }}>
      {showEvents.map((ev, ei) => (
        <div key={ei} style={{
          fontSize: 9, fontWeight: 700,
          color: ev.type === "holiday" ? "#dc2626" : T.text,
          background: "transparent",
          borderLeft: `2px solid ${ev.color}`,
          paddingLeft: 3,
          flex: 1,           // 남은 공간 균등 배분
          overflow: "hidden",
          wordBreak: "break-all",
          lineHeight: `${LINE_H}px`,
        }}>{ev.label}</div>
      ))}
      {overflow && (
        <div style={{ fontSize: 8, color: T.muted, fontWeight: 800, paddingLeft: 3, lineHeight: `${LINE_H}px`, flexShrink: 0 }}>···</div>
      )}
    </div>
  );
}

// ── 캘린더 뷰 컴포넌트 ───────────────────────────────────────────
function CalendarView({ monthDays, leaves, selectedMonth, settings, onSelectDate }) {
  const T_local = T;
  const [y, m] = selectedMonth.split("-").map(Number);
  const firstDay = new Date(y, m - 1, 1).getDay();
  const daysInMonth = new Date(y, m, 0).getDate();
  const recMap = Object.fromEntries(monthDays);
  const holidayList = settings.holidays || [];

  const statusConfig = {
    normal:      { color: "#16a34a", label: "출근" },
    overtime:    { color: "#2563eb", label: "잔업" },
    annual:      { color: "#7c3aed", label: "연차" },
    holidayWork: { color: "#ea580c", label: "휴일근무" },
  };

  const getStatus = (dateStr) => {
    const rec = recMap[dateStr];
    const leave = leaves[dateStr];
    const isHol = isHoliday(dateStr, holidayList);
    if (leave) return "annual";
    if (!rec) return null;
    if (isHol && rec.in) return "holidayWork";
    const om = calcTotalOvertimeMin(rec.in, rec.out, settings.workStart, settings.workEnd);
    if (om >= 30) return "overtime";
    return "normal";
  };

  const getHolidayMemo = (dateStr) => {
    const h = holidayList.find(h => (typeof h === "string" ? h : h.date) === dateStr);
    return h && typeof h === "object" ? h.memo : null;
  };

  const weeks = [];
  let days = [];
  for (let i = 0; i < firstDay; i++) days.push(null);
  for (let d = 1; d <= daysInMonth; d++) {
    days.push(d);
    if (days.length === 7) { weeks.push(days); days = []; }
  }
  while (days.length < 7) days.push(null);
  if (days.some(d => d)) weeks.push(days);

  const pad = (n) => String(n).padStart(2, "0");
  const getDateStr = (d) => `${y}-${pad(m)}-${pad(d)}`;

  return (
    <div>
      {/* 캘린더 */}
      <div style={{ background: T_local.card, borderRadius: 14, border: `1px solid ${T_local.border}`, overflow: "hidden", marginBottom: 12 }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(7,1fr)" }}>
          {["일","월","화","수","목","금","토"].map((d, i) => (
            <div key={d} style={{ textAlign: "center", padding: "7px 0", fontSize: 10, fontWeight: 700,
              color: i === 0 ? "#dc2626" : i === 6 ? "#2563eb" : T_local.muted,
              borderBottom: `1px solid ${T_local.border}` }}>{d}</div>
          ))}
        </div>
        {weeks.map((week, wi) => (
          <div key={wi} style={{ display: "grid", gridTemplateColumns: "repeat(7,1fr)" }}>
            {week.map((d, di) => {
              const dateStr = d ? getDateStr(d) : null;
              const status = d ? getStatus(dateStr) : null;
              const cfg = status ? statusConfig[status] : null;
              const memo = d ? getHolidayMemo(dateStr) : null;
              const isWeekend = di === 0 || di === 6;
              const isHolDay = d ? isHoliday(dateStr, holidayList) : false;
              let dateColor = T_local.text;
              if (cfg) dateColor = cfg.color;
              else if (di === 0) dateColor = "#dc2626";
              else if (di === 6) dateColor = "#2563eb";
              else if (isHolDay) dateColor = "#dc2626";
              else dateColor = T_local.muted;
              return (
                <div key={di} onClick={() => d && recMap[dateStr] && onSelectDate(dateStr)}
                  style={{ padding: "5px 3px", borderBottom: `1px solid ${T_local.border}`,
                    borderRight: di < 6 ? `1px solid ${T_local.border}` : "none",
                    minHeight: 48, display: "flex", flexDirection: "column", alignItems: "center", gap: 1,
                    cursor: d && recMap[dateStr] ? "pointer" : "default",
                    background: d && recMap[dateStr] ? "transparent" : "transparent" }}>
                  {d && <>
                    <div style={{ fontSize: 12, fontWeight: 600, color: dateColor }}>{d}</div>
                    {cfg && <div style={{ fontSize: 9, fontWeight: 700, color: cfg.color }}>{cfg.label}</div>}
                    {memo && <div style={{ fontSize: 8, color: T_local.muted, textAlign: "center", lineHeight: 1.2 }}>{memo}</div>}
                  </>}
                </div>
              );
            })}
          </div>
        ))}
      </div>
      {/* 범례 - 오른쪽 아래 */}
      <div style={{ display: "flex", flexWrap: "wrap", gap: "2px 10px", marginBottom: 10, justifyContent: "flex-end" }}>
        {Object.entries(statusConfig).map(([k, v]) => (
          <div key={k} style={{ fontSize: 10, fontWeight: 700, color: v.color }}>{v.label}</div>
        ))}
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

  // 반차 자동 제외 처리
  const isHalfAM = leaveType === "반차(오전)";
  const isHalfPM = leaveType === "반차(오후)";
  const autoLate = isHalfAM ? false : calcLateMin(inIso, settings.workStart) > 0;
  const autoEarly = isHalfPM ? false : calcEarlyOutMin(outIso, settings.workEnd) > 0;
  const autoOvertime = calcTotalOvertimeMin(inIso, outIso, settings.workStart, settings.workEnd) >= 30;
  const lm = isHalfAM ? 0 : calcLateMin(inIso, settings.workStart);
  const em = isHalfPM ? 0 : calcEarlyOutMinWithLeave(outIso, settings.workEnd, inIso, settings.workStart, userLeaves[date], settings);
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
  const [addDateMode, setAddDateMode] = useState(false);
  const [addDate, setAddDate] = useState("");

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
      const header = ["날짜", "요일", "출근", "퇴근", "지각", "지각시간", "조퇴", "조퇴시간", "잔업", "잔업시간", "외출횟수", "외출시간", "연차/반차", "메모"];
      const rows = days.map(([date, rec]) => {
        const dow = new Date(date).toLocaleDateString("ko-KR", { weekday: "short" });
        const lm = calcLateMin(rec.in, settings.workStart), em = calcEarlyOutMin(rec.out, settings.workEnd), om = calcTotalOvertimeMin(rec.in, rec.out, settings.workStart, settings.workEnd);
        const leave = userLeaves[date];
        const outings = rec.outing || []; const outingStr = outings.map(o => formatTime(o.out) + "~" + formatTime(o.in)).join(" | "); return [date, dow, formatTime(rec.in), formatTime(rec.out), lm > 0 ? "O" : "", lm > 0 ? fmtMinutes(lm) : "", em > 0 ? "O" : "", em > 0 ? fmtMinutes(em) : "", om >= 30 ? "O" : "", om >= 30 ? fmtMinutes(roundTo30(om)) : "", outings.length > 0 ? outings.length + "회" : "", outingStr, leave ? leave.type : "", rec.note || ""];
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

        {/* 기록 없는 날짜 직접 추가 */}
        {!editTarget && (
          <div style={{ marginTop: 8 }}>
            {!addDateMode ? (
              <button onClick={() => setAddDateMode(true)}
                style={{ width: "100%", padding: "11px 0", borderRadius: 12, border: `2px dashed ${T.border}`, background: "none", color: T.muted, fontSize: 13, cursor: "pointer", fontWeight: 600 }}>
                + 날짜 직접 추가 (기록 없는 날)
              </button>
            ) : (
              <div style={{ background: T.card, borderRadius: 12, padding: 14, border: `1px solid ${T.border}` }}>
                <div style={{ fontSize: 13, fontWeight: 700, color: T.text, marginBottom: 10 }}>날짜 선택</div>
                <input type="date" value={addDate} onChange={e => setAddDate(e.target.value)}
                  style={{ width: "100%", padding: "9px 12px", borderRadius: 10, border: `1px solid ${T.border}`, fontSize: 13, boxSizing: "border-box", fontFamily: "inherit", marginBottom: 10 }} />
                <div style={{ display: "flex", gap: 8 }}>
                  <button onClick={() => { setAddDateMode(false); setAddDate(""); }}
                    style={{ flex: 1, padding: "9px 0", borderRadius: 10, border: `1px solid ${T.border}`, background: "none", color: T.muted, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>취소</button>
                  <button onClick={() => {
                    if (!addDate) return;
                    setEditTarget({ user: drillUser, date: addDate });
                    setAddDateMode(false); setAddDate("");
                  }}
                    style={{ flex: 2, padding: "9px 0", borderRadius: 10, border: "none", background: T.headerBg, color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                    이 날짜 수정
                  </button>
                </div>
              </div>
            )}
          </div>
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
        const header = ["이름", "날짜", "요일", "출근", "퇴근", "지각", "지각시간", "조퇴", "조퇴시간", "잔업", "잔업시간", "외출횟수", "외출시간", "연차/반차", "메모"];
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
              (rec.outing||[]).length>0?(rec.outing||[]).length+"회":"", (rec.outing||[]).map(o=>formatTime(o.out)+"~"+formatTime(o.in)).join(" | "),
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
              {[["전체", (ms.totalDays||0) + "일", T.text], ["근무", ((ms.totalDays||0)-(ms.offDays||0)) + "일", T.green], ["출근", ms.days + "일", "#2563eb"], ["휴일", ms.holiday + "일", T.red], ["연차", ms.annualDays > 0 ? ms.annualDays + "일" : "0일", "#7c3aed"]].map(([l, v, c]) => (
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
  const [pensionBase, setPensionBase] = useState("");
  const [insuranceBase, setInsuranceBase] = useState("");

  // 관리자 기초데이터 불러오기
  useEffect(() => {
    if (!admin.id) return;
    getDoc(doc(db, COL_MEMBER_INFO, admin.id)).then(snap => {
      if (snap.exists()) {
        const d = snap.data();
        if (d.pensionBase) setPensionBase(String(d.pensionBase));
        if (d.insuranceBase) setInsuranceBase(String(d.insuranceBase));
      }
    });
  }, [admin.id]);

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
          {tabBtn("info", "이름변경")}{tabBtn("pin", "PIN변경")}{tabBtn("lost", "PIN분실")}{tabBtn("insurance", "보험기준")}
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
        {tab === "insurance" && <>
          <div style={{ fontSize: 12, color: T.muted, marginBottom: 14, lineHeight: 1.6 }}>
            4대보험 계산기에서 자동으로 불러옵니다.
          </div>
          {[
            ["국민연금 기준소득월액", pensionBase, setPensionBase, "× 4.75% = 국민연금"],
            ["건강/고용보험 보수월액", insuranceBase, setInsuranceBase, "건강 ×3.595% / 고용 ×0.9%"],
          ].map(([label, val, setter, sub]) => (
            <div key={label} style={{ marginBottom: 12 }}>
              <div style={{ fontSize: 13, color: T.sub, marginBottom: 4, fontWeight: 600 }}>{label}</div>
              <div style={{ fontSize: 11, color: T.muted, marginBottom: 6 }}>{sub}</div>
              <input value={val} onChange={e => setter(e.target.value.replace(/[^0-9]/g, ""))}
                placeholder="0" style={{ ...iStyle, textAlign: "right", marginBottom: 0 }} />
            </div>
          ))}
          {err && <div style={{ color: T.red, fontSize: 13, marginBottom: 8, fontWeight: 600 }}>{err}</div>}
          {ok && <div style={{ color: T.green, fontSize: 13, marginBottom: 8, fontWeight: 600 }}>{ok}</div>}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 4 }}>
            <Btn variant="ghost" onClick={onClose}>닫기</Btn>
            <Btn variant="admin" onClick={async () => {
              if (!admin.id) return;
              await setDoc(doc(db, COL_MEMBER_INFO, admin.id), {
                pensionBase: Number(pensionBase) || 0,
                insuranceBase: Number(insuranceBase) || 0,
              }, { merge: true });
              setErr(""); setOk("저장됐어요 ✓"); setTimeout(() => setOk(""), 2000);
            }}>저장</Btn>
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
        {/* 급여 지급일 */}
        <div style={{ background: T.bg, border: `1px solid ${T.border}`, borderRadius: 14, padding: "14px 16px", marginBottom: 16 }}>
          <div style={{ fontSize: 13, color: T.text, fontWeight: 700, marginBottom: 4 }}>💰 급여 지급일</div>
          <div style={{ fontSize: 11, color: T.muted, marginBottom: 8 }}>익월 N일 지급 · 주말/공휴일이면 다음 근무일로 자동 조정</div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <div style={{ fontSize: 13, color: T.muted, fontWeight: 600 }}>익월</div>
            <input type="number" min="1" max="31" value={s.payDay ?? 15}
              onChange={e => setS(p => ({ ...p, payDay: Number(e.target.value) }))}
              style={{ ...iStyle, width: 80, textAlign: "center" }} placeholder="15" />
            <div style={{ fontSize: 13, color: T.muted, fontWeight: 600 }}>일</div>
          </div>
        </div>
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
            <button onClick={() => { if (window.confirm("등록된 위치를 삭제할까요?")) { setS(p => ({ ...p, officeLat: null, officeLng: null })); setGpsMsg("위치가 삭제됐어요"); } }} style={{ padding: "10px 0", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: T.muted, fontSize: 12, fontWeight: 600, cursor: "pointer" }}>위치 삭제</button>
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
              <button onClick={() => { setDelTarget(u); setMode("retire"); }} style={{ background: "#fff7ed", border: "none", color: "#d97706", borderRadius: 8, padding: "5px 10px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>퇴직</button>
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
        {mode === "retire" && <>
          <div style={{ textAlign: "center", padding: "20px 0" }}>
            <div style={{ fontSize: 36, marginBottom: 14 }}>👤</div>
            <div style={{ fontWeight: 800, fontSize: 17, color: T.text, marginBottom: 10 }}>{delTarget?.name} 퇴직 처리</div>
            <div style={{ fontSize: 13, color: T.muted, lineHeight: 1.7 }}>
              출퇴근·급여·계약서 등 모든 데이터는<br />보관됩니다. 언제든 복원 가능합니다.
            </div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
            <Btn variant="ghost" onClick={() => { setMode("list"); setDelTarget(null); }}>취소</Btn>
            <Btn variant="yellow" onClick={async () => {
              await fbRetireUser(delTarget);
              setMode("list"); setDelTarget(null); onClose();
            }}>퇴직 처리</Btn>
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
  const members = users.filter(u => u.role === "member" && (!u.status || u.status === "active"));

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
function AdminHome({ user, onLogout, onSection, leaveRequests = [], board = [], reads = {}, contracts = [] }) {
  const now = new Date();
  const kst = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  const dateStr = kst.toLocaleDateString("ko-KR", { year: "numeric", month: "long", day: "numeric", weekday: "long" });

  const pendingSign = contracts.filter(c => c.status === "sent").length;

  const sections = [
    { key: "attendance", icon: "📋", label: "근태",   desc: "출퇴근 현황 · 월별 기록", color: "#2563eb" },
    { key: "wage",       icon: "💰", label: "임금",   desc: "급여 계산 · 임금대장",   color: "#16a34a" },
    { key: "members",    icon: "👥", label: "팀원",   desc: "직원 정보 · 기초 데이터", color: "#7c3aed" },
    { key: "annual",     icon: "📅", label: "연차",   desc: "연차 현황 · 신청 승인",   color: "#0284c7", badge: leaveRequests.filter(r => r.status === "대기").length },
    { key: "notice",     icon: "📢", label: "공지",   desc: "공지사항 작성 · 관리",   color: "#ea580c" },
    { key: "board",      icon: "💬", label: "게시판", desc: "자유게시판",              color: "#0891b2",
      badge: board.filter(b => !reads[`${user.id}_board_${b.id}`]).length },
    { key: "contract",   icon: "📄", label: "문서함", desc: "계약서 · 동의서 · 확인서",  color: "#0891b2", badge: pendingSign },
    { key: "vault",      icon: "📁", label: "보관함", desc: "메모 · 파일 보관",        color: "#7c3aed" },
    { key: "settings",   icon: "⚙",  label: "설정",   desc: "근무시간 · GPS · 공휴일", color: "#6b7280" },
    { key: "schedule",   icon: "🗓", label: "일정",    desc: "캘린더 · 리마인더",     color: "#7c3aed" },
    { key: "severance",  icon: "💼", label: "퇴직금", desc: "퇴직금 계산",            color: "#b45309" },
    { key: "insurance",  icon: "🏦", label: "4대보험", desc: "보험료 계산 · 납부 요약", color: "#16a34a" },
    { key: "education",  icon: "🎓", label: "교육",    desc: "교육 개설 · 완료 현황",   color: "#7c3aed" },
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
  const members = users.filter(u => u.role === "member" && (!u.status || u.status === "active"));

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
                    {/* 알림 재발송 버튼 */}
                    {(() => {
                      const [sh, sm] = (settings.workStart || "09:00").split(":").map(Number);
                      const [eh, em2] = (settings.workEnd || "18:00").split(":").map(Number);
                      const nowMin = now.getHours() * 60 + now.getMinutes();
                      let startMin = sh * 60 + sm;
                      const endMin = eh * 60 + em2;
                      const todayIsHoliday = isHoliday(today, settings.holidays || []);
                      // 반차(오전)이면 점심시간 끝(lunchEnd) 기준으로 출근 알림
                      if (todayLeave?.type === "반차(오전)") {
                        const [lh, lm2] = (settings.lunchEnd || "12:30").split(":").map(Number);
                        startMin = lh * 60 + lm2;
                      }
                      // 휴일이면 출근 알림 안 보임, 퇴근 알림은 출근한 경우만
                      const showCheckin = !rec.in && nowMin >= startMin && !todayIsHoliday;
                      const showCheckout = rec.in && !rec.out && nowMin >= endMin;
                      if (showCheckin) return (
                        <button onClick={() => { if (window.confirm(`${u.name}님께 출근 알림을 보낼까요?`)) sendPush({ title: "🌅 출근 알림", message: `${u.name}님, 출근 기록을 잊지 마세요!`, targetUserId: u.id }); }}
                          style={{ background: "#fef9c3", border: "1px solid #fde68a", color: "#92400e", borderRadius: 8, padding: "4px 10px", fontSize: 11, cursor: "pointer", fontWeight: 700 }}>🔔 출근</button>
                      );
                      if (showCheckout) return (
                        <button onClick={() => { if (window.confirm(`${u.name}님께 퇴근 알림을 보낼까요?`)) sendPush({ title: "🏠 퇴근 알림", message: `${u.name}님, 퇴근 기록을 잊지 마세요!`, targetUserId: u.id }); }}
                          style={{ background: "#dbeafe", border: "1px solid #93c5fd", color: "#1e40af", borderRadius: 8, padding: "4px 10px", fontSize: 11, cursor: "pointer", fontWeight: 700 }}>🔔 퇴근</button>
                      );
                      return null;
                    })()}
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
function getPayDate(yearMonth, holidays = [], payDay = 15) {
  const [y, m] = yearMonth.split("-").map(Number);
  let d = new Date(y, m, payDay); // 익월 payDay일
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
          <div style={{ fontSize: 12, color: "#ffffff80", marginTop: 4 }}>지급일 {getPayDate(yearMonth, holidays, settings?.payDay ?? 15)} · {info?.bank ? `${info.bank}은행 ${info.account}` : "계좌 미등록"}</div>
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
              monthStats, payDate: getPayDate(yearMonth, holidays, settings?.payDay ?? 15),
              createdAt: new Date().toISOString()
            })}>확정</Btn>
          </div>
        </div>
      </div>
    </div>
  );
}

// ── 관리자 임금 섹션 ──────────────────────────────────────────
function AdminWage({ users, records, leaves, settings, memberInfo, annual, leaveRequests, payslips, reads, onBack }) {
  const members = users.filter(u => u.role === "member" && (!u.status || u.status === "active"));
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
      if (!s) return [u.name, getPayDate(selectedMonth, settings?.holidays ?? [], settings?.payDay ?? 15), ...Array(20).fill("-")];
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
        {/* 월 선택 — 급여계산/임금대장에서만 표시 */}
        {tab !== "payslip" && (
          <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 12 }}>
            <button onClick={prevMonth} style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 10, padding: "10px 16px", fontSize: 16, cursor: "pointer", fontWeight: 700, color: T.text }}>‹</button>
            <div style={{ flex: 1, textAlign: "center", fontSize: 16, fontWeight: 800, color: T.text }}>{monthLabel(selectedMonth)}</div>
            <button onClick={nextMonth} style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 10, padding: "10px 16px", fontSize: 16, cursor: "pointer", fontWeight: 700, color: isCurrentMonth ? T.muted : T.text, opacity: isCurrentMonth ? 0.3 : 1 }}>›</button>
          </div>
        )}

        {/* 급여 계산 탭 */}
        {tab === "calc" && members.map(u => {
          const info = memberInfo[u.id] || {};
          const ms = getMonthStats(u.id);
          const key = `${u.id}_${selectedMonth}`;
          const saved = savedWages[key];
          const hourlyWage = Number(info.hourlyWage || 0);

          // 근태 변경 감지 — 확정 후 근태가 바뀌었으면 "재확인 필요"
          const needsRecheck = saved && (() => {
            const cur = ms;
            const sv = saved.monthStats || {};
            return (
              cur.days !== sv.days ||
              cur.lateMin !== sv.lateMin ||
              cur.earlyMin !== sv.earlyMin ||
              cur.otMin !== sv.otMin ||
              cur.holiday !== sv.holiday ||
              cur.absentDays !== sv.absentDays
            );
          })();

          // 상태: "미확정" | "재확인필요" | "확정"
          const statusLabel = !saved ? "미확정" : needsRecheck ? "재확인 필요" : "확정";
          const statusBg    = !saved ? "#f3f4f6" : needsRecheck ? "#fef9c3" : "#dcfce7";
          const statusColor = !saved ? "#6b7280" : needsRecheck ? "#b45309" : "#16a34a";
          const borderColor = !saved ? T.border : needsRecheck ? "#fde68a" : "#16a34a44";
          return (
            <div key={u.id} style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 12, border: `1px solid ${borderColor}`, boxShadow: "0 1px 4px #0000000a" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 12 }}>
                <div style={{ width: 44, height: 44, borderRadius: "50%", background: "#16a34a", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 17, fontWeight: 800, color: "#fff" }}>{u.name[0]}</div>
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 800, fontSize: 16, color: T.text }}>{u.name}</div>
                </div>
                <span style={{ fontSize: 13, fontWeight: 800, padding: "5px 12px", borderRadius: 10, background: statusBg, color: statusColor, whiteSpace: "nowrap" }}>{statusLabel}</span>
              </div>
              {needsRecheck && (
                <div style={{ fontSize: 12, color: "#b45309", background: "#fef9c3", borderRadius: 8, padding: "8px 12px", marginBottom: 10 }}>
                  ⚠ 확정 후 근태가 변경됐어요. 재확인 후 다시 확정해주세요.
                </div>
              )}
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
                  style={{ padding: "10px 0", borderRadius: 10, border: "none", background: !saved ? "#16a34a" : needsRecheck ? "#b45309" : "#e5e7eb", color: !saved ? "#fff" : needsRecheck ? "#fff" : "#6b7280", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                  {!saved ? "급여 계산" : needsRecheck ? "⚠ 재확정" : "✏ 수정"}
                </button>
                {saved && (() => {
                  const alreadySent = payslips.some(p => p.userId === u.id && p.month === selectedMonth);
                  return (
                    <button onClick={() => {
                      if (alreadySent && !window.confirm(`${u.name}님께 이미 전송된 명세서가 있어요.\n재전송할까요?`)) return;
                      sendPayslip(u.id, u.name, saved);
                    }} disabled={sending === u.id}
                      style={{ padding: "10px 0", borderRadius: 10, border: "none",
                        background: sending === u.id ? T.muted : alreadySent ? "#e5e7eb" : "#2563eb",
                        color: sending === u.id ? "#fff" : alreadySent ? "#6b7280" : "#fff",
                        fontSize: 13, fontWeight: 700, cursor: sending === u.id ? "not-allowed" : "pointer",
                        opacity: sending === u.id ? 0.6 : 1 }}>
                      {sending === u.id ? "전송중..." : alreadySent ? "🔄 재전송" : "📤 명세서 전송"}
                    </button>
                  );
                })()}
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
                    <span style={{ fontSize: 14, fontWeight: 800, padding: "5px 14px", borderRadius: 10, background: s ? "#dcfce7" : "#f3f4f6", color: s ? "#16a34a" : "#6b7280", whiteSpace: "nowrap" }}>{s ? "확정" : "미확정"}</span>
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
  const [tab, setTab] = useState("active"); // active | retired
  const [deleteModal, setDeleteModal] = useState(null); // { user }
  const [deleteInput, setDeleteInput] = useState("");
  const [deletePin, setDeletePin] = useState("");
  const [deleteErr, setDeleteErr] = useState("");
  const admin = users.find(u => u.role === "admin");

  const activeMembers = users.filter(u => u.role === "member" && (!u.status || u.status === "active"));
  const retiredMembers = users.filter(u => u.role === "member" && u.status === "retired");

  const saveInfo = async (userId, data) => {
    await setDoc(doc(db, COL_MEMBER_INFO, userId), data);
    setEditInfo(null);
  };

  const handleCompleteDelete = async () => {
    if (deleteInput !== deleteModal.user.name) { setDeleteErr("이름이 일치하지 않습니다."); return; }
    if (deletePin !== admin?.pin && deletePin !== MASTER_CODE) { setDeleteErr("PIN이 맞지 않습니다."); return; }
    try {
      await fbDeleteUserCompletely(deleteModal.user.id);
      setDeleteModal(null); setDeleteInput(""); setDeletePin(""); setDeleteErr("");
    } catch(e) { setDeleteErr("삭제 실패: " + e.message); }
  };

  return (
    <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'Noto Sans KR',sans-serif" }}>
      <div style={{ background: "#7c3aed", paddingTop: "calc(16px + env(safe-area-inset-top))", paddingBottom: 20, paddingLeft: 16, paddingRight: 16 }}>
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
            {tab === "active" && <button onClick={() => setShowUserModal(true)} style={{ background: "#ffffff18", border: "none", color: "#fff", padding: "7px 14px", borderRadius: 18, fontSize: 12, cursor: "pointer", fontWeight: 600 }}>+ 관리</button>}
          </div>
        </div>
        {/* 탭 */}
        <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
          <button onClick={() => setTab("active")}
            style={{ padding: "6px 16px", borderRadius: 20, border: "none", background: tab === "active" ? "#fff" : "#ffffff25", color: tab === "active" ? "#7c3aed" : "#fff", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
            재직중 {activeMembers.length}명
          </button>
          <button onClick={() => setTab("retired")}
            style={{ padding: "6px 16px", borderRadius: 20, border: "none", background: tab === "retired" ? "#fff" : "#ffffff25", color: tab === "retired" ? "#7c3aed" : "#fff", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
            퇴직자 {retiredMembers.length}명
          </button>
        </div>
      </div>

      <div style={{ padding: 16 }}>
        {/* 재직중 탭 */}
        {tab === "active" && activeMembers.map(u => {
          const pending = leaveRequests.filter(r => r.userId === u.id && r.status === "대기").length;
          const info = memberInfo[u.id] || {};
          return (
            <div key={u.id} style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 14, border: `1px solid ${T.border}`, boxShadow: "0 1px 4px #0000000a" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 12 }}>
                <div style={{ width: 44, height: 44, borderRadius: "50%", background: "#7c3aed", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 17, fontWeight: 800, color: "#fff" }}>{u.name[0]}</div>
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 800, fontSize: 16, color: T.text }}>{u.name}</div>
                  <div style={{ fontSize: 12, color: T.muted }}>{info.empNo && `사번 ${info.empNo} · `}{info.employType || ""}</div>
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 6, alignItems: "flex-end" }}>
                  {pending > 0 && <Badge label={`연차신청 ${pending}건`} color="yellow" />}
                  <button onClick={() => setEditInfo({ user: u })}
                    style={{ background: T.bg, border: `1px solid ${T.border}`, color: T.text, borderRadius: 8, padding: "5px 12px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>기초데이터</button>
                </div>
              </div>
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

        {/* 퇴직자 탭 */}
        {tab === "retired" && (
          retiredMembers.length === 0 ? (
            <div style={{ textAlign: "center", padding: 40 }}>
              <div style={{ fontSize: 40, marginBottom: 10 }}>👤</div>
              <div style={{ fontSize: 14, color: T.muted, fontWeight: 600 }}>퇴직자가 없습니다</div>
            </div>
          ) : retiredMembers.map(u => (
            <div key={u.id} style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 14, border: `1px solid ${T.border}`, opacity: 0.85 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                <div style={{ width: 44, height: 44, borderRadius: "50%", background: "#9ca3af", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 17, fontWeight: 800, color: "#fff" }}>{u.name[0]}</div>
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 800, fontSize: 16, color: T.text }}>{u.name}</div>
                  <div style={{ fontSize: 11, color: T.muted, marginTop: 2 }}>
                    퇴직일: {u.retiredAt ? new Date(u.retiredAt).toLocaleDateString("ko-KR") : "-"}
                  </div>
                  {u.retiredAt && (() => {
                    const deleteAvailDate = new Date(u.retiredAt);
                    deleteAvailDate.setFullYear(deleteAvailDate.getFullYear() + 3);
                    const canDelete = new Date() >= deleteAvailDate;
                    return canDelete ? (
                      <div style={{ fontSize: 11, color: "#16a34a", fontWeight: 700, marginTop: 2 }}>✅ 3년 경과 — 삭제 가능</div>
                    ) : (
                      <div style={{ fontSize: 11, color: "#d97706", fontWeight: 600, marginTop: 2 }}>
                        ⏳ 삭제 가능: {deleteAvailDate.toLocaleDateString("ko-KR")}
                      </div>
                    );
                  })()}
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                  <button onClick={async () => { await fbRestoreUser(u); }}
                    style={{ padding: "6px 12px", borderRadius: 8, border: "none", background: "#dcfce7", color: "#16a34a", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
                    복원
                  </button>
                  {(() => {
                    const deleteAvailDate = new Date(u.retiredAt || Date.now());
                    deleteAvailDate.setFullYear(deleteAvailDate.getFullYear() + 3);
                    const canDelete = new Date() >= deleteAvailDate;
                    return (
                      <button onClick={() => {
                        if (!canDelete && !window.confirm(`아직 3년 보관 의무 기간입니다.\n(삭제 가능일: ${deleteAvailDate.toLocaleDateString("ko-KR")})\n\n그래도 삭제하시겠습니까?`)) return;
                        setDeleteModal({ user: u }); setDeleteInput(""); setDeletePin(""); setDeleteErr("");
                      }}
                        style={{ padding: "6px 12px", borderRadius: 8, border: "none", background: canDelete ? "#fee2e2" : "#f3f4f6", color: canDelete ? "#b91c1c" : "#9ca3af", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
                        완전삭제
                      </button>
                    );
                  })()}
                </div>
              </div>
            </div>
          ))
        )}
      </div>

      {/* 완전삭제 모달 */}
      {deleteModal && (
        <div style={{ position: "fixed", inset: 0, background: "#00000066", zIndex: 999, display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
          <div style={{ background: "#fff", borderRadius: 20, padding: 24, width: "100%", maxWidth: 340 }}>
            <div style={{ fontSize: 32, textAlign: "center", marginBottom: 8 }}>⚠️</div>
            <div style={{ fontSize: 15, fontWeight: 800, color: T.text, textAlign: "center", marginBottom: 6 }}>완전 삭제</div>
            <div style={{ fontSize: 12, color: "#dc2626", textAlign: "center", marginBottom: 4, fontWeight: 600 }}>이 작업은 되돌릴 수 없습니다</div>
            <div style={{ fontSize: 12, color: T.muted, textAlign: "center", marginBottom: 16, lineHeight: 1.7 }}>
              모든 데이터가 영구 삭제됩니다.<br />계속하려면 아래를 입력하세요.
            </div>
            <div style={{ fontSize: 12, color: T.muted, marginBottom: 4, fontWeight: 600 }}>팀원 이름 입력</div>
            <input value={deleteInput} onChange={e => { setDeleteInput(e.target.value); setDeleteErr(""); }}
              placeholder={deleteModal.user.name}
              style={{ width: "100%", padding: "10px 12px", borderRadius: 10, border: `1px solid ${deleteErr ? "#fca5a5" : T.border}`, fontSize: 14, marginBottom: 10, boxSizing: "border-box", fontFamily: "inherit" }} />
            <div style={{ fontSize: 12, color: T.muted, marginBottom: 4, fontWeight: 600 }}>관리자 PIN</div>
            <input type="password" value={deletePin} onChange={e => { setDeletePin(e.target.value); setDeleteErr(""); }}
              placeholder="관리자 PIN 입력"
              style={{ width: "100%", padding: "10px 12px", borderRadius: 10, border: `1px solid ${deleteErr ? "#fca5a5" : T.border}`, fontSize: 14, marginBottom: 10, boxSizing: "border-box", fontFamily: "inherit", letterSpacing: 4 }} />
            {deleteErr && <div style={{ fontSize: 12, color: "#dc2626", textAlign: "center", marginBottom: 10, fontWeight: 600 }}>{deleteErr}</div>}
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              <button onClick={() => setDeleteModal(null)}
                style={{ padding: "12px 0", borderRadius: 12, border: `1px solid ${T.border}`, background: T.bg, color: T.text, fontSize: 14, fontWeight: 700, cursor: "pointer" }}>취소</button>
              <button onClick={handleCompleteDelete}
                disabled={deleteInput !== deleteModal.user.name}
                style={{ padding: "12px 0", borderRadius: 12, border: "none", background: deleteInput === deleteModal.user.name ? "#dc2626" : "#e5e7eb", color: deleteInput === deleteModal.user.name ? "#fff" : T.muted, fontSize: 14, fontWeight: 700, cursor: deleteInput === deleteModal.user.name ? "pointer" : "default" }}>
                완전 삭제
              </button>
            </div>
          </div>
        </div>
      )}

      {editInfo && (
        <MemberInfoModal user={editInfo.user} info={memberInfo[editInfo.user.id] || {}}
          onSave={data => saveInfo(editInfo.user.id, data)} onClose={() => setEditInfo(null)} />
      )}

      {showUserModal && <UserManageModal users={activeMembers} onSave={async u => { await fbSaveUsers(u, activeMembers); setShowUserModal(false); }} onClose={() => setShowUserModal(false)} />}
      {showAccount && <AdminAccountModal users={users} onUpdateUsers={onSaveUsers} onClose={() => setShowAccount(false)} />}
    </div>
  );
}

// ── 관리자 일반 섹션 ───────────────────────────────────────────

// ── 일정 (캘린더 + 리마인더) ─────────────────────────────────────
function AdminSchedule({ reminders = [], users = [], settings = {}, scheduleEvents = [] }) {
  const [subTab, setSubTab] = useState("calendar");
  const [calClickDate, setCalClickDate] = useState(null);

  return (
    <div>
      {/* 서브탭 */}
      <div style={{ display: "flex", background: T.card, borderBottom: `1px solid ${T.border}` }}>
        {[["calendar","🗓 캘린더"],["reminder","🔔 리마인더"]].map(([key,label]) => (
          <button key={key} onClick={() => setSubTab(key)}
            style={{ flex: 1, padding: "12px 0", border: "none", background: "none", cursor: "pointer",
              fontWeight: subTab === key ? 800 : 500, fontSize: 13,
              color: subTab === key ? "#7c3aed" : T.muted,
              borderBottom: subTab === key ? "3px solid #7c3aed" : "3px solid transparent",
              fontFamily: "inherit" }}>
            {label}
          </button>
        ))}
      </div>
      {subTab === "calendar" && (
        <ScheduleCalendar
          reminders={reminders}
          settings={settings}
          scheduleEvents={scheduleEvents}
          users={users}
          onSwitchToReminder={(date) => { setCalClickDate(date); setSubTab("reminder"); }}
        />
      )}
      {subTab === "reminder" && (
        <AdminReminder
          reminders={reminders}
          users={users}
          presetDate={calClickDate}
          onClearPreset={() => setCalClickDate(null)}
        />
      )}
    </div>
  );
}

// ── 일정 캘린더 뷰 ───────────────────────────────────────────────
function ScheduleCalendar({ reminders = [], settings = {}, scheduleEvents = [], users = [], onSwitchToReminder }) {
  const kstNow = new Date(new Date().getTime() + 9 * 60 * 60 * 1000);
  const currentMonth = kstNow.toISOString().slice(0, 7);
  const [selectedMonth, setSelectedMonth] = useState(currentMonth);
  const [selDate, setSelDate] = useState(null);

  // 이벤트 편집 상태
  const [editMode, setEditMode] = useState(null); // null | "add" | "edit"
  const [editTarget, setEditTarget] = useState(null); // 수정 대상 event
  const [evTitle, setEvTitle] = useState("");
  const [evColor, setEvColor] = useState("#7c3aed");
  const [evNote, setEvNote]   = useState("");
  const [evTarget, setEvTarget] = useState("admin"); // "admin"|"all"|팀원ID
  const [saving, setSaving]   = useState(false);

  const [y, m] = selectedMonth.split("-").map(Number);
  const holidays = settings.holidays || [];
  const pad = (n) => String(n).padStart(2, "0");
  const getDateStr = (d) => `${y}-${pad(m)}-${pad(d)}`;
  const firstDay   = new Date(y, m - 1, 1).getDay();
  const daysInMonth = new Date(y, m, 0).getDate();
  const today = kstNow.toISOString().slice(0, 10);

  const prevMonth = () => {
    const d = new Date(y, m - 2, 1);
    setSelectedMonth(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`);
    setSelDate(null); setEditMode(null);
  };
  const nextMonth = () => {
    const d = new Date(y, m, 1);
    setSelectedMonth(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`);
    setSelDate(null); setEditMode(null);
  };

  // 날짜별 이벤트 맵
  const eventMap = {};
  // 1) 공휴일
  holidays.forEach(h => {
    const date = typeof h === "string" ? h : h.date;
    const memo = typeof h === "object" ? h.memo : null;
    if (date?.startsWith(selectedMonth)) {
      if (!eventMap[date]) eventMap[date] = [];
      eventMap[date].push({ type: "holiday", label: memo || "공휴일", color: "#dc2626", id: null });
    }
  });
  // 2) 리마인더 (반복 규칙)
  reminders.forEach(r => {
    if (!r.active) return;
    getOccurrencesInMonth(r, y, m, daysInMonth, holidays).forEach(d => {
      if (!eventMap[d]) eventMap[d] = [];
      eventMap[d].push({ type: "reminder", label: r.title, color: "#7c3aed", id: r.id });
    });
  });
  // 3) 단건 일정 이벤트 — 관리자는 isAdminPost 또는 target=admin 것만
  scheduleEvents
    .filter(ev => ev.isAdminPost || ev.target === "admin" || !ev.target)
    .forEach(ev => {
    if (ev.date?.startsWith(selectedMonth)) {
      if (!eventMap[ev.date]) eventMap[ev.date] = [];
      eventMap[ev.date].push({ type: "event", label: ev.title, color: ev.color || "#0891b2", id: ev.id, note: ev.note, target: ev.target || ev.userId || "admin" });
    }
  });

  // 주간 그리드
  const weeks = [];
  let days = [];
  for (let i = 0; i < firstDay; i++) days.push(null);
  for (let d = 1; d <= daysInMonth; d++) {
    days.push(d);
    if (days.length === 7) { weeks.push(days); days = []; }
  }
  while (days.length < 7) days.push(null);
  if (days.some(d => d)) weeks.push(days);

  const selEvents = selDate ? (eventMap[selDate] || []) : [];

  const openAdd = () => {
    setEvTitle(""); setEvColor("#0891b2"); setEvNote(""); setEvTarget("admin");
    setEditTarget(null); setEditMode("add");
  };
  const openEdit = (ev) => {
    if (ev.type !== "event") return;
    setEvTitle(ev.label); setEvColor(ev.color || "#0891b2"); setEvNote(ev.note || "");
    setEvTarget(ev.target || "admin");
    setEditTarget(ev); setEditMode("edit");
  };
  const cancelEdit = () => { setEditMode(null); setEditTarget(null); };

  const saveEvent = async () => {
    if (!evTitle.trim() || !selDate) return;
    setSaving(true);
    try {
      if (editMode === "add") {
        await addDoc(collection(db, COL_EVENTS), {
          date: selDate, title: evTitle.trim(), color: evColor, note: evNote.trim(),
          target: evTarget, isAdminPost: true, createdAt: new Date().toISOString()
        });
      } else if (editMode === "edit" && editTarget?.id) {
        await setDoc(doc(db, COL_EVENTS, editTarget.id), {
          date: selDate, title: evTitle.trim(), color: evColor, note: evNote.trim(),
          target: evTarget, isAdminPost: true, createdAt: editTarget.createdAt || new Date().toISOString()
        });
      }
      cancelEdit();
    } catch(e) { console.error(e); }
    setSaving(false);
  };

  const deleteEvent = async (id) => {
    if (!id) return;
    await deleteDoc(doc(db, COL_EVENTS, id));
  };

  const EVENT_COLORS = [
    ["#0891b2","청록"],["#7c3aed","보라"],["#16a34a","초록"],
    ["#d97706","주황"],["#dc2626","빨강"],["#1d4ed8","파랑"],
  ];

  return (
    <div style={{ padding: 16, paddingBottom: 140 }}>
      {/* 월 이동 */}
      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 14 }}>
        <button onClick={prevMonth}
          style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 10, padding: "9px 14px", fontSize: 16, cursor: "pointer", fontWeight: 700, color: T.text }}>‹</button>
        <div style={{ flex: 1, textAlign: "center", fontSize: 16, fontWeight: 800, color: T.text }}>{y}년 {m}월</div>
        <button onClick={nextMonth}
          style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 10, padding: "9px 14px", fontSize: 16, cursor: "pointer", fontWeight: 700, color: T.text }}>›</button>
      </div>

      {/* 캘린더 그리드 — 날짜칸 고정 크기, 내용은 도트로만 */}
      <div style={{ background: T.card, borderRadius: 14, border: `1px solid ${T.border}`, overflow: "hidden", marginBottom: 12 }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(7,1fr)" }}>
          {["일","월","화","수","목","금","토"].map((d, i) => (
            <div key={d} style={{ textAlign: "center", padding: "7px 0", fontSize: 10, fontWeight: 700,
              color: i === 0 ? "#dc2626" : i === 6 ? "#2563eb" : T.muted,
              borderBottom: `1px solid ${T.border}` }}>{d}</div>
          ))}
        </div>
        {weeks.map((week, wi) => (
          <div key={wi} style={{ display: "grid", gridTemplateColumns: "repeat(7,1fr)" }}>
            {week.map((d, di) => {
              const dateStr = d ? getDateStr(d) : null;
              const events  = d ? (eventMap[dateStr] || []) : [];
              const isToday    = dateStr === today;
              const isSelected = !!d && dateStr === selDate;
              const isHol = d ? isHoliday(dateStr, holidays) : false;
              const hasHol = events.some(e => e.type === "holiday");
              let dateColor = T.muted;
              if (di === 6) dateColor = "#2563eb";                          // 토요일 → 파랑 (최우선)
              else if (di === 0 || isHol || hasHol) dateColor = "#dc2626";  // 일요일·공휴일 → 빨강
              return (
                <div key={di}
                  onClick={() => { if (!d) return; setSelDate(selDate === dateStr ? null : dateStr); setEditMode(null); setEditTarget(null); }}
                  style={{
                    /* 고정 높이 — 내용 많아도 칸 크기 불변 */
                    height: 80, boxSizing: "border-box",
                    padding: "4px 3px 3px",
                    borderBottom: wi < weeks.length - 1 ? `1px solid ${T.border}` : "none",
                    borderRight: di < 6 ? `1px solid ${T.border}` : "none",
                    display: "flex", flexDirection: "column", alignItems: "stretch", gap: 1,
                    cursor: d ? "pointer" : "default",
                    background: isSelected ? "#ede9fe" : isToday ? "#f0f9ff" : "transparent",
                    overflow: "hidden", transition: "background 0.15s"
                  }}>
                  {d && <>
                    {/* 날짜 숫자 */}
                    <div style={{
                      fontSize: 12, fontWeight: isToday ? 900 : 600,
                      color: di === 6 ? "#2563eb" : (di === 0 || isHol || hasHol) ? "#dc2626" : isToday ? "#7c3aed" : dateColor,
                      width: 20, height: 20, borderRadius: "50%",
                      display: "flex", alignItems: "center", justifyContent: "center",
                      background: isToday ? (di === 6 ? "#dbeafe" : (di === 0 || isHol || hasHol) ? "#fee2e2" : "#ede9fe") : "transparent",
                      flexShrink: 0, alignSelf: "center"
                    }}>{d}</div>
                    <CalEventLabels events={events} />
                  </>}
                </div>
              );
            })}
          </div>
        ))}
      </div>

      {/* 범례 */}
      <div style={{ display: "flex", gap: 10, marginBottom: 14, justifyContent: "flex-end", flexWrap: "wrap" }}>
        {[["#dc2626","공휴일"],["#7c3aed","리마인더"],["#0891b2","일정"]].map(([color,label]) => (
          <div key={label} style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 10, fontWeight: 700 }}>
            <div style={{ width: 7, height: 7, borderRadius: "50%", background: color }} />
            <span style={{ color: T.muted }}>{label}</span>
          </div>
        ))}
      </div>

      {/* 선택 날짜 패널 */}
      {selDate && (
        <div style={{ background: T.card, borderRadius: 14, border: `1px solid ${T.border}`, overflow: "hidden", marginBottom: 12 }}>
          {/* 패널 헤더 */}
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between",
            padding: "12px 14px", borderBottom: `1px solid ${T.border}`, background: "#faf5ff" }}>
            <div style={{ fontSize: 14, fontWeight: 800, color: T.text }}>{formatDate(selDate)}</div>
            <button onClick={openAdd}
              style={{ background: "#7c3aed", border: "none", color: "#fff", borderRadius: 10,
                padding: "6px 14px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
              + 일정 추가
            </button>
          </div>

          {/* 이벤트 목록 */}
          <div style={{ padding: "8px 14px" }}>
            {selEvents.length === 0 && editMode !== "add" ? (
              <div style={{ fontSize: 13, color: T.muted, textAlign: "center", padding: "16px 0" }}>등록된 일정 없음</div>
            ) : selEvents.map((ev, i) => (
              <div key={i}>
                {/* 이벤트 행 */}
                {editMode === "edit" && editTarget?.id === ev.id ? (
                  <EventForm
                    evTitle={evTitle} setEvTitle={setEvTitle}
                    evColor={evColor} setEvColor={setEvColor}
                    evNote={evNote}  setEvNote={setEvNote}
                    evTarget={evTarget} setEvTarget={setEvTarget} users={users}
                    EVENT_COLORS={EVENT_COLORS}
                    saving={saving} onSave={saveEvent} onCancel={cancelEdit}
                    label="수정"
                  />
                ) : (
                  <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "9px 0",
                    borderBottom: i < selEvents.length - 1 ? `1px solid ${T.border}` : "none" }}>
                    <div style={{ width: 8, height: 8, borderRadius: "50%", background: ev.color, flexShrink: 0 }} />
                    <div style={{ flex: 1 }}>
                      <div style={{ fontSize: 13, fontWeight: 700,
                        color: ev.type === "holiday" ? "#dc2626" : T.text }}>{ev.label}</div>
                      {ev.note && <div style={{ fontSize: 11, color: T.muted, marginTop: 2 }}>{ev.note}</div>}
                    </div>
                    {ev.type === "holiday" && <span style={{ fontSize: 10, color: "#dc2626", fontWeight: 700 }}>공휴일</span>}
                    {ev.type === "reminder" && (
                      <span style={{ fontSize: 10, color: "#7c3aed", fontWeight: 700 }}>🔔</span>
                    )}
                    {ev.type === "event" && (
                      <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 3 }}>
                        <span style={{ fontSize: 9, color: T.muted, fontWeight: 600 }}>
                          {ev.target === "admin" ? "나만" : ev.target === "all" ? "전체" : users.find(u=>u.id===ev.target)?.name || ev.target}
                        </span>
                        <div style={{ display: "flex", gap: 4 }}>
                          <button onClick={() => openEdit(ev)}
                            style={{ background: T.bg, border: `1px solid ${T.border}`, color: T.muted,
                              borderRadius: 7, padding: "3px 9px", fontSize: 11, cursor: "pointer", fontWeight: 600 }}>수정</button>
                          <button onClick={() => { if (window.confirm("일정을 삭제할까요?")) deleteEvent(ev.id); }}
                            style={{ background: T.redBg, border: "none", color: T.red,
                              borderRadius: 7, padding: "3px 9px", fontSize: 11, cursor: "pointer", fontWeight: 700 }}>삭제</button>
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </div>
            ))}

            {/* 추가 폼 */}
            {editMode === "add" && (
              <div style={{ borderTop: selEvents.length > 0 ? `1px solid ${T.border}` : "none", paddingTop: selEvents.length > 0 ? 10 : 0 }}>
                <EventForm
                  evTitle={evTitle} setEvTitle={setEvTitle}
                  evColor={evColor} setEvColor={setEvColor}
                  evNote={evNote}  setEvNote={setEvNote}
                  evTarget={evTarget} setEvTarget={setEvTarget} users={users}
                  EVENT_COLORS={EVENT_COLORS}
                  saving={saving} onSave={saveEvent} onCancel={cancelEdit}
                  label="추가"
                />
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

// ── 이벤트 입력 폼 (인라인) ───────────────────────────────────────
function EventForm({ evTitle, setEvTitle, evColor, setEvColor, evNote, setEvNote, evTarget, setEvTarget, users, EVENT_COLORS, saving, onSave, onCancel, label }) {
  const iS = { width: "100%", padding: "9px 12px", borderRadius: 10, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 13, fontWeight: 600, boxSizing: "border-box", fontFamily: "inherit" };
  return (
    <div style={{ background: "#f5f3ff", borderRadius: 10, padding: 12, marginTop: 4 }}>
      <input value={evTitle} onChange={e => setEvTitle(e.target.value)}
        placeholder="일정 제목" autoFocus style={{ ...iS, marginBottom: 8 }} />
      <input value={evNote} onChange={e => setEvNote(e.target.value)}
        placeholder="메모 (선택)" style={{ ...iS, marginBottom: 8 }} />
      {/* 수신인 선택 — 관리자 캘린더에서만 표시 */}
      {setEvTarget && users && (
        <select value={evTarget} onChange={e => setEvTarget(e.target.value)}
          style={{ ...iS, marginBottom: 8 }}>
          <option value="admin">나만 (관리자 개인)</option>
          <option value="all">전체 팀원</option>
          {users.filter(u => u.role !== "admin").map(u => (
            <option key={u.id} value={u.id}>{u.name}</option>
          ))}
        </select>
      )}
      {/* 색상 선택 */}
      <div style={{ display: "flex", gap: 6, marginBottom: 10, alignItems: "center" }}>
        <span style={{ fontSize: 11, color: T.muted, fontWeight: 600, whiteSpace: "nowrap" }}>색상</span>
        {EVENT_COLORS.map(([c, name]) => (
          <button key={c} onClick={() => setEvColor(c)} title={name}
            style={{ width: 22, height: 22, borderRadius: "50%", background: c, border: evColor === c ? "3px solid #1a1a2e" : "2px solid #fff",
              cursor: "pointer", flexShrink: 0, boxShadow: evColor === c ? "0 0 0 1px " + c : "none" }} />
        ))}
      </div>
      <div style={{ display: "flex", gap: 8 }}>
        <button onClick={onCancel}
          style={{ flex: 1, background: "#fff", border: `1px solid ${T.border}`, color: T.muted,
            borderRadius: 8, padding: "8px 0", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>취소</button>
        <button onClick={onSave} disabled={saving || !evTitle.trim()}
          style={{ flex: 2, background: saving ? T.muted : "#7c3aed", border: "none", color: "#fff",
            borderRadius: 8, padding: "8px 0", fontSize: 12, fontWeight: 700,
            cursor: saving ? "not-allowed" : "pointer", opacity: !evTitle.trim() ? 0.5 : 1 }}>
          {saving ? "저장 중..." : label}
        </button>
      </div>
    </div>
  );
}

// 리마인더가 해당 월에 발생하는 날짜 목록 반환
function getOccurrencesInMonth(r, y, m, daysInMonth, holidays) {
  const pad = (n) => String(n).padStart(2, "0");
  const results = [];

  // 근무일 목록
  const workDays = [];
  for (let d = 1; d <= daysInMonth; d++) {
    const dateStr = `${y}-${pad(m)}-${pad(d)}`;
    const dow = new Date(y, m - 1, d).getDay();
    if (dow === 0 || dow === 6) continue;
    if (isHoliday(dateStr, holidays)) continue;
    workDays.push(d);
  }
  const lastWorkDay = workDays[workDays.length - 1] || null;

  for (let d = 1; d <= daysInMonth; d++) {
    const dateStr = `${y}-${pad(m)}-${pad(d)}`;
    const dow = new Date(y, m - 1, d).getDay();
    let matches = false;
    if (r.repeat === "daily") matches = true;
    else if (r.repeat === "weekly") matches = (dow === (r.weekDay ?? 1));
    else if (r.repeat === "monthly") matches = (d === (r.monthDay || 1));
    else if (r.repeat === "monthly_nth_work") {
      const nth = r.monthWorkDay ?? 1;
      if (workDays.length >= nth && workDays[nth - 1] === d) matches = true;
    }
    else if (r.repeat === "monthly_last_work") matches = (d === lastWorkDay);
    if (matches) results.push(dateStr);
  }
  return results;
}

// ── 리마인더 ────────────────────────────────────────────────────
function AdminReminder({ reminders = [], users = [], presetDate = null, onClearPreset }) {
  const EMPTY = { title: "", time: "09:00", repeat: "daily", monthDay: 1, weekDay: 1, monthWorkDay: 1, target: "admin", sendBeforeHoliday: false };
  const [form, setForm] = useState(EMPTY);
  const [editId, setEditId] = useState(null); // null=추가, id=수정
  const [adding, setAdding] = useState(false);
  const [loading, setLoading] = useState(false);

  // 캘린더에서 날짜 클릭 시 자동으로 추가 폼 열기
  useEffect(() => {
    if (presetDate) {
      const [, , dd] = presetDate.split("-").map(Number);
      setForm(p => ({ ...p, repeat: "monthly", monthDay: dd }));
      setAdding(true);
      setEditId(null);
      if (onClearPreset) onClearPreset();
    }
  }, [presetDate]);

  const DOW = ["일","월","화","수","목","금","토"];

  const openEdit = (r) => {
    setForm({ title: r.title, time: r.time, repeat: r.repeat, monthDay: r.monthDay || 1, weekDay: r.weekDay || 1, monthWeek: r.monthWeek || 1, monthWorkDay: r.monthWorkDay || 1, target: r.target || "admin", sendBeforeHoliday: r.sendBeforeHoliday || false });
    setEditId(r.id);
    setAdding(false);
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
    const DOW_LABEL = ["일","월","화","수","목","금","토"];
    if (r.repeat === "daily") return "매일";
    if (r.repeat === "weekly") return `매주 ${DOW_LABEL[r.weekDay]}요일`;
    if (r.repeat === "monthly") return `매월 ${r.monthDay}일`;
    if (r.repeat === "monthly_nth_work") {
      return r.monthWorkDay === 0 ? "매월 마지막 근무일" : `매월 ${r.monthWorkDay}번째 근무일`;
    }
    if (r.repeat === "monthly_last_work") return "매월 마지막 근무일";
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
                <option value="monthly">매월 (날짜)</option>
                <option value="monthly_nth_work">매월 (N번째 근무일)</option>
                <option value="monthly_last_work">매월 마지막 근무일</option>
              </select>
            </div>
          </div>

          {/* 매주 → 요일 선택 */}
          {(form.repeat === "weekly") && (
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

          {/* 매월 날짜 */}
          {form.repeat === "monthly" && (
            <div style={{ marginBottom: 10 }}>
              <div style={{ fontSize: 11, color: T.muted, marginBottom: 4 }}>날짜</div>
              <select value={form.monthDay} onChange={e => setForm(p => ({...p, monthDay: Number(e.target.value)}))}
                style={{ width: "100%", padding: "10px 8px", borderRadius: 10, border: `1px solid ${T.border}`, fontSize: 13, fontWeight: 700, background: "#fff", color: T.text }}>
                {Array.from({length: 31}, (_, i) => i+1).map(d => <option key={d} value={d}>{d}일</option>)}
              </select>
            </div>
          )}

          {/* 매월 N번째 근무일 */}
          {form.repeat === "monthly_nth_work" && (
            <div style={{ marginBottom: 10 }}>
              <div style={{ fontSize: 11, color: T.muted, marginBottom: 6 }}>몇 번째 근무일</div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                {Array.from({length: 20}, (_, i) => i + 1).map(n => (
                  <button key={n} onClick={() => setForm(p => ({...p, monthWorkDay: n}))}
                    style={{ width: 40, height: 36, borderRadius: 8,
                      border: `1px solid ${form.monthWorkDay === n ? "#7c3aed" : T.border}`,
                      background: form.monthWorkDay === n ? "#7c3aed" : "#fff",
                      color: form.monthWorkDay === n ? "#fff" : T.text,
                      fontSize: 12, fontWeight: 700, cursor: "pointer" }}>{n}</button>
                ))}
              </div>
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
      ) : [...reminders].sort((a, b) => {
          // 매월→매주→매일 순, 같은 타입은 날짜/요일 순
          const typeOrder = { monthly: 0, monthly_nth_work: 1, monthly_last_work: 2, weekly: 3, daily: 4 };
          const ta = typeOrder[a.repeat] ?? 9, tb = typeOrder[b.repeat] ?? 9;
          if (ta !== tb) return ta - tb;
          if (a.repeat === "monthly") return (a.monthDay || 1) - (b.monthDay || 1);
          if (a.repeat === "weekly") return (a.weekDay ?? 0) - (b.weekDay ?? 0);
          return 0;
        }).map(r => {
          // 날짜 동그라미 라벨
          const DOW = ["일","월","화","수","목","금","토"];
          const badgeLabel = r.repeat === "monthly" ? `${r.monthDay || 1}`
            : r.repeat === "weekly" ? DOW[r.weekDay ?? 1]
            : r.repeat === "monthly_nth_work" ? (r.monthWorkDay === 0 ? "말" : `${r.monthWorkDay}근`)
            : r.repeat === "monthly_last_work" ? "말"
            : "매";
        return (
        <div key={r.id} style={{ background: T.card, border: `1px solid ${editId === r.id ? "#7c3aed" : r.active ? "#7c3aed44" : T.border}`, borderRadius: 14, marginBottom: 10, overflow: "hidden", opacity: r.active ? 1 : 0.6 }}>
          {/* 요약 행 */}
          <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "12px 14px" }}>
            {/* 날짜 동그라미 */}
            <div style={{ width: 36, height: 36, borderRadius: "50%", background: r.active ? "#7c3aed" : T.muted, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
              <span style={{ fontSize: r.repeat === "monthly" ? 11 : 13, fontWeight: 800, color: "#fff" }}>{badgeLabel}</span>
            </div>
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
              <button onClick={() => { if (window.confirm(`"${r.title}" 리마인더를 삭제할까요?`)) deleteReminder(r.id); }}
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
                  <option value="monthly">매월 (날짜)</option>
                  <option value="monthly_nth_work">매월 (N번째 근무일)</option>
                  <option value="monthly_last_work">매월 마지막 근무일</option>
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
              {form.repeat === "monthly_nth_work" && (
                <div style={{ marginBottom: 8 }}>
                  <div style={{ fontSize: 11, color: T.muted, marginBottom: 6 }}>몇 번째 근무일</div>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 4 }}>
                    {Array.from({length: 20}, (_, i) => i + 1).map(n => (
                      <button key={n} onClick={() => setForm(p => ({...p, monthWorkDay: n}))}
                        style={{ width: 36, height: 32, borderRadius: 8,
                          border: `1px solid ${form.monthWorkDay === n ? "#7c3aed" : T.border}`,
                          background: form.monthWorkDay === n ? "#7c3aed" : "#fff",
                          color: form.monthWorkDay === n ? "#fff" : T.text,
                          fontSize: 11, fontWeight: 700, cursor: "pointer" }}>{n}</button>
                    ))}
                  </div>
                </div>
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
      )})}
    </div>
  );
}

// ── 관리자 섹션 래퍼 (공지/게시판용) ─────────────────────────────
function AdminSectionWrap({ title, color, onBack, children }) {
  return (
    <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'Noto Sans KR',sans-serif" }}>
      <div style={{ background: color || T.adminHeader, paddingTop: "calc(16px + env(safe-area-inset-top))", paddingBottom: "14px", paddingLeft: "16px", paddingRight: "16px" }}>
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

// ── 교육 관리 (관리자) ─────────────────────────────────────────
const CONTRACT_ISTYLE = { width: "100%", padding: "10px 12px", borderRadius: 10, border: `1px solid ${T.border}`, fontSize: 13, fontWeight: 600, color: T.text, background: "#fff", boxSizing: "border-box", fontFamily: "inherit" };
function ContractLabel({ children }) { return <div style={{ fontSize: 12, color: T.sub, fontWeight: 600, marginBottom: 4 }}>{children}</div>; }
function ContractField({ label, fkey, type = "text", placeholder = "", form, setForm }) {
  return (
    <div style={{ marginBottom: 12 }}>
      <ContractLabel>{label}</ContractLabel>
      <input type={type} value={form[fkey] || ""} onChange={e => setForm(p => ({ ...p, [fkey]: e.target.value }))}
        placeholder={placeholder} style={CONTRACT_ISTYLE} />
    </div>
  );
}

// ── 근로계약서 (관리자) ─────────────────────────────────────────
function DocSection({ users, memberInfo, settings, contracts, onBack }) {
  const members = users.filter(u => u.role === "member" && (!u.status || u.status === "active"));
  const [docTypeFilter, setDocTypeFilter] = useState("contract"); // 상단 탭
  const [selUser, setSelUser] = useState(null);
  const [editing, setEditing] = useState(false);
  const [saving, setSaving] = useState(false);
  const [pinModal, setPinModal] = useState(null); // { contract } 또는 null
  const [pinInput, setPinInput] = useState("");
  const [pinErr, setPinErr] = useState("");
  const [pdfLoading, setPdfLoading] = useState(null);
  const admin = users.find(u => u.role === "admin");
  const [form, setForm] = useState(null);

  const downloadContractPDF = async (contract) => {
    setPdfLoading(contract.id);
    try {
      const el = document.getElementById(`contract-print-${contract.id}`);
      if (!el) { alert("PDF 생성 영역을 찾을 수 없습니다."); setPdfLoading(null); return; }
      const pdfBtn = document.getElementById(`contract-pdf-btn-${contract.id}`);
      if (pdfBtn) pdfBtn.style.visibility = "hidden";
      const canvas = await html2canvas(el, { scale: 2.5, useCORS: true, backgroundColor: "#ffffff" });
      if (pdfBtn) pdfBtn.style.visibility = "visible";
      const imgData = canvas.toDataURL("image/png");
      const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
      const pageW = pdf.internal.pageSize.getWidth();
      const pageH = pdf.internal.pageSize.getHeight();
      const margin = 8;
      const maxW = pageW - margin * 2;
      const maxH = pageH - margin * 2;
      const ratio = canvas.width / canvas.height;
      let imgW = maxW, imgH = imgW / ratio;
      if (imgH > maxH) { imgH = maxH; imgW = imgH * ratio; }
      pdf.addImage(imgData, "PNG", (pageW - imgW) / 2, margin, imgW, imgH);
      const docLabel = DOC_TYPES.find(d => d.key === (contract.docType || "contract"))?.label || "문서";
      const docName = (!contract.docType || contract.docType === "contract")
        ? `근로계약서_${contract.contractStart || ""}`
        : `${contract.docTitle || docLabel}_${contract.createdAt?.slice(0,10) || ""}`;
      const fileName = `${contract.userName}_${docName}.pdf`;
      pdf.save(fileName);
      // ✅ 다운로드 완료 즉시 버튼 정상화
      setPdfLoading(null);
      // 🔄 Storage 업로드는 백그라운드로 처리
      try {
        const pdfBlob = pdf.output("blob");
        const storageRef = ref(storage, `contracts/${contract.userId}/${contract.id}.pdf`);
        await uploadBytes(storageRef, pdfBlob);
        const url = await getDownloadURL(storageRef);
        const cacheBustUrl = `${url}&t=${Date.now()}`;
        await setDoc(doc(db, COL_CONTRACTS, contract.id), { ...contract, pdfUrl: cacheBustUrl, pdfGeneratedAt: new Date().toISOString() });
      } catch(e) { console.warn("Storage 저장 실패:", e.message); }
    } catch(e) {
      alert("PDF 생성 실패: " + e.message);
      setPdfLoading(null);
    }
  };

  const openNew = (u, dtype = docTypeFilter) => {
    const info = memberInfo[u.id] || {};
    const now = new Date();
    const thisYear = now.getFullYear();
    const janFirst = `${thisYear}-01-01`;
    const hourly = Number(info.hourlyWage || 0);
    const wkHours = Number(info.weeklyHours || 40);
    const moHours = Math.round(wkHours / 40 * 209);
    const calcMonthly = hourly ? String(Math.round(hourly * moHours)) : "";
    setForm({
      userId: u.id,
      userName: u.name,
      companyName: settings.companyName || "하나기업",
      ownerName: settings.ownerName || "박용균",
      bizAddress: settings.bizAddress || "경남 양산시 어곡공단2길 28",
      empAddress: "",
      empPhone: "",
      workPlace: "경남 양산시 어곡공단2길 28",
      jobType: info.jobType || "포장직",
      joinDate: info.joinDate || "",
      contractStart: janFirst,
      contractEnd: "",
      workStart: settings.workStart || "07:30",
      workEnd: settings.workEnd || "16:30",
      dailyHours: "8",
      weekDays: "월~금",
      breakLunch: "60분(11:40~12:40)",
      breakSnack: "20분(16:10~16:30)",
      payType: "월급제",
      monthlyWage: calcMonthly,
      hourlyWage: String(info.hourlyWage || ""),
      weeklyHours: String(wkHours),
      monthlyHours: String(moHours),
      wage1: "시급*근로일수 (주휴수당 포함)",
      wage2: "시급*8시간 / 별도 지급",
      wage3: "시급*근로시간*1.5 / 별도 지급",
      wage4: "시급*근로시간*1.5 / 별도 지급",
      wage5: "기본급의 100% / 수시 지급",
      payCalcPeriod: "전월 1일부터 전월 말일까지",
      payDay: String(settings.payDay || "15"),
      payHoliday: "공휴일은 익일 지급",
      payMethod: "계좌이체",
      bankName: info.bank || "",
      bankAccount: info.account || "",
      annualLeave: "근로기준법에서 정하는 바에 따라 부여함",
      insurance: "4대보험 의무가입",
      welfare: "교통비, 식대 지원",
      severancePay: "1년 이상 근무하고 퇴직하였을 때는 1년에 대하여 평균임금 1개월분의 퇴직금을 지급한다",
      resignNotice: "퇴사하기 30일전에 통보한다",
      retirementAge: "만 60세가 되는 해, 년도말일 기준으로 정년퇴임 한다",
      terminationReasons: "1. 정당한 업무명령을 위반하였을 때\n2. 무단결근 계속 3일 이상(지각, 조퇴 3회는 결근 1일로 간주)\n3. 근로계약기간이 종료되었을 때\n4. \"갑\"의 발주처로부터 \"을\"의 귀책사유로 교체요청이 있을 때\n5. \"갑\"의 발주처와의 도급계약의 해지가 있을 때 또는 계약의 연장이 안되었을 때\n6. 상기 해지 사유 발생 시 30일전에 통보한다",
      bonus: "설 50%, 추석 50% / 지급시기 변동시 통보 후 변경가능",
      specialTerms: "",
      docType: dtype,
      status: "draft",
      signedAt: null,
      createdAt: now.toISOString(),
    });
    setSelUser(u);
    setEditing(true);
  };

  const openEdit = (contract) => {
    setForm({ ...contract });
    setSelUser(users.find(u => u.id === contract.userId) || { id: contract.userId, name: contract.userName });
    setEditing(true);
  };

  const saveContract = async () => {
    if (!form) return;
    setSaving(true);
    try {
      const ref = await addDoc(collection(db, COL_CONTRACTS), { ...form, id: "" });
      await setDoc(doc(db, COL_CONTRACTS, ref.id), { ...form, id: ref.id });
      setEditing(false);
      setForm(null);
    } catch(e) { alert("저장 실패: " + e.message); }
    setSaving(false);
  };

  const sendToEmployee = async (contract) => {
    const docLabel = DOC_TYPES.find(d => d.key === (contract.docType || "contract"))?.label || "문서";
    const docTitle = contract.docTitle || docLabel;
    if (!window.confirm(`${contract.userName}님께 ${docLabel} 서명 요청을 보낼까요?`)) return;
    try {
      await setDoc(doc(db, COL_CONTRACTS, contract.id), { ...contract, status: "sent", sentAt: new Date().toISOString() });
      await addDoc(collection(db, COL_NOTICES), {
        title: `📄 ${docTitle} 서명 요청`,
        content: `${docTitle}이(가) 발송되었습니다.\n문서함 탭에서 내용을 확인하고 동의(서명)해주세요.`,
        recipient: contract.userId,
        author: "관리자",
        createdAt: new Date().toISOString(), auto: true,
      });
      alert(`${contract.userName}님께 ${docLabel} 서명 요청을 보냈습니다.`);
    } catch(e) { alert("발송 실패: " + e.message); }
  };

  const deleteContract = async (contract) => {
    const docLabel = DOC_TYPES.find(d => d.key === (contract.docType || "contract"))?.label || "문서";
    const docName = contract.docTitle || docLabel;
    if (contract.status === "signed") {
      setPinInput("");
      setPinErr("");
      setPinModal({ contract });
      return;
    }
    if (!window.confirm(`${contract.userName}님의 ${docName}을(를) 삭제할까요?`)) return;
    try { await deleteDoc(doc(db, COL_CONTRACTS, contract.id)); } catch(e) { alert("삭제 실패: " + e.message); }
  };

  const confirmPinDelete = async () => {
    if (pinInput !== admin?.pin && pinInput !== MASTER_CODE) {
      setPinErr("PIN이 맞지 않습니다.");
      return;
    }
    if (!pinModal?.contract?.id) {
      setPinErr("계약서 ID를 찾을 수 없습니다.");
      return;
    }
    try {
      await deleteDoc(doc(db, COL_CONTRACTS, pinModal.contract.id));
      setPinModal(null);
    } catch(e) {
      console.error("계약서 삭제 실패:", e);
      setPinErr("삭제 실패: " + e.message);
    }
  };

  const statusLabel = (s) => s === "signed" ? { text: "✅ 서명완료", color: "#16a34a", bg: "#dcfce7" } : s === "sent" ? { text: "📨 서명대기", color: "#d97706", bg: "#fef3c7" } : { text: "📝 초안", color: "#6b7280", bg: "#f3f4f6" };
  const [expandedHistory, setExpandedHistory] = useState({});

  // 편집 화면
  if (editing && form) {
    const docTypeInfo = DOC_TYPES.find(d => d.key === (form.docType || "contract")) || DOC_TYPES[0];
    return (
      <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'Noto Sans KR',sans-serif", paddingBottom: 90 }}>
        <div style={{ background: T.adminHeader, paddingTop: "calc(16px + env(safe-area-inset-top))", paddingBottom: "14px", paddingLeft: "16px", paddingRight: "16px" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <button onClick={() => { setEditing(false); setForm(null); }} style={{ background: "#ffffff18", border: "none", color: "#fff", fontSize: 18, cursor: "pointer", padding: "8px 14px", borderRadius: 12, fontWeight: 700 }}>‹</button>
            <div>
              <div style={{ fontSize: 11, color: "#ffffff40", letterSpacing: 3 }}>ADMIN</div>
              <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>{docTypeInfo.icon} {selUser?.name} {docTypeInfo.label}</div>
            </div>
          </div>
        </div>

        <div style={{ padding: "16px 16px 0" }}>
          {/* 문서 종류 선택 */}
          <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 12, border: `1px solid ${T.border}` }}>
            <div style={{ fontSize: 13, fontWeight: 800, color: T.text, marginBottom: 10 }}>📌 문서 종류</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
              {DOC_TYPES.map(dt => (
                <button key={dt.key} onClick={() => setForm(p => ({ ...p, docType: dt.key }))}
                  style={{ padding: "10px 0", borderRadius: 10, border: `2px solid ${form.docType === dt.key ? T.adminHeader : T.border}`, background: form.docType === dt.key ? T.adminHeader : T.bg, color: form.docType === dt.key ? "#fff" : T.text, fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                  {dt.icon} {dt.label}
                </button>
              ))}
            </div>
          </div>

          {/* 근로계약서 폼 */}
          {(form.docType || "contract") === "contract" && (<>

          {/* (갑) 사업주 정보 */}
          <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 12, border: `1px solid ${T.border}` }}>
            <div style={{ fontSize: 13, fontWeight: 800, color: T.adminHeader, marginBottom: 12 }}>🏢 (갑) 사업주</div>
            <ContractField form={form} setForm={setForm} label="사업체명" fkey="companyName" placeholder="회사명 입력" />
            <ContractField form={form} setForm={setForm} label="대표자" fkey="ownerName" placeholder="대표자 이름" />
            <ContractField form={form} setForm={setForm} label="주소" fkey="bizAddress" placeholder="주소 입력" />
          </div>

          {/* (을) 근로자 정보 */}
          <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 12, border: `1px solid ${T.border}` }}>
            <div style={{ fontSize: 13, fontWeight: 800, color: "#7c3aed", marginBottom: 12 }}>👤 (을) 근로자</div>
            <div style={{ marginBottom: 12, padding: "8px 12px", background: T.bg, borderRadius: 10 }}>
              <span style={{ fontSize: 12, color: T.muted }}>성명: </span>
              <span style={{ fontSize: 13, fontWeight: 700, color: T.text }}>{form.userName}</span>
            </div>
          </div>

          {/* 근로 장소 / 직종 / 계약기간 */}
          <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 12, border: `1px solid ${T.border}` }}>
            <div style={{ fontSize: 13, fontWeight: 800, color: "#0891b2", marginBottom: 12 }}>📋 근로 조건</div>
            <ContractField form={form} setForm={setForm} label="근로 장소" fkey="workPlace" placeholder="근무 장소" />
            <ContractField form={form} setForm={setForm} label="직종" fkey="jobType" placeholder="포장직" />
            <ContractField form={form} setForm={setForm} label="입사일 (근로개시일)" fkey="joinDate" type="date" />
            <div style={{ marginBottom: 12, padding: "10px 12px", background: "#f0f9ff", borderRadius: 10, border: "1px solid #bae6fd" }}>
              <div style={{ fontSize: 11, color: "#0369a1", lineHeight: 1.8 }}>
                본 근로계약은 <b>{form.contractStart || "____년 __월 __일"}</b>에 체결되었으며, 근로개시일은 <b>{form.joinDate || "____년 __월 __일"}</b>로 한다.
                {form.contractStart && form.joinDate && form.contractStart !== form.joinDate && (
                  <><br />본 계약의 내용은 입사일부터 적용하며, 기존 근로계약을 본 계약으로 대체한다.</>
                )}
              </div>
              {form.contractStart && form.joinDate && (
                <div style={{ fontSize: 10, color: "#0891b2", marginTop: 6, fontWeight: 700 }}>
                  {form.contractStart === form.joinDate ? "✅ 신규 입사 계약" : "🔄 재계약 / 계약 변경"}
                </div>
              )}
            </div>
            <ContractField form={form} setForm={setForm} label="계약 시작일" fkey="contractStart" type="date" />
            <div style={{ marginBottom: 12 }}>
              <ContractLabel>계약 종료일 (정규직은 비워두세요)</ContractLabel>
              <input type="date" value={form.contractEnd || ""} onChange={e => setForm(p => ({ ...p, contractEnd: e.target.value }))} style={CONTRACT_ISTYLE} />
            </div>
          </div>

          {/* 임금 구성 */}
          <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 12, border: `1px solid ${T.border}` }}>
            <div style={{ fontSize: 13, fontWeight: 800, color: "#ea580c", marginBottom: 12 }}>💰 임금 구성</div>
            <div style={{ marginBottom: 12 }}>
              <ContractLabel>임금 계산 방법</ContractLabel>
              <select value={form.payType || "월급제"} onChange={e => setForm(p => ({ ...p, payType: e.target.value }))} style={CONTRACT_ISTYLE}>
                {["월급제", "시급제"].map(t => <option key={t} value={t}>{t}</option>)}
              </select>
            </div>
            <ContractField form={form} setForm={setForm} label="시급 (원)" fkey="hourlyWage" type="number" placeholder="9,160" />
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
              <div>
                <ContractLabel>주 근로시간</ContractLabel>
                <input type="number" value={form.weeklyHours || ""} onChange={e => {
                  const wk = Number(e.target.value);
                  const mo = Math.round(wk / 40 * 209);
                  setForm(p => ({ ...p, weeklyHours: e.target.value, monthlyHours: String(mo), monthlyWage: p.hourlyWage ? String(Math.round(Number(p.hourlyWage) * mo)) : "" }));
                }} placeholder="40" style={CONTRACT_ISTYLE} />
              </div>
              <div>
                <ContractLabel>월 근로시간</ContractLabel>
                <input type="number" value={form.monthlyHours || ""} onChange={e => {
                  const mo = Number(e.target.value);
                  setForm(p => ({ ...p, monthlyHours: e.target.value, monthlyWage: p.hourlyWage ? String(Math.round(Number(p.hourlyWage) * mo)) : "" }));
                }} placeholder="209" style={CONTRACT_ISTYLE} />
              </div>
            </div>
            <div style={{ marginBottom: 12 }}>
              <ContractLabel>월급 (시급 × 월근로시간, 자동계산)</ContractLabel>
              <div style={{ padding: "10px 12px", borderRadius: 10, border: `1px solid ${T.border}`, fontSize: 14, fontWeight: 800, color: "#ea580c", background: "#fff7ed" }}>
                {form.hourlyWage && form.monthlyHours
                  ? `${Math.round(Number(form.hourlyWage) * Number(form.monthlyHours)).toLocaleString()}원 (주휴수당 포함)`
                  : "시급과 월근로시간 입력 시 자동계산"}
              </div>
            </div>
            <div style={{ fontSize: 12, color: T.muted, fontWeight: 700, marginBottom: 8, marginTop: 4 }}>임금 구성 항목</div>
            <ContractField form={form} setForm={setForm} label="1. 기본급" fkey="wage1" placeholder="시급*근로일수 (주휴수당 포함)" />
            <ContractField form={form} setForm={setForm} label="2. 연차수당" fkey="wage2" placeholder="시급*8시간 / 별도 지급" />
            <ContractField form={form} setForm={setForm} label="3. 잔업수당" fkey="wage3" placeholder="시급*근로시간*1.5 / 별도 지급" />
            <ContractField form={form} setForm={setForm} label="4. 특근수당" fkey="wage4" placeholder="시급*근로시간*1.5 / 별도 지급" />
            <ContractField form={form} setForm={setForm} label="5. 상여금" fkey="wage5" placeholder="기본급의 100% / 수시 지급" />
            <div style={{ marginTop: 8, padding: "10px 12px", background: "#fff7ed", borderRadius: 10, border: "1px solid #fed7aa" }}>
              <div style={{ fontSize: 11, color: "#92400e", lineHeight: 1.7 }}>
                📌 임금은 관계 법령에 따른 최저임금 변동 및 당사자 간 합의에 의해 변경될 수 있으며, 변경 시 사전 서면 통보로 본 계약의 해당 조항을 갈음한다. 단, 임금의 감액은 근로자의 서면 동의를 요한다.
              </div>
            </div>
          </div>

          {/* 근로시간 / 휴게시간 */}
          <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 12, border: `1px solid ${T.border}` }}>
            <div style={{ fontSize: 13, fontWeight: 800, color: "#16a34a", marginBottom: 12 }}>⏰ 근로시간 · 휴게시간</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
              <div>
                <ContractLabel>출근 시간</ContractLabel>
                <input type="time" value={form.workStart || "07:30"} onChange={e => setForm(p => ({ ...p, workStart: e.target.value }))} style={CONTRACT_ISTYLE} />
              </div>
              <div>
                <ContractLabel>퇴근 시간</ContractLabel>
                <input type="time" value={form.workEnd || "16:30"} onChange={e => setForm(p => ({ ...p, workEnd: e.target.value }))} style={CONTRACT_ISTYLE} />
              </div>
            </div>
            <ContractField form={form} setForm={setForm} label="1일 근로시간 (시간)" fkey="dailyHours" type="number" placeholder="8" />
            <ContractField form={form} setForm={setForm} label="근무 요일" fkey="weekDays" placeholder="월~금" />
            <ContractField form={form} setForm={setForm} label="휴게 - 식사시간" fkey="breakLunch" placeholder="60분(11:40~12:40)" />
            <ContractField form={form} setForm={setForm} label="휴게 - 참시간" fkey="breakSnack" placeholder="20분(16:10~16:30)" />
          </div>

          {/* 연차 / 4대보험 / 복리후생 */}
          <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 12, border: `1px solid ${T.border}` }}>
            <div style={{ fontSize: 13, fontWeight: 800, color: "#0284c7", marginBottom: 12 }}>📅 휴가 · 보험 · 복지</div>
            <ContractField form={form} setForm={setForm} label="연차유급휴가" fkey="annualLeave" placeholder="근로기준법에서 정하는 바에 따라 부여함" />
            <ContractField form={form} setForm={setForm} label="4대보험" fkey="insurance" placeholder="4대보험 의무가입" />
            <ContractField form={form} setForm={setForm} label="복리후생" fkey="welfare" placeholder="교통비, 식대 지원" />
          </div>

          {/* 임금 지급 */}
          <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 12, border: `1px solid ${T.border}` }}>
            <div style={{ fontSize: 13, fontWeight: 800, color: "#ea580c", marginBottom: 12 }}>💳 임금 지급</div>
            <ContractField form={form} setForm={setForm} label="임금 계산 기간" fkey="payCalcPeriod" placeholder="전월 1일부터 전월 말일까지" />
            <ContractField form={form} setForm={setForm} label="지급일 (매월 __일)" fkey="payDay" type="number" placeholder="15" />
            <ContractField form={form} setForm={setForm} label="공휴일 처리" fkey="payHoliday" placeholder="공휴일은 익일 지급" />
            <div style={{ marginBottom: 12 }}>
              <ContractLabel>지급 방법</ContractLabel>
              <select value={form.payMethod || "계좌이체"} onChange={e => setForm(p => ({ ...p, payMethod: e.target.value }))} style={CONTRACT_ISTYLE}>
                {["계좌이체", "현금"].map(t => <option key={t} value={t}>{t}</option>)}
              </select>
            </div>
            <ContractField form={form} setForm={setForm} label="은행명" fkey="bankName" placeholder="국민은행" />
            <ContractField form={form} setForm={setForm} label="계좌번호" fkey="bankAccount" placeholder="계좌번호 입력" />
          </div>

          {/* 퇴직금 / 퇴직절차 / 정년 */}
          <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 12, border: `1px solid ${T.border}` }}>
            <div style={{ fontSize: 13, fontWeight: 800, color: "#b45309", marginBottom: 12 }}>💼 퇴직 · 정년</div>
            <div style={{ marginBottom: 12 }}>
              <ContractLabel>퇴직금</ContractLabel>
              <textarea value={form.severancePay || ""} onChange={e => setForm(p => ({ ...p, severancePay: e.target.value }))}
                style={{ ...CONTRACT_ISTYLE, minHeight: 60, resize: "vertical" }} />
            </div>
            <ContractField form={form} setForm={setForm} label="퇴직 절차" fkey="resignNotice" placeholder="퇴사하기 30일전에 통보한다" />
            <ContractField form={form} setForm={setForm} label="정년" fkey="retirementAge" placeholder="만 60세가 되는 해, 년도말일 기준" />
          </div>

          {/* 근로계약 해지사유 */}
          <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 12, border: `1px solid ${T.border}` }}>
            <div style={{ fontSize: 13, fontWeight: 800, color: "#dc2626", marginBottom: 12 }}>⚠️ 근로계약 해지사유</div>
            <div style={{ marginBottom: 12 }}>
              <ContractLabel>해지 사유</ContractLabel>
              <textarea value={form.terminationReasons || ""} onChange={e => setForm(p => ({ ...p, terminationReasons: e.target.value }))}
                style={{ ...CONTRACT_ISTYLE, minHeight: 140, resize: "vertical" }} />
            </div>
          </div>

          {/* 기타 / 특약 */}
          <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 12, border: `1px solid ${T.border}` }}>
            <div style={{ fontSize: 13, fontWeight: 800, color: T.text, marginBottom: 12 }}>📝 기타 사항</div>
            <ContractField form={form} setForm={setForm} label="상여금 지급 시기" fkey="bonus" placeholder="설 50%, 추석 50%" />
            <div style={{ marginBottom: 12 }}>
              <ContractLabel>특약 사항 (선택)</ContractLabel>
              <textarea value={form.specialTerms || ""} onChange={e => setForm(p => ({ ...p, specialTerms: e.target.value }))}
                placeholder="특약 사항이 있으면 입력하세요"
                style={{ ...CONTRACT_ISTYLE, minHeight: 80, resize: "vertical" }} />
            </div>
          </div>
          </>)}

          {/* 동의서 폼 */}
          {form.docType === "agreement" && (
            <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
              <div style={{ background: T.card, borderRadius: 16, padding: 16, border: `1px solid ${T.border}` }}>
                <div style={{ fontSize: 13, fontWeight: 800, color: "#7c3aed", marginBottom: 10 }}>📋 템플릿 선택</div>
                {AGREEMENT_TEMPLATES.map(t => (
                  <button key={t.key} onClick={() => setForm(p => ({ ...p, docTitle: t.label, docContent: t.content }))}
                    style={{ display: "block", width: "100%", padding: "10px 12px", marginBottom: 8, borderRadius: 10, border: `1px solid ${form.docTitle === t.label ? "#7c3aed" : T.border}`, background: form.docTitle === t.label ? "#f5f3ff" : T.bg, color: T.text, fontSize: 13, fontWeight: 600, cursor: "pointer", textAlign: "left" }}>
                    {form.docTitle === t.label ? "✅ " : ""}{t.label}
                  </button>
                ))}
              </div>
              <div style={{ background: T.card, borderRadius: 16, padding: 16, border: `1px solid ${T.border}` }}>
                <div style={{ fontSize: 13, fontWeight: 800, color: T.text, marginBottom: 10 }}>✍️ 문서 내용</div>
                <div style={{ marginBottom: 12, padding: "8px 12px", background: T.bg, borderRadius: 10 }}>
                  <span style={{ fontSize: 12, color: T.muted }}>대상: </span>
                  <span style={{ fontSize: 13, fontWeight: 700, color: T.text }}>{form.userName}</span>
                </div>
                <ContractField form={form} setForm={setForm} label="제목" fkey="docTitle" placeholder="동의서 제목" />
                <div style={{ marginBottom: 12 }}>
                  <ContractLabel>내용</ContractLabel>
                  <textarea value={form.docContent || ""} onChange={e => setForm(p => ({ ...p, docContent: e.target.value }))}
                    placeholder="동의서 내용을 입력하세요" style={{ ...CONTRACT_ISTYLE, minHeight: 200, resize: "vertical" }} />
                </div>
              </div>
            </div>
          )}

          {/* 확인서 폼 */}
          {form.docType === "confirm" && (() => {
            const now = new Date();
            const year = now.getFullYear();
            const quarter = Math.ceil((now.getMonth() + 1) / 3);
            const autoTitle = (key, label) => {
              if (key === "safety") return `${label} ${year}-${quarter}분기`;
              if (key === "harassment") return `${label} ${year}`;
              if (key === "disability") return `${label} ${year}`;
              if (key === "privacy") return `${label} ${year}`;
              if (key === "wage") return `${label} ${year}`;
              if (key === "rule") return `${label} ${year}`;
              return label;
            };
            return (
            <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
              <div style={{ background: T.card, borderRadius: 16, padding: 16, border: `1px solid ${T.border}` }}>
                <div style={{ fontSize: 13, fontWeight: 800, color: "#16a34a", marginBottom: 10 }}>📋 템플릿 선택</div>
                {CONFIRM_TEMPLATES.map(t => {
                  const title = autoTitle(t.key, t.label);
                  return (
                    <button key={t.key} onClick={() => setForm(p => ({ ...p, docTitle: title, docContent: t.content }))}
                      style={{ display: "block", width: "100%", padding: "10px 12px", marginBottom: 8, borderRadius: 10, border: `1px solid ${form.docTitle === title ? "#16a34a" : T.border}`, background: form.docTitle === title ? "#f0fdf4" : T.bg, color: T.text, fontSize: 13, fontWeight: 600, cursor: "pointer", textAlign: "left" }}>
                      {form.docTitle === title ? "✅ " : ""}{t.label}
                      {t.key !== "custom" && <span style={{ fontSize: 11, color: T.muted, marginLeft: 6 }}>({title.replace(t.label, "").trim()})</span>}
                    </button>
                  );
                })}
              </div>
              <div style={{ background: T.card, borderRadius: 16, padding: 16, border: `1px solid ${T.border}` }}>
                <div style={{ fontSize: 13, fontWeight: 800, color: T.text, marginBottom: 10 }}>✅ 문서 내용</div>
                <div style={{ marginBottom: 12, padding: "8px 12px", background: T.bg, borderRadius: 10 }}>
                  <span style={{ fontSize: 12, color: T.muted }}>대상: </span>
                  <span style={{ fontSize: 13, fontWeight: 700, color: T.text }}>{form.userName}</span>
                </div>
                <ContractField form={form} setForm={setForm} label="제목" fkey="docTitle" placeholder="확인서 제목" />
                <div style={{ marginBottom: 12 }}>
                  <ContractLabel>내용</ContractLabel>
                  <textarea value={form.docContent || ""} onChange={e => setForm(p => ({ ...p, docContent: e.target.value }))}
                    placeholder="확인서 내용을 입력하세요" style={{ ...CONTRACT_ISTYLE, minHeight: 200, resize: "vertical" }} />
                </div>
              </div>
            </div>
            );
          })()}

          {/* 기타 문서 폼 */}
          {form.docType === "other" && (
            <div style={{ background: T.card, borderRadius: 16, padding: 16, border: `1px solid ${T.border}` }}>
              <div style={{ fontSize: 13, fontWeight: 800, color: T.text, marginBottom: 10 }}>📄 문서 내용</div>
              <div style={{ marginBottom: 12, padding: "8px 12px", background: T.bg, borderRadius: 10 }}>
                <span style={{ fontSize: 12, color: T.muted }}>대상: </span>
                <span style={{ fontSize: 13, fontWeight: 700, color: T.text }}>{form.userName}</span>
              </div>
              <ContractField form={form} setForm={setForm} label="제목" fkey="docTitle" placeholder="문서 제목" />
              <div style={{ marginBottom: 12 }}>
                <ContractLabel>내용</ContractLabel>
                <textarea value={form.docContent || ""} onChange={e => setForm(p => ({ ...p, docContent: e.target.value }))}
                  placeholder="내용을 자유롭게 입력하세요" style={{ ...CONTRACT_ISTYLE, minHeight: 200, resize: "vertical" }} />
              </div>
            </div>
          )}

        </div>

        <div style={{ position: "fixed", bottom: 0, left: 0, right: 0, background: T.card, borderTop: `1px solid ${T.border}`, padding: "12px 16px", paddingBottom: "calc(12px + env(safe-area-inset-bottom))" }}>
          <button onClick={saveContract} disabled={saving}
            style={{ width: "100%", padding: "14px 0", borderRadius: 14, border: "none", background: T.adminHeader, color: "#fff", fontSize: 15, fontWeight: 800, cursor: "pointer", opacity: saving ? 0.6 : 1 }}>
            {saving ? "저장 중..." : "💾 저장"}
          </button>
        </div>
      </div>
    );
  }

  // 목록 화면
  const filteredContracts = (userId) => contracts.filter(c => c.userId === userId && (c.docType || "contract") === docTypeFilter);
  const isContractTab = docTypeFilter === "contract";

  return (
    <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'Noto Sans KR',sans-serif", paddingBottom: 30 }}>
      <div style={{ background: T.adminHeader, paddingTop: "calc(16px + env(safe-area-inset-top))", paddingBottom: "14px", paddingLeft: "16px", paddingRight: "16px" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <button onClick={onBack} style={{ background: "#ffffff18", border: "none", color: "#fff", fontSize: 18, cursor: "pointer", padding: "8px 14px", borderRadius: 12, fontWeight: 700 }}>‹</button>
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 11, color: "#ffffff40", letterSpacing: 3 }}>ADMIN</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>📄 문서함</div>
          </div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6, marginTop: 12 }}>
          {DOC_TYPES.map(dt => (
            <button key={dt.key} onClick={() => setDocTypeFilter(dt.key)}
              style={{ padding: "7px 0", borderRadius: 10, border: "none", background: docTypeFilter === dt.key ? "#fff" : "#ffffff25", color: docTypeFilter === dt.key ? T.adminHeader : "#fff", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
              {dt.icon} {dt.label}
            </button>
          ))}
        </div>
      </div>

      {/* PIN 확인 모달 */}
      {pinModal && (
        <div style={{ position: "fixed", inset: 0, background: "#00000060", zIndex: 999, display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
          <div style={{ background: "#fff", borderRadius: 20, padding: 24, width: "100%", maxWidth: 340, boxShadow: "0 8px 32px #0000003a" }}>
            <div style={{ fontSize: 32, textAlign: "center", marginBottom: 8 }}>🔐</div>
            <div style={{ fontSize: 15, fontWeight: 800, color: T.text, textAlign: "center", marginBottom: 6 }}>서명완료 문서 삭제</div>
            <div style={{ fontSize: 12, color: "#dc2626", textAlign: "center", marginBottom: 4, fontWeight: 600 }}>⚠️ 근로기준법상 3년 보관 의무 대상</div>
            <div style={{ fontSize: 12, color: T.muted, textAlign: "center", marginBottom: 16 }}>
              {pinModal.contract.userName}님의 서명완료 문서입니다.<br />삭제하려면 관리자 PIN을 입력하세요.
            </div>
            <input type="password" value={pinInput} onChange={e => { setPinInput(e.target.value); setPinErr(""); }}
              placeholder="관리자 PIN 입력"
              style={{ width: "100%", padding: "12px 14px", borderRadius: 12, border: `1px solid ${pinErr ? "#fca5a5" : T.border}`, fontSize: 15, textAlign: "center", letterSpacing: 6, boxSizing: "border-box", fontFamily: "inherit" }}
              autoFocus />
            {pinErr && <div style={{ fontSize: 12, color: "#dc2626", textAlign: "center", marginTop: 6, fontWeight: 600 }}>{pinErr}</div>}
            <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
              <button onClick={() => setPinModal(null)}
                style={{ flex: 1, padding: "12px 0", borderRadius: 12, border: `1px solid ${T.border}`, background: T.bg, color: T.text, fontSize: 14, fontWeight: 700, cursor: "pointer" }}>취소</button>
              <button onClick={confirmPinDelete}
                style={{ flex: 1, padding: "12px 0", borderRadius: 12, border: "none", background: "#dc2626", color: "#fff", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>삭제</button>
            </div>
          </div>
        </div>
      )}

      <div style={{ padding: 16 }}>

        {members.map(u => {
          const myDocs = filteredContracts(u.id);
          const showHistory = expandedHistory[`hist_${docTypeFilter}_${u.id}`];

          return (
            <div key={u.id} style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 12, border: `1px solid ${T.border}`, boxShadow: "0 2px 8px #0000000d" }}>
              {/* 직원명 + 작성 버튼 */}
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: myDocs.length > 0 ? 10 : 0 }}>
                <div>
                  <div style={{ fontWeight: 800, fontSize: 15, color: T.text }}>{u.name}</div>
                  {isContractTab && myDocs[0] && (
                    <div style={{ fontSize: 11, color: T.muted, marginTop: 2 }}>
                      {myDocs[0].contractStart} ~ {myDocs[0].contractEnd || "기간 없음"} · {myDocs[0].jobType || ""}
                    </div>
                  )}
                  {!isContractTab && myDocs.length > 0 && (
                    <div style={{ fontSize: 11, color: T.muted, marginTop: 2 }}>총 {myDocs.length}건</div>
                  )}
                </div>
                {isContractTab && myDocs[0] && (
                  <span style={{ fontSize: 11, fontWeight: 700, color: statusLabel(myDocs[0].status).color, background: statusLabel(myDocs[0].status).bg, borderRadius: 8, padding: "3px 8px" }}>
                    {statusLabel(myDocs[0].status).text}
                  </span>
                )}
              </div>

              {/* ── 근로계약서: 최신 1건 + 히스토리 ── */}
              {isContractTab && (() => {
                const latest = myDocs[0];
                const history = myDocs.slice(1);
                return (
                  <>
                    {latest?.status === "signed" && latest.signedAt && (
                      <div style={{ fontSize: 11, color: "#16a34a", marginBottom: 8, background: "#dcfce7", padding: "8px 10px", borderRadius: 8 }}>
                        <div>✅ 서명일시: {new Date(latest.signedAt).toLocaleString("ko-KR")}</div>
                        {latest.empAddress && <div style={{ marginTop: 3 }}>📍 {latest.empAddress}</div>}
                        {latest.empPhone && <div style={{ marginTop: 3 }}>📞 {latest.empPhone}</div>}
                        {latest.signIp && <div style={{ marginTop: 3 }}>🌐 {latest.signIp} · {latest.signDevice}</div>}
                      </div>
                    )}
                    <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                      <button onClick={() => openNew(u, "contract")}
                        style={{ padding: "8px 14px", borderRadius: 10, border: "none", background: T.adminHeader, color: "#fff", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
                        {latest ? "✏️ 새 계약서" : "📝 작성"}
                      </button>
                      {latest && latest.status === "draft" && (
                        <button onClick={() => sendToEmployee(latest)}
                          style={{ padding: "8px 14px", borderRadius: 10, border: "none", background: "#0891b2", color: "#fff", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
                          📨 서명 요청
                        </button>
                      )}
                      {latest && latest.status !== "signed" && (
                        <button onClick={() => openEdit(latest)}
                          style={{ padding: "8px 14px", borderRadius: 10, border: `1px solid ${T.border}`, background: T.bg, color: T.text, fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
                          수정
                        </button>
                      )}
                      {latest && (
                        <>
                          <button id={`contract-pdf-btn-${latest.id}`} onClick={() => downloadContractPDF(latest)} disabled={pdfLoading === latest.id}
                            style={{ padding: "8px 14px", borderRadius: 10, border: "none", background: "#16a34a", color: "#fff", fontSize: 12, fontWeight: 700, cursor: "pointer", opacity: pdfLoading === latest.id ? 0.6 : 1 }}>
                            {pdfLoading === latest.id ? "생성 중..." : "⬇ PDF"}
                          </button>
                          {latest.pdfUrl && (
                            <a href={latest.pdfUrl} target="_blank" rel="noreferrer"
                              style={{ padding: "8px 14px", borderRadius: 10, border: `1px solid #16a34a`, background: "#f0fdf4", color: "#16a34a", fontSize: 12, fontWeight: 700, textDecoration: "none" }}>
                              🔗 저장본
                            </a>
                          )}
                          <button onClick={() => deleteContract(latest)}
                            style={{ padding: "8px 14px", borderRadius: 10, border: "none", background: "#fee2e2", color: "#b91c1c", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
                            삭제
                          </button>
                        </>
                      )}
                    </div>
                    {/* 숨김 인쇄 */}
                    {latest && (
                      <div id={`contract-print-${latest.id}`} style={{ position: "fixed", left: -9999, top: 0, width: 794, background: "#fff", padding: "40px 50px", fontFamily: "'Noto Sans KR', sans-serif", fontSize: 12, color: "#111", lineHeight: 1.8 }}>
                        <div style={{ textAlign: "center", fontSize: 20, fontWeight: 900, marginBottom: 24, letterSpacing: 8 }}>근 로 계 약 서</div>
                        <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: 16, fontSize: 12 }}>
                          <tbody>
                            <tr><td style={{ padding: "4px 8px", fontWeight: 700, width: 80 }}>(갑) 사용자</td><td style={{ padding: "4px 8px" }}>사업체명: {latest.companyName} &nbsp;&nbsp; 대표자: {latest.ownerName}</td></tr>
                            <tr><td style={{ padding: "4px 8px" }}></td><td style={{ padding: "4px 8px" }}>주소: {latest.bizAddress}</td></tr>
                            <tr><td style={{ padding: "4px 8px", fontWeight: 700 }}>(을) 근로자</td><td style={{ padding: "4px 8px" }}>성명: {latest.userName} &nbsp;&nbsp; 연락처: {latest.empPhone || "__________"}</td></tr>
                            <tr><td style={{ padding: "4px 8px" }}></td><td style={{ padding: "4px 8px" }}>주소: {latest.empAddress || "__________________________________________"}</td></tr>
                          </tbody>
                        </table>
                        <div style={{ fontSize: 11, marginBottom: 16 }}>위 당사자는 아래의 근로조건을 성실히 이행할 것을 약정하고 근로계약을 체결한다.</div>
                        {latest.contractStart && latest.joinDate && (
                          <div style={{ fontSize: 11, marginBottom: 12, padding: "6px 10px", background: "#f0f9ff", borderRadius: 4 }}>
                            본 근로계약은 {latest.contractStart}에 체결되었으며, 근로개시일은 {latest.joinDate}로 한다.
                            {latest.contractStart !== latest.joinDate && " 본 계약의 내용은 입사일부터 적용하며, 기존 근로계약을 본 계약으로 대체한다."}
                          </div>
                        )}
                        {[
                          ["근로 장소", latest.workPlace], ["직종", latest.jobType],
                          ["계약 기간", `${latest.contractStart} ~ ${latest.contractEnd || "기간 없음 (정규직)"}`],
                          ["근로 시간", `${latest.workStart} ~ ${latest.workEnd} (1일 ${latest.dailyHours || 8}시간, ${latest.weekDays || "월~금"})`],
                          ["휴게 시간", `식사 ${latest.breakLunch || ""}  참 ${latest.breakSnack || ""}`],
                          ["시급", latest.hourlyWage ? `${Number(latest.hourlyWage).toLocaleString()}원` : ""],
                          ["월급", latest.monthlyWage ? `${Number(latest.monthlyWage).toLocaleString()}원 (주휴수당 포함)` : ""],
                          ["연차수당", latest.wage2], ["잔업수당", latest.wage3], ["특근수당", latest.wage4], ["상여금", latest.wage5],
                          ["임금 계산기간", latest.payCalcPeriod],
                          ["임금 지급일", latest.payDay ? `매월 ${latest.payDay}일 (${latest.payHoliday || ""})` : ""],
                          ["지급 방법", `${latest.payMethod || ""} ${latest.bankName ? `/ ${latest.bankName} ${latest.bankAccount || ""}` : ""}`],
                          ["연차유급휴가", latest.annualLeave], ["4대보험", latest.insurance], ["복리후생", latest.welfare],
                          ["퇴직금", latest.severancePay], ["퇴직 절차", latest.resignNotice], ["정년", latest.retirementAge],
                          ["상여금 지급시기", latest.bonus],
                        ].filter(([, v]) => v).map(([label, value]) => (
                          <div key={label} style={{ display: "flex", borderBottom: "1px solid #e5e7eb", padding: "3px 0" }}>
                            <span style={{ minWidth: 100, fontWeight: 700, fontSize: 11 }}>{label}</span>
                            <span style={{ fontSize: 11, flex: 1 }}>{value}</span>
                          </div>
                        ))}
                        {latest.terminationReasons && (
                          <div style={{ marginTop: 10 }}>
                            <div style={{ fontWeight: 700, fontSize: 11, marginBottom: 4 }}>근로계약 해지사유</div>
                            <div style={{ fontSize: 10, whiteSpace: "pre-wrap", lineHeight: 1.7 }}>{latest.terminationReasons}</div>
                          </div>
                        )}
                        <div style={{ marginTop: 10, padding: "6px 10px", background: "#fff7ed", borderRadius: 4, fontSize: 10, color: "#92400e" }}>
                          📌 임금은 관계 법령에 따른 최저임금 변동 및 당사자 간 합의에 의해 변경될 수 있으며, 변경 시 사전 서면 통보로 본 계약의 해당 조항을 갈음한다. 단, 임금의 감액은 근로자의 서면 동의를 요한다.
                        </div>
                        <div style={{ marginTop: 8, fontSize: 10, color: "#555" }}>이 계약에 정함이 없는 사항은 근로기준법에 의함</div>
                        {latest.specialTerms && <div style={{ marginTop: 6, fontSize: 10 }}><b>특약:</b> {latest.specialTerms}</div>}
                        <div style={{ marginTop: 24, fontSize: 12 }}>{latest.contractStart?.replace(/-/g, "년 ").replace(/-/, "월 ")}일</div>
                        <div style={{ marginTop: 16, display: "flex", justifyContent: "space-around", fontSize: 12 }}>
                          <div>(사용자) {latest.ownerName} &nbsp;&nbsp; {latest.sentAt ? `📨 ${new Date(latest.sentAt).toLocaleDateString("ko-KR")} 발송` : ""}</div>
                          <div>(근로자) {latest.userName} &nbsp;&nbsp; {latest.status === "signed" && latest.signedAt ? `✅ ${new Date(latest.signedAt).toLocaleDateString("ko-KR")} 전자서명` : ""}</div>
                        </div>
                      </div>
                    )}
                    {/* 이전 근로계약서 히스토리 */}
                    {history.length > 0 && (
                      <div style={{ marginTop: 12 }}>
                        <button onClick={() => setExpandedHistory(p => ({ ...p, [`hist_contract_${u.id}`]: !p[`hist_contract_${u.id}`] }))}
                          style={{ width: "100%", padding: "7px 0", borderRadius: 8, border: `1px dashed ${T.border}`, background: "transparent", color: T.muted, fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
                          📁 이전 근로계약서 {history.length}건 {showHistory ? "▲ 접기" : "▼ 보기"}
                        </button>
                        {showHistory && (
                          <div style={{ marginTop: 8, display: "flex", flexDirection: "column", gap: 6 }}>
                            {history.map(c => {
                              const hst = statusLabel(c.status);
                              const dk = `hist_det_${c.id}`;
                              const sd = expandedHistory[dk];
                              return (
                                <div key={c.id} style={{ background: T.bg, borderRadius: 10, border: `1px solid ${T.border}`, overflow: "hidden" }}>
                                  <div style={{ padding: "10px 12px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                                    <div>
                                      <div style={{ fontSize: 12, fontWeight: 700, color: T.text }}>{c.contractStart} ~ {c.contractEnd || "기간 없음"}</div>
                                      <div style={{ fontSize: 11, color: T.muted, marginTop: 2 }}>
                                        작성: {c.createdAt ? new Date(c.createdAt).toLocaleDateString("ko-KR") : "-"}
                                        {c.signedAt ? ` · 서명: ${new Date(c.signedAt).toLocaleDateString("ko-KR")}` : ""}
                                      </div>
                                    </div>
                                    <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                                      <span style={{ fontSize: 10, fontWeight: 700, color: hst.color, background: hst.bg, borderRadius: 6, padding: "2px 7px" }}>{hst.text}</span>
                                      <button onClick={() => setExpandedHistory(p => ({ ...p, [dk]: !p[dk] }))}
                                        style={{ padding: "4px 8px", borderRadius: 6, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 11, fontWeight: 700, cursor: "pointer" }}>
                                        {sd ? "▲" : "▼"}
                                      </button>
                                      {c.status !== "signed" && (
                                        <button onClick={() => deleteContract(c)}
                                          style={{ padding: "4px 8px", borderRadius: 6, border: "none", background: "#fee2e2", color: "#b91c1c", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>삭제</button>
                                      )}
                                    </div>
                                  </div>
                                  {sd && (
                                    <div style={{ borderTop: `1px solid ${T.border}`, padding: "10px 12px", fontSize: 11 }}>
                                      {[["직종", c.jobType], ["시급", c.hourlyWage ? `${Number(c.hourlyWage).toLocaleString()}원` : null],
                                        ["월급", c.monthlyWage ? `${Number(c.monthlyWage).toLocaleString()}원` : null],
                                        ["지급일", c.payDay ? `매월 ${c.payDay}일` : null],
                                      ].filter(([, v]) => v).map(([label, value]) => (
                                        <div key={label} style={{ display: "flex", gap: 8, padding: "3px 0", borderBottom: `1px solid ${T.border}` }}>
                                          <span style={{ color: T.muted, minWidth: 60 }}>{label}</span>
                                          <span style={{ fontWeight: 600 }}>{value}</span>
                                        </div>
                                      ))}
                                      {c.signedAt && (
                                        <div style={{ marginTop: 6, padding: "6px 8px", background: "#dcfce7", borderRadius: 6 }}>
                                          <div style={{ color: "#15803d", fontWeight: 700 }}>✅ {new Date(c.signedAt).toLocaleString("ko-KR")}</div>
                                          {c.empAddress && <div style={{ color: "#15803d" }}>📍 {c.empAddress}</div>}
                                          {c.signIp && <div style={{ color: "#15803d" }}>🌐 {c.signIp} · {c.signDevice}</div>}
                                        </div>
                                      )}
                                    </div>
                                  )}
                                </div>
                              );
                            })}
                          </div>
                        )}
                      </div>
                    )}
                  </>
                );
              })()}

              {/* ── 동의서/확인서/기타: 모두 나란히 ── */}
              {!isContractTab && (
                <>
                  <button onClick={() => openNew(u, docTypeFilter)}
                    style={{ padding: "8px 14px", borderRadius: 10, border: "none", background: T.adminHeader, color: "#fff", fontSize: 12, fontWeight: 700, cursor: "pointer", marginBottom: myDocs.length > 0 ? 10 : 0 }}>
                    📝 {DOC_TYPES.find(d => d.key === docTypeFilter)?.label || "문서"} 작성
                  </button>
                  {myDocs.length === 0 && (
                    <div style={{ fontSize: 12, color: T.muted }}>작성된 문서가 없습니다</div>
                  )}
                  {/* docTitle 기준으로 그룹핑 */}
                  {(() => {
                    const groups = [];
                    const seen = {};
                    myDocs.forEach(c => {
                      const key = c.docTitle || "__no_title__";
                      if (!seen[key]) { seen[key] = []; groups.push(seen[key]); }
                      seen[key].push(c);
                    });
                    return groups.map((group, gi) => {
                      const latest = group[0];
                      const history = group.slice(1);
                      const dst = statusLabel(latest.status);
                      const dk = `det_${latest.id}`;
                      const sd = expandedHistory[dk];
                      const histKey = `ghist_${latest.docTitle || gi}_${u.id}`;
                      const showGHist = expandedHistory[histKey];
                      return (
                        <div key={latest.id} style={{ background: T.bg, borderRadius: 12, border: `1px solid ${T.border}`, overflow: "hidden", marginBottom: 8 }}>
                          {/* 최신 문서 헤더 */}
                          <div style={{ padding: "12px 14px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                            <div style={{ flex: 1, minWidth: 0 }}>
                              <div style={{ fontSize: 13, fontWeight: 700, color: T.text, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                                {latest.docTitle || DOC_TYPES.find(d => d.key === latest.docType)?.label || ""}
                              </div>
                              <div style={{ fontSize: 11, color: T.muted, marginTop: 2 }}>
                                {latest.createdAt ? new Date(latest.createdAt).toLocaleDateString("ko-KR") : ""}
                                {latest.signedAt ? ` · 서명 ${new Date(latest.signedAt).toLocaleDateString("ko-KR")}` : ""}
                              </div>
                            </div>
                            <div style={{ display: "flex", alignItems: "center", gap: 6, flexShrink: 0 }}>
                              <span style={{ fontSize: 10, fontWeight: 700, color: dst.color, background: dst.bg, borderRadius: 6, padding: "2px 7px" }}>{dst.text}</span>
                              <button onClick={() => setExpandedHistory(p => ({ ...p, [dk]: !p[dk] }))}
                                style={{ padding: "4px 8px", borderRadius: 6, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 11, fontWeight: 700, cursor: "pointer" }}>
                                {sd ? "▲" : "▼"}
                              </button>
                            </div>
                          </div>
                          {/* 최신 문서 상세 */}
                          {sd && (
                            <div style={{ borderTop: `1px solid ${T.border}`, padding: "12px 14px" }}>
                              {latest.docContent && (
                                <div style={{ fontSize: 12, color: T.text, lineHeight: 1.7, whiteSpace: "pre-wrap", marginBottom: 10 }}>{latest.docContent}</div>
                              )}
                              {latest.signedAt && (
                                <div style={{ padding: "8px 10px", background: "#dcfce7", borderRadius: 8, marginBottom: 8 }}>
                                  <div style={{ fontSize: 11, color: "#15803d", fontWeight: 700 }}>✅ {new Date(latest.signedAt).toLocaleString("ko-KR")}</div>
                                  {latest.signIp && <div style={{ fontSize: 11, color: "#16a34a", marginTop: 2 }}>🌐 {latest.signIp} · {latest.signDevice}</div>}
                                  {latest.attachment && (
                                    <a href={latest.attachment.url} target="_blank" rel="noreferrer"
                                      style={{ display: "inline-flex", alignItems: "center", gap: 4, marginTop: 6, padding: "4px 10px", borderRadius: 6, background: "#fff", color: "#0369a1", fontSize: 11, fontWeight: 700, textDecoration: "none", border: "1px solid #bae6fd" }}>
                                      📎 {latest.attachment.name}
                                    </a>
                                  )}
                                </div>
                              )}
                              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                                {latest.status === "draft" && (
                                  <button onClick={() => sendToEmployee(latest)}
                                    style={{ padding: "6px 12px", borderRadius: 8, border: "none", background: "#0891b2", color: "#fff", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>
                                    📨 서명 요청
                                  </button>
                                )}
                                {latest.status !== "signed" && (
                                  <button onClick={() => openEdit(latest)}
                                    style={{ padding: "6px 12px", borderRadius: 8, border: `1px solid ${T.border}`, background: T.bg, color: T.text, fontSize: 11, fontWeight: 700, cursor: "pointer" }}>
                                    수정
                                  </button>
                                )}
                                <button id={`contract-pdf-btn-${latest.id}`} onClick={() => downloadContractPDF(latest)} disabled={pdfLoading === latest.id}
                                  style={{ padding: "6px 12px", borderRadius: 8, border: "none", background: "#16a34a", color: "#fff", fontSize: 11, fontWeight: 700, cursor: "pointer", opacity: pdfLoading === latest.id ? 0.6 : 1 }}>
                                  {pdfLoading === latest.id ? "생성 중..." : "⬇ PDF"}
                                </button>
                                {latest.pdfUrl && (
                                  <a href={latest.pdfUrl} target="_blank" rel="noreferrer"
                                    style={{ padding: "6px 12px", borderRadius: 8, border: `1px solid #16a34a`, background: "#f0fdf4", color: "#16a34a", fontSize: 11, fontWeight: 700, textDecoration: "none" }}>
                                    🔗 저장본
                                  </a>
                                )}
                                <button onClick={() => deleteContract(latest)}
                                  style={{ padding: "6px 12px", borderRadius: 8, border: "none", background: "#fee2e2", color: "#b91c1c", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>
                                  🗑 삭제
                                </button>
                              </div>
                              {/* 숨김 인쇄 */}
                              <div id={`contract-print-${latest.id}`} style={{ position: "fixed", left: -9999, top: 0, width: 794, background: "#fff", padding: "40px 50px", fontFamily: "'Noto Sans KR', sans-serif", fontSize: 12, color: "#111", lineHeight: 1.8 }}>
                                <div style={{ textAlign: "center", fontSize: 20, fontWeight: 900, marginBottom: 8, letterSpacing: 4 }}>
                                  {DOC_TYPES.find(d => d.key === latest.docType)?.label || "문서"}
                                </div>
                                {latest.docTitle && <div style={{ textAlign: "center", fontSize: 14, fontWeight: 700, marginBottom: 24, color: "#444" }}>{latest.docTitle}</div>}
                                <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: 20, fontSize: 12 }}>
                                  <tbody>
                                    <tr><td style={{ padding: "4px 8px", fontWeight: 700, width: 80 }}>사업체명</td><td style={{ padding: "4px 8px" }}>하나기업</td></tr>
                                    <tr><td style={{ padding: "4px 8px", fontWeight: 700 }}>대표자</td><td style={{ padding: "4px 8px" }}>박용균</td></tr>
                                    <tr><td style={{ padding: "4px 8px", fontWeight: 700 }}>성명</td><td style={{ padding: "4px 8px" }}>{latest.userName}</td></tr>
                                  </tbody>
                                </table>
                                {latest.docContent && (
                                  <div style={{ fontSize: 12, lineHeight: 1.9, whiteSpace: "pre-wrap", marginBottom: 24, padding: "12px 0", borderTop: "1px solid #e5e7eb", borderBottom: "1px solid #e5e7eb" }}>
                                    {latest.docContent}
                                  </div>
                                )}
                                <div style={{ marginTop: 24, fontSize: 12 }}>{latest.createdAt ? new Date(latest.createdAt).toLocaleDateString("ko-KR", { year: "numeric", month: "long", day: "numeric" }) : ""}</div>
                                <div style={{ marginTop: 16, display: "flex", justifyContent: "space-around", fontSize: 12 }}>
                                  <div>(사업자) 박용균 &nbsp;&nbsp; {latest.sentAt ? `📨 ${new Date(latest.sentAt).toLocaleDateString("ko-KR")} 발송` : ""}</div>
                                  <div>(서명인) {latest.userName} &nbsp;&nbsp; {latest.status === "signed" && latest.signedAt ? `✅ ${new Date(latest.signedAt).toLocaleDateString("ko-KR")} 전자서명` : ""}</div>
                                </div>
                              </div>
                            </div>
                          )}
                          {/* 이전 버전 히스토리 */}
                          {history.length > 0 && (
                            <div style={{ borderTop: `1px solid ${T.border}`, padding: "8px 14px", background: "#fafafa" }}>
                              <button onClick={() => setExpandedHistory(p => ({ ...p, [histKey]: !p[histKey] }))}
                                style={{ width: "100%", padding: "5px 0", background: "transparent", border: "none", color: T.muted, fontSize: 11, fontWeight: 700, cursor: "pointer", textAlign: "left" }}>
                                📁 이전 버전 {history.length}건 {showGHist ? "▲ 접기" : "▼ 보기"}
                              </button>
                              {showGHist && (
                                <div style={{ marginTop: 6, display: "flex", flexDirection: "column", gap: 6 }}>
                                  {history.map(c => {
                                    const hst = statusLabel(c.status);
                                    const hdk = `hdet_${c.id}`;
                                    const hsd = expandedHistory[hdk];
                                    return (
                                      <div key={c.id} style={{ background: "#fff", borderRadius: 8, border: `1px solid ${T.border}`, overflow: "hidden" }}>
                                        <div style={{ padding: "8px 12px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                                          <div>
                                            <div style={{ fontSize: 11, color: T.muted }}>
                                              작성: {c.createdAt ? new Date(c.createdAt).toLocaleDateString("ko-KR") : "-"}
                                              {c.signedAt ? ` · 서명: ${new Date(c.signedAt).toLocaleDateString("ko-KR")}` : ""}
                                            </div>
                                          </div>
                                          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                                            <span style={{ fontSize: 10, fontWeight: 700, color: hst.color, background: hst.bg, borderRadius: 6, padding: "2px 6px" }}>{hst.text}</span>
                                            <button onClick={() => setExpandedHistory(p => ({ ...p, [hdk]: !p[hdk] }))}
                                              style={{ padding: "3px 7px", borderRadius: 5, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 10, fontWeight: 700, cursor: "pointer" }}>
                                              {hsd ? "▲" : "▼"}
                                            </button>
                                            <button onClick={() => deleteContract(c)}
                                              style={{ padding: "3px 7px", borderRadius: 5, border: "none", background: "#fee2e2", color: "#b91c1c", fontSize: 10, fontWeight: 700, cursor: "pointer" }}>삭제</button>
                                          </div>
                                        </div>
                                        {hsd && (
                                          <div style={{ borderTop: `1px solid ${T.border}`, padding: "8px 12px" }}>
                                            {c.docContent && (
                                              <div style={{ fontSize: 11, color: T.text, lineHeight: 1.6, whiteSpace: "pre-wrap", marginBottom: 6 }}>{c.docContent}</div>
                                            )}
                                            {c.signedAt && (
                                              <div style={{ padding: "5px 8px", background: "#dcfce7", borderRadius: 6 }}>
                                                <div style={{ fontSize: 10, color: "#15803d", fontWeight: 700 }}>✅ {new Date(c.signedAt).toLocaleString("ko-KR")}</div>
                                                {c.signIp && <div style={{ fontSize: 10, color: "#16a34a" }}>🌐 {c.signIp} · {c.signDevice}</div>}
                                                {c.attachment && (
                                                  <a href={c.attachment.url} target="_blank" rel="noreferrer"
                                                    style={{ display: "inline-flex", alignItems: "center", gap: 4, marginTop: 4, fontSize: 10, color: "#0369a1", fontWeight: 700, textDecoration: "none" }}>
                                                    📎 {c.attachment.name}
                                                  </a>
                                                )}
                                              </div>
                                            )}
                                          </div>
                                        )}
                                      </div>
                                    );
                                  })}
                                </div>
                              )}
                            </div>
                          )}
                        </div>
                      );
                    });
                  })()}
                </>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}
function ContractViewScreen({ user, contracts }) {
  const [docTypeTab, setDocTypeTab] = useState("contract");
  const myDocs = contracts.filter(c => c.userId === user.id && (c.docType || "contract") === docTypeTab && (c.status === "sent" || c.status === "signed"));
  const contract = myDocs[0]; // 근로계약서용 최신 1건
  const [signing, setSigning] = useState(false);
  const [confirmed, setConfirmed] = useState(false);
  const [showDetail, setShowDetail] = useState(false);
  const [signAddr, setSignAddr] = useState("");
  const [signPhone, setSignPhone] = useState("");
  const [pdfLoading, setPdfLoading] = useState(false);
  const [expandedDoc, setExpandedDoc] = useState({});

  const downloadMyPDF = async () => {
    if (!contract) return;
    setPdfLoading(true);
    try {
      const el = document.getElementById(`contract-member-print-${contract.id}`);
      if (!el) { alert("PDF 생성 영역을 찾을 수 없습니다."); setPdfLoading(false); return; }
      const btn = document.getElementById("member-pdf-btn");
      if (btn) btn.style.visibility = "hidden";
      const canvas = await html2canvas(el, { scale: 2.5, useCORS: true, backgroundColor: "#ffffff" });
      if (btn) btn.style.visibility = "visible";
      const imgData = canvas.toDataURL("image/png");
      const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
      const pageW = pdf.internal.pageSize.getWidth();
      const pageH = pdf.internal.pageSize.getHeight();
      const margin = 8;
      const maxW = pageW - margin * 2;
      const maxH = pageH - margin * 2;
      const ratio = canvas.width / canvas.height;
      let imgW = maxW, imgH = imgW / ratio;
      if (imgH > maxH) { imgH = maxH; imgW = imgH * ratio; }
      pdf.addImage(imgData, "PNG", (pageW - imgW) / 2, margin, imgW, imgH);
      pdf.save(`${user.name}_근로계약서_${contract.contractStart || ""}.pdf`);
      // ✅ 다운로드 완료 즉시 버튼 정상화
      setPdfLoading(false);
    } catch(e) {
      alert("PDF 생성 실패: " + e.message);
      setPdfLoading(false);
    }
  };

  const sign = async () => {
    if (!contract) return;
    if (!confirmed) { alert("내용에 동의 체크 후 서명해주세요."); return; }
    const isContract = !contract.docType || contract.docType === "contract";
    if (isContract && !signAddr.trim()) { alert("주소를 입력해주세요."); return; }
    if (isContract && !signPhone.trim()) { alert("전화번호를 입력해주세요."); return; }
    const docLabel = DOC_TYPES.find(d => d.key === (contract.docType || "contract"))?.label || "문서";
    if (!window.confirm(`${docLabel}에 전자서명(동의)하시겠습니까?`)) return;
    setSigning(true);
    try {
      let signIp = "";
      try {
        const ipRes = await fetch("https://api.ipify.org?format=json");
        const ipData = await ipRes.json();
        signIp = ipData.ip || "";
      } catch { signIp = "알 수 없음"; }

      const signMeta = {
        signedAt: new Date().toISOString(),
        signIp,
        signUserAgent: navigator.userAgent,
        signDevice: /Mobi|Android/i.test(navigator.userAgent) ? "모바일" : "PC",
        signBrowser: navigator.userAgent.match(/(Chrome|Safari|Firefox|Edge|Samsung)/)?.[1] || "기타",
      };

      await setDoc(doc(db, COL_CONTRACTS, contract.id), {
        ...contract,
        status: "signed",
        ...(isContract ? { empAddress: signAddr.trim(), empPhone: signPhone.trim() } : {}),
        ...signMeta,
      });
      await addDoc(collection(db, COL_NOTICES), {
        title: `✅ ${user.name}님 ${docLabel} 서명 완료`,
        content: `${user.name}님이 ${docLabel}에 전자서명(동의)하였습니다.\n서명일시: ${new Date().toLocaleString("ko-KR")}\nIP: ${signIp}`,
        recipient: "admin",
        author: user.name, auto: true,
        createdAt: new Date().toISOString(),
      });
      await sendPush({ title: `✅ ${docLabel} 서명 완료`, message: `${user.name}님이 ${docLabel}에 서명하였습니다.`, targetUserId: "admin" });
      alert("서명이 완료되었습니다.");
    } catch(e) { alert("서명 실패: " + e.message); }
    setSigning(false);
  };

  if (!contract && (docTypeTab === "contract" || myDocs.length === 0)) return (
    <div style={{ padding: 32, textAlign: "center" }}>
      <div style={{ fontSize: 48, marginBottom: 12 }}>📄</div>
      <div style={{ fontSize: 15, color: T.muted, fontWeight: 600 }}>등록된 {DOC_TYPES.find(d=>d.key===docTypeTab)?.label||"문서"}가 없습니다</div>
      <div style={{ fontSize: 13, color: T.muted, marginTop: 6 }}>관리자가 발송하면 여기서 확인할 수 있어요</div>
    </div>
  );

  const st = contract.status === "signed" ? { text: "✅ 서명완료", color: "#16a34a", bg: "#dcfce7" }
    : contract.status === "sent" ? { text: "📨 서명 대기", color: "#d97706", bg: "#fef3c7" }
    : { text: "📝 초안", color: "#6b7280", bg: "#f3f4f6" };

  const Row = ({ label, value }) => value ? (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", padding: "8px 0", borderBottom: `1px solid ${T.border}` }}>
      <span style={{ fontSize: 13, color: T.muted, minWidth: 90 }}>{label}</span>
      <span style={{ fontSize: 13, fontWeight: 600, color: T.text, textAlign: "right", flex: 1 }}>{value}</span>
    </div>
  ) : null;

  return (
    <div style={{ padding: "0 0 30px" }}>
      {/* 문서 종류 탭 — 두 줄 */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6, padding: "12px 16px 0" }}>
        {DOC_TYPES.map(dt => {
          const hasPending = contracts.some(c => c.userId === user.id && (c.docType || "contract") === dt.key && c.status === "sent");
          return (
            <button key={dt.key} onClick={() => { setDocTypeTab(dt.key); setConfirmed(false); setSignAddr(""); setSignPhone(""); }}
              style={{ padding: "8px 0", borderRadius: 10, border: `2px solid ${docTypeTab === dt.key ? T.adminHeader : T.border}`, background: docTypeTab === dt.key ? T.adminHeader : T.bg, color: docTypeTab === dt.key ? "#fff" : T.text, fontSize: 12, fontWeight: 700, cursor: "pointer", position: "relative" }}>
              {dt.icon} {dt.label}
              {hasPending && <span style={{ position: "absolute", top: -4, right: -4, width: 8, height: 8, borderRadius: "50%", background: "#ef4444" }} />}
            </button>
          );
        })}
      </div>

      {/* 동의서/확인서/기타 탭 — 그룹핑해서 전부 표시 */}
      {docTypeTab !== "contract" && docTypeTab !== "education" && myDocs.length > 0 && (() => {
        const groups = [];
        const seen = {};
        myDocs.forEach(c => {
          const key = c.docTitle || "__no_title__";
          if (!seen[key]) { seen[key] = []; groups.push(seen[key]); }
          seen[key].push(c);
        });
        return (
          <div style={{ padding: "12px 16px 0" }}>
            {groups.map((group, gi) => {
              const latest = group[0];
              const history = group.slice(1);
              const dst = latest.status === "signed"
                ? { text: "✅ 서명완료", color: "#16a34a", bg: "#dcfce7" }
                : latest.status === "sent"
                ? { text: "📨 서명 대기", color: "#d97706", bg: "#fef3c7" }
                : { text: "📝 초안", color: "#6b7280", bg: "#f3f4f6" };
              const dk = `mdoc_${latest.id}`;
              const sd = expandedDoc[dk];
              const histKey = `mhist_${gi}_${latest.docTitle}`;
              const showHist = expandedDoc[histKey];
              return (
                <div key={latest.id} style={{ background: T.card, borderRadius: 14, border: `1px solid ${T.border}`, marginBottom: 10, overflow: "hidden" }}>
                  {/* 헤더 */}
                  <div style={{ padding: "12px 14px", display: "flex", justifyContent: "space-between", alignItems: "center" }}
                    onClick={() => setExpandedDoc(p => ({ ...p, [dk]: !p[dk] }))}>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: 14, fontWeight: 800, color: T.text }}>{latest.docTitle || ""}</div>
                      <div style={{ fontSize: 11, color: T.muted, marginTop: 2 }}>
                        {latest.createdAt ? new Date(latest.createdAt).toLocaleDateString("ko-KR") : ""}
                        {latest.signedAt ? ` · 서명 ${new Date(latest.signedAt).toLocaleDateString("ko-KR")}` : ""}
                      </div>
                    </div>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <span style={{ fontSize: 11, fontWeight: 700, color: dst.color, background: dst.bg, borderRadius: 8, padding: "3px 8px" }}>{dst.text}</span>
                      <span style={{ fontSize: 14, color: T.muted }}>{sd ? "▲" : "▼"}</span>
                    </div>
                  </div>
                  {/* 내용 + 서명 */}
                  {sd && (
                    <div style={{ borderTop: `1px solid ${T.border}`, padding: "12px 14px" }}>
                      {latest.docContent && (
                        <div style={{ fontSize: 13, color: T.text, lineHeight: 1.8, whiteSpace: "pre-wrap", marginBottom: 12 }}>
                          {latest.docContent}
                        </div>
                      )}
                      {/* 서명 대기 */}
                      {latest.status === "sent" && (() => {
                        const needsFile = latest.docType === "confirm" || latest.docType === "other";
                        const fileKey = `file_${latest.id}`;
                        const attachedFile = expandedDoc[fileKey];
                        const canSign = !!expandedDoc[`agree_${latest.id}`];
                        return (
                          <div style={{ background: T.card, borderRadius: 12, padding: 14, border: `2px solid #0891b2`, marginTop: 8 }}>
                            <div style={{ fontSize: 13, fontWeight: 800, color: "#0891b2", marginBottom: 10 }}>✍️ 전자서명 (동의)</div>
                            <div style={{ marginBottom: 10, padding: "8px 12px", background: T.bg, borderRadius: 8, fontSize: 13, fontWeight: 700, color: T.text }}>{user.name}</div>

                              {/* 파일 첨부 — 확인서/기타만 */}
                              {needsFile && (
                                <div style={{ marginBottom: 12 }}>
                                  <div style={{ fontSize: 12, color: T.sub, fontWeight: 600, marginBottom: 6 }}>
                                    📎 파일 첨부 <span style={{ color: T.muted, fontWeight: 400 }}>(수료증 등 선택사항)</span>
                                  </div>
                                  {attachedFile ? (
                                    <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 12px", background: T.bg, borderRadius: 8, border: `1px solid ${T.border}` }}>
                                      <span style={{ fontSize: 18 }}>{attachedFile.type?.startsWith("image/") ? "🖼" : "📄"}</span>
                                      <span style={{ fontSize: 12, fontWeight: 600, color: T.text, flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{attachedFile.name}</span>
                                      <button onClick={() => setExpandedDoc(p => ({ ...p, [fileKey]: null }))}
                                        style={{ background: "none", border: "none", color: "#b91c1c", cursor: "pointer", fontSize: 15 }}>✕</button>
                                    </div>
                                  ) : (
                                    <div style={{ display: "flex", gap: 8 }}>
                                      <label style={{ flex: 1, padding: "10px 0", borderRadius: 10, border: `1px dashed ${T.border}`, background: T.bg, color: T.sub, fontSize: 12, fontWeight: 600, cursor: "pointer", textAlign: "center" }}>
                                        📷 사진
                                        <input type="file" accept="image/*" style={{ display: "none" }}
                                          onChange={e => { const f = e.target.files?.[0]; if (f) setExpandedDoc(p => ({ ...p, [fileKey]: f })); }} />
                                      </label>
                                      <label style={{ flex: 1, padding: "10px 0", borderRadius: 10, border: `1px dashed ${T.border}`, background: T.bg, color: T.sub, fontSize: 12, fontWeight: 600, cursor: "pointer", textAlign: "center" }}>
                                        📁 파일
                                        <input type="file" accept=".pdf,.doc,.docx,.jpg,.jpeg,.png,.gif,.webp,image/*,application/pdf" style={{ display: "none" }}
                                          onChange={e => { const f = e.target.files?.[0]; if (f) setExpandedDoc(p => ({ ...p, [fileKey]: f })); }} />
                                      </label>
                                    </div>
                                  )}
                                </div>
                              )}

                            <div style={{ display: "flex", alignItems: "flex-start", gap: 10, marginBottom: 12, padding: 10, background: "#f0f9ff", borderRadius: 8 }}>
                              <input type="checkbox" id={`agree_${latest.id}`}
                                checked={!!expandedDoc[`agree_${latest.id}`]}
                                onChange={e => setExpandedDoc(p => ({ ...p, [`agree_${latest.id}`]: e.target.checked }))}
                                style={{ width: 18, height: 18, marginTop: 1, cursor: "pointer" }} />
                              <label htmlFor={`agree_${latest.id}`} style={{ fontSize: 13, color: T.text, lineHeight: 1.6, cursor: "pointer" }}>
                                본인은 위 {latest.docTitle || "문서"}의 내용을 충분히 읽고 이해하였으며, 이에 동의합니다.
                              </label>
                            </div>
                            <button
                              onClick={async () => {
                                if (!canSign) { alert("동의 체크 후 서명해주세요."); return; }
                                if (!window.confirm(`${latest.docTitle || "문서"}에 전자서명(동의)하시겠습니까?`)) return;
                                try {
                                  let signIp = "";
                                  try { const r = await fetch("https://api.ipify.org?format=json"); signIp = (await r.json()).ip || ""; } catch {}
                                  // 파일 업로드 (확인서/기타)
                                  let attachmentData = null;
                                  if (needsFile && attachedFile) {
                                    const ext = attachedFile.name.split(".").pop().toLowerCase();
                                    const sRef = ref(storage, `doc_attachments/${latest.userId}/${latest.id}_${Date.now()}.${ext}`);
                                    await uploadBytes(sRef, attachedFile);
                                    const url = await getDownloadURL(sRef);
                                    attachmentData = { url, name: attachedFile.name, type: attachedFile.type, size: attachedFile.size };
                                  }
                                  await setDoc(doc(db, COL_CONTRACTS, latest.id), {
                                    ...latest, status: "signed",
                                    signedAt: new Date().toISOString(), signIp,
                                    signUserAgent: navigator.userAgent,
                                    signDevice: /Mobi|Android/i.test(navigator.userAgent) ? "모바일" : "PC",
                                    signBrowser: navigator.userAgent.match(/(Chrome|Safari|Firefox|Edge|Samsung)/)?.[1] || "기타",
                                    ...(attachmentData ? { attachment: attachmentData } : {}),
                                  });
                                  await addDoc(collection(db, COL_NOTICES), {
                                    title: `✅ ${user.name}님 ${latest.docTitle || "문서"} 서명 완료`,
                                    content: `${user.name}님이 ${latest.docTitle || "문서"}에 전자서명하였습니다.${attachmentData ? "\n📎 첨부파일: " + attachmentData.name : ""}`,
                                    recipient: "admin", author: user.name, auto: true, createdAt: new Date().toISOString(),
                                  });
                                  await sendPush({ title: `✅ ${latest.docTitle || "문서"} 서명 완료`, message: `${user.name}님이 서명하였습니다.${attachmentData ? " (첨부파일 있음)" : ""}`, targetUserId: "admin" });
                                  alert("서명이 완료되었습니다.");
                                } catch(e) { alert("서명 실패: " + e.message); }
                              }}
                              disabled={!canSign}
                              style={{ width: "100%", padding: "12px 0", borderRadius: 12, border: "none",
                                background: canSign ? "#0891b2" : "#e5e7eb",
                                color: canSign ? "#fff" : T.muted,
                                fontSize: 14, fontWeight: 800, cursor: canSign ? "pointer" : "default" }}>
                              📝 서명하기 (동의)
                            </button>
                          </div>
                        );
                      })()}
                      {/* 서명 완료 */}
                      {latest.status === "signed" && (
                        <div style={{ padding: 12, background: "#dcfce7", borderRadius: 10, textAlign: "center" }}>
                          <div style={{ fontSize: 24, marginBottom: 4 }}>✅</div>
                          <div style={{ fontSize: 13, fontWeight: 800, color: "#15803d" }}>서명 완료</div>
                          <div style={{ fontSize: 11, color: "#16a34a", marginTop: 2 }}>
                            {latest.signedAt && new Date(latest.signedAt).toLocaleString("ko-KR")}
                          </div>
                          {latest.attachment && (
                            <a href={latest.attachment.url} target="_blank" rel="noreferrer"
                              style={{ display: "inline-flex", alignItems: "center", gap: 6, marginTop: 8, padding: "6px 14px", borderRadius: 8, background: "#fff", color: "#15803d", fontSize: 12, fontWeight: 700, textDecoration: "none", border: "1px solid #86efac" }}>
                              📎 {latest.attachment.name}
                            </a>
                          )}
                          <button onClick={async () => {
                            const el = document.getElementById(`doc-print-${latest.id}`);
                            if (!el) { alert("PDF 생성 영역을 찾을 수 없습니다."); return; }
                            try {
                              const canvas = await html2canvas(el, { scale: 2.5, useCORS: true, backgroundColor: "#ffffff" });
                              const imgData = canvas.toDataURL("image/png");
                              const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
                              const pageW = pdf.internal.pageSize.getWidth();
                              const pageH = pdf.internal.pageSize.getHeight();
                              const margin = 8;
                              const ratio = canvas.width / canvas.height;
                              let imgW = pageW - margin * 2, imgH = imgW / ratio;
                              if (imgH > pageH - margin * 2) { imgH = pageH - margin * 2; imgW = imgH * ratio; }
                              pdf.addImage(imgData, "PNG", (pageW - imgW) / 2, margin, imgW, imgH);
                              pdf.save(`${user.name}_${latest.docTitle || "문서"}_${latest.createdAt?.slice(0,10) || ""}.pdf`);
                            } catch(e) { alert("PDF 생성 실패: " + e.message); }
                          }}
                            style={{ display: "block", width: "100%", marginTop: 8, padding: "8px 0", borderRadius: 10, border: "none", background: "#16a34a", color: "#fff", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
                            ⬇ PDF
                          </button>
                        </div>
                      )}
                      {/* 동의서/확인서 숨김 인쇄 영역 */}
                      <div id={`doc-print-${latest.id}`} style={{ position: "fixed", left: -9999, top: 0, width: 794, background: "#fff", padding: "40px 50px", fontFamily: "'Noto Sans KR', sans-serif", fontSize: 12, color: "#111", lineHeight: 1.8 }}>
                        <div style={{ textAlign: "center", fontSize: 20, fontWeight: 900, marginBottom: 8, letterSpacing: 4 }}>
                          {DOC_TYPES.find(d => d.key === latest.docType)?.label || "문서"}
                        </div>
                        {latest.docTitle && <div style={{ textAlign: "center", fontSize: 14, fontWeight: 700, marginBottom: 24, color: "#444" }}>{latest.docTitle}</div>}
                        <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: 20, fontSize: 12 }}>
                          <tbody>
                            <tr><td style={{ padding: "4px 8px", fontWeight: 700, width: 80 }}>사업체명</td><td style={{ padding: "4px 8px" }}>하나기업</td></tr>
                            <tr><td style={{ padding: "4px 8px", fontWeight: 700 }}>대표자</td><td style={{ padding: "4px 8px" }}>박용균</td></tr>
                            <tr><td style={{ padding: "4px 8px", fontWeight: 700 }}>성명</td><td style={{ padding: "4px 8px" }}>{user.name}</td></tr>
                          </tbody>
                        </table>
                        {latest.docContent && (
                          <div style={{ fontSize: 12, lineHeight: 1.9, whiteSpace: "pre-wrap", marginBottom: 24, padding: "12px 0", borderTop: "1px solid #e5e7eb", borderBottom: "1px solid #e5e7eb" }}>
                            {latest.docContent}
                          </div>
                        )}
                        <div style={{ marginTop: 24, fontSize: 12 }}>{latest.createdAt ? new Date(latest.createdAt).toLocaleDateString("ko-KR", { year: "numeric", month: "long", day: "numeric" }) : ""}</div>
                        <div style={{ marginTop: 16, display: "flex", justifyContent: "space-around", fontSize: 12 }}>
                          <div>(사업자) 박용균</div>
                          <div>(서명인) {user.name} &nbsp;&nbsp; {latest.signedAt ? `✅ ${new Date(latest.signedAt).toLocaleDateString("ko-KR")} 전자서명` : ""}</div>
                        </div>
                      </div>
                    </div>
                  )}
                  {/* 이전 버전 히스토리 */}
                  {history.length > 0 && (
                    <div style={{ borderTop: `1px solid ${T.border}`, padding: "8px 14px", background: "#fafafa" }}>
                      <button onClick={() => setExpandedDoc(p => ({ ...p, [histKey]: !p[histKey] }))}
                        style={{ width: "100%", padding: "4px 0", background: "transparent", border: "none", color: T.muted, fontSize: 11, fontWeight: 700, cursor: "pointer", textAlign: "left" }}>
                        📁 이전 버전 {history.length}건 {showHist ? "▲ 접기" : "▼ 보기"}
                      </button>
                      {showHist && (
                        <div style={{ marginTop: 6, display: "flex", flexDirection: "column", gap: 4 }}>
                          {history.map(c => {
                            const hdk = `mhdet_${c.id}`;
                            const hsd = expandedDoc[hdk];
                            return (
                              <div key={c.id} style={{ background: "#fff", borderRadius: 8, border: `1px solid ${T.border}`, overflow: "hidden" }}>
                                <div style={{ padding: "8px 10px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                                  <div>
                                    <div style={{ fontSize: 11, color: T.muted }}>
                                      작성: {c.createdAt ? new Date(c.createdAt).toLocaleDateString("ko-KR") : "-"}
                                      {c.signedAt ? ` · 서명: ${new Date(c.signedAt).toLocaleDateString("ko-KR")}` : ""}
                                    </div>
                                    {c.signedAt && <div style={{ fontSize: 11, color: "#15803d", marginTop: 2 }}>✅ 서명완료</div>}
                                  </div>
                                  <button onClick={() => setExpandedDoc(p => ({ ...p, [hdk]: !p[hdk] }))}
                                    style={{ padding: "3px 7px", borderRadius: 5, border: `1px solid ${T.border}`, background: "#fff", color: T.text, fontSize: 10, fontWeight: 700, cursor: "pointer", flexShrink: 0 }}>
                                    {hsd ? "▲" : "▼"}
                                  </button>
                                </div>
                                {hsd && (
                                  <div style={{ borderTop: `1px solid ${T.border}`, padding: "8px 10px" }}>
                                    {c.docContent && (
                                      <div style={{ fontSize: 11, color: T.text, lineHeight: 1.6, whiteSpace: "pre-wrap", marginBottom: 6 }}>{c.docContent}</div>
                                    )}
                                    {c.signedAt && (
                                      <div style={{ padding: "5px 8px", background: "#dcfce7", borderRadius: 6 }}>
                                        <div style={{ fontSize: 10, color: "#15803d", fontWeight: 700 }}>✅ {new Date(c.signedAt).toLocaleString("ko-KR")}</div>
                                        {c.signIp && <div style={{ fontSize: 10, color: "#16a34a" }}>🌐 {c.signIp} · {c.signDevice}</div>}
                                      </div>
                                    )}
                                  </div>
                                )}
                              </div>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        );
      })()}

      {/* 상태 배너 ~ 서명/PDF — 근로계약서 탭만 */}
      {docTypeTab === "contract" && contract && (<>
      {/* 상태 배너 */}
      <div style={{ margin: 16, padding: "14px 16px", borderRadius: 14, background: st.bg, border: `1px solid ${st.color}30`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div>
          <div style={{ fontSize: 14, fontWeight: 800, color: st.color }}>{st.text}</div>
          {contract.status === "signed" && contract.signedAt && (
            <div style={{ fontSize: 11, color: T.muted, marginTop: 2 }}>
              서명일시: {new Date(contract.signedAt).toLocaleString("ko-KR")}
            </div>
          )}
          {contract.status === "sent" && (
            <div style={{ fontSize: 11, color: T.muted, marginTop: 2 }}>내용을 확인하고 서명해주세요</div>
          )}
        </div>
        {(!contract.docType || contract.docType === "contract") && (
          <button onClick={() => setShowDetail(!showDetail)}
            style={{ padding: "7px 14px", borderRadius: 10, border: "none", background: "#fff", color: T.text, fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
            {showDetail ? "접기 ▲" : "상세보기 ▼"}
          </button>
        )}
      </div>

      {/* 근로계약서일 때만 주요내용/상세보기/PDF 표시 */}
      {(!contract.docType || contract.docType === "contract") && (<>
      {/* 계약서 요약 */}
      <div style={{ margin: "0 16px 12px", background: T.card, borderRadius: 16, padding: 16, border: `1px solid ${T.border}` }}>
        <div style={{ fontSize: 13, fontWeight: 800, color: T.text, marginBottom: 12 }}>📋 계약 주요 내용</div>
        <Row label="사업체명" value={contract.companyName} />
        <Row label="직종" value={contract.jobType} />
        <Row label="입사일(근로개시일)" value={contract.joinDate} />
        {contract.contractStart && contract.joinDate && (
          <div style={{ padding: "8px 0", borderBottom: `1px solid ${T.border}` }}>
            <div style={{ fontSize: 12, color: "#0369a1", lineHeight: 1.8 }}>
              본 근로계약은 <b>{contract.contractStart}</b>에 체결되었으며, 근로개시일은 <b>{contract.joinDate}</b>로 한다.
              {contract.contractStart !== contract.joinDate && (
                <><br />본 계약의 내용은 입사일부터 적용하며, 기존 근로계약을 본 계약으로 대체한다.</>
              )}
            </div>
            <div style={{ fontSize: 10, color: "#0891b2", marginTop: 4, fontWeight: 700 }}>
              {contract.contractStart === contract.joinDate ? "✅ 신규 입사 계약" : "🔄 재계약 / 계약 변경"}
            </div>
          </div>
        )}
        <Row label="계약 시작" value={contract.contractStart} />
        <Row label="계약 종료" value={contract.contractEnd || "기간 없음 (정규직)"} />
        <Row label="근로시간" value={`${contract.workStart} ~ ${contract.workEnd} (1일 ${contract.dailyHours || 8}시간, ${contract.weekDays || "월~금"})`} />
        <Row label="시급" value={contract.hourlyWage ? `${Number(contract.hourlyWage).toLocaleString()}원` : null} />
        <Row label="월급" value={contract.monthlyWage ? `${Number(contract.monthlyWage).toLocaleString()}원 (주휴수당 포함)` : null} />
        <Row label="임금 지급일" value={contract.payDay ? `매월 ${contract.payDay}일 (${contract.payHoliday || ""})` : null} />
      </div>

      {/* 상세 내용 펼치기 */}
      {showDetail && (
        <div style={{ margin: "0 16px 12px" }}>
          {/* 임금 구성 */}
          <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 10, border: `1px solid ${T.border}` }}>
            <div style={{ fontSize: 13, fontWeight: 800, color: "#ea580c", marginBottom: 10 }}>💰 임금 구성</div>
            {contract.wage1 && <Row label="기본급" value={contract.wage1} />}
            {contract.wage2 && <Row label="연차수당" value={contract.wage2} />}
            {contract.wage3 && <Row label="잔업수당" value={contract.wage3} />}
            {contract.wage4 && <Row label="특근수당" value={contract.wage4} />}
            {contract.wage5 && <Row label="상여금" value={contract.wage5} />}
            <div style={{ marginTop: 8, padding: "10px 12px", background: "#fff7ed", borderRadius: 10, border: "1px solid #fed7aa" }}>
              <div style={{ fontSize: 11, color: "#92400e", lineHeight: 1.7 }}>
                📌 임금은 관계 법령에 따른 최저임금 변동 및 당사자 간 합의에 의해 변경될 수 있으며, 변경 시 사전 서면 통보로 본 계약의 해당 조항을 갈음한다. 단, 임금의 감액은 근로자의 서면 동의를 요한다.
              </div>
            </div>
          </div>
          {/* 휴게 / 지급 */}
          <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 10, border: `1px solid ${T.border}` }}>
            <div style={{ fontSize: 13, fontWeight: 800, color: "#16a34a", marginBottom: 10 }}>⏰ 휴게시간 · 지급</div>
            {contract.breakLunch && <Row label="식사시간" value={contract.breakLunch} />}
            {contract.breakSnack && <Row label="참시간" value={contract.breakSnack} />}
            <Row label="임금 계산기간" value={contract.payCalcPeriod} />
            <Row label="지급 방법" value={contract.payMethod} />
            {contract.bankName && <Row label="은행" value={`${contract.bankName} ${contract.bankAccount || ""}`} />}
          </div>
          {/* 휴가/보험/복지 */}
          <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 10, border: `1px solid ${T.border}` }}>
            <div style={{ fontSize: 13, fontWeight: 800, color: "#0284c7", marginBottom: 10 }}>📅 휴가 · 보험 · 복지</div>
            <Row label="연차유급휴가" value={contract.annualLeave} />
            <Row label="4대보험" value={contract.insurance} />
            <Row label="복리후생" value={contract.welfare} />
          </div>
          {/* 퇴직/정년 */}
          <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 10, border: `1px solid ${T.border}` }}>
            <div style={{ fontSize: 13, fontWeight: 800, color: "#b45309", marginBottom: 10 }}>💼 퇴직 · 정년</div>
            <Row label="퇴직금" value={contract.severancePay} />
            <Row label="퇴직 절차" value={contract.resignNotice} />
            <Row label="정년" value={contract.retirementAge} />
          </div>
          {/* 해지사유 */}
          {contract.terminationReasons && (
            <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 10, border: `1px solid ${T.border}` }}>
              <div style={{ fontSize: 13, fontWeight: 800, color: "#dc2626", marginBottom: 10 }}>⚠️ 근로계약 해지사유</div>
              <div style={{ fontSize: 13, color: T.text, whiteSpace: "pre-wrap", lineHeight: 1.8 }}>{contract.terminationReasons}</div>
            </div>
          )}
          {/* 기타 */}
          {(contract.bonus || contract.specialTerms) && (
            <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 10, border: `1px solid ${T.border}` }}>
              <div style={{ fontSize: 13, fontWeight: 800, color: T.text, marginBottom: 10 }}>📝 기타</div>
              {contract.bonus && <Row label="상여금 지급시기" value={contract.bonus} />}
              {contract.specialTerms && (
                <div style={{ marginTop: 8, padding: 10, background: T.bg, borderRadius: 8 }}>
                  <div style={{ fontSize: 12, color: T.muted, fontWeight: 600, marginBottom: 4 }}>특약 사항</div>
                  <div style={{ fontSize: 13, color: T.text, whiteSpace: "pre-wrap" }}>{contract.specialTerms}</div>
                </div>
              )}
            </div>
          )}
          {/* 법적 고지 */}
          <div style={{ padding: 14, background: "#f0fdf4", borderRadius: 14, border: "1px solid #bbf7d0", marginBottom: 10 }}>
            <div style={{ fontSize: 11, color: "#15803d", lineHeight: 1.7 }}>
              사용자는 근로계약을 체결함과 동시에 본 계약서를 사본하여 근로자의 교부요구와 관계없이 서면(전자문서 포함)으로 작성하여 근로자에게 교부하여야 함(근로기준법 제17조 이행)<br />
              이 계약에 정함이 없는 사항은 근로기준법에 의함
            </div>
          </div>
        </div>
      )}
      </>)}

      {/* 서명 영역 — 모든 문서 공통 */}
      {contract.status === "sent" && (() => {
        const isContract = !contract.docType || contract.docType === "contract";
        const docLabel = DOC_TYPES.find(d => d.key === (contract.docType || "contract"))?.label || "문서";
        const canSign = confirmed && (!isContract || (signAddr && signPhone));
        return (
          <div style={{ margin: "0 16px", background: T.card, borderRadius: 16, padding: 16, border: `2px solid #0891b2` }}>
            <div style={{ fontSize: 14, fontWeight: 800, color: "#0891b2", marginBottom: 12 }}>✍️ 전자서명 (동의)</div>
            <div style={{ marginBottom: 12 }}>
              <div style={{ fontSize: 12, color: T.sub, fontWeight: 600, marginBottom: 4 }}>성명</div>
              <div style={{ padding: "10px 12px", borderRadius: 10, border: `1px solid ${T.border}`, fontSize: 13, fontWeight: 700, color: T.text, background: "#f8fafc" }}>{user.name}</div>
            </div>

            {/* 주소/전화 — 근로계약서만 */}
            {isContract && (<>
              <div style={{ marginBottom: 12 }}>
                <div style={{ fontSize: 12, color: T.sub, fontWeight: 600, marginBottom: 4 }}>주소 <span style={{ color: "#dc2626" }}>*</span></div>
                <input type="text" value={signAddr} onChange={e => setSignAddr(e.target.value)}
                  placeholder="주소를 입력하세요"
                  style={{ width: "100%", padding: "10px 12px", borderRadius: 10, border: `1px solid ${signAddr ? T.border : "#fca5a5"}`, fontSize: 13, fontWeight: 600, color: T.text, background: "#fff", boxSizing: "border-box", fontFamily: "inherit" }} />
              </div>
              <div style={{ marginBottom: 16 }}>
                <div style={{ fontSize: 12, color: T.sub, fontWeight: 600, marginBottom: 4 }}>전화번호 <span style={{ color: "#dc2626" }}>*</span></div>
                <input type="tel" value={signPhone} onChange={e => {
                  const digits = e.target.value.replace(/\D/g, "").slice(0, 11);
                  let formatted = digits;
                  if (digits.length <= 3) formatted = digits;
                  else if (digits.length <= 7) formatted = `${digits.slice(0,3)}-${digits.slice(3)}`;
                  else formatted = `${digits.slice(0,3)}-${digits.slice(3,7)}-${digits.slice(7)}`;
                  setSignPhone(formatted);
                }}
                  placeholder="010-0000-0000"
                  style={{ width: "100%", padding: "10px 12px", borderRadius: 10, border: `1px solid ${signPhone ? T.border : "#fca5a5"}`, fontSize: 13, fontWeight: 600, color: T.text, background: "#fff", boxSizing: "border-box", fontFamily: "inherit" }} />
              </div>
            </>)}

            {/* 동의 체크 */}
            <div style={{ display: "flex", alignItems: "flex-start", gap: 10, marginBottom: 16, padding: 12, background: "#f0f9ff", borderRadius: 10 }}>
              <input type="checkbox" id="agreeCheck" checked={confirmed} onChange={e => setConfirmed(e.target.checked)}
                style={{ width: 18, height: 18, marginTop: 1, cursor: "pointer" }} />
              <label htmlFor="agreeCheck" style={{ fontSize: 13, color: T.text, lineHeight: 1.6, cursor: "pointer" }}>
                본인은 위 {docLabel}의 내용을 충분히 읽고 이해하였으며, 이에 동의합니다.
              </label>
            </div>
            <button onClick={sign} disabled={signing || !canSign}
              style={{ width: "100%", padding: "14px 0", borderRadius: 14, border: "none",
                background: canSign ? "#0891b2" : "#e5e7eb",
                color: canSign ? "#fff" : T.muted,
                fontSize: 15, fontWeight: 800, cursor: canSign ? "pointer" : "default" }}>
              {signing ? "처리 중..." : "📝 서명하기 (동의)"}
            </button>
            <div style={{ fontSize: 11, color: T.muted, textAlign: "center", marginTop: 8 }}>
              서명 후에는 취소할 수 없습니다
            </div>
          </div>
        );
      })()}

      {contract.status === "signed" && (
        <div style={{ margin: "0 16px", padding: 16, background: "#dcfce7", borderRadius: 16, textAlign: "center" }}>
          <div style={{ fontSize: 32, marginBottom: 6 }}>✅</div>
          <div style={{ fontSize: 14, fontWeight: 800, color: "#15803d" }}>서명 완료</div>
          <div style={{ fontSize: 12, color: "#16a34a", marginTop: 4 }}>
            {contract.signedAt && new Date(contract.signedAt).toLocaleString("ko-KR")}
          </div>
          <button id="member-pdf-btn" onClick={downloadMyPDF} disabled={pdfLoading}
            style={{ marginTop: 12, padding: "10px 24px", borderRadius: 12, border: "none", background: "#16a34a", color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer", opacity: pdfLoading ? 0.6 : 1 }}>
            {pdfLoading ? "생성 중..." : "⬇ PDF 다운로드"}
          </button>
        </div>
      )}

      {/* 팀원용 숨김 인쇄 영역 — 모든 문서 공통 */}
      {contract && (
        <div id={`contract-member-print-${contract.id}`} style={{ position: "fixed", left: -9999, top: 0, width: 794, background: "#fff", padding: "40px 50px", fontFamily: "'Noto Sans KR', sans-serif", fontSize: 12, color: "#111", lineHeight: 1.8 }}>
          {(!contract.docType || contract.docType === "contract") && (<>
            <div style={{ textAlign: "center", fontSize: 20, fontWeight: 900, marginBottom: 24, letterSpacing: 8 }}>근 로 계 약 서</div>
            <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: 16, fontSize: 12 }}>
              <tbody>
                <tr><td style={{ padding: "4px 8px", fontWeight: 700, width: 80 }}>(갑) 사용자</td><td style={{ padding: "4px 8px" }}>사업체명: {contract.companyName} &nbsp;&nbsp; 대표자: {contract.ownerName}</td></tr>
                <tr><td style={{ padding: "4px 8px" }}></td><td style={{ padding: "4px 8px" }}>주소: {contract.bizAddress}</td></tr>
                <tr><td style={{ padding: "4px 8px", fontWeight: 700 }}>(을) 근로자</td><td style={{ padding: "4px 8px" }}>성명: {contract.userName} &nbsp;&nbsp; 연락처: {contract.empPhone || "__________"}</td></tr>
                <tr><td style={{ padding: "4px 8px" }}></td><td style={{ padding: "4px 8px" }}>주소: {contract.empAddress || "__________________________________________"}</td></tr>
              </tbody>
            </table>
            <div style={{ fontSize: 11, marginBottom: 16 }}>위 당사자는 아래의 근로조건을 성실히 이행할 것을 약정하고 근로계약을 체결한다.</div>
            {contract.contractStart && contract.joinDate && (
              <div style={{ fontSize: 11, marginBottom: 12, padding: "6px 10px", background: "#f0f9ff", borderRadius: 4 }}>
                본 근로계약은 {contract.contractStart}에 체결되었으며, 근로개시일은 {contract.joinDate}로 한다.
                {contract.contractStart !== contract.joinDate && " 본 계약의 내용은 입사일부터 적용하며, 기존 근로계약을 본 계약으로 대체한다."}
              </div>
            )}
            {[
              ["근로 장소", contract.workPlace], ["직종", contract.jobType],
              ["계약 기간", `${contract.contractStart} ~ ${contract.contractEnd || "기간 없음 (정규직)"}`],
              ["근로 시간", `${contract.workStart} ~ ${contract.workEnd} (1일 ${contract.dailyHours || 8}시간, ${contract.weekDays || "월~금"})`],
              ["휴게 시간", `식사 ${contract.breakLunch || ""}  참 ${contract.breakSnack || ""}`],
              ["시급", contract.hourlyWage ? `${Number(contract.hourlyWage).toLocaleString()}원` : ""],
              ["월급", contract.monthlyWage ? `${Number(contract.monthlyWage).toLocaleString()}원 (주휴수당 포함)` : ""],
              ["연차수당", contract.wage2], ["잔업수당", contract.wage3], ["특근수당", contract.wage4], ["상여금", contract.wage5],
              ["임금 계산기간", contract.payCalcPeriod],
              ["임금 지급일", contract.payDay ? `매월 ${contract.payDay}일 (${contract.payHoliday || ""})` : ""],
              ["지급 방법", `${contract.payMethod || ""} ${contract.bankName ? `/ ${contract.bankName} ${contract.bankAccount || ""}` : ""}`],
              ["연차유급휴가", contract.annualLeave], ["4대보험", contract.insurance], ["복리후생", contract.welfare],
              ["퇴직금", contract.severancePay], ["퇴직 절차", contract.resignNotice], ["정년", contract.retirementAge],
              ["상여금 지급시기", contract.bonus],
            ].filter(([, v]) => v).map(([label, value]) => (
              <div key={label} style={{ display: "flex", borderBottom: "1px solid #e5e7eb", padding: "3px 0" }}>
                <span style={{ minWidth: 100, fontWeight: 700, fontSize: 11 }}>{label}</span>
                <span style={{ fontSize: 11, flex: 1 }}>{value}</span>
              </div>
            ))}
            {contract.terminationReasons && (
              <div style={{ marginTop: 10 }}>
                <div style={{ fontWeight: 700, fontSize: 11, marginBottom: 4 }}>근로계약 해지사유</div>
                <div style={{ fontSize: 10, whiteSpace: "pre-wrap", lineHeight: 1.7 }}>{contract.terminationReasons}</div>
              </div>
            )}
            <div style={{ marginTop: 10, padding: "6px 10px", background: "#fff7ed", borderRadius: 4, fontSize: 10, color: "#92400e" }}>
              📌 임금은 관계 법령에 따른 최저임금 변동 및 당사자 간 합의에 의해 변경될 수 있으며, 변경 시 사전 서면 통보로 본 계약의 해당 조항을 갈음한다. 단, 임금의 감액은 근로자의 서면 동의를 요한다.
            </div>
            <div style={{ marginTop: 8, fontSize: 10, color: "#555" }}>이 계약에 정함이 없는 사항은 근로기준법에 의함</div>
            {contract.specialTerms && <div style={{ marginTop: 6, fontSize: 10 }}><b>특약:</b> {contract.specialTerms}</div>}
            <div style={{ marginTop: 24, fontSize: 12 }}>{contract.contractStart?.replace(/-/g, "년 ").replace(/-/, "월 ")}일</div>
            <div style={{ marginTop: 16, display: "flex", justifyContent: "space-around", fontSize: 12 }}>
              <div>(사용자) {contract.ownerName} &nbsp;&nbsp; {contract.sentAt ? `📨 ${new Date(contract.sentAt).toLocaleDateString("ko-KR")} 발송` : ""}</div>
              <div>(근로자) {contract.userName} &nbsp;&nbsp; {contract.status === "signed" && contract.signedAt ? `✅ ${new Date(contract.signedAt).toLocaleDateString("ko-KR")} 전자서명` : ""}</div>
            </div>
          </>)}
          {contract.docType && contract.docType !== "contract" && (<>
            <div style={{ textAlign: "center", fontSize: 20, fontWeight: 900, marginBottom: 8, letterSpacing: 4 }}>
              {DOC_TYPES.find(d => d.key === contract.docType)?.label || "문서"}
            </div>
            {contract.docTitle && <div style={{ textAlign: "center", fontSize: 14, fontWeight: 700, marginBottom: 24, color: "#444" }}>{contract.docTitle}</div>}
            <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: 20, fontSize: 12 }}>
              <tbody>
                <tr><td style={{ padding: "4px 8px", fontWeight: 700, width: 80 }}>사업체명</td><td style={{ padding: "4px 8px" }}>하나기업</td></tr>
                <tr><td style={{ padding: "4px 8px", fontWeight: 700 }}>대표자</td><td style={{ padding: "4px 8px" }}>박용균</td></tr>
                <tr><td style={{ padding: "4px 8px", fontWeight: 700 }}>성명</td><td style={{ padding: "4px 8px" }}>{contract.userName}</td></tr>
              </tbody>
            </table>
            {contract.docContent && (
              <div style={{ fontSize: 12, lineHeight: 1.9, whiteSpace: "pre-wrap", marginBottom: 24, padding: "12px 0", borderTop: "1px solid #e5e7eb", borderBottom: "1px solid #e5e7eb" }}>
                {contract.docContent}
              </div>
            )}
            <div style={{ marginTop: 24, fontSize: 12 }}>{contract.createdAt ? new Date(contract.createdAt).toLocaleDateString("ko-KR", { year: "numeric", month: "long", day: "numeric" }) : ""}</div>
            <div style={{ marginTop: 16, display: "flex", justifyContent: "space-around", fontSize: 12 }}>
              <div>(사업자) 박용균 &nbsp;&nbsp; {contract.sentAt ? `📨 ${new Date(contract.sentAt).toLocaleDateString("ko-KR")} 발송` : ""}</div>
              <div>(서명인) {contract.userName} &nbsp;&nbsp; {contract.status === "signed" && contract.signedAt ? `✅ ${new Date(contract.signedAt).toLocaleDateString("ko-KR")} 전자서명` : ""}</div>
            </div>
          </>)}
        </div>
      )}
      </>)}
    </div>
  );
}
// ── 교육 (독립 섹션) ────────────────────────────────────────────
function EducationSection({ users, reads, onBack }) {
  const members = users.filter(u => u.role === "member" && (!u.status || u.status === "active"));
  const [educations, setEducations] = useState([]);
  const [showForm, setShowForm] = useState(false);
  const [title, setTitle] = useState("");
  const [content, setContent] = useState("");
  const [eduDate, setEduDate] = useState("");
  const [file, setFile] = useState(null);
  const [uploading, setUploading] = useState(false);
  const [expanded, setExpanded] = useState({});

  useEffect(() => {
    const unsub = onSnapshot(
      query(collection(db, "education"), orderBy("createdAt", "desc")),
      snap => setEducations(snap.docs.map(d => ({ id: d.id, ...d.data() }))),
      () => setEducations([])
    );
    return () => unsub();
  }, []);

  const reset = () => { setTitle(""); setContent(""); setEduDate(""); setFile(null); setShowForm(false); };

  const submit = async () => {
    if (!title.trim()) { alert("교육명을 입력하세요."); return; }
    setUploading(true);
    try {
      let fileUrl = null, fileName = null;
      if (file) {
        const sRef = ref(storage, `education/${Date.now()}_${file.name}`);
        await uploadBytes(sRef, file);
        fileUrl = await getDownloadURL(sRef);
        fileName = file.name;
      }
      await addDoc(collection(db, "education"), {
        title: title.trim(), content: content.trim(), eduDate,
        fileUrl, fileName, createdAt: new Date().toISOString(),
      });
      for (const m of members) {
        await addDoc(collection(db, COL_NOTICES), {
          title: `🎓 교육 안내: ${title.trim()}`,
          content: `교육일: ${eduDate || "미정"}\n\n${content.trim()}\n\n교육 탭에서 자료 확인 후 완료 보고해주세요.`,
          recipient: m.id, author: "관리자", auto: true, createdAt: new Date().toISOString(),
        });
      }
      reset();
    } catch(e) { alert("오류: " + e.message); }
    setUploading(false);
  };

  return (
    <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'Noto Sans KR',sans-serif", paddingBottom: 40 }}>
      <div style={{ background: "#7c3aed", paddingTop: "calc(16px + env(safe-area-inset-top))", paddingBottom: 14, paddingLeft: 16, paddingRight: 16 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <button onClick={onBack} style={{ background: "#ffffff18", border: "none", color: "#fff", fontSize: 18, cursor: "pointer", padding: "8px 14px", borderRadius: 12, fontWeight: 700 }}>‹</button>
          <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>🎓 교육 관리</div>
        </div>
      </div>
      <div style={{ padding: 16 }}>
        {!showForm && (
          <button onClick={() => setShowForm(true)}
            style={{ width: "100%", padding: "12px 0", borderRadius: 12, border: "2px dashed #c4b5fd", background: "none", color: "#7c3aed", fontSize: 14, fontWeight: 700, cursor: "pointer", marginBottom: 16 }}>
            🎓 + 교육 개설
          </button>
        )}
        {showForm && (
          <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 16, border: `1px solid ${T.border}` }}>
            <div style={{ fontSize: 14, fontWeight: 800, color: T.text, marginBottom: 12 }}>🎓 교육 개설</div>
            {[["교육명", title, setTitle, "예: 안전보건교육"], ["교육일", eduDate, setEduDate, "예: 2026-07-15"]].map(([label, val, setter, ph]) => (
              <div key={label} style={{ marginBottom: 10 }}>
                <div style={{ fontSize: 12, color: T.muted, fontWeight: 600, marginBottom: 4 }}>{label}</div>
                <input value={val} onChange={e => setter(e.target.value)} placeholder={ph}
                  style={{ width: "100%", padding: "10px 12px", borderRadius: 10, border: `1px solid ${T.border}`, fontSize: 13, boxSizing: "border-box", fontFamily: "inherit" }} />
              </div>
            ))}
            <div style={{ marginBottom: 10 }}>
              <div style={{ fontSize: 12, color: T.muted, fontWeight: 600, marginBottom: 4 }}>교육 내용</div>
              <textarea value={content} onChange={e => setContent(e.target.value)} placeholder="내용 입력" rows={4}
                style={{ width: "100%", padding: "10px 12px", borderRadius: 10, border: `1px solid ${T.border}`, fontSize: 13, boxSizing: "border-box", fontFamily: "inherit", resize: "vertical" }} />
            </div>
            <div style={{ marginBottom: 12 }}>
              <div style={{ fontSize: 12, color: T.muted, fontWeight: 600, marginBottom: 4 }}>자료 첨부 (선택)</div>
              <label style={{ display: "block", padding: "10px 0", borderRadius: 10, border: `1px dashed ${T.border}`, background: T.bg, color: file ? "#16a34a" : T.muted, fontSize: 12, fontWeight: 600, cursor: "pointer", textAlign: "center" }}>
                {file ? `📎 ${file.name}` : "📁 파일 선택"}
                <input type="file" accept=".pdf,.ppt,.pptx,.doc,.docx,image/*" style={{ display: "none" }} onChange={e => setFile(e.target.files?.[0])} />
              </label>
              {file && <button onClick={() => setFile(null)} style={{ marginTop: 4, fontSize: 11, color: "#b91c1c", background: "none", border: "none", cursor: "pointer" }}>✕ 제거</button>}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              <Btn variant="ghost" onClick={reset}>취소</Btn>
              <Btn variant="admin" onClick={submit} disabled={uploading}>{uploading ? "등록 중..." : "📢 개설 + 공지"}</Btn>
            </div>
          </div>
        )}
        {educations.length === 0 && !showForm && (
          <div style={{ textAlign: "center", padding: 40, color: T.muted, fontSize: 14 }}>등록된 교육이 없습니다</div>
        )}
        {educations.map(edu => {
          const isOpen = expanded[edu.id];
          const done = members.filter(m => reads[`${m.id}_edu_${edu.id}`]);
          const notDone = members.filter(m => !reads[`${m.id}_edu_${edu.id}`]);
          return (
            <div key={edu.id} style={{ background: T.card, borderRadius: 16, marginBottom: 12, border: `1px solid ${T.border}`, overflow: "hidden" }}>
              <div style={{ padding: "14px 16px", display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer" }}
                onClick={() => setExpanded(p => ({ ...p, [edu.id]: !p[edu.id] }))}>
                <div>
                  <div style={{ fontSize: 14, fontWeight: 800, color: T.text }}>🎓 {edu.title}</div>
                  <div style={{ fontSize: 11, color: T.muted, marginTop: 2 }}>
                    {edu.eduDate && `${edu.eduDate} · `}
                    <span style={{ color: done.length === members.length && members.length > 0 ? "#16a34a" : "#d97706", fontWeight: 700 }}>
                      완료 {done.length}/{members.length}명
                    </span>
                  </div>
                </div>
                <span style={{ color: T.muted }}>{isOpen ? "▲" : "▼"}</span>
              </div>
              {isOpen && (
                <div style={{ borderTop: `1px solid ${T.border}`, padding: "14px 16px" }}>
                  {edu.content && <div style={{ fontSize: 13, color: T.text, lineHeight: 1.7, whiteSpace: "pre-wrap", marginBottom: 12 }}>{edu.content}</div>}
                  {edu.fileUrl && (
                    <a href={edu.fileUrl} target="_blank" rel="noreferrer"
                      style={{ display: "flex", alignItems: "center", gap: 6, padding: "10px 14px", background: "#eff6ff", borderRadius: 10, color: "#2563eb", fontSize: 13, fontWeight: 700, textDecoration: "none", marginBottom: 12 }}>
                      📎 {edu.fileName || "자료 다운로드"}
                    </a>
                  )}
                  <div style={{ background: T.bg, borderRadius: 10, padding: "10px 12px", marginBottom: 10 }}>
                    <div style={{ fontSize: 12, fontWeight: 700, color: T.text, marginBottom: 8 }}>완료 현황</div>
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                      {members.map(m => (
                        <span key={m.id} style={{ fontSize: 12, padding: "4px 10px", borderRadius: 20, fontWeight: 600,
                          background: reads[`${m.id}_edu_${edu.id}`] ? "#dcfce7" : "#fee2e2",
                          color: reads[`${m.id}_edu_${edu.id}`] ? "#16a34a" : "#b91c1c" }}>
                          {reads[`${m.id}_edu_${edu.id}`] ? "✅" : "⏳"} {m.name}
                        </span>
                      ))}
                    </div>
                  </div>
                  {notDone.length > 0 && (
                    <button onClick={async () => {
                      for (const m of notDone) {
                        await sendPush({ title: "📣 교육 완료 독촉", message: `"${edu.title}" 교육 완료 보고를 해주세요.`, targetUserId: m.id });
                      }
                      alert(`${notDone.map(m => m.name).join(", ")}님께 독촉했습니다.`);
                    }}
                      style={{ width: "100%", padding: "9px 0", borderRadius: 10, border: "none", background: "#fff7ed", color: "#d97706", fontSize: 13, fontWeight: 700, cursor: "pointer", marginBottom: 8 }}>
                      📣 미완료 독촉 ({notDone.length}명)
                    </button>
                  )}
                  <button onClick={async () => {
                    if (!window.confirm(`"${edu.title}" 교육을 삭제할까요?`)) return;
                    await deleteDoc(doc(db, "education", edu.id));
                  }}
                    style={{ width: "100%", padding: "9px 0", borderRadius: 10, border: "none", background: "#fee2e2", color: "#b91c1c", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                    🗑 삭제
                  </button>
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ── 교육 (팀원) ─────────────────────────────────────────────────
function MemberEducationTab({ user, reads }) {
  const [educations, setEducations] = useState([]);

  useEffect(() => {
    const unsub = onSnapshot(
      query(collection(db, "education"), orderBy("createdAt", "desc")),
      snap => setEducations(snap.docs.map(d => ({ id: d.id, ...d.data() }))),
      () => setEducations([])
    );
    return () => unsub();
  }, []);

  return (
    <div style={{ paddingBottom: 80 }}>
      <div style={{ background: "#7c3aed", paddingTop: "calc(18px + env(safe-area-inset-top))", paddingBottom: 14, paddingLeft: 16, paddingRight: 16 }}>
        <div style={{ fontSize: 11, color: "#ffffff50", letterSpacing: 3 }}>ATTENDANCE</div>
        <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>🎓 교육</div>
      </div>
      <div style={{ padding: 16 }}>
        {educations.length === 0 && (
          <div style={{ textAlign: "center", padding: 40, color: T.muted, fontSize: 14 }}>등록된 교육이 없습니다</div>
        )}
        {educations.map(edu => {
          const myDone = reads[`${user.id}_edu_${edu.id}`];
          return (
            <div key={edu.id} style={{ background: T.card, borderRadius: 16, marginBottom: 12, border: `1px solid ${T.border}`, overflow: "hidden" }}>
              <div style={{ padding: "14px 16px" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
                  <div style={{ fontSize: 14, fontWeight: 800, color: T.text }}>🎓 {edu.title}</div>
                  {myDone
                    ? <span style={{ fontSize: 11, padding: "3px 8px", borderRadius: 20, background: "#dcfce7", color: "#16a34a", fontWeight: 700 }}>✅ 완료</span>
                    : <span style={{ fontSize: 11, padding: "3px 8px", borderRadius: 20, background: "#fef3c7", color: "#d97706", fontWeight: 700 }}>⏳ 미완료</span>
                  }
                </div>
                {edu.eduDate && <div style={{ fontSize: 11, color: T.muted, marginBottom: 8 }}>📅 교육일: {edu.eduDate}</div>}
                {edu.content && <div style={{ fontSize: 13, color: T.text, lineHeight: 1.7, whiteSpace: "pre-wrap", marginBottom: 12 }}>{edu.content}</div>}
                {edu.fileUrl && (
                  <a href={edu.fileUrl} target="_blank" rel="noreferrer"
                    style={{ display: "flex", alignItems: "center", gap: 6, padding: "10px 14px", background: "#eff6ff", borderRadius: 10, color: "#2563eb", fontSize: 13, fontWeight: 700, textDecoration: "none", marginBottom: 12 }}>
                    📎 {edu.fileName || "자료 열람 / 다운로드"}
                  </a>
                )}
                {myDone ? (
                  <div style={{ padding: "12px", background: "#f0fdf4", borderRadius: 10, textAlign: "center" }}>
                    <div style={{ fontSize: 13, fontWeight: 700, color: "#16a34a" }}>
                      ✅ {new Date(myDone.readAt).toLocaleString("ko-KR")} 완료 보고
                    </div>
                  </div>
                ) : (
                  <button onClick={async () => {
                    if (!window.confirm(`"${edu.title}" 교육 완료를 보고할까요?`)) return;
                    await setDoc(doc(db, COL_READS, `${user.id}_edu_${edu.id}`), {
                      userId: user.id, type: "edu", docId: edu.id,
                      readAt: new Date().toISOString(), userName: user.name,
                    });
                    await sendPush({ title: "✅ 교육 완료 보고", message: `${user.name}님이 "${edu.title}" 교육을 완료하였습니다.`, targetUserId: "admin" });
                    alert("교육 완료 보고 완료!");
                  }}
                    style={{ width: "100%", padding: "13px 0", borderRadius: 12, border: "none", background: "#7c3aed", color: "#fff", fontSize: 14, fontWeight: 800, cursor: "pointer" }}>
                    ✅ 교육 완료 보고
                  </button>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ── 4대보험료 계산기 ────────────────────────────────────────────
function InsuranceSection({ users, memberInfo, onBack }) {
  const 장기요양요율 = 0.1314;
  const activeMembers = users.filter(u => u.role === "member" && (!u.status || u.status === "active"));

  function calcOne(pension, health, isOwner = false) {
    const 국민연금_근로자 = Math.floor(pension * 0.0475 / 10) * 10;
    const 국민연금_사업주 = Math.floor(pension * 0.0475 / 10) * 10;
    const 건강보험_근로자 = Math.floor(health * 0.03595 / 10) * 10;
    const 건강보험_사업주 = Math.floor(health * 0.03595 / 10) * 10;
    const 장기요양_근로자 = Math.floor(건강보험_근로자 * 장기요양요율 / 10) * 10;
    const 장기요양_사업주 = Math.floor(건강보험_사업주 * 장기요양요율 / 10) * 10;
    const 실업급여_근로자 = isOwner ? 0 : Math.floor(health * 0.009 / 10) * 10;
    const 실업급여_사업주 = isOwner ? 0 : Math.floor(health * 0.009 / 10) * 10;
    const 고안_사업주     = isOwner ? 0 : Math.floor(health * 0.0025 / 10) * 10;
    const 합계_근로자 = 국민연금_근로자 + 건강보험_근로자 + 장기요양_근로자 + 실업급여_근로자;
    const 합계_사업주 = 국민연금_사업주 + 건강보험_사업주 + 장기요양_사업주 + 실업급여_사업주 + 고안_사업주;
    return {
      rows: [
        { 항목: "국민연금",            요율: "각 4.75%",                             근로자: 국민연금_근로자, 사업주: 국민연금_사업주 },
        { 항목: "건강보험",            요율: "각 3.595%",                            근로자: 건강보험_근로자, 사업주: 건강보험_사업주 },
        { 항목: "장기요양",            요율: "건강료×13.14%",                        근로자: 장기요양_근로자, 사업주: 장기요양_사업주 },
        { 항목: "고용보험(실업급여)",  요율: isOwner ? "적용제외" : "각 0.9%",      근로자: 실업급여_근로자, 사업주: 실업급여_사업주 },
        { 항목: "고용보험(고안·직능)", 요율: isOwner ? "적용제외" : "사업주 0.25%", 근로자: 0,               사업주: 고안_사업주 },
      ],
      합계_근로자, 합계_사업주,
    };
  }

  const admin = users.find(u => u.role === "admin");
  const [전자통보, set전자통보] = useState(true);
  const [산재율, set산재율] = useState("8.6");
  const [임채율, set임채율] = useState("0.9");
  const [ownerPension, setOwnerPension] = useState("");
  const [ownerHealth, setOwnerHealth] = useState("");
  const [memberInputs, setMemberInputs] = useState(() =>
    activeMembers.map(m => {
      const info = memberInfo[m.id] || {};
      const pension = Number(info.pensionBase) || 0;
      const health = Number(info.insuranceBase) || 0;
      return { id: m.id, name: m.name, pension: pension ? String(pension) : "", health: health ? String(health) : "" };
    })
  );
  const [results, setResults] = useState(null);

  // 관리자 기초데이터 불러오기
  useEffect(() => {
    if (!admin?.id) return;
    getDoc(doc(db, COL_MEMBER_INFO, admin.id)).then(snap => {
      if (snap.exists()) {
        const d = snap.data();
        setOwnerPension(String(Number(d.pensionBase) || 0));
        setOwnerHealth(String(Number(d.insuranceBase) || 0));
      }
    });
  }, [admin?.id]);

  const fmt = (n) => n === 0 ? "-" : n.toLocaleString() + "원";
  const num = (v) => Number(v) || 0;
  const iStyle = { width: "100%", padding: "9px 12px", borderRadius: 10, border: `1px solid ${T.border}`, fontSize: 13, color: T.text, textAlign: "right", boxSizing: "border-box", fontFamily: "inherit", background: "#fff" };

  const calculate = () => {
    try {
      const ownerResult = { name: "관리자(사업주)", isOwner: true, pension: num(ownerPension), health: num(ownerHealth), ...calcOne(num(ownerPension), num(ownerHealth), true) };
      const memberResults = memberInputs.filter(m => m.pension || m.health).map(m => ({
        name: m.name, isOwner: false, pension: num(m.pension), health: num(m.health),
        ...calcOne(num(m.pension), num(m.health), false),
      }));
      const all = [ownerResult, ...memberResults];
      const 팀원합산보수 = memberResults.reduce((s, r) => s + r.health, 0);
      const 산재보험료 = Math.floor(팀원합산보수 * (parseFloat(산재율) || 0) / 1000 / 10) * 10;
      const 임금채권료 = Math.floor(팀원합산보수 * (parseFloat(임채율) || 0) / 1000 / 10) * 10;
      const 산재임채합계 = 산재보험료 + 임금채권료;
      const 전자통보감액 = 전자통보 ? 200 : 0;
      const 보험항목 = ["국민연금", "건강보험", "장기요양", "고용보험(실업급여)", "고용보험(고안·직능)"];
      const 보험별합계 = 보험항목.map(항목 => {
        const 근로자합계 = all.reduce((s, r) => s + (r.rows.find(row => row.항목 === 항목)?.근로자 || 0), 0);
        const 사업주합계 = all.reduce((s, r) => s + (r.rows.find(row => row.항목 === 항목)?.사업주 || 0), 0);
        return { 항목, 근로자합계, 사업주합계, 합계: 근로자합계 + 사업주합계 };
      });
      const total = {
        직원합계_근로자: memberResults.reduce((s, r) => s + r.합계_근로자, 0),
        직원합계_사업주: memberResults.reduce((s, r) => s + r.합계_사업주, 0),
        관리자_근로자: ownerResult.합계_근로자,
        관리자_사업주: ownerResult.합계_사업주,
        산재보험료, 임금채권료, 산재임채합계, 팀원합산보수, 전자통보감액, 보험별합계,
      };
      total.회사총부담 = total.직원합계_사업주 + total.관리자_사업주 + total.관리자_근로자 - 전자통보감액 + 산재임채합계;
      total.전체보험료 = total.직원합계_근로자 + total.직원합계_사업주 + total.관리자_근로자 + total.관리자_사업주 - 전자통보감액 + 산재임채합계;
      setResults({ all, total });
    } catch(e) {
      alert("계산 오류: " + e.message);
    }
  };

  const downloadExcel = () => {
    if (!results) return;
    const { all, total } = results;
    const wb = XLSX.utils.book_new();
    const now = new Date();
    const rows = [];

    // 제목
    rows.push([`${now.getFullYear()}년 ${now.getMonth()+1}월 4대보험료 계산 내역`]);
    rows.push([`산재 ${산재율}‰  /  임금채권 ${임채율}‰  /  전자통보 감액 ${total.전자통보감액 ? "적용(-200원)" : "미적용"}  /  2026년 기준`]);
    rows.push([]);

    // 헤더행: 항목 | 구분 | 관리자 | 이현주 | 김재우 | ... | 합계
    const names = all.map(r => r.name);
    rows.push(["항목", "구분", ...names, "합계"]);

    // 보험 항목별 근로자/사업주 행
    const 항목목록 = ["국민연금", "건강보험", "장기요양", "고용보험(실업급여)", "고용보험(고안·직능)"];
    항목목록.forEach(항목 => {
      const 근로자행 = [항목, "근로자"];
      const 사업주행 = ["", "사업주"];
      let 근로자합계 = 0, 사업주합계 = 0;
      all.forEach(r => {
        const row = r.rows.find(row => row.항목 === 항목);
        const 근 = row?.근로자 || 0;
        const 사 = row?.사업주 || 0;
        근로자행.push(근 === 0 ? "-" : 근);
        사업주행.push(사 === 0 ? "-" : 사);
        근로자합계 += 근;
        사업주합계 += 사;
      });
      근로자행.push(근로자합계 || "-");
      사업주행.push(사업주합계 || "-");
      rows.push(근로자행);
      rows.push(사업주행);
    });

    // 산재/임채 (사업주만, 팀원 합산 기준)
    rows.push([`산재보험 (${산재율}‰)`, "사업주", ...all.map((r, i) => i === 0 ? "-" : ""), `(팀원합산 ${total.팀원합산보수.toLocaleString()}원→${total.산재보험료.toLocaleString()})`]);
    rows.push([`임금채권 (${임채율}‰)`, "사업주", ...all.map((r, i) => i === 0 ? "-" : ""), `(팀원합산→${total.임금채권료.toLocaleString()})`]);

    rows.push([]);

    // 소계행
    const 근로자소계행 = ["소  계", "근로자"];
    const 사업주소계행 = ["", "사업주"];
    all.forEach(r => {
      근로자소계행.push(r.합계_근로자);
      사업주소계행.push(r.합계_사업주);
    });
    근로자소계행.push(total.보험별합계.reduce((s, r) => s + r.근로자합계, 0));
    사업주소계행.push(total.보험별합계.reduce((s, r) => s + r.사업주합계, 0));
    rows.push(근로자소계행);
    rows.push(사업주소계행);
    rows.push([]);

    // 납부 요약
    rows.push(["납부 요약", "", ...all.map(() => ""), ""]);
    rows.push(["급여 공제 합계", "(직원 근로자 부담)", ...all.map(r => r.isOwner ? "" : r.합계_근로자), total.직원합계_근로자]);
    rows.push(["회사 납부 보험료", "(직원 사업주분+관리자)", ...all.map(() => ""), total.회사총부담]);
    rows.push([`산재+임채 (별도)`, `팀원합산 ${total.팀원합산보수.toLocaleString()}원`, ...all.map(() => ""), total.산재임채합계]);
    rows.push(["★ 전체 보험료 합계", "(산재·임채 포함)", ...all.map(() => ""), total.전체보험료]);

    const ws = XLSX.utils.aoa_to_sheet(rows);
    const colW = [{ wch: 22 }, { wch: 12 }, ...all.map(() => ({ wch: 13 })), { wch: 20 }];
    ws["!cols"] = colW;

    XLSX.utils.book_append_sheet(wb, ws, `${now.getFullYear()}.${String(now.getMonth()+1).padStart(2,"0")} 보험료`);
    XLSX.writeFile(wb, `4대보험료_${now.getFullYear()}${String(now.getMonth()+1).padStart(2,"0")}.xlsx`);
  };

  return (
    <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'Noto Sans KR',sans-serif", paddingBottom: 40 }}>
      <div style={{ background: "#16a34a", paddingTop: "calc(16px + env(safe-area-inset-top))", paddingBottom: "14px", paddingLeft: "16px", paddingRight: "16px" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
          <button onClick={onBack} style={{ background: "#ffffff18", border: "none", color: "#fff", fontSize: 18, cursor: "pointer", padding: "8px 14px", borderRadius: 12, fontWeight: 700 }}>‹</button>
          <div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>💰 4대보험료 계산기</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 12px", background: "#ffffff18", borderRadius: 10, cursor: "pointer", marginBottom: 8 }}
          onClick={() => set전자통보(p => !p)}>
          <input type="checkbox" checked={전자통보} onChange={() => {}} style={{ width: 16, height: 16, cursor: "pointer" }} />
          <span style={{ fontSize: 12, fontWeight: 700, color: "#fff" }}>국민연금 전자통보 감액 (-200원)</span>
        </div>
        <div style={{ background: "#ffffff18", borderRadius: 10, padding: "10px 12px" }}>
          <div style={{ fontSize: 11, fontWeight: 700, color: "#ffffff90", marginBottom: 8 }}>🏭 산재·임채 (‰, 팀원 합산보수 기준)</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {[["산재보험", 산재율, set산재율], ["임금채권부담금", 임채율, set임채율]].map(([label, val, setter]) => (
              <div key={label} style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                <span style={{ fontSize: 12, color: "#ffffff90", fontWeight: 600 }}>{label}</span>
                <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                  <input value={val} onChange={e => setter(e.target.value.replace(/[^0-9.]/g, ""))}
                    style={{ width: 80, padding: "7px 10px", borderRadius: 8, border: "none", background: "#ffffff25", color: "#fff", fontSize: 14, fontWeight: 800, textAlign: "right", fontFamily: "inherit" }} />
                  <span style={{ fontSize: 11, color: "#ffffff70" }}>‰</span>
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>

      <div style={{ padding: 16 }}>
        {/* 관리자 */}
        <div style={{ background: "#faf5ff", borderRadius: 14, padding: 14, marginBottom: 12, border: "2px solid #c4b5fd" }}>
          <div style={{ fontSize: 13, fontWeight: 800, color: "#7c3aed", marginBottom: 4 }}>👑 관리자 (사업주)</div>
          <div style={{ fontSize: 11, color: "#94a3b8", marginBottom: 10 }}>※ 산재·임채·고용보험 적용 제외</div>
          {[["국민연금 기준소득월액", ownerPension, setOwnerPension], ["건강보험·고용보험 보수월액", ownerHealth, setOwnerHealth]].map(([label, val, setter]) => (
            <div key={label} style={{ marginBottom: 8 }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: "#7c3aed", marginBottom: 4 }}>{label}</div>
              <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                <input value={val} onChange={e => setter(e.target.value.replace(/[^0-9]/g, ""))} placeholder="0"
                  style={{ ...iStyle, border: "1px solid #c4b5fd", background: "#faf5ff" }} />
                <span style={{ fontSize: 12, color: T.muted }}>원</span>
              </div>
            </div>
          ))}
        </div>

        {/* 팀원 */}
        <div style={{ background: T.card, borderRadius: 14, padding: 14, marginBottom: 16, border: `1px solid ${T.border}` }}>
          <div style={{ fontSize: 13, fontWeight: 800, color: T.text, marginBottom: 4 }}>👤 팀원</div>
          <div style={{ fontSize: 11, color: "#94a3b8", marginBottom: 10 }}>※ 기초데이터 보험 과세표준 자동 입력 · 수정 가능</div>
          {memberInputs.map((m, i) => (
            <div key={m.id} style={{ background: T.bg, borderRadius: 12, padding: 12, marginBottom: 10, border: `1px solid ${T.border}` }}>
              <div style={{ fontSize: 13, fontWeight: 800, color: T.text, marginBottom: 10 }}>{m.name}</div>
              {[["국민연금 기준소득월액", "pension"], ["건강보험·고용보험 보수월액", "health"]].map(([label, field]) => (
                <div key={field} style={{ marginBottom: 8 }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: "#475569", marginBottom: 4 }}>{label}</div>
                  <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                    <input value={m[field]} onChange={e => setMemberInputs(p => p.map((x, j) => j === i ? { ...x, [field]: e.target.value.replace(/[^0-9]/g, "") } : x))}
                      placeholder="0" style={iStyle} />
                    <span style={{ fontSize: 12, color: T.muted }}>원</span>
                  </div>
                </div>
              ))}
            </div>
          ))}
        </div>

        <button onClick={downloadExcel} disabled={!results}
          style={{ display: "block", width: "100%", padding: "14px 0", borderRadius: 14, border: "none", background: results ? "#0369a1" : "#e5e7eb", color: results ? "#fff" : "#9ca3af", fontSize: 14, fontWeight: 800, cursor: results ? "pointer" : "default", marginBottom: 10 }}>
          📥 엑셀 다운로드
        </button>
        <button onClick={calculate}
          style={{ display: "block", width: "100%", padding: "14px 0", borderRadius: 14, border: "none", background: "#16a34a", color: "#fff", fontSize: 15, fontWeight: 800, cursor: "pointer", marginBottom: 20 }}>
          💰 계산하기
        </button>

        {results && (<>
          <div style={{ fontSize: 14, fontWeight: 800, color: T.text, marginBottom: 10 }}>📊 개인별 내역</div>
          {results.all.map((r, ri) => (
            <div key={ri} style={{ background: T.card, borderRadius: 14, padding: 14, marginBottom: 12, border: r.isOwner ? "2px solid #c4b5fd" : `1px solid ${T.border}` }}>
              <div style={{ marginBottom: 10 }}>
                <div style={{ fontSize: 14, fontWeight: 800, color: T.text }}>
                  {r.isOwner && <span style={{ fontSize: 10, background: "#7c3aed", color: "#fff", borderRadius: 4, padding: "1px 6px", marginRight: 6 }}>사업주</span>}
                  {r.name}
                </div>
                <div style={{ fontSize: 10, color: T.muted, marginTop: 2 }}>
                  국민연금: {r.pension.toLocaleString()}원 · 건강/고용: {r.health.toLocaleString()}원
                </div>
              </div>
              <div style={{ borderRadius: 10, overflow: "hidden", border: `1px solid ${T.border}` }}>
                <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr 1fr", background: "#f1f5f9", padding: "7px 10px", fontSize: 10, fontWeight: 700, color: "#475569" }}>
                  <span>항목</span><span style={{ textAlign: "right" }}>근로자</span><span style={{ textAlign: "right" }}>사업주</span>
                </div>
                {r.rows.map((row, i) => (
                  <div key={i} style={{ display: "grid", gridTemplateColumns: "2fr 1fr 1fr", padding: "8px 10px", borderTop: `1px solid ${T.border}`, background: i % 2 === 0 ? "#fff" : "#fafafa" }}>
                    <div>
                      <div style={{ fontSize: 11, fontWeight: 700, color: T.text }}>{row.항목}</div>
                      <div style={{ fontSize: 10, color: T.muted }}>{row.요율}</div>
                    </div>
                    <div style={{ textAlign: "right", fontSize: 11, fontWeight: 700, color: row.근로자 === 0 ? T.muted : "#dc2626" }}>{fmt(row.근로자)}</div>
                    <div style={{ textAlign: "right", fontSize: 11, fontWeight: 700, color: row.사업주 === 0 ? T.muted : "#d97706" }}>{fmt(row.사업주)}</div>
                  </div>
                ))}
                <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr 1fr", padding: "10px", borderTop: `2px solid ${T.border}`, background: "#f8fafc" }}>
                  <div style={{ fontSize: 12, fontWeight: 800, color: T.text }}>합계</div>
                  <div style={{ textAlign: "right", fontSize: 12, fontWeight: 800, color: "#dc2626" }}>{r.합계_근로자.toLocaleString()}원</div>
                  <div style={{ textAlign: "right", fontSize: 12, fontWeight: 800, color: "#d97706" }}>{r.합계_사업주.toLocaleString()}원</div>
                </div>
              </div>
            </div>
          ))}

          {/* 보험별 합계 */}
          <div style={{ background: T.card, borderRadius: 14, padding: 14, marginBottom: 12, border: `1px solid ${T.border}` }}>
            <div style={{ fontSize: 14, fontWeight: 800, color: T.text, marginBottom: 12 }}>📋 보험별 합계 (전원)</div>
            <div style={{ borderRadius: 10, overflow: "hidden", border: `1px solid ${T.border}` }}>
              <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr 1fr 1fr", background: "#f1f5f9", padding: "7px 10px", fontSize: 10, fontWeight: 700, color: "#475569" }}>
                <span>보험</span><span style={{ textAlign: "right" }}>근로자</span><span style={{ textAlign: "right" }}>사업주</span><span style={{ textAlign: "right" }}>합계</span>
              </div>
              {results.total.보험별합계.map((row, i) => (
                <div key={i} style={{ display: "grid", gridTemplateColumns: "2fr 1fr 1fr 1fr", padding: "8px 10px", borderTop: `1px solid ${T.border}`, background: i % 2 === 0 ? "#fff" : "#fafafa" }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: T.text }}>{row.항목}</div>
                  <div style={{ textAlign: "right", fontSize: 11, color: row.근로자합계 === 0 ? T.muted : "#dc2626", fontWeight: 600 }}>{row.근로자합계 === 0 ? "-" : row.근로자합계.toLocaleString()}</div>
                  <div style={{ textAlign: "right", fontSize: 11, color: "#d97706", fontWeight: 600 }}>{row.사업주합계.toLocaleString()}</div>
                  <div style={{ textAlign: "right", fontSize: 11, fontWeight: 800, color: T.text }}>{row.합계.toLocaleString()}</div>
                </div>
              ))}
              <div style={{ padding: "8px 10px", borderTop: `1px solid ${T.border}`, background: "#fafafa" }}>
                <div style={{ fontSize: 11, fontWeight: 700, color: "#475569", marginBottom: 6 }}>
                  🏭 산재·임채 (팀원합산 {results.total.팀원합산보수.toLocaleString()}원)
                </div>
                {[["산재보험", 산재율, results.total.산재보험료], ["임금채권", 임채율, results.total.임금채권료]].map(([label, rate, amt]) => (
                  <div key={label} style={{ display: "grid", gridTemplateColumns: "2fr 1fr 1fr 1fr", fontSize: 11, marginBottom: 4 }}>
                    <span style={{ color: T.text, fontWeight: 600 }}>{label} ({rate}‰)</span>
                    <span style={{ textAlign: "right", color: T.muted }}>-</span>
                    <span style={{ textAlign: "right", color: "#d97706", fontWeight: 700 }}>{amt.toLocaleString()}</span>
                    <span style={{ textAlign: "right", fontWeight: 800 }}>{amt.toLocaleString()}</span>
                  </div>
                ))}
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr 1fr 1fr", padding: "10px", borderTop: `2px solid ${T.border}`, background: "#f8fafc" }}>
                <div style={{ fontSize: 12, fontWeight: 800, color: T.text }}>합계</div>
                <div style={{ textAlign: "right", fontSize: 12, fontWeight: 800, color: "#dc2626" }}>{results.total.보험별합계.reduce((s, r) => s + r.근로자합계, 0).toLocaleString()}</div>
                <div style={{ textAlign: "right", fontSize: 12, fontWeight: 800, color: "#d97706" }}>{(results.total.보험별합계.reduce((s, r) => s + r.사업주합계, 0) + results.total.산재임채합계).toLocaleString()}</div>
                <div style={{ textAlign: "right", fontSize: 13, fontWeight: 800, color: "#7c3aed" }}>{results.total.전체보험료.toLocaleString()}</div>
              </div>
            </div>
          </div>

          {/* 납부 요약 */}
          <div style={{ background: T.adminHeader, borderRadius: 16, padding: 20 }}>
            <div style={{ fontSize: 14, fontWeight: 800, color: "#fff", marginBottom: 14 }}>💡 납부 요약</div>
            {[
              { label: "급여 공제 합계", sub: "직원 근로자 부담분", value: results.total.직원합계_근로자, color: "#fca5a5" },
              { label: "회사 납부 보험료", sub: `직원 사업주분 + 관리자${results.total.전자통보감액 ? " (전자통보 -200원)" : ""}`, value: results.total.회사총부담, color: "#fcd34d" },
              { label: "산재 + 임채", sub: `팀원 합산 ${results.total.팀원합산보수.toLocaleString()}원 기준`, value: results.total.산재임채합계, color: "#f9a8d4" },
              { label: "전체 보험료 합계", sub: "산재·임채 포함", value: results.total.전체보험료, color: "#a5f3fc", big: true },
            ].map((s, i) => (
              <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: s.big ? "12px 14px" : "10px 0", background: s.big ? "#ffffff20" : "none", borderRadius: s.big ? 10 : 0, borderTop: i > 0 && !s.big ? "1px solid #ffffff20" : "none" }}>
                <div>
                  <div style={{ fontSize: s.big ? 13 : 12, fontWeight: 700, color: "#e2e8f0" }}>{s.label}</div>
                  <div style={{ fontSize: 10, color: "#94a3b8" }}>{s.sub}</div>
                </div>
                <span style={{ fontSize: s.big ? 16 : 14, fontWeight: 800, color: s.color }}>{s.value.toLocaleString()}원</span>
              </div>
            ))}
          </div>
        </>)}
      </div>
    </div>
  );
}
function VaultSection({ onBack }) {
  const [vault, setVaultLocal] = useState([]);
  const [loading, setLoading] = useState(true);
  const [adding, setAdding] = useState(false); // 새 항목 입력창 토글
  const [newTitle, setNewTitle] = useState("");
  const [newText, setNewText] = useState("");
  const [newFiles, setNewFiles] = useState([]);
  const [saving, setSaving] = useState(false);
  const [deleting, setDeleting] = useState(null);
  const [preview, setPreview] = useState(null);
  const [expanded, setExpanded] = useState({});
  const [editId, setEditId] = useState(null); // 인라인 수정 중인 항목 id
  const [editTitle, setEditTitle] = useState("");
  const [editText, setEditText] = useState("");
  const [editFiles, setEditFiles] = useState([]); // 새로 추가할 파일
  const [editSaving, setEditSaving] = useState(false);
  const newFileRef = useRef(null);
  const editFileRef = useRef(null);

  useEffect(() => {
    const unsub = onSnapshot(query(collection(db, COL_VAULT), orderBy("createdAt", "desc")), snap => {
      setVaultLocal(snap.docs.map(d => ({ id: d.id, ...d.data() })));
      setLoading(false);
    });
    return () => unsub();
  }, []);

  // 새 항목 저장
  const handleAdd = async () => {
    if (!newTitle.trim() && !newText.trim() && newFiles.length === 0) return;
    setSaving(true);
    try {
      const fileDataList = [];
      for (const f of newFiles) {
        const ext = f.name.split(".").pop().toLowerCase();
        const sRef = ref(storage, `vault/${Date.now()}_${f.name}`);
        await uploadBytes(sRef, f);
        const url = await getDownloadURL(sRef);
        fileDataList.push({ url, name: f.name, size: f.size, ext, type: f.type });
      }
      await addDoc(collection(db, COL_VAULT), {
        title: newTitle.trim(), text: newText.trim(), files: fileDataList,
        createdAt: new Date().toISOString(),
      });
      setNewTitle(""); setNewText(""); setNewFiles([]); setAdding(false);
      if (newFileRef.current) newFileRef.current.value = "";
    } catch(e) { alert("저장 실패: " + e.message); }
    setSaving(false);
  };

  // 인라인 수정 시작
  const startEdit = (item) => {
    setEditId(item.id);
    setEditTitle(item.title || "");
    setEditText(item.text || "");
    setEditFiles([]);
  };

  // 인라인 수정 저장
  const handleEditSave = async (item) => {
    setEditSaving(true);
    try {
      const fileDataList = [];
      for (const f of editFiles) {
        const ext = f.name.split(".").pop().toLowerCase();
        const sRef = ref(storage, `vault/${Date.now()}_${f.name}`);
        await uploadBytes(sRef, f);
        const url = await getDownloadURL(sRef);
        fileDataList.push({ url, name: f.name, size: f.size, ext, type: f.type });
      }
      const existingFiles = item.files?.length > 0 ? item.files : (item.file ? [item.file] : []);
      await setDoc(doc(db, COL_VAULT, item.id), {
        ...item, title: editTitle.trim(), text: editText.trim(),
        files: [...existingFiles, ...fileDataList],
        updatedAt: new Date().toISOString(),
      });
      setEditId(null); setEditFiles([]);
      if (editFileRef.current) editFileRef.current.value = "";
    } catch(e) { alert("수정 실패: " + e.message); }
    setEditSaving(false);
  };

  const handleDeleteFile = async (item, fileIdx) => {
    if (!window.confirm("파일을 삭제할까요?")) return;
    try {
      const fileList = item.files?.length > 0 ? item.files : (item.file ? [item.file] : []);
      try { await deleteObject(ref(storage, fileList[fileIdx].url)); } catch {}
      await setDoc(doc(db, COL_VAULT, item.id), { ...item, files: fileList.filter((_, i) => i !== fileIdx), file: null });
    } catch(e) { alert("파일 삭제 실패: " + e.message); }
  };

  const handleDelete = async (item) => {
    if (!window.confirm("삭제할까요?")) return;
    setDeleting(item.id);
    try {
      const fileList = item.files?.length > 0 ? item.files : (item.file ? [item.file] : []);
      for (const f of fileList) { try { await deleteObject(ref(storage, f.url)); } catch {} }
      await deleteDoc(doc(db, COL_VAULT, item.id));
    } catch(e) { alert("삭제 실패: " + e.message); }
    setDeleting(null);
  };

  const isImage = (f) => f?.type?.startsWith("image/");
  const fileIcon = (f) => {
    if (!f) return "📎";
    if (isImage(f)) return "🖼";
    if (f.ext === "pdf") return "📄";
    if (["doc","docx"].includes(f.ext)) return "📝";
    if (["xls","xlsx"].includes(f.ext)) return "📊";
    return "📎";
  };
  const formatSize = (b) => b > 1024*1024 ? `${(b/1024/1024).toFixed(1)}MB` : `${Math.round(b/1024)}KB`;
  const formatDate = (iso) => new Date(iso).toLocaleDateString("ko-KR", { month: "short", day: "numeric", hour: "2-digit", minute: "2-digit" });

  // 공통 파일 목록 UI
  const FileList = ({ files: fList, item }) => (
    <div style={{ display: "flex", flexDirection: "column", gap: 8, marginTop: 10 }}>
      {fList.map((f, i) => (
        <div key={i} style={{ position: "relative" }}>
          {isImage(f) ? (
            <>
              <img src={f.url} alt={f.name} onClick={() => setPreview(f)}
                style={{ width: "100%", maxHeight: 200, objectFit: "cover", borderRadius: 10, cursor: "pointer", display: "block" }} />
              <button onClick={() => handleDeleteFile(item, i)}
                style={{ position: "absolute", top: 6, right: 6, background: "#000000aa", border: "none", color: "#fff", borderRadius: 8, padding: "3px 8px", fontSize: 12, cursor: "pointer" }}>✕</button>
            </>
          ) : (
            <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", background: T.bg, borderRadius: 10, border: `1px solid ${T.border}` }}>
              <span style={{ fontSize: 22 }}>{fileIcon(f)}</span>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontSize: 13, fontWeight: 700, color: T.text, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{f.name}</div>
                <div style={{ fontSize: 11, color: T.muted }}>{formatSize(f.size)}</div>
              </div>
              <a href={f.url} target="_blank" rel="noreferrer" style={{ fontSize: 12, color: "#0891b2", fontWeight: 700, textDecoration: "none" }}>열기</a>
              <button onClick={() => handleDeleteFile(item, i)} style={{ background: "none", border: "none", color: "#b91c1c", cursor: "pointer", fontSize: 15 }}>✕</button>
            </div>
          )}
        </div>
      ))}
    </div>
  );

  // 새 파일 선택 목록 UI
  const NewFileList = ({ files: fList, onRemove }) => fList.length === 0 ? null : (
    <div style={{ display: "flex", flexDirection: "column", gap: 6, marginTop: 8 }}>
      {fList.map((f, i) => (
        <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 10px", background: T.bg, borderRadius: 8, border: `1px solid ${T.border}` }}>
          <span>{fileIcon({ type: f.type, ext: f.name.split(".").pop().toLowerCase() })}</span>
          {isImage({ type: f.type }) && <img src={URL.createObjectURL(f)} alt="" style={{ width: 32, height: 32, objectFit: "cover", borderRadius: 6 }} />}
          <span style={{ fontSize: 12, fontWeight: 600, color: T.text, flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{f.name}</span>
          <span style={{ fontSize: 11, color: T.muted }}>{formatSize(f.size)}</span>
          <button onClick={() => onRemove(i)} style={{ background: "none", border: "none", color: "#b91c1c", cursor: "pointer", fontSize: 15 }}>✕</button>
        </div>
      ))}
    </div>
  );

  return (
    <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'Noto Sans KR',sans-serif", paddingBottom: 30 }}>
      <div style={{ background: T.adminHeader, paddingTop: "calc(16px + env(safe-area-inset-top))", paddingBottom: "14px", paddingLeft: "16px", paddingRight: "16px" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <button onClick={onBack} style={{ background: "#ffffff18", border: "none", color: "#fff", fontSize: 18, cursor: "pointer", padding: "8px 14px", borderRadius: 12, fontWeight: 700 }}>‹</button>
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 11, color: "#ffffff40", letterSpacing: 3 }}>ADMIN</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>📁 보관함</div>
          </div>
          <button onClick={() => { setAdding(p => !p); setNewTitle(""); setNewText(""); setNewFiles([]); }}
            style={{ background: adding ? "#ffffff30" : "#ffffff18", border: "none", color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer", padding: "8px 16px", borderRadius: 12 }}>
            {adding ? "✕ 취소" : "+ 새 항목"}
          </button>
        </div>
      </div>

      <div style={{ padding: "16px 16px 0" }}>
        {/* 새 항목 입력 (토글) */}
        {adding && (
          <div style={{ background: T.card, borderRadius: 16, padding: 16, marginBottom: 16, border: `2px solid ${T.adminHeader}` }}>
            <input type="text" value={newTitle} onChange={e => setNewTitle(e.target.value)}
              placeholder="제목 (선택)"
              style={{ width: "100%", padding: "10px 12px", borderRadius: 10, border: `1px solid ${T.border}`, fontSize: 15, fontWeight: 700, color: T.text, background: T.bg, boxSizing: "border-box", fontFamily: "inherit", marginBottom: 8 }} />
            <textarea value={newText} onChange={e => setNewText(e.target.value)}
              placeholder="메모를 입력하세요..."
              style={{ width: "100%", minHeight: 160, padding: "10px 12px", borderRadius: 10, border: `1px solid ${T.border}`, fontSize: 14, color: T.text, background: T.bg, resize: "vertical", boxSizing: "border-box", fontFamily: "inherit", lineHeight: 1.7 }} />
            <div style={{ marginTop: 10 }}>
              <button onClick={() => newFileRef.current?.click()}
                style={{ padding: "8px 16px", borderRadius: 10, border: `1px dashed ${T.border}`, background: T.bg, color: T.sub, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
                📎 파일 첨부
              </button>
              <input ref={newFileRef} type="file" accept="*/*" multiple style={{ display: "none" }}
                onChange={e => setNewFiles(prev => [...prev, ...Array.from(e.target.files || [])])} />
            </div>
            <NewFileList files={newFiles} onRemove={i => setNewFiles(prev => prev.filter((_, j) => j !== i))} />
            <button onClick={handleAdd} disabled={saving || (!newTitle.trim() && !newText.trim() && newFiles.length === 0)}
              style={{ width: "100%", marginTop: 12, padding: "13px 0", borderRadius: 12, border: "none", background: T.adminHeader, color: "#fff", fontSize: 14, fontWeight: 800, cursor: "pointer" }}>
              {saving ? "저장 중..." : "💾 저장"}
            </button>
          </div>
        )}

        {/* 피드 */}
        {loading ? (
          <div style={{ textAlign: "center", padding: 40, color: T.muted }}>불러오는 중...</div>
        ) : vault.length === 0 ? (
          <div style={{ textAlign: "center", padding: 40 }}>
            <div style={{ fontSize: 40, marginBottom: 10 }}>📁</div>
            <div style={{ fontSize: 14, color: T.muted, fontWeight: 600 }}>저장된 항목이 없습니다<br /><span style={{ fontSize: 12 }}>우측 상단 "+ 새 항목"으로 추가하세요</span></div>
          </div>
        ) : (
          vault.map(item => {
            const isCert = item.title?.includes("수료증") || item.text?.includes("수료증");
            const isExp = expanded[item.id];
            const isEditing = editId === item.id;
            const fileList = item.files?.length > 0 ? item.files : (item.file ? [item.file] : []);
            const previewText = item.text?.slice(0, 80) + (item.text?.length > 80 ? "..." : "");
            return (
              <div key={item.id} style={{ background: isCert ? "#fefce8" : T.card, borderRadius: 16, marginBottom: 10, border: `2px solid ${isEditing ? "#7c3aed" : isCert ? "#fbbf24" : T.border}`, boxShadow: "0 2px 8px #0000000d", overflow: "hidden", opacity: deleting === item.id ? 0.5 : 1 }}>
                {/* 헤더 — 클릭하면 펼치기 */}
                <div onClick={() => { if (!isEditing) setExpanded(p => ({ ...p, [item.id]: !p[item.id] })); }}
                  style={{ padding: "12px 14px", cursor: isEditing ? "default" : "pointer", display: "flex", alignItems: "center", gap: 10 }}>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    {item.title && <div style={{ fontSize: 14, fontWeight: 800, color: T.text, marginBottom: 2 }}>{item.title}</div>}
                    {!isExp && !isEditing && previewText && (
                      <div style={{ fontSize: 12, color: T.muted, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{previewText}</div>
                    )}
                    <div style={{ fontSize: 11, color: T.muted, marginTop: 3 }}>
                      {formatDate(item.createdAt)}
                      {item.updatedAt && " · 수정됨"}
                      {fileList.length > 0 && ` · 📎 ${fileList.length}개`}
                    </div>
                  </div>
                  {!isEditing && <span style={{ fontSize: 14, color: T.muted }}>{isExp ? "▲" : "▼"}</span>}
                </div>

                {/* 펼친 상태 — 보기 or 수정 */}
                {(isExp || isEditing) && (
                  <div style={{ padding: "0 14px 14px", borderTop: `1px solid ${T.border}` }}>
                    {isEditing ? (
                      /* ── 인라인 수정 폼 ── */
                      <>
                        <input type="text" value={editTitle} onChange={e => setEditTitle(e.target.value)}
                          placeholder="제목 (선택)"
                          style={{ width: "100%", padding: "10px 12px", borderRadius: 10, border: `1px solid #7c3aed`, fontSize: 15, fontWeight: 700, color: T.text, background: T.bg, boxSizing: "border-box", fontFamily: "inherit", marginTop: 12, marginBottom: 8 }} />
                        <textarea value={editText} onChange={e => setEditText(e.target.value)}
                          style={{ width: "100%", minHeight: 140, padding: "10px 12px", borderRadius: 10, border: `1px solid #7c3aed`, fontSize: 14, color: T.text, background: T.bg, resize: "vertical", boxSizing: "border-box", fontFamily: "inherit", lineHeight: 1.7 }} />
                        {/* 기존 파일 */}
                        {fileList.length > 0 && <FileList files={fileList} item={item} />}
                        {/* 파일 추가 */}
                        <div style={{ marginTop: 10 }}>
                          <button onClick={() => editFileRef.current?.click()}
                            style={{ padding: "7px 14px", borderRadius: 10, border: `1px dashed ${T.border}`, background: T.bg, color: T.sub, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
                            📎 파일 추가
                          </button>
                          <input ref={editFileRef} type="file" accept="*/*" multiple style={{ display: "none" }}
                            onChange={e => setEditFiles(prev => [...prev, ...Array.from(e.target.files || [])])} />
                        </div>
                        <NewFileList files={editFiles} onRemove={i => setEditFiles(prev => prev.filter((_, j) => j !== i))} />
                        <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
                          <button onClick={() => { setEditId(null); setEditFiles([]); }}
                            style={{ flex: 1, padding: "10px 0", borderRadius: 10, border: `1px solid ${T.border}`, background: T.bg, color: T.text, fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                            취소
                          </button>
                          <button onClick={() => handleEditSave(item)} disabled={editSaving}
                            style={{ flex: 2, padding: "10px 0", borderRadius: 10, border: "none", background: "#7c3aed", color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                            {editSaving ? "저장 중..." : "✏️ 수정 완료"}
                          </button>
                        </div>
                      </>
                    ) : (
                      /* ── 보기 모드 ── */
                      <>
                        {item.text && (
                          <div style={{ fontSize: 14, color: T.text, lineHeight: 1.7, whiteSpace: "pre-wrap", paddingTop: 12, marginBottom: fileList.length > 0 ? 0 : 0 }}>
                            {item.text}
                          </div>
                        )}
                        {fileList.length > 0 && <FileList files={fileList} item={item} />}
                        <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
                          <button onClick={() => startEdit(item)}
                            style={{ flex: 1, padding: "9px 0", borderRadius: 10, border: `1px solid ${T.border}`, background: T.bg, color: T.text, fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                            ✏️ 수정
                          </button>
                          <button onClick={() => handleDelete(item)} disabled={deleting === item.id}
                            style={{ flex: 1, padding: "9px 0", borderRadius: 10, border: "none", background: "#fee2e2", color: "#b91c1c", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                            🗑 삭제
                          </button>
                        </div>
                      </>
                    )}
                  </div>
                )}
              </div>
            );
          })
        )}
      </div>

      {/* 이미지 전체화면 */}
      {preview && (
        <div onClick={() => setPreview(null)}
          style={{ position: "fixed", inset: 0, background: "#000000cc", zIndex: 9999, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
          <img src={preview.url} alt={preview.name} style={{ maxWidth: "100%", maxHeight: "90vh", borderRadius: 12, objectFit: "contain" }} />
          <button onClick={() => setPreview(null)}
            style={{ position: "absolute", top: 20, right: 20, background: "#ffffff30", border: "none", color: "#fff", fontSize: 24, cursor: "pointer", borderRadius: 10, padding: "4px 10px" }}>✕</button>
        </div>
      )}
    </div>
  );
}

// ── 관리자 화면 (라우터) ───────────────────────────────────────
function AdminScreen({ user, users, settings, records, leaves, notices, board, payslips, annual, leaveRequests, memberInfo, reads, reminders = [], scheduleEvents = [], contracts = [], onSaveRecord, onSaveLeave, onSaveUsers, onSaveSettings, onLogout }) {
  const [section, setSection] = useState(null);
  const back = () => { setSection(null); window.scrollTo(0,0); };
  if (!section) return <AdminHome user={user} onLogout={onLogout} onSection={s => { setSection(s); window.scrollTo(0,0); }} leaveRequests={leaveRequests} board={board} reads={reads} contracts={contracts} />;
  if (section === "attendance") return <><AdminAttendance users={users} settings={settings} records={records} leaves={leaves} leaveRequests={leaveRequests} onSaveRecord={onSaveRecord} onSaveLeave={onSaveLeave} onSaveSettings={onSaveSettings} onBack={back} /><FloatBack onClick={back} /></>;
  if (section === "wage") return <><AdminWage users={users} records={records} leaves={leaves} settings={settings} memberInfo={memberInfo} annual={annual} leaveRequests={leaveRequests} payslips={payslips} reads={reads} onBack={back} /><FloatBack onClick={back} /></>;
  if (section === "members") return <><AdminMembers users={users} annual={annual} leaveRequests={leaveRequests} memberInfo={memberInfo} onSaveUsers={onSaveUsers} onBack={back} /><FloatBack onClick={back} /></>;
  if (section === "annual") return <><AnnualScreen user={user} users={users} annual={annual} leaveRequests={leaveRequests} onBack={back} /><FloatBack onClick={back} /></>;
  if (section === "severance") return <><AdminSeverance users={users} memberInfo={memberInfo} annual={annual} onBack={back} /><FloatBack onClick={back} /></>;
  if (section === "notice") return <AdminSectionWrap title="📢 공지사항" color="#ea580c" onBack={back}><NoticeScreen user={user} users={users} notices={notices} reads={reads} /></AdminSectionWrap>;
  if (section === "board") return <AdminSectionWrap title="💬 게시판" color="#0891b2" onBack={back}><BoardScreen user={user} board={board} reads={reads} /></AdminSectionWrap>;
  if (section === "settings") return <><SettingsModal settings={settings} onSave={async s => { await onSaveSettings(s); back(); }} onClose={back} /></>;
  if (section === "schedule") return <AdminSectionWrap title="🗓 일정" color="#7c3aed" onBack={back}><AdminSchedule reminders={reminders} users={users} settings={settings} scheduleEvents={scheduleEvents} /></AdminSectionWrap>;
  if (section === "contract") return <><DocSection users={users} memberInfo={memberInfo} settings={settings} contracts={contracts} onBack={back} /><FloatBack onClick={back} /></>;
  if (section === "vault") return <><VaultSection onBack={back} /><FloatBack onClick={back} /></>;
  if (section === "insurance") return <><InsuranceSection users={users} memberInfo={memberInfo} onBack={back} /><FloatBack onClick={back} /></>;
  if (section === "education") return <><EducationSection users={users} reads={reads} onBack={back} /><FloatBack onClick={back} /></>;
  return null;
}

// ── 공지사항 ────────────────────────────────────────────────────
// ── 공지 + 게시판 통합 ────────────────────────────────────────────
function NoticeBoardScreen({ user, users, notices, board, reads }) {
  const [subTab, setSubTab] = useState("notice");
  return (
    <div>
      <div style={{ display: "flex", background: T.card, borderBottom: `1px solid ${T.border}` }}>
        {[["notice","📢 공지사항"],["board","💬 게시판"]].map(([key,label]) => (
          <button key={key} onClick={() => setSubTab(key)}
            style={{ flex: 1, padding: "12px 0", border: "none", background: "none", cursor: "pointer",
              fontWeight: subTab === key ? 800 : 500, fontSize: 13,
              color: subTab === key ? T.headerBg : T.muted,
              borderBottom: subTab === key ? `3px solid ${T.headerBg}` : "3px solid transparent",
              fontFamily: "inherit" }}>
            {label}
          </button>
        ))}
      </div>
      {subTab === "notice" && <NoticeScreen user={user} users={users} notices={notices} reads={reads} />}
      {subTab === "board"  && <BoardScreen  user={user} board={board} reads={reads} />}
    </div>
  );
}

function NoticeScreen({ user, users, notices, reads }) {
  const isAdmin = user.role === "admin";
  const members = users.filter(u => u.role === "member" && (!u.status || u.status === "active"));
  const [showWrite, setShowWrite] = useState(false);
  const [editTarget, setEditTarget] = useState(null);
  const [title, setTitle] = useState("");
  const [content, setContent] = useState("");
  const [recipient, setRecipient] = useState("all"); // "all" or userId
  const [file, setFile] = useState(null);
  const [uploading, setUploading] = useState(false);
  const [expanded, setExpanded] = useState(null);
  const [requireConfirm, setRequireConfirm] = useState(false);
  const [requireReply, setRequireReply] = useState(false); // 회신 요청
  const [replyType, setReplyType] = useState("both"); // text | file | both
  const [nudgeModal, setNudgeModal] = useState(null);
  const [nudgeSending, setNudgeSending] = useState(false);
  const [replyInputs, setReplyInputs] = useState({}); // { noticeId: { text, file } }
  const [replySending, setReplySending] = useState(null);
  const [noticeTab, setNoticeTab] = useState("manual"); // manual | auto

  // 내가 볼 수 있는 공지 필터
  const visibleNotices = notices.filter(n =>
    n.recipient === "all" || n.recipient === user.id || user.role === "admin"
  );

  // 관리자: 직접작성 vs 자동생성 분리
  const isAutoNotice = (n) => n.auto === true;
  const manualNotices = visibleNotices.filter(n => !isAutoNotice(n));
  const autoNotices = visibleNotices.filter(n => isAutoNotice(n));
  const displayNotices = isAdmin ? (noticeTab === "manual" ? manualNotices : autoNotices) : visibleNotices;

  const resetForm = () => { setTitle(""); setContent(""); setRecipient("all"); setFile(null); setShowWrite(false); setEditTarget(null); setRequireConfirm(false); setRequireReply(false); setReplyType("both"); };

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
    const data = { title: title.trim(), content: content.trim(), recipient, author: user.name, createdAt: new Date().toISOString(), requireConfirm, requireReply, replyType };
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
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: isAdmin ? 10 : 16 }}>
        <div style={{ fontSize: 16, fontWeight: 800, color: T.text }}>📢 공지사항</div>
        {isAdmin && !showWrite && noticeTab === "manual" && <button onClick={() => { resetForm(); setShowWrite(true); }} style={{ background: T.adminHeader, border: "none", color: "#fff", borderRadius: 10, padding: "8px 14px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>+ 작성</button>}
      </div>

      {/* 관리자 탭 */}
      {isAdmin && (
        <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
          <button onClick={() => { setNoticeTab("manual"); setExpanded(null); }}
            style={{ flex: 1, padding: "9px 0", borderRadius: 10, border: `2px solid ${noticeTab === "manual" ? T.adminHeader : T.border}`, background: noticeTab === "manual" ? T.adminHeader : T.bg, color: noticeTab === "manual" ? "#fff" : T.text, fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
            📢 작성 공지 {manualNotices.length > 0 && `(${manualNotices.length})`}
          </button>
          <button onClick={() => { setNoticeTab("auto"); setExpanded(null); }}
            style={{ flex: 1, padding: "9px 0", borderRadius: 10, border: `2px solid ${noticeTab === "auto" ? "#6b7280" : T.border}`, background: noticeTab === "auto" ? "#6b7280" : T.bg, color: noticeTab === "auto" ? "#fff" : T.text, fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
            🤖 자동 알림 {autoNotices.length > 0 && `(${autoNotices.length})`}
          </button>
        </div>
      )}

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
          {/* 읽음 확인 요청 */}
          <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", background: requireConfirm ? "#fffbeb" : T.bg, borderRadius: 10, border: `1px solid ${requireConfirm ? "#fbbf24" : T.border}`, marginBottom: 8, cursor: "pointer" }}
            onClick={() => setRequireConfirm(p => !p)}>
            <input type="checkbox" checked={requireConfirm} onChange={() => {}} style={{ width: 16, height: 16, cursor: "pointer" }} />
            <div>
              <div style={{ fontSize: 13, fontWeight: 700, color: T.text }}>📋 읽음 확인 요청</div>
              <div style={{ fontSize: 11, color: T.muted }}>팀원이 "확인했습니다" 버튼을 눌러야 완료됩니다</div>
            </div>
          </div>
          {/* 회신 요청 */}
          <div style={{ marginBottom: 10 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", background: requireReply ? "#f0f9ff" : T.bg, borderRadius: 10, border: `1px solid ${requireReply ? "#0891b2" : T.border}`, cursor: "pointer" }}
              onClick={() => setRequireReply(p => !p)}>
              <input type="checkbox" checked={requireReply} onChange={() => {}} style={{ width: 16, height: 16, cursor: "pointer" }} />
              <div>
                <div style={{ fontSize: 13, fontWeight: 700, color: T.text }}>💬 회신 요청</div>
                <div style={{ fontSize: 11, color: T.muted }}>팀원이 텍스트/파일로 회신할 수 있습니다</div>
              </div>
            </div>
            {requireReply && (
              <div style={{ display: "flex", gap: 6, marginTop: 8, paddingLeft: 4 }}>
                {[["text", "💬 텍스트"], ["file", "📎 파일"], ["both", "💬+📎 둘 다"]].map(([val, label]) => (
                  <button key={val} onClick={() => setReplyType(val)}
                    style={{ flex: 1, padding: "7px 0", borderRadius: 8, border: `1px solid ${replyType === val ? "#0891b2" : T.border}`, background: replyType === val ? "#0891b2" : T.bg, color: replyType === val ? "#fff" : T.text, fontSize: 11, fontWeight: 700, cursor: "pointer" }}>
                    {label}
                  </button>
                ))}
              </div>
            )}
          </div>
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

      {displayNotices.length === 0
        ? <div style={{ textAlign: "center", color: T.muted, padding: 40 }}>{isAdmin && noticeTab === "auto" ? "자동 알림이 없어요" : "공지사항이 없어요"}</div>
        : displayNotices.map(n => {
          // 확인 현황 계산
          const targetMembers = n.recipient === "all"
            ? members
            : members.filter(m => m.id === n.recipient);
          const confirmedIds = targetMembers.filter(m => reads[`${m.id}_confirm_${n.id}`]).map(m => m.id);
          const unconfirmedMembers = targetMembers.filter(m => !reads[`${m.id}_confirm_${n.id}`]);
          const myConfirmed = reads[`${user.id}_confirm_${n.id}`];
          // 회신 관련
          const myReply = reads[`${user.id}_reply_${n.id}`];
          const memberReplies = targetMembers.map(m => ({ member: m, reply: reads[`${m.id}_reply_${n.id}`] }));

          return (
          <div key={n.id} style={{ background: T.card, borderRadius: 14, padding: "14px 16px", marginBottom: 10, border: `1px solid ${isUnread(n) ? T.blue : T.border}` }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }} onClick={() => toggleExpanded(n.id)}>
              <div style={{ flex: 1 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 3 }}>
                  {isUnread(n) && <div style={{ width: 8, height: 8, borderRadius: "50%", background: T.blue, flexShrink: 0 }} />}
                  <div style={{ fontWeight: isUnread(n) ? 800 : 700, fontSize: 14, color: T.text }}>{n.title}</div>
                  {recipientLabel(n.recipient)}
                  {n.fileName && <Badge label="📎" color="gray" />}
                  {n.requireConfirm && <Badge label="📋 확인요청" color="yellow" />}
                  {n.requireReply && <Badge label="💬 회신요청" color="blue" />}
                </div>
                <div style={{ fontSize: 11, color: T.muted }}>
                  {n.author} · {n.createdAt?.slice(0,10)}
                  {n.requireConfirm && isAdmin && targetMembers.length > 0 && (
                    <span style={{ marginLeft: 8, fontWeight: 700, color: confirmedIds.length === targetMembers.length ? "#16a34a" : "#d97706" }}>
                      ✅ {confirmedIds.length}/{targetMembers.length}명 확인
                    </span>
                  )}
                </div>
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

                {/* 팀원: 확인했습니다 버튼 */}
                {!isAdmin && n.requireConfirm && (
                  <div style={{ marginBottom: 10 }}>
                    {myConfirmed ? (
                      <div style={{ padding: "10px 14px", background: "#dcfce7", borderRadius: 10, fontSize: 13, fontWeight: 700, color: "#16a34a" }}>
                        ✅ 확인완료 · {new Date(reads[`${user.id}_confirm_${n.id}`]?.readAt || "").toLocaleString("ko-KR")}
                      </div>
                    ) : (
                      <button onClick={async () => {
                        const key = `${user.id}_confirm_${n.id}`;
                        await setDoc(doc(db, COL_READS, key), { userId: user.id, type: "confirm", docId: n.id, readAt: new Date().toISOString() });
                        await addDoc(collection(db, COL_NOTICES), {
                          title: `✅ ${user.name}님 공지 확인 완료`,
                          content: `"${n.title}" 공지를 확인하였습니다.`,
                          recipient: "admin", author: user.name, auto: true, createdAt: new Date().toISOString(),
                        });
                        await sendPush({ title: `✅ ${user.name}님 공지 확인`, message: `"${n.title}" 공지를 확인하였습니다.`, targetUserId: "admin" });
                      }}
                        style={{ width: "100%", padding: "12px 0", borderRadius: 12, border: "none", background: "#0891b2", color: "#fff", fontSize: 14, fontWeight: 800, cursor: "pointer" }}>
                        ✅ 확인했습니다
                      </button>
                    )}
                  </div>
                )}

                {/* 팀원: 회신 입력 */}
                {!isAdmin && n.requireReply && (
                  <div style={{ marginBottom: 10 }}>
                    {myReply ? (
                      <div style={{ padding: "10px 14px", background: "#f0f9ff", borderRadius: 10, border: "1px solid #bae6fd" }}>
                        <div style={{ fontSize: 12, fontWeight: 700, color: "#0891b2", marginBottom: 4 }}>💬 회신 완료</div>
                        {myReply.text && <div style={{ fontSize: 13, color: T.text, marginBottom: 4 }}>{myReply.text}</div>}
                        {myReply.fileUrl && (
                          <a href={myReply.fileUrl} target="_blank" rel="noreferrer"
                            style={{ fontSize: 12, color: "#0891b2", fontWeight: 700, textDecoration: "none" }}>📎 {myReply.fileName}</a>
                        )}
                        <div style={{ fontSize: 11, color: T.muted, marginTop: 4 }}>{myReply.repliedAt ? new Date(myReply.repliedAt).toLocaleString("ko-KR") : ""}</div>
                      </div>
                    ) : (
                      <div style={{ padding: "12px 14px", background: T.bg, borderRadius: 10, border: `1px solid #0891b2` }}>
                        <div style={{ fontSize: 13, fontWeight: 700, color: "#0891b2", marginBottom: 10 }}>💬 회신하기</div>
                        {(n.replyType === "text" || n.replyType === "both" || !n.replyType) && (
                          <textarea
                            value={replyInputs[n.id]?.text || ""}
                            onChange={e => setReplyInputs(p => ({ ...p, [n.id]: { ...p[n.id], text: e.target.value } }))}
                            placeholder="회신 내용을 입력하세요"
                            style={{ width: "100%", padding: "10px 12px", borderRadius: 8, border: `1px solid ${T.border}`, fontSize: 13, resize: "vertical", minHeight: 80, boxSizing: "border-box", fontFamily: "inherit", marginBottom: 8 }} />
                        )}
                        {(n.replyType === "file" || n.replyType === "both") && (
                          <div style={{ marginBottom: 8 }}>
                            {replyInputs[n.id]?.file ? (
                              <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 10px", background: "#fff", borderRadius: 8, border: `1px solid ${T.border}` }}>
                                <span>📎</span>
                                <span style={{ fontSize: 12, flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{replyInputs[n.id].file.name}</span>
                                <button onClick={() => setReplyInputs(p => ({ ...p, [n.id]: { ...p[n.id], file: null } }))}
                                  style={{ background: "none", border: "none", color: "#b91c1c", cursor: "pointer" }}>✕</button>
                              </div>
                            ) : (
                              <label style={{ display: "flex", alignItems: "center", gap: 8, padding: "10px 12px", borderRadius: 8, border: `1px dashed ${T.border}`, cursor: "pointer", fontSize: 12, color: T.muted, fontWeight: 600 }}>
                                📎 파일 첨부
                                <input type="file" style={{ display: "none" }}
                                  onChange={e => setReplyInputs(p => ({ ...p, [n.id]: { ...p[n.id], file: e.target.files?.[0] } }))} />
                              </label>
                            )}
                          </div>
                        )}
                        <button onClick={async () => {
                          const input = replyInputs[n.id] || {};
                          if (!input.text?.trim() && !input.file) { alert("내용 또는 파일을 입력해주세요."); return; }
                          setReplySending(n.id);
                          try {
                            let fileUrl = null, fileName = null;
                            if (input.file) {
                              const sRef = ref(storage, `notice_replies/${n.id}/${user.id}_${Date.now()}_${input.file.name}`);
                              await uploadBytes(sRef, input.file);
                              fileUrl = await getDownloadURL(sRef);
                              fileName = input.file.name;
                            }
                            const replyData = { text: input.text?.trim() || "", fileUrl, fileName, repliedAt: new Date().toISOString(), userId: user.id, userName: user.name };
                            await setDoc(doc(db, COL_READS, `${user.id}_reply_${n.id}`), { ...replyData, type: "reply", docId: n.id });
                            await addDoc(collection(db, COL_NOTICES), {
                              title: `💬 ${user.name}님 회신`,
                              content: `"${n.title}" 공지에 회신하였습니다.${input.text?.trim() ? `\n내용: ${input.text.trim()}` : ""}${fileName ? `\n파일: ${fileName}` : ""}`,
                              recipient: "admin", author: user.name, auto: true, createdAt: new Date().toISOString(),
                            });
                            await sendPush({ title: `💬 ${user.name}님 회신`, message: `"${n.title}" 공지에 회신하였습니다.`, targetUserId: "admin" });
                            setReplyInputs(p => ({ ...p, [n.id]: {} }));
                            alert("회신이 완료되었습니다.");
                          } catch(e) { alert("회신 실패: " + e.message); }
                          setReplySending(null);
                        }} disabled={replySending === n.id}
                          style={{ width: "100%", padding: "11px 0", borderRadius: 10, border: "none", background: "#0891b2", color: "#fff", fontSize: 13, fontWeight: 800, cursor: "pointer", opacity: replySending === n.id ? 0.6 : 1 }}>
                          {replySending === n.id ? "전송 중..." : "💬 회신하기"}
                        </button>
                      </div>
                    )}
                  </div>
                )}

                {/* 관리자: 회신 현황 */}
                {isAdmin && n.requireReply && targetMembers.length > 0 && (
                  <div style={{ background: T.bg, borderRadius: 10, padding: "10px 12px", marginBottom: 10 }}>
                    <div style={{ fontSize: 12, fontWeight: 700, color: T.text, marginBottom: 8 }}>💬 회신 현황</div>
                    {memberReplies.map(({ member: m, reply }) => (
                      <div key={m.id} style={{ padding: "8px 10px", borderRadius: 8, background: reply ? "#f0f9ff" : "#fff", border: `1px solid ${reply ? "#bae6fd" : T.border}`, marginBottom: 6 }}>
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                          <span style={{ fontSize: 13, fontWeight: 700, color: T.text }}>{m.name}</span>
                          <span style={{ fontSize: 11, fontWeight: 700, color: reply ? "#0891b2" : T.muted }}>{reply ? "✅ 회신완료" : "⏳ 미회신"}</span>
                        </div>
                        {reply?.text && <div style={{ fontSize: 12, color: T.text, marginTop: 4, lineHeight: 1.5 }}>{reply.text}</div>}
                        {reply?.fileUrl && (
                          <a href={reply.fileUrl} target="_blank" rel="noreferrer"
                            style={{ display: "inline-block", marginTop: 4, fontSize: 12, color: "#0891b2", fontWeight: 700, textDecoration: "none" }}>📎 {reply.fileName}</a>
                        )}
                        {reply?.repliedAt && <div style={{ fontSize: 10, color: T.muted, marginTop: 4 }}>{new Date(reply.repliedAt).toLocaleString("ko-KR")}</div>}
                      </div>
                    ))}
                  </div>
                )}

                {/* 관리자: 확인 현황 + 독촉 버튼 */}
                {isAdmin && n.requireConfirm && targetMembers.length > 0 && (
                  <div style={{ background: T.bg, borderRadius: 10, padding: "10px 12px", marginBottom: 10 }}>
                    <div style={{ fontSize: 12, fontWeight: 700, color: T.text, marginBottom: 8 }}>📋 확인 현황</div>
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 8 }}>
                      {targetMembers.map(m => (
                        <span key={m.id} style={{ fontSize: 12, padding: "4px 10px", borderRadius: 20, fontWeight: 600,
                          background: reads[`${m.id}_confirm_${n.id}`] ? "#dcfce7" : "#fee2e2",
                          color: reads[`${m.id}_confirm_${n.id}`] ? "#16a34a" : "#b91c1c" }}>
                          {reads[`${m.id}_confirm_${n.id}`] ? "✅" : "⏳"} {m.name}
                        </span>
                      ))}
                    </div>
                    {unconfirmedMembers.length > 0 && (
                      <button onClick={async () => {
                        for (const m of unconfirmedMembers) {
                          await sendPush({ title: `📣 확인 요청`, message: `"${n.title}" 공지를 아직 확인하지 않으셨습니다. 확인 부탁드립니다.`, targetUserId: m.id });
                        }
                        alert(`${unconfirmedMembers.map(m => m.name).join(", ")}님께 독촉 알림을 보냈습니다.`);
                      }}
                        style={{ width: "100%", padding: "9px 0", borderRadius: 10, border: "none", background: "#d97706", color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                        📣 미확인자 독촉 ({unconfirmedMembers.length}명)
                      </button>
                    )}
                  </div>
                )}

                {isAdmin && (
                  <div>
                    {/* 독촉 알림 버튼 — 작성 공지에만 */}
                    {!editTarget && !n.auto && (() => {
                      const targets = n.recipient === "all" ? members : members.filter(m => m.id === n.recipient);
                      if (targets.length === 0) return null;
                      if (n.nudgeDone) return (
                        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "8px 12px", background: "#f0fdf4", borderRadius: 10, border: "1px solid #86efac", marginBottom: 8 }}>
                          <span style={{ fontSize: 12, fontWeight: 700, color: "#16a34a" }}>✅ 독촉 완료 처리됨</span>
                          <button onClick={async () => await setDoc(doc(db, COL_NOTICES, n.id), { ...n, nudgeDone: false }, { merge: true })}
                            style={{ background: "none", border: "none", color: T.muted, fontSize: 11, cursor: "pointer", textDecoration: "underline" }}>되돌리기</button>
                        </div>
                      );
                      return (
                        <div style={{ display: "flex", gap: 6, marginBottom: 8 }}>
                          <button onClick={() => setNudgeModal({ notice: n, targets, selected: targets.map(m => m.id) })}
                            style={{ flex: 1, padding: "9px 0", borderRadius: 10, border: "none", background: "#fff7ed", color: "#d97706", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                            📣 독촉 알림 발송
                          </button>
                          <button onClick={async () => {
                            if (!window.confirm("독촉 완료 처리하면 독촉 버튼이 숨겨집니다. 계속할까요?")) return;
                            await setDoc(doc(db, COL_NOTICES, n.id), { ...n, nudgeDone: true }, { merge: true });
                          }}
                            style={{ padding: "9px 12px", borderRadius: 10, border: `1px solid ${T.border}`, background: T.bg, color: T.muted, fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
                            완료
                          </button>
                        </div>
                      );
                    })()}
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
                        <button onClick={() => { if (window.confirm(`"${n.title}" 공지사항을 삭제할까요?`)) del(n); }} style={{ background: T.redBg, border: "none", color: T.red, borderRadius: 8, padding: "5px 12px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>삭제</button>
                      </div>
                    )}
                  </div>
                )}
              </div>
            )}
          </div>
          );
        })
      }
      {/* 독촉 알림 선택 모달 */}
      {nudgeModal && (
        <div style={{ position: "fixed", inset: 0, background: "#00000066", zIndex: 999, display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
          <div style={{ background: "#fff", borderRadius: 20, padding: 24, width: "100%", maxWidth: 340 }}>
            <div style={{ fontSize: 18, textAlign: "center", marginBottom: 6 }}>📣</div>
            <div style={{ fontSize: 15, fontWeight: 800, color: T.text, textAlign: "center", marginBottom: 4 }}>독촉 알림 발송</div>
            <div style={{ fontSize: 12, color: T.muted, textAlign: "center", marginBottom: 16, lineHeight: 1.6 }}>
              "{nudgeModal.notice.title}"
            </div>
            <div style={{ marginBottom: 16 }}>
              {nudgeModal.targets.map(m => {
                const checked = nudgeModal.selected.includes(m.id);
                return (
                  <div key={m.id} onClick={() => setNudgeModal(p => ({
                    ...p,
                    selected: checked ? p.selected.filter(id => id !== m.id) : [...p.selected, m.id]
                  }))}
                    style={{ display: "flex", alignItems: "center", gap: 12, padding: "12px 14px", borderRadius: 10, border: `1px solid ${checked ? "#d97706" : T.border}`, background: checked ? "#fff7ed" : T.bg, marginBottom: 8, cursor: "pointer" }}>
                    <div style={{ width: 20, height: 20, borderRadius: 6, border: `2px solid ${checked ? "#d97706" : T.border}`, background: checked ? "#d97706" : "#fff", display: "flex", alignItems: "center", justifyContent: "center" }}>
                      {checked && <span style={{ color: "#fff", fontSize: 13, fontWeight: 800 }}>✓</span>}
                    </div>
                    <div style={{ width: 32, height: 32, borderRadius: "50%", background: "#d97706", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14, fontWeight: 800, color: "#fff" }}>{m.name[0]}</div>
                    <span style={{ fontSize: 14, fontWeight: 700, color: T.text }}>{m.name}</span>
                  </div>
                );
              })}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              <button onClick={() => setNudgeModal(null)}
                style={{ padding: "12px 0", borderRadius: 12, border: `1px solid ${T.border}`, background: T.bg, color: T.text, fontSize: 14, fontWeight: 700, cursor: "pointer" }}>취소</button>
              <button onClick={async () => {
                if (nudgeModal.selected.length === 0) { alert("보낼 대상을 선택해주세요."); return; }
                setNudgeSending(true);
                for (const id of nudgeModal.selected) {
                  await sendPush({
                    title: `📣 업무 독촉`,
                    message: `"${nudgeModal.notice.title}" 공지 내용을 아직 처리하지 않으셨습니다. 확인 부탁드립니다.`,
                    targetUserId: id
                  });
                }
                setNudgeSending(false);
                setNudgeModal(null);
                alert(`${nudgeModal.selected.length}명에게 독촉 알림을 보냈습니다.`);
              }} disabled={nudgeSending || nudgeModal.selected.length === 0}
                style={{ padding: "12px 0", borderRadius: 12, border: "none", background: nudgeModal.selected.length > 0 ? "#d97706" : "#e5e7eb", color: nudgeModal.selected.length > 0 ? "#fff" : T.muted, fontSize: 14, fontWeight: 700, cursor: nudgeModal.selected.length > 0 ? "pointer" : "default" }}>
                {nudgeSending ? "발송 중..." : `발송 (${nudgeModal.selected.length}명)`}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
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
      author: user.name, auto: true, userId: user.id, createdAt: new Date().toISOString()
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
        {!showWrite && <button onClick={() => setShowWrite(!showWrite)} style={{ background: T.adminHeader, border: "none", color: "#fff", borderRadius: 10, padding: "8px 14px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>+ 글쓰기</button>}
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
                      <button onClick={() => { if (window.confirm(`"${b.title}" 게시글을 삭제할까요?`)) del(b.id); }} style={{ background: T.redBg, border: "none", color: T.red, borderRadius: 8, padding: "5px 12px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>삭제</button>
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
  const kstNow = new Date(new Date().getTime() + 9*60*60*1000);
  const [selectedYear, setSelectedYear] = useState(kstNow.getFullYear());

  const members = users.filter(u => u.role === "member" && (!u.status || u.status === "active"));
  const allPayslips = isAdmin ? payslips : payslips.filter(p => p.userId === user.id);
  // 연도 필터 적용
  const myPayslips = allPayslips.filter(p => (p.month || "").startsWith(String(selectedYear)));

  // 선택 가능한 연도 목록
  const years = [...new Set(allPayslips.map(p => (p.month || "").slice(0, 4)).filter(Boolean))].sort((a, b) => b - a);
  if (!years.includes(String(selectedYear))) years.unshift(String(selectedYear));

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
      // PDF 버튼 임시 숨기기
      const pdfBtn = document.getElementById(`pdf-btn-${p.id}`);
      if (pdfBtn) pdfBtn.style.visibility = "hidden";
      const canvas = await html2canvas(el, { scale: 3, useCORS: true, backgroundColor: "#ffffff" });
      if (pdfBtn) pdfBtn.style.visibility = "visible";
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

      {/* 연도 선택 */}
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16, background: T.card, borderRadius: 12, padding: "10px 14px", border: `1px solid ${T.border}` }}>
        <button onClick={() => setSelectedYear(y => y - 1)}
          style={{ background: "none", border: "none", fontSize: 18, cursor: "pointer", color: T.text, fontWeight: 700 }}>‹</button>
        <div style={{ flex: 1, textAlign: "center", fontSize: 16, fontWeight: 800, color: T.text }}>{selectedYear}년</div>
        <button onClick={() => { if (selectedYear < kstNow.getFullYear()) setSelectedYear(y => y + 1); }}
          style={{ background: "none", border: "none", fontSize: 18, cursor: "pointer", color: selectedYear >= kstNow.getFullYear() ? T.muted : T.text, fontWeight: 700 }}>›</button>
      </div>

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

                  {isAdmin && (
                    p.paidConfirm
                      ? <span style={{ background: "#dcfce7", color: "#16a34a", borderRadius: 8, padding: "5px 10px", fontSize: 11, fontWeight: 700 }}>✓ 송금완료</span>
                      : <button onClick={async e => {
                          e.stopPropagation();
                          if (!window.confirm(`${member?.name}님께 ${monthLabel(p.month)} 급여 입금 알림을 보낼까요?`)) return;
                          await setDoc(doc(db, COL_PAYSLIPS, p.id), { ...p, paidConfirm: true, paidAt: new Date().toISOString() });
                          await sendPush({ title: "💰 급여 입금 안내", message: `${monthLabel(p.month)} 급여가 입금되었습니다. 확인해주세요.`, targetUserId: p.userId });
                        }}
                          style={{ background: "#fef3c7", border: "1px solid #fde68a", color: "#92400e", borderRadius: 8, padding: "5px 10px", fontSize: 11, cursor: "pointer", fontWeight: 700 }}>💰 송금완료 알림</button>
                  )}

                  {isAdmin && <button onClick={e => { e.stopPropagation(); if (window.confirm(`${users.find(u => u.id === p.userId)?.name} · ${monthLabel(p.month)} 명세서를 삭제할까요?\n삭제 후 복구할 수 없어요.`)) deleteDoc(doc(db, COL_PAYSLIPS, p.id)); }}
                    style={{ background: T.redBg, border: "none", color: T.red, borderRadius: 8, padding: "5px 8px", fontSize: 11, cursor: "pointer", fontWeight: 700 }}>삭제</button>}
                  <span style={{ color: T.muted, fontSize: 14 }}>{isOpen ? "▲" : "▼"}</span>
                </div>
              </div>
              {isOpen && w && (
                <div id={`payslip-content-${p.id}`} style={{ background: T.bg, padding: "12px 16px", borderTop: `1px solid ${T.border}` }}>
                  {/* PDF 헤더 */}
                  <div style={{ textAlign: "center", marginBottom: 14, paddingBottom: 12, borderBottom: `2px solid ${T.border}` }}>
                    <div style={{ fontSize: 18, fontWeight: 800, color: T.text, letterSpacing: 2 }}>급 여 명 세 서</div>
                    <div style={{ fontSize: 12, color: T.muted, marginTop: 4 }}>
                      {isAdmin ? `${users.find(u => u.id === p.userId)?.name} · ` : ""}{monthLabel(p.month)} · 지급일 {w.payDate || "-"}
                    </div>
                  </div>
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
                  <button id={`pdf-btn-${p.id}`} onClick={() => downloadPDF(p)} disabled={pdfLoading === p.id}
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
  const members = users.filter(u => u.role === "member" && (!u.status || u.status === "active"));
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
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
            <div style={{ fontSize: 13, color: T.muted, fontWeight: 600 }}>신청 내역</div>
            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <button onClick={() => setSelectedYear(y => String(Number(y) - 1))}
                style={{ background: "none", border: "none", fontSize: 16, cursor: "pointer", color: T.text, fontWeight: 700 }}>‹</button>
              <span style={{ fontSize: 13, fontWeight: 800, color: T.text }}>{selectedYear}년</span>
              <button onClick={() => { if (Number(selectedYear) < Number(thisYear)) setSelectedYear(y => String(Number(y) + 1)); }}
                style={{ background: "none", border: "none", fontSize: 16, cursor: "pointer", color: Number(selectedYear) >= Number(thisYear) ? T.muted : T.text, fontWeight: 700 }}>›</button>
            </div>
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
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", margin: "16px 0 8px" }}>
            <div style={{ fontSize: 13, color: T.muted, fontWeight: 600 }}>연차 신청 목록</div>
            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <button onClick={() => setSelectedYear(y => String(Number(y) - 1))}
                style={{ background: "none", border: "none", fontSize: 16, cursor: "pointer", color: T.text, fontWeight: 700 }}>‹</button>
              <span style={{ fontSize: 13, fontWeight: 800, color: T.text }}>{selectedYear}년</span>
              <button onClick={() => { if (Number(selectedYear) < Number(thisYear)) setSelectedYear(y => String(Number(y) + 1)); }}
                style={{ background: "none", border: "none", fontSize: 16, cursor: "pointer", color: Number(selectedYear) >= Number(thisYear) ? T.muted : T.text, fontWeight: 700 }}>›</button>
            </div>
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
function TabBar({ tab, setTab, isAdmin, leaveRequests, notices, board, payslips, user, reads, contracts = [] }) {
  const pendingCount = leaveRequests.filter(r => r.status === "대기").length;

  const unreadCount = (items, type) => {
    if (isAdmin || !user || !reads) return 0;
    return items.filter(item => !reads[`${user.id}_${type}_${item.id}`] &&
      (type !== "notice" || item.recipient === "all" || item.recipient === user.id) &&
      (type !== "board" || item.userId !== user.id)
    ).length;
  };

  const unreadNotice = unreadCount(notices, "notice") + unreadCount(board, "board");
  const unreadPayslip = unreadCount(payslips.filter(p => p.userId === user?.id), "payslip");
  const contractBadge = contracts.filter(c => c.userId === user?.id && c.status === "sent").length;

  const tabs = [
    ["att",      "🏠", "출퇴근", 0],
    ["notice",   "📢", "공지",   unreadNotice],
    ["annual",   "📅", "연차",   isAdmin ? pendingCount : 0],
    ["payslip",  "💰", "명세서", unreadPayslip],
    ["contract", "📄", "문서함", contractBadge],
    ["schedule", "🗓", "일정",   0],
    ["education","🎓", "교육",   0],
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
function App({ users, settings, records, leaves, notices, board, payslips, annual, leaveRequests, memberInfo, reads, reminders = [], scheduleEvents = [], contracts = [], onSaveUsers, onSaveSettings, onSaveRecord, onSaveLeave }) {
  const [user, setUser] = useState(null);
  const [userLoaded, setUserLoaded] = useState(false);
  
  useEffect(() => {
    if (users && users.length > 0 && !userLoaded) {
      setUserLoaded(true);
      try {
        const saved = localStorage.getItem("loggedInUser");
        if (saved) {
          const savedUser = JSON.parse(saved);
          const freshUser = users.find(u => u.id === savedUser.id);
          if (freshUser) setUser(freshUser);
          else localStorage.removeItem("loggedInUser");
        }
      } catch { localStorage.removeItem("loggedInUser"); }
    }
  }, [users]);

  const setUserWithStorage = (u) => {
    if (u) localStorage.setItem("loggedInUser", JSON.stringify(u));
    else localStorage.removeItem("loggedInUser");
    setUser(u);
  };
  const [tab, setTab] = useState("att");

  if (!user) return <LoginScreen users={users} onLogin={setUserWithStorage} onUpdateUsers={onSaveUsers} />;

  const isAdmin = user.role === "admin";

  // 관리자는 대문+섹션 구조 (탭바 없음)
  if (isAdmin) return (
    <AdminScreen user={user} users={users} settings={settings} records={records} leaves={leaves}
      notices={notices} board={board} payslips={payslips} annual={annual} leaveRequests={leaveRequests} memberInfo={memberInfo} reads={reads}
      reminders={reminders} scheduleEvents={scheduleEvents} contracts={contracts}
      onSaveRecord={onSaveRecord} onSaveLeave={onSaveLeave}
      onSaveUsers={onSaveUsers} onSaveSettings={onSaveSettings}
      onLogout={() => { setUserWithStorage(null); setTab("att"); }} />
  );

  // 팀원은 탭바 구조
  return (
    <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'Noto Sans KR',sans-serif", paddingBottom: 80 }}>
      {tab === "att" && (
        <MemberScreen user={user} settings={settings} records={records} leaves={leaves}
          scheduleEvents={scheduleEvents}
          onSaveRecord={onSaveRecord} onLogout={() => { setUserWithStorage(null); setTab("att"); }} />
      )}
      {tab !== "att" && <FloatBack onClick={() => setTab("att")} />}
      {tab === "notice" && (
        <>
          <div style={{ background: T.headerBg, paddingTop: "calc(18px + env(safe-area-inset-top))", paddingBottom: "14px", paddingLeft: "16px", paddingRight: "16px" }}>
            <div style={{ fontSize: 11, color: "#ffffff50", letterSpacing: 3 }}>ATTENDANCE</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>공지</div>
          </div>
          <NoticeBoardScreen user={user} users={users} notices={notices} board={board} reads={reads} />
        </>
      )}
      {tab === "annual" && (
        <>
          <div style={{ background: T.headerBg, paddingTop: "calc(18px + env(safe-area-inset-top))", paddingBottom: "14px", paddingLeft: "16px", paddingRight: "16px" }}>
            <div style={{ fontSize: 11, color: "#ffffff50", letterSpacing: 3 }}>ATTENDANCE</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>연차</div>
          </div>
          <AnnualScreen user={user} users={users} annual={annual} leaveRequests={leaveRequests} />
        </>
      )}
      {tab === "payslip" && (
        <>
          <div style={{ background: T.headerBg, paddingTop: "calc(18px + env(safe-area-inset-top))", paddingBottom: "14px", paddingLeft: "16px", paddingRight: "16px" }}>
            <div style={{ fontSize: 11, color: "#ffffff50", letterSpacing: 3 }}>ATTENDANCE</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>급여명세서</div>
          </div>
          <PayslipScreen user={user} users={users} payslips={payslips} reads={reads} />
        </>
      )}
      {tab === "contract" && (
        <>
          <div style={{ background: T.headerBg, paddingTop: "calc(18px + env(safe-area-inset-top))", paddingBottom: "14px", paddingLeft: "16px", paddingRight: "16px" }}>
            <div style={{ fontSize: 11, color: "#ffffff50", letterSpacing: 3 }}>ATTENDANCE</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>📄 근로계약서</div>
          </div>
          <ContractViewScreen user={user} contracts={contracts} />
        </>
      )}
      {tab === "schedule" && (
        <>
          <div style={{ background: T.headerBg, paddingTop: "calc(18px + env(safe-area-inset-top))", paddingBottom: "14px", paddingLeft: "16px", paddingRight: "16px" }}>
            <div style={{ fontSize: 11, color: "#ffffff50", letterSpacing: 3 }}>ATTENDANCE</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#fff" }}>🗓 일정</div>
          </div>
          <div style={{ padding: 16 }}>
            <MemberScheduleCalendar settings={settings} scheduleEvents={scheduleEvents} userId={user?.id} />
          </div>
        </>
      )}
      {tab === "education" && (
        <MemberEducationTab user={user} reads={reads} />
      )}
      <TabBar tab={tab} setTab={t => { setTab(t); window.scrollTo(0, 0); }} isAdmin={isAdmin} leaveRequests={leaveRequests} notices={notices} board={board} payslips={payslips} user={user} reads={reads} contracts={contracts} />
    </div>
  );
}

export default AppLoader;
