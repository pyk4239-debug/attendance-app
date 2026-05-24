import https from 'https';

const FIREBASE_API_KEY = process.env.FIREBASE_API_KEY;
const FIREBASE_PROJECT_ID = process.env.FIREBASE_PROJECT_ID;
const ONESIGNAL_APP_ID = process.env.ONESIGNAL_APP_ID;
const ONESIGNAL_API_KEY = process.env.ONESIGNAL_API_KEY;

function httpsGet(url) {
  return new Promise((resolve, reject) => {
    https.get(url, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => resolve(JSON.parse(data)));
    }).on('error', reject);
  });
}

function httpsPost(options, body) {
  return new Promise((resolve, reject) => {
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => resolve(JSON.parse(data)));
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

function parseField(field) {
  if (!field) return null;
  if (field.stringValue !== undefined) return field.stringValue;
  if (field.integerValue !== undefined) return Number(field.integerValue);
  if (field.booleanValue !== undefined) return field.booleanValue;
  if (field.arrayValue) return (field.arrayValue.values || []).map(parseField);
  if (field.mapValue) return parseFields(field.mapValue.fields || {});
  return null;
}

function parseFields(fields) {
  const result = {};
  for (const [k, v] of Object.entries(fields)) result[k] = parseField(v);
  return result;
}

async function fetchDoc(collection, docId) {
  const url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/${collection}/${docId}?key=${FIREBASE_API_KEY}`;
  const res = await httpsGet(url);
  return res.fields ? parseFields(res.fields) : null;
}

async function fetchCollection(collection) {
  const url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/${collection}?key=${FIREBASE_API_KEY}&pageSize=300`;
  const res = await httpsGet(url);
  return (res.documents || []).map(doc => ({
    id: doc.name.split('/').pop(),
    ...parseFields(doc.fields || {})
  }));
}

async function sendPush(title, message, userIds) {
  if (!userIds || userIds.length === 0) return;
  const filters = userIds.map((id, i) => [
    ...(i > 0 ? [{ operator: 'OR' }] : []),
    { field: 'tag', key: 'userId', relation: '=', value: id }
  ]).flat();

  const body = JSON.stringify({
    app_id: ONESIGNAL_APP_ID,
    headings: { en: title, ko: title },
    contents: { en: message, ko: message },
    filters
  });

  const options = {
    hostname: 'onesignal.com',
    path: '/api/v1/notifications',
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Basic ${ONESIGNAL_API_KEY}`,
      'Content-Length': Buffer.byteLength(body)
    }
  };

  const result = await httpsPost(options, body);
  console.log('발송 결과:', JSON.stringify(result));
}

async function main() {
  const now = new Date();
  const kst = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  const hh = kst.getUTCHours();
  const mm = kst.getUTCMinutes();
  const today = kst.toISOString().slice(0, 10);
  const dayOfWeek = kst.getUTCDay();

  console.log(`KST: ${hh}:${String(mm).padStart(2, '0')} (${today}) 요일:${dayOfWeek}`);

  if (dayOfWeek === 0 || dayOfWeek === 6) { console.log('주말'); return; }

  const settings = await fetchDoc('app', 'settings');
  if (!settings) { console.log('settings 없음'); return; }

  const workStart = settings.workStart || '09:00';
  const workEnd = settings.workEnd || '18:00';
  const holidays = settings.holidays || [];

  console.log(`출근: ${workStart}, 퇴근: ${workEnd}`);

  // 날짜가 공휴일/주말인지 확인하는 헬퍼
  const isHolidayDate = (dateStr) => {
    const d = new Date(dateStr + 'T00:00:00Z');
    const dow = d.getUTCDay();
    if (dow === 0 || dow === 6) return true;
    return holidays.some(h => (typeof h === 'string' ? h : h.date) === dateStr);
  };

  // 특정 날짜의 리마인더 발송 대상 날짜 계산 (전날 대신 발송 로직)
  // originalDate의 리마인더를 오늘 발송해야 하는지 확인
  const shouldSendToday = (r) => {
    const currentHHMM = `${String(hh).padStart(2,'0')}:${String(mm).padStart(2,'0')}`;
    if (r.time !== currentHHMM) return false;

    // 오늘이 공휴일/주말이면 리마인더 발송 안 함 (단, 전날대신 발송 옵션 있는 것들은 아래서 처리)
    // 반복 조건 체크할 날짜 결정
    // "전날 대신 발송" 옵션이 있으면: 내일~최대 7일 후까지 공휴일인 날의 리마인더를 오늘 발송
    if (r.sendBeforeHoliday) {
      // 내일부터 최대 7일 후까지 공휴일인지 확인, 첫 번째 공휴일의 전날(=오늘)이면 발송
      for (let i = 1; i <= 7; i++) {
        const targetDate = new Date(kst.getTime() + i * 24 * 60 * 60 * 1000);
        const targetStr = targetDate.toISOString().slice(0, 10);
        const targetDow = targetDate.getUTCDay();
        if (!isHolidayDate(targetStr)) {
          // 첫 번째 평일이 i일 후면, i-1일 전(오늘)이 전날
          if (i === 1) break; // 내일이 평일이면 해당 없음
          // i>1 이면 오늘이 공휴일 직전 평일인지 확인
          // 오늘은 이미 평일(주말/공휴일 return 안 했으면)
          // targetDate가 첫 번째 평일이고 i>1이면 중간에 공휴일 있음
          // 그 공휴일들의 날짜 중 반복 조건 맞는 게 있으면 오늘 발송
          for (let j = 1; j < i; j++) {
            const holidayDate = new Date(kst.getTime() + j * 24 * 60 * 60 * 1000);
            const holidayStr = holidayDate.toISOString().slice(0, 10);
            const holidayDow = holidayDate.getUTCDay();
            if (matchesRepeat(r, holidayStr, holidayDow)) return true;
          }
          break;
        }
      }
    }

    // 오늘 발송 조건 (오늘이 공휴일이면 해당 없음 - 이미 위에서 걸림)
    if (isHolidayDate(today)) return false;
    return matchesRepeat(r, today, dayOfWeek);
  };

  const matchesRepeat = (r, dateStr, dow) => {
    const date = new Date(dateStr + 'T00:00:00Z');
    const dom = date.getUTCDate();
    if (r.repeat === 'weekly' && r.weekDay !== dow) return false;
    if (r.repeat === 'monthly' && r.monthDay !== dom) return false;
    return true;
  };

  const todayIsHoliday = isHolidayDate(today);
  if (todayIsHoliday) { console.log('공휴일 - 출퇴근 알림 생략'); }

  if (!todayIsHoliday) {
    const [startH, startM] = workStart.split(':').map(Number);
    const [endH, endM] = workEnd.split(':').map(Number);
    const currentMinutes = hh * 60 + mm;
    const startMinutes = startH * 60 + startM;
    const endMinutes = endH * 60 + endM;
    const isCheckinAlert = currentMinutes === startMinutes - 5;
    const isCheckoutAlert = currentMinutes === endMinutes + 1;

    console.log(`현재: ${currentMinutes}분, 출근알림: ${isCheckinAlert}, 퇴근알림: ${isCheckoutAlert}`);

    const users = await fetchCollection('users');
    const members = users.filter(u => u.role === 'member');
    console.log('팀원:', members.map(u => u.name));

    const records = await fetchCollection('records');
    const recordMap = {};
    for (const rec of records) recordMap[rec.id] = rec;

    const leaves = await fetchCollection('leaves');
    const todayLeaves = leaves.filter(l => l.date === today && !l.deleted);
    const leaveSet = new Set(todayLeaves.map(l => l.userId));
    const amLeaveSet = new Set(todayLeaves.filter(l => l.type && l.type.includes('오전')).map(l => l.userId));
    const pmLeaveSet = new Set(todayLeaves.filter(l => l.type && l.type.includes('오후')).map(l => l.userId));

    if (isCheckinAlert) {
      const absentIds = members.filter(u => {
        const rec = recordMap[`${u.id}_${today}`];
        if (leaveSet.has(u.id)) return false;
        if (amLeaveSet.has(u.id)) return false;
        if (rec?.in) return false;
        return true;
      }).map(u => u.id);
      console.log('출근 알림 대상:', absentIds);
      if (absentIds.length > 0) await sendPush('⏰ 출근 시간 알림', `출근 시간 ${workStart} 5분 전입니다. 출근 기록을 해주세요!`, absentIds);
    }

    if (isCheckoutAlert) {
      const notOutIds = members.filter(u => {
        const rec = recordMap[`${u.id}_${today}`];
        if (leaveSet.has(u.id)) return false;
        if (pmLeaveSet.has(u.id)) return false;
        if (!rec?.in) return false;
        if (rec?.out) return false;
        const outings = rec?.outing || [];
        if (outings.length > 0 && !outings[outings.length - 1].in) return false;
        return true;
      }).map(u => u.id);
      console.log('퇴근 알림 대상:', notOutIds);
      if (notOutIds.length > 0) await sendPush('🏠 퇴근 시간 알림', `퇴근 시간 ${workEnd}이 지났습니다. 퇴근 기록을 해주세요!`, notOutIds);
    }

    // users 변수 재사용을 위해 위에서 선언한 경우 아래서 재사용
    var _users = users;
  } else {
    var _users = null;
  }

  // ── 리마인더 체크 (공휴일 여부 무관하게 항상 실행) ───────────────
  const reminders = await fetchCollection('reminders');
  const reminderUsers = _users || await fetchCollection('users');
  const adminIds = reminderUsers.filter(u => u.role === 'admin').map(u => u.id);
  const allIds = reminderUsers.map(u => u.id);
  const currentHHMM = `${String(hh).padStart(2,'0')}:${String(mm).padStart(2,'0')}`;

  for (const r of reminders) {
    if (!r.active) continue;
    if (!shouldSendToday(r)) continue;
    const targetIds = r.target === 'all' ? allIds : adminIds;
    if (targetIds.length === 0) continue;
    console.log(`리마인더 발송: ${r.title} → ${r.target}`);
    await sendPush(`🔔 ${r.title}`, r.title, targetIds);
  }
}

main().catch(err => { console.error('오류:', err); process.exit(1); });
