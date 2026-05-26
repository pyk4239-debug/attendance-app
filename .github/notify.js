import https from 'https';

const FIREBASE_API_KEY = (process.env.FIREBASE_API_KEY || '').trim();
const FIREBASE_PROJECT_ID = (process.env.FIREBASE_PROJECT_ID || '').trim();
const ONESIGNAL_APP_ID = (process.env.ONESIGNAL_APP_ID || '').trim();
const ONESIGNAL_API_KEY = (process.env.ONESIGNAL_API_KEY || '').trim();

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

  // userIds가 null이면 전체 발송, 아니면 태그 필터
  const payload = {
    app_id: ONESIGNAL_APP_ID,
    headings: { en: title, ko: title },
    contents: { en: message, ko: message },
  };
  if (userIds === null) {
    payload.included_segments = ['All'];
  } else {
    payload.filters = userIds.reduce((acc, id, i) => {
      if (i > 0) acc.push({ operator: 'OR' });
      acc.push({ field: 'tag', key: 'userId', relation: '=', value: String(id) });
      return acc;
    }, []);
  }
  const body = JSON.stringify(payload);
  console.log('sendPush body:', body);
  console.log('API_KEY 앞8자:', ONESIGNAL_API_KEY.slice(0,8), '길이:', ONESIGNAL_API_KEY.length);

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

  const currentHHMM = `${String(hh).padStart(2,'0')}:${String(mm).padStart(2,'0')}`;
  const todayDom = kst.getUTCDate();

  const matchesRepeat = (r, dom2, dow2) => {
    if (r.repeat === 'weekly' && Number(r.weekDay) !== dow2) return false;
    if (r.repeat === 'monthly' && Number(r.monthDay) !== dom2) return false;
    return true;
  };

  const shouldSendToday = (r) => {
    // 시간 체크 (±4분 허용 - GitHub Actions 지연 대응)
    const [rh, rm] = r.time.split(':').map(Number);
    const rMin = rh * 60 + rm;
    const nowMin = hh * 60 + mm;
    if (Math.abs(rMin - nowMin) > 4) return false;

    // 오늘 공휴일 아니면 → 오늘 날짜로 반복 조건 체크
    if (!isHolidayDate(today)) {
      if (!matchesRepeat(r, todayDom, dayOfWeek)) return false;
      return true;
    }

    // 오늘이 공휴일인데 전날발송 옵션 없으면 → 발송 안 함
    if (!r.sendBeforeHoliday) return false;

    // 오늘이 공휴일 + 전날발송: 오늘이 공휴일 바로 다음 평일이 있는 연속 공휴일의 마지막 전날인지 확인
    // 어제가 평일이어야 함 (오늘이 첫 번째 공휴일의 전날 = 어제가 평일)
    // 사실상 공휴일 당일에는 발송 안 함 - 전날(평일)에 발송하는 구조
    return false;
  };

  const todayIsHoliday = isHolidayDate(today);

  // ── 출퇴근 알림 (평일+비공휴일만) ────────────────────────────
  const allUsers = await fetchCollection('users');

  if (!todayIsHoliday) {
    const [startH, startM] = workStart.split(':').map(Number);
    const [endH, endM] = workEnd.split(':').map(Number);
    const currentMinutes = hh * 60 + mm;
    const startMinutes = startH * 60 + startM;
    const endMinutes = endH * 60 + endM;
    const isCheckinAlert = currentMinutes >= startMinutes - 7 && currentMinutes <= startMinutes - 3;
    const isCheckoutAlert = currentMinutes >= endMinutes - 1 && currentMinutes <= endMinutes + 3;

    console.log(`현재: ${currentMinutes}분, 출근알림: ${isCheckinAlert}, 퇴근알림: ${isCheckoutAlert}`);

    const members = allUsers.filter(u => u.role === 'member');
    console.log('팀원:', members.map(u => u.name));

    if (isCheckinAlert || isCheckoutAlert) {
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
    }
  } else {
    console.log('공휴일 - 출퇴근 알림 생략');
  }

  // ── 리마인더 체크 (공휴일 여부 무관하게 항상 실행) ───────────────
  console.log('리마인더 체크 시작');
  const reminders = await fetchCollection('reminders');
  console.log('리마인더 수:', reminders.length);
  const adminIds = allUsers.filter(u => u.role === 'admin').map(u => u.id);
  const allIds = allUsers.map(u => u.id);
  console.log('adminIds:', adminIds, 'allIds:', allIds);
  console.log('현재시간(HHMM):', currentHHMM);

  for (const r of reminders) {
    if (!r.active) continue;
    const [rh2, rm2] = r.time.split(':').map(Number); const timeMatch = Math.abs((rh2*60+rm2) - (hh*60+mm)) <= 4;
    const domMatch = r.repeat !== 'monthly' || Number(r.monthDay) === todayDom;
    const dowMatch = r.repeat !== 'weekly' || Number(r.weekDay) === dayOfWeek;
    console.log(`리마인더: ${r.title} | 시간:${r.time}==${currentHHMM}(${timeMatch}) | 날짜:${r.monthDay}==${todayDom}(${domMatch}) | 요일:${r.weekDay}==${dayOfWeek}(${dowMatch})`);
    if (!shouldSendToday(r)) { console.log(`  → 스킵`); continue; }
    const targetIds = r.target === 'all' ? allIds : adminIds;
    if (targetIds.length === 0) continue;
    // 중복 발송 방지 - Firestore에 오늘 발송 기록 저장
    const firedDocId = `reminder_${r.id}_${today}`;
    const firedUrl = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/reminder_fired/${firedDocId}?key=${FIREBASE_API_KEY}`;
    const firedCheck = await httpsGet(firedUrl);
    if (firedCheck.fields) { console.log(`  → 오늘 이미 발송됨 스킵`); continue; }
    // 발송 기록 저장
    await new Promise((resolve, reject) => {
      const body = JSON.stringify({ fields: { firedAt: { stringValue: new Date().toISOString() } } });
      const url = new URL(firedUrl);
      const req = https.request({ hostname: url.hostname, path: url.pathname + url.search, method: 'PATCH',
        headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) }
      }, res => { res.on('data', () => {}); res.on('end', resolve); });
      req.on('error', reject); req.write(body); req.end();
    });
    console.log(`  → 발송! ${r.title}`);
    await sendPush(`🔔 ${r.title}`, r.title, targetIds);
  }
}

main().catch(err => { console.error('오류:', err); process.exit(1); });
