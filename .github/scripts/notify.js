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

  const settings = await fetchDoc('settings', 'main');
  if (!settings) { console.log('settings 없음'); return; }

  const workStart = settings.workStart || '09:00';
  const workEnd = settings.workEnd || '18:00';
  const holidays = settings.holidays || [];

  console.log(`출근: ${workStart}, 퇴근: ${workEnd}`);

  const isHoliday = holidays.some(h => (typeof h === 'string' ? h : h.date) === today);
  if (isHoliday) { console.log('공휴일'); return; }

  const [startH, startM] = workStart.split(':').map(Number);
  const [endH, endM] = workEnd.split(':').map(Number);

  const currentMinutes = hh * 60 + mm;
  const startMinutes = startH * 60 + startM;
  const endMinutes = endH * 60 + endM;

  const isCheckinAlert = currentMinutes === startMinutes - 5;
  const isCheckoutAlert = currentMinutes === endMinutes + 1;

  console.log(`현재: ${currentMinutes}분, 출근알림: ${isCheckinAlert}, 퇴근알림: ${isCheckoutAlert}`);

  if (!isCheckinAlert && !isCheckoutAlert) { console.log('알림 시간 아님'); return; }

  const users = await fetchCollection('users');
  const members = users.filter(u => u.role === 'member');
  console.log('팀원:', members.map(u => u.name));

  const records = await fetchCollection('records');
  const recordMap = {};
  for (const rec of records) recordMap[rec.id] = rec;

  const leaves = await fetchCollection('leaves');

  // 오늘 연차/반차 목록
  const todayLeaves = leaves.filter(l => l.date === today && !l.deleted);
  const leaveSet = new Set(todayLeaves.map(l => l.userId)); // 연차(전일)
  const amLeaveSet = new Set(todayLeaves.filter(l => l.type && l.type.includes('오전')).map(l => l.userId)); // 반차(오전)
  const pmLeaveSet = new Set(todayLeaves.filter(l => l.type && l.type.includes('오후')).map(l => l.userId)); // 반차(오후)

  if (isCheckinAlert) {
    const absentIds = members
      .filter(u => {
        const rec = recordMap[`${u.id}_${today}`];
        if (leaveSet.has(u.id)) return false;   // 1. 연차 제외
        if (amLeaveSet.has(u.id)) return false;  // 2. 반차(오전) 제외
        if (rec?.in) return false;               // 3. 이미 출근 제외
        return true;
      })
      .map(u => u.id);
    console.log('출근 알림 대상:', absentIds);
    if (absentIds.length > 0) {
      await sendPush('⏰ 출근 시간 알림', `출근 시간 ${workStart} 5분 전입니다. 출근 기록을 해주세요!`, absentIds);
    }
  }

  if (isCheckoutAlert) {
    const notOutIds = members
      .filter(u => {
        const rec = recordMap[`${u.id}_${today}`];
        if (leaveSet.has(u.id)) return false;    // 1. 연차 제외
        if (pmLeaveSet.has(u.id)) return false;  // 2. 반차(오후) 제외
        if (!rec?.in) return false;              // 3. 출근 안 한 팀원 제외
        if (rec?.out) return false;              // 4. 이미 퇴근 제외
        // 5. 외출 중인 팀원 제외 (outing 배열 마지막이 out만 있고 in 없는 경우)
        const outings = rec?.outing || [];
        if (outings.length > 0 && !outings[outings.length - 1].in) return false;
        return true;
      })
      .map(u => u.id);
    console.log('퇴근 알림 대상:', notOutIds);
    if (notOutIds.length > 0) {
      await sendPush('🏠 퇴근 시간 알림', `퇴근 시간 ${workEnd}이 지났습니다. 퇴근 기록을 해주세요!`, notOutIds);
    }
  }

  // ── 리마인더 체크 ──────────────────────────────────────────────
  const reminders = await fetchCollection('reminders');
  const currentHHMM = `${String(hh).padStart(2,'0')}:${String(mm).padStart(2,'0')}`;
  const adminIds = users.filter(u => u.role === 'admin').map(u => u.id);
  const allIds = users.map(u => u.id);

  for (const r of reminders) {
    if (!r.active) continue;
    if (r.time !== currentHHMM) continue;

    // 반복 조건
    if (r.repeat === 'weekly' && r.weekDay !== dayOfWeek) continue;
    if (r.repeat === 'monthly' && r.monthDay !== kst.getUTCDate()) continue;

    // 대상
    const targetIds = r.target === 'all' ? allIds : adminIds;
    if (targetIds.length === 0) continue;

    console.log(`리마인더 발송: ${r.title} → ${r.target}`);
    await sendPush(`📅 ${r.title}`, r.title, targetIds);
  }
}

main().catch(err => { console.error('오류:', err); process.exit(1); });
