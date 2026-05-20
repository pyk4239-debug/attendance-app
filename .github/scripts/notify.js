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

  if (holidays.includes(today)) { console.log('공휴일'); return; }

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
  const leaveSet = new Set(
    leaves.filter(l => l.date === today && !l.deleted).map(l => l.userId)
  );

  if (isCheckinAlert) {
    const absentIds = members
      .filter(u => !leaveSet.has(u.id) && !recordMap[`${u.id}_${today}`]?.in)
      .map(u => u.id);
    console.log('출근 알림 대상:', absentIds);
    if (absentIds.length > 0) {
      await sendPush('⏰ 출근 시간 알림', `출근 시간 ${workStart} 5분 전입니다. 출근 기록을 해주세요!`, absentIds);
    }
  }

  if (isCheckoutAlert) {
    const notOutIds = members
      .filter(u => !leaveSet.has(u.id) && recordMap[`${u.id}_${today}`]?.in && !recordMap[`${u.id}_${today}`]?.out)
      .map(u => u.id);
    console.log('퇴근 알림 대상:', notOutIds);
    if (notOutIds.length > 0) {
      await sendPush('🏠 퇴근 시간 알림', `퇴근 시간 ${workEnd}이 지났습니다. 퇴근 기록을 해주세요!`, notOutIds);
    }
  }
}

main().catch(err => { console.error('오류:', err); process.exit(1); });
