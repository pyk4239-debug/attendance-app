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
      res.on('end', () => {
        try { resolve(JSON.parse(data)); } catch(e) { resolve({}); }
      });
    }).on('error', reject);
  });
}

function httpsPost(options, body) {
  return new Promise((resolve, reject) => {
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try { resolve(JSON.parse(data)); } catch(e) { resolve({ error: data.slice(0, 200) }); }
      });
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
  // userIds: string[] → 특정 유저, null → 전체
  const payload = {
    app_id: ONESIGNAL_APP_ID,
    headings: { en: title, ko: title },
    contents: { en: message, ko: message },
  };
  if (userIds === null) {
    payload.included_segments = ['All'];
  } else {
    if (!userIds || userIds.length === 0) return;
    payload.filters = userIds.reduce((acc, id, i) => {
      if (i > 0) acc.push({ operator: 'OR' });
      acc.push({ field: 'tag', key: 'userId', relation: '=', value: String(id) });
      return acc;
    }, []);
  }
  const body = JSON.stringify(payload);
  const options = {
    hostname: 'onesignal.com',
    path: '/api/v1/notifications',
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Key ${ONESIGNAL_API_KEY}`,
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
  const todayDom = kst.getUTCDate();
  const nowMin = hh * 60 + mm;

  console.log(`KST: ${hh}:${String(mm).padStart(2,'0')} (${today}) 요일:${dayOfWeek}`);

  if (dayOfWeek === 0 || dayOfWeek === 6) { console.log('주말'); return; }

  const settings = await fetchDoc('app', 'settings');
  if (!settings) { console.log('settings 없음'); return; }

  const workStart = settings.workStart || '09:00';
  const workEnd = settings.workEnd || '18:00';
  const holidays = settings.holidays || [];

  console.log(`출근: ${workStart}, 퇴근: ${workEnd}`);

  // 공휴일 여부 확인
  const isHolidayDate = (dateStr) => {
    const d = new Date(dateStr + 'T00:00:00Z');
    if (d.getUTCDay() === 0 || d.getUTCDay() === 6) return true;
    return holidays.some(h => (typeof h === 'string' ? h : h.date) === dateStr);
  };

  const todayIsHoliday = isHolidayDate(today);
  const allUsers = await fetchCollection('users');

  // ── 출퇴근 알림 (평일/비공휴일만) ─────────────────────────────
  if (!todayIsHoliday) {
    const [startH, startM] = workStart.split(':').map(Number);
    const [endH, endM] = workEnd.split(':').map(Number);
    const startMin = startH * 60 + startM;
    const endMin = endH * 60 + endM;
    const isCheckinAlert  = nowMin >= startMin - 5 && nowMin <= startMin - 3;
    const isCheckoutAlert = nowMin >= endMin       && nowMin <= endMin + 3;

    console.log(`현재: ${nowMin}분, 출근알림: ${isCheckinAlert}, 퇴근알림: ${isCheckoutAlert}`);

    const members = allUsers.filter(u => u.role === 'member');
    console.log('팀원:', members.map(u => u.name));

    if (isCheckinAlert || isCheckoutAlert) {
      const records = await fetchCollection('records');
      const recordMap = {};
      for (const rec of records) recordMap[rec.id] = rec;

      const leaves = await fetchCollection('leaves');
      const todayLeaves = leaves.filter(l => l.date === today && !l.deleted);
      const leaveSet   = new Set(todayLeaves.map(l => l.userId));
      const amLeaveSet = new Set(todayLeaves.filter(l => l.type?.includes('오전')).map(l => l.userId));
      const pmLeaveSet = new Set(todayLeaves.filter(l => l.type?.includes('오후')).map(l => l.userId));

      if (isCheckinAlert) {
        const ids = members.filter(u => {
          const rec = recordMap[`${u.id}_${today}`];
          if (leaveSet.has(u.id) || amLeaveSet.has(u.id)) return false;
          if (rec?.in) return false;
          return true;
        }).map(u => u.id);
        console.log('출근 알림 대상:', ids);
        if (ids.length > 0) await sendPush('⏰ 출근 시간 알림', `출근 시간 ${workStart} 5분 전입니다. 출근 기록을 해주세요!`, ids);
      }

      if (isCheckoutAlert) {
        const ids = members.filter(u => {
          const rec = recordMap[`${u.id}_${today}`];
          if (leaveSet.has(u.id) || pmLeaveSet.has(u.id)) return false;
          if (!rec?.in || rec?.out) return false;
          const outings = rec?.outing || [];
          if (outings.length > 0 && !outings[outings.length - 1].in) return false;
          return true;
        }).map(u => u.id);
        console.log('퇴근 알림 대상:', ids);
        if (ids.length > 0) await sendPush('🏠 퇴근 시간 알림', `퇴근 시간 ${workEnd}이 지났습니다. 퇴근 기록을 해주세요!`, ids);
      }
    }
  } else {
    console.log('공휴일 - 출퇴근 알림 생략');
  }

  // ── 리마인더 체크 (공휴일 여부 무관하게 항상 실행) ───────────────
  const reminders = await fetchCollection('reminders');
  console.log(`리마인더 ${reminders.length}개, 현재 ${String(hh).padStart(2,'0')}:${String(mm).padStart(2,'0')}`);

  const adminIds = allUsers.filter(u => u.role === 'admin').map(u => u.id);
  const allIds   = allUsers.map(u => u.id);

  for (const r of reminders) {
    if (!r.active) continue;

    // 시간 체크 (±4분 허용)
    const [rh, rm] = r.time.split(':').map(Number);
    const rMin = rh * 60 + rm;
    if (Math.abs(rMin - nowMin) > 4) continue;

    // 오늘이 공휴일이면 리마인더 발송 안 함 (전날발송도 공휴일 당일엔 안 보냄)
    if (todayIsHoliday) continue;

    // 전날발송 옵션: 내일~최대7일 후 중 첫 공휴일이 있으면 오늘 그 날짜의 리마인더 발송
    if (r.sendBeforeHoliday) {
      let sent = false;
      for (let i = 1; i <= 7; i++) {
        const checkDate = new Date(kst.getTime() + i * 24 * 60 * 60 * 1000);
        const checkStr  = checkDate.toISOString().slice(0, 10);
        if (!isHolidayDate(checkStr)) break; // 첫 번째 평일 만나면 종료
        // checkDate가 공휴일 → 이 날의 반복 조건 확인
        const checkDom = checkDate.getUTCDate();
        const checkDow = checkDate.getUTCDay();
        const matches = (r.repeat === 'daily') ||
                        (r.repeat === 'weekly'  && Number(r.weekDay)  === checkDow) ||
                        (r.repeat === 'monthly' && Number(r.monthDay) === checkDom);
        if (matches) { sent = true; break; }
      }
      if (!sent) {
        // 전날발송 조건 미해당 → 원래 날짜로 체크
        if (r.repeat === 'weekly'  && Number(r.weekDay)  !== dayOfWeek) continue;
        if (r.repeat === 'monthly' && Number(r.monthDay) !== todayDom)  continue;
      }
    } else {
      // 반복 조건 체크 (일반)
      if (r.repeat === 'weekly'  && Number(r.weekDay)  !== dayOfWeek) continue;
      if (r.repeat === 'monthly' && Number(r.monthDay) !== todayDom)  continue;
    }

    // 중복 발송 방지
    const firedDocId = `reminder_${r.id}_${today}`;
    const firedUrl = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/reminder_fired/${firedDocId}?key=${FIREBASE_API_KEY}`;
    const firedCheck = await httpsGet(firedUrl);
    if (firedCheck.fields) { console.log(`  이미 발송됨: ${r.title}`); continue; }

    // 발송 기록 저장
    await new Promise((resolve, reject) => {
      const body = JSON.stringify({ fields: { firedAt: { stringValue: new Date().toISOString() } } });
      const url = new URL(firedUrl);
      const req = https.request({
        hostname: url.hostname, path: url.pathname + url.search, method: 'PATCH',
        headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) }
      }, res => { res.on('data', () => {}); res.on('end', resolve); });
      req.on('error', reject); req.write(body); req.end();
    });

    const targetIds = r.target === 'all' ? allIds : adminIds;
    console.log(`리마인더 발송: ${r.title} → ${r.target}`);
    await sendPush(`🔔 ${r.title}`, r.title, targetIds);
  }
}

main().catch(err => { console.error('오류:', err); process.exit(1); });
