export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { title, message, targetUserId } = req.body;

  const body = {
    app_id: process.env.ONESIGNAL_APP_ID,
    headings: { en: title, ko: title },
    contents: { en: message, ko: message },
    collapse_id: Date.now().toString(), // 알림 누적 표시 (덮어쓰기 방지)
  };

  if (targetUserId) {
    body.filters = [{ field: 'tag', key: 'userId', relation: '=', value: targetUserId }];
  } else {
    body.included_segments = ['All'];
  }

  const response = await fetch('https://onesignal.com/api/v1/notifications', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Key ${process.env.ONESIGNAL_API_KEY}`,
    },
    body: JSON.stringify(body),
  });

  const data = await response.json();
  console.log('OneSignal 응답:', JSON.stringify(data));
  return res.status(200).json(data);
}
