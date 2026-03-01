const router = require('express').Router();

// POST /send — send LINE message
router.post('/send', async (req, res, next) => {
  try {
    const { message } = req.body;
    const token = process.env.LINE_CHANNEL_ACCESS_TOKEN;
    const targetId = process.env.LINE_TARGET_ID;

    if (!token || !targetId) {
      return res.status(400).json({ success: false, error: 'LINE credentials not configured' });
    }

    const response = await fetch('https://api.line.me/v2/bot/message/push', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${token}`,
      },
      body: JSON.stringify({
        to: targetId,
        messages: [{ type: 'text', text: message }],
      }),
    });

    if (!response.ok) {
      const err = await response.text();
      return res.status(response.status).json({ success: false, error: err });
    }

    res.json({ success: true, data: { sent: true } });
  } catch (err) { next(err); }
});

// POST /test — test LINE connection
router.post('/test', async (req, res, next) => {
  try {
    const { token, targetId } = req.body;

    if (!token || token.length < 50) {
      return res.json({
        success: false,
        error: 'Token ต้องมีความยาวอย่างน้อย 50 ตัวอักษร',
        hint: 'กรุณาตรวจสอบ LINE Channel Access Token',
      });
    }

    if (!targetId || (!/^C/.test(targetId) && !/^U/.test(targetId))) {
      return res.json({
        success: false,
        error: 'Target ID ต้องเริ่มต้นด้วย C (Group) หรือ U (User)',
        hint: 'ใช้ LINE Official Account Manager เพื่อดู User ID หรือ Group ID',
      });
    }

    const response = await fetch('https://api.line.me/v2/bot/message/push', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${token}`,
      },
      body: JSON.stringify({
        to: targetId,
        messages: [{ type: 'text', text: '🔔 ทดสอบการเชื่อมต่อ LINE สำเร็จ!' }],
      }),
    });

    if (!response.ok) {
      const errText = await response.text();
      return res.json({
        success: false,
        error: `LINE API Error: ${response.status}`,
        hint: errText,
      });
    }

    res.json({ success: true, data: { sent: true, message: 'ส่งข้อความทดสอบสำเร็จ!' } });
  } catch (err) { next(err); }
});

module.exports = router;
