const router = require('express').Router();

// POST /report — send HTML email via Resend API
router.post('/report', async (req, res, next) => {
  try {
    const { to, subject, html } = req.body;
    const apiKey = process.env.RESEND_API_KEY;

    if (!apiKey) {
      return res.status(400).json({ success: false, error: 'RESEND_API_KEY not configured' });
    }

    const response = await fetch('https://api.resend.com/emails', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        from: 'pharmacy@resend.dev',
        to: Array.isArray(to) ? to : [to],
        subject,
        html,
      }),
    });

    if (!response.ok) {
      const errText = await response.text();
      return res.status(response.status).json({ success: false, error: errText });
    }

    const data = await response.json();
    res.json({ success: true, data });
  } catch (err) { next(err); }
});

module.exports = router;
