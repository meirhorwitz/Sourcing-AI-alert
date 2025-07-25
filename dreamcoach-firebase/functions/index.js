const functions = require('firebase-functions');
const admin = require('firebase-admin');
const fetch = require('node-fetch');
const nodemailer = require('nodemailer');

admin.initializeApp();
const db = admin.firestore();

const OPENAI_API_KEY = functions.config().openai.key;
const SMTP_USER = functions.config().smtp.user;
const SMTP_PASS = functions.config().smtp.pass;

const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: { user: SMTP_USER, pass: SMTP_PASS }
});

exports.submitDream = functions.https.onRequest(async (req, res) => {
  if (req.method !== 'POST') {
    return res.status(405).send('Method Not Allowed');
  }

  const { name, email, dreamDate, description } = req.body;
  if (!name || !email || !dreamDate || !description) {
    return res.status(400).json({ error: 'Missing fields' });
  }

  try {
    // Analyze dream with OpenAI
    const aiResp = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${OPENAI_API_KEY}`
      },
      body: JSON.stringify({
        model: 'gpt-4',
        messages: [
          { role: 'system', content: 'You are a professional Jungian dream analyst.' },
          { role: 'user', content: description },
          { role: 'system', content: 'Also provide alternative interpretations using Freudian and Gestalt frameworks. Then generate three reflection questions.' }
        ]
      })
    });

    const aiData = await aiResp.json();
    const analysis = aiData.choices?.[0]?.message?.content || '';

    const docRef = await db.collection('dreams').add({
      name,
      email,
      dreamDate,
      description,
      analysis,
      created: admin.firestore.FieldValue.serverTimestamp()
    });

    await transporter.sendMail({
      from: SMTP_USER,
      to: email,
      subject: 'Your Dream Coaching Plan',
      html: `<p>Hi ${name},</p><p>Your dream analysis is ready.</p><pre>${analysis}</pre>`
    });

    res.json({ success: true, id: docRef.id, analysis });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Processing failed' });
  }
});
