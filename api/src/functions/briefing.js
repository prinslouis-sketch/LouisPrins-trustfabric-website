/* ═══════════════════════════════════════════════════════════════════
   POST /api/briefing
   Azure Static Web Apps managed function (Node 20, Functions v4 model).

   Receives the Book a Briefing form, sends the lead into the TrustFabric
   mailbox via Microsoft Graph, and sends the prospect an acknowledgement.
   No third-party form service: lead data never leaves your tenant.

   Zero npm dependencies — Node 20 has fetch built in.

   App settings required (Static Web App → Configuration):
     TENANT_ID       Directory (tenant) ID
     CLIENT_ID       App registration (client) ID
     CLIENT_SECRET   Client secret value
     MAIL_FROM       Sending mailbox, e.g. hello@trustfabric.co.za
     MAIL_TO         Where leads land, e.g. hello@trustfabric.co.za
     ACK_ENABLED     "true" to send the prospect an acknowledgement
   ══════════════════════════════════════════════════════════════════ */

const { app } = require('@azure/functions');

const GRAPH = 'https://graph.microsoft.com/v1.0';
const MAX = { name: 120, email: 160, company: 160, country: 80, message: 4000 };

/* Crude in-memory throttle. A single SWA instance handles this site's
   volume comfortably; it resets on cold start, which is fine — it exists
   to blunt scripted floods, not to be an audit control. */
const hits = new Map();
function throttled(ip, limit = 5, windowMs = 10 * 60 * 1000) {
  const now = Date.now();
  const recent = (hits.get(ip) || []).filter((t) => now - t < windowMs);
  recent.push(now);
  hits.set(ip, recent);
  if (hits.size > 5000) hits.clear();
  return recent.length > limit;
}

const clean = (v, max) =>
  String(v == null ? '' : v).replace(/[\u0000-\u001f\u007f]/g, ' ').trim().slice(0, max);

const escapeHtml = (s) =>
  String(s).replace(/[&<>"']/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

async function graphToken() {
  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    throw new Error('Graph credentials are not configured');
  }
  const res = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: CLIENT_ID,
        client_secret: CLIENT_SECRET,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials'
      })
    }
  );
  if (!res.ok) throw new Error(`Token request failed: ${res.status} ${await res.text()}`);
  return (await res.json()).access_token;
}

async function sendMail(token, from, message) {
  const res = await fetch(`${GRAPH}/users/${encodeURIComponent(from)}/sendMail`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ message, saveToSentItems: true })
  });
  if (!res.ok) throw new Error(`sendMail failed: ${res.status} ${await res.text()}`);
}

app.http('briefing', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'briefing',
  handler: async (request, context) => {
    const json = (status, body) => ({ status, jsonBody: body });

    let payload;
    try {
      payload = await request.json();
    } catch {
      return json(400, { error: 'Invalid request body.' });
    }

    /* Bot filters — answer 200 so scripts get no signal to tune against. */
    if (clean(payload.website, 200)) {
      context.log('Honeypot triggered; discarded.');
      return json(200, { ok: true });
    }
    if (Number(payload.elapsed) < 2000) {
      context.log('Submitted in under 2s; discarded.');
      return json(200, { ok: true });
    }

    const ip =
      (request.headers.get('x-forwarded-for') || '').split(',')[0].trim() || 'unknown';
    if (throttled(ip)) return json(429, { error: 'Too many requests. Please email us instead.' });

    const data = {
      name: clean(payload.name, MAX.name),
      email: clean(payload.email, MAX.email),
      company: clean(payload.company, MAX.company),
      country: clean(payload.country, MAX.country),
      message: clean(payload.message, MAX.message),
      page: clean(payload.page, 200)
    };

    if (!data.name || !data.company || !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(data.email)) {
      return json(400, { error: 'Please complete the required fields.' });
    }

    const { MAIL_FROM, MAIL_TO, ACK_ENABLED } = process.env;
    if (!MAIL_FROM || !MAIL_TO) {
      context.error('MAIL_FROM / MAIL_TO are not configured.');
      return json(500, { error: 'Form is not configured.' });
    }

    const rows = [
      ['Name', data.name],
      ['Work email', data.email],
      ['Company', data.company],
      ['Country', data.country],
      ['Source page', data.page || '/'],
      ['Received', new Date().toISOString()]
    ]
      .map(
        ([k, v]) =>
          `<tr><td style="padding:6px 16px 6px 0;color:#6b7280;white-space:nowrap">${k}</td>` +
          `<td style="padding:6px 0"><strong>${escapeHtml(v)}</strong></td></tr>`
      )
      .join('');

    const leadHtml = `
      <div style="font-family:Inter,Segoe UI,Arial,sans-serif;font-size:15px;color:#111">
        <h2 style="margin:0 0 4px;font-size:18px">Executive briefing request</h2>
        <p style="margin:0 0 18px;color:#6b7280">${escapeHtml(data.company)} &middot; ${escapeHtml(data.country)}</p>
        <table style="border-collapse:collapse;margin-bottom:18px">${rows}</table>
        <div style="padding:14px 16px;background:#f5f3ff;border-left:3px solid #6A3DFF;border-radius:4px">
          <div style="color:#6b7280;font-size:13px;margin-bottom:6px">What they are looking to solve</div>
          <div style="white-space:pre-wrap">${escapeHtml(data.message || 'Not specified')}</div>
        </div>
      </div>`;

    try {
      const token = await graphToken();

      await sendMail(token, MAIL_FROM, {
        subject: `Briefing request — ${data.company} (${data.name})`,
        body: { contentType: 'HTML', content: leadHtml },
        toRecipients: [{ emailAddress: { address: MAIL_TO } }],
        replyTo: [{ emailAddress: { address: data.email, name: data.name } }]
      });

      if (String(ACK_ENABLED).toLowerCase() === 'true') {
        const first = data.name.split(' ')[0];
        const ackHtml = `
          <div style="font-family:Inter,Segoe UI,Arial,sans-serif;font-size:15px;color:#111;line-height:1.6">
            <p>Hi ${escapeHtml(first)},</p>
            <p>Thank you for your interest in TrustFabric. We've received your request for an
               executive briefing and will come back to you within one business day to arrange a time.</p>
            <p>If it's easier, you're welcome to reply to this email with any additional context
               on your environment or timelines.</p>
            <p style="margin-top:22px">Regards,<br><strong>TrustFabric</strong><br>
               <span style="color:#6b7280">Digital Trust, Cybersecurity &amp; Cryptographic Infrastructure</span></p>
          </div>`;
        try {
          await sendMail(token, MAIL_FROM, {
            subject: 'We received your briefing request — TrustFabric',
            body: { contentType: 'HTML', content: ackHtml },
            toRecipients: [{ emailAddress: { address: data.email, name: data.name } }]
          });
        } catch (ackErr) {
          /* The lead is already safely delivered — never fail the request
             because the courtesy acknowledgement bounced. */
          context.warn('Acknowledgement failed:', ackErr.message);
        }
      }

      return json(200, { ok: true });
    } catch (err) {
      context.error('Briefing submission failed:', err.message);
      return json(502, { error: 'We could not send your request. Please email hello@trustfabric.co.za.' });
    }
  }
});