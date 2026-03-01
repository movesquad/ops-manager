const https = require('https');

let tokenCache = { token: null, expiresAt: 0 };

async function getAccessToken(tenantId, clientId, clientSecret) {
  const now = Date.now();
  if (tokenCache.token && now < tokenCache.expiresAt - 60000) {
    return { token: tokenCache.token, error: null };
  }
  const body = [
    'grant_type=client_credentials',
    'scope=' + encodeURIComponent('https://graph.microsoft.com/.default'),
    'client_id=' + encodeURIComponent(clientId),
    'client_secret=' + encodeURIComponent(clientSecret)
  ].join('&');
  const result = await new Promise((resolve, reject) => {
    const options = {
      hostname: 'login.microsoftonline.com',
      path: '/' + tenantId + '/oauth2/v2.0/token',
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) }
    };
    const req = https.request(options, (res) => {
      let data = ''; res.on('data', c => { data += c; }); res.on('end', () => resolve({ status: res.statusCode, body: data }));
    });
    req.on('error', reject);
    req.setTimeout(15000, () => { req.destroy(); reject(new Error('Token timeout')); });
    req.write(body); req.end();
  });
  let parsed; try { parsed = JSON.parse(result.body); } catch(e) { parsed = {}; }
  if (result.status !== 200 || !parsed.access_token) {
    return { token: null, error: parsed.error_description || parsed.error || 'Token failed: HTTP ' + result.status };
  }
  tokenCache.token = parsed.access_token;
  tokenCache.expiresAt = now + (parsed.expires_in * 1000);
  return { token: tokenCache.token, error: null };
}

async function graphRequest(token, method, path, body) {
  const bodyStr = body ? JSON.stringify(body) : null;
  return new Promise((resolve, reject) => {
    const options = {
      hostname: 'graph.microsoft.com', path, method,
      headers: {
        'Authorization': 'Bearer ' + token, 'Accept': 'application/json',
        ...(bodyStr ? { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(bodyStr) } : {})
      }
    };
    const req = https.request(options, (res) => {
      let data = ''; res.on('data', c => { data += c; }); res.on('end', () => resolve({ status: res.statusCode, body: data }));
    });
    req.on('error', reject);
    req.setTimeout(30000, () => { req.destroy(); reject(new Error('Graph timeout')); });
    if (bodyStr) req.write(bodyStr);
    req.end();
  });
}

module.exports = async function (context, req) {
  context.log('Email proxy called');
  if (req.method !== 'POST') { context.res = { status: 405, body: 'Method not allowed' }; return; }

  const tenantId     = process.env.SP_TENANT_ID;
  const clientId     = process.env.SP_CLIENT_ID;
  const clientSecret = process.env.SP_CLIENT_SECRET;
  const mailFrom     = process.env.MAIL_FROM || 'updates@onwards.network';

  const missing = ['SP_TENANT_ID','SP_CLIENT_ID','SP_CLIENT_SECRET'].filter(k => !process.env[k]);
  if (missing.length) {
    context.res = { status: 500, headers: {'Content-Type':'application/json'}, body: JSON.stringify({ error: 'Missing env vars: ' + missing.join(', ') }) };
    return;
  }

  const { to, cc, subject, html, text, replyTo } = req.body || {};
  if (!to || !subject || (!html && !text)) {
    context.res = { status: 400, headers: {'Content-Type':'application/json'}, body: JSON.stringify({ error: 'Missing required fields: to, subject, html or text' }) };
    return;
  }

  try {
    const { token, error: tokenError } = await getAccessToken(tenantId, clientId, clientSecret);
    if (!token) {
      context.res = { status: 401, headers: {'Content-Type':'application/json'}, body: JSON.stringify({ error: tokenError }) };
      return;
    }

    // Build recipient arrays
    const toRecipients = (Array.isArray(to) ? to : [to])
      .filter(Boolean)
      .map(addr => ({ emailAddress: { address: addr } }));

    const ccRecipients = cc ? (Array.isArray(cc) ? cc : [cc])
      .filter(Boolean)
      .map(addr => ({ emailAddress: { address: addr } })) : [];

    const message = {
      subject,
      body: { contentType: html ? 'HTML' : 'Text', content: html || text },
      toRecipients,
      ...(ccRecipients.length ? { ccRecipients } : {}),
      ...(replyTo ? { replyTo: [{ emailAddress: { address: replyTo } }] } : {}),
      from: { emailAddress: { address: mailFrom } }
    };

    // Send via Graph — /users/{mailFrom}/sendMail
    const result = await graphRequest(token, 'POST',
      '/v1.0/users/' + encodeURIComponent(mailFrom) + '/sendMail',
      { message, saveToSentItems: true }
    );

    if (result.status === 202) {
      context.res = { status: 200, headers: {'Content-Type':'application/json'}, body: JSON.stringify({ ok: true }) };
    } else {
      let parsed; try { parsed = JSON.parse(result.body); } catch(e) { parsed = {}; }
      const errMsg = (parsed.error && (parsed.error.message || parsed.error.code)) || result.body.slice(0,200);
      context.log.error('Send mail error:', result.status, errMsg);
      context.res = { status: result.status, headers: {'Content-Type':'application/json'}, body: JSON.stringify({ error: errMsg }) };
    }

  } catch (err) {
    context.log.error('Email proxy exception:', err.message);
    context.res = { status: 502, headers: {'Content-Type':'application/json'}, body: JSON.stringify({ error: err.message }) };
  }
};
