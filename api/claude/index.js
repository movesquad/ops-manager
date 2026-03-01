const https = require('https');

module.exports = async function (context, req) {
  if (req.method !== 'POST') { context.res = { status: 405, body: 'Method not allowed' }; return; }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    context.res = { status: 500, headers: {'Content-Type':'application/json'}, body: JSON.stringify({ error: 'ANTHROPIC_API_KEY not configured' }) };
    return;
  }

  const { model, max_tokens, messages, system } = req.body || {};
  if (!messages) {
    context.res = { status: 400, headers: {'Content-Type':'application/json'}, body: JSON.stringify({ error: 'Missing messages' }) };
    return;
  }

  const payload = JSON.stringify({
    model:      model      || 'claude-sonnet-4-20250514',
    max_tokens: max_tokens || 2000,
    messages,
    ...(system ? { system } : {})
  });

  try {
    const result = await new Promise((resolve, reject) => {
      const opts = {
        hostname: 'api.anthropic.com',
        path:     '/v1/messages',
        method:   'POST',
        headers: {
          'Content-Type':      'application/json',
          'x-api-key':         apiKey,
          'anthropic-version': '2023-06-01',
          'Content-Length':    Buffer.byteLength(payload)
        }
      };
      const req2 = https.request(opts, (res) => {
        let data = ''; res.on('data', c => { data += c; }); res.on('end', () => resolve({ status: res.statusCode, body: data }));
      });
      req2.on('error', reject);
      req2.setTimeout(60000, () => { req2.destroy(); reject(new Error('Claude API timeout')); });
      req2.write(payload); req2.end();
    });

    context.res = { status: result.status, headers: {'Content-Type':'application/json'}, body: result.body };
  } catch (err) {
    context.res = { status: 502, headers: {'Content-Type':'application/json'}, body: JSON.stringify({ error: err.message }) };
  }
};
