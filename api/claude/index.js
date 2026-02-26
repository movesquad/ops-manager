const https = require('https');

module.exports = async function (context, req) {
  context.log('Claude proxy called, method:', req.method);

  if (req.method !== 'POST') {
    context.res = { status: 405, body: 'Method not allowed' };
    return;
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    context.log.error('ANTHROPIC_API_KEY not set');
    context.res = { status: 500, body: 'ANTHROPIC_API_KEY environment variable not configured' };
    return;
  }

  // Validate request body
  const payload = req.body;
  if (!payload || !payload.messages) {
    context.res = { status: 400, body: 'Missing messages in request body' };
    return;
  }

  const bodyStr = JSON.stringify(payload);
  context.log('Sending to Anthropic, model:', payload.model);

  try {
    const result = await new Promise((resolve, reject) => {
      const options = {
        hostname: 'api.anthropic.com',
        path: '/v1/messages',
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01',
          'Content-Length': Buffer.byteLength(bodyStr)
        }
      };

      const req = https.request(options, (res) => {
        let data = '';
        res.on('data', chunk => { data += chunk; });
        res.on('end', () => {
          context.log('Anthropic responded with status:', res.statusCode);
          resolve({ status: res.statusCode, body: data });
        });
      });

      req.on('error', (e) => {
        context.log.error('HTTPS request error:', e.message);
        reject(e);
      });

      req.setTimeout(30000, () => {
        req.destroy();
        reject(new Error('Request timed out'));
      });

      req.write(bodyStr);
      req.end();
    });

    context.res = {
      status: result.status,
      headers: { 'Content-Type': 'application/json' },
      body: result.body
    };

  } catch (err) {
    context.log.error('Proxy error:', err.message);
    context.res = {
      status: 502,
      body: 'Proxy error: ' + err.message
    };
  }
};
