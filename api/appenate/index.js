const https = require('https');

module.exports = async function (context, req) {
  context.log('Appenate proxy called');

  if (req.method !== 'POST') {
    context.res = { status: 405, body: 'Method not allowed' };
    return;
  }

  const integrationKey = process.env.APPENATE_INTEGRATION_KEY;
  const providerId     = process.env.APPENATE_PROVIDER_ID;

  if (!integrationKey || !providerId) {
    context.log.error('Appenate credentials not configured');
    context.res = { status: 500, body: 'APPENATE_INTEGRATION_KEY or APPENATE_PROVIDER_ID not configured in Azure' };
    return;
  }

  const { action, payload } = req.body || {};
  if (!action || !payload) {
    context.res = { status: 400, body: 'Missing action or payload in request body' };
    return;
  }

  // Inject credentials into payload
  payload.IntegrationKey = integrationKey;
  payload.ProviderId     = parseInt(providerId, 10);

  let method = 'POST';
  let path   = '/api/v1/stask?format=json';

  if (action === 'getTask') {
    method = 'GET';
    path   = '/api/v1/stask?format=json&Id=' + encodeURIComponent(payload.Id)
           + '&ProviderId=' + providerId
           + '&Integrationkey=' + encodeURIComponent(integrationKey);
  } else if (action === 'deleteTask') {
    method = 'DELETE';
    path   = '/api/v1/stask?format=json&Id=' + encodeURIComponent(payload.Id)
           + '&ProviderId=' + providerId
           + '&Integrationkey=' + encodeURIComponent(integrationKey);
  }

  const bodyStr = (method === 'POST') ? JSON.stringify(payload) : null;

  try {
    const result = await new Promise((resolve, reject) => {
      const options = {
        hostname: 'secure.appenate.com',
        path:     path,
        method:   method,
        headers:  {
          'Content-Type': 'application/json',
          'Accept':        'application/json'
        }
      };

      if (bodyStr) {
        options.headers['Content-Length'] = Buffer.byteLength(bodyStr);
      }

      const apReq = https.request(options, (res) => {
        let data = '';
        res.on('data', chunk => { data += chunk; });
        res.on('end', () => {
          context.log('Appenate status:', res.statusCode);
          resolve({ status: res.statusCode, body: data });
        });
      });

      apReq.on('error', (e) => {
        context.log.error('Appenate request error:', e.message);
        reject(e);
      });

      apReq.setTimeout(30000, () => {
        apReq.destroy();
        reject(new Error('Timed out'));
      });

      if (bodyStr) apReq.write(bodyStr);
      apReq.end();
    });

    context.res = {
      status:  result.status,
      headers: { 'Content-Type': 'application/json' },
      body:    result.body
    };

  } catch (err) {
    context.log.error('Appenate proxy error:', err.message);
    context.res = { status: 502, body: 'Proxy error: ' + err.message };
  }
};
