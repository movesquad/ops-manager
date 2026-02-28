const https = require('https');

// ── Token cache (in-memory, per function instance) ────────────────────────
let tokenCache = { token: null, expiresAt: 0 };

async function getAccessToken(tenantId, clientId, clientSecret) {
  const now = Date.now();
  if (tokenCache.token && now < tokenCache.expiresAt - 60000) {
    return tokenCache.token;
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
      path:     '/' + tenantId + '/oauth2/v2.0/token',
      method:   'POST',
      headers:  {
        'Content-Type':   'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(body)
      }
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => resolve({ status: res.statusCode, body: data }));
    });
    req.on('error', reject);
    req.setTimeout(15000, () => { req.destroy(); reject(new Error('Token request timed out')); });
    req.write(body);
    req.end();
  });

  if (result.status !== 200) {
    // Parse error for better message
    let errMsg = 'Token fetch failed (' + result.status + ')';
    try {
      const parsed = JSON.parse(result.body);
      errMsg = parsed.error_description || parsed.error || errMsg;
    } catch(e) {}
    throw new Error(errMsg);
  }

  const parsed = JSON.parse(result.body);
  tokenCache.token     = parsed.access_token;
  tokenCache.expiresAt = now + (parsed.expires_in * 1000);
  return tokenCache.token;
}

async function graphRequest(token, method, path, body) {
  const bodyStr = body ? JSON.stringify(body) : null;

  return new Promise((resolve, reject) => {
    const options = {
      hostname: 'graph.microsoft.com',
      path:     path,
      method:   method,
      headers:  {
        'Authorization': 'Bearer ' + token,
        'Accept':        'application/json'
      }
    };

    if (bodyStr) {
      options.headers['Content-Type']   = 'application/json';
      options.headers['Content-Length'] = Buffer.byteLength(bodyStr);
    }

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => resolve({ status: res.statusCode, body: data }));
    });

    req.on('error', reject);
    req.setTimeout(30000, () => { req.destroy(); reject(new Error('Graph request timed out')); });
    if (bodyStr) req.write(bodyStr);
    req.end();
  });
}

module.exports = async function (context, req) {
  context.log('SharePoint proxy called:', req.body && req.body.action);

  if (req.method !== 'POST') {
    context.res = { status: 405, body: 'Method not allowed' };
    return;
  }

  // ── Credentials from environment variables ────────────────────────────
  const tenantId     = process.env.SP_TENANT_ID;
  const clientId     = process.env.SP_CLIENT_ID;
  const clientSecret = process.env.SP_CLIENT_SECRET;
  const siteUrl      = process.env.SP_SITE_URL;

  if (!tenantId || !clientId || !clientSecret || !siteUrl) {
    const missing = ['SP_TENANT_ID','SP_CLIENT_ID','SP_CLIENT_SECRET','SP_SITE_URL']
      .filter(k => !process.env[k]).join(', ');
    context.res = { status: 500, body: JSON.stringify({ error: 'Missing env vars: ' + missing }) };
    return;
  }

  const { action, payload } = req.body || {};
  if (!action) {
    context.res = { status: 400, body: JSON.stringify({ error: 'Missing action' }) };
    return;
  }

  try {
    const token = await getAccessToken(tenantId, clientId, clientSecret);

    // Derive site path from URL
    const urlParts  = siteUrl.replace('https://', '').split('/');
    const spHost    = urlParts[0];
    const spPath    = '/' + urlParts.slice(1).join('/');
    const siteIdPath = '/v1.0/sites/' + spHost + ':' + spPath;

    let result;

    if (action === 'getSiteId') {
      result = await graphRequest(token, 'GET', siteIdPath, null);

    } else if (action === 'getDriveId') {
      result = await graphRequest(token, 'GET',
        '/v1.0/sites/' + payload.siteId + '/drives', null);

    } else if (action === 'createFolder') {
      const parentPath = payload.parentPath.replace(/^\//, '');
      result = await graphRequest(token, 'POST',
        '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId +
        '/root:/' + parentPath + ':/children',
        {
          name:   payload.folderName,
          folder: {},
          '@microsoft.graph.conflictBehavior': 'rename'
        }
      );

    } else if (action === 'listFolder') {
      const folderPath = payload.folderPath.replace(/^\//, '');
      result = await graphRequest(token, 'GET',
        '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId +
        '/root:/' + folderPath + ':/children?$orderby=name&$top=200', null);

    } else if (action === 'uploadFile') {
      const fileBuffer  = Buffer.from(payload.fileContent, 'base64');
      const filePath    = payload.filePath.replace(/^\//, '');

      result = await new Promise((resolve, reject) => {
        const options = {
          hostname: 'graph.microsoft.com',
          path:     '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId +
                    '/root:/' + filePath + ':/content',
          method:   'PUT',
          headers:  {
            'Authorization':  'Bearer ' + token,
            'Content-Type':   payload.contentType || 'application/octet-stream',
            'Content-Length': fileBuffer.length
          }
        };
        const req = https.request(options, (res) => {
          let data = '';
          res.on('data', chunk => { data += chunk; });
          res.on('end', () => resolve({ status: res.statusCode, body: data }));
        });
        req.on('error', reject);
        req.setTimeout(60000, () => { req.destroy(); reject(new Error('Upload timed out')); });
        req.write(fileBuffer);
        req.end();
      });

    } else if (action === 'updateMetadata') {
      result = await graphRequest(token, 'PATCH',
        '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId +
        '/items/' + payload.itemId + '/listItem/fields',
        payload.metadata
      );

    } else if (action === 'createShareLink') {
      const expiry = new Date();
      expiry.setDate(expiry.getDate() + 90);
      result = await graphRequest(token, 'POST',
        '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId +
        '/items/' + payload.itemId + '/createLink',
        { type: 'view', scope: 'anonymous', expirationDateTime: expiry.toISOString() }
      );

    } else if (action === 'deleteItem') {
      result = await graphRequest(token, 'DELETE',
        '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId +
        '/items/' + payload.itemId, null);

    } else if (action === 'getDownloadUrl') {
      result = await graphRequest(token, 'GET',
        '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId +
        '/items/' + payload.itemId + '?select=id,name,@microsoft.graph.downloadUrl', null);

    } else if (action === 'searchFiles') {
      result = await graphRequest(token, 'GET',
        '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId +
        '/root/search(q=\'' + encodeURIComponent(payload.query) + '\')?$top=50', null);

    } else if (action === 'testConnection') {
      // Diagnostic action — returns full detail of what's happening
      const tokenOk = !!token;
      const siteResult = await graphRequest(token, 'GET', siteIdPath, null);
      let siteBody;
      try { siteBody = JSON.parse(siteResult.body); } catch(e) { siteBody = siteResult.body; }
      context.res = {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          tokenOk,
          siteUrl,
          siteIdPath,
          siteStatus: siteResult.status,
          siteResponse: siteBody
        })
      };
      return;

    } else {
      context.res = { status: 400, body: JSON.stringify({ error: 'Unknown action: ' + action }) };
      return;
    }

    // Parse body and surface any Graph errors clearly
    let parsed;
    try { parsed = JSON.parse(result.body); } catch(e) { parsed = { rawBody: result.body }; }

    if (result.status >= 400) {
      const errMsg = (parsed.error && (parsed.error.message || parsed.error.code)) || result.body;
      context.log.error('Graph API error:', result.status, errMsg);
    }

    context.res = {
      status:  result.status,
      headers: { 'Content-Type': 'application/json' },
      body:    result.body
    };

  } catch (err) {
    context.log.error('SharePoint proxy error:', err.message);
    context.res = {
      status: 502,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ error: err.message })
    };
  }
};
