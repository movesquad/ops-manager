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
    throw new Error('Token fetch failed: ' + result.body);
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
  const siteUrl      = process.env.SP_SITE_URL; // https://movesquadgroup.sharepoint.com/sites/MoveSquadGroup

  if (!tenantId || !clientId || !clientSecret || !siteUrl) {
    context.log.error('SharePoint credentials not fully configured');
    context.res = { status: 500, body: 'SharePoint environment variables not configured' };
    return;
  }

  const { action, payload } = req.body || {};
  if (!action) {
    context.res = { status: 400, body: 'Missing action in request body' };
    return;
  }

  try {
    const token = await getAccessToken(tenantId, clientId, clientSecret);

    // ── Derive site ID from site URL ────────────────────────────────────
    // GET /v1.0/sites/{hostname}:{path}
    const urlObj    = new URL(siteUrl);
    const hostname  = urlObj.hostname;                          // movesquadgroup.sharepoint.com
    const sitePath  = urlObj.pathname;                          // /sites/MoveSquadGroup
    const siteIdPath = '/v1.0/sites/' + hostname + ':' + sitePath;

    let result;

    // ════════════════════════════════════════════════════════════════════
    // ACTION: getSiteId — resolves the Graph site ID for the SharePoint site
    // ════════════════════════════════════════════════════════════════════
    if (action === 'getSiteId') {
      result = await graphRequest(token, 'GET', siteIdPath, null);

    // ════════════════════════════════════════════════════════════════════
    // ACTION: getDriveId — gets the document library drive ID
    // payload: { siteId }
    // ════════════════════════════════════════════════════════════════════
    } else if (action === 'getDriveId') {
      result = await graphRequest(token, 'GET',
        '/v1.0/sites/' + payload.siteId + '/drives', null);

    // ════════════════════════════════════════════════════════════════════
    // ACTION: createFolder — creates a folder (and parents if needed)
    // payload: { siteId, driveId, parentPath, folderName }
    // ════════════════════════════════════════════════════════════════════
    } else if (action === 'createFolder') {
      const encodedPath = encodeURIComponent(payload.parentPath).replace(/%2F/g, '/');
      result = await graphRequest(token, 'POST',
        '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId +
        '/root:/' + encodedPath + ':/children',
        {
          name:   payload.folderName,
          folder: {},
          '@microsoft.graph.conflictBehavior': 'rename'
        }
      );

    // ════════════════════════════════════════════════════════════════════
    // ACTION: listFolder — lists contents of a folder
    // payload: { siteId, driveId, folderPath }
    // ════════════════════════════════════════════════════════════════════
    } else if (action === 'listFolder') {
      const encodedPath = encodeURIComponent(payload.folderPath).replace(/%2F/g, '/');
      result = await graphRequest(token, 'GET',
        '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId +
        '/root:/' + encodedPath + ':/children?$orderby=name&$top=200', null);

    // ════════════════════════════════════════════════════════════════════
    // ACTION: uploadFile — uploads a file (up to ~4MB via simple upload)
    // payload: { siteId, driveId, filePath, fileContent (base64), contentType }
    // ════════════════════════════════════════════════════════════════════
    } else if (action === 'uploadFile') {
      const fileBuffer  = Buffer.from(payload.fileContent, 'base64');
      const encodedPath = encodeURIComponent(payload.filePath).replace(/%2F/g, '/');

      // Simple upload (files < 4MB)
      result = await new Promise((resolve, reject) => {
        const options = {
          hostname: 'graph.microsoft.com',
          path:     '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId +
                    '/root:/' + encodedPath + ':/content',
          method:   'PUT',
          headers:  {
            'Authorization': 'Bearer ' + token,
            'Content-Type':  payload.contentType || 'application/octet-stream',
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

    // ════════════════════════════════════════════════════════════════════
    // ACTION: updateMetadata — writes custom metadata columns to a file
    // payload: { siteId, driveId, itemId, metadata: { key: value } }
    // ════════════════════════════════════════════════════════════════════
    } else if (action === 'updateMetadata') {
      result = await graphRequest(token, 'PATCH',
        '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId +
        '/items/' + payload.itemId + '/listItem/fields',
        payload.metadata
      );

    // ════════════════════════════════════════════════════════════════════
    // ACTION: createShareLink — generates a 90-day view-only sharing link
    // payload: { siteId, driveId, itemId }
    // ════════════════════════════════════════════════════════════════════
    } else if (action === 'createShareLink') {
      const expiryDate = new Date();
      expiryDate.setDate(expiryDate.getDate() + 90);

      result = await graphRequest(token, 'POST',
        '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId +
        '/items/' + payload.itemId + '/createLink',
        {
          type:            'view',
          scope:           'anonymous',
          expirationDateTime: expiryDate.toISOString()
        }
      );

    // ════════════════════════════════════════════════════════════════════
    // ACTION: deleteItem — deletes a file or folder
    // payload: { siteId, driveId, itemId }
    // ════════════════════════════════════════════════════════════════════
    } else if (action === 'deleteItem') {
      result = await graphRequest(token, 'DELETE',
        '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId +
        '/items/' + payload.itemId, null);

    // ════════════════════════════════════════════════════════════════════
    // ACTION: getDownloadUrl — gets a temporary download URL for a file
    // payload: { siteId, driveId, itemId }
    // ════════════════════════════════════════════════════════════════════
    } else if (action === 'getDownloadUrl') {
      result = await graphRequest(token, 'GET',
        '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId +
        '/items/' + payload.itemId + '?select=id,name,@microsoft.graph.downloadUrl', null);

    // ════════════════════════════════════════════════════════════════════
    // ACTION: searchFiles — searches across the library
    // payload: { siteId, driveId, query }
    // ════════════════════════════════════════════════════════════════════
    } else if (action === 'searchFiles') {
      result = await graphRequest(token, 'GET',
        '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId +
        '/root/search(q=\'' + encodeURIComponent(payload.query) + '\')?$top=50', null);

    } else {
      context.res = { status: 400, body: 'Unknown action: ' + action };
      return;
    }

    context.res = {
      status:  result.status,
      headers: { 'Content-Type': 'application/json' },
      body:    result.body
    };

  } catch (err) {
    context.log.error('SharePoint proxy error:', err.message);
    context.res = { status: 502, body: 'Proxy error: ' + err.message };
  }
};
