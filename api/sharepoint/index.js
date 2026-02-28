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
    req.setTimeout(15000, () => { req.destroy(); reject(new Error('Token request timed out')); });
    req.write(body); req.end();
  });
  let parsed; try { parsed = JSON.parse(result.body); } catch(e) { parsed = {}; }
  if (result.status !== 200 || !parsed.access_token) {
    const errMsg = parsed.error_description || parsed.error || ('HTTP ' + result.status + ': ' + result.body.slice(0,200));
    return { token: null, error: 'Token error: ' + errMsg };
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
    req.setTimeout(30000, () => { req.destroy(); reject(new Error('Graph request timed out')); });
    if (bodyStr) req.write(bodyStr);
    req.end();
  });
}

module.exports = async function (context, req) {
  context.log('SharePoint proxy:', req.body && req.body.action);
  if (req.method !== 'POST') { context.res = { status: 405, body: 'Method not allowed' }; return; }

  const tenantId     = process.env.SP_TENANT_ID;
  const clientId     = process.env.SP_CLIENT_ID;
  const clientSecret = process.env.SP_CLIENT_SECRET;
  const siteUrl      = process.env.SP_SITE_URL;

  const missing = ['SP_TENANT_ID','SP_CLIENT_ID','SP_CLIENT_SECRET','SP_SITE_URL'].filter(k => !process.env[k]);
  if (missing.length) {
    context.res = { status: 500, headers: {'Content-Type':'application/json'}, body: JSON.stringify({ error: 'Missing env vars: ' + missing.join(', ') }) };
    return;
  }

  const { action, payload } = req.body || {};
  if (!action) { context.res = { status: 400, headers: {'Content-Type':'application/json'}, body: JSON.stringify({ error: 'Missing action' }) }; return; }

  try {
    const { token, error: tokenError } = await getAccessToken(tenantId, clientId, clientSecret);
    if (!token) {
      context.log.error('Token failure:', tokenError);
      context.res = { status: 401, headers: {'Content-Type':'application/json'}, body: JSON.stringify({ error: tokenError }) };
      return;
    }

    const urlParts   = siteUrl.replace('https://', '').split('/');
    const spHost     = urlParts[0];
    const spPath     = '/' + urlParts.slice(1).join('/');
    const siteIdPath = '/v1.0/sites/' + spHost + ':' + spPath;

    let result;

    if (action === 'testConnection') {
      const siteResult = await graphRequest(token, 'GET', siteIdPath, null);
      let siteBody; try { siteBody = JSON.parse(siteResult.body); } catch(e) { siteBody = { raw: siteResult.body.slice(0,300) }; }
      context.res = { status: 200, headers: {'Content-Type':'application/json'},
        body: JSON.stringify({ tokenOk: true, siteUrl, siteIdPath, siteStatus: siteResult.status, siteResponse: siteBody }) };
      return;

    } else if (action === 'getSiteId') {
      result = await graphRequest(token, 'GET', siteIdPath, null);

    } else if (action === 'getDriveId') {
      result = await graphRequest(token, 'GET', '/v1.0/sites/' + payload.siteId + '/drives', null);

    } else if (action === 'createFolder') {
      // Create folder by path (for top-level folders where path is simple)
      const parentPath = (payload.parentPath || '').replace(/^\//, '');
      const encodedPath = parentPath.split('/').map(p => encodeURIComponent(p)).join('/');
      result = await graphRequest(token, 'POST',
        '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId + '/root:/' + encodedPath + ':/children',
        { name: payload.folderName, folder: {}, '@microsoft.graph.conflictBehavior': 'rename' });

    } else if (action === 'createFolderById') {
      // Create folder by parent item ID — avoids path encoding issues with special characters
      result = await graphRequest(token, 'POST',
        '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId + '/items/' + payload.parentId + '/children',
        { name: payload.folderName, folder: {}, '@microsoft.graph.conflictBehavior': 'rename' });

    } else if (action === 'listFolder') {
      const folderPath = (payload.folderPath || '').replace(/^\//, '');
      const encodedPath = folderPath.split('/').map(p => encodeURIComponent(p)).join('/');
      result = await graphRequest(token, 'GET',
        '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId + '/root:/' + encodedPath + ':/children?$orderby=name&$top=200', null);

    } else if (action === 'uploadFile') {
      // Upload by parent folder ID to avoid path encoding issues
      const fileBuffer = Buffer.from(payload.fileContent, 'base64');
      const fileName   = payload.fileName || 'upload';
      const uploadPath = payload.parentId
        ? '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId + '/items/' + payload.parentId + ':/' + encodeURIComponent(fileName) + ':/content'
        : '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId + '/root:/' + (payload.filePath || '').replace(/^\//, '') + ':/content';

      result = await new Promise((resolve, reject) => {
        const opts = {
          hostname: 'graph.microsoft.com', path: uploadPath, method: 'PUT',
          headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': payload.contentType || 'application/octet-stream', 'Content-Length': fileBuffer.length }
        };
        const req = https.request(opts, (res) => { let d=''; res.on('data',c=>{d+=c;}); res.on('end',()=>resolve({status:res.statusCode,body:d})); });
        req.on('error', reject);
        req.setTimeout(60000, () => { req.destroy(); reject(new Error('Upload timed out')); });
        req.write(fileBuffer); req.end();
      });

    } else if (action === 'getFolderChildren') {
      // List children of a folder by ID — for finding sub-folder IDs
      result = await graphRequest(token, 'GET',
        '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId + '/items/' + payload.itemId + '/children?$top=50', null);

    } else if (action === 'updateMetadata') {
      result = await graphRequest(token, 'PATCH',
        '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId + '/items/' + payload.itemId + '/listItem/fields',
        payload.metadata);

    } else if (action === 'createShareLink') {
      const expiry = new Date(); expiry.setDate(expiry.getDate() + 90);
      result = await graphRequest(token, 'POST',
        '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId + '/items/' + payload.itemId + '/createLink',
        { type: 'view', scope: 'anonymous', expirationDateTime: expiry.toISOString() });

    } else if (action === 'deleteItem') {
      result = await graphRequest(token, 'DELETE',
        '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId + '/items/' + payload.itemId, null);

    } else if (action === 'getDownloadUrl') {
      result = await graphRequest(token, 'GET',
        '/v1.0/sites/' + payload.siteId + '/drives/' + payload.driveId + '/items/' + payload.itemId + '?select=id,name,@microsoft.graph.downloadUrl', null);

    } else {
      context.res = { status: 400, headers: {'Content-Type':'application/json'}, body: JSON.stringify({ error: 'Unknown action: ' + action }) };
      return;
    }

    let parsed; try { parsed = JSON.parse(result.body); } catch(e) { parsed = { rawBody: result.body.slice(0,300) }; }

    if (result.status >= 400) {
      const errMsg = (parsed.error && (parsed.error.message || parsed.error.code)) || result.body.slice(0,200);
      context.log.error('Graph error ' + result.status + ' on ' + action + ':', errMsg);
      context.res = { status: result.status, headers: {'Content-Type':'application/json'}, body: JSON.stringify({ error: errMsg, graphStatus: result.status, action }) };
      return;
    }

    context.res = { status: result.status, headers: {'Content-Type':'application/json'}, body: result.body };

  } catch (err) {
    context.log.error('SharePoint proxy exception:', err.message);
    context.res = { status: 502, headers: {'Content-Type':'application/json'}, body: JSON.stringify({ error: err.message }) };
  }
};
