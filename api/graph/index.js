const https = require('https');

let tokenCache = { token: null, expiresAt: 0 };

async function getToken(tenantId, clientId, clientSecret) {
  const now = Date.now();
  if (tokenCache.token && now < tokenCache.expiresAt - 60000) return tokenCache.token;
  const body = ['grant_type=client_credentials','scope='+encodeURIComponent('https://graph.microsoft.com/.default'),
    'client_id='+encodeURIComponent(clientId),'client_secret='+encodeURIComponent(clientSecret)].join('&');
  const result = await new Promise((resolve, reject) => {
    const opts = { hostname:'login.microsoftonline.com', path:'/'+tenantId+'/oauth2/v2.0/token', method:'POST',
      headers:{'Content-Type':'application/x-www-form-urlencoded','Content-Length':Buffer.byteLength(body)} };
    const req = https.request(opts, res => { let d=''; res.on('data',c=>{d+=c;}); res.on('end',()=>resolve({status:res.statusCode,body:d})); });
    req.on('error', reject); req.setTimeout(15000, () => { req.destroy(); reject(new Error('timeout')); });
    req.write(body); req.end();
  });
  const parsed = JSON.parse(result.body);
  if (!parsed.access_token) throw new Error('Token failed: ' + (parsed.error_description||parsed.error||result.body.slice(0,100)));
  tokenCache.token = parsed.access_token;
  tokenCache.expiresAt = now + (parsed.expires_in * 1000);
  return tokenCache.token;
}

async function graphRequest(token, method, path, body) {
  const bodyStr = body ? JSON.stringify(body) : null;
  return new Promise((resolve, reject) => {
    const opts = { hostname:'graph.microsoft.com', path, method,
      headers:{ 'Authorization':'Bearer '+token, 'Accept':'application/json',
        ...(bodyStr ? {'Content-Type':'application/json','Content-Length':Buffer.byteLength(bodyStr)} : {}) } };
    const req = https.request(opts, res => { let d=''; res.on('data',c=>{d+=c;}); res.on('end',()=>resolve({status:res.statusCode,body:d})); });
    req.on('error', reject); req.setTimeout(30000, () => { req.destroy(); reject(new Error('timeout')); });
    if (bodyStr) req.write(bodyStr); req.end();
  });
}

module.exports = async function(context, req) {
  if (req.method !== 'POST') { context.res = {status:405,body:'Method not allowed'}; return; }

  const tenantId = process.env.SP_TENANT_ID, clientId = process.env.SP_CLIENT_ID, clientSecret = process.env.SP_CLIENT_SECRET;
  const mailFrom = process.env.MAIL_FROM || 'updates@onwards.network';

  try {
    const token = await getToken(tenantId, clientId, clientSecret);
    const { action, payload, userEmail } = req.body || {};

    let result;

    if (action === 'createCalendarEvent') {
      // Create event in the sending mailbox and invite the move manager as attendee
      // Using mailFrom as the organiser calendar
      result = await graphRequest(token, 'POST',
        '/v1.0/users/' + encodeURIComponent(mailFrom) + '/events',
        payload
      );

    } else if (action === 'updateCalendarEvent') {
      result = await graphRequest(token, 'PATCH',
        '/v1.0/users/' + encodeURIComponent(mailFrom) + '/events/' + payload.eventId,
        payload.updates
      );

    } else if (action === 'deleteCalendarEvent') {
      result = await graphRequest(token, 'DELETE',
        '/v1.0/users/' + encodeURIComponent(mailFrom) + '/events/' + payload.eventId,
        null
      );

    } else if (action === 'upsertContact') {
      // Search for existing contact by email first
      const searchResult = await graphRequest(token, 'GET',
        '/v1.0/users/' + encodeURIComponent(mailFrom) + '/contacts?$filter=emailAddresses/any(e:e/address eq \'' + encodeURIComponent(payload.email||'') + '\')&$top=1',
        null
      );
      let searchBody; try { searchBody = JSON.parse(searchResult.body); } catch(e) { searchBody = {}; }
      const existing = searchBody.value && searchBody.value[0];

      const contactBody = {
        displayName:    payload.displayName || '',
        emailAddresses: payload.email ? [{ address: payload.email, name: payload.displayName||'' }] : [],
        businessPhones: payload.phone ? [payload.phone] : [],
        companyName:    payload.company || '',
        personalNotes:  payload.notes  || ''
      };

      if (existing) {
        // Update existing
        result = await graphRequest(token, 'PATCH',
          '/v1.0/users/' + encodeURIComponent(mailFrom) + '/contacts/' + existing.id,
          contactBody
        );
      } else {
        // Create new
        result = await graphRequest(token, 'POST',
          '/v1.0/users/' + encodeURIComponent(mailFrom) + '/contacts',
          contactBody
        );
      }

    } else {
      context.res = {status:400,headers:{'Content-Type':'application/json'},body:JSON.stringify({error:'Unknown action: '+action})};
      return;
    }

    let parsed; try { parsed = JSON.parse(result.body); } catch(e) { parsed = {}; }
    if (result.status >= 400) {
      const errMsg = (parsed.error && (parsed.error.message || parsed.error.code)) || result.body.slice(0,200);
      context.log.error('Graph error', result.status, action, errMsg);
      context.res = {status:result.status,headers:{'Content-Type':'application/json'},body:JSON.stringify({error:errMsg})};
      return;
    }

    context.res = {status:result.status||200,headers:{'Content-Type':'application/json'},body:result.body||'{}'};

  } catch(err) {
    context.log.error('Graph function error:', err.message);
    context.res = {status:502,headers:{'Content-Type':'application/json'},body:JSON.stringify({error:err.message})};
  }
};
