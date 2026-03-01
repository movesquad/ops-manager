// Timer-triggered Azure Function — runs daily at 07:00 UTC
// Checks all jobs for missing required docs at 7-day and 48-hour thresholds
const https = require('https');

async function getToken(tenantId, clientId, clientSecret) {
  const body = ['grant_type=client_credentials','scope='+encodeURIComponent('https://graph.microsoft.com/.default'),
    'client_id='+encodeURIComponent(clientId),'client_secret='+encodeURIComponent(clientSecret)].join('&');
  return new Promise((resolve, reject) => {
    const opts = { hostname:'login.microsoftonline.com', path:'/'+tenantId+'/oauth2/v2.0/token', method:'POST',
      headers:{'Content-Type':'application/x-www-form-urlencoded','Content-Length':Buffer.byteLength(body)} };
    const req = https.request(opts, res => { let d=''; res.on('data',c=>{d+=c;}); res.on('end',()=>{ try{const p=JSON.parse(d); resolve(p.access_token||null);}catch{resolve(null);} }); });
    req.on('error', reject); req.write(body); req.end();
  });
}

async function graphPost(token, path, body) {
  const bodyStr = JSON.stringify(body);
  return new Promise((resolve, reject) => {
    const opts = { hostname:'graph.microsoft.com', path, method:'POST',
      headers:{'Authorization':'Bearer '+token,'Content-Type':'application/json','Content-Length':Buffer.byteLength(bodyStr)} };
    const req = https.request(opts, res => { let d=''; res.on('data',c=>{d+=c;}); res.on('end',()=>resolve({status:res.statusCode,body:d})); });
    req.on('error', reject); req.write(bodyStr); req.end();
  });
}

async function tableGet(accountName, accountKey, table) {
  // Simple Azure Table Storage REST GET
  const date = new Date().toUTCString();
  const resource = '/'+accountName+'/'+table+'()';
  const crypto = require('crypto');
  const strToSign = 'GET\n\n\n'+date+'\n'+resource;
  const sig = crypto.createHmac('sha256', Buffer.from(accountKey,'base64')).update(strToSign,'utf8').digest('base64');
  const auth = 'SharedKey '+accountName+':'+sig;
  return new Promise((resolve, reject) => {
    const opts = { hostname: accountName+'.table.core.windows.net', path: '/'+table+'()?$top=1000',
      method:'GET', headers:{'Authorization':auth,'Date':date,'Accept':'application/json;odata=nometadata','x-ms-version':'2019-02-02'} };
    const req = https.request(opts, res => { let d=''; res.on('data',c=>{d+=c;}); res.on('end',()=>{ try{resolve(JSON.parse(d).value||[]);}catch{resolve([]);} }); });
    req.on('error', reject); req.end();
  });
}

module.exports = async function (context, myTimer) {
  context.log('Doc reminder function running');

  const tenantId     = process.env.SP_TENANT_ID;
  const clientId     = process.env.SP_CLIENT_ID;
  const clientSecret = process.env.SP_CLIENT_SECRET;
  const mailFrom     = process.env.MAIL_FROM || 'updates@onwards.network';
  const storageAcct  = process.env.STORAGE_ACCOUNT;
  const storageKey   = process.env.STORAGE_KEY;

  if (!tenantId || !clientId || !clientSecret || !storageAcct || !storageKey) {
    context.log.error('Missing env vars'); return;
  }

  const token = await getToken(tenantId, clientId, clientSecret);
  if (!token) { context.log.error('Token failed'); return; }

  // Load client jobs and move managers from table storage
  const [clientJobs, moveManagers] = await Promise.all([
    tableGet(storageAcct, storageKey, 'OpsClientJobs'),
    tableGet(storageAcct, storageKey, 'OpsMoveManagers')
  ]);

  const now = new Date();
  const today = now.toISOString().split('T')[0];
  let sent = 0;

  // Required docs by slug — must match SP_CHECKLIST in the app
  const REQUIRED_SLUGS = ['packing-list','insurance','survey-report','delivery-order'];

  for (const rawJob of clientJobs) {
    let cj;
    try { cj = JSON.parse(rawJob.data || '{}'); } catch { continue; }

    if (!cj.startDate || cj.status === 'Cancelled' || cj.status === 'Completed' || cj.status === 'Draft') continue;
    if (!cj.moveManager) continue;

    // Calculate days until job
    const jobDate = new Date(cj.startDate + 'T12:00:00');
    const daysUntil = Math.round((jobDate - now) / (1000 * 60 * 60 * 24));

    // Only send at 7-day or 2-day threshold
    if (daysUntil !== 7 && daysUntil !== 2) continue;

    // Check which required docs are missing
    const checklist = cj.docChecklist || {};
    const missingDocs = REQUIRED_SLUGS.filter(slug => !checklist[slug] || checklist[slug] === 'missing');
    if (!missingDocs.length) continue; // All docs present — no reminder needed

    // Find move manager email
    const mm = moveManagers.find(m => { try { const p=JSON.parse(m.data||'{}'); return p.name===cj.moveManager; } catch{return false;} });
    let mmEmail = '';
    if (mm) { try { mmEmail = JSON.parse(mm.data||'{}').email||''; } catch{} }
    if (!mmEmail) continue;

    const thresholdLabel = daysUntil === 7 ? '7 days' : '48 hours';
    const jobDateNice    = jobDate.toLocaleDateString('en-GB',{weekday:'long',day:'numeric',month:'long',year:'numeric'});
    const clientName     = cj.clientName || '—';
    const partnerRef     = cj.partnerRef || cj.clientRef || '—';
    const ourRef         = cj.masterJobId || '—';

    // Map slugs to readable names
    const slugLabels = {
      'packing-list':   'Packing List / Inventory',
      'insurance':      'Insurance Certificate',
      'survey-report':  'Survey Report',
      'delivery-order': 'Delivery Order / Instructions'
    };
    const missingList = missingDocs.map(s => slugLabels[s] || s).join(', ');
    const docsUrl = cj.spFolderUrl || '';

    const subject = 'Action Required — Missing Documents: ' + partnerRef + ' / ' + ourRef + ' — ' + clientName;
    const bodyHtml = '<p>Dear ' + mmEmail.split('@')[0] + ',</p>'
      + '<p>This is an automated reminder that the following job is due in <strong>' + thresholdLabel + '</strong> and has outstanding required documents:</p>'
      + '<table style="width:100%;border-collapse:collapse;margin:16px 0">'
      + '<tr><td style="padding:8px 14px;border:1px solid #e8eaed;font-weight:600;font-size:13px;color:#374151;background:#f8f9fb;width:150px">Client</td><td style="padding:8px 14px;border:1px solid #e8eaed;font-size:13px">' + clientName + '</td></tr>'
      + '<tr><td style="padding:8px 14px;border:1px solid #e8eaed;font-weight:600;font-size:13px;color:#374151;background:#f8f9fb">Your Reference</td><td style="padding:8px 14px;border:1px solid #e8eaed;font-size:13px">' + partnerRef + '</td></tr>'
      + '<tr><td style="padding:8px 14px;border:1px solid #e8eaed;font-weight:600;font-size:13px;color:#374151;background:#f8f9fb">Our Reference</td><td style="padding:8px 14px;border:1px solid #e8eaed;font-size:13px">' + ourRef + '</td></tr>'
      + '<tr><td style="padding:8px 14px;border:1px solid #e8eaed;font-weight:600;font-size:13px;color:#374151;background:#f8f9fb">Job Date</td><td style="padding:8px 14px;border:1px solid #e8eaed;font-size:13px">' + jobDateNice + '</td></tr>'
      + '</table>'
      + '<div style="background:#fff3cd;border:1px solid #ffc107;border-radius:8px;padding:16px;margin:20px 0">'
      + '<div style="font-weight:700;color:#856404;margin-bottom:8px">⚠ Missing Documents</div>'
      + '<div style="color:#533f03;font-size:13px">' + missingList + '</div>'
      + '</div>'
      + (docsUrl ? '<p>Please upload the missing documents to the job folder:<br><a href="' + docsUrl + '" style="color:#0073EA">' + docsUrl + '</a></p>' : '')
      + '<p>If these documents have already been sent please disregard this message.</p>'
      + '<p>Kind regards,<br><strong>Onwards Operations</strong></p>';

    const emailBody = JSON.stringify({
      message: {
        subject,
        body: { contentType: 'HTML', content: bodyHtml },
        toRecipients: [{ emailAddress: { address: mmEmail } }],
        from: { emailAddress: { address: mailFrom } }
      },
      saveToSentItems: true
    });

    const result = await graphPost(token,
      '/v1.0/users/' + encodeURIComponent(mailFrom) + '/sendMail',
      JSON.parse(emailBody)
    );

    if (result.status === 202) {
      context.log('Reminder sent to', mmEmail, 'for job', ourRef, '('+daysUntil+' days)');
      sent++;
    } else {
      context.log.error('Reminder failed for', ourRef, result.status, result.body.slice(0,200));
    }
  }

  context.log('Doc reminder run complete —', sent, 'reminders sent');
};
