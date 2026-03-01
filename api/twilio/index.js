const https = require('https');

async function twilioRequest(accountSid, authToken, path, params) {
  const body = Object.entries(params).map(([k,v]) => encodeURIComponent(k)+'='+encodeURIComponent(v)).join('&');
  const auth  = Buffer.from(accountSid+':'+authToken).toString('base64');
  return new Promise((resolve, reject) => {
    const opts = {
      hostname: 'api.twilio.com',
      path: '/2010-04-01/Accounts/' + accountSid + path,
      method: 'POST',
      headers: { 'Authorization': 'Basic '+auth, 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) }
    };
    const req = https.request(opts, res => { let d=''; res.on('data',c=>{d+=c;}); res.on('end',()=>resolve({status:res.statusCode,body:d})); });
    req.on('error', reject);
    req.setTimeout(30000, () => { req.destroy(); reject(new Error('Twilio timeout')); });
    req.write(body); req.end();
  });
}

module.exports = async function(context, req) {
  if (req.method !== 'POST') { context.res = {status:405,body:'Method not allowed'}; return; }

  const accountSid = process.env.TWILIO_ACCOUNT_SID;
  const authToken  = process.env.TWILIO_AUTH_TOKEN;
  const fromNumber = process.env.TWILIO_FROM_NUMBER;

  if (!accountSid || !authToken || !fromNumber) {
    context.res = {status:500,headers:{'Content-Type':'application/json'},body:JSON.stringify({error:'Twilio env vars not configured: TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN, TWILIO_FROM_NUMBER'})};
    return;
  }

  const { action, to, body: msgBody, from, mediaUrl } = req.body || {};

  try {
    let result;

    if (action === 'sendSms' || !action) {
      result = await twilioRequest(accountSid, authToken, '/Messages.json', {
        To:   to,
        From: from || fromNumber,
        Body: msgBody || ''
      });

    } else if (action === 'sendWhatsApp') {
      const fromWa = process.env.TWILIO_WHATSAPP_NUMBER || ('whatsapp:' + fromNumber);
      result = await twilioRequest(accountSid, authToken, '/Messages.json', {
        To:   'whatsapp:' + to.replace('whatsapp:',''),
        From: fromWa,
        Body: msgBody || ''
      });

    } else if (action === 'makeCall') {
      // Initiate an outbound call with TwiML
      const twiml = req.body.twiml || '<Response><Say>' + (msgBody||'Hello from Onwards Operations') + '</Say></Response>';
      result = await twilioRequest(accountSid, authToken, '/Calls.json', {
        To:   to,
        From: from || fromNumber,
        Twiml: twiml
      });

    } else {
      context.res = {status:400,headers:{'Content-Type':'application/json'},body:JSON.stringify({error:'Unknown action: '+action})};
      return;
    }

    let parsed; try { parsed = JSON.parse(result.body); } catch(e) { parsed = {}; }

    if (result.status >= 400) {
      const errMsg = parsed.message || parsed.error_message || result.body.slice(0,200);
      context.log.error('Twilio error', result.status, action, errMsg);
      context.res = {status:result.status,headers:{'Content-Type':'application/json'},body:JSON.stringify({error:errMsg,twilioStatus:result.status})};
      return;
    }

    context.res = {status:200,headers:{'Content-Type':'application/json'},body:JSON.stringify({ok:true,sid:parsed.sid,status:parsed.status})};

  } catch(err) {
    context.log.error('Twilio function error:', err.message);
    context.res = {status:502,headers:{'Content-Type':'application/json'},body:JSON.stringify({error:err.message})};
  }
};
