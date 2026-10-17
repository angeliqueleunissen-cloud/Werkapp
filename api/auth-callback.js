import crypto from 'crypto';
function cookie(req,name){const h=req.headers.cookie||'';const m=h.split(';').map(x=>x.trim()).find(x=>x.startsWith(name+'='));return m?decodeURIComponent(m.slice(name.length+1)):'';}
export default async function handler(req,res){
  try{
    const {code,state,error,error_description} = req.query || {};
    if(error) return res.status(400).send(`Microsoft gaf een fout: ${error_description || error}`);
    if (!code || !state || String(state) !== String(cookie(req, 'mrm_oauth_state'))) {
  return res.status(400).send('Ongeldige of verlopen autorisatie.');
}
    const body=new URLSearchParams({client_id:process.env.MS_CLIENT_ID,client_secret:process.env.MS_CLIENT_SECRET,grant_type:'authorization_code',code:String(code),redirect_uri:process.env.MS_REDIRECT_URI,scope:'openid offline_access Files.ReadWrite'});
    const r=await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body});
    const d=await r.json();
    if(!r.ok || !d.refresh_token) return res.status(400).send(`Token ophalen mislukt: ${d.error_description || 'onbekende fout'}`);
    res.setHeader('Set-Cookie','mrm_oauth_state=; HttpOnly; Secure; SameSite=Lax; Path=/api; Max-Age=0');
    res.setHeader('Content-Type','text/html; charset=utf-8');
    res.status(200).send(`<!doctype html><meta charset="utf-8"><title>MRM-01 koppeling</title><style>body{font-family:Arial;max-width:760px;margin:40px auto;padding:20px}textarea{width:100%;height:130px}button{padding:12px 18px}</style><h2>Microsoft-koppeling gelukt</h2><p>Kopieer deze <b>refresh token</b> en zet hem als gevoelige Vercel-variabele <code>MS_REFRESH_TOKEN</code>. Deel deze token met niemand.</p><textarea id="t" readonly>${String(d.refresh_token).replace(/&/g,'&amp;').replace(/</g,'&lt;')}</textarea><button onclick="navigator.clipboard.writeText(document.getElementById('t').value)">Kopiëren</button><p>Daarna moet de Vercel-deployment opnieuw worden uitgevoerd.</p>`);
  }catch(e){res.status(500).send('Fout: '+e.message)}
}
