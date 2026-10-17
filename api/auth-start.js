import crypto from 'crypto';
export default function handler(req,res) {
  const clientId = process.env.MS_CLIENT_ID;
  const redirectUri = process.env.MS_REDIRECT_URI;
  if (!clientId || !redirectUri) return res.status(503).send('MS_CLIENT_ID en MS_REDIRECT_URI ontbreken.');
  const state = crypto.randomBytes(24).toString('hex');
  const auth = new URL('https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize');
  auth.searchParams.set('client_id',clientId);
  auth.searchParams.set('response_type','code');
  auth.searchParams.set('redirect_uri',redirectUri);
  auth.searchParams.set('response_mode','query');
  auth.searchParams.set('scope','openid offline_access Files.ReadWrite');
  auth.searchParams.set('state',state);
  res.setHeader('Set-Cookie',`mrm_oauth_state=${state}; HttpOnly; Secure; SameSite=Lax; Path=/api; Max-Age=600`);
  res.redirect(auth.toString());
}
