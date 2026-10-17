export default function handler(req,res) {
  const configured = Boolean(process.env.MS_REFRESH_TOKEN && process.env.MS_CLIENT_ID && process.env.MS_CLIENT_SECRET);
  res.status(200).json({ok:true,microsoftConfigured:configured});
}
