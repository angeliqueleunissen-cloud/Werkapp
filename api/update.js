import { setCell } from './_graph.js';

const allowed = new Set(['Civiel','Blazen','Lassen','Hassen','Gereed','Probleem','Schouw']);
export default async function handler(req,res) {
  if (req.method !== 'POST') return res.status(405).json({error:'POST vereist.'});
  try {
    const {action,sheet,row,status,note} = req.body || {};
    if (action !== 'update') return res.status(400).json({error:'Ongeldige actie.'});
    if (sheet !== 'MRM-01') return res.status(400).json({error:'Ongeldig werkblad.'});
    const n = Number(row);
    if (!Number.isInteger(n) || n < 2) return res.status(400).json({error:'Ongeldig rijnummer.'});
    if (!allowed.has(String(status))) return res.status(400).json({error:'Ongeldige status.'});
    await setCell(n,'M',String(status));
    if (note !== undefined) await setCell(n,'W',String(note));
    return res.status(200).json({ok:true,row:n,status:String(status)});
  } catch(e) {
    const code = e.code === 'NOT_CONFIGURED' ? 503 : (e.status || 500);
    return res.status(code).json({ok:false,error:e.message});
  }
}
