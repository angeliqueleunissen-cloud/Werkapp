import { getWorkbookUsedRange, setCell } from './_graph.js';

const allowed = new Set([
  'Civiel',
  'Blazen',
  'Lassen',
  'Hassen',
  'Gereed',
  'Probleem',
  'Schouw'
]);

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'POST vereist.' });
  }

  try {
    const { action, sheet, row, status, note } = req.body || {};

    if (action !== 'update') {
      return res.status(400).json({ error: 'Ongeldige actie.' });
    }

    if (sheet !== 'MRM-01') {
      return res.status(400).json({ error: 'Ongeldig werkblad.' });
    }

    const n = Number(row);

    if (!Number.isInteger(n) || n < 2) {
      return res.status(400).json({ error: 'Ongeldig rijnummer.' });
    }

    if (!allowed.has(String(status))) {
      return res.status(400).json({ error: 'Ongeldige status.' });
    }

    // Huidige gegevens uit Excel ophalen
    const used = await getWorkbookUsedRange();
    const values = used.values || [];

    // Bepalen vanaf welke Excel-rij het gebruikte bereik begint
    const address = String(used.address || 'A1');
    const match =
      address.match(/![A-Z]+(\d+)/i) ||
      address.match(/^[A-Z]+(\d+)/i);

    const startRow = match ? Number(match[1]) : 1;

    // Rij in het gebruikte bereik bepalen
    const arrayIndex = n - startRow;

    if (arrayIndex < 0 || arrayIndex >= values.length) {
      return res.status(400).json({
        error: 'De opgegeven Excel-rij bestaat niet in het gebruikte bereik.'
      });
    }

    const currentRow = values[arrayIndex] || [];

    // M = Status
    const oldStatus = String(currentRow[12] ?? '').trim();

    // W = Opmerkingen
    const oldNote = String(currentRow[22] ?? '').trim();

    // Nederlandse datum/tijd
    const timestamp = new Intl.DateTimeFormat('nl-NL', {
      timeZone: 'Europe/Amsterdam',
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      hour12: false
    }).format(new Date());

    // Melding maken
    let logText = String(note ?? '').trim();

    if (!logText) {
      logText = `Status gewijzigd: ${oldStatus || '-'} → ${String(status)}`;
    }

    const logEntry = `${timestamp} - ${logText}`;

    // Bestaande opmerkingen behouden en nieuwe melding eronder zetten
    const newNote = oldNote
      ? `${oldNote}\n${logEntry}`
      : logEntry;

    // Eerst status wijzigen
    await setCell(n, 'M', String(status));

    // Daarna de statuswijziging toevoegen aan Opmerkingen
    await setCell(n, 'W', newNote);

    return res.status(200).json({
      ok: true,
      row: n,
      oldStatus,
      status: String(status),
      logged: logEntry
    });

  } catch (e) {
    const code =
      e.code === 'NOT_CONFIGURED'
        ? 503
        : (e.status || 500);

    return res.status(code).json({
      ok: false,
      error: e.message
    });
  }
}
