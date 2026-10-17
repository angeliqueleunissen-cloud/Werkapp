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

    // Huidige Excelgegevens ophalen
    const used = await getWorkbookUsedRange();
    const values = used.values || [];

    const address = String(used.address || 'A1');
    const match =
      address.match(/![A-Z]+(\d+)/i) ||
      address.match(/^[A-Z]+(\d+)/i);

    const startRow = match ? Number(match[1]) : 1;
    const arrayIndex = n - startRow;

    if (arrayIndex < 0 || arrayIndex >= values.length) {
      return res.status(400).json({
        error: 'De opgegeven Excel-rij bestaat niet.'
      });
    }

    const currentRow = values[arrayIndex] || [];

    // M = Status
    const oldStatus = String(currentRow[12] ?? '').trim();

    // N = Type
    const type = String(currentRow[13] ?? '').trim();

    // W = Opmerkingen
    const oldNote = String(currentRow[22] ?? '').trim();

    /*
     * BELANGRIJKE REGEL:
     *
     * Als Blazen TB = Nee heeft gekozen,
     * staat er "TB niet gemaakt" in Opmerkingen.
     *
     * Zodra Civiel daarna gereed meldt,
     * moet de opdracht naar LASSEN.
     *
     * Dit wordt hier in de backend afgedwongen.
     */
    let finalStatus = String(status);
    let finalNote = String(note ?? '').trim();

    if (
      oldStatus === 'Civiel' &&
      finalStatus === 'Blazen' &&
      oldNote.toLowerCase().includes('tb niet gemaakt')
    ) {
      finalStatus = 'Lassen';
      finalNote = 'Civiel gereed - Blazen was al gereed, TB niet gemaakt - naar Lassen';
    }

    // Extra beveiliging:
    // DAC mag nooit naar Blazen.
    if (
      finalStatus === 'Blazen' &&
      type.toUpperCase() === 'DAC'
    ) {
      finalStatus = 'Lassen';
      finalNote = 'DAC - automatisch naar Lassen';
    }

    // Datum/tijd Nederland
    const timestamp = new Intl.DateTimeFormat('nl-NL', {
      timeZone: 'Europe/Amsterdam',
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      hour12: false
    }).format(new Date());

    if (!finalNote) {
      finalNote =
        `Status gewijzigd: ${oldStatus || '-'} → ${finalStatus}`;
    }

    const logEntry = `${timestamp} - ${finalNote}`;

    // Bestaande opmerkingen behouden
    const newNote = oldNote
      ? `${oldNote}\n${logEntry}`
      : logEntry;

    // Status opslaan
    await setCell(n, 'M', finalStatus);

    // Opmerking toevoegen
    await setCell(n, 'W', newNote);

    return res.status(200).json({
      ok: true,
      row: n,
      oldStatus,
      status: finalStatus,
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
