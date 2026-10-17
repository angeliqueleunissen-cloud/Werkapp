import { getWorkbookUsedRange, setCell } from './_graph.js';

const stages = {
  Civiel: 'Civiel',
  Blazen: 'Blazen',
  Lasser: 'Lassen',
  Hasser: 'Hassen',
  Schouw: 'Schouw',
  Schouwer: 'Schouw'
};

export default async function handler(req, res) {
  try {
    const worker = String(req.query?.worker || '');
    const wanted = stages[worker];

    if (!wanted) {
      return res.status(400).json({ error: 'Onbekende medewerker.' });
    }

    const used = await getWorkbookUsedRange();
    const values = used.values || [];

    if (!values.length) {
      return res.status(200).json({ rows: [] });
    }

    const address = String(used.address || 'A1');
    const m =
      address.match(/![A-Z]+(\d+)/i) ||
      address.match(/^[A-Z]+(\d+)/i);

    const startRow = m ? Number(m[1]) : 1;

    const headers = values[0].map(
      (v, i) => String(v ?? '').trim() || `Kolom ${i + 1}`
    );

    const rows = [];

    for (let i = 1; i < values.length; i++) {
      const row = values[i];

      const status = String(row[12] ?? '').trim(); // M = Status
      const type = String(row[13] ?? '').trim();   // N = Type

      /*
       * DAC mag NOOIT bij Blazen staan.
       *
       * Staat er toch:
       * Status = Blazen
       * Type   = DAC
       *
       * dan wordt de status automatisch Lassen.
       */
      if (
        status === 'Blazen' &&
        type.toUpperCase() === 'DAC'
      ) {
        const excelRow = startRow + i;

        await setCell(excelRow, 'M', 'Lassen');

        // Als de gebruiker Lasser is,
        // tonen we hem meteen in de Lassen-lijst.
        if (wanted === 'Lassen') {
          const obj = { _row: excelRow };

          headers.forEach((h, j) => {
            obj[h] = row[j] ?? '';
          });

          obj['Status'] = 'Lassen';
          obj['Type'] = row[13] ?? '';
          obj['Opmerkingen'] = row[22] ?? '';

          rows.push(obj);
        }

        continue;
      }

      if (status !== wanted) continue;

      const obj = { _row: startRow + i };

      headers.forEach((h, j) => {
        obj[h] = row[j] ?? '';
      });

      obj['Status'] = row[12] ?? '';
      obj['Type'] = row[13] ?? '';
      obj['Opmerkingen'] = row[22] ?? '';

      rows.push(obj);
    }

    res.setHeader('Cache-Control', 'no-store');

    return res.status(200).json({ rows });

  } catch (e) {
    const code =
      e.code === 'NOT_CONFIGURED'
        ? 503
        : (e.status || 500);

    return res.status(code).json({
      error: e.message
    });
  }
}
