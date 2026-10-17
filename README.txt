MRM-01 Werkapp

Frontend + Vercel serverless backend.
Monteurs hoeven niet in te loggen bij Microsoft.

Backend endpoints:
GET  /api/health
GET  /api/list?worker=Civiel
POST /api/update

Excel configuration:
EXCEL_FILENAME = MRM-01 Openstaande schouw .xlsx
EXCEL_SHEET = MRM-01

Microsoft setup is intentionally not included in source code. Secrets/tokens belong in Vercel Environment Variables.
