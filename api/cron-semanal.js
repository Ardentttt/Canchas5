// api/cron-semanal.js — rotación mensual
// Vercel cron: "0 6 1 * *" (3am Argentina el día 1 de cada mes)
const { google } = require("googleapis");

const GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID;
const GOOGLE_SA_EMAIL = process.env.GOOGLE_SA_EMAIL;
const GOOGLE_SA_KEY   = process.env.GOOGLE_SA_KEY;

const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
               "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

const HEADERS = ["ID","Hora","Cancha","Cancha ID","Fecha Turno",
                 "Horario","Nombre","Teléfono","Precio Total","Estado","Notas"];

module.exports = async function handler(req, res) {
  if (req.method !== "GET") return res.status(405).end();

  try {
    const sheets = await getSheetsClient();
    const ahora  = new Date();
    const nombre = MESES[ahora.getMonth()] + " " + ahora.getFullYear();

    const meta  = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
    const hojas = meta.data.sheets.map(function(s) { return s.properties.title; });

    if (hojas.includes(nombre)) {
      return res.status(200).json({ ok: true, mensaje: "Ya existe: " + nombre });
    }

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: GOOGLE_SHEET_ID,
      requestBody: { requests: [{ addSheet: {
        properties: { title: nombre, index: hojas.length,
                      gridProperties: { rowCount: 1000, columnCount: 11 } }
      }}]}
    });

    await sheets.spreadsheets.values.update({
      spreadsheetId: GOOGLE_SHEET_ID, range: nombre + "!A1:K1",
      valueInputOption: "RAW", requestBody: { values: [HEADERS] }
    });

    const metaNew = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
    const hojaNew = metaNew.data.sheets.find(function(s) { return s.properties.title === nombre; });
    const sid     = hojaNew.properties.sheetId;

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: GOOGLE_SHEET_ID,
      requestBody: { requests: [
        { repeatCell: {
            range: { sheetId: sid, startRowIndex: 0, endRowIndex: 1 },
            cell: { userEnteredFormat: {
              backgroundColor: { red: 0.067, green: 0.122, blue: 0.082 },
              textFormat: { bold: true, foregroundColor: { red: 0.0, green: 0.91, blue: 0.478 } }
            }},
            fields: "userEnteredFormat(backgroundColor,textFormat)"
        }},
        { updateSheetProperties: {
            properties: { sheetId: sid, gridProperties: { frozenRowCount: 1 } },
            fields: "gridProperties.frozenRowCount"
        }}
      ]}
    });

    console.log("Hoja mensual creada:", nombre);
    return res.status(200).json({ ok: true, creada: nombre });

  } catch (err) {
    console.error("Error cron:", err);
    return res.status(500).json({ error: err.message });
  }
};

async function getSheetsClient() {
  const key = GOOGLE_SA_KEY.replace(/\\n/g, "\n").trim();
  const auth = new google.auth.GoogleAuth({
    credentials: { client_email: GOOGLE_SA_EMAIL, private_key: key },
    scopes: ["https://www.googleapis.com/auth/spreadsheets"]
  });
  return google.sheets({ version: "v4", auth });
}
