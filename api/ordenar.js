// api/ordenar.js
// Endpoint manual para ordenar una hoja existente por fecha + horario.
// Llamar con GET /api/ordenar?sheet=Mayo%202026
// Si no se pasa ?sheet= usa la hoja del mes actual.

const { google } = require("googleapis");

const GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID;
const GOOGLE_SA_EMAIL = process.env.GOOGLE_SA_EMAIL;
const GOOGLE_SA_KEY   = process.env.GOOGLE_SA_KEY;

const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
               "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  if (req.method !== "GET") return res.status(405).end();

  try {
    const sheets = await getSheetsClient();

    // Si se pasa ?sheet=NombreHoja lo usa, sino usa el mes actual
    let sheetName = req.query.sheet || "";
    if (!sheetName) {
      const ahora = new Date();
      sheetName = MESES[ahora.getMonth()] + " " + ahora.getFullYear();
    }

    const r    = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: sheetName + "!A2:K1000"
    });
    const rows = r.data.values || [];

    if (rows.length === 0) {
      return res.status(200).json({ ok: true, mensaje: "Hoja vacía, nada que ordenar" });
    }

    const sorted = rows.slice().sort(function(a, b) {
      const dA = a[4] || "", dB = b[4] || "";
      if (dA !== dB) return dA < dB ? -1 : 1;
      const sA = a[5] || "", sB = b[5] || "";
      return sA < sB ? -1 : sA > sB ? 1 : 0;
    });

    const padded = sorted.map(function(row) {
      const r = row.slice();
      while (r.length < 11) r.push("");
      return r;
    });

    await sheets.spreadsheets.values.update({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: sheetName + "!A2:K" + (padded.length + 1),
      valueInputOption: "RAW",
      requestBody: { values: padded }
    });

    return res.status(200).json({
      ok: true,
      hoja: sheetName,
      filas: padded.length,
      mensaje: "Hoja ordenada por fecha y horario correctamente"
    });

  } catch (err) {
    console.error("Error ordenar:", err);
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
