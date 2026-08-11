// api/ocupados.js
// Lee hoja mensual (CONFIRMADA/PENDIENTE) + hoja Temp (bloqueos activos).

const { google } = require("googleapis");

const GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID;
const GOOGLE_SA_EMAIL = process.env.GOOGLE_SA_EMAIL;
const GOOGLE_SA_KEY   = process.env.GOOGLE_SA_KEY;

const EXPIRACION_MS = 10 * 60 * 1000;
const TEMP_SHEET    = "Temp";

const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
               "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  if (req.method === "OPTIONS") return res.status(200).end();

  const { courtId, date } = req.query;
  if (!courtId) return res.status(400).json({ error: "Falta courtId" });

  try {
    const sheets = await getSheetsClient();
    const result = {};
    const ahora  = Date.now();

    // 1. Leer hoja mensual del mes del turno que se está consultando
    // Si no hay date en query, usar mes actual
    const fechaRef  = date || new Date().toISOString().slice(0, 10);
    const sheetName = await getSheetName(sheets, fechaRef);

    if (sheetName) {
      const rMes = await sheets.spreadsheets.values.get({
        spreadsheetId: GOOGLE_SHEET_ID, range: sheetName + "!A2:K1000"
      });
      for (const row of rMes.data.values || []) {
        if (String(row[3]||"") !== String(courtId)) continue;
        const est = row[9] || "";
        if (est === "CONFIRMADA" || est === "PENDIENTE") {
          result[(row[4]||"") + "|" + (row[5]||"")] = est === "CONFIRMADA" ? "confirmed" : "pending";
        }
      }
    }

    // 2. Leer Temp (bloqueos activos de 10 min)
    try {
      const rTemp = await sheets.spreadsheets.values.get({
        spreadsheetId: GOOGLE_SHEET_ID, range: TEMP_SHEET + "!A2:E1000"
      });
      for (const row of rTemp.data.values || []) {
        if (String(row[1]||"") !== String(courtId)) continue;
        const ts   = row[4] || "";
        const edad = ahora - new Date(ts).getTime();
        if (edad <= EXPIRACION_MS) {
          const key = (row[2]||"") + "|" + (row[3]||"");
          if (!result[key]) result[key] = "pending"; // solo si no está ya confirmado
        }
      }
    } catch (e) {
      // Temp puede no existir todavía, ignorar
    }

    return res.status(200).json(result);
  } catch (err) {
    console.error("Error en ocupados:", err);
    return res.status(500).json({ error: err.message });
  }
};

async function getSheetName(sheets, fechaRef) {
  try {
    const base   = new Date(fechaRef + "T00:00:00");
    const nombre = MESES[base.getMonth()] + " " + base.getFullYear();
    const meta   = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
    const hojas  = meta.data.sheets.map(function(s) { return s.properties.title; });
    if (hojas.includes(nombre)) return nombre;
    return null; // hoja del mes no existe todavía = no hay ocupados
  } catch (e) {
    return null;
  }
}

async function getSheetsClient() {
  const key = GOOGLE_SA_KEY.replace(/\\n/g, "\n").replace(/\r\n/g, "\n").trim();
  const auth = new google.auth.GoogleAuth({
    credentials: { client_email: GOOGLE_SA_EMAIL, private_key: key },
    scopes: ["https://www.googleapis.com/auth/spreadsheets"]
  });
  return google.sheets({ version: "v4", auth });
}
