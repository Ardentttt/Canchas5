// api/ocupados.js
// Solo lee CONFIRMADA y PENDIENTE del Sheet.
// RESERVANDO ya no existe — el turno bloqueado se maneja solo con la preferencia de MP activa.

const { google } = require("googleapis");

const GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID;
const GOOGLE_SA_EMAIL = process.env.GOOGLE_SA_EMAIL;
const GOOGLE_SA_KEY   = process.env.GOOGLE_SA_KEY;

const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
               "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  if (req.method === "OPTIONS") return res.status(200).end();

  const { courtId } = req.query;
  if (!courtId) return res.status(400).json({ error: "Falta courtId" });

  try {
    const sheets    = await getSheetsClient();
    const sheetName = await getOrCreateSheetName(sheets);

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: sheetName + "!A2:K1000"
    });

    const rows   = response.data.values || [];
    const result = {};

    for (const row of rows) {
      const rCourt = String(row[3] || "");
      const rDate  = row[4] || "";
      const rSlot  = row[5] || "";
      const rEst   = row[9] || "";

      // Solo mostrar como ocupado lo que está CONFIRMADA o PENDIENTE
      if (rCourt === String(courtId) &&
          (rEst === "CONFIRMADA" || rEst === "PENDIENTE")) {
        result[rDate + "|" + rSlot] = rEst === "CONFIRMADA" ? "confirmed" : "pending";
      }
    }

    return res.status(200).json(result);
  } catch (err) {
    console.error("Error en ocupados:", err);
    return res.status(500).json({ error: err.message });
  }
};

async function getOrCreateSheetName(sheets) {
  const ahora  = new Date();
  const nombre = MESES[ahora.getMonth()] + " " + ahora.getFullYear();
  const meta   = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
  const hojas  = meta.data.sheets.map(function(s) { return s.properties.title; });
  if (hojas.includes(nombre)) return nombre;
  // Si no existe la hoja del mes, devolver la última disponible
  if (hojas.length > 0) return hojas[hojas.length - 1];
  return "Reservas";
}

async function getSheetsClient() {
  const key = GOOGLE_SA_KEY.replace(/\\n/g, "\n").replace(/\r\n/g, "\n").trim();
  const auth = new google.auth.GoogleAuth({
    credentials: { client_email: GOOGLE_SA_EMAIL, private_key: key },
    scopes: ["https://www.googleapis.com/auth/spreadsheets"]
  });
  return google.sheets({ version: "v4", auth });
}
