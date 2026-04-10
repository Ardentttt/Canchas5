// api/ocupados.js
// Devuelve los turnos ocupados de una cancha desde Google Sheets.

const { google } = require("googleapis");

const GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID;
const GOOGLE_SA_EMAIL = process.env.GOOGLE_SA_EMAIL;
const GOOGLE_SA_KEY   = process.env.GOOGLE_SA_KEY;

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  if (req.method === "OPTIONS") return res.status(200).end();

  const { courtId } = req.query;
  if (!courtId) return res.status(400).json({ error: "Falta courtId" });

  try {
    const sheets   = await getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: "Reservas!A2:L1000"
    });

    const rows   = response.data.values || [];
    const result = {};

    for (const row of rows) {
      const rCourtId = String(row[3] || "");
      const rDate    = row[4]  || "";
      const rSlot    = row[5]  || "";
      const rEstado  = row[10] || "";

      if (rCourtId === String(courtId) && rEstado !== "CANCELADA") {
        const key    = rDate + "|" + rSlot;
        result[key]  = rEstado === "CONFIRMADA" ? "confirmed" : "pending";
      }
    }

    return res.status(200).json(result);
  } catch (err) {
    console.error("Error en ocupados:", err);
    return res.status(500).json({ error: err.message });
  }
};

async function getSheetsClient() {
  const key = GOOGLE_SA_KEY
    .replace(/\\n/g, "\n")
    .replace(/\r\n/g, "\n")
    .trim();
  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: GOOGLE_SA_EMAIL,
      private_key: key
    },
    scopes: ["https://www.googleapis.com/auth/spreadsheets"]
  });
  return google.sheets({ version: "v4", auth });
}
