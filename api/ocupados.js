// api/ocupados.js
const { google } = require("googleapis");

const GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID;
const GOOGLE_SA_EMAIL = process.env.GOOGLE_SA_EMAIL;
const GOOGLE_SA_KEY   = process.env.GOOGLE_SA_KEY;

const EXPIRACION_MS = 10 * 60 * 1000; // 10 minutos

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  if (req.method === "OPTIONS") return res.status(200).end();

  const { courtId } = req.query;
  if (!courtId) return res.status(400).json({ error: "Falta courtId" });

  try {
    const sheets    = await getSheetsClient();
    const sheetName = await getSheetName(sheets);

    // =============================================================
    // AUTO-LIMPIEZA PASIVA: Borramos las expiradas antes de leer
    // =============================================================
    const responseInicial = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: sheetName + "!A2:M1000"
    });

    const rowsInicial = responseInicial.data.values || [];
    const ahora = Date.now();
    const aBorrar = [];

    for (let i = 0; i < rowsInicial.length; i++) {
      const rEstado = String(rowsInicial[i][10] || "");
      const rTs     = rowsInicial[i][12] || ""; // Columna M (Timestamp)

      if (rEstado === "RESERVANDO" && rTs) {
        const timestampReserva = Number(rTs);
        const tiempoReservaMS = isNaN(timestampReserva) ? new Date(rTs).getTime() : timestampReserva;
        const edad = ahora - tiempoReservaMS;

        if (edad > EXPIRACION_MS) {
          aBorrar.push(i + 1); // Fila indexada (i=0 es fila 2, startIndex=1)
        }
      }
    }

    // Si encontramos celdas viejas, las borramos en lote antes de armar la grilla
    if (aBorrar.length > 0) {
      console.log("Limpieza pasiva: Borrando expiradas detectadas:", aBorrar.length);
      const metaHoja = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
      const hoja = metaHoja.data.sheets.find(s => s.properties.title === sheetName);
      
      if (hoja) {
        const sheetId = hoja.properties.sheetId;
        const sorted = aBorrar.slice().sort((a, b) => b - a);
        const requests = sorted.map(idx => ({
          deleteDimension: { range: { sheetId, dimension: "ROWS", startIndex: idx, endIndex: idx + 1 } }
        }));

        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: GOOGLE_SHEET_ID,
          requestBody: { requests }
        });
      }
    }

    // =============================================================
    // LECTURA FINAL: Traemos los datos limpios para el frontend
    // =============================================================
    const responseFinal = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: sheetName + "!A2:M1000"
    });

    const rows = responseFinal.data.values || [];
    const result = {};

    for (let i = 0; i < rows.length; i++) {
      const row      = rows[i];
      const rCourtId = String(row[3] || "");
      const rDate    = row[4]  || "";
      const rSlot    = row[5]  || "";
      const rEstado  = row[10] || "";

      if (rCourtId === String(courtId) && rEstado !== "CANCELADA") {
        const key = rDate + "|" + rSlot;
        result[key] = rEstado === "CONFIRMADA" ? "confirmed" : "pending";
      }
    }

    return res.status(200).json(result);

  } catch (err) {
    console.error("Error en ocupados con auto-limpieza:", err);
    return res.status(500).json({ error: err.message });
  }
};

async function getSheetName(sheets) {
  try {
    const meta  = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
    const hojas = meta.data.sheets.map(s => s.properties.title);
    const meses = hojas.filter(h => h.startsWith("Mes "));
    if (meses.length > 0) return meses[meses.length - 1];
    if (hojas.includes("Reservas")) return "Reservas";
    return hojas[0];
  } catch (e) {
    return "Reservas";
  }
}

async function getSheetsClient() {
  const key = GOOGLE_SA_KEY.replace(/\\n/g, "\n").trim();
  const auth = new google.auth.GoogleAuth({
    credentials: { client_email: GOOGLE_SA_EMAIL, private_key: key },
    scopes: ["https://www.googleapis.com/auth/spreadsheets"]
  });
  return google.sheets({ version: "v4", auth });
}
