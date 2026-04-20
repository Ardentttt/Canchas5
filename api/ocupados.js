// api/ocupados.js
// Devuelve turnos ocupados. Borra filas RESERVANDO expiradas (sin rastro en el Sheet).

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

    const response  = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: sheetName + "!A2:M1000"
    });

    const rows   = response.data.values || [];
    const ahora  = Date.now();
    const result = {};
    const aBorrar = []; // índices 0-based (fila 2 del sheet = índice 1)

    for (let i = 0; i < rows.length; i++) {
      const row      = rows[i];
      const rCourtId = String(row[3] || "");
      const rDate    = row[4]  || "";
      const rSlot    = row[5]  || "";
      const rEstado  = row[10] || "";
      const rTs      = row[12] || ""; // columna M: timestamp de inicio de reserva

      // RESERVANDO expirada → borrar, no bloquear el turno
      if (rEstado === "RESERVANDO" && rTs) {
        const edad = ahora - new Date(rTs).getTime();
        if (edad > EXPIRACION_MS) {
          aBorrar.push(i + 1); // i+1 porque fila 1 es header (0-based)
          continue;
        }
      }

      if (rCourtId === String(courtId) && rEstado !== "CANCELADA") {
        const key   = rDate + "|" + rSlot;
        result[key] = rEstado === "CONFIRMADA" ? "confirmed" : "pending";
      }
    }

    // Borrar filas expiradas en background — no bloqueamos la respuesta
    if (aBorrar.length > 0) {
      borrarFilas(sheets, sheetName, aBorrar).catch(console.error);
    }

    return res.status(200).json(result);
  } catch (err) {
    console.error("Error en ocupados:", err);
    return res.status(500).json({ error: err.message });
  }
};

// Borra filas del Sheet de abajo hacia arriba para no desplazar índices
async function borrarFilas(sheets, sheetName, rowIndexes) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
  const hoja = meta.data.sheets.find(function(s) { return s.properties.title === sheetName; });
  if (!hoja) return;
  const sheetId = hoja.properties.sheetId;

  // Ordenar de mayor a menor
  const sorted = rowIndexes.slice().sort(function(a, b) { return b - a; });

  const requests = sorted.map(function(rowIndex) {
    return {
      deleteDimension: {
        range: {
          sheetId,
          dimension: "ROWS",
          startIndex: rowIndex,
          endIndex:   rowIndex + 1
        }
      }
    };
  });

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: GOOGLE_SHEET_ID,
    requestBody: { requests }
  });
  console.log("Filas expiradas borradas:", sorted);
}

async function getSheetName(sheets) {
  try {
    const meta    = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
    const hojas   = meta.data.sheets.map(function(s) { return s.properties.title; });
    const semanas = hojas.filter(function(h) { return h.startsWith("Semana"); });
    if (semanas.length > 0) return semanas[semanas.length - 1];
    if (hojas.includes("Reservas")) return "Reservas";
    return hojas[0];
  } catch (e) {
    return "Reservas";
  }
}

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
