// api/ocupados.js
// Devuelve los turnos ocupados de una cancha desde Google Sheets.
// También limpia reservas RESERVANDO con más de 10 minutos (expiradas).

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
    const sheets     = await getSheetsClient();
    const sheetName  = await getSheetName(sheets);
    const range      = sheetName + "!A2:M1000";

    const response   = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID,
      range
    });

    const rows    = response.data.values || [];
    const ahora   = Date.now();
    const result  = {};
    const updates = []; // filas a marcar como EXPIRADA

    for (let i = 0; i < rows.length; i++) {
      const row      = rows[i];
      const rCourtId = String(row[3] || "");
      const rDate    = row[4]  || "";
      const rSlot    = row[5]  || "";
      const rEstado  = row[10] || "";
      const rTs      = row[12] || ""; // columna M: timestamp ISO de cuando se reservó

      // Limpiar reservas temporales expiradas
      if (rEstado === "RESERVANDO" && rTs) {
        const edad = ahora - new Date(rTs).getTime();
        if (edad > EXPIRACION_MS) {
          updates.push(i + 2); // número de fila en Sheet (1-indexed + header)
          continue; // no la incluimos como ocupada
        }
      }

      if (rCourtId === String(courtId) && rEstado !== "CANCELADA" && rEstado !== "EXPIRADA") {
        const key   = rDate + "|" + rSlot;
        result[key] = rEstado === "CONFIRMADA" ? "confirmed" : "pending";
      }
    }

    // Marcar expiradas en background (no bloqueamos la respuesta)
    if (updates.length > 0) {
      marcarExpiradas(sheets, sheetName, updates).catch(console.error);
    }

    return res.status(200).json(result);
  } catch (err) {
    console.error("Error en ocupados:", err);
    return res.status(500).json({ error: err.message });
  }
};

// Marca filas como EXPIRADA en la columna K
async function marcarExpiradas(sheets, sheetName, rowNums) {
  const data = rowNums.map(function(rowNum) {
    return {
      range: sheetName + "!K" + rowNum,
      values: [["EXPIRADA"]]
    };
  });
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: GOOGLE_SHEET_ID,
    requestBody: {
      valueInputOption: "USER_ENTERED",
      data
    }
  });
}

// Obtiene el nombre de la hoja activa (la más reciente que empiece con "Semana")
async function getSheetName(sheets) {
  try {
    const meta = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
    const hojas = meta.data.sheets.map(function(s) { return s.properties.title; });
    // Busca hojas que empiecen con "Semana", toma la última
    const semanas = hojas.filter(function(h) { return h.startsWith("Semana"); });
    if (semanas.length > 0) return semanas[semanas.length - 1];
    // Fallback a "Reservas" si no hay hojas de semana
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
