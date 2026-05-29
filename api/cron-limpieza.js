// api/cron-limpieza.js
const { google } = require("googleapis");

const GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID;
const GOOGLE_SA_EMAIL = process.env.GOOGLE_SA_EMAIL;
const GOOGLE_SA_KEY   = process.env.GOOGLE_SA_KEY;

const EXPIRACION_MS = 10 * 60 * 1000; // 10 minutos

module.exports = async function handler(req, res) {
  // Solo permitir solicitudes GET (Vercel Crons utiliza GET)
  if (req.method !== "GET") return res.status(405).end();

  try {
    const sheets    = await getSheetsClient();
    const sheetName = await getSheetName(sheets);

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: sheetName + "!A2:M1000"
    });
    
    const rows  = response.data.values || [];
    const ahora = Date.now();
    const aBorrar = [];

    for (let i = 0; i < rows.length; i++) {
      const rEstado = String(rows[i][10] || "");
      const rTs     = rows[i][12] || ""; // Columna M (Timestamp)

      if (rEstado === "RESERVANDO" && rTs) {
        const timestampReserva = Number(rTs);
        
        // Soporte de compatibilidad si queda algún formato ISO viejo ("2026-05-29...")
        const tiempoReservaMS = isNaN(timestampReserva) ? new Date(rTs).getTime() : timestampReserva;

        const edad = ahora - tiempoReservaMS;

        if (edad > EXPIRACION_MS) {
          aBorrar.push(i + 1); // Fila 2 de la hoja corresponde a i=0, por ende startIndex=1
        }
      }
    }

    if (aBorrar.length === 0) {
      return res.status(200).json({ ok: true, mensaje: "No había reservas expiradas." });
    }

    const meta    = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
    const hoja    = meta.data.sheets.find(s => s.properties.title === sheetName);
    if (!hoja) return res.status(404).json({ error: "Hoja no encontrada" });
    const sheetId = hoja.properties.sheetId;

    // Ordenar de mayor a menor para borrar de abajo hacia arriba y no desplazar las filas remanentes
    const sorted   = aBorrar.slice().sort((a, b) => b - a);
    const requests = sorted.map(idx => ({
      deleteDimension: { range: { sheetId, dimension: "ROWS", startIndex: idx, endIndex: idx + 1 } }
    }));

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: GOOGLE_SHEET_ID,
      requestBody: { requests }
    });

    console.log("Cron Limpieza: Expiradas borradas:", sorted.length);
    return res.status(200).json({ ok: true, borradas: sorted.length });

  } catch (err) {
    console.error("Error en cron-limpieza:", err);
    return res.status(500).json({ error: err.message });
  }
};

async function getSheetName(sheets) {
  const meta  = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
  const hojas = meta.data.sheets.map(s => s.properties.title);
  const meses = hojas.filter(h => h.startsWith("Mes "));
  if (meses.length > 0) return meses[meses.length - 1];
  return hojas[0];
}

async function getSheetsClient() {
  const key = GOOGLE_SA_KEY.replace(/\\n/g, "\n").trim();
  const auth = new google.auth.GoogleAuth({
    credentials: { client_email: GOOGLE_SA_EMAIL, private_key: key },
    scopes: ["https://www.googleapis.com/auth/spreadsheets"]
  });
  return google.sheets({ version: "v4", auth });
}
