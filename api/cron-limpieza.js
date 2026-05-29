// api/cron-semanal.js
const { google } = require("googleapis");

const GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID;
const GOOGLE_SA_EMAIL = process.env.GOOGLE_SA_EMAIL;
const GOOGLE_SA_KEY   = process.env.GOOGLE_SA_KEY;

const EXPIRACION_MS = 10 * 60 * 1000; // 10 minutos
const HEADERS = [
  "ID", "Fecha Reserva", "Cancha", "Cancha ID",
  "Fecha Turno", "Horario", "Nombre", "Teléfono",
  "Precio Total", "Seña 50%", "Estado", "Notas", "Timestamp"
];

module.exports = async function handler(req, res) {
  if (req.method !== "GET") return res.status(405).end();

  try {
    const sheets    = await getSheetsClient();
    let sheetName   = await getSheetName(sheets);
    const ahora     = new Date();
    const ahoraMS   = Date.now();

    // ==========================================
    // TAREA 1: ROTACIÓN MENSUAL (Solo si es el día 1 y la hoja no existe)
    // ==========================================
    const year        = ahora.getFullYear();
    const mes         = String(ahora.getMonth() + 1).padStart(2, "0");
    const nombreNuevo = "Mes " + year + "-" + mes;

    const metaInicial = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
    const titles      = metaInicial.data.sheets.map(s => s.properties.title);

    if (!titles.includes(nombreNuevo) && ahora.getDate() === 1) {
      console.log("Es día 1 de mes. Creando nueva hoja:", nombreNuevo);
      
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: GOOGLE_SHEET_ID,
        requestBody: {
          requests: [{
            addSheet: { properties: { title: nombreNuevo, gridProperties: { rowCount: 1000, columnCount: 13 } } }
          }]
        }
      });

      await sheets.spreadsheets.values.update({
        spreadsheetId: GOOGLE_SHEET_ID,
        range: nombreNuevo + "!A1:M1",
        valueInputOption: "RAW",
        requestBody: { values: [HEADERS] }
      });

      sheetName = nombreNuevo; // Cambiamos el puntero a la nueva hoja
    }

    // ==========================================
    // TAREA 2: LIMPIEZA DE RESERVAS EXPIRADAS
    // ==========================================
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: sheetName + "!A2:M1000"
    });
    const rows = response.data.values || [];
    const aBorrar = [];

    for (let i = 0; i < rows.length; i++) {
      const rEstado = String(rows[i][10] || "");
      const rTs     = rows[i][12] || ""; // Columna M

      if (rEstado === "RESERVANDO" && rTs) {
        const timestampReserva = Number(rTs);
        const tiempoReservaMS = isNaN(timestampReserva) ? new Date(rTs).getTime() : timestampReserva;
        const edad = ahoraMS - tiempoReservaMS;

        if (edad > EXPIRACION_MS) {
          aBorrar.push(i + 1);
        }
      }
    }

    let borradasCount = 0;
    if (aBorrar.length > 0) {
      const metaFinal = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
      const hoja = metaFinal.data.sheets.find(s => s.properties.title === sheetName);
      
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
        borradasCount = sorted.length;
        console.log("Cron: Se borraron", borradasCount, "filas expiradas.");
      }
    }

    return res.status(200).json({
      ok: true,
      hojaActual: sheetName,
      expiradasBorradas: borradasCount
    });

  } catch (err) {
    console.error("Error en el cron unificado:", err);
    return res.status(500).json({ error: err.message });
  }
};

async function getSheetName(sheets) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
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
