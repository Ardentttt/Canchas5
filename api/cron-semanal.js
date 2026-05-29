// api/cron-semanal.js
// Ahora hace rotación MENSUAL, no semanal.
// El 1ro de cada mes a las 3am Argentina crea una hoja nueva "Mes YYYY-MM"
// y borra la del mes anterior. Solo mantiene 1 mes activo + 1 de historial (2 total).

const { google } = require("googleapis");

const GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID;
const GOOGLE_SA_EMAIL = process.env.GOOGLE_SA_EMAIL;
const GOOGLE_SA_KEY   = process.env.GOOGLE_SA_KEY;

const HEADERS = [
  "ID", "Fecha Reserva", "Cancha", "Cancha ID",
  "Fecha Turno", "Horario", "Nombre", "Teléfono",
  "Precio Total", "Seña 50%", "Estado", "Notas", "Timestamp"
];

module.exports = async function handler(req, res) {
  if (req.method !== "GET") return res.status(405).end();

  try {
    const sheets = await getSheetsClient();
    const meta   = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
    const hojas  = meta.data.sheets;
    const titles = hojas.map(function(s) { return s.properties.title; });

    // Nombre de la hoja para el mes actual
    const ahora    = new Date();
    const year     = ahora.getFullYear();
    const mes      = String(ahora.getMonth() + 1).padStart(2, "0");
    const nombreNuevo = "Mes " + year + "-" + mes;

    // Si ya existe la hoja de este mes, no hacer nada
    if (titles.includes(nombreNuevo)) {
      return res.status(200).json({ ok: true, mensaje: "Hoja del mes ya existe: " + nombreNuevo });
    }

    // Crear nueva hoja del mes
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: GOOGLE_SHEET_ID,
      requestBody: {
        requests: [{
          addSheet: {
            properties: {
              title: nombreNuevo,
              gridProperties: { rowCount: 1000, columnCount: 13 }
            }
          }
        }]
      }
    });

    // Agregar headers
    await sheets.spreadsheets.values.update({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: nombreNuevo + "!A1:M1",
      valueInputOption: "RAW",
      requestBody: { values: [HEADERS] }
    });

    // Formatear header
    const metaNew  = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
    const hojaNew  = metaNew.data.sheets.find(function(s) { return s.properties.title === nombreNuevo; });
    const sheetIdNew = hojaNew.properties.sheetId;

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: GOOGLE_SHEET_ID,
      requestBody: {
        requests: [
          {
            repeatCell: {
              range: { sheetId: sheetIdNew, startRowIndex: 0, endRowIndex: 1 },
              cell: {
                userEnteredFormat: {
                  backgroundColor: { red: 0.067, green: 0.122, blue: 0.082 },
                  textFormat: { bold: true, foregroundColor: { red: 0.0, green: 0.91, blue: 0.478 } }
                }
              },
              fields: "userEnteredFormat(backgroundColor,textFormat)"
            }
          },
          {
            updateSheetProperties: {
              properties: { sheetId: sheetIdNew, gridProperties: { frozenRowCount: 1 } },
              fields: "gridProperties.frozenRowCount"
            }
          }
        ]
      }
    });

    console.log("Hoja creada:", nombreNuevo);

    // Borrar hojas viejas — conservar solo el mes actual y el anterior
    const metaFinal  = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
    const hojasFinal = metaFinal.data.sheets;
    const meses = hojasFinal
      .filter(function(s) { return s.properties.title.startsWith("Mes "); })
      .sort(function(a, b) { return a.properties.title.localeCompare(b.properties.title); });

    // También borrar hojas "Semana X" viejas si quedaron
    const semanas = hojasFinal.filter(function(s) { return s.properties.title.startsWith("Semana"); });

    const aBorrar = [];

    // Conservar solo los últimos 2 meses
    if (meses.length > 2) {
      const viejas = meses.slice(0, meses.length - 2);
      viejas.forEach(function(h) { aBorrar.push(h.properties.sheetId); });
    }

    // Borrar todas las hojas "Semana X" viejas
    semanas.forEach(function(h) { aBorrar.push(h.properties.sheetId); });

    if (aBorrar.length > 0) {
      const requests = aBorrar.map(function(sheetId) {
        return { deleteSheet: { sheetId } };
      });
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: GOOGLE_SHEET_ID,
        requestBody: { requests }
      });
      console.log("Hojas borradas:", aBorrar.length);
    }

    return res.status(200).json({
      ok: true,
      creada: nombreNuevo,
      borradas: aBorrar.length
    });

  } catch (err) {
    console.error("Error en cron-mensual:", err);
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
