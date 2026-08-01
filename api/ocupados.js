// api/ocupados.js
const { google } = require("googleapis");

const GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID;
const GOOGLE_SA_EMAIL = process.env.GOOGLE_SA_EMAIL;
const GOOGLE_SA_KEY   = process.env.GOOGLE_SA_KEY;

const EXPIRACION_MS = 10 * 60 * 1000;

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

    const rows    = response.data.values || [];
    const ahora   = Date.now();
    const result  = {};
    const aBorrar = [];

    for (let i = 0; i < rows.length; i++) {
      const row    = rows[i];
      const rCourt = String(row[3] || "");
      const rDate  = row[4] || "";
      const rSlot  = row[5] || "";
      const rEst   = row[9] || "";  // col J = Estado
      const notas  = row[10] || ""; // col K = Notas (contiene TS)

      if (rEst === "RESERVANDO") {
        const match = notas.match(/TS:(.+)$/);
        if (match) {
          const edad = ahora - new Date(match[1]).getTime();
          if (edad > EXPIRACION_MS) {
            aBorrar.push(i + 1);
            continue;
          }
        }
        if (rCourt === String(courtId)) {
          result[rDate + "|" + rSlot] = "pending";
        }
        continue;
      }

      if (rCourt === String(courtId) && rEst !== "CANCELADA") {
        result[rDate + "|" + rSlot] = rEst === "CONFIRMADA" ? "confirmed" : "pending";
      }
    }

    if (aBorrar.length > 0) borrarFilas(sheets, sheetName, aBorrar).catch(console.error);

    return res.status(200).json(result);
  } catch (err) {
    console.error("Error en ocupados:", err);
    return res.status(500).json({ error: err.message });
  }
};

async function borrarFilas(sheets, sheetName, indexes) {
  const meta    = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
  const hoja    = meta.data.sheets.find(function(s) { return s.properties.title === sheetName; });
  if (!hoja) return;
  const sheetId = hoja.properties.sheetId;
  const sorted  = indexes.slice().sort(function(a, b) { return b - a; });
  const reqs    = sorted.map(function(idx) {
    return { deleteDimension: { range: { sheetId, dimension: "ROWS", startIndex: idx, endIndex: idx + 1 } } };
  });
  await sheets.spreadsheets.batchUpdate({ spreadsheetId: GOOGLE_SHEET_ID, requestBody: { requests: reqs } });
  console.log("Expiradas borradas:", sorted.length);
}

async function getOrCreateSheetName(sheets) {
  const ahora  = new Date();
  const nombre = MESES[ahora.getMonth()] + " " + ahora.getFullYear();

  const meta  = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
  const hojas = meta.data.sheets.map(function(s) { return s.properties.title; });

  if (hojas.includes(nombre)) return nombre;

  const HEADERS = ["ID","Hora","Cancha","Cancha ID","Fecha Turno",
                   "Horario","Nombre","Teléfono","Precio Total","Estado","Notas"];

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: GOOGLE_SHEET_ID,
    requestBody: { requests: [{ addSheet: {
      properties: { title: nombre, gridProperties: { rowCount: 1000, columnCount: 11 } }
    }}]}
  });
  await sheets.spreadsheets.values.update({
    spreadsheetId: GOOGLE_SHEET_ID, range: nombre + "!A1:K1",
    valueInputOption: "RAW", requestBody: { values: [HEADERS] }
  });

  const metaNew = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
  const hojaNew = metaNew.data.sheets.find(function(s) { return s.properties.title === nombre; });
  const sid     = hojaNew.properties.sheetId;

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: GOOGLE_SHEET_ID,
    requestBody: { requests: [
      { repeatCell: {
          range: { sheetId: sid, startRowIndex: 0, endRowIndex: 1 },
          cell: { userEnteredFormat: {
            backgroundColor: { red: 0.067, green: 0.122, blue: 0.082 },
            textFormat: { bold: true, foregroundColor: { red: 0.0, green: 0.91, blue: 0.478 } }
          }},
          fields: "userEnteredFormat(backgroundColor,textFormat)"
      }},
      { updateSheetProperties: {
          properties: { sheetId: sid, gridProperties: { frozenRowCount: 1 } },
          fields: "gridProperties.frozenRowCount"
      }}
    ]}
  });

  console.log("Hoja creada:", nombre);
  return nombre;
}

async function getSheetsClient() {
  const key = GOOGLE_SA_KEY.replace(/\\n/g, "\n").replace(/\r\n/g, "\n").trim();
  const auth = new google.auth.GoogleAuth({
    credentials: { client_email: GOOGLE_SA_EMAIL, private_key: key },
    scopes: ["https://www.googleapis.com/auth/spreadsheets"]
  });
  return google.sheets({ version: "v4", auth });
}
