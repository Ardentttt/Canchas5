// api/webhook.js
// Escribe la fila en el Sheet SOLO cuando el pago está aprobado.
// Usa pago.metadata para obtener los datos (no depende de nada en el Sheet).

const { MercadoPagoConfig, Payment } = require("mercadopago");
const { google } = require("googleapis");

const MP_ACCESS_TOKEN = process.env.MP_ACCESS_TOKEN;
const GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID;
const GOOGLE_SA_EMAIL = process.env.GOOGLE_SA_EMAIL;
const GOOGLE_SA_KEY   = process.env.GOOGLE_SA_KEY;

const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
               "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

const HEADERS = ["ID","Hora","Cancha","Cancha ID","Fecha Turno",
                 "Horario","Nombre","Teléfono","Precio Total","Estado","Notas"];

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method === "GET")     return res.status(200).send("OK");
  if (req.method !== "POST")    return res.status(405).end();

  try {
    const { data } = req.body;
    if (!data?.id) return res.status(200).json({ received: true });

    const client  = new MercadoPagoConfig({ accessToken: MP_ACCESS_TOKEN });
    const payment = new Payment(client);
    const pago    = await payment.get({ id: data.id });

    const estado = pago.status;
    const extRef = pago.external_reference || "";
    const meta   = pago.metadata || {};

    console.log("Webhook:", data.id, "estado:", estado);

    if (estado === "approved") {
      const sheets    = await getSheetsClient();
      const sheetName = await getOrCreateSheetName(sheets, meta.sheet_name);

      // Verificar que no esté ya registrada (evitar duplicados por webhook repetido)
      const yaExiste = await buscarFila(sheets, sheetName, extRef);
      if (yaExiste) {
        console.log("Ya registrada, ignorando duplicado");
        return res.status(200).json({ received: true });
      }

      const reservaId = "R" + data.id.toString().slice(-7);

      // Escribir fila directamente como CONFIRMADA
      await sheets.spreadsheets.values.append({
        spreadsheetId: GOOGLE_SHEET_ID,
        range: sheetName + "!A:K",
        valueInputOption: "RAW",
        requestBody: { values: [[
          reservaId,
          meta.hora_ar || "",
          meta.court_name || "",
          meta.court_id || "",
          meta.date || "",
          meta.slot || "",
          meta.name || "",
          meta.phone || "",
          meta.full_price || "",
          "CONFIRMADA",
          "Pago MP #" + data.id
        ]]}
      });

      console.log("Fila CONFIRMADA escrita en:", sheetName);

      // Ordenar la hoja por fecha + horario
      await ordenarHoja(sheets, sheetName);

    } else if (estado === "rejected" || estado === "cancelled") {
      // No hay nada en el Sheet que borrar — simplemente ignorar
      console.log("Pago rechazado/cancelado, nada que hacer en Sheet");
    } else {
      console.log("Estado pendiente:", estado, "— esperando aprobación");
    }

    return res.status(200).json({ received: true });
  } catch (err) {
    console.error("Error webhook:", err);
    return res.status(500).json({ error: err.message });
  }
};

// Busca si ya existe una fila con ese extRef en Notas para evitar duplicados
async function buscarFila(sheets, sheetName, extRef) {
  const r    = await sheets.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_ID, range: sheetName + "!A2:K1000"
  });
  const rows = r.data.values || [];
  // Extraemos el courtId|date|slot del extRef para comparar
  const parts = extRef.split("|");
  const courtId = parts[0] || "";
  const date    = parts[1] || "";
  const slot    = parts[2] || "";
  for (const row of rows) {
    if (String(row[3]||"") === courtId &&
        String(row[4]||"") === date &&
        String(row[5]||"") === slot &&
        String(row[9]||"") === "CONFIRMADA") {
      return true;
    }
  }
  return false;
}

// Ordena la hoja por Fecha Turno (col E) + Horario (col F)
async function ordenarHoja(sheets, sheetName) {
  const r    = await sheets.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_ID, range: sheetName + "!A2:K1000"
  });
  const rows = r.data.values || [];
  if (rows.length === 0) return;

  const sorted = rows.slice().sort(function(a, b) {
    const dA = a[4] || "", dB = b[4] || "";
    if (dA !== dB) return dA < dB ? -1 : 1;
    const sA = a[5] || "", sB = b[5] || "";
    return sA < sB ? -1 : sA > sB ? 1 : 0;
  });

  const padded = sorted.map(function(row) {
    const r = row.slice();
    while (r.length < 11) r.push("");
    return r;
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: sheetName + "!A2:K" + (padded.length + 1),
    valueInputOption: "RAW",
    requestBody: { values: padded }
  });
  console.log("Hoja ordenada:", padded.length, "filas");
}

async function getOrCreateSheetName(sheets, metaSheetName) {
  // Intentar usar el sheetName guardado en metadata primero
  // MP convierte las claves a snake_case, por eso meta.sheet_name
  const ahora  = new Date();
  const nombre = metaSheetName || (MESES[ahora.getMonth()] + " " + ahora.getFullYear());

  const meta  = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
  const hojas = meta.data.sheets.map(function(s) { return s.properties.title; });
  if (hojas.includes(nombre)) return nombre;

  // Si no existe, crear
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
      { repeatCell: { range: { sheetId: sid, startRowIndex: 0, endRowIndex: 1 },
          cell: { userEnteredFormat: {
            backgroundColor: { red: 0.067, green: 0.122, blue: 0.082 },
            textFormat: { bold: true, foregroundColor: { red: 0.0, green: 0.91, blue: 0.478 } }
          }}, fields: "userEnteredFormat(backgroundColor,textFormat)" }},
      { updateSheetProperties: { properties: { sheetId: sid, gridProperties: { frozenRowCount: 1 } },
          fields: "gridProperties.frozenRowCount" }}
    ]}
  });
  console.log("Hoja creada:", nombre);
  return nombre;
}

async function getSheetsClient() {
  const key = GOOGLE_SA_KEY.replace(/\\n/g, "\n").trim();
  const auth = new google.auth.GoogleAuth({
    credentials: { client_email: GOOGLE_SA_EMAIL, private_key: key },
    scopes: ["https://www.googleapis.com/auth/spreadsheets"]
  });
  return google.sheets({ version: "v4", auth });
}
