// api/webhook.js
const { MercadoPagoConfig, Payment } = require("mercadopago");
const { google } = require("googleapis");

const MP_ACCESS_TOKEN = process.env.MP_ACCESS_TOKEN;
const GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID;
const GOOGLE_SA_EMAIL = process.env.GOOGLE_SA_EMAIL;
const GOOGLE_SA_KEY   = process.env.GOOGLE_SA_KEY;

const EXPIRACION_MS = 10 * 60 * 1000;

const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
               "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

// Columnas: A=ID B=Hora C=Cancha D=CanchaID E=FechaTurno F=Horario
//           G=Nombre H=Tel I=Precio J=Estado K=Notas

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
    console.log("Webhook:", data.id, "estado:", estado, "ref:", extRef);

    const sheets    = await getSheetsClient();
    const sheetName = await getOrCreateSheetName(sheets);

    await limpiarExpiradas(sheets, sheetName);

    if (estado === "approved") {
      await actualizarEstado(sheets, sheetName, extRef, "CONFIRMADA", data.id);
      await ordenarHoja(sheets, sheetName);
    } else if (estado === "rejected" || estado === "cancelled") {
      await borrarReserva(sheets, sheetName, extRef);
    } else {
      await actualizarEstado(sheets, sheetName, extRef, "PENDIENTE", data.id);
    }

    return res.status(200).json({ received: true });
  } catch (err) {
    console.error("Error webhook:", err);
    return res.status(500).json({ error: err.message });
  }
};

async function actualizarEstado(sheets, sheetName, extRef, nuevoEstado, pagoId) {
  const r    = await sheets.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_ID, range: sheetName + "!A2:K1000"
  });
  const rows = r.data.values || [];

  const parts   = extRef.split("|");
  const courtId = parts[0] || "";
  const date    = parts[1] || "";
  const slot    = parts[2] || "";

  for (let i = 0; i < rows.length; i++) {
    const row    = rows[i];
    const rCourt = String(row[3] || "");
    const rDate  = String(row[4] || "");
    const rSlot  = String(row[5] || "");
    const rEst   = String(row[9] || "");

    if (rCourt === courtId && rDate === date && rSlot === slot
        && rEst !== "CONFIRMADA" && rEst !== "CANCELADA") {
      const n = i + 2;
      // Col J = Estado, Col K = Notas
      await sheets.spreadsheets.values.update({
        spreadsheetId: GOOGLE_SHEET_ID,
        range: sheetName + "!J" + n,
        valueInputOption: "RAW",
        requestBody: { values: [[nuevoEstado]] }
      });
      await sheets.spreadsheets.values.update({
        spreadsheetId: GOOGLE_SHEET_ID,
        range: sheetName + "!K" + n,
        valueInputOption: "RAW",
        requestBody: { values: [["Pago MP #" + pagoId]] }
      });
      console.log("Fila", n, "->", nuevoEstado);
      return;
    }
  }
  console.log("No se encontro fila para:", courtId, date, slot);
}

async function borrarReserva(sheets, sheetName, extRef) {
  const r    = await sheets.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_ID, range: sheetName + "!A2:K1000"
  });
  const rows = r.data.values || [];

  const parts   = extRef.split("|");
  const courtId = parts[0] || "";
  const date    = parts[1] || "";
  const slot    = parts[2] || "";

  for (let i = 0; i < rows.length; i++) {
    const row  = rows[i];
    if (String(row[3]||"") === courtId && String(row[4]||"") === date
        && String(row[5]||"") === slot && String(row[9]||"") !== "CONFIRMADA") {
      await borrarFila(sheets, sheetName, i + 1);
      console.log("Fila borrada (rechazado/cancelado)");
      return;
    }
  }
}

async function limpiarExpiradas(sheets, sheetName) {
  const r    = await sheets.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_ID, range: sheetName + "!A2:K1000"
  });
  const rows    = r.data.values || [];
  const ahora   = Date.now();
  const aBorrar = [];

  for (let i = 0; i < rows.length; i++) {
    const rEst  = String(rows[i][9] || "");
    const notas = rows[i][10] || "";
    if (rEst === "RESERVANDO") {
      const match = notas.match(/TS:(.+)$/);
      if (match) {
        const edad = ahora - new Date(match[1]).getTime();
        if (edad > EXPIRACION_MS) aBorrar.push(i + 1);
      }
    }
  }

  if (aBorrar.length === 0) return;
  await borrarVariasFilas(sheets, sheetName, aBorrar);
  console.log("Expiradas borradas:", aBorrar.length);
}

// Ordena la hoja por Fecha Turno (col E) + Horario (col F) ascendente
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

async function borrarFila(sheets, sheetName, idx) {
  await borrarVariasFilas(sheets, sheetName, [idx]);
}

async function borrarVariasFilas(sheets, sheetName, indexes) {
  const meta    = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
  const hoja    = meta.data.sheets.find(function(s) { return s.properties.title === sheetName; });
  if (!hoja) return;
  const sheetId = hoja.properties.sheetId;
  const sorted  = indexes.slice().sort(function(a, b) { return b - a; });
  const reqs    = sorted.map(function(i) {
    return { deleteDimension: { range: { sheetId, dimension: "ROWS", startIndex: i, endIndex: i + 1 } } };
  });
  await sheets.spreadsheets.batchUpdate({ spreadsheetId: GOOGLE_SHEET_ID, requestBody: { requests: reqs } });
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
  const key = GOOGLE_SA_KEY.replace(/\\n/g, "\n").trim();
  const auth = new google.auth.GoogleAuth({
    credentials: { client_email: GOOGLE_SA_EMAIL, private_key: key },
    scopes: ["https://www.googleapis.com/auth/spreadsheets"]
  });
  return google.sheets({ version: "v4", auth });
}
