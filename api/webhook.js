// api/webhook.js
const { MercadoPagoConfig, Payment } = require("mercadopago");
const { google } = require("googleapis");

const MP_ACCESS_TOKEN = process.env.MP_ACCESS_TOKEN;
const GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID;
const GOOGLE_SA_EMAIL = process.env.GOOGLE_SA_EMAIL;
const GOOGLE_SA_KEY   = process.env.GOOGLE_SA_KEY;

const EXPIRACION_MS = 10 * 60 * 1000;

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method === "GET")     return res.status(200).send("OK");
  if (req.method !== "POST")    return res.status(405).end();

  try {
    const { type, data } = req.body;
    if (!data?.id) return res.status(200).json({ received: true });

    const client  = new MercadoPagoConfig({ accessToken: MP_ACCESS_TOKEN });
    const payment = new Payment(client);
    const pago    = await payment.get({ id: data.id });

    const estado = pago.status;
    const extRef = pago.external_reference || "";

    console.log("Webhook pago:", data.id, "estado:", estado, "ref:", extRef);

    const sheets    = await getSheetsClient();
    const sheetName = await getSheetName(sheets);

    // Siempre limpiar expiradas al inicio (resuelve el caso de "abandonó sin pagar")
    await limpiarExpiradas(sheets, sheetName);

    if (estado === "approved") {
      await actualizarEstado(sheets, sheetName, extRef, "CONFIRMADA", data.id);
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
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: sheetName + "!A2:M1000"
  });
  const rows = response.data.values || [];

  const parts   = extRef.split("|");
  const courtId = parts[0] || "";
  const date    = parts[1] || "";
  const slot    = parts[2] || "";
  console.log("Buscando -> courtId:[" + courtId + "] date:[" + date + "] slot:[" + slot + "]");

  for (let i = 0; i < rows.length; i++) {
    const row     = rows[i];
    const rCourt  = String(row[3] || "");
    const rDate   = String(row[4] || "");
    const rSlot   = String(row[5] || "");
    const rEstado = String(row[10] || "");
    console.log("Fila " + (i+2) + " -> court:[" + rCourt + "] date:[" + rDate + "] slot:[" + rSlot + "] estado:[" + rEstado + "]");

    if (
      rCourt === courtId &&
      rDate  === date    &&
      rSlot  === slot    &&
      rEstado !== "CONFIRMADA" &&
      rEstado !== "CANCELADA"
    ) {
      const rowNum = i + 2;
      await sheets.spreadsheets.values.update({
        spreadsheetId: GOOGLE_SHEET_ID,
        range: sheetName + "!K" + rowNum,
        valueInputOption: "RAW",
        requestBody: { values: [[nuevoEstado]] }
      });
      await sheets.spreadsheets.values.update({
        spreadsheetId: GOOGLE_SHEET_ID,
        range: sheetName + "!L" + rowNum,
        valueInputOption: "RAW",
        requestBody: { values: [["Pago MP #" + pagoId + " — " + nuevoEstado]] }
      });
      console.log("Sheet actualizado fila", rowNum, "->", nuevoEstado);
      return;
    }
  }
  console.log("No se encontro fila para actualizar");
}

// Borra la fila cuando el pago es rechazado o cancelado
async function borrarReserva(sheets, sheetName, extRef) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: sheetName + "!A2:M1000"
  });
  const rows = response.data.values || [];

  const parts   = extRef.split("|");
  const courtId = parts[0] || "";
  const date    = parts[1] || "";
  const slot    = parts[2] || "";

  for (let i = 0; i < rows.length; i++) {
    const row     = rows[i];
    const rCourt  = String(row[3] || "");
    const rDate   = String(row[4] || "");
    const rSlot   = String(row[5] || "");
    const rEstado = String(row[10] || "");

    if (rCourt === courtId && rDate === date && rSlot === slot && rEstado !== "CONFIRMADA") {
      await borrarFila(sheets, sheetName, i + 1);
      console.log("Fila borrada (pago rechazado/cancelado), índice", i + 1);
      return;
    }
  }
  console.log("borrarReserva: no se encontro fila para borrar");
}

// Borra todas las filas RESERVANDO con más de 10 minutos
async function limpiarExpiradas(sheets, sheetName) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: sheetName + "!A2:M1000"
  });
  const rows  = response.data.values || [];
  const ahora = Date.now();
  const aBorrar = [];

  for (let i = 0; i < rows.length; i++) {
    const rEstado = String(rows[i][10] || "");
    const rTs     = rows[i][12] || "";
    if (rEstado === "RESERVANDO" && rTs) {
      const edad = ahora - new Date(rTs).getTime();
      if (edad > EXPIRACION_MS) {
        aBorrar.push(i + 1); // índice 0-based, fila 1 = header
      }
    }
  }

  if (aBorrar.length === 0) return;

  const meta    = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
  const hoja    = meta.data.sheets.find(function(s) { return s.properties.title === sheetName; });
  if (!hoja) return;
  const sheetId = hoja.properties.sheetId;

  // Borrar de abajo hacia arriba para no desplazar índices
  const sorted   = aBorrar.slice().sort(function(a, b) { return b - a; });
  const requests = sorted.map(function(idx) {
    return { deleteDimension: { range: { sheetId, dimension: "ROWS", startIndex: idx, endIndex: idx + 1 } } };
  });

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: GOOGLE_SHEET_ID,
    requestBody: { requests }
  });
  console.log("Expiradas borradas:", sorted.length);
}

async function borrarFila(sheets, sheetName, rowIndex) {
  const meta    = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
  const hoja    = meta.data.sheets.find(function(s) { return s.properties.title === sheetName; });
  if (!hoja) return;
  const sheetId = hoja.properties.sheetId;

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: GOOGLE_SHEET_ID,
    requestBody: {
      requests: [{
        deleteDimension: {
          range: { sheetId, dimension: "ROWS", startIndex: rowIndex, endIndex: rowIndex + 1 }
        }
      }]
    }
  });
}

async function getSheetName(sheets) {
  try {
    const meta  = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
    const hojas = meta.data.sheets.map(function(s) { return s.properties.title; });
    // Buscar hojas "Mes YYYY-MM" primero
    const meses = hojas.filter(function(h) { return h.startsWith("Mes "); });
    if (meses.length > 0) return meses[meses.length - 1];
    // Fallback a "Semana X"
    const semanas = hojas.filter(function(h) { return h.startsWith("Semana"); });
    if (semanas.length > 0) return semanas[semanas.length - 1];
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
