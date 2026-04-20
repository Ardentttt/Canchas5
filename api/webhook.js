// api/webhook.js
const { MercadoPagoConfig, Payment } = require("mercadopago");
const { google } = require("googleapis");

const MP_ACCESS_TOKEN = process.env.MP_ACCESS_TOKEN;
const GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID;
const GOOGLE_SA_EMAIL = process.env.GOOGLE_SA_EMAIL;
const GOOGLE_SA_KEY   = process.env.GOOGLE_SA_KEY;

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

    if (estado === "approved") {
      await actualizarEstado(extRef, "CONFIRMADA", data.id);
    } else if (estado === "rejected" || estado === "cancelled") {
      await actualizarEstado(extRef, "CANCELADA", data.id);
    } else {
      await actualizarEstado(extRef, "PENDIENTE", data.id);
    }

    return res.status(200).json({ received: true });
  } catch (err) {
    console.error("Error webhook:", err);
    return res.status(500).json({ error: err.message });
  }
};

async function actualizarEstado(extRef, nuevoEstado, pagoId) {
  const sheets    = await getSheetsClient();
  const sheetName = await getSheetName(sheets);

  console.log("Buscando en hoja:", sheetName);

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: sheetName + "!A2:M1000"
  });
  const rows = response.data.values || [];
  console.log("Total filas encontradas:", rows.length);

  const parts   = extRef.split("|");
  const courtId = parts[0] || "";
  const date    = parts[1] || "";
  const slot    = parts[2] || "";
  console.log("Buscando courtId:", courtId, "date:", date, "slot:", slot);

  for (let i = 0; i < rows.length; i++) {
    const row     = rows[i];
    const rCourt  = String(row[3] || "");
    const rDate   = String(row[4] || "");
    const rSlot   = String(row[5] || "");
    const rEstado = String(row[10] || "");

    console.log("Fila", i+2, "-> court:", rCourt, "date:", rDate, "slot:", rSlot, "estado:", rEstado);

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
        valueInputOption: "USER_ENTERED",
        requestBody: { values: [[nuevoEstado]] }
      });

      await sheets.spreadsheets.values.update({
        spreadsheetId: GOOGLE_SHEET_ID,
        range: sheetName + "!L" + rowNum,
        valueInputOption: "USER_ENTERED",
        requestBody: { values: [["Pago MP #" + pagoId + " — " + nuevoEstado]] }
      });

      console.log("Sheet actualizado fila", rowNum, "->", nuevoEstado);
      return;
    }
  }
  console.log("ADVERTENCIA: No se encontro fila para actualizar");
}

async function getSheetName(sheets) {
  try {
    const meta  = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
    const hojas = meta.data.sheets.map(function(s) { return s.properties.title; });
    console.log("Hojas disponibles:", JSON.stringify(hojas));
    const semanas = hojas.filter(function(h) { return h.startsWith("Semana"); });
    console.log("Semanas encontradas:", JSON.stringify(semanas));
    if (semanas.length > 0) return semanas[semanas.length - 1];
    if (hojas.includes("Reservas")) return "Reservas";
    return hojas[0];
  } catch (e) {
    console.log("Error getSheetName:", e.message);
    return "Reservas";
  }
}

async function getSheetsClient() {
  const key = GOOGLE_SA_KEY.replace(/\\n/g, "\n").trim();
  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: GOOGLE_SA_EMAIL,
      private_key: key
    },
    scopes: ["https://www.googleapis.com/auth/spreadsheets"]
  });
  return google.sheets({ version: "v4", auth });
}
