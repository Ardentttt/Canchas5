// api/webhook.js
// Mercado Pago llama a este endpoint automáticamente cuando se confirma un pago.
// Actualiza el estado en Google Sheets a CONFIRMADA.

const { MercadoPagoConfig, Payment } = require("mercadopago");
const { google } = require("googleapis");

const MP_ACCESS_TOKEN = process.env.MP_ACCESS_TOKEN;
const GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID;
const GOOGLE_SA_EMAIL = process.env.GOOGLE_SA_EMAIL;
const GOOGLE_SA_KEY   = process.env.GOOGLE_SA_KEY;
const WA_NUMBER       = process.env.WA_NUMBER || "5491166140749";

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  if (req.method === "OPTIONS") return res.status(200).end();

  // MP envía GET para validar la URL
  if (req.method === "GET") return res.status(200).send("OK");

  if (req.method !== "POST") return res.status(405).end();

  try {
    const { type, data } = req.body;

    // Solo procesar notificaciones de pago
    if (type !== "payment" || !data?.id) {
      return res.status(200).json({ received: true });
    }

    // ── Consultar el pago a MP ──
    const client  = new MercadoPagoConfig({ accessToken: MP_ACCESS_TOKEN });
    const payment = new Payment(client);
    const pago    = await payment.get({ id: data.id });

    const estado  = pago.status;          // "approved", "pending", "rejected"
    const meta    = pago.metadata || {};
    const extRef  = pago.external_reference || "";

    console.log("Webhook pago:", data.id, "estado:", estado, "ref:", extRef);

    if (estado === "approved") {
      // ── Actualizar Sheet a CONFIRMADA ──
      await actualizarEstado(extRef, "CONFIRMADA", pago.id);
      console.log("Reserva CONFIRMADA:", extRef);

    } else if (estado === "rejected" || estado === "cancelled") {
      await actualizarEstado(extRef, "CANCELADA", pago.id);
      console.log("Reserva CANCELADA:", extRef);

    } else {
      // pending, in_process, etc.
      await actualizarEstado(extRef, "PENDIENTE", pago.id);
    }

    return res.status(200).json({ received: true });

  } catch (err) {
    console.error("Error en webhook:", err);
    return res.status(500).json({ error: err.message });
  }
};

// ── Busca la fila por external_reference y actualiza el estado ──
async function actualizarEstado(extRef, nuevoEstado, pagoId) {
  const sheets = await getSheetsClient();

  // Leer todas las filas
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: "Reservas!A2:L1000"
  });
  const rows = response.data.values || [];

  // Buscar por referencia externa (columna A = ID, columna L = notas con pref MP)
  for (let i = 0; i < rows.length; i++) {
    const notaCol = rows[i][11] || "";
    // La nota tiene "Pref MP: XXX" — buscamos por extRef que también contiene el prefId
    if (notaCol.includes(extRef.split("|")[3]) || rows[i][0] === extRef) {
      const rowNum = i + 2; // +2 porque empezamos en fila 2

      // Actualizar estado (columna K = 11)
      await sheets.spreadsheets.values.update({
        spreadsheetId: GOOGLE_SHEET_ID,
        range: `Reservas!K${rowNum}`,
        valueInputOption: "USER_ENTERED",
        requestBody: { values: [[nuevoEstado]] }
      });

      // Actualizar notas con el ID del pago (columna L = 12)
      await sheets.spreadsheets.values.update({
        spreadsheetId: GOOGLE_SHEET_ID,
        range: `Reservas!L${rowNum}`,
        valueInputOption: "USER_ENTERED",
        requestBody: { values: [["Pago MP #" + pagoId + " – " + nuevoEstado]] }
      });

      console.log("Sheet actualizado fila", rowNum, "→", nuevoEstado);
      break;
    }
  }
}

async function getSheetsClient() {
  const auth = new google.auth.JWT(
    GOOGLE_SA_EMAIL,
    null,
    GOOGLE_SA_KEY.replace(/\\n/g, "\n"),
    ["https://www.googleapis.com/auth/spreadsheets"]
  );
  return google.sheets({ version: "v4", auth });
}
