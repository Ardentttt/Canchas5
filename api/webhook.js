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
  const sheets = await getSheetsClient();
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: "Reservas!A2:L1000"
  });
  const rows = response.data.values || [];

  // extRef = "courtId|date|slot|timestamp"
  // La nota en columna L tiene "Pref MP: prefId"
  // Buscamos por courtId+date+slot que están en columnas D, E, F
  const parts   = extRef.split("|");
  const courtId = parts[0] || "";
  const date    = parts[1] || "";
  const slot    = parts[2] || "";

  for (let i = 0; i < rows.length; i++) {
    const row      = rows[i];
    const rCourt   = String(row[3] || "");
    const rDate    = String(row[4] || "");
    const rSlot    = String(row[5] || "");
    const rEstado  = String(row[10] || "");

    // Buscar por cancha + fecha + slot que no estén ya CONFIRMADA/CANCELADA
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
        range: `Reservas!K${rowNum}`,
        valueInputOption: "USER_ENTERED",
        requestBody: { values: [[nuevoEstado]] }
      });

      await sheets.spreadsheets.values.update({
        spreadsheetId: GOOGLE_SHEET_ID,
        range: `Reservas!L${rowNum}`,
        valueInputOption: "USER_ENTERED",
        requestBody: { values: [["Pago MP #" + pagoId + " — " + nuevoEstado]] }
      });

      console.log("Sheet actualizado fila", rowNum, "->", nuevoEstado);
      break;
    }
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
