// api/crear-pago.js
const { MercadoPagoConfig, Preference } = require("mercadopago");
const { google } = require("googleapis");

const MP_ACCESS_TOKEN  = process.env.MP_ACCESS_TOKEN;
const GOOGLE_SHEET_ID  = process.env.GOOGLE_SHEET_ID;
const GOOGLE_SA_EMAIL  = process.env.GOOGLE_SA_EMAIL;
const GOOGLE_SA_KEY    = process.env.GOOGLE_SA_KEY;
const BASE_URL         = process.env.BASE_URL;

const EXPIRACION_MS = 10 * 60 * 1000;

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST")   return res.status(405).json({ error: "Método no permitido" });

  try {
    const { courtId, courtName, date, slot, name, phone, halfPrice, fullPrice } = req.body;

    if (!courtId || !date || !slot || !name || !phone || !halfPrice) {
      return res.status(400).json({ error: "Faltan datos de la reserva" });
    }

    const sheets    = await getSheetsClient();
    const sheetName = await getSheetName(sheets);
    console.log("crear-pago usando hoja:", sheetName);

    const disponible = await checkDisponibilidad(sheets, sheetName, courtId, date, slot);
    if (!disponible) {
      return res.status(409).json({ error: "Ese turno ya fue reservado. Elegí otro." });
    }

    const client     = new MercadoPagoConfig({ accessToken: MP_ACCESS_TOKEN });
    const preference = new Preference(client);
    const externalRef = courtId + "|" + date + "|" + slot + "|" + Date.now();

    const mpResponse = await preference.create({
      body: {
        items: [{
          title: "Seña – " + courtName + " – " + date + " " + slot + "hs",
          quantity: 1,
          unit_price: halfPrice,
          currency_id: "ARS",
          description: "Reserva a nombre de " + name
        }],
        payer: { name },
        external_reference: externalRef,
        back_urls: {
          success: BASE_URL + "/success.html",
          failure: BASE_URL + "/failure.html",
          pending: BASE_URL + "/pending.html"
        },
        auto_return: "approved",
        notification_url: BASE_URL + "/api/webhook",
        statement_descriptor: "CANCHA5",
        metadata: { courtId, courtName, date, slot, name, phone, halfPrice, fullPrice }
      }
    });

    const reservaId = "R" + Date.now().toString().slice(-7);
    const now       = new Date().toISOString();

    // IMPORTANTE: guardamos date con apóstrofe para que Sheets no lo convierta a formato fecha
    // Así el webhook puede comparar "2026-04-20" === "2026-04-20" sin problemas
    await sheets.spreadsheets.values.append({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: sheetName + "!A:M",
      valueInputOption: "RAW",  // RAW = sin interpretar, guarda texto exacto
      requestBody: {
        values: [[
          reservaId,
          now,
          courtName,
          String(courtId),
          date,        // se guarda como texto exacto gracias a RAW
          slot,
          name,
          String(phone),
          fullPrice,
          halfPrice,
          "RESERVANDO",
          "Pref MP: " + mpResponse.id,
          now
        ]]
      }
    });

    console.log("Fila RESERVANDO guardada OK en:", sheetName, "date:", date, "slot:", slot);

    return res.status(200).json({
      init_point:    mpResponse.init_point,
      preference_id: mpResponse.id,
      reserva_id:    reservaId,
      expires_at:    Date.now() + EXPIRACION_MS
    });

  } catch (err) {
    console.error("Error en crear-pago:", err);
    return res.status(500).json({ error: "Error interno: " + err.message });
  }
};

async function checkDisponibilidad(sheets, sheetName, courtId, date, slot) {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: sheetName + "!A2:M1000"
    });
    const rows  = response.data.values || [];
    const ahora = Date.now();

    for (const row of rows) {
      const rCourtId = String(row[3] || "");
      const rDate    = row[4]  || "";
      const rSlot    = row[5]  || "";
      const rEstado  = row[10] || "";
      const rTs      = row[12] || "";

      if (rCourtId !== String(courtId) || rDate !== date || rSlot !== slot) continue;
      if (rEstado === "CANCELADA") continue;
      if (rEstado === "RESERVANDO" && rTs) {
        const edad = ahora - new Date(rTs).getTime();
        if (edad > EXPIRACION_MS) continue;
      }
      return false;
    }
    return true;
  } catch (e) {
    console.error("Error verificando disponibilidad:", e);
    return true;
  }
}

async function getSheetName(sheets) {
  try {
    const meta  = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
    const hojas = meta.data.sheets.map(function(s) { return s.properties.title; });
    console.log("Hojas disponibles:", JSON.stringify(hojas));
    const semanas = hojas.filter(function(h) { return h.startsWith("Semana"); });
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
