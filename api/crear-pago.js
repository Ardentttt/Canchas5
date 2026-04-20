// api/crear-pago.js
// Recibe los datos de la reserva, crea una preferencia de pago en MP
// y devuelve el link para redirigir al cliente.

const { MercadoPagoConfig, Preference } = require("mercadopago");
const { google } = require("googleapis");

const MP_ACCESS_TOKEN  = process.env.MP_ACCESS_TOKEN;
const GOOGLE_SHEET_ID  = process.env.GOOGLE_SHEET_ID;
const GOOGLE_SA_EMAIL  = process.env.GOOGLE_SA_EMAIL;
const GOOGLE_SA_KEY    = process.env.GOOGLE_SA_KEY;
const BASE_URL         = process.env.BASE_URL;

const EXPIRACION_MS = 10 * 60 * 1000; // 10 minutos

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

    // ── 1. Verificar disponibilidad ──
    const disponible = await checkDisponibilidad(sheets, sheetName, courtId, date, slot);
    if (!disponible) {
      return res.status(409).json({ error: "Ese turno ya fue reservado. Elegí otro." });
    }

    // ── 2. Crear preferencia de pago en Mercado Pago ──
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

    // ── 3. Pre-registrar en Sheet como RESERVANDO con timestamp ──
    const reservaId = "R" + Date.now().toString().slice(-7);
    const now       = new Date().toISOString();

    await sheets.spreadsheets.values.append({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: sheetName + "!A:M",
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[
          reservaId,
          now,
          courtName,
          courtId,
          date,
          slot,
          name,
          phone,
          fullPrice,
          halfPrice,
          "RESERVANDO",
          "Pref MP: " + mpResponse.id,
          now   // columna M: timestamp de inicio para calcular expiración
        ]]
      }
    });

    // Aplicar formato de colores a la hoja
    applyConditionalFormatting(sheets, sheetName).catch(console.error);

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
      if (rEstado === "CANCELADA" || rEstado === "EXPIRADA") continue;

      // RESERVANDO expirada no bloquea
      if (rEstado === "RESERVANDO" && rTs) {
        const edad = ahora - new Date(rTs).getTime();
        if (edad > 10 * 60 * 1000) continue;
      }

      return false; // ocupado
    }
    return true;
  } catch (e) {
    console.error("Error verificando disponibilidad:", e);
    return true;
  }
}

// Aplica formato condicional de colores una vez por hoja
// Verde = CONFIRMADA, Rojo = CANCELADA/EXPIRADA, Naranja = PENDIENTE, Amarillo = RESERVANDO
async function applyConditionalFormatting(sheets, sheetName) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
  const hoja = meta.data.sheets.find(function(s) { return s.properties.title === sheetName; });
  if (!hoja) return;
  const sheetId = hoja.properties.sheetId;

  // Columna K = índice 10
  const col = 10;

  const reglas = [
    { valor: "CONFIRMADA", r: 0.204, g: 0.659, b: 0.325 },   // verde
    { valor: "CANCELADA",  r: 0.957, g: 0.263, b: 0.212 },   // rojo
    { valor: "EXPIRADA",   r: 0.957, g: 0.263, b: 0.212 },   // rojo
    { valor: "PENDIENTE",  r: 1.0,   g: 0.596, b: 0.0   },   // naranja
    { valor: "RESERVANDO", r: 1.0,   g: 0.898, b: 0.2   }    // amarillo
  ];

  const requests = reglas.map(function(r) {
    return {
      addConditionalFormatRule: {
        rule: {
          ranges: [{
            sheetId,
            startRowIndex: 1,
            startColumnIndex: col,
            endColumnIndex: col + 1
          }],
          booleanRule: {
            condition: {
              type: "TEXT_EQ",
              values: [{ userEnteredValue: r.valor }]
            },
            format: {
              backgroundColor: { red: r.r, green: r.g, blue: r.b },
              textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } }
            }
          }
        },
        index: 0
      }
    };
  });

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: GOOGLE_SHEET_ID,
    requestBody: { requests }
  });
}

async function getSheetName(sheets) {
  try {
    const meta  = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
    const hojas = meta.data.sheets.map(function(s) { return s.properties.title; });
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
    credentials: {
      client_email: GOOGLE_SA_EMAIL,
      private_key: key
    },
    scopes: ["https://www.googleapis.com/auth/spreadsheets"]
  });
  return google.sheets({ version: "v4", auth });
}
