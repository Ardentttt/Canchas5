// api/crear-pago.js
// Recibe los datos de la reserva, crea una preferencia de pago en MP
// y devuelve el link para redirigir al cliente.

const { MercadoPagoConfig, Preference } = require("mercadopago");
const { google } = require("googleapis");

// ── Lee variables de entorno (las configurás en Vercel, no acá) ──
const MP_ACCESS_TOKEN  = process.env.MP_ACCESS_TOKEN;   // Tu Access Token de MP
const GOOGLE_SHEET_ID  = process.env.GOOGLE_SHEET_ID;   // ID de tu Google Sheet
const GOOGLE_SA_EMAIL  = process.env.GOOGLE_SA_EMAIL;   // Email de la service account
const GOOGLE_SA_KEY    = process.env.GOOGLE_SA_KEY;     // Clave privada (con \n reales)
const BASE_URL         = process.env.BASE_URL;           // URL de tu app en Vercel

module.exports = async function handler(req, res) {
  // CORS
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST")   return res.status(405).json({ error: "Método no permitido" });

  try {
    const { courtId, courtName, date, slot, name, phone, halfPrice, fullPrice } = req.body;

    // Validar campos
    if (!courtId || !date || !slot || !name || !phone || !halfPrice) {
      return res.status(400).json({ error: "Faltan datos de la reserva" });
    }

    // ── 1. Verificar que el turno esté disponible en Google Sheets ──
    const disponible = await checkDisponibilidad(courtId, date, slot);
    if (!disponible) {
      return res.status(409).json({ error: "Ese turno ya fue reservado. Elegí otro." });
    }

    // ── 2. Crear preferencia de pago en Mercado Pago ──
    const client = new MercadoPagoConfig({ accessToken: MP_ACCESS_TOKEN });
    const preference = new Preference(client);

    const externalRef = `${courtId}|${date}|${slot}|${Date.now()}`;

    const mpResponse = await preference.create({
      body: {
        items: [
          {
            title: `Seña – ${courtName} – ${date} ${slot}hs`,
            quantity: 1,
            unit_price: halfPrice,
            currency_id: "ARS",
            description: `Reserva a nombre de ${name}`
          }
        ],
        payer: { name },
        external_reference: externalRef,
        back_urls: {
          success: `${BASE_URL}/success.html`,
          failure: `${BASE_URL}/failure.html`,
          pending: `${BASE_URL}/pending.html`
        },
        auto_return: "approved",
        notification_url: `${BASE_URL}/api/webhook`,
        statement_descriptor: "CANCHA5",
        metadata: { courtId, courtName, date, slot, name, phone, halfPrice, fullPrice }
      }
    });

    // ── 3. Pre-registrar en Sheet como RESERVANDO (bloquea el turno) ──
    const reservaId = "R" + Date.now().toString().slice(-7);
    await registrarEnSheet({
      id: reservaId,
      courtId, courtName, date, slot,
      name, phone, halfPrice, fullPrice,
      estado: "RESERVANDO",
      mpPrefId: mpResponse.id
    });

    return res.status(200).json({
      init_point: mpResponse.init_point,   // Link para redirigir al cliente
      preference_id: mpResponse.id,
      reserva_id: reservaId
    });

  } catch (err) {
    console.error("Error en crear-pago:", err);
    return res.status(500).json({ error: "Error interno: " + err.message });
  }
};

// ── Verifica disponibilidad en Google Sheets ──
async function checkDisponibilidad(courtId, date, slot) {
  try {
    const sheets = await getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: "Reservas!A2:L1000"
    });
    const rows = response.data.values || [];
    for (const row of rows) {
      const rCourtId = String(row[3]);  // Columna D: Cancha ID
      const rDate    = row[4];           // Columna E: Fecha Turno
      const rSlot    = row[5];           // Columna F: Horario
      const rEstado  = row[10];          // Columna K: Estado
      if (
        rCourtId === String(courtId) &&
        rDate    === date &&
        rSlot    === slot &&
        rEstado  !== "CANCELADA"
      ) {
        return false; // Ya ocupado
      }
    }
    return true;
  } catch (e) {
    console.error("Error verificando disponibilidad:", e);
    return true; // Si falla, deja pasar (el webhook lo corrige)
  }
}

// ── Registra la reserva en Google Sheets ──
async function registrarEnSheet(data) {
  const sheets = await getSheetsClient();
  const now    = new Date().toISOString();
  await sheets.spreadsheets.values.append({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: "Reservas!A:L",
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [[
        data.id,
        now,
        data.courtName,
        data.courtId,
        data.date,
        data.slot,
        data.name,
        data.phone,
        data.fullPrice,
        data.halfPrice,
        data.estado,
        "Pref MP: " + data.mpPrefId
      ]]
    }
  });
}

// ── Cliente autenticado de Google Sheets ──
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
