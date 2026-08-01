// api/crear-pago.js
// NO escribe en el Sheet. Solo crea la preferencia de MP con los datos en metadata.
// El webhook escribe la fila directamente cuando el pago se aprueba.

const { MercadoPagoConfig, Preference } = require("mercadopago");
const { google } = require("googleapis");

const MP_ACCESS_TOKEN = process.env.MP_ACCESS_TOKEN;
const GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID;
const GOOGLE_SA_EMAIL = process.env.GOOGLE_SA_EMAIL;
const GOOGLE_SA_KEY   = process.env.GOOGLE_SA_KEY;
const BASE_URL        = process.env.BASE_URL;

const EXPIRACION_MS = 10 * 60 * 1000;

const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
               "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

const HEADERS = ["ID","Hora","Cancha","Cancha ID","Fecha Turno",
                 "Horario","Nombre","Teléfono","Precio Total","Estado","Notas"];

function aHoraArgentina(isoStr) {
  return new Date(isoStr).toLocaleString("es-AR", {
    timeZone: "America/Argentina/Buenos_Aires",
    day: "2-digit", month: "2-digit", year: "numeric",
    hour: "2-digit", minute: "2-digit", hour12: false
  });
}

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
    const sheetName = await getOrCreateSheetName(sheets, date);

    // Verificar que el turno no esté ya CONFIRMADA en el Sheet
    const disponible = await checkDisponibilidad(sheets, sheetName, courtId, date, slot);
    if (!disponible) {
      return res.status(409).json({ error: "Ese turno ya fue reservado. Elegí otro." });
    }

    const client = new MercadoPagoConfig({ accessToken: MP_ACCESS_TOKEN });
    const pref   = new Preference(client);
    const extRef = courtId + "|" + date + "|" + slot + "|" + Date.now();
    const nowAR  = aHoraArgentina(new Date().toISOString());

    const mpResponse = await pref.create({
      body: {
        items: [{ title: "Seña – " + courtName + " – " + date + " " + slot + "hs",
                  quantity: 1, unit_price: halfPrice, currency_id: "ARS",
                  description: "Reserva a nombre de " + name }],
        payer: { name },
        external_reference: extRef,
        back_urls: { success: BASE_URL + "/success.html",
                     failure: BASE_URL + "/failure.html",
                     pending: BASE_URL + "/pending.html" },
        auto_return: "approved",
        notification_url: BASE_URL + "/api/webhook",
        statement_descriptor: "CANCHA5",
        // Todos los datos van en metadata para que el webhook los use
        metadata: {
          courtId: String(courtId),
          courtName, date, slot, name,
          phone: String(phone),
          fullPrice, halfPrice,
          horaAR: nowAR,
          sheetName
        }
      }
    });

    console.log("Preferencia creada, sin escribir en Sheet:", extRef);

    return res.status(200).json({
      init_point:    mpResponse.init_point,
      preference_id: mpResponse.id,
      expires_at:    Date.now() + EXPIRACION_MS
    });

  } catch (err) {
    console.error("Error en crear-pago:", err);
    return res.status(500).json({ error: "Error interno: " + err.message });
  }
};

async function checkDisponibilidad(sheets, sheetName, courtId, date, slot) {
  try {
    const r    = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID, range: sheetName + "!A2:K1000"
    });
    const rows = r.data.values || [];
    for (const row of rows) {
      if (String(row[3]||"") !== String(courtId)) continue;
      if ((row[4]||"") !== date) continue;
      if ((row[5]||"") !== slot) continue;
      const estado = row[9] || "";
      // Solo bloquear si está CONFIRMADA o PENDIENTE
      if (estado === "CONFIRMADA" || estado === "PENDIENTE") return false;
    }
    return true;
  } catch (e) {
    console.error("Error disponibilidad:", e);
    return true;
  }
}

async function getOrCreateSheetName(sheets, fechaTurno) {
  const ahora  = fechaTurno ? new Date(fechaTurno + "T00:00:00") : new Date();
  const nombre = MESES[ahora.getMonth()] + " " + ahora.getFullYear();
  const meta   = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
  const hojas  = meta.data.sheets.map(function(s) { return s.properties.title; });
  if (hojas.includes(nombre)) return nombre;

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
