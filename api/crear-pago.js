// api/crear-pago.js
// Guarda en hoja "Temp" para bloquear el turno 10 min.
// La hoja mensual solo recibe entradas CONFIRMADAS (desde el webhook).

const { MercadoPagoConfig, Preference } = require("mercadopago");
const { google } = require("googleapis");

const MP_ACCESS_TOKEN = process.env.MP_ACCESS_TOKEN;
const GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID;
const GOOGLE_SA_EMAIL = process.env.GOOGLE_SA_EMAIL;
const GOOGLE_SA_KEY   = process.env.GOOGLE_SA_KEY;
const BASE_URL        = process.env.BASE_URL;

const EXPIRACION_MS = 10 * 60 * 1000;
const TEMP_SHEET    = "Temp";

const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
               "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

const HEADERS_MENSUAL = ["ID","Hora","Cancha","Cancha ID","Fecha Turno",
                         "Horario","Nombre","Teléfono","Precio Total","Estado","Notas"];

const HEADERS_TEMP = ["ExtRef","CanchaID","Fecha","Slot","Timestamp"];

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

    const sheets = await getSheetsClient();

    // Asegurar que exista la hoja Temp
    await ensureTempSheet(sheets);

    // Limpiar expiradas de Temp
    await limpiarTempExpiradas(sheets);

    // Verificar disponibilidad: hoja mensual + Temp
    const sheetMes   = await getOrCreateSheetName(sheets, date);
    const disponible = await checkDisponibilidad(sheets, sheetMes, courtId, date, slot);
    if (!disponible) {
      return res.status(409).json({ error: "Ese turno ya fue reservado. Elegí otro." });
    }

    // Crear preferencia de MP
    const client = new MercadoPagoConfig({ accessToken: MP_ACCESS_TOKEN });
    const pref   = new Preference(client);
    const extRef = courtId + "|" + date + "|" + slot + "|" + Date.now();
    const nowISO = new Date().toISOString();
    const nowAR  = aHoraArgentina(nowISO);

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
        metadata: {
          court_id:   String(courtId),
          court_name: courtName,
          date, slot, name,
          phone:      String(phone),
          full_price: fullPrice,
          half_price: halfPrice,
          hora_ar:    nowAR,
          sheet_name: sheetMes
        }
      }
    });

    // Escribir en Temp para bloquear el turno 10 min
    await sheets.spreadsheets.values.append({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: TEMP_SHEET + "!A:E",
      valueInputOption: "RAW",
      requestBody: { values: [[extRef, String(courtId), date, slot, nowISO]] }
    });

    console.log("Turno bloqueado en Temp:", date, slot, "por 10 min");

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

async function checkDisponibilidad(sheets, sheetMes, courtId, date, slot) {
  try {
    // 1. Verificar en hoja mensual (CONFIRMADA / PENDIENTE)
    const rMes = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID, range: sheetMes + "!A2:K1000"
    });
    for (const row of rMes.data.values || []) {
      if (String(row[3]||"") !== String(courtId)) continue;
      if ((row[4]||"") !== date) continue;
      if ((row[5]||"") !== slot) continue;
      const est = row[9] || "";
      if (est === "CONFIRMADA" || est === "PENDIENTE") return false;
    }

    // 2. Verificar en Temp (bloqueos activos)
    const rTemp = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID, range: TEMP_SHEET + "!A2:E1000"
    });
    const ahora = Date.now();
    for (const row of rTemp.data.values || []) {
      if (String(row[1]||"") !== String(courtId)) continue;
      if ((row[2]||"") !== date) continue;
      if ((row[3]||"") !== slot) continue;
      const ts  = row[4] || "";
      const edad = ahora - new Date(ts).getTime();
      if (edad <= EXPIRACION_MS) return false; // bloqueado activo
    }

    return true;
  } catch (e) {
    console.error("Error disponibilidad:", e);
    return true;
  }
}

async function limpiarTempExpiradas(sheets) {
  try {
    const r    = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID, range: TEMP_SHEET + "!A2:E1000"
    });
    const rows  = r.data.values || [];
    const ahora = Date.now();
    const aBorrar = [];

    for (let i = 0; i < rows.length; i++) {
      const ts   = rows[i][4] || "";
      const edad = ahora - new Date(ts).getTime();
      if (edad > EXPIRACION_MS) aBorrar.push(i + 1);
    }

    if (aBorrar.length === 0) return;

    const meta    = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
    const hoja    = meta.data.sheets.find(function(s) { return s.properties.title === TEMP_SHEET; });
    if (!hoja) return;
    const sheetId = hoja.properties.sheetId;

    const sorted = aBorrar.slice().sort(function(a, b) { return b - a; });
    const reqs   = sorted.map(function(idx) {
      return { deleteDimension: { range: { sheetId, dimension: "ROWS", startIndex: idx, endIndex: idx + 1 } } };
    });
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: GOOGLE_SHEET_ID, requestBody: { requests: reqs }
    });
    console.log("Temp: expiradas borradas:", sorted.length);
  } catch (e) {
    console.error("Error limpiando Temp:", e);
  }
}

async function ensureTempSheet(sheets) {
  const meta  = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
  const hojas = meta.data.sheets.map(function(s) { return s.properties.title; });
  if (hojas.includes(TEMP_SHEET)) return;

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: GOOGLE_SHEET_ID,
    requestBody: { requests: [{ addSheet: {
      properties: { title: TEMP_SHEET, gridProperties: { rowCount: 200, columnCount: 5 } }
    }}]}
  });
  await sheets.spreadsheets.values.update({
    spreadsheetId: GOOGLE_SHEET_ID, range: TEMP_SHEET + "!A1:E1",
    valueInputOption: "RAW",
    requestBody: { values: [HEADERS_TEMP] }
  });
  console.log("Hoja Temp creada");
}

async function getOrCreateSheetName(sheets, fechaTurno) {
  const base   = fechaTurno ? new Date(fechaTurno + "T00:00:00") : new Date();
  const nombre = MESES[base.getMonth()] + " " + base.getFullYear();
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
    valueInputOption: "RAW", requestBody: { values: [HEADERS_MENSUAL] }
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
  console.log("Hoja mensual creada:", nombre);
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
