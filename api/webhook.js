// api/webhook.js
// Al aprobar: escribe en hoja mensual + crea evento en Google Calendar + borra de Temp + ordena.
// Al rechazar/cancelar: solo borra de Temp.

const { MercadoPagoConfig, Payment } = require("mercadopago");
const { google } = require("googleapis");

const MP_ACCESS_TOKEN    = process.env.MP_ACCESS_TOKEN;
const GOOGLE_SHEET_ID    = process.env.GOOGLE_SHEET_ID;
const GOOGLE_CALENDAR_ID = process.env.GOOGLE_CALENDAR_ID;
const GOOGLE_SA_EMAIL    = process.env.GOOGLE_SA_EMAIL;
const GOOGLE_SA_KEY      = process.env.GOOGLE_SA_KEY;

const TEMP_SHEET = "Temp";

const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
               "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

const HEADERS_MENSUAL = ["ID","Hora","Cancha","Cancha ID","Fecha Turno",
                         "Horario","Nombre","Teléfono","Precio Total","Estado","Notas"];

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method === "GET")     return res.status(200).send("OK");
  if (req.method !== "POST")    return res.status(405).end();

  try {
    const paymentId = req.body?.data?.id || req.body?.id || req.query?.id || req.query?.["data.id"];
    if (!paymentId) return res.status(200).json({ received: true });

    const client  = new MercadoPagoConfig({ accessToken: MP_ACCESS_TOKEN });
    const payment = new Payment(client);
    const pago    = await payment.get({ id: paymentId });

    const estado = pago.status;
    const extRef = pago.external_reference || "";
    const meta   = pago.metadata || {};

    console.log("Webhook:", paymentId, "estado:", estado);

    const auth = getGoogleAuth();
    const sheets = google.sheets({ version: "v4", auth });

    if (estado === "approved") {
      const fechaTurno = meta.date || "";
      const sheetName  = await getOrCreateSheetName(sheets, fechaTurno);

      // Evitar duplicados
      const yaExiste = await buscarConfirmada(sheets, sheetName, extRef);
      if (!yaExiste) {
        const reservaId = "R" + paymentId.toString().slice(-7);
        await sheets.spreadsheets.values.append({
          spreadsheetId: GOOGLE_SHEET_ID,
          range: sheetName + "!A:K",
          valueInputOption: "RAW",
          requestBody: { values: [[
            reservaId,
            meta.hora_ar    || "",
            meta.court_name || "",
            meta.court_id   || "",
            meta.date       || "",
            meta.slot       || "",
            meta.name       || "",
            meta.phone      || "",
            meta.full_price || "",
            "CONFIRMADA",
            "Pago MP #" + paymentId
          ]]}
        });
        console.log("CONFIRMADA escrita en:", sheetName);
        
        // Crear evento en Google Calendar con horario exacto
        if (GOOGLE_CALENDAR_ID && meta.date && meta.slot) {
          await crearEventoCalendar(auth, {
            courtName: meta.court_name,
            date: meta.date,
            slot: meta.slot,
            name: meta.name,
            phone: meta.phone,
            reservaId: reservaId,
            paymentId: paymentId
          });
        }

        await ordenarHoja(sheets, sheetName);
      } else {
        console.log("Duplicado ignorado");
      }

      // Borrar de Temp en cualquier caso
      await borrarDeTemp(sheets, extRef);

    } else if (estado === "rejected" || estado === "cancelled") {
      await borrarDeTemp(sheets, extRef);
      console.log("Borrado de Temp por rechazo/cancelación");
    } else {
      console.log("Estado pendiente:", estado);
    }

    return res.status(200).json({ received: true });
  } catch (err) {
    console.error("Error webhook:", err);
    return res.status(500).json({ error: err.message });
  }
};

async function crearEventoCalendar(auth, info) {
  try {
    const calendar = google.calendar({ version: "v3", auth });

    const slotLimpio = String(info.slot).replace(" hs", "").trim();
    const partesHora = slotLimpio.split(":");
    const horaInicio = parseInt(partesHora[0], 10);
    const minutoInicio = partesHora[1] || "00";

    const horaFin = horaInicio + 1;

    // Offset explícito para Argentina (-03:00)
    const startStr = `${info.date}T${String(horaInicio).padStart(2, "0")}:${minutoInicio}:00-03:00`;
    const endStr   = `${info.date}T${String(horaFin).padStart(2, "0")}:${minutoInicio}:00-03:00`;

    await calendar.events.insert({
      calendarId: GOOGLE_CALENDAR_ID,
      requestBody: {
        summary: `⚽ ${info.courtName} - ${info.name}`,
        description: `Reserva: ${info.reservaId}\nCliente: ${info.name}\nTeléfono: ${info.phone}\nPago MP: #${info.paymentId}`,
        start: {
          dateTime: startStr,
          timeZone: "America/Argentina/Buenos_Aires"
        },
        end: {
          dateTime: endStr,
          timeZone: "America/Argentina/Buenos_Aires"
        }
      }
    });

    console.log("Evento creado en Google Calendar en horario correcto:", startStr);
  } catch (err) {
    console.error("Error creando evento en Google Calendar:", err);
  }
}

async function buscarConfirmada(sheets, sheetName, extRef) {
  const parts   = extRef.split("|");
  const courtId = parts[0] || "";
  const date    = parts[1] || "";
  const slot    = parts[2] || "";
  const r       = await sheets.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_ID, range: sheetName + "!A2:K1000"
  });
  for (const row of r.data.values || []) {
    if (String(row[3]||"") === courtId && (row[4]||"") === date &&
        (row[5]||"") === slot && (row[9]||"") === "CONFIRMADA") return true;
  }
  return false;
}

async function borrarDeTemp(sheets, extRef) {
  try {
    const r    = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID, range: TEMP_SHEET + "!A2:E1000"
    });
    const rows = r.data.values || [];
    const aBorrar = [];
    for (let i = 0; i < rows.length; i++) {
      if ((rows[i][0] || "") === extRef) aBorrar.push(i + 1);
    }
    if (aBorrar.length === 0) return;

    const meta    = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
    const hoja    = meta.data.sheets.find(function(s) { return s.properties.title === TEMP_SHEET; });
    if (!hoja) return;
    const sheetId = hoja.properties.sheetId;
    const sorted  = aBorrar.slice().sort(function(a, b) { return b - a; });
    const reqs    = sorted.map(function(idx) {
      return { deleteDimension: { range: { sheetId, dimension: "ROWS", startIndex: idx, endIndex: idx + 1 } } };
    });
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: GOOGLE_SHEET_ID, requestBody: { requests: reqs }
    });
    console.log("Borrado de Temp:", extRef);
  } catch (e) {
    console.error("Error borrando de Temp:", e);
  }
}

async function ordenarHoja(sheets, sheetName) {
  const r    = await sheets.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_ID, range: sheetName + "!A2:K1000"
  });
  const rows = r.data.values || [];
  if (rows.length === 0) return;

  const sorted = rows.slice().sort(function(a, b) {
    const dA = a[4]||"", dB = b[4]||"";
    if (dA !== dB) return dA < dB ? -1 : 1;
    const sA = a[5]||"", sB = b[5]||"";
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

function getGoogleAuth() {
  const key = GOOGLE_SA_KEY.replace(/\\n/g, "\n").replace(/\r\n/g, "\n").trim();
  return new google.auth.GoogleAuth({
    credentials: { client_email: GOOGLE_SA_EMAIL, private_key: key },
    scopes: [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/calendar.events"
    ]
  });
}
