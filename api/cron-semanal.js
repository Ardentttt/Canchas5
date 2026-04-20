// api/cron-semanal.js
// Crea una nueva hoja "Semana X" cada domingo a las 3am (Argentina = UTC-3).
// Mantiene un máximo de 4 semanas; si hay 5, borra la más vieja.
// Disparado por Vercel Cron Jobs (ver vercel.json).

const { google } = require("googleapis");

const GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID;
const GOOGLE_SA_EMAIL = process.env.GOOGLE_SA_EMAIL;
const GOOGLE_SA_KEY   = process.env.GOOGLE_SA_KEY;

const HEADERS = [
  "ID", "Fecha Reserva", "Cancha", "Cancha ID",
  "Fecha Turno", "Horario", "Nombre", "Teléfono",
  "Precio Total", "Seña 50%", "Estado", "Notas", "Timestamp"
];

const MAX_SEMANAS = 4;

module.exports = async function handler(req, res) {
  // Solo permitir GET (Vercel Cron usa GET)
  if (req.method !== "GET") return res.status(405).end();

  // Verificar que sea domingo 3am en Argentina (UTC-3)
  // Vercel Cron garantiza el horario, pero validamos igual
  const ahora = new Date();
  const horaAR = (ahora.getUTCHours() - 3 + 24) % 24;
  const diaUTC = ahora.getUTCDay(); // 0 = domingo
  const diaAR  = horaAR < 0 ? (diaUTC + 6) % 7 : diaUTC; // ajuste por timezone

  console.log("Cron ejecutado. Hora AR:", horaAR, "Día AR:", diaAR);

  try {
    const sheets = await getSheetsClient();
    const meta   = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
    const hojas  = meta.data.sheets;
    const titles = hojas.map(function(s) { return s.properties.title; });

    // Encontrar semanas existentes
    const semanas = titles
      .filter(function(t) { return t.startsWith("Semana"); })
      .sort(function(a, b) {
        return numDeSemana(a) - numDeSemana(b);
      });

    // Determinar número de la nueva semana
    const ultimoNum = semanas.length > 0
      ? numDeSemana(semanas[semanas.length - 1])
      : 0;
    const nuevoNum  = ultimoNum + 1;
    const nuevoNombre = "Semana " + nuevoNum;

    // Crear la nueva hoja
    const addRequest = {
      addSheet: {
        properties: {
          title: nuevoNombre,
          gridProperties: { rowCount: 1000, columnCount: 13 }
        }
      }
    };

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: GOOGLE_SHEET_ID,
      requestBody: { requests: [addRequest] }
    });

    // Agregar headers a la nueva hoja
    await sheets.spreadsheets.values.update({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: nuevoNombre + "!A1:M1",
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [HEADERS] }
    });

    // Formatear header (negrita + fondo oscuro)
    const metaNew   = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
    const hojaNew   = metaNew.data.sheets.find(function(s) { return s.properties.title === nuevoNombre; });
    const sheetIdNew = hojaNew.properties.sheetId;

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: GOOGLE_SHEET_ID,
      requestBody: {
        requests: [
          {
            repeatCell: {
              range: { sheetId: sheetIdNew, startRowIndex: 0, endRowIndex: 1 },
              cell: {
                userEnteredFormat: {
                  backgroundColor: { red: 0.067, green: 0.122, blue: 0.082 },
                  textFormat: { bold: true, foregroundColor: { red: 0.0, green: 0.91, blue: 0.478 } }
                }
              },
              fields: "userEnteredFormat(backgroundColor,textFormat)"
            }
          },
          // Congelar primera fila
          {
            updateSheetProperties: {
              properties: {
                sheetId: sheetIdNew,
                gridProperties: { frozenRowCount: 1 }
              },
              fields: "gridProperties.frozenRowCount"
            }
          }
        ]
      }
    });

    console.log("Hoja creada:", nuevoNombre);

    // Si hay más de MAX_SEMANAS, borrar la más vieja
    let borrada = null;
    if (semanas.length >= MAX_SEMANAS) {
      const masVieja    = semanas[0];
      const sheetVieja  = hojas.find(function(s) { return s.properties.title === masVieja; });
      if (sheetVieja) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: GOOGLE_SHEET_ID,
          requestBody: {
            requests: [{
              deleteSheet: { sheetId: sheetVieja.properties.sheetId }
            }]
          }
        });
        borrada = masVieja;
        console.log("Hoja borrada:", masVieja);
      }
    }

    return res.status(200).json({
      ok: true,
      creada: nuevoNombre,
      borrada: borrada || "ninguna",
      semanasTotales: semanas.length + 1 - (borrada ? 1 : 0)
    });

  } catch (err) {
    console.error("Error en cron-semanal:", err);
    return res.status(500).json({ error: err.message });
  }
};

function numDeSemana(titulo) {
  var match = titulo.match(/\d+/);
  return match ? parseInt(match[0]) : 0;
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
