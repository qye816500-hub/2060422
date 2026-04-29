// api/bookmarks/update.js
// PATCH /api/bookmarks/update  { userId, id, category }
// Updates the category of a specific bookmark

const { google } = require("googleapis");

const GOOGLE_CREDENTIALS = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
const SHEETS_ID = process.env.GOOGLE_SHEETS_ID;
const SHEET_NAME = "Threads收藏";

async function getSheetsClient() {
  const credentials = JSON.parse(GOOGLE_CREDENTIALS);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  const authClient = await auth.getClient();
  return google.sheets({ version: "v4", auth: authClient });
}

module.exports = async function(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "PATCH, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "PATCH") return res.status(405).json({ error: "Method not allowed" });

  const { userId, id, category } = req.body || {};
  if (!userId || !id || !category) {
    return res.status(400).json({ error: "Missing userId, id or category" });
  }

  try {
    const sheets = await getSheetsClient();
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEETS_ID,
      range: SHEET_NAME + "!A:K",
    });
    const rows = result.data.values || [];

    let targetRow = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0] || "") === id && String(rows[i][7] || "") === userId) {
        targetRow = i + 1; // 1-based
        break;
      }
    }

    if (targetRow === -1) {
      return res.status(404).json({ error: "Bookmark not found" });
    }

    const now = new Date().toLocaleString("zh-TW", { timeZone: "Asia/Taipei" });
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SHEETS_ID,
      requestBody: {
        valueInputOption: "RAW",
        data: [
          { range: SHEET_NAME + "!G" + targetRow, values: [[category]] },
          { range: SHEET_NAME + "!J" + targetRow, values: [[now]] },
        ],
      },
    });

    return res.status(200).json({ success: true });
  } catch (err) {
    console.error("Bookmark update error:", err);
    return res.status(500).json({ error: err.message });
  }
};
