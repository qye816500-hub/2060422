// api/bookmarks.js
// GET /api/bookmarks?userId=xxx

const { google } = require("googleapis");

const GOOGLE_CREDENTIALS = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
const SHEETS_ID = process.env.GOOGLE_SHEETS_ID;
const SHEET_NAME = "Threads\u6536\u85CF";

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
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  if (req.method !== "GET") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const userId = req.query && req.query.userId;
  if (!userId) {
    return res.status(400).json({ error: "Missing userId" });
  }

  try {
    const sheets = await getSheetsClient();
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEETS_ID,
      range: SHEET_NAME + "!A:K",
    });

    const rows = result.data.values || [];
    const bookmarks = rows.slice(1)
      .filter(function(r) { return String(r[7] || "") === userId; })
      .reverse()
      .slice(0, 50)
      .map(function(r) {
        return {
          id: String(r[0] || ""),
          url: String(r[1] || ""),
          platform: String(r[2] || "Link"),
          title: String(r[3] || ""),
          note: String(r[4] || ""),
          tags: String(r[5] || ""),
          category: String(r[6] || "\u672A\u5206\u985E"),
          userId: String(r[7] || ""),
          createdAt: String(r[8] || ""),
          updatedAt: String(r[9] || ""),
        };
      });

    return res.status(200).json({ success: true, bookmarks: bookmarks });
  } catch (err) {
    console.error("Bookmarks API error:", err);
    return res.status(500).json({ error: err.message });
  }
};
