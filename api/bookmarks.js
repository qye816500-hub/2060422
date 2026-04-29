// api/bookmarks.js
// GET  /api/bookmarks?userId=xxx
// DELETE /api/bookmarks?userId=xxx&id=xxx

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
  res.setHeader("Access-Control-Allow-Methods", "GET, DELETE, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  const userId = req.query && req.query.userId;
  if (!userId) {
    return res.status(400).json({ error: "Missing userId" });
  }

  // ── GET ──────────────────────────────────────────────────
  if (req.method === "GET") {
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
            category: String(r[6] || "未分類"),
            userId: String(r[7] || ""),
            createdAt: String(r[8] || ""),
            updatedAt: String(r[9] || ""),
          };
        });

      return res.status(200).json({ success: true, bookmarks: bookmarks });
    } catch (err) {
      console.error("Bookmarks GET error:", err);
      return res.status(500).json({ error: err.message });
    }
  }

  // ── DELETE ───────────────────────────────────────────────
  if (req.method === "DELETE") {
    const bookmarkId = req.query && req.query.id;
    if (!bookmarkId) {
      return res.status(400).json({ error: "Missing id" });
    }

    try {
      const sheets = await getSheetsClient();

      // 找出要刪除的那列
      const result = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEETS_ID,
        range: SHEET_NAME + "!A:K",
      });
      const rows = result.data.values || [];

      // 找到符合 id 且 userId 相符的列（+1 因為 header，+1 因為 Sheets 從 1 開始）
      let targetRowIndex = -1;
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0] || "") === bookmarkId && String(rows[i][7] || "") === userId) {
          targetRowIndex = i + 1; // Sheets row number (1-based)
          break;
        }
      }

      if (targetRowIndex === -1) {
        return res.status(404).json({ error: "Bookmark not found" });
      }

      // 取得 spreadsheetId 對應的 sheetId
      const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SHEETS_ID });
      const sheet = spreadsheet.data.sheets.find(function(s) {
        return s.properties.title === SHEET_NAME;
      });
      if (!sheet) {
        return res.status(500).json({ error: "Sheet not found" });
      }
      const sheetId = sheet.properties.sheetId;

      // 刪除那列
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SHEETS_ID,
        requestBody: {
          requests: [{
            deleteDimension: {
              range: {
                sheetId: sheetId,
                dimension: "ROWS",
                startIndex: targetRowIndex - 1,
                endIndex: targetRowIndex,
              }
            }
          }]
        }
      });

      return res.status(200).json({ success: true });
    } catch (err) {
      console.error("Bookmarks DELETE error:", err);
      return res.status(500).json({ error: err.message });
    }
  }

  return res.status(405).json({ error: "Method not allowed" });
};
