// api/files.js
// GET    /api/files?userId=xxx
// DELETE /api/files?userId=xxx&id=xxx

const { google } = require("googleapis");
const CREDS = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
const SID = process.env.GOOGLE_SHEETS_ID;
const SNAME = "檔案備份";

async function getSheets() {
  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(CREDS),
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  return google.sheets({ version: "v4", auth: await auth.getClient() });
}

module.exports = async function(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, DELETE, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();

  const uid = req.query && req.query.userId;
  if (!uid) return res.status(400).json({ error: "Missing userId" });

  // ── GET ──────────────────────────────────────────────────
  if (req.method === "GET") {
    try {
      const sheets = await getSheets();
      const r = await sheets.spreadsheets.values.get({ spreadsheetId: SID, range: SNAME + "!A:L" });
      const rows = (r.data.values || []).slice(1);
      const files = rows
        .filter(function(r) { return String(r[9] || "") === uid; })
        .reverse().slice(0, 30)
        .map(function(r) {
          return {
            id: String(r[0] || ""),
            fileName: String(r[1] || ""),
            fileType: String(r[2] || ""),
            mimeType: String(r[3] || ""),
            size: Number(r[4] || 0),
            driveFileId: String(r[5] || ""),
            driveUrl: String(r[6] || ""),
            userId: String(r[9] || ""),
            lineMessageId: String(r[10] || ""),
            createdAt: String(r[11] || ""),
          };
        });
      return res.status(200).json({ success: true, files: files });
    } catch(e) {
      console.error(e);
      return res.status(500).json({ error: e.message });
    }
  }

  // ── DELETE ───────────────────────────────────────────────
  if (req.method === "DELETE") {
    const fileId = req.query && req.query.id;
    if (!fileId) return res.status(400).json({ error: "Missing id" });

    try {
      const sheets = await getSheets();
      const r = await sheets.spreadsheets.values.get({ spreadsheetId: SID, range: SNAME + "!A:L" });
      const rows = r.data.values || [];

      let targetRow = -1;
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0] || "") === fileId && String(rows[i][9] || "") === uid) {
          targetRow = i + 1;
          break;
        }
      }
      if (targetRow === -1) return res.status(404).json({ error: "File not found" });

      const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SID });
      const sheet = spreadsheet.data.sheets.find(function(s) { return s.properties.title === SNAME; });
      if (!sheet) return res.status(500).json({ error: "Sheet not found" });

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SID,
        requestBody: {
          requests: [{
            deleteDimension: {
              range: {
                sheetId: sheet.properties.sheetId,
                dimension: "ROWS",
                startIndex: targetRow - 1,
                endIndex: targetRow,
              }
            }
          }]
        }
      });

      return res.status(200).json({ success: true });
    } catch(e) {
      console.error(e);
      return res.status(500).json({ error: e.message });
    }
  }

  return res.status(405).json({ error: "Method not allowed" });
};
