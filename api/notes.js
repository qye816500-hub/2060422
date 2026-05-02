// api/notes.js
// GET    /api/notes?userId=xxx
// POST   /api/notes { userId, content, tags }
// PATCH  /api/notes/:id { content, tags }
// DELETE /api/notes?userId=xxx&id=xxx

const { google } = require("googleapis");

const GOOGLE_CREDENTIALS = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
const SHEETS_ID = process.env.GOOGLE_SHEETS_ID;
const SHEET_NAME = "收藏筆記";

async function getSheetsClient() {
  const credentials = JSON.parse(GOOGLE_CREDENTIALS);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  const authClient = await auth.getClient();
  return google.sheets({ version: "v4", auth: authClient });
}

async function ensureSheet(sheets) {
  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SHEETS_ID });
  const exists = spreadsheet.data.sheets.some(function(s) {
    return s.properties.title === SHEET_NAME;
  });
  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEETS_ID,
      requestBody: { requests: [{ addSheet: { properties: { title: SHEET_NAME } } }] }
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEETS_ID,
      range: SHEET_NAME + "!A1:F1",
      valueInputOption: "RAW",
      requestBody: { values: [["id", "content", "tags", "userId", "createdAt", "updatedAt"]] }
    });
  }
}

function generateId() {
  const ts = new Date().toISOString().replace(/[:\-T.Z]/g, "").substring(0, 14);
  const rnd = Math.floor(Math.random() * 1000).toString().padStart(3, "0");
  return "N" + ts + rnd;
}

function formatDateTime(date) {
  return new Date(date).toLocaleString("zh-TW", {
    timeZone: "Asia/Taipei",
    year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit",
  });
}

module.exports = async function(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, PATCH, DELETE, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();

  try {
    const sheets = await getSheetsClient();
    await ensureSheet(sheets);

    // ── GET ──────────────────────────────────────────────────
    if (req.method === "GET") {
      const userId = req.query && req.query.userId;
      if (!userId) return res.status(400).json({ error: "Missing userId" });

      const result = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEETS_ID,
        range: SHEET_NAME + "!A:F",
      });
      const rows = result.data.values || [];
      const notes = rows.slice(1)
        .filter(function(r) { return String(r[3] || "") === userId; })
        .reverse()
        .slice(0, 100)
        .map(function(r) {
          return {
            id: String(r[0] || ""),
            content: String(r[1] || ""),
            tags: String(r[2] || ""),
            userId: String(r[3] || ""),
            createdAt: String(r[4] || ""),
            updatedAt: String(r[5] || ""),
          };
        });
      return res.status(200).json({ success: true, notes });
    }

    // ── POST ─────────────────────────────────────────────────
    if (req.method === "POST") {
      const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
      const { userId, content, tags } = body || {};
      if (!userId || !content) return res.status(400).json({ error: "Missing userId or content" });

      const id = generateId();
      const now = formatDateTime(new Date());
      await sheets.spreadsheets.values.append({
        spreadsheetId: SHEETS_ID,
        range: SHEET_NAME + "!A:F",
        valueInputOption: "RAW",
        requestBody: { values: [[id, content, tags || "", userId, now, now]] },
      });
      return res.status(200).json({ success: true, id });
    }

    // ── PATCH ────────────────────────────────────────────────
    if (req.method === "PATCH") {
      const url = req.url || "";
      const noteId = url.split("/").pop().split("?")[0];
      const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
      const { content, tags } = body || {};

      if (!noteId) return res.status(400).json({ error: "Missing id" });

      const result = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEETS_ID, range: SHEET_NAME + "!A:F",
      });
      const rows = result.data.values || [];
      const idx = rows.findIndex(function(r) { return String(r[0] || "") === noteId; });
      if (idx === -1) return res.status(404).json({ error: "Note not found" });

      const rowNum = idx + 1;
      const now = formatDateTime(new Date());
      const updates = [{ range: SHEET_NAME + "!F" + rowNum, values: [[now]] }];
      if (content !== undefined) updates.push({ range: SHEET_NAME + "!B" + rowNum, values: [[content]] });
      if (tags !== undefined) updates.push({ range: SHEET_NAME + "!C" + rowNum, values: [[tags]] });

      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SHEETS_ID,
        requestBody: { valueInputOption: "RAW", data: updates },
      });
      return res.status(200).json({ success: true });
    }

    // ── DELETE ───────────────────────────────────────────────
    if (req.method === "DELETE") {
      const noteId = req.query && req.query.id;
      const userId = req.query && req.query.userId;
      if (!noteId || !userId) return res.status(400).json({ error: "Missing id or userId" });

      const result = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEETS_ID, range: SHEET_NAME + "!A:F",
      });
      const rows = result.data.values || [];
      let targetRow = -1;
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0] || "") === noteId && String(rows[i][3] || "") === userId) {
          targetRow = i + 1; break;
        }
      }
      if (targetRow === -1) return res.status(404).json({ error: "Note not found" });

      const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SHEETS_ID });
      const sheet = spreadsheet.data.sheets.find(function(s) { return s.properties.title === SHEET_NAME; });

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SHEETS_ID,
        requestBody: {
          requests: [{
            deleteDimension: {
              range: { sheetId: sheet.properties.sheetId, dimension: "ROWS", startIndex: targetRow - 1, endIndex: targetRow }
            }
          }]
        }
      });
      return res.status(200).json({ success: true });
    }

    return res.status(405).json({ error: "Method not allowed" });
  } catch (err) {
    console.error("Notes API error:", err);
    return res.status(500).json({ error: err.message });
  }
};
