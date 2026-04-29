// api/categories.js
// GET    /api/categories?userId=xxx        → 取得該用戶的自訂分類
// POST   /api/categories                  → 新增分類 { userId, name }
// DELETE /api/categories?userId=xxx&name=xxx → 刪除分類

const { google } = require("googleapis");

const GOOGLE_CREDENTIALS = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
const SHEETS_ID = process.env.GOOGLE_SHEETS_ID;
const SHEET_NAME = "收藏分類";

const DEFAULT_CATEGORIES = ["個人", "工作", "家庭", "未分類"];

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
      requestBody: {
        requests: [{ addSheet: { properties: { title: SHEET_NAME } } }]
      }
    });
    // 加入 header
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEETS_ID,
      range: SHEET_NAME + "!A1:B1",
      valueInputOption: "RAW",
      requestBody: { values: [["userId", "name"]] }
    });
  }
}

module.exports = async function(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, DELETE, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") return res.status(200).end();

  const userId = (req.query && req.query.userId) || (req.body && req.body.userId);
  if (!userId) return res.status(400).json({ error: "Missing userId" });

  try {
    const sheets = await getSheetsClient();
    await ensureSheet(sheets);

    // ── GET ──────────────────────────────────────────────
    if (req.method === "GET") {
      const result = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEETS_ID,
        range: SHEET_NAME + "!A:B",
      });
      const rows = result.data.values || [];
      const custom = rows.slice(1)
        .filter(function(r) { return String(r[0] || "") === userId; })
        .map(function(r) { return String(r[1] || ""); })
        .filter(Boolean);

      // 合併預設 + 自訂（去重）
      const all = [...DEFAULT_CATEGORIES];
      custom.forEach(function(c) {
        if (!all.includes(c)) all.push(c);
      });

      return res.status(200).json({ success: true, categories: all, custom: custom });
    }

    // ── POST ─────────────────────────────────────────────
    if (req.method === "POST") {
      const name = req.body && req.body.name;
      if (!name || !name.trim()) return res.status(400).json({ error: "Missing name" });
      const trimmed = name.trim();

      // 檢查是否重複
      const result = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEETS_ID,
        range: SHEET_NAME + "!A:B",
      });
      const rows = result.data.values || [];
      const exists = rows.slice(1).some(function(r) {
        return String(r[0] || "") === userId && String(r[1] || "") === trimmed;
      });
      if (exists || DEFAULT_CATEGORIES.includes(trimmed)) {
        return res.status(400).json({ error: "Category already exists" });
      }

      await sheets.spreadsheets.values.append({
        spreadsheetId: SHEETS_ID,
        range: SHEET_NAME + "!A:B",
        valueInputOption: "RAW",
        requestBody: { values: [[userId, trimmed]] }
      });

      return res.status(200).json({ success: true });
    }

    // ── DELETE ───────────────────────────────────────────
    if (req.method === "DELETE") {
      const name = req.query.name;
      if (!name) return res.status(400).json({ error: "Missing name" });

      if (DEFAULT_CATEGORIES.includes(name)) {
        return res.status(400).json({ error: "Cannot delete default category" });
      }

      const result = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEETS_ID,
        range: SHEET_NAME + "!A:B",
      });
      const rows = result.data.values || [];

      let targetRow = -1;
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0] || "") === userId && String(rows[i][1] || "") === name) {
          targetRow = i + 1;
          break;
        }
      }
      if (targetRow === -1) return res.status(404).json({ error: "Category not found" });

      const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SHEETS_ID });
      const sheet = spreadsheet.data.sheets.find(function(s) {
        return s.properties.title === SHEET_NAME;
      });
      const sheetId = sheet.properties.sheetId;

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SHEETS_ID,
        requestBody: {
          requests: [{
            deleteDimension: {
              range: {
                sheetId: sheetId,
                dimension: "ROWS",
                startIndex: targetRow - 1,
                endIndex: targetRow,
              }
            }
          }]
        }
      });

      return res.status(200).json({ success: true });
    }

    return res.status(405).json({ error: "Method not allowed" });

  } catch (err) {
    console.error("Categories API error:", err);
    return res.status(500).json({ error: err.message });
  }
};
