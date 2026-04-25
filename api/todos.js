// api/todos.js
// GET /api/todos?userId=xxx
// POST /api/todos { userId, content, remindAt }
// PATCH /api/todos/:id { status }

const { google } = require(“googleapis”);

const GOOGLE_CREDENTIALS = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
const SHEETS_ID = process.env.GOOGLE_SHEETS_ID;
const SHEET_NAME = “\u5F85\u8FA6\u6E05\u55AE”;

async function getSheetsClient() {
const credentials = JSON.parse(GOOGLE_CREDENTIALS);
const auth = new google.auth.GoogleAuth({
credentials,
scopes: [“https://www.googleapis.com/auth/spreadsheets”],
});
const authClient = await auth.getClient();
return google.sheets({ version: “v4”, auth: authClient });
}

function generateId(prefix) {
const ts = new Date().toISOString().replace(/[:-T.Z]/g, “”).substring(0, 14);
const rnd = Math.floor(Math.random() * 1000).toString().padStart(3, “0”);
return prefix + ts + rnd;
}

function formatDateTime(date) {
return new Date(date).toLocaleString(“zh-TW”, {
timeZone: “Asia/Taipei”,
year: “numeric”, month: “2-digit”, day: “2-digit”,
hour: “2-digit”, minute: “2-digit”,
});
}

module.exports = async function(req, res) {
res.setHeader(“Access-Control-Allow-Origin”, “*”);
res.setHeader(“Access-Control-Allow-Methods”, “GET, POST, PATCH, OPTIONS”);
res.setHeader(“Access-Control-Allow-Headers”, “Content-Type”);

if (req.method === “OPTIONS”) {
return res.status(200).end();
}

try {
const sheets = await getSheetsClient();

```
// GET - list todos
if (req.method === "GET") {
  const userId = req.query && req.query.userId;
  if (!userId) return res.status(400).json({ error: "Missing userId" });

  const result = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEETS_ID,
    range: SHEET_NAME + "!A:H",
  });

  const rows = result.data.values || [];
  const todos = rows.slice(1)
    .filter(function(r) {
      return String(r[5] || "") === userId &&
             String(r[6] || "") === "todo" &&
             String(r[2] || "pending") !== "deleted";
    })
    .reverse()
    .slice(0, 50)
    .map(function(r) {
      return {
        todoId: String(r[0] || ""),
        content: String(r[1] || ""),
        status: String(r[2] || "pending"),
        createdAt: String(r[3] || ""),
        remindAt: String(r[4] || ""),
        userId: String(r[5] || ""),
        googleEventId: String(r[7] || ""),
      };
    });

  return res.status(200).json({ success: true, todos: todos });
}

// POST - create todo
if (req.method === "POST") {
  const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
  const userId = body && body.userId;
  const content = body && body.content;
  const remindAt = body && body.remindAt ? body.remindAt.replace("T", " ") : "";

  if (!userId || !content) {
    return res.status(400).json({ error: "Missing userId or content" });
  }

  const id = generateId("TD");
  const now = formatDateTime(new Date());
  const row = [id, content, "pending", now, remindAt, userId, "todo", ""];

  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEETS_ID,
    range: SHEET_NAME + "!A:H",
    valueInputOption: "RAW",
    requestBody: { values: [row] },
  });

  return res.status(200).json({ success: true, todoId: id });
}

// PATCH - update todo status
if (req.method === "PATCH") {
  const url = req.url || "";
  const parts = url.split("/");
  const todoId = parts[parts.length - 1].split("?")[0];

  const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
  const status = body && body.status;

  if (!todoId || !status) {
    return res.status(400).json({ error: "Missing todoId or status" });
  }

  const result = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEETS_ID,
    range: SHEET_NAME + "!A:H",
  });

  const rows = result.data.values || [];
  const rowIndex = rows.findIndex(function(r) { return String(r[0] || "") === todoId; });

  if (rowIndex === -1) {
    return res.status(404).json({ error: "Todo not found" });
  }

  const rowNum = rowIndex + 1;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEETS_ID,
    range: SHEET_NAME + "!C" + rowNum,
    valueInputOption: "RAW",
    requestBody: { values: [[status]] },
  });

  return res.status(200).json({ success: true });
}

return res.status(405).json({ error: "Method not allowed" });
```

} catch (err) {
console.error(“Todos API error:”, err);
return res.status(500).json({ error: err.message });
}
};
