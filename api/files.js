// api/files.js
// GET /api/files?userId=xxx

const { google } = require(“googleapis”);

const GOOGLE_CREDENTIALS = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
const SHEETS_ID = process.env.GOOGLE_SHEETS_ID;
const SHEET_NAME = “\u6A94\u6848\u5099\u4EFD”;

async function getSheetsClient() {
const credentials = JSON.parse(GOOGLE_CREDENTIALS);
const auth = new google.auth.GoogleAuth({
credentials,
scopes: [“https://www.googleapis.com/auth/spreadsheets”],
});
const authClient = await auth.getClient();
return google.sheets({ version: “v4”, auth: authClient });
}

module.exports = async function(req, res) {
res.setHeader(“Access-Control-Allow-Origin”, “*”);
res.setHeader(“Access-Control-Allow-Methods”, “GET, OPTIONS”);
res.setHeader(“Access-Control-Allow-Headers”, “Content-Type”);

if (req.method === “OPTIONS”) {
return res.status(200).end();
}

if (req.method !== “GET”) {
return res.status(405).json({ error: “Method not allowed” });
}

const userId = req.query && req.query.userId;
if (!userId) {
return res.status(400).json({ error: “Missing userId” });
}

try {
const sheets = await getSheetsClient();
const result = await sheets.spreadsheets.values.get({
spreadsheetId: SHEETS_ID,
range: SHEET_NAME + “!A:L”,
});

```
const rows = result.data.values || [];
const files = rows.slice(1)
  .filter(function(r) { return String(r[9] || "") === userId; })
  .reverse()
  .slice(0, 30)
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
```

} catch (err) {
console.error(“Files API error:”, err);
return res.status(500).json({ error: err.message });
}
};
