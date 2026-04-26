// api/todos.js
const { google } = require("googleapis");
const CREDS = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
const SID = process.env.GOOGLE_SHEETS_ID;
const SNAME = "\u5F85\u8FA6\u6E05\u55AE";
async function getSheets() {
  const auth = new google.auth.GoogleAuth({ credentials: JSON.parse(CREDS), scopes: ["https://www.googleapis.com/auth/spreadsheets"] });
  return google.sheets({ version: "v4", auth: await auth.getClient() });
}
function genId(p) {
  const ts = new Date().toISOString().replace(/[:\-T.Z]/g,"").substring(0,14);
  return p + ts + Math.floor(Math.random()*1000).toString().padStart(3,"0");
}
function fmtDt(d) {
  return new Date(d).toLocaleString("zh-TW", { timeZone:"Asia/Taipei", year:"numeric", month:"2-digit", day:"2-digit", hour:"2-digit", minute:"2-digit" });
}
module.exports = async function(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,PATCH,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  try {
    const sheets = await getSheets();
    if (req.method === "GET") {
      const uid = req.query && req.query.userId;
      if (!uid) return res.status(400).json({ error: "Missing userId" });
      const r = await sheets.spreadsheets.values.get({ spreadsheetId: SID, range: SNAME + "!A:H" });
      const rows = (r.data.values || []).slice(1);
      const todos = rows
        .filter(function(r) { return String(r[5]||"")===uid && String(r[6]||"")==="todo" && String(r[2]||"pending")!=="deleted"; })
        .reverse().slice(0,50)
        .map(function(r) { return { todoId:String(r[0]||""), content:String(r[1]||""), status:String(r[2]||"pending"), createdAt:String(r[3]||""), remindAt:String(r[4]||""), userId:String(r[5]||""), googleEventId:String(r[7]||"") }; });
      return res.status(200).json({ success: true, todos: todos });
    }
    if (req.method === "POST") {
      const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
      const uid = body && body.userId;
      const content = body && body.content;
      const remindAt = body && body.remindAt ? String(body.remindAt).replace("T"," ") : "";
      if (!uid || !content) return res.status(400).json({ error: "Missing fields" });
      const id = genId("TD");
      const now = fmtDt(new Date());
      await sheets.spreadsheets.values.append({ spreadsheetId: SID, range: SNAME + "!A:H", valueInputOption: "RAW", requestBody: { values: [[id, content, "pending", now, remindAt, uid, "todo", ""]] } });
      return res.status(200).json({ success: true, todoId: id });
    }
    if (req.method === "PATCH") {
      const urlParts = (req.url || "").split("/");
      const todoId = urlParts[urlParts.length - 1].split("?")[0];
      const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
      const status = body && body.status;
      if (!todoId || !status) return res.status(400).json({ error: "Missing fields" });
      const r = await sheets.spreadsheets.values.get({ spreadsheetId: SID, range: SNAME + "!A:H" });
      const rows = r.data.values || [];
      const idx = rows.findIndex(function(r) { return String(r[0]||"") === todoId; });
      if (idx === -1) return res.status(404).json({ error: "Not found" });
      await sheets.spreadsheets.values.update({ spreadsheetId: SID, range: SNAME + "!C" + (idx + 1), valueInputOption: "RAW", requestBody: { values: [[status]] } });
      return res.status(200).json({ success: true });
    }
    return res.status(405).json({ error: "Method not allowed" });
  } catch(e) {
    console.error(e);
    return res.status(500).json({ error: e.message });
  }
};
