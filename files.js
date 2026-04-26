// api/files.js
const { google } = require("googleapis");
const CREDS = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
const SID = process.env.GOOGLE_SHEETS_ID;
const SNAME = "\u6A94\u6848\u5099\u4EFD";
async function getSheets() {
  const auth = new google.auth.GoogleAuth({ credentials: JSON.parse(CREDS), scopes: ["https://www.googleapis.com/auth/spreadsheets"] });
  return google.sheets({ version: "v4", auth: await auth.getClient() });
}
module.exports = async function(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "GET") return res.status(405).json({ error: "Method not allowed" });
  const uid = req.query && req.query.userId;
  if (!uid) return res.status(400).json({ error: "Missing userId" });
  try {
    const sheets = await getSheets();
    const r = await sheets.spreadsheets.values.get({ spreadsheetId: SID, range: SNAME + "!A:L" });
    const rows = (r.data.values || []).slice(1);
    const files = rows
      .filter(function(r) { return String(r[9]||"") === uid; })
      .reverse().slice(0,30)
      .map(function(r) { return { id:String(r[0]||""), fileName:String(r[1]||""), fileType:String(r[2]||""), mimeType:String(r[3]||""), size:Number(r[4]||0), driveFileId:String(r[5]||""), driveUrl:String(r[6]||""), userId:String(r[9]||""), lineMessageId:String(r[10]||""), createdAt:String(r[11]||"") }; });
    return res.status(200).json({ success: true, files: files });
  } catch(e) {
    console.error(e);
    return res.status(500).json({ error: e.message });
  }
};
