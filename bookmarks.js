// api/bookmarks.js
const { google } = require("googleapis");
const CREDS = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
const SID = process.env.GOOGLE_SHEETS_ID;
const SNAME = "Threads\u6536\u85CF";
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
    const r = await sheets.spreadsheets.values.get({ spreadsheetId: SID, range: SNAME + "!A:K" });
    const rows = (r.data.values || []).slice(1);
    const bookmarks = rows
      .filter(function(r) { return String(r[7] || "") === uid; })
      .reverse().slice(0, 50)
      .map(function(r) {
        return { id: String(r[0]||""), url: String(r[1]||""), platform: String(r[2]||"Link"), title: String(r[3]||""), note: String(r[4]||""), tags: String(r[5]||""), category: String(r[6]||""), userId: String(r[7]||""), createdAt: String(r[8]||""), updatedAt: String(r[9]||"") };
      });
    return res.status(200).json({ success: true, bookmarks: bookmarks });
  } catch(e) {
    console.error(e);
    return res.status(500).json({ error: e.message });
  }
};
