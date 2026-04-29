// api/thumb.js
// GET /api/thumb?fileId=xxx
// Proxies Google Drive thumbnail using Service Account auth

const { google } = require("googleapis");

const CREDS = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;

async function getDriveClient() {
  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(CREDS),
    scopes: ["https://www.googleapis.com/auth/drive.readonly"],
  });
  return google.drive({ version: "v3", auth: await auth.getClient() });
}

module.exports = async function(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  if (req.method === "OPTIONS") return res.status(200).end();

  const fileId = req.query && req.query.fileId;
  if (!fileId) return res.status(400).json({ error: "Missing fileId" });

  try {
    const drive = await getDriveClient();

    // 先取得檔案 metadata，拿縮圖 URL
    const meta = await drive.files.get({
      fileId: fileId,
      fields: "id,name,mimeType,thumbnailLink,hasThumbnail",
    });

    const thumbnailLink = meta.data.thumbnailLink;

    if (!thumbnailLink) {
      // 沒有縮圖，回傳 404
      return res.status(404).json({ error: "No thumbnail available" });
    }

    // thumbnailLink 是帶 token 的公開 URL，可以直接 fetch
    const https = require("https");
    const url = new URL(thumbnailLink);

    https.get(thumbnailLink, function(imgRes) {
      const contentType = imgRes.headers["content-type"] || "image/jpeg";
      res.setHeader("Content-Type", contentType);
      res.setHeader("Cache-Control", "public, max-age=3600");
      imgRes.pipe(res);
    }).on("error", function(err) {
      console.error("Thumb fetch error:", err);
      res.status(500).json({ error: "Failed to fetch thumbnail" });
    });

  } catch (e) {
    console.error("Thumb API error:", e);
    return res.status(500).json({ error: e.message });
  }
};
