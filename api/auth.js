// api/auth.js
// Google OAuth 2.0 - 取得 refresh token

const { google } = require("googleapis");

const CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const REDIRECT_URI = "https://2060422.vercel.app/api/auth/callback";

function getOAuthClient() {
  return new google.auth.OAuth2(CLIENT_ID, CLIENT_SECRET, REDIRECT_URI);
}

module.exports = async function(req, res) {
  const url = req.url || "";

  // /api/auth/callback - Google 授權後回調
  if (url.includes("/callback")) {
    const code = new URL(req.url, "https://2060422.vercel.app").searchParams.get("code");

    if (!code) {
      return res.status(400).send("Missing code parameter");
    }

    try {
      const oauth2Client = getOAuthClient();
      const { tokens } = await oauth2Client.getToken(code);

      const refreshToken = tokens.refresh_token || "(no refresh token)";

      return res.status(200).send(`
        <html>
        <body style="font-family:sans-serif;padding:40px;max-width:600px">
          <h2>授權成功！</h2>
          <p>請複製以下 Refresh Token，加到 Vercel 環境變數：</p>
          <p><strong>變數名稱：</strong> GOOGLE_REFRESH_TOKEN</p>
          <p><strong>變數值：</strong></p>
          <textarea style="width:100%;height:80px;font-size:12px">${refreshToken}</textarea>
          <br><br>
          <p style="color:#666">複製後到 Vercel → Settings → Environment Variables 新增。</p>
          ${refreshToken === "(no refresh token)" ? "<p style='color:red'>⚠️ 沒有取得 refresh token，請重新授權（確保已勾選 offline access）</p>" : ""}
        </body>
        </html>
      `);
    } catch (err) {
      return res.status(500).send("Error getting tokens: " + err.message);
    }
  }

  // /api/auth - 導向 Google 授權
  const oauth2Client = getOAuthClient();
  const authUrl = oauth2Client.generateAuthUrl({
    access_type: "offline",
    prompt: "consent",
    scope: [
      "https://www.googleapis.com/auth/drive.file",
      "https://www.googleapis.com/auth/calendar.events",
    ],
  });

  return res.redirect(authUrl);
};
