// api/cron/remind.js
// Cron Job - 每天台灣時間早上9點執行
// 掃描今天有提醒時間的待辦，推播給對應用戶

const { google } = require("googleapis");
const https = require("https");

const GOOGLE_CREDENTIALS = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
const SHEETS_ID = process.env.GOOGLE_SHEETS_ID;
const LINE_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN;
const SHEET_NAME = "待辦清單";

async function getSheetsClient() {
  const credentials = JSON.parse(GOOGLE_CREDENTIALS);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  const authClient = await auth.getClient();
  return google.sheets({ version: "v4", auth: authClient });
}

function getTodayString() {
  // 台灣時間今天日期 YYYY-MM-DD
  return new Date().toLocaleDateString("zh-TW", {
    timeZone: "Asia/Taipei",
    year: "numeric", month: "2-digit", day: "2-digit",
  }).replace(/\//g, "-");
}

function pushMessage(userId, messages) {
  return new Promise(function(resolve, reject) {
    const body = JSON.stringify({ to: userId, messages: messages });
    const req = https.request({
      hostname: "api.line.me",
      path: "/v2/bot/message/push",
      method: "POST",
      headers: {
        Authorization: "Bearer " + LINE_TOKEN,
        "Content-Type": "application/json",
        "Content-Length": Buffer.byteLength(body),
      },
    }, function(res) {
      let data = "";
      res.on("data", function(c) { data += c; });
      res.on("end", function() {
        if (res.statusCode >= 200 && res.statusCode < 300) resolve(data);
        else reject(new Error("LINE Push error " + res.statusCode + ": " + data));
      });
    });
    req.on("error", reject);
    req.write(body);
    req.end();
  });
}

module.exports = async function(req, res) {
  // 驗證是 Vercel Cron 呼叫（防止外部亂呼叫）
  const authHeader = req.headers["authorization"];
  if (authHeader !== "Bearer " + process.env.CRON_SECRET) {
    return res.status(401).json({ error: "Unauthorized" });
  }

  try {
    const sheets = await getSheetsClient();
    const today = getTodayString();

    // 讀取待辦清單
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEETS_ID,
      range: SHEET_NAME + "!A:H",
    });

    const rows = result.data.values || [];

    // 找出今天有提醒且還是 pending 的待辦，依用戶分組
    const userTodos = {};
    rows.slice(1).forEach(function(r) {
      const todoId = String(r[0] || "");
      const content = String(r[1] || "");
      const status = String(r[2] || "pending");
      const remindAt = String(r[4] || "");
      const userId = String(r[5] || "");
      const type = String(r[6] || "");

      if (type !== "todo" || status !== "pending" || !remindAt || !userId) return;
      if (!remindAt.startsWith(today)) return;

      if (!userTodos[userId]) userTodos[userId] = [];
      userTodos[userId].push({ content, remindAt });
    });

    // 對每個用戶推播
    const userIds = Object.keys(userTodos);
    let pushed = 0;

    for (const userId of userIds) {
      const todos = userTodos[userId];
      if (!todos.length) continue;

      const lines = todos.map(function(t, i) {
        const timeStr = t.remindAt.substring(11, 16); // HH:MM
        return (i + 1) + ". " + t.content + (timeStr ? "（" + timeStr + "）" : "");
      });

      const messages = [
        {
          type: "flex",
          altText: "🔔 今日待辦提醒，共 " + todos.length + " 筆",
          contents: {
            type: "bubble",
            header: {
              type: "box", layout: "horizontal", paddingAll: "16px",
              backgroundColor: "#EDE4FF",
              contents: [
                { type: "text", text: "🔔", size: "xl", flex: 0 },
                {
                  type: "box", layout: "vertical", flex: 1, paddingStart: "10px",
                  contents: [
                    { type: "text", text: "今日待辦提醒", weight: "bold", size: "lg", color: "#7B5EA7" },
                    { type: "text", text: today + "，共 " + todos.length + " 筆", size: "xs", color: "#9B72CF" },
                  ],
                },
              ],
            },
            body: {
              type: "box", layout: "vertical", paddingAll: "16px", spacing: "md",
              backgroundColor: "#FDFAFF",
              contents: todos.slice(0, 10).map(function(t, i) {
                const timeStr = t.remindAt.length >= 16 ? t.remindAt.substring(11, 16) : "";
                return {
                  type: "box", layout: "vertical", paddingAll: "10px",
                  backgroundColor: i % 2 === 0 ? "#F5F0FF" : "#FDFAFF",
                  cornerRadius: "12px",
                  contents: [
                    { type: "text", text: (i + 1) + ". " + t.content, size: "sm", weight: "bold", color: "#4A3728", wrap: true },
                    timeStr ? { type: "text", text: "⏰ " + timeStr, size: "xs", color: "#FF8FAB", margin: "xs" } : null,
                  ].filter(Boolean),
                };
              }),
            },
            footer: {
              type: "box", layout: "vertical", paddingAll: "12px",
              backgroundColor: "#FDFAFF",
              contents: [
                {
                  type: "button", style: "primary", height: "sm", color: "#9B72CF",
                  action: { type: "uri", label: "📋 開啟待辦清單", uri: "https://liff.line.me/2009891497-Bd5P0goB?page=todos" },
                },
              ],
            },
          },
        },
      ];

      try {
        await pushMessage(userId, messages);
        pushed++;
      } catch(e) {
        console.error("Push failed for " + userId + ":", e.message);
      }
    }

    console.log("Remind cron done. Pushed to " + pushed + " users, today=" + today);
    return res.status(200).json({ success: true, pushed, today, users: userIds.length });

  } catch (err) {
    console.error("Cron remind error:", err);
    return res.status(500).json({ error: err.message });
  }
};
