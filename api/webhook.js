// api/webhook.js
// LINE Bot — 個人管理工具 v3.0.0
// 功能：Threads 收藏 + 標籤管理、檔案永久備份、Google 日曆待辦
//
// ============================================================
//  環境變數（Vercel 後台 Environment Variables）
//
//  LINE_CHANNEL_ACCESS_TOKEN   → LINE Developers 取得
//  GOOGLE_SERVICE_ACCOUNT_JSON → GCP 服務帳號 JSON（整份內容）
//  GOOGLE_SHEETS_ID            → Google Sheets 試算表 ID
//  GOOGLE_DRIVE_FOLDER_ID      → Drive 備份根資料夾 ID
//  GOOGLE_CALENDAR_ID          → 日曆 ID（或填 "primary"）
// ============================================================

const { google } = require("googleapis");
const https = require("https");
const path = require("path");

const LINE_TOKEN         = process.env.LINE_CHANNEL_ACCESS_TOKEN;
const GOOGLE_CREDENTIALS = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
const SHEETS_ID          = process.env.GOOGLE_SHEETS_ID;
const DRIVE_FOLDER_ID    = process.env.GOOGLE_DRIVE_FOLDER_ID;
const CALENDAR_ID        = process.env.GOOGLE_CALENDAR_ID || "primary";

// 工作表分頁名稱（請在試算表內建立這三個分頁）
const SHEET = {
  THREADS:  "Threads收藏",
  FILES:    "檔案備份",
  CALENDAR: "日曆待辦",
};

// ============================================================
//  Vercel Serverless 進入點
// ============================================================
module.exports = async (req, res) => {
  if (req.method === "GET") {
    return res.status(200).json({ status: "ok", version: "3.0.0" });
  }
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
    if (!body.events || body.events.length === 0) {
      return res.status(200).json({ status: "ok" });
    }

    const clients = await getGoogleClients();

    for (const event of body.events) {
      await handleEvent(event, clients);
    }
  } catch (err) {
    console.error("Global error:", err);
  }

  return res.status(200).json({ status: "ok" });
};

// ============================================================
//  Google API 初始化
// ============================================================
async function getGoogleClients() {
  const credentials = JSON.parse(GOOGLE_CREDENTIALS);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/drive",
      "https://www.googleapis.com/auth/calendar",
    ],
  });
  const authClient = await auth.getClient();
  return {
    sheets:   google.sheets({ version: "v4", auth: authClient }),
    drive:    google.drive({ version: "v3", auth: authClient }),
    calendar: google.calendar({ version: "v3", auth: authClient }),
  };
}

// ============================================================
//  事件路由
// ============================================================
async function handleEvent(event, clients) {
  const { replyToken, type, message, source } = event;
  const userId = source?.userId || "unknown";

  try {
    if (type === "message" && message.type === "text") {
      await handleTextMessage(replyToken, message.text.trim(), userId, clients);
    } else if (type === "message" && ["image", "video", "file"].includes(message.type)) {
      await handleFileMessage(replyToken, message, message.type, userId, clients);
    } else if (type === "message" && message.type === "audio") {
      await replyText(replyToken, "🎙 語音功能尚未開放，請用文字輸入待辦事項。");
    }
  } catch (err) {
    console.error("handleEvent error:", err);
    await replyText(replyToken, "⚠️ 發生錯誤：" + err.message);
  }
}

// ============================================================
//  文字訊息路由
// ============================================================
async function handleTextMessage(replyToken, text, userId, clients) {
  const { sheets, calendar } = clients;
  const lower = text.toLowerCase();

  // ── Threads 連結偵測（優先）────────────────────────────
  const threadsUrls = extractThreadsUrls(text);
  if (threadsUrls.length > 0) {
    await handleThreadsSave(replyToken, text, threadsUrls, userId, sheets);
    return;
  }

  // ── 標籤操作 ─────────────────────────────────────────
  if (/^(標籤|加標籤|幫我加標籤)\s+\S+/.test(text)) {
    await handleAddTag(replyToken, text, userId, sheets);
    return;
  }

  // ── 修改標籤 ─────────────────────────────────────────
  if (/^(改標籤|修改標籤|更改標籤)\s+\S+/.test(text)) {
    await handleEditTag(replyToken, text, userId, sheets);
    return;
  }

  // ── 查標籤 ───────────────────────────────────────────
  if (/^(查標籤|查詢標籤)\s+\S+/.test(text)) {
    const tag = text.replace(/^(查標籤|查詢標籤)\s+/, "").trim();
    await handleQueryTag(replyToken, tag, sheets);
    return;
  }
  if (/^#\S+/.test(text)) {
    const tag = text.replace(/^#/, "").trim();
    await handleQueryTag(replyToken, tag, sheets);
    return;
  }

  // ── 搜尋 Threads ─────────────────────────────────────
  if (/^(搜尋|搜索|找)\s+\S+/.test(text) && !text.includes("檔案")) {
    const keyword = text.replace(/^(搜尋|搜索|找)\s+/, "").trim();
    await handleSearchThreads(replyToken, keyword, sheets);
    return;
  }

  // ── 最近 Threads 收藏 ─────────────────────────────────
  if (/最近.*收藏|最新.*threads|最近.*threads/i.test(lower) || text === "最近收藏") {
    await handleRecentThreads(replyToken, sheets);
    return;
  }

  // ── 刪除最後一筆 ─────────────────────────────────────
  if (/^(刪除|delete)\s*(上一筆|最後一筆|last)?$/.test(text.trim())) {
    await handleDeleteLast(replyToken, userId, sheets);
    return;
  }

  // ── 最近檔案 ─────────────────────────────────────────
  if (/最近.*檔案|檔案.*最近/.test(text) || text === "最近檔案") {
    await handleRecentFiles(replyToken, sheets);
    return;
  }

  // ── 搜尋檔案 ─────────────────────────────────────────
  if (/^(搜尋檔案|找檔案|搜檔案)\s+\S+/.test(text)) {
    const keyword = text.replace(/^(搜尋檔案|找檔案|搜檔案)\s+/, "").trim();
    await handleSearchFiles(replyToken, keyword, sheets);
    return;
  }

  // ── 日曆待辦 ─────────────────────────────────────────
  if (/提醒|待辦|日曆|行程|開會|會議|早上|下午|晚上|明天|今天|後天|下週|下禮拜/.test(text)) {
    await handleCalendar(replyToken, text, userId, calendar, sheets);
    return;
  }

  // ── 說明 ─────────────────────────────────────────────
  if (/^(說明|help|指令|怎麼用|功能)$/.test(lower)) {
    await replyHelp(replyToken);
    return;
  }

  // ── 預設提示 ─────────────────────────────────────────
  await replyText(replyToken,
    "你好！我可以幫你做這些事：\n\n" +
    "📌 貼 Threads 連結 → 自動收藏\n" +
    "📁 傳圖片／檔案 → 自動備份到 Drive\n" +
    "📅 說待辦事項 → 新增到 Google 日曆\n\n" +
    "輸入「說明」查看完整指令"
  );
}

// ============================================================
//  A. Threads 收藏
// ============================================================

function extractThreadsUrls(text) {
  const regex = /https?:\/\/(www\.)?threads\.(net|com)\/[^\s]*/gi;
  return text.match(regex) || [];
}

async function handleThreadsSave(replyToken, rawText, urls, userId, sheets) {
  // 連結以外的文字當備註
  let note = rawText;
  urls.forEach(u => { note = note.replace(u, "").trim(); });
  note = note.replace(/^[\s\-—:：]+/, "").trim();

  const saved = [];
  for (const url of urls) {
    const isDup = await checkDuplicateThreads(url, sheets);
    if (isDup) { saved.push({ url, duplicate: true }); continue; }

    const id  = generateId("T");
    const now = formatDateTime(new Date());
    // 欄位順序：編號, 連結, 標題, 摘要, 備註, 標籤, userId, 建立時間, 更新時間, 原始訊息
    const row = [id, url, "", "", note, "", userId, now, now, rawText];
    await appendRow(sheets, SHEET.THREADS, row);
    saved.push({ url, id, duplicate: false });
  }

  if (saved.length === 1 && !saved[0].duplicate) {
    await replyFlex(replyToken, buildThreadsSavedFlex(saved[0].url, note));
  } else if (saved.length === 1 && saved[0].duplicate) {
    await replyText(replyToken, "⚠️ 這個 Threads 連結之前已經收藏過了！");
  } else {
    const lines = saved.map((s, i) =>
      s.duplicate ? `${i + 1}. ⚠️ 已收藏過（略過）` : `${i + 1}. ✅ 已收藏`
    );
    await replyText(replyToken,
      `📌 收到 ${urls.length} 個連結\n\n${lines.join("\n")}\n\n💡 輸入「標籤 美妝 靈感」幫最後一筆加標籤`
    );
  }
}

function buildThreadsSavedFlex(url, note) {
  const now = formatDateTime(new Date());
  return {
    type: "flex", altText: "✅ 已收藏這篇 Threads",
    contents: {
      type: "bubble",
      header: {
        type: "box", layout: "vertical",
        backgroundColor: "#000000", paddingAll: "16px",
        contents: [{ type: "text", text: "📌 Threads 已收藏", color: "#ffffff", weight: "bold", size: "md" }],
      },
      body: {
        type: "box", layout: "vertical", paddingAll: "16px", spacing: "md",
        contents: [
          makeRow("備註", note || "（無備註）"),
          makeRow("標籤", "未分類"),
          makeRow("時間", now),
        ],
      },
      footer: {
        type: "box", layout: "vertical", paddingAll: "12px", spacing: "sm",
        contents: [
          { type: "button", style: "link", height: "sm",
            action: { type: "uri", label: "🔗 開啟 Threads", uri: url } },
          { type: "text", text: "💡 回覆「標籤 美妝 靈感」來加標籤", size: "xs", color: "#aaaaaa", align: "center" },
        ],
      },
    },
  };
}

async function handleAddTag(replyToken, text, userId, sheets) {
  const tagStr = text.replace(/^(標籤|加標籤|幫我加標籤)\s+/, "").trim();
  const tags   = tagStr.split(/[\s,，、]+/).filter(t => t.length > 0);

  const { row: lastRow, index: lastIndex } = await findLastUserRow(sheets, SHEET.THREADS, userId);
  if (lastIndex === -1) {
    await replyText(replyToken, "找不到最近收藏的 Threads，請先貼上連結。");
    return;
  }

  const tagValue = tags.join("、");
  const rowNum   = lastIndex + 1;
  const now      = formatDateTime(new Date());

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SHEETS_ID,
    requestBody: {
      valueInputOption: "RAW",
      data: [
        { range: `${SHEET.THREADS}!F${rowNum}`, values: [[tagValue]] },
        { range: `${SHEET.THREADS}!I${rowNum}`, values: [[now]] },
      ],
    },
  });

  const url = String(lastRow[1] || "");
  await replyFlex(replyToken, {
    type: "flex", altText: "✅ 標籤已更新",
    contents: {
      type: "bubble",
      body: {
        type: "box", layout: "vertical", paddingAll: "16px", spacing: "md",
        contents: [
          { type: "text", text: "✅ 標籤已更新", weight: "bold", size: "lg" },
          makeRow("標籤", tags.map(t => `#${t}`).join(" ")),
          { type: "button", style: "link", height: "sm",
            action: { type: "uri", label: "🔗 開啟 Threads", uri: url } },
        ],
      },
    },
  });
}

async function handleEditTag(replyToken, text, userId, sheets) {
  const tagStr = text.replace(/^(改標籤|修改標籤|更改標籤)\s+/, "").trim();
  const tags   = tagStr.split(/[\s,，、]+/).filter(t => t.length > 0);

  // 和 addTag 邏輯相同，直接覆寫標籤欄
  const fakeText = "標籤 " + tagStr;
  await handleAddTag(replyToken, fakeText, userId, sheets);
}

async function handleQueryTag(replyToken, tag, sheets) {
  if (!tag) {
    await replyText(replyToken, "請輸入要查詢的標籤，例如：查標籤 木地板");
    return;
  }

  const data    = await getSheetData(sheets, SHEET.THREADS);
  const all     = data.slice(1).filter(row => String(row[5] || "").includes(tag));
  const results = all.reverse().slice(0, 10);

  if (results.length === 0) {
    await replyText(replyToken, `找不到標籤「${tag}」的 Threads。\n輸入「最近收藏」查看所有收藏。`);
    return;
  }

  const lines = results.map((r, i) => {
    const url  = String(r[1] || "");
    const note = String(r[4] || "").substring(0, 30) || "（無備註）";
    const tags = String(r[5] || "未分類");
    const date = String(r[7] || "").substring(0, 10);
    return `${i + 1}. ${note}\n   🏷 ${tags}　📅 ${date}\n   🔗 ${url}`;
  });

  await replyText(replyToken,
    `🏷 標籤「${tag}」共 ${all.length} 筆（顯示最新 ${results.length} 筆）\n\n${lines.join("\n\n")}`
  );
}

async function handleRecentThreads(replyToken, sheets) {
  const data    = await getSheetData(sheets, SHEET.THREADS);
  const results = data.slice(1).reverse().slice(0, 10);

  if (results.length === 0) {
    await replyText(replyToken, "還沒有收藏任何 Threads！");
    return;
  }

  const lines = results.map((r, i) => {
    const url  = String(r[1] || "");
    const note = String(r[4] || "").substring(0, 30) || "（無備註）";
    const tags = String(r[5] || "未分類");
    const date = String(r[7] || "").substring(0, 10);
    return `${i + 1}. ${note}\n   🏷 ${tags}　📅 ${date}\n   🔗 ${url}`;
  });

  await replyText(replyToken, `📌 最近 ${results.length} 筆 Threads\n\n${lines.join("\n\n")}`);
}

async function handleSearchThreads(replyToken, keyword, sheets) {
  const data    = await getSheetData(sheets, SHEET.THREADS);
  const results = data.slice(1)
    .filter(row => [row[1], row[4], row[5], row[9]].join(" ").includes(keyword))
    .reverse()
    .slice(0, 8);

  if (results.length === 0) {
    await replyText(replyToken, `找不到包含「${keyword}」的 Threads。`);
    return;
  }

  const lines = results.map((r, i) => {
    const url  = String(r[1] || "");
    const note = String(r[4] || "").substring(0, 30) || "（無備註）";
    const tags = String(r[5] || "未分類");
    const date = String(r[7] || "").substring(0, 10);
    return `${i + 1}. ${note}\n   🏷 ${tags}　📅 ${date}\n   🔗 ${url}`;
  });

  await replyText(replyToken,
    `🔍 「${keyword}」找到 ${results.length} 筆\n\n${lines.join("\n\n")}`
  );
}

async function handleDeleteLast(replyToken, userId, sheets) {
  const { row: lastRow, index: lastIndex } = await findLastUserRow(sheets, SHEET.THREADS, userId);

  if (lastIndex === -1) {
    await replyText(replyToken, "找不到可以刪除的 Threads。");
    return;
  }

  const url    = String(lastRow[1] || "");
  const rowNum = lastIndex + 1;
  await sheets.spreadsheets.values.clear({
    spreadsheetId: SHEETS_ID,
    range: `${SHEET.THREADS}!A${rowNum}:J${rowNum}`,
  });

  await replyText(replyToken, `🗑 已刪除最後一筆\n${url}`);
}

async function checkDuplicateThreads(url, sheets) {
  const data = await getSheetData(sheets, SHEET.THREADS);
  return data.slice(1).some(row => String(row[1] || "") === url);
}

// ============================================================
//  B. 檔案備份
// ============================================================

async function handleFileMessage(replyToken, message, type, userId, clients) {
  const { sheets, drive } = clients;

  await replyText(replyToken, "📁 收到檔案，正在備份到 Google Drive，請稍候...");

  const fileBuffer = await downloadLineContent(message.id);
  const fileName   = message.fileName || generateFileName(type, message.id);
  const mimeType   = getMimeType(type, fileName);
  const fileType   = classifyFileType(type, fileName);

  const driveFile  = await uploadToDrive(drive, fileName, mimeType, fileBuffer, fileType);

  const id  = generateId("F");
  const now = formatDateTime(new Date());
  // 欄位：編號, 檔名, 類型, MIME, 大小(bytes), DriveID, 永久連結, 標籤, 備註, userId, messageId, 建立時間
  const row = [
    id, fileName, fileType, mimeType,
    fileBuffer.length, driveFile.id, driveFile.webViewLink,
    "", "", userId, message.id, now,
  ];
  await appendRow(sheets, SHEET.FILES, row);

  await replyFlex(replyToken, buildFileSavedFlex(fileName, fileType, driveFile.webViewLink, now));
}

async function downloadLineContent(messageId) {
  return new Promise((resolve, reject) => {
    https.get({
      hostname: "api-data.line.me",
      path: `/v2/bot/message/${messageId}/content`,
      method: "GET",
      headers: { Authorization: `Bearer ${LINE_TOKEN}` },
    }, res => {
      const chunks = [];
      res.on("data", chunk => chunks.push(chunk));
      res.on("end",  () => resolve(Buffer.concat(chunks)));
      res.on("error", reject);
    }).on("error", reject);
  });
}

async function uploadToDrive(drive, fileName, mimeType, buffer, fileType) {
  const folderId = await getOrCreateFolder(drive, fileType);
  const { Readable } = require("stream");
  const stream = new Readable();
  stream.push(buffer);
  stream.push(null);

  const { data } = await drive.files.create({
    requestBody: { name: fileName, parents: [folderId] },
    media: { mimeType, body: stream },
    fields: "id, webViewLink, name",
  });

  // 設定任何人可讀（永久連結）
  await drive.permissions.create({
    fileId: data.id,
    requestBody: { role: "reader", type: "anyone" },
  });

  return data;
}

async function getOrCreateFolder(drive, folderName) {
  const res = await drive.files.list({
    q: `'${DRIVE_FOLDER_ID}' in parents and name='${folderName}' and mimeType='application/vnd.google-apps.folder' and trashed=false`,
    fields: "files(id)",
  });

  if (res.data.files.length > 0) return res.data.files[0].id;

  const folder = await drive.files.create({
    requestBody: {
      name: folderName,
      mimeType: "application/vnd.google-apps.folder",
      parents: [DRIVE_FOLDER_ID],
    },
    fields: "id",
  });
  return folder.data.id;
}

function buildFileSavedFlex(fileName, fileType, driveUrl, now) {
  const emoji = { image: "🖼", pdf: "📄", excel: "📊", word: "📝", video: "🎬", other: "📁" };
  return {
    type: "flex", altText: "📁 檔案備份成功",
    contents: {
      type: "bubble",
      header: {
        type: "box", layout: "vertical",
        backgroundColor: "#1a73e8", paddingAll: "16px",
        contents: [{ type: "text", text: "📁 檔案備份成功", color: "#ffffff", weight: "bold", size: "md" }],
      },
      body: {
        type: "box", layout: "vertical", paddingAll: "16px", spacing: "md",
        contents: [
          makeRow("檔名", fileName),
          makeRow("類型", `${emoji[fileType] || "📁"} ${fileType.toUpperCase()}`),
          makeRow("時間", now),
        ],
      },
      footer: {
        type: "box", layout: "vertical", paddingAll: "12px",
        contents: [
          { type: "button", style: "primary", height: "sm", color: "#1a73e8",
            action: { type: "uri", label: "🔗 開啟 Google Drive", uri: driveUrl } },
        ],
      },
    },
  };
}

async function handleRecentFiles(replyToken, sheets) {
  const data    = await getSheetData(sheets, SHEET.FILES);
  const results = data.slice(1).reverse().slice(0, 10);

  if (results.length === 0) {
    await replyText(replyToken, "還沒有備份任何檔案。");
    return;
  }

  const lines = results.map((r, i) => {
    const name = String(r[1] || "");
    const type = String(r[2] || "");
    const url  = String(r[6] || "");
    const date = String(r[11] || "").substring(0, 10);
    return `${i + 1}. ${name}\n   📁 ${type}　📅 ${date}\n   🔗 ${url}`;
  });

  await replyText(replyToken, `📁 最近 ${results.length} 筆備份\n\n${lines.join("\n\n")}`);
}

async function handleSearchFiles(replyToken, keyword, sheets) {
  const data    = await getSheetData(sheets, SHEET.FILES);
  const results = data.slice(1)
    .filter(row => [row[1], row[2], row[7], row[8]].join(" ").includes(keyword))
    .reverse()
    .slice(0, 8);

  if (results.length === 0) {
    await replyText(replyToken, `找不到「${keyword}」相關檔案。`);
    return;
  }

  const lines = results.map((r, i) => {
    const name = String(r[1] || "");
    const url  = String(r[6] || "");
    const date = String(r[11] || "").substring(0, 10);
    return `${i + 1}. ${name}\n   📅 ${date}\n   🔗 ${url}`;
  });

  await replyText(replyToken, `🔍 「${keyword}」找到 ${results.length} 筆\n\n${lines.join("\n\n")}`);
}

// ============================================================
//  C. Google 日曆待辦
// ============================================================

async function handleCalendar(replyToken, text, userId, calendar, sheets) {
  const parsed = parseCalendarInput(text);

  if (!parsed) {
    await replyText(replyToken,
      "⚠️ 我看到待辦，但無法解析時間。\n\n請用這種格式：\n" +
      "「明天下午3點 開會」\n" +
      "「4/25 10:00 打電話給客戶」\n" +
      "「下週一 確認報價」"
    );
    return;
  }

  try {
    const event = await createCalendarEvent(calendar, parsed);

    const id  = generateId("C");
    const now = formatDateTime(new Date());
    const row = [
      id, parsed.title,
      formatDate(parsed.start),
      parsed.isAllDay ? "" : formatTime(parsed.start),
      parsed.isAllDay ? "TRUE" : "FALSE",
      event.id, text, "text", userId, now,
    ];
    await appendRow(sheets, SHEET.CALENDAR, row);

    await replyFlex(replyToken, buildCalendarFlex(parsed, event));
  } catch (err) {
    console.error("Calendar error:", err);
    await replyText(replyToken, "⚠️ 建立日曆失敗：" + err.message);
  }
}

function parseCalendarInput(text) {
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  let targetDate = null;
  let targetTime = null;
  let title = text;

  // ── 日期解析 ──
  if (/今天|今日/.test(text)) {
    targetDate = new Date(now);
    title = title.replace(/今天|今日/g, "");
  } else if (/明天|明日/.test(text)) {
    targetDate = new Date(now);
    targetDate.setDate(targetDate.getDate() + 1);
    title = title.replace(/明天|明日/g, "");
  } else if (/後天/.test(text)) {
    targetDate = new Date(now);
    targetDate.setDate(targetDate.getDate() + 2);
    title = title.replace(/後天/g, "");
  } else {
    const weekMatch = text.match(/下[週禮拜]+([一二三四五六日天])/);
    if (weekMatch) {
      const dayMap = { 一: 1, 二: 2, 三: 3, 四: 4, 五: 5, 六: 6, 日: 0, 天: 0 };
      const target = dayMap[weekMatch[1]];
      targetDate   = new Date(now);
      let diff = (target - targetDate.getDay() + 7) % 7;
      if (diff === 0) diff = 7;
      diff += 7;
      targetDate.setDate(targetDate.getDate() + diff);
      title = title.replace(weekMatch[0], "");
    } else {
      const mdMatch = text.match(/(\d{1,2})[\/月](\d{1,2})[日号]?/);
      if (mdMatch) {
        targetDate = new Date(now.getFullYear(), parseInt(mdMatch[1]) - 1, parseInt(mdMatch[2]));
        title = title.replace(mdMatch[0], "");
      }
    }
  }

  if (!targetDate) return null;

  // ── 時間解析 ──
  const isPM = /下午|晚上|傍晚/.test(text);
  const isAM = /上午|早上/.test(text);
  const timeMatch = text.match(/(\d{1,2})[:點時](\d{0,2})/);

  if (timeMatch) {
    let hour  = parseInt(timeMatch[1]);
    const min = parseInt(timeMatch[2] || "0");
    if (isPM && hour < 12) hour += 12;
    if (isAM && hour === 12) hour = 0;
    targetTime = { hour, min };
    title = title.replace(timeMatch[0], "");
  }

  // ── 清理標題 ──
  title = title
    .replace(/上午|早上|下午|晚上|傍晚|中午|提醒我|提醒/g, "")
    .replace(/\s+/g, " ")
    .trim();
  if (!title) title = "待辦事項";

  const start    = new Date(targetDate);
  const isAllDay = !targetTime;

  if (targetTime) start.setHours(targetTime.hour, targetTime.min, 0, 0);

  const end = new Date(start);
  if (isAllDay) end.setDate(end.getDate() + 1);
  else          end.setHours(end.getHours() + 1);

  return { title, start, end, isAllDay };
}

async function createCalendarEvent(calendar, parsed) {
  const body = {
    summary: parsed.title,
    reminders: {
      useDefault: false,
      overrides: parsed.isAllDay ? [] : [{ method: "popup", minutes: 10 }],
    },
  };

  if (parsed.isAllDay) {
    body.start = { date: formatDate(parsed.start) };
    body.end   = { date: formatDate(parsed.end) };
  } else {
    body.start = { dateTime: parsed.start.toISOString(), timeZone: "Asia/Taipei" };
    body.end   = { dateTime: parsed.end.toISOString(),   timeZone: "Asia/Taipei" };
  }

  const { data } = await calendar.events.insert({
    calendarId: CALENDAR_ID,
    requestBody: body,
  });
  return data;
}

function buildCalendarFlex(parsed, event) {
  const timeStr = parsed.isAllDay
    ? `${formatDate(parsed.start)}（全天）`
    : `${formatDate(parsed.start)} ${formatTime(parsed.start)}`;

  return {
    type: "flex", altText: `✅ 已新增：${parsed.title}`,
    contents: {
      type: "bubble",
      header: {
        type: "box", layout: "vertical",
        backgroundColor: "#0F9D58", paddingAll: "16px",
        contents: [{ type: "text", text: "✅ 已新增到 Google 日曆", color: "#ffffff", weight: "bold", size: "md" }],
      },
      body: {
        type: "box", layout: "vertical", paddingAll: "16px", spacing: "md",
        contents: [
          makeRow("標題", parsed.title),
          makeRow("時間", timeStr),
          makeRow("提醒", parsed.isAllDay ? "（全天，無提醒）" : "事件前 10 分鐘"),
        ],
      },
      footer: event.htmlLink
        ? {
            type: "box", layout: "vertical", paddingAll: "12px",
            contents: [
              { type: "button", style: "primary", height: "sm", color: "#0F9D58",
                action: { type: "uri", label: "📅 開啟 Google 日曆", uri: event.htmlLink } },
            ],
          }
        : undefined,
    },
  };
}

// ============================================================
//  說明訊息
// ============================================================

async function replyHelp(replyToken) {
  await replyFlex(replyToken, {
    type: "flex", altText: "📋 使用說明",
    contents: {
      type: "bubble", size: "mega",
      header: {
        type: "box", layout: "vertical",
        backgroundColor: "#2C3E50", paddingAll: "16px",
        contents: [{ type: "text", text: "📋 使用說明", color: "#ffffff", weight: "bold", size: "lg" }],
      },
      body: {
        type: "box", layout: "vertical", paddingAll: "16px", spacing: "lg",
        contents: [
          helpSection("📌 Threads 收藏", [
            "貼上連結 → 自動收藏",
            "標籤 美妝 靈感 → 加標籤",
            "改標籤 企劃 素材 → 修改標籤",
            "查標籤 木地板 或 #木地板 → 查詢",
            "搜尋 保養 → 關鍵字搜尋",
            "最近收藏 → 最新 10 筆",
            "刪除上一筆 → 刪除最後一筆",
          ]),
          helpSection("📁 檔案備份", [
            "傳圖片／PDF／檔案 → 自動備份到 Drive",
            "最近檔案 → 最新 10 筆",
            "找檔案 報價 → 搜尋檔名",
          ]),
          helpSection("📅 日曆待辦", [
            "明天下午3點 開會",
            "4/25 10:00 打電話給客戶",
            "下週一 確認報價",
            "今天晚上8點 整理文件",
          ]),
        ],
      },
    },
  });
}

function helpSection(title, items) {
  return {
    type: "box", layout: "vertical", spacing: "xs",
    contents: [
      { type: "text", text: title, size: "sm", weight: "bold", color: "#2C3E50" },
      ...items.map(t => ({
        type: "text", text: `• ${t}`, size: "xs", color: "#555555", wrap: true,
      })),
      { type: "separator" },
    ],
  };
}

// ============================================================
//  共用 UI 元件
// ============================================================

function makeRow(label, value) {
  return {
    type: "box", layout: "horizontal",
    contents: [
      { type: "text", text: label,             size: "sm", color: "#888888", flex: 1 },
      { type: "text", text: String(value || ""), size: "sm", color: "#333333", flex: 3, wrap: true },
    ],
  };
}

// ============================================================
//  工具函式
// ============================================================

function generateId(prefix) {
  const ts  = new Date().toISOString().replace(/[-:T.Z]/g, "").substring(0, 14);
  const rnd = Math.floor(Math.random() * 1000).toString().padStart(3, "0");
  return `${prefix}${ts}${rnd}`;
}

function formatDateTime(date) {
  return new Date(date).toLocaleString("zh-TW", {
    timeZone: "Asia/Taipei",
    year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit",
  });
}

function formatDate(date) {
  const d   = new Date(date);
  const y   = d.getFullYear();
  const m   = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function formatTime(date) {
  const d = new Date(date);
  return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
}

function generateFileName(type, messageId) {
  const ext = { image: "jpg", video: "mp4", audio: "m4a", file: "bin" };
  const now  = new Date().toISOString().substring(0, 10).replace(/-/g, "");
  return `${type}_${now}_${messageId}.${ext[type] || "bin"}`;
}

function getMimeType(type, fileName) {
  const ext = path.extname(fileName || "").toLowerCase();
  const map = {
    ".jpg":  "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png":  "image/png",
    ".gif":  "image/gif",
    ".webp": "image/webp",
    ".pdf":  "application/pdf",
    ".doc":  "application/msword",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".xls":  "application/vnd.ms-excel",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".mp4":  "video/mp4",
    ".m4a":  "audio/m4a",
  };
  if (map[ext]) return map[ext];
  if (type === "image") return "image/jpeg";
  if (type === "video") return "video/mp4";
  return "application/octet-stream";
}

function classifyFileType(type, fileName) {
  if (type === "image") return "image";
  if (type === "video") return "video";
  const ext = path.extname(fileName || "").toLowerCase();
  if ([".pdf"].includes(ext)) return "pdf";
  if ([".xls", ".xlsx", ".csv"].includes(ext)) return "excel";
  if ([".doc", ".docx"].includes(ext)) return "word";
  return "other";
}

// ── Sheets 工具 ───────────────────────────────────────────

async function getSheetData(sheets, sheetName) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEETS_ID,
    range: `${sheetName}!A1:Z`,
  });
  return res.data.values || [];
}

async function appendRow(sheets, sheetName, row) {
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEETS_ID,
    range: `${sheetName}!A1`,
    valueInputOption: "RAW",
    insertDataOption: "INSERT_ROWS",
    requestBody: { values: [row] },
  });
}

// 找某個 userId 在指定工作表的最後一筆資料
async function findLastUserRow(sheets, sheetName, userId) {
  const data = await getSheetData(sheets, sheetName);
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][6] || "") === userId) {
      return { row: data[i], index: i };
    }
  }
  return { row: null, index: -1 };
}

// ============================================================
//  LINE API
// ============================================================

async function replyText(replyToken, text) {
  await callLineAPI({ replyToken, messages: [{ type: "text", text }] });
}

async function replyFlex(replyToken, flexObj) {
  await callLineAPI({ replyToken, messages: [flexObj] });
}

function callLineAPI(payload) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify(payload);
    const req  = https.request({
      hostname: "api.line.me",
      path:     "/v2/bot/message/reply",
      method:   "POST",
      headers: {
        "Content-Type":   "application/json",
        "Authorization":  `Bearer ${LINE_TOKEN}`,
        "Content-Length": Buffer.byteLength(body),
      },
    }, res => {
      let data = "";
      res.on("data", c => data += c);
      res.on("end",  () => {
        if (res.statusCode !== 200) {
          console.error("LINE API Error:", res.statusCode, data);
        }
        resolve();
      });
    });
    req.on("error", reject);
    req.write(body);
    req.end();
  });
}
