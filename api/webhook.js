// api/webhook.js
// LINE Bot — 個人管理工具 v6.0.0
// 最終完整覆蓋版

const { google } = require("googleapis");
const https = require("https");
const path = require("path");
const { Readable } = require("stream");

const LINE_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN;
const GOOGLE_CREDENTIALS = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
const SHEETS_ID = process.env.GOOGLE_SHEETS_ID;
const DRIVE_FOLDER_ID = process.env.GOOGLE_DRIVE_FOLDER_ID;
const CALENDAR_ID = process.env.GOOGLE_CALENDAR_ID || "primary";

const SHEET = {
  THREADS: "Threads收藏",
  FILES: "檔案備份",
  CALENDAR: "日曆待辦",
  TODOS: "待辦清單",
};

const CATEGORIES = ["未分類", "個人", "工作", "家庭"];
const CATEGORY_COLORS = {
  未分類: "#888888",
  個人: "#FF6B6B",
  工作: "#4ECDC4",
  家庭: "#45B7D1",
};
const CATEGORY_EMOJI = {
  未分類: "📋",
  個人: "👤",
  工作: "💼",
  家庭: "🏠",
};

module.exports = async (req, res) => {
  if (req.method === "GET") {
    return res.status(200).json({ status: "ok", version: "6.0.0" });
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
    if (!body?.events?.length) {
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
    sheets: google.sheets({ version: "v4", auth: authClient }),
    drive: google.drive({ version: "v3", auth: authClient }),
    calendar: google.calendar({ version: "v3", auth: authClient }),
  };
}

async function handleEvent(event, clients) {
  const { replyToken, type, message, source, postback } = event;
  const userId = source?.userId || "unknown";

  try {
    if (type === "follow") {
      await replyFlex(replyToken, buildMainMenuFlex());
      return;
    }

    if (type === "postback" && postback?.data) {
      await handlePostback(replyToken, postback.data, userId, clients);
      return;
    }

    if (type !== "message" || !message) return;

    if (message.type === "text") {
      await handleTextMessage(replyToken, message.text.trim(), userId, clients);
      return;
    }

    if (["image", "video", "file", "audio"].includes(message.type)) {
      await handleFileMessage(replyToken, message, message.type, userId, clients);
      return;
    }

    await replyText(replyToken, "目前支援文字、圖片、影片、音訊與一般檔案。");
  } catch (err) {
    console.error("handleEvent error:", err);
    if (replyToken) {
      await replyText(replyToken, "⚠️ 發生錯誤，請稍後再試。\n" + err.message);
    }
  }
}

async function handlePostback(replyToken, data, userId, clients) {
  const { sheets } = clients;
  const params = new URLSearchParams(data);
  const action = params.get("action");

  if (action === "set_category") {
    await handleSetCategory(replyToken, params.get("cat"), userId, sheets);
    return;
  }

  if (action === "view_category") {
    await handleQueryCategory(replyToken, params.get("cat"), userId, sheets);
    return;
  }

  await replyText(replyToken, "⚠️ 未知的操作。");
}

async function handleTextMessage(replyToken, text, userId, clients) {
  const { sheets, calendar } = clients;
  const lower = text.toLowerCase();

  // 主功能說明
  if (text === "功能說明" || /^(說明|help|使用說明|功能|如何使用)$/i.test(text)) {
    await replyFlex(replyToken, buildMainMenuFlex());
    return;
  }

  // ===== 待辦功能選單 =====
  if (text === "待辦清單") {
    await replyFlex(replyToken, buildTodoMenuFlex());
    return;
  }

  if (text === "查看待辦") {
    await handleShowTodos(replyToken, userId, sheets);
    return;
  }

  if (text === "新增待辦") {
    await replyText(
      replyToken,
      "請直接輸入：\n新增待辦 今天整理報價單\n或\n新增待辦 明天下午3點提醒我回覆客戶"
    );
    return;
  }

  if (/^新增待辦\s+.+/.test(text)) {
    const content = text.replace(/^新增待辦\s+/, "").trim();
    await handleCreateTodo(replyToken, content, userId, sheets);
    return;
  }

  if (text === "完成待辦") {
    await handlePromptCompleteTodo(replyToken, userId, sheets);
    return;
  }

  if (/^完成待辦\s+\d+$/.test(text)) {
    const index = Number(text.replace(/^完成待辦\s+/, "").trim());
    await handleCompleteTodoByIndex(replyToken, index, userId, sheets);
    return;
  }

  if (text === "刪除待辦") {
    await handlePromptDeleteTodo(replyToken, userId, sheets);
    return;
  }

  if (/^刪除待辦\s+\d+$/.test(text)) {
    const index = Number(text.replace(/^刪除待辦\s+/, "").trim());
    await handleDeleteTodoByIndex(replyToken, index, userId, sheets);
    return;
  }

  // ===== 分類查詢 =====
  for (const cat of CATEGORIES) {
    if (text === `${cat}收藏` || text === cat) {
      await handleQueryCategory(replyToken, cat, userId, sheets);
      return;
    }
  }

  // 文字分類：把最近一筆改分類
  if (/^(分類|設為|歸類|標記為)\s*(個人|工作|家庭|未分類)/.test(text)) {
    const catMatch = text.match(/(個人|工作|家庭|未分類)/);
    if (catMatch) {
      await handleSetCategory(replyToken, catMatch[1], userId, sheets);
      return;
    }
  }

  // ===== 連結收藏 =====
  const urls = extractSupportedUrls(text);
  if (urls.length > 0) {
    await handleLinkSave(replyToken, text, urls, userId, sheets);
    return;
  }

  // ===== 標籤 =====
  if (/^(標籤|加標籤|幫我加標籤)\s+\S+/.test(text)) {
    await handleAddTag(replyToken, text, userId, sheets);
    return;
  }

  if (/^(改標籤|修改標籤|更改標籤)\s+\S+/.test(text)) {
    const tagStr = text.replace(/^(改標籤|修改標籤|更改標籤)\s+/, "").trim();
    await handleAddTag(replyToken, `標籤 ${tagStr}`, userId, sheets);
    return;
  }

  // ===== 查標籤 =====
  if (/^(查標籤|查詢標籤)\s+\S+/.test(text)) {
    const tag = text.replace(/^(查標籤|查詢標籤)\s+/, "").trim();
    await handleQueryTag(replyToken, tag, userId, sheets);
    return;
  }

  if (/^#\S+/.test(text)) {
    const tag = text.replace(/^#/, "").trim();
    await handleQueryTag(replyToken, tag, userId, sheets);
    return;
  }

  // ===== 搜尋收藏 =====
  if (/^(搜尋|搜索|找)\s+\S+/.test(text) && !text.includes("檔案")) {
    const keyword = text.replace(/^(搜尋|搜索|找)\s+/, "").trim();
    await handleSearchLinks(replyToken, keyword, userId, sheets);
    return;
  }

  // ===== 最近收藏 =====
  if (/最近.*收藏|最近.*threads|最近.*thr/i.test(lower) || text === "最近收藏") {
    await handleRecentLinks(replyToken, userId, sheets);
    return;
  }

  // ===== 刪除最近一筆收藏 =====
  if (/^(刪除|delete)\s*(最新|最近一筆|last)?$/.test(text.trim())) {
    await handleDeleteLast(replyToken, userId, sheets);
    return;
  }

  // ===== 最近檔案 =====
  if (/最近.*檔案|檔案.*最近/.test(text) || text === "最近檔案") {
    await handleRecentFiles(replyToken, userId, sheets);
    return;
  }

  // ===== 搜尋檔案 =====
  if (/^(搜尋檔案|找檔案|搜檔案)\s+\S+/.test(text)) {
    const keyword = text.replace(/^(搜尋檔案|找檔案|搜檔案)\s+/, "").trim();
    await handleSearchFiles(replyToken, keyword, userId, sheets);
    return;
  }

  // ===== Google 日曆 =====
  if (/明天|行程|提醒|會議|今天|下午|上午|早上|明後天|後天|下週|下周|下星期|週一|週二|週三|週四|週五|週六|週日/.test(text)) {
    await handleCalendar(replyToken, text, userId, calendar, sheets);
    return;
  }

  await replyFlex(replyToken, buildMainMenuFlex());
}

// ============================================================
//  待辦功能（簡易版）
//  Sheet: 待辦清單
//  A: todoId | B: content | C: status | D: createdAt | E: remindText | F: userId | G: todo
// ============================================================

function buildTodoMenuFlex() {
  return {
    type: "flex",
    altText: "待辦功能選單",
    contents: {
      type: "bubble",
      styles: { header: { backgroundColor: "#2D3561" } },
      header: {
        type: "box",
        layout: "vertical",
        paddingAll: "16px",
        contents: [
          { type: "text", text: "📋 待辦功能選單", color: "#FFFFFF", weight: "bold", size: "md" },
          { type: "text", text: "查看、新增、完成、刪除待辦", color: "#FFFFFFCC", size: "xs" },
        ],
      },
      body: {
        type: "box",
        layout: "vertical",
        paddingAll: "16px",
        spacing: "md",
        contents: [
          { type: "text", text: "新增待辦範例", weight: "bold", size: "sm", color: "#333333" },
          { type: "text", text: "新增待辦 今天整理報價單", size: "sm", color: "#666666", wrap: true },
          { type: "text", text: "新增待辦 明天下午3點提醒我回覆客戶", size: "sm", color: "#666666", wrap: true },
        ],
      },
      footer: {
        type: "box",
        layout: "vertical",
        spacing: "sm",
        paddingAll: "12px",
        contents: [
          {
            type: "box",
            layout: "horizontal",
            spacing: "sm",
            contents: [
              {
                type: "button",
                style: "secondary",
                height: "sm",
                flex: 1,
                action: { type: "message", label: "查看待辦", text: "查看待辦" },
              },
              {
                type: "button",
                style: "secondary",
                height: "sm",
                flex: 1,
                action: { type: "message", label: "新增待辦", text: "新增待辦" },
              },
            ],
          },
          {
            type: "box",
            layout: "horizontal",
            spacing: "sm",
            contents: [
              {
                type: "button",
                style: "secondary",
                height: "sm",
                flex: 1,
                action: { type: "message", label: "完成待辦", text: "完成待辦" },
              },
              {
                type: "button",
                style: "primary",
                height: "sm",
                flex: 1,
                color: "#2D3561",
                action: { type: "message", label: "刪除待辦", text: "刪除待辦" },
              },
            ],
          },
        ],
      },
    },
  };
}

async function getTodoRows(sheets, userId) {
  const todoData = await getSheetData(sheets, SHEET.TODOS);
  return todoData
    .slice(1)
    .filter((r) => String(r[6] || "") === "todo" && String(r[5] || "") === userId)
    .filter((r) => String(r[2] || "pending") !== "deleted");
}

async function handleCreateTodo(replyToken, content, userId, sheets) {
  const id = generateId("TD");
  const now = formatDateTime(new Date());
  const row = [id, content, "pending", now, "", userId, "todo"];
  await appendRow(sheets, SHEET.TODOS, row);

  await replyFlex(replyToken, {
    type: "flex",
    altText: `已新增待辦：${content}`,
    contents: {
      type: "bubble",
      body: {
        type: "box",
        layout: "vertical",
        paddingAll: "16px",
        spacing: "md",
        contents: [
          { type: "text", text: "✅ 已新增待辦", weight: "bold", size: "md", color: "#0F9D58" },
          { type: "text", text: content, size: "sm", color: "#333333", wrap: true },
        ],
      },
      footer: {
        type: "box",
        layout: "horizontal",
        spacing: "sm",
        paddingAll: "12px",
        contents: [
          {
            type: "button",
            style: "secondary",
            height: "sm",
            flex: 1,
            action: { type: "message", label: "查看待辦", text: "查看待辦" },
          },
          {
            type: "button",
            style: "primary",
            height: "sm",
            flex: 1,
            color: "#2D3561",
            action: { type: "message", label: "完成待辦", text: "完成待辦" },
          },
        ],
      },
    },
  });
}

async function handleShowTodos(replyToken, userId, sheets) {
  const rows = (await getTodoRows(sheets, userId)).filter((r) => String(r[2] || "pending") === "pending");

  if (!rows.length) {
    await replyFlex(
      replyToken,
      buildEmptyStateFlex(
        "目前沒有待辦",
        "你可以直接輸入：新增待辦 今天整理報價單",
        [{ label: "新增待辦", text: "新增待辦" }]
      )
    );
    return;
  }

  const lines = rows.slice(0, 10).map((r, i) => `${i + 1}. ${String(r[1] || "")}`);
  await replyText(replyToken, `📋 目前待辦\n\n${lines.join("\n")}`);
}

async function handlePromptCompleteTodo(replyToken, userId, sheets) {
  const rows = (await getTodoRows(sheets, userId)).filter((r) => String(r[2] || "pending") === "pending");

  if (!rows.length) {
    await replyText(replyToken, "目前沒有可完成的待辦。");
    return;
  }

  const lines = rows.slice(0, 10).map((r, i) => `${i + 1}. ${String(r[1] || "")}`);
  await replyText(replyToken, `請輸入：完成待辦 序號\n\n${lines.join("\n")}`);
}

async function handleCompleteTodoByIndex(replyToken, index, userId, sheets) {
  const allData = await getSheetData(sheets, SHEET.TODOS);
  const pendingRows = allData
    .map((row, idx) => ({ row, idx }))
    .filter(({ row }) =>
      String(row[6] || "") === "todo" &&
      String(row[5] || "") === userId &&
      String(row[2] || "pending") === "pending"
    );

  if (!index || index < 1 || index > pendingRows.length) {
    await replyText(replyToken, "序號無效，請重新輸入。");
    return;
  }

  const target = pendingRows[index - 1];
  const rowNum = target.idx + 2;
  const content = String(target.row[1] || "");

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEETS_ID,
    range: `${SHEET.TODOS}!C${rowNum}`,
    valueInputOption: "RAW",
    requestBody: { values: [["done"]] },
  });

  await replyText(replyToken, `✅ 已完成待辦\n${content}`);
}

async function handlePromptDeleteTodo(replyToken, userId, sheets) {
  const rows = (await getTodoRows(sheets, userId)).filter((r) => String(r[2] || "pending") === "pending");

  if (!rows.length) {
    await replyText(replyToken, "目前沒有可刪除的待辦。");
    return;
  }

  const lines = rows.slice(0, 10).map((r, i) => `${i + 1}. ${String(r[1] || "")}`);
  await replyText(replyToken, `請輸入：刪除待辦 序號\n\n${lines.join("\n")}`);
}

async function handleDeleteTodoByIndex(replyToken, index, userId, sheets) {
  const allData = await getSheetData(sheets, SHEET.TODOS);
  const pendingRows = allData
    .map((row, idx) => ({ row, idx }))
    .filter(({ row }) =>
      String(row[6] || "") === "todo" &&
      String(row[5] || "") === userId &&
      String(row[2] || "pending") === "pending"
    );

  if (!index || index < 1 || index > pendingRows.length) {
    await replyText(replyToken, "序號無效，請重新輸入。");
    return;
  }

  const target = pendingRows[index - 1];
  const rowNum = target.idx + 2;
  const content = String(target.row[1] || "");

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEETS_ID,
    range: `${SHEET.TODOS}!C${rowNum}`,
    valueInputOption: "RAW",
    requestBody: { values: [["deleted"]] },
  });

  await replyText(replyToken, `🗑️ 已刪除待辦\n${content}`);
}

// ============================================================
//  連結收藏
// ============================================================

function extractSupportedUrls(text) {
  const regex = /https?:\/\/(www\.)?(threads\.(net|com)|facebook\.com|fb\.com|youtube\.com|youtu\.be|instagram\.com|twitter\.com|x\.com)\/[^\s]*/gi;
  return text.match(regex) || [];
}

function detectPlatform(url) {
  if (/threads\.(net|com)/i.test(url)) return { name: "Threads", emoji: "🧵", color: "#000000" };
  if (/facebook\.com|fb\.com/i.test(url)) return { name: "Facebook", emoji: "📘", color: "#1877F2" };
  if (/youtube\.com|youtu\.be/i.test(url)) return { name: "YouTube", emoji: "📺", color: "#FF0000" };
  if (/instagram\.com/i.test(url)) return { name: "Instagram", emoji: "📸", color: "#E1306C" };
  if (/twitter\.com|x\.com/i.test(url)) return { name: "X (Twitter)", emoji: "🐦", color: "#000000" };
  return { name: "連結", emoji: "🔗", color: "#555555" };
}

async function handleLinkSave(replyToken, rawText, urls, userId, sheets) {
  let note = rawText;
  urls.forEach((u) => {
    note = note.replace(u, "").trim();
  });
  note = note.replace(/^[\s\-—：:]+/, "").trim();

  const saved = [];
  for (const url of urls) {
    const isDup = await checkDuplicateThreads(url, userId, sheets);
    if (isDup) {
      saved.push({ url, duplicate: true });
      continue;
    }

    const platform = detectPlatform(url);
    const id = generateId("T");
    const now = formatDateTime(new Date());
    const row = [id, url, platform.name, "", note, "", "未分類", userId, now, now, rawText];
    await appendRow(sheets, SHEET.THREADS, row);
    saved.push({ url, id, platform, duplicate: false, note });
  }

  if (saved.length === 1 && !saved[0].duplicate) {
    await replyFlex(replyToken, buildLinkSavedFlex(saved[0].url, saved[0].platform, saved[0].note));
    return;
  }

  if (saved.length === 1 && saved[0].duplicate) {
    await replyFlex(
      replyToken,
      buildEmptyStateFlex("⚠️ 已收藏過了", "這則連結之前已經收藏過", [{ label: "查看最近收藏", text: "最近收藏" }])
    );
    return;
  }

  const lines = saved.map((s, i) => {
    if (s.duplicate) return `${i + 1}. ⚠️ 已收藏過`;
    return `${i + 1}. ✅ ${s.platform.emoji} 已收藏`;
  });

  await replyText(replyToken, `📌 收藏 ${urls.length} 則連結\n\n${lines.join("\n")}`);
}

function buildLinkSavedFlex(url, platform, note) {
  const now = formatDateTime(new Date());

  return {
    type: "flex",
    altText: `${platform.emoji} 已收藏成功`,
    contents: {
      type: "bubble",
      size: "mega",
      header: {
        type: "box",
        layout: "vertical",
        paddingAll: "16px",
        backgroundColor: "#1A1A2E",
        contents: [
          {
            type: "text",
            text: "✅ 已收藏成功",
            color: "#FFFFFF",
            weight: "bold",
            size: "md"
          },
          {
            type: "text",
            text: `${platform.emoji} ${platform.name}`,
            color: "#FFFFFFCC",
            size: "sm"
          }
        ]
      },
      body: {
        type: "box",
        layout: "vertical",
        paddingAll: "16px",
        spacing: "md",
        contents: [
          note ? makeRow("備註", note) : makeRow("備註", "（未填寫）"),
          makeRow("分類", "未分類"),
          makeRow("時間", now),
          {
            type: "separator"
          },
          {
            type: "text",
            text: "請選擇收藏分類",
            size: "sm",
            color: "#555555",
            weight: "bold"
          },
          {
            type: "box",
            layout: "horizontal",
            spacing: "sm",
            contents: ["個人", "工作", "家庭"].map(cat => ({
              type: "button",
              style: "secondary",
              height: "sm",
              flex: 1,
              action: {
                type: "postback",
                label: cat,
                data: `action=set_category&id=LAST&cat=${cat}`,
                displayText: `設為${cat}`
              }
            }))
          }
        ]
      },
      footer: {
        type: "box",
        layout: "horizontal",
        spacing: "sm",
        paddingAll: "12px",
        contents: [
          {
            type: "button",
            style: "primary",
            flex: 2,
            color: "#1A1A2E",
            action: {
              type: "uri",
              label: "開啟連結",
              uri: url
            }
          },
          {
            type: "button",
            style: "secondary",
            flex: 1,
            action: {
              type: "message",
              label: "加標籤",
              text: "加標籤 "
            }
          }
        ]
      }
    }
  };
}

async function handleSetCategory(replyToken, category, userId, sheets) {
  const { row: lastRow, index: lastIndex } = await findLastUserRow(sheets, SHEET.THREADS, userId);
  if (lastIndex === -1) {
    await replyText(replyToken, "找不到最近的收藏。");
    return;
  }

  const rowNum = lastIndex + 1;
  const now = formatDateTime(new Date());

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SHEETS_ID,
    requestBody: {
      valueInputOption: "RAW",
      data: [
        { range: `${SHEET.THREADS}!G${rowNum}`, values: [[category]] },
        { range: `${SHEET.THREADS}!J${rowNum}`, values: [[now]] },
      ],
    },
  });

  const url = String(lastRow[1] || "");
  await replyFlex(replyToken, {
    type: "flex",
    altText: `✅ 已設為${category}`,
    contents: {
      type: "bubble",
      body: {
        type: "box",
        layout: "vertical",
        paddingAll: "20px",
        spacing: "md",
        contents: [
          {
            type: "box",
            layout: "horizontal",
            spacing: "md",
            contents: [
              { type: "text", text: CATEGORY_EMOJI[category], size: "xxl", flex: 0 },
              {
                type: "box",
                layout: "vertical",
                flex: 1,
                contents: [
                  { type: "text", text: `已歸類到「${category}」`, weight: "bold", size: "md", color: "#333333" },
                  { type: "text", text: `之後可用「${category}收藏」查詢`, size: "xs", color: "#888888" },
                ],
              },
            ],
          },
        ],
      },
      footer: {
        type: "box",
        layout: "horizontal",
        spacing: "sm",
        paddingAll: "12px",
        contents: [
          {
            type: "button",
            style: "primary",
            height: "sm",
            flex: 1,
            color: CATEGORY_COLORS[category] || "#1A1A2E",
            action: {
              type: "postback",
              label: `查看${category}`,
              data: `action=view_category&cat=${category}`,
              displayText: `${category}收藏`,
            },
          },
          {
            type: "button",
            style: "secondary",
            height: "sm",
            flex: 1,
            action: { type: "uri", label: "開啟連結", uri: url },
          },
        ],
      },
    },
  });
}

async function handleQueryCategory(replyToken, category, userId, sheets) {
  const data = await getSheetData(sheets, SHEET.THREADS);
  const all = data
    .slice(1)
    .filter((row) => String(row[7] || "") === userId && String(row[6] || "未分類") === category)
    .reverse();

  const results = all.slice(0, 10);

  if (!results.length) {
    await replyFlex(
      replyToken,
      buildEmptyStateFlex(
        `${CATEGORY_EMOJI[category]} ${category}還沒有收藏`,
        "收藏連結後，點成功卡片下方的分類按鈕即可歸類",
        [{ label: "查看最近收藏", text: "最近收藏" }]
      )
    );
    return;
  }

  await replyFlex(replyToken, buildLinksCarousel(results, `${CATEGORY_EMOJI[category]} ${category} — ${all.length} 筆`));
}

async function handleAddTag(replyToken, text, userId, sheets) {
  const tagStr = text.replace(/^(標籤|加標籤|幫我加標籤)\s+/, "").trim();
  const tags = tagStr.split(/[\s,，、]+/).filter((t) => t.length > 0);

  const { row: lastRow, index: lastIndex } = await findLastUserRow(sheets, SHEET.THREADS, userId);
  if (lastIndex === -1) {
    await replyText(replyToken, "找不到最近的收藏，請先貼上連結再加標籤。");
    return;
  }

  const tagValue = tags.join("、");
  const rowNum = lastIndex + 1;
  const now = formatDateTime(new Date());

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SHEETS_ID,
    requestBody: {
      valueInputOption: "RAW",
      data: [
        { range: `${SHEET.THREADS}!F${rowNum}`, values: [[tagValue]] },
        { range: `${SHEET.THREADS}!J${rowNum}`, values: [[now]] },
      ],
    },
  });

  const url = String(lastRow[1] || "");
  await replyFlex(replyToken, {
    type: "flex",
    altText: "✅ 標籤已更新",
    contents: {
      type: "bubble",
      body: {
        type: "box",
        layout: "vertical",
        paddingAll: "16px",
        spacing: "md",
        contents: [
          { type: "text", text: "🏷 標籤已更新", weight: "bold", size: "lg" },
          makeRow("標籤", tags.map((t) => `#${t}`).join(" ")),
          { type: "button", style: "link", height: "sm", action: { type: "uri", label: "🔗 開啟連結", uri: url } },
        ],
      },
    },
  });
}

async function handleQueryTag(replyToken, tag, userId, sheets) {
  if (!tag) {
    await replyText(replyToken, "請告訴我要查詢的標籤，例如：查標籤 旅遊");
    return;
  }

  const data = await getSheetData(sheets, SHEET.THREADS);
  const all = data.slice(1).filter((row) => String(row[7] || "") === userId && String(row[5] || "").includes(tag));
  const results = all.reverse().slice(0, 10);

  if (!results.length) {
    await replyFlex(
      replyToken,
      buildEmptyStateFlex("找不到符合的內容", `目前沒有標籤「${tag}」的收藏資料`, [{ label: "查看最近收藏", text: "最近收藏" }])
    );
    return;
  }

  await replyFlex(replyToken, buildLinksCarousel(results, `🏷 標籤「${tag}」— ${all.length} 筆`));
}

async function handleRecentLinks(replyToken, userId, sheets) {
  const data = await getSheetData(sheets, SHEET.THREADS);
  const results = data.slice(1).filter((r) => String(r[7] || "") === userId).reverse().slice(0, 10);

  if (!results.length) {
    await replyFlex(
      replyToken,
      buildEmptyStateFlex("還沒有收藏", "貼上連結就會自動收藏", [{ label: "使用說明", text: "說明" }])
    );
    return;
  }

  await replyFlex(replyToken, buildLinksCarousel(results, `📌 最近 ${results.length} 則收藏`));
}

async function handleSearchLinks(replyToken, keyword, userId, sheets) {
  const data = await getSheetData(sheets, SHEET.THREADS);
  const results = data
    .slice(1)
    .filter((row) => String(row[7] || "") === userId)
    .filter((row) => [row[1], row[4], row[5], row[6], row[10]].join(" ").includes(keyword))
    .reverse()
    .slice(0, 8);

  if (!results.length) {
    await replyFlex(
      replyToken,
      buildEmptyStateFlex("找不到符合的內容", `沒有找到包含「${keyword}」的收藏`, [{ label: "查看最近收藏", text: "最近收藏" }])
    );
    return;
  }

  await replyFlex(replyToken, buildLinksCarousel(results, `🔍「${keyword}」— ${results.length} 筆`));
}

function buildLinksCarousel(rows, headerTitle) {
  if (rows.length <= 5) {
    return {
      type: "flex",
      altText: headerTitle,
      contents: {
        type: "carousel",
        contents: rows.map((r) => buildLinkCardBubble(r)),
      },
    };
  }

  return {
    type: "flex",
    altText: headerTitle,
    contents: {
      type: "bubble",
      size: "mega",
      header: {
        type: "box",
        layout: "vertical",
        backgroundColor: "#1A1A2E",
        paddingAll: "14px",
        contents: [{ type: "text", text: headerTitle, color: "#FFFFFF", weight: "bold", size: "sm" }],
      },
      body: {
        type: "box",
        layout: "vertical",
        paddingAll: "12px",
        spacing: "sm",
        contents: rows.slice(0, 8).map((r, i) => buildLinkListRow(r, i)),
      },
      footer: {
        type: "box",
        layout: "vertical",
        paddingAll: "10px",
        contents: [{ type: "button", style: "link", height: "sm", action: { type: "message", label: "查看最近收藏", text: "最近收藏" } }],
      },
    },
  };
}

function buildLinkCardBubble(r) {
  const url = String(r[1] || "");
  const platform = detectPlatform(url);
  const note = String(r[4] || "").substring(0, 40) || "（未填備註）";
  const tags = String(r[5] || "");
  const category = String(r[6] || "未分類");
  const date = String(r[8] || r[7] || "").substring(0, 10);

  return {
    type: "bubble",
    size: "kilo",
    styles: { header: { backgroundColor: platform.color } },
    header: {
      type: "box",
      layout: "horizontal",
      paddingAll: "10px",
      contents: [
        { type: "text", text: platform.emoji, size: "sm", flex: 0, color: "#FFFFFF" },
        { type: "text", text: platform.name, size: "xs", flex: 1, color: "#FFFFFF", paddingStart: "6px", gravity: "center" },
        { type: "text", text: `${CATEGORY_EMOJI[category] || "📋"} ${category}`, size: "xs", flex: 0, color: "#FFFFFF99" },
      ],
    },
    body: {
      type: "box",
      layout: "vertical",
      paddingAll: "12px",
      spacing: "sm",
      contents: [
        { type: "text", text: note, size: "sm", color: "#333333", wrap: true, maxLines: 3 },
        tags ? { type: "text", text: `🏷 ${tags}`, size: "xs", color: "#888888" } : null,
        makeRow("日期", date),
      ].filter(Boolean),
    },
    footer: {
      type: "box",
      layout: "vertical",
      paddingAll: "10px",
      contents: [{ type: "button", style: "primary", height: "sm", color: platform.color || "#1A1A2E", action: { type: "uri", label: "開啟連結", uri: url } }],
    },
  };
}

function buildLinkListRow(r, i) {
  const url = String(r[1] || "");
  const platform = detectPlatform(url);
  const note = String(r[4] || "").substring(0, 30) || "（未填備註）";
  const category = String(r[6] || "未分類");
  const date = String(r[8] || r[7] || "").substring(0, 10);

  return {
    type: "box",
    layout: "vertical",
    spacing: "xs",
    paddingTop: i === 0 ? "0px" : "8px",
    contents: [
      {
        type: "box",
        layout: "horizontal",
        contents: [
          { type: "text", text: `${platform.emoji} ${note}`, size: "sm", color: "#333333", flex: 1, wrap: true, maxLines: 2 },
          { type: "button", style: "link", height: "sm", flex: 0, action: { type: "uri", label: "開啟", uri: url } },
        ],
      },
      { type: "text", text: `${CATEGORY_EMOJI[category] || "📋"} ${category}  📅 ${date}`, size: "xs", color: "#888888" },
      { type: "separator" },
    ],
  };
}

async function handleDeleteLast(replyToken, userId, sheets) {
  const { row: lastRow, index: lastIndex } = await findLastUserRow(sheets, SHEET.THREADS, userId);
  if (lastIndex === -1) {
    await replyText(replyToken, "找不到你最近的收藏。");
    return;
  }

  const url = String(lastRow[1] || "");
  const rowNum = lastIndex + 1;
  await sheets.spreadsheets.values.clear({
    spreadsheetId: SHEETS_ID,
    range: `${SHEET.THREADS}!A${rowNum}:K${rowNum}`,
  });

  await replyText(replyToken, `🗑 已刪除最近一筆\n${url}`);
}

async function checkDuplicateThreads(url, userId, sheets) {
  const data = await getSheetData(sheets, SHEET.THREADS);
  return data.slice(1).some((row) => String(row[1] || "") === url && String(row[7] || "") === userId);
}

// ============================================================
//  檔案備份
// ============================================================

async function handleFileMessage(replyToken, message, type, userId, clients) {
  const { sheets, drive } = clients;

  try {
    const fileBuffer = await downloadLineContent(message.id);
    const fileName = message.fileName || generateFileName(type, message.id);
    const mimeType = getMimeType(type, fileName);
    const fileType = classifyFileType(type, fileName);

    const driveFile = await uploadToDrive(drive, fileName, mimeType, fileBuffer, fileType);

    const id = generateId("F");
    const now = formatDateTime(new Date());
    const row = [id, fileName, fileType, mimeType, fileBuffer.length, driveFile.id, driveFile.webViewLink, "", "", userId, message.id, now];
    await appendRow(sheets, SHEET.FILES, row);

    await replyFlex(replyToken, buildFileSavedFlex(fileName, fileType, driveFile.webViewLink, now));
  } catch (err) {
    const msg = String(err?.message || err || "");

    if (/Google Drive API has not been used|accessNotConfigured|SERVICE_DISABLED/i.test(msg)) {
      await replyText(replyToken, "⚠️ Google Drive API 尚未完全啟用，請等 5～10 分鐘後再測試一次。");
      return;
    }

    if (/insufficient.*permission|forbidden|The caller does not have permission/i.test(msg)) {
      await replyText(replyToken, "⚠️ Google Drive 資料夾權限不足，請再確認 service account 是否已加入該資料夾並設為編輯者。");
      return;
    }

    await replyText(replyToken, `⚠️ 檔案備份失敗：${msg}`);
  }
}

function downloadLineContent(messageId) {
  return new Promise((resolve, reject) => {
    https.get(
      {
        hostname: "api-data.line.me",
        path: `/v2/bot/message/${messageId}/content`,
        method: "GET",
        headers: { Authorization: `Bearer ${LINE_TOKEN}` },
      },
      (res) => {
        if (res.statusCode && res.statusCode >= 400) {
          reject(new Error(`LINE content download failed: ${res.statusCode}`));
          return;
        }
        const chunks = [];
        res.on("data", (chunk) => chunks.push(chunk));
        res.on("end", () => resolve(Buffer.concat(chunks)));
        res.on("error", reject);
      }
    ).on("error", reject);
  });
}

async function uploadToDrive(drive, fileName, mimeType, buffer, fileType) {
  const folderId = await getOrCreateFolder(drive, fileType);
  const stream = Readable.from(buffer);

  const { data } = await drive.files.create({
    requestBody: { name: fileName, parents: [folderId] },
    media: { mimeType, body: stream },
    fields: "id,webViewLink,name",
  });

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
  const emoji = { image: "🖼", pdf: "📄", excel: "📊", word: "📝", video: "🎬", audio: "🎵", other: "📁" };
  return {
    type: "flex",
    altText: "📁 檔案備份完成",
    contents: {
      type: "bubble",
      styles: { header: { backgroundColor: "#1a73e8" } },
      header: {
        type: "box",
        layout: "vertical",
        paddingAll: "16px",
        contents: [{ type: "text", text: "📁 檔案備份完成", color: "#ffffff", weight: "bold", size: "md" }],
      },
      body: {
        type: "box",
        layout: "vertical",
        paddingAll: "16px",
        spacing: "md",
        contents: [
          makeRow("檔案名", fileName),
          makeRow("類型", `${emoji[fileType] || "📁"} ${fileType.toUpperCase()}`),
          makeRow("時間", now),
        ],
      },
      footer: {
        type: "box",
        layout: "vertical",
        paddingAll: "12px",
        contents: [{ type: "button", style: "primary", height: "sm", color: "#1a73e8", action: { type: "uri", label: "🔗 查看 Google Drive", uri: driveUrl } }],
      },
    },
  };
}

async function handleRecentFiles(replyToken, userId, sheets) {
  const data = await getSheetData(sheets, SHEET.FILES);
  const results = data.slice(1).filter((r) => String(r[9] || "") === userId).reverse().slice(0, 10);

  if (!results.length) {
    await replyFlex(
      replyToken,
      buildEmptyStateFlex("還沒有備份檔案", "傳送圖片、PDF、影片、語音或其他檔案就會自動備份", [{ label: "使用說明", text: "說明" }])
    );
    return;
  }

  const lines = results.map((r, i) => {
    const name = String(r[1] || "");
    const type = String(r[2] || "");
    const url = String(r[6] || "");
    const date = String(r[11] || "").substring(0, 10);
    return `${i + 1}. ${name}\n   📁 ${type}  📅 ${date}\n   🔗 ${url}`;
  });

  await replyText(replyToken, `📁 最近 ${results.length} 筆檔案\n\n${lines.join("\n\n")}`);
}

async function handleSearchFiles(replyToken, keyword, userId, sheets) {
  const data = await getSheetData(sheets, SHEET.FILES);
  const results = data
    .slice(1)
    .filter((row) => String(row[9] || "") === userId)
    .filter((row) => [row[1], row[2], row[7], row[8]].join(" ").includes(keyword))
    .reverse()
    .slice(0, 8);

  if (!results.length) {
    await replyFlex(
      replyToken,
      buildEmptyStateFlex("找不到符合的內容", `沒有找到包含「${keyword}」的檔案`, [{ label: "查看最近檔案", text: "最近檔案" }])
    );
    return;
  }

  const lines = results.map((r, i) => {
    const name = String(r[1] || "");
    const url = String(r[6] || "");
    const date = String(r[11] || "").substring(0, 10);
    return `${i + 1}. ${name}\n   📅 ${date}\n   🔗 ${url}`;
  });

  await replyText(replyToken, `🔍「${keyword}」共 ${results.length} 筆\n\n${lines.join("\n\n")}`);
}

// ============================================================
//  Google 日曆
// ============================================================

async function handleCalendar(replyToken, text, userId, calendar, sheets) {
  const parsed = parseCalendarInput(text);

  if (!parsed) {
    await replyText(
      replyToken,
      "⚠️ 我找不到日期，無法建立行程。\n\n請用這樣的格式：\n明天 下午3點 開會\n4/25 10:00 客戶簡報\n下週三 記得續約"
    );
    return;
  }

  try {
    const event = await createCalendarEvent(calendar, parsed);

    const id = generateId("C");
    const now = formatDateTime(new Date());
    const row = [
      id,
      parsed.title,
      formatDate(parsed.start),
      parsed.isAllDay ? "" : formatTime(parsed.start),
      parsed.isAllDay ? "TRUE" : "FALSE",
      event.id,
      text,
      "text",
      userId,
      now,
    ];
    await appendRow(sheets, SHEET.CALENDAR, row);

    await replyFlex(replyToken, buildCalendarFlex(parsed, event));
  } catch (err) {
    console.error("Calendar error:", err);
    await replyText(replyToken, "⚠️ 建立行程失敗：" + err.message);
  }
}

function parseCalendarInput(text) {
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  let targetDate = null;
  let targetTime = null;
  let title = text;

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
    const weekMatch = text.match(/下[週周]?(一|二|三|四|五|六|日|天)?/);
    if (weekMatch && weekMatch[1]) {
      const dayMap = { 一: 1, 二: 2, 三: 3, 四: 4, 五: 5, 六: 6, 日: 0, 天: 0 };
      const target = dayMap[weekMatch[1]];
      targetDate = new Date(now);
      let diff = (target - targetDate.getDay() + 7) % 7;
      if (diff === 0) diff = 7;
      diff += 7;
      targetDate.setDate(targetDate.getDate() + diff);
      title = title.replace(weekMatch[0], "");
    } else {
      const mdMatch = text.match(/(\d{1,2})[\/月](\d{1,2})[日]?/);
      if (mdMatch) {
        targetDate = new Date(now.getFullYear(), parseInt(mdMatch[1], 10) - 1, parseInt(mdMatch[2], 10));
        title = title.replace(mdMatch[0], "");
      }
    }
  }

  if (!targetDate) return null;

  const isPM = /下午|晚上|pm|傍晚/i.test(text);
  const isAM = /上午|早上|am/i.test(text);
  const timeMatch = text.match(/(\d{1,2})[:點時](\d{0,2})/);

  if (timeMatch) {
    let hour = parseInt(timeMatch[1], 10);
    const min = parseInt(timeMatch[2] || "0", 10);
    if (isPM && hour < 12) hour += 12;
    if (isAM && hour === 12) hour = 0;
    targetTime = { hour, min };
    title = title.replace(timeMatch[0], "");
  }

  title = title.replace(/上午|早上|下午|晚上|傍晚|pm|am/gi, "").replace(/\s+/g, " ").trim();
  if (!title) title = "行程待辦";

  const start = new Date(targetDate);
  const isAllDay = !targetTime;
  if (targetTime) start.setHours(targetTime.hour, targetTime.min, 0, 0);

  const end = new Date(start);
  if (isAllDay) end.setDate(end.getDate() + 1);
  else end.setHours(end.getHours() + 1);

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
    body.end = { date: formatDate(parsed.end) };
  } else {
    body.start = { dateTime: parsed.start.toISOString(), timeZone: "Asia/Taipei" };
    body.end = { dateTime: parsed.end.toISOString(), timeZone: "Asia/Taipei" };
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
    type: "flex",
    altText: `✅ 已新增：${parsed.title}`,
    contents: {
      type: "bubble",
      styles: { header: { backgroundColor: "#0F9D58" } },
      header: {
        type: "box",
        layout: "vertical",
        paddingAll: "16px",
        contents: [{ type: "text", text: "✅ 已新增到 Google 日曆", color: "#ffffff", weight: "bold", size: "md" }],
      },
      body: {
        type: "box",
        layout: "vertical",
        paddingAll: "16px",
        spacing: "md",
        contents: [
          makeRow("行程", parsed.title),
          makeRow("時間", timeStr),
          makeRow("提醒", parsed.isAllDay ? "（全天，無提醒）" : "開始前 10 分鐘"),
        ],
      },
      footer: event.htmlLink
        ? {
            type: "box",
            layout: "vertical",
            paddingAll: "12px",
            contents: [{ type: "button", style: "primary", height: "sm", color: "#0F9D58", action: { type: "uri", label: "📅 查看 Google 日曆", uri: event.htmlLink } }],
          }
        : undefined,
    },
  };
}

// ============================================================
//  主選單
// ============================================================

function buildMainMenuFlex() {
  return {
    type: "flex",
    altText: "你好！我可以幫你做這些事",
    contents: {
      type: "bubble",
      size: "mega",
      styles: { header: { backgroundColor: "#1A1A2E" } },
      header: {
        type: "box",
        layout: "vertical",
        paddingAll: "20px",
        contents: [
          { type: "text", text: "你好！我可以幫你做這些事 👋", color: "#FFFFFF", weight: "bold", size: "md" },
          { type: "text", text: "選擇功能或直接輸入", color: "#FFFFFF88", size: "xs" },
        ],
      },
      body: {
        type: "box",
        layout: "vertical",
        paddingAll: "16px",
        spacing: "lg",
        contents: [
          menuItem("🧵", "貼連結收藏", "Threads / FB / YouTube / IG 自動收藏，並可直接分類"),
          { type: "separator" },
          menuItem("📋", "待辦清單", "查看、新增、完成、刪除待辦"),
          { type: "separator" },
          menuItem("📁", "傳圖片／檔案", "自動備份到 Google Drive"),
          { type: "separator" },
          menuItem("📅", "說待辦事項", "新增到 Google 日曆"),
        ],
      },
      footer: {
        type: "box",
        layout: "vertical",
        paddingAll: "12px",
        spacing: "sm",
        contents: [
          {
            type: "box",
            layout: "horizontal",
            spacing: "sm",
            contents: CATEGORIES.filter((c) => c !== "未分類").map((cat) => ({
              type: "button",
              style: "secondary",
              height: "sm",
              flex: 1,
              action: {
                type: "postback",
                label: `${CATEGORY_EMOJI[cat]}${cat}`,
                data: `action=view_category&cat=${cat}`,
                displayText: `${cat}收藏`,
              },
            })),
          },
          {
            type: "box",
            layout: "horizontal",
            spacing: "sm",
            contents: [
              {
                type: "button",
                style: "secondary",
                height: "sm",
                flex: 1,
                action: { type: "message", label: "最近收藏", text: "最近收藏" },
              },
              {
                type: "button",
                style: "secondary",
                height: "sm",
                flex: 1,
                action: { type: "message", label: "查看待辦", text: "查看待辦" },
              },
              {
                type: "button",
                style: "primary",
                height: "sm",
                flex: 1,
                color: "#1A1A2E",
                action: { type: "message", label: "待辦清單", text: "待辦清單" },
              },
            ],
          },
        ],
      },
    },
  };
}

function menuItem(icon, title, desc) {
  return {
    type: "box",
    layout: "horizontal",
    spacing: "md",
    contents: [
      { type: "text", text: icon, size: "xl", flex: 0, gravity: "center" },
      {
        type: "box",
        layout: "vertical",
        flex: 1,
        contents: [
          { type: "text", text: title, size: "sm", weight: "bold", color: "#1A1A2E" },
          { type: "text", text: desc, size: "xs", color: "#888888", wrap: true },
        ],
      },
    ],
  };
}

function buildEmptyStateFlex(title, description, buttons = []) {
  return {
    type: "flex",
    altText: title,
    contents: {
      type: "bubble",
      body: {
        type: "box",
        layout: "vertical",
        paddingAll: "24px",
        spacing: "md",
        contents: [
          { type: "text", text: title, weight: "bold", size: "md", color: "#333333", align: "center" },
          { type: "text", text: description, size: "sm", color: "#888888", wrap: true, align: "center" },
        ],
      },
      footer:
        buttons.length > 0
          ? {
              type: "box",
              layout: "vertical",
              paddingAll: "12px",
              spacing: "sm",
              contents: buttons.map((b) => ({
                type: "button",
                style: "secondary",
                height: "sm",
                action: { type: "message", label: b.label, text: b.text },
              })),
            }
          : undefined,
    },
  };
}

function makeRow(label, value) {
  return {
    type: "box",
    layout: "horizontal",
    contents: [
      { type: "text", text: label, size: "sm", color: "#888888", flex: 1 },
      { type: "text", text: String(value || ""), size: "sm", color: "#333333", flex: 3, wrap: true },
    ],
  };
}

// ============================================================
//  工具函式
// ============================================================

async function getSheetData(sheets, sheetName) {
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEETS_ID,
    range: `${sheetName}!A:Z`,
  });
  return data.values || [];
}

async function appendRow(sheets, sheetName, row) {
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEETS_ID,
    range: `${sheetName}!A:Z`,
    valueInputOption: "RAW",
    requestBody: { values: [row] },
  });
}

async function findLastUserRow(sheets, sheetName, userId) {
  const data = await getSheetData(sheets, sheetName);
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    if (String(row[7] || row[9] || "") === userId) {
      return { row, index: i };
    }
  }
  return { row: null, index: -1 };
}

function generateId(prefix) {
  const ts = new Date().toISOString().replace(/[:\-T.Z]/g, "").substring(0, 14);
  const rnd = Math.floor(Math.random() * 1000).toString().padStart(3, "0");
  return `${prefix}${ts}${rnd}`;
}

function formatDateTime(date) {
  return new Date(date).toLocaleString("zh-TW", {
    timeZone: "Asia/Taipei",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function formatDate(date) {
  const d = new Date(date);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function formatTime(date) {
  const d = new Date(date);
  return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
}

function generateFileName(type, messageId) {
  const ext = { image: "jpg", video: "mp4", audio: "m4a", file: "bin" };
  const now = new Date().toISOString().substring(0, 10).replace(/-/g, "");
  return `${type}_${now}_${messageId}.${ext[type] || "bin"}`;
}

function getMimeType(type, fileName) {
  const ext = path.extname(fileName || "").toLowerCase();
  const map = {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".gif": "image/gif",
    ".webp": "image/webp",
    ".pdf": "application/pdf",
    ".doc": "application/msword",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".xls": "application/vnd.ms-excel",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".ppt": "application/vnd.ms-powerpoint",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".mp4": "video/mp4",
    ".mov": "video/quicktime",
    ".m4a": "audio/m4a",
    ".mp3": "audio/mpeg",
    ".wav": "audio/wav",
    ".txt": "text/plain",
    ".zip": "application/zip",
  };

  if (map[ext]) return map[ext];
  if (type === "image") return "image/jpeg";
  if (type === "video") return "video/mp4";
  if (type === "audio") return "audio/m4a";
  return "application/octet-stream";
}

function classifyFileType(type, fileName) {
  const ext = path.extname(fileName || "").toLowerCase();
  if (type === "image") return "image";
  if (type === "video") return "video";
  if (type === "audio") return "audio";
  if (ext === ".pdf") return "pdf";
  if ([".xls", ".xlsx", ".csv"].includes(ext)) return "excel";
  if ([".doc", ".docx"].includes(ext)) return "word";
  return "other";
}

function replyText(replyToken, text) {
  return replyMessages(replyToken, [{ type: "text", text }]);
}

function replyFlex(replyToken, flexMessage) {
  return replyMessages(replyToken, [flexMessage]);
}

function replyMessages(replyToken, messages) {
  return lineApiRequest("/v2/bot/message/reply", {
    replyToken,
    messages,
  });
}

function lineApiRequest(apiPath, payload) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify(payload);
    const req = https.request(
      {
        hostname: "api.line.me",
        path: apiPath,
        method: "POST",
        headers: {
          Authorization: `Bearer ${LINE_TOKEN}`,
          "Content-Type": "application/json",
          "Content-Length": Buffer.byteLength(body),
        },
      },
      (res) => {
        let data = "";
        res.on("data", (chunk) => (data += chunk));
        res.on("end", () => {
          if (res.statusCode >= 200 && res.statusCode < 300) {
            resolve(data ? JSON.parse(data || "{}") : {});
          } else {
            reject(new Error(`LINE API error ${res.statusCode}: ${data}`));
          }
        });
      }
    );

    req.on("error", reject);
    req.write(body);
    req.end();
  });
}
