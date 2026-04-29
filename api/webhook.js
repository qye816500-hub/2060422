// api/webhook.js
// LINE Bot - v7.3.0
// bot state + any URL bookmark + todo direct input + OAuth Drive

const { google } = require("googleapis");
const https = require("https");
const path = require("path");
const { Readable } = require("stream");

const LINE_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN;
const GOOGLE_CREDENTIALS = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
const SHEETS_ID = process.env.GOOGLE_SHEETS_ID;
const DRIVE_FOLDER_ID = process.env.GOOGLE_DRIVE_FOLDER_ID;
const CALENDAR_ID = process.env.GOOGLE_CALENDAR_ID || "primary";

const GOOGLE_CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const GOOGLE_CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const GOOGLE_REFRESH_TOKEN = process.env.GOOGLE_REFRESH_TOKEN;
const REDIRECT_URI = "https://2060422.vercel.app/api/auth/callback";

const SHEET = {
  THREADS: "Threads收藏",
  FILES: "檔案備份",
  CALENDAR: "日曆待辦",
  TODOS: "待辦清單",
  BOT_STATE: "機器人狀態",
  CATEGORIES: "收藏分類",
};

const DEFAULT_CATEGORIES = ["未分類", "個人", "工作", "家庭"];
const CATEGORY_COLORS = {
  "未分類": "#888888",
  "個人": "#FF6B6B",
  "工作": "#4ECDC4",
  "家庭": "#45B7D1",
};

// ============================================================
//  Google Clients
// ============================================================

async function getGoogleClients() {
  const credentials = JSON.parse(GOOGLE_CREDENTIALS);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/calendar",
    ],
  });
  const authClient = await auth.getClient();
  return {
    sheets: google.sheets({ version: "v4", auth: authClient }),
    calendar: google.calendar({ version: "v3", auth: authClient }),
  };
}

function getOAuthDriveClient() {
  const oauth2Client = new google.auth.OAuth2(GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, REDIRECT_URI);
  oauth2Client.setCredentials({ refresh_token: GOOGLE_REFRESH_TOKEN });
  return google.drive({ version: "v3", auth: oauth2Client });
}

// ============================================================
//  Dynamic Categories
// ============================================================

async function getUserCategories(sheets, userId) {
  try {
    const data = await getSheetData(sheets, SHEET.CATEGORIES);
    const custom = data.slice(1)
      .filter(function(r) { return String(r[0] || "") === userId; })
      .map(function(r) { return String(r[1] || ""); })
      .filter(Boolean);
    const all = [...DEFAULT_CATEGORIES];
    custom.forEach(function(c) { if (!all.includes(c)) all.push(c); });
    return all;
  } catch(e) {
    return [...DEFAULT_CATEGORIES];
  }
}

// ============================================================
//  Bot State
// ============================================================

async function getBotState(sheets, userId) {
  const data = await getSheetData(sheets, SHEET.BOT_STATE);
  const row = data.slice(1).find(function(r) { return String(r[0] || "") === userId; });
  return row ? String(row[1] || "") : "";
}

async function setBotState(sheets, userId, state) {
  const data = await getSheetData(sheets, SHEET.BOT_STATE);
  const now = formatDateTime(new Date());
  const rowIndex = data.slice(1).findIndex(function(r) { return String(r[0] || "") === userId; });

  if (rowIndex === -1) {
    await appendRow(sheets, SHEET.BOT_STATE, [userId, state, now]);
  } else {
    const rowNum = rowIndex + 2;
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SHEETS_ID,
      requestBody: {
        valueInputOption: "RAW",
        data: [
          { range: SHEET.BOT_STATE + "!B" + rowNum, values: [[state]] },
          { range: SHEET.BOT_STATE + "!C" + rowNum, values: [[now]] },
        ],
      },
    });
  }
}

async function clearBotState(sheets, userId) {
  await setBotState(sheets, userId, "");
}

// ============================================================
//  Main Entry
// ============================================================

module.exports = async function(req, res) {
  if (req.method === "GET") {
    return res.status(200).json({ status: "ok", version: "7.3.0" });
  }
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
    if (!body || !body.events || !body.events.length) {
      return res.status(200).json({ status: "ok" });
    }
    const clients = await getGoogleClients();
    for (let i = 0; i < body.events.length; i++) {
      await handleEvent(body.events[i], clients);
    }
  } catch (err) {
    console.error("Global error:", err);
  }

  return res.status(200).json({ status: "ok" });
};

async function handleEvent(event, clients) {
  const replyToken = event.replyToken;
  const type = event.type;
  const message = event.message;
  const source = event.source;
  const postback = event.postback;
  const userId = source && source.userId ? source.userId : "unknown";

  try {
    if (type === "follow") {
      await replyFlex(replyToken, buildMainMenuFlex());
      return;
    }
    if (type === "postback" && postback && postback.data) {
      await handlePostback(replyToken, postback.data, userId, clients);
      return;
    }
    if (type !== "message" || !message) return;

    if (message.type === "text") {
      await handleTextMessage(replyToken, message.text.trim(), userId, clients);
      return;
    }
    if (message.type === "image" || message.type === "video" || message.type === "file" || message.type === "audio") {
      await handleFileMessage(replyToken, message, message.type, userId, clients);
      return;
    }
  } catch (err) {
    console.error("handleEvent error:", err);
    if (replyToken) {
      await replyText(replyToken, "ERROR: " + err.message);
    }
  }
}

// ============================================================
//  Postback
// ============================================================

async function handlePostback(replyToken, data, userId, clients) {
  const sheets = clients.sheets;
  const params = new URLSearchParams(data);
  const action = params.get("action");

  if (action === "add_tag_prompt") {
    await setBotState(sheets, userId, "awaiting_tag_input");
    await replyText(replyToken, "請輸入要加的標籤（可多個，用空格分隔）：\n\n例如：旅遊 美食 推薦");
    return;
  }
  if (action === "set_category") {
    await handleSetCategory(replyToken, params.get("cat"), userId, sheets);
    return;
  }
  if (action === "view_category") {
    await handleQueryCategory(replyToken, params.get("cat"), userId, sheets);
    return;
  }
  if (action === "todo_new") {
    await setBotState(sheets, userId, "awaiting_todo_input");
    await replyText(replyToken, "請直接輸入待辦內容：\n\n例如：整理報價單\n或：明天下午3點提醒我回覆客戶");
    return;
  }
  if (action === "todo_view") {
    await handleShowTodos(replyToken, userId, sheets);
    return;
  }
  if (action === "todo_complete") {
    await handlePromptCompleteTodo(replyToken, userId, sheets);
    return;
  }
  if (action === "todo_delete") {
    await handlePromptDeleteTodo(replyToken, userId, sheets);
    return;
  }
}

// ============================================================
//  Text Message Handler
// ============================================================

async function handleTextMessage(replyToken, text, userId, clients) {
  const sheets = clients.sheets;
  const calendar = clients.calendar;

  const state = await getBotState(sheets, userId);

  if (state === "awaiting_todo_input") {
    await clearBotState(sheets, userId);
    await handleCreateTodo(replyToken, text, userId, sheets, calendar);
    return;
  }

  if (state === "awaiting_tag_input") {
    await clearBotState(sheets, userId);
    await handleAddTag(replyToken, "加標籤 " + text, userId, sheets);
    return;
  }

  if (state === "awaiting_todo_complete") {
    const num = parseInt(text.trim(), 10);
    if (!isNaN(num)) {
      await clearBotState(sheets, userId);
      await handleCompleteTodoByIndex(replyToken, num, userId, sheets);
      return;
    }
    await replyText(replyToken, "請輸入數字序號，例如：1\n\n或輸入「取消」離開");
    return;
  }

  if (state === "awaiting_todo_delete") {
    const num = parseInt(text.trim(), 10);
    if (!isNaN(num)) {
      await clearBotState(sheets, userId);
      await handleDeleteTodoByIndex(replyToken, num, userId, sheets);
      return;
    }
    await replyText(replyToken, "請輸入數字序號，例如：1\n\n或輸入「取消」離開");
    return;
  }

  if (text === "取消") {
    await clearBotState(sheets, userId);
    await replyText(replyToken, "已取消操作。");
    return;
  }

  if (text === "功能說明" || text === "說明" || text === "help") {
    await replyFlex(replyToken, buildMainMenuFlex());
    return;
  }

  if (text === "待辦清單") {
    await replyFlex(replyToken, buildTodoMenuFlex());
    return;
  }
  if (text === "查看待辦") {
    await handleShowTodos(replyToken, userId, sheets);
    return;
  }
  if (text === "新增待辦") {
    await setBotState(sheets, userId, "awaiting_todo_input");
    await replyText(replyToken, "請直接輸入待辦內容：\n\n例如：整理報價單\n或：明天下午3點提醒我回覆客戶");
    return;
  }
  if (text.indexOf("新增待辦 ") === 0) {
    const content = text.slice(5).trim();
    await handleCreateTodo(replyToken, content, userId, sheets, calendar);
    return;
  }
  if (text === "完成待辦") {
    await handlePromptCompleteTodo(replyToken, userId, sheets);
    return;
  }
  if (/^完成待辦\s+\d+$/.test(text)) {
    const index = parseInt(text.replace(/^完成待辦\s+/, ""), 10);
    await handleCompleteTodoByIndex(replyToken, index, userId, sheets);
    return;
  }
  if (text === "刪除待辦") {
    await handlePromptDeleteTodo(replyToken, userId, sheets);
    return;
  }
  if (/^刪除待辦\s+\d+$/.test(text)) {
    const index = parseInt(text.replace(/^刪除待辦\s+/, ""), 10);
    await handleDeleteTodoByIndex(replyToken, index, userId, sheets);
    return;
  }

  // 分類查詢（支援動態分類）
  const userCats = await getUserCategories(sheets, userId);
  for (let i = 0; i < userCats.length; i++) {
    const cat = userCats[i];
    if (text === cat + "收藏" || text === cat) {
      await handleQueryCategory(replyToken, cat, userId, sheets);
      return;
    }
  }

  const urls = extractUrls(text);
  if (urls.length > 0) {
    await handleLinkSave(replyToken, text, urls, userId, sheets);
    return;
  }

  // 加標籤 - 修正：只匹配有內容的標籤指令
  if (/^(標籤|加標籤)\s+\S/.test(text)) {
    await handleAddTag(replyToken, text, userId, sheets);
    return;
  }
  if (text.indexOf("查標籤 ") === 0) {
    const tag = text.slice(4).trim();
    await handleQueryTag(replyToken, tag, userId, sheets);
    return;
  }
  if (text.charAt(0) === "#" && text.length > 1) {
    const tag = text.slice(1).trim();
    await handleQueryTag(replyToken, tag, userId, sheets);
    return;
  }

  if (text.indexOf("搜尋 ") === 0 || text.indexOf("找 ") === 0) {
    const keyword = text.indexOf("搜尋 ") === 0 ? text.slice(3).trim() : text.slice(2).trim();
    await handleSearchLinks(replyToken, keyword, userId, sheets);
    return;
  }

  if (text === "最近收藏") {
    await handleRecentLinks(replyToken, userId, sheets);
    return;
  }
  if (text === "最近檔案") {
    await handleRecentFiles(replyToken, userId, sheets);
    return;
  }

  if (/明天|行程|提醒|會議|今天|下午|上午|早上|後天|下週|下周|週一|週二|週三|週四|週五|週六|週日/.test(text)) {
    await handleCalendar(replyToken, text, userId, calendar, sheets);
    return;
  }

  await replyFlex(replyToken, buildMainMenuFlex());
}

// ============================================================
//  Todo
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
          { type: "text", text: "待辦功能選單", color: "#FFFFFF", weight: "bold", size: "md" },
          { type: "text", text: "點下方按鈕操作", color: "#FFFFFFCC", size: "xs" },
        ],
      },
      body: {
        type: "box",
        layout: "vertical",
        paddingAll: "16px",
        spacing: "md",
        contents: [
          { type: "text", text: "新增待辦範例", weight: "bold", size: "sm", color: "#333333" },
          { type: "text", text: "整理報價單", size: "sm", color: "#666666", wrap: true },
          { type: "text", text: "明天下午3點提醒我回覆客戶", size: "sm", color: "#666666", wrap: true },
          { type: "text", text: "點「新增待辦」後直接輸入內容即可", size: "xs", color: "#999999", wrap: true },
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
              { type: "button", style: "secondary", height: "sm", flex: 1, action: { type: "postback", label: "查看待辦", data: "action=todo_view", displayText: "查看待辦" } },
              { type: "button", style: "primary", height: "sm", flex: 1, color: "#2D3561", action: { type: "postback", label: "新增待辦", data: "action=todo_new", displayText: "新增待辦" } },
            ],
          },
          {
            type: "box",
            layout: "horizontal",
            spacing: "sm",
            contents: [
              { type: "button", style: "secondary", height: "sm", flex: 1, action: { type: "postback", label: "完成待辦", data: "action=todo_complete", displayText: "完成待辦" } },
              { type: "button", style: "secondary", height: "sm", flex: 1, action: { type: "postback", label: "刪除待辦", data: "action=todo_delete", displayText: "刪除待辦" } },
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
    .filter(function(r) { return String(r[6] || "") === "todo" && String(r[5] || "") === userId; })
    .filter(function(r) { return String(r[2] || "pending") !== "deleted"; });
}

async function handleCreateTodo(replyToken, content, userId, sheets, calendar) {
  const id = generateId("TD");
  const now = formatDateTime(new Date());

  const parsed = parseCalendarInput(content);
  let remindText = "";
  let googleEventId = "";

  if (parsed) {
    remindText = parsed.isAllDay ? formatDate(parsed.start) : formatDate(parsed.start) + " " + formatTime(parsed.start);
    try {
      const event = await createCalendarEvent(calendar, parsed);
      googleEventId = event.id || "";
    } catch (e) {
      console.error("Calendar create error:", e);
    }
  }

  const row = [id, content, "pending", now, remindText, userId, "todo", googleEventId];
  await appendRow(sheets, SHEET.TODOS, row);

  const bodyContents = [
    { type: "text", text: "已新增待辦", weight: "bold", size: "md", color: "#0F9D58" },
    { type: "text", text: content, size: "sm", color: "#333333", wrap: true },
  ];
  if (remindText) bodyContents.push(makeRow("提醒", remindText));
  if (googleEventId) bodyContents.push({ type: "text", text: "已同步到 Google 日曆", size: "xs", color: "#0F9D58" });

  await replyFlex(replyToken, {
    type: "flex",
    altText: "已新增待辦",
    contents: {
      type: "bubble",
      body: { type: "box", layout: "vertical", paddingAll: "16px", spacing: "md", contents: bodyContents },
      footer: {
        type: "box", layout: "horizontal", spacing: "sm", paddingAll: "12px",
        contents: [
          { type: "button", style: "secondary", height: "sm", flex: 1, action: { type: "postback", label: "查看待辦", data: "action=todo_view", displayText: "查看待辦" } },
          { type: "button", style: "primary", height: "sm", flex: 1, color: "#2D3561", action: { type: "postback", label: "再新增一筆", data: "action=todo_new", displayText: "新增待辦" } },
        ],
      },
    },
  });
}

async function handleShowTodos(replyToken, userId, sheets) {
  const rows = (await getTodoRows(sheets, userId)).filter(function(r) { return String(r[2] || "pending") === "pending"; });
  if (!rows.length) {
    await replyText(replyToken, "目前沒有待辦。\n\n輸入「新增待辦」新增第一筆");
    return;
  }
  const lines = rows.slice(0, 10).map(function(r, i) {
    const content = String(r[1] || "");
    const remind = String(r[4] || "");
    return remind ? (i + 1) + ". " + content + "\n   [" + remind + "]" : (i + 1) + ". " + content;
  });
  await replyText(replyToken, "目前待辦（" + rows.length + " 筆）\n\n" + lines.join("\n\n") + "\n\n完成請輸入：完成待辦 序號");
}

async function handlePromptCompleteTodo(replyToken, userId, sheets) {
  const rows = (await getTodoRows(sheets, userId)).filter(function(r) { return String(r[2] || "pending") === "pending"; });
  if (!rows.length) { await replyText(replyToken, "目前沒有可完成的待辦。"); return; }
  const lines = rows.slice(0, 10).map(function(r, i) {
    const remind = String(r[4] || "");
    return remind ? (i + 1) + ". " + String(r[1] || "") + " (" + remind + ")" : (i + 1) + ". " + String(r[1] || "");
  });
  await replyText(replyToken, "請輸入要完成的序號：\n\n" + lines.join("\n") + "\n\n例如輸入：完成待辦 1");
}

async function handleCompleteTodoByIndex(replyToken, index, userId, sheets) {
  const allData = await getSheetData(sheets, SHEET.TODOS);
  const pendingRows = allData.map(function(row, idx) { return { row: row, idx: idx }; }).filter(function(item) {
    return String(item.row[6] || "") === "todo" && String(item.row[5] || "") === userId && String(item.row[2] || "pending") === "pending";
  });

  if (!index || index < 1 || index > pendingRows.length) {
    await replyText(replyToken, "序號無效，請輸入 1~" + pendingRows.length + " 之間的數字。");
    return;
  }

  const target = pendingRows[index - 1];
  const rowNum = target.idx + 1;
  const content = String(target.row[1] || "");

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEETS_ID,
    range: SHEET.TODOS + "!C" + rowNum,
    valueInputOption: "RAW",
    requestBody: { values: [["done"]] },
  });

  await replyText(replyToken, "已完成待辦\n" + content);
}

async function handlePromptDeleteTodo(replyToken, userId, sheets) {
  const rows = (await getTodoRows(sheets, userId)).filter(function(r) { return String(r[2] || "pending") === "pending"; });
  if (!rows.length) { await replyText(replyToken, "目前沒有可刪除的待辦。"); return; }
  const lines = rows.slice(0, 10).map(function(r, i) {
    const remind = String(r[4] || "");
    return remind ? (i + 1) + ". " + String(r[1] || "") + " (" + remind + ")" : (i + 1) + ". " + String(r[1] || "");
  });
  await replyText(replyToken, "請輸入要刪除的序號：\n\n" + lines.join("\n") + "\n\n例如輸入：刪除待辦 1");
}

async function handleDeleteTodoByIndex(replyToken, index, userId, sheets) {
  const allData = await getSheetData(sheets, SHEET.TODOS);
  const pendingRows = allData.map(function(row, idx) { return { row: row, idx: idx }; }).filter(function(item) {
    return String(item.row[6] || "") === "todo" && String(item.row[5] || "") === userId && String(item.row[2] || "pending") === "pending";
  });

  if (!index || index < 1 || index > pendingRows.length) {
    await replyText(replyToken, "序號無效，請輸入 1~" + pendingRows.length + " 之間的數字。");
    return;
  }

  const target = pendingRows[index - 1];
  const rowNum = target.idx + 1;
  const content = String(target.row[1] || "");

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEETS_ID,
    range: SHEET.TODOS + "!C" + rowNum,
    valueInputOption: "RAW",
    requestBody: { values: [["deleted"]] },
  });

  await replyText(replyToken, "已刪除待辦\n" + content);
}

// ============================================================
//  Bookmarks
// ============================================================

function extractUrls(text) {
  const regex = /https?:\/\/[^\s]+/gi;
  return text.match(regex) || [];
}

function detectPlatform(url) {
  if (/threads\.(net|com)/i.test(url)) return { name: "Threads", color: "#000000" };
  if (/facebook\.com|fb\.com/i.test(url)) return { name: "Facebook", color: "#1877F2" };
  if (/youtube\.com|youtu\.be/i.test(url)) return { name: "YouTube", color: "#FF0000" };
  if (/instagram\.com/i.test(url)) return { name: "Instagram", color: "#E1306C" };
  if (/twitter\.com|x\.com/i.test(url)) return { name: "X", color: "#000000" };
  if (/linkedin\.com/i.test(url)) return { name: "LinkedIn", color: "#0077B5" };
  if (/github\.com/i.test(url)) return { name: "GitHub", color: "#333333" };
  if (/notion\.so/i.test(url)) return { name: "Notion", color: "#000000" };
  if (/medium\.com/i.test(url)) return { name: "Medium", color: "#000000" };
  return { name: "Link", color: "#555555" };
}

async function handleLinkSave(replyToken, rawText, urls, userId, sheets) {
  let note = rawText;
  urls.forEach(function(u) { note = note.replace(u, "").trim(); });
  note = note.replace(/^[\s\-]+/, "").trim();

  const saved = [];
  for (let i = 0; i < urls.length; i++) {
    const url = urls[i];
    const isDup = await checkDuplicateThreads(url, userId, sheets);
    if (isDup) { saved.push({ url: url, duplicate: true }); continue; }
    const platform = detectPlatform(url);
    const id = generateId("T");
    const now = formatDateTime(new Date());
    const row = [id, url, platform.name, "", note, "", "未分類", userId, now, now, rawText];
    await appendRow(sheets, SHEET.THREADS, row);
    saved.push({ url: url, platform: platform, duplicate: false, note: note, id: id });
  }

  if (saved.length === 1 && !saved[0].duplicate) {
    // 動態讀取分類
    const userCats = await getUserCategories(sheets, userId);
    await replyFlex(replyToken, buildLinkSavedFlex(saved[0].url, saved[0].platform, saved[0].note, userCats));
    return;
  }
  if (saved.length === 1 && saved[0].duplicate) {
    await replyText(replyToken, "這則連結之前已經收藏過。");
    return;
  }
  const lines = saved.map(function(s, i) {
    return s.duplicate ? (i + 1) + ". 已收藏過" : (i + 1) + ". [" + s.platform.name + "] 已收藏";
  });
  await replyText(replyToken, "收藏 " + urls.length + " 則連結\n\n" + lines.join("\n"));
}

// ★ 修正1：動態分類按鈕 + 修正2：加標籤按鈕不再有尾隨空格
function buildLinkSavedFlex(url, platform, note, userCats) {
  const now = formatDateTime(new Date());
  // 排除「未分類」，最多顯示5個分類按鈕（LINE Flex 限制）
  const catButtons = (userCats || DEFAULT_CATEGORIES)
    .filter(function(c) { return c !== "未分類"; })
    .slice(0, 5)
    .map(function(cat) {
      return {
        type: "button", style: "secondary", height: "sm", flex: 1,
        action: { type: "postback", label: cat, data: "action=set_category&id=LAST&cat=" + cat, displayText: "設為" + cat }
      };
    });

  // 超過3個分成兩排
  let catRows = [];
  if (catButtons.length <= 3) {
    catRows = [{ type: "box", layout: "horizontal", spacing: "sm", contents: catButtons }];
  } else {
    catRows = [
      { type: "box", layout: "horizontal", spacing: "sm", contents: catButtons.slice(0, 3) },
      { type: "box", layout: "horizontal", spacing: "sm", contents: catButtons.slice(3) },
    ];
  }

  return {
    type: "flex",
    altText: "已收藏成功",
    contents: {
      type: "bubble",
      size: "mega",
      header: {
        type: "box", layout: "vertical", paddingAll: "16px", backgroundColor: "#1A1A2E",
        contents: [
          { type: "text", text: "已收藏成功", color: "#FFFFFF", weight: "bold", size: "md" },
          { type: "text", text: platform.name, color: "#FFFFFFCC", size: "sm" },
        ],
      },
      body: {
        type: "box", layout: "vertical", paddingAll: "16px", spacing: "md",
        contents: [
          makeRow("備註", note || "(未填寫)"),
          makeRow("分類", "未分類"),
          makeRow("時間", now),
          { type: "separator" },
          { type: "text", text: "請選擇收藏分類", size: "sm", color: "#555555", weight: "bold" },
          ...catRows,
        ],
      },
      footer: {
        type: "box", layout: "horizontal", spacing: "sm", paddingAll: "12px",
        contents: [
          { type: "button", style: "primary", flex: 2, color: "#1A1A2E", action: { type: "uri", label: "開啟連結", uri: url } },
          // ★ 修正2：加標籤按鈕改用 postback，避免空白字串觸發主選單
          { type: "button", style: "secondary", flex: 1, action: { type: "postback", label: "加標籤", data: "action=add_tag_prompt", displayText: "加標籤" } },
        ],
      },
    },
  };
}

// ★ 修正2：加標籤 postback 處理
// 在 handlePostback 裡加上這個 action：
// 已在 handlePostback 函數中補上 add_tag_prompt 處理

async function handleSetCategory(replyToken, category, userId, sheets) {
  const result = await findLastUserRow(sheets, SHEET.THREADS, userId);
  const lastRow = result.row;
  const lastIndex = result.index;
  if (lastIndex === -1) { await replyText(replyToken, "找不到最近的收藏。"); return; }

  const rowNum = lastIndex + 1;
  const now = formatDateTime(new Date());
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SHEETS_ID,
    requestBody: {
      valueInputOption: "RAW",
      data: [
        { range: SHEET.THREADS + "!G" + rowNum, values: [[category]] },
        { range: SHEET.THREADS + "!J" + rowNum, values: [[now]] },
      ],
    },
  });

  const url = String(lastRow[1] || "");
  const catColor = CATEGORY_COLORS[category] || "#1A1A2E";
  await replyFlex(replyToken, {
    type: "flex",
    altText: "已設為" + category,
    contents: {
      type: "bubble",
      body: {
        type: "box", layout: "vertical", paddingAll: "20px", spacing: "md",
        contents: [
          { type: "text", text: "已歸類到「" + category + "」", weight: "bold", size: "md", color: "#333333" },
          { type: "text", text: "之後可用「" + category + "收藏」查詢", size: "xs", color: "#888888" },
        ],
      },
      footer: {
        type: "box", layout: "horizontal", spacing: "sm", paddingAll: "12px",
        contents: [
          { type: "button", style: "primary", height: "sm", flex: 1, color: catColor, action: { type: "postback", label: "查看" + category, data: "action=view_category&cat=" + category, displayText: category + "收藏" } },
          { type: "button", style: "secondary", height: "sm", flex: 1, action: { type: "uri", label: "開啟連結", uri: url } },
        ],
      },
    },
  });
}

async function handleQueryCategory(replyToken, category, userId, sheets) {
  const data = await getSheetData(sheets, SHEET.THREADS);
  const all = data.slice(1).filter(function(row) { return String(row[7] || "") === userId && String(row[6] || "未分類") === category; }).reverse();
  const results = all.slice(0, 10);

  if (!results.length) {
    await replyText(replyToken, category + " 還沒有收藏。\n\n收藏連結後，點成功卡片下方的分類按鈕即可歸類。");
    return;
  }
  await replyFlex(replyToken, buildLinksCarousel(results, category + " - " + all.length + " 筆"));
}

async function handleAddTag(replyToken, text, userId, sheets) {
  const tagStr = text.replace(/^(標籤|加標籤)\s+/, "").trim();
  const tags = tagStr.split(/[\s,]+/).filter(function(t) { return t.length > 0; });
  const result = await findLastUserRow(sheets, SHEET.THREADS, userId);
  if (result.index === -1) { await replyText(replyToken, "找不到最近的收藏。"); return; }

  const tagValue = tags.join("、");
  const rowNum = result.index + 1;
  const now = formatDateTime(new Date());
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SHEETS_ID,
    requestBody: {
      valueInputOption: "RAW",
      data: [
        { range: SHEET.THREADS + "!F" + rowNum, values: [[tagValue]] },
        { range: SHEET.THREADS + "!J" + rowNum, values: [[now]] },
      ],
    },
  });
  await replyText(replyToken, "標籤已更新\n" + tags.map(function(t) { return "#" + t; }).join(" "));
}

async function handleQueryTag(replyToken, tag, userId, sheets) {
  if (!tag) { await replyText(replyToken, "請告訴我要查詢的標籤"); return; }
  const data = await getSheetData(sheets, SHEET.THREADS);
  const all = data.slice(1).filter(function(row) { return String(row[7] || "") === userId && String(row[5] || "").includes(tag); });
  const results = all.reverse().slice(0, 10);
  if (!results.length) { await replyText(replyToken, "沒有標籤「" + tag + "」的收藏。"); return; }
  await replyFlex(replyToken, buildLinksCarousel(results, "標籤 " + tag + " - " + all.length + " 筆"));
}

async function handleRecentLinks(replyToken, userId, sheets) {
  const data = await getSheetData(sheets, SHEET.THREADS);
  const results = data.slice(1).filter(function(r) { return String(r[7] || "") === userId; }).reverse().slice(0, 10);
  if (!results.length) { await replyText(replyToken, "還沒有收藏。\n\n貼上任意連結就會自動收藏。"); return; }
  await replyFlex(replyToken, buildLinksCarousel(results, "最近 " + results.length + " 則收藏"));
}

async function handleSearchLinks(replyToken, keyword, userId, sheets) {
  const data = await getSheetData(sheets, SHEET.THREADS);
  const results = data.slice(1)
    .filter(function(row) { return String(row[7] || "") === userId; })
    .filter(function(row) { return [row[1], row[4], row[5], row[6], row[10]].join(" ").includes(keyword); })
    .reverse().slice(0, 8);
  if (!results.length) { await replyText(replyToken, "沒有找到「" + keyword + "」的收藏。"); return; }
  await replyFlex(replyToken, buildLinksCarousel(results, keyword + " - " + results.length + " 筆"));
}

function buildLinksCarousel(rows, headerTitle) {
  if (rows.length <= 5) {
    return { type: "flex", altText: headerTitle, contents: { type: "carousel", contents: rows.map(function(r) { return buildLinkCardBubble(r); }) } };
  }
  return {
    type: "flex", altText: headerTitle,
    contents: {
      type: "bubble", size: "mega",
      header: { type: "box", layout: "vertical", backgroundColor: "#1A1A2E", paddingAll: "14px", contents: [{ type: "text", text: headerTitle, color: "#FFFFFF", weight: "bold", size: "sm" }] },
      body: { type: "box", layout: "vertical", paddingAll: "12px", spacing: "sm", contents: rows.slice(0, 8).map(function(r, i) { return buildLinkListRow(r, i); }) },
      footer: { type: "box", layout: "vertical", paddingAll: "10px", contents: [{ type: "button", style: "link", height: "sm", action: { type: "message", label: "查看最近收藏", text: "最近收藏" } }] },
    },
  };
}

function buildLinkCardBubble(r) {
  const url = String(r[1] || "");
  const platform = detectPlatform(url);
  const note = String(r[4] || "").substring(0, 40) || "(未填備註)";
  const tags = String(r[5] || "");
  const category = String(r[6] || "未分類");
  const date = String(r[8] || r[7] || "").substring(0, 10);

  const bodyContents = [
    { type: "text", text: note, size: "sm", color: "#333333", wrap: true, maxLines: 3 },
    makeRow("日期", date),
  ];
  if (tags) bodyContents.splice(1, 0, { type: "text", text: "#" + tags, size: "xs", color: "#888888" });

  return {
    type: "bubble", size: "kilo",
    styles: { header: { backgroundColor: platform.color } },
    header: { type: "box", layout: "horizontal", paddingAll: "10px", contents: [
      { type: "text", text: platform.name, size: "xs", flex: 1, color: "#FFFFFF", gravity: "center" },
      { type: "text", text: category, size: "xs", flex: 0, color: "#FFFFFF99" },
    ]},
    body: { type: "box", layout: "vertical", paddingAll: "12px", spacing: "sm", contents: bodyContents },
    footer: { type: "box", layout: "vertical", paddingAll: "10px", contents: [{ type: "button", style: "primary", height: "sm", color: platform.color || "#1A1A2E", action: { type: "uri", label: "開啟連結", uri: url } }] },
  };
}

function buildLinkListRow(r, i) {
  const url = String(r[1] || "");
  const platform = detectPlatform(url);
  const note = String(r[4] || "").substring(0, 30) || "(未填備註)";
  const category = String(r[6] || "未分類");
  const date = String(r[8] || r[7] || "").substring(0, 10);
  return {
    type: "box", layout: "vertical", spacing: "xs", paddingTop: i === 0 ? "0px" : "8px",
    contents: [
      { type: "box", layout: "horizontal", contents: [
        { type: "text", text: "[" + platform.name + "] " + note, size: "sm", color: "#333333", flex: 1, wrap: true, maxLines: 2 },
        { type: "button", style: "link", height: "sm", flex: 0, action: { type: "uri", label: "開啟", uri: url } },
      ]},
      { type: "text", text: category + "  " + date, size: "xs", color: "#888888" },
      { type: "separator" },
    ],
  };
}

async function checkDuplicateThreads(url, userId, sheets) {
  const data = await getSheetData(sheets, SHEET.THREADS);
  return data.slice(1).some(function(row) { return String(row[1] || "") === url && String(row[7] || "") === userId; });
}

// ============================================================
//  File Backup - OAuth 2.0
// ============================================================

async function handleFileMessage(replyToken, message, type, userId, clients) {
  const sheets = clients.sheets;

  try {
    const drive = getOAuthDriveClient();
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
    const msg = String(err && err.message ? err.message : err || "");
    console.error("File upload error:", msg);
    await replyText(replyToken, "檔案備份失敗：" + msg);
  }
}

function downloadLineContent(messageId) {
  return new Promise(function(resolve, reject) {
    https.get({
      hostname: "api-data.line.me",
      path: "/v2/bot/message/" + messageId + "/content",
      method: "GET",
      headers: { Authorization: "Bearer " + LINE_TOKEN },
    }, function(res) {
      if (res.statusCode && res.statusCode >= 400) { reject(new Error("LINE content download failed: " + res.statusCode)); return; }
      const chunks = [];
      res.on("data", function(chunk) { chunks.push(chunk); });
      res.on("end", function() { resolve(Buffer.concat(chunks)); });
      res.on("error", reject);
    }).on("error", reject);
  });
}

async function uploadToDrive(drive, fileName, mimeType, buffer, fileType) {
  const folderId = await getOrCreateFolder(drive, fileType);
  const stream = Readable.from(buffer);
  const result = await drive.files.create({
    requestBody: { name: fileName, parents: [folderId] },
    media: { mimeType: mimeType, body: stream },
    fields: "id,webViewLink,name",
  });
  return result.data;
}

async function getOrCreateFolder(drive, folderName) {
  if (!DRIVE_FOLDER_ID) {
    const folder = await drive.files.create({
      requestBody: { name: "LINE Bot 備份", mimeType: "application/vnd.google-apps.folder" },
      fields: "id",
    });
    return folder.data.id;
  }

  const res = await drive.files.list({
    q: "'" + DRIVE_FOLDER_ID + "' in parents and name='" + folderName + "' and mimeType='application/vnd.google-apps.folder' and trashed=false",
    fields: "files(id)",
  });
  if (res.data.files.length > 0) return res.data.files[0].id;

  const folder = await drive.files.create({
    requestBody: { name: folderName, mimeType: "application/vnd.google-apps.folder", parents: [DRIVE_FOLDER_ID] },
    fields: "id",
  });
  return folder.data.id;
}

function buildFileSavedFlex(fileName, fileType, driveUrl, now) {
  return {
    type: "flex", altText: "檔案備份完成",
    contents: {
      type: "bubble",
      styles: { header: { backgroundColor: "#1a73e8" } },
      header: { type: "box", layout: "vertical", paddingAll: "16px", contents: [{ type: "text", text: "檔案備份完成", color: "#ffffff", weight: "bold", size: "md" }] },
      body: { type: "box", layout: "vertical", paddingAll: "16px", spacing: "md", contents: [makeRow("檔案名", fileName), makeRow("類型", fileType.toUpperCase()), makeRow("時間", now)] },
      footer: { type: "box", layout: "vertical", paddingAll: "12px", contents: [{ type: "button", style: "primary", height: "sm", color: "#1a73e8", action: { type: "uri", label: "查看 Google Drive", uri: driveUrl } }] },
    },
  };
}

async function handleRecentFiles(replyToken, userId, sheets) {
  const data = await getSheetData(sheets, SHEET.FILES);
  const results = data.slice(1).filter(function(r) { return String(r[9] || "") === userId; }).reverse().slice(0, 10);
  if (!results.length) { await replyText(replyToken, "還沒有備份檔案。"); return; }
  const lines = results.map(function(r, i) {
    return (i + 1) + ". " + String(r[1] || "") + "\n   " + String(r[2] || "") + "  " + String(r[11] || "").substring(0, 10) + "\n   " + String(r[6] || "");
  });
  await replyText(replyToken, "最近 " + results.length + " 筆檔案\n\n" + lines.join("\n\n"));
}

// ============================================================
//  Calendar
// ============================================================

async function handleCalendar(replyToken, text, userId, calendar, sheets) {
  const parsed = parseCalendarInput(text);
  if (!parsed) {
    await replyText(replyToken, "我找不到日期，無法建立行程。\n\n請用這樣的格式：\n明天 下午3點 開會\n4/25 10:00 客戶簡報");
    return;
  }
  try {
    const event = await createCalendarEvent(calendar, parsed);
    const id = generateId("C");
    const now = formatDateTime(new Date());
    const row = [id, parsed.title, formatDate(parsed.start), parsed.isAllDay ? "" : formatTime(parsed.start), parsed.isAllDay ? "TRUE" : "FALSE", event.id, text, "text", userId, now];
    await appendRow(sheets, SHEET.CALENDAR, row);
    await replyFlex(replyToken, buildCalendarFlex(parsed, event));
  } catch (err) {
    console.error("Calendar error:", err);
    await replyText(replyToken, "建立行程失敗：" + err.message);
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
    const weekMatch = text.match(/下[週周]?([一二三四五六日天])?/);
    if (weekMatch && weekMatch[1]) {
      const dayMap = { "一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "日": 0, "天": 0 };
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

  const isPM = /下午|晚上|pm/i.test(text);
  const isAM = /上午|早上|am/i.test(text);
  const timeMatch = text.match(/(\d{1,2})[:點時](\d{0,2})/);

  if (timeMatch) {
    let hour = parseInt(timeMatch[1], 10);
    const min = parseInt(timeMatch[2] || "0", 10);
    if (isPM && hour < 12) hour += 12;
    if (isAM && hour === 12) hour = 0;
    targetTime = { hour: hour, min: min };
    title = title.replace(timeMatch[0], "");
  }

  title = title.replace(/上午|早上|下午|晚上|pm|am/gi, "").replace(/\s+/g, " ").trim();
  if (!title) title = "行程待辦";

  const start = new Date(targetDate);
  const isAllDay = !targetTime;
  if (targetTime) start.setHours(targetTime.hour, targetTime.min, 0, 0);
  const end = new Date(start);
  if (isAllDay) end.setDate(end.getDate() + 1);
  else end.setHours(end.getHours() + 1);

  return { title: title, start: start, end: end, isAllDay: isAllDay };
}

async function createCalendarEvent(calendar, parsed) {
  const body = {
    summary: parsed.title,
    reminders: { useDefault: false, overrides: parsed.isAllDay ? [] : [{ method: "popup", minutes: 10 }] },
  };
  if (parsed.isAllDay) {
    body.start = { date: formatDate(parsed.start) };
    body.end = { date: formatDate(parsed.end) };
  } else {
    body.start = { dateTime: parsed.start.toISOString(), timeZone: "Asia/Taipei" };
    body.end = { dateTime: parsed.end.toISOString(), timeZone: "Asia/Taipei" };
  }
  const result = await calendar.events.insert({ calendarId: CALENDAR_ID, requestBody: body });
  return result.data;
}

function buildCalendarFlex(parsed, event) {
  const timeStr = parsed.isAllDay ? formatDate(parsed.start) + "(全天)" : formatDate(parsed.start) + " " + formatTime(parsed.start);
  return {
    type: "flex", altText: "已新增：" + parsed.title,
    contents: {
      type: "bubble",
      styles: { header: { backgroundColor: "#0F9D58" } },
      header: { type: "box", layout: "vertical", paddingAll: "16px", contents: [{ type: "text", text: "已新增到 Google 日曆", color: "#ffffff", weight: "bold", size: "md" }] },
      body: { type: "box", layout: "vertical", paddingAll: "16px", spacing: "md", contents: [makeRow("行程", parsed.title), makeRow("時間", timeStr), makeRow("提醒", parsed.isAllDay ? "(全天，無提醒)" : "開始前 10 分鐘")] },
      footer: event.htmlLink ? { type: "box", layout: "vertical", paddingAll: "12px", contents: [{ type: "button", style: "primary", height: "sm", color: "#0F9D58", action: { type: "uri", label: "查看 Google 日曆", uri: event.htmlLink } }] } : undefined,
    },
  };
}

// ============================================================
//  Main Menu
// ============================================================

function buildMainMenuFlex() {
  return {
    type: "flex",
    altText: "你好！我可以幫你做這些事",
    contents: {
      type: "bubble", size: "mega",
      styles: { header: { backgroundColor: "#1A1A2E" } },
      header: {
        type: "box", layout: "vertical", paddingAll: "20px",
        contents: [
          { type: "text", text: "你好！我可以幫你做這些事", color: "#FFFFFF", weight: "bold", size: "md" },
          { type: "text", text: "選擇功能或直接輸入", color: "#FFFFFF88", size: "xs" },
        ],
      },
      body: {
        type: "box", layout: "vertical", paddingAll: "16px", spacing: "lg",
        contents: [
          menuItem("[LINK]", "貼連結收藏", "任意網址自動收藏，可直接分類"),
          { type: "separator" },
          menuItem("[LIST]", "待辦清單", "查看、新增、完成、刪除待辦"),
          { type: "separator" },
          menuItem("[FILE]", "傳圖片/檔案", "自動備份到 Google Drive"),
          { type: "separator" },
          menuItem("[CAL]", "說待辦事項", "含時間則新增到 Google 日曆"),
        ],
      },
      footer: {
        type: "box", layout: "vertical", paddingAll: "12px", spacing: "sm",
        contents: [
          {
            type: "box", layout: "horizontal", spacing: "sm",
            contents: DEFAULT_CATEGORIES.filter(function(c) { return c !== "未分類"; }).map(function(cat) {
              return { type: "button", style: "secondary", height: "sm", flex: 1, action: { type: "postback", label: cat, data: "action=view_category&cat=" + cat, displayText: cat + "收藏" } };
            }),
          },
          {
            type: "box", layout: "horizontal", spacing: "sm",
            contents: [
              { type: "button", style: "secondary", height: "sm", flex: 1, action: { type: "message", label: "最近收藏", text: "最近收藏" } },
              { type: "button", style: "secondary", height: "sm", flex: 1, action: { type: "message", label: "查看待辦", text: "查看待辦" } },
              { type: "button", style: "primary", height: "sm", flex: 1, color: "#1A1A2E", action: { type: "message", label: "待辦清單", text: "待辦清單" } },
            ],
          },
        ],
      },
    },
  };
}

function menuItem(icon, title, desc) {
  return {
    type: "box", layout: "horizontal", spacing: "md",
    contents: [
      { type: "text", text: icon, size: "sm", flex: 0, gravity: "center", color: "#1A1A2E" },
      { type: "box", layout: "vertical", flex: 1, contents: [
        { type: "text", text: title, size: "sm", weight: "bold", color: "#1A1A2E" },
        { type: "text", text: desc, size: "xs", color: "#888888", wrap: true },
      ]},
    ],
  };
}

// ============================================================
//  Utilities
// ============================================================

async function getSheetData(sheets, sheetName) {
  const result = await sheets.spreadsheets.values.get({ spreadsheetId: SHEETS_ID, range: sheetName + "!A:Z" });
  return result.data.values || [];
}

async function appendRow(sheets, sheetName, row) {
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEETS_ID,
    range: sheetName + "!A:Z",
    valueInputOption: "RAW",
    requestBody: { values: [row] },
  });
}

async function findLastUserRow(sheets, sheetName, userId) {
  const data = await getSheetData(sheets, sheetName);
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    if (String(row[7] || row[9] || "") === userId) return { row: row, index: i };
  }
  return { row: null, index: -1 };
}

function generateId(prefix) {
  const ts = new Date().toISOString().replace(/[:\-T.Z]/g, "").substring(0, 14);
  const rnd = Math.floor(Math.random() * 1000).toString().padStart(3, "0");
  return prefix + ts + rnd;
}

function formatDateTime(date) {
  return new Date(date).toLocaleString("zh-TW", { timeZone: "Asia/Taipei", year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" });
}

function formatDate(date) {
  const d = new Date(date);
  return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
}

function formatTime(date) {
  const d = new Date(date);
  return String(d.getHours()).padStart(2, "0") + ":" + String(d.getMinutes()).padStart(2, "0");
}

function generateFileName(type, messageId) {
  const ext = { image: "jpg", video: "mp4", audio: "m4a", file: "bin" };
  const now = new Date().toISOString().substring(0, 10).replace(/-/g, "");
  return type + "_" + now + "_" + messageId + "." + (ext[type] || "bin");
}

function getMimeType(type, fileName) {
  const ext = path.extname(fileName || "").toLowerCase();
  const map = {
    ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png", ".gif": "image/gif", ".webp": "image/webp",
    ".pdf": "application/pdf", ".doc": "application/msword",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".xls": "application/vnd.ms-excel", ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".mp4": "video/mp4", ".mov": "video/quicktime", ".m4a": "audio/m4a", ".mp3": "audio/mpeg",
    ".wav": "audio/wav", ".txt": "text/plain", ".zip": "application/zip",
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

function makeRow(label, value) {
  return {
    type: "box", layout: "horizontal",
    contents: [
      { type: "text", text: label, size: "sm", color: "#888888", flex: 1 },
      { type: "text", text: String(value || ""), size: "sm", color: "#333333", flex: 3, wrap: true },
    ],
  };
}

function replyText(replyToken, text) { return replyMessages(replyToken, [{ type: "text", text: text }]); }
function replyFlex(replyToken, flexMessage) { return replyMessages(replyToken, [flexMessage]); }
function replyMessages(replyToken, messages) { return lineApiRequest("/v2/bot/message/reply", { replyToken: replyToken, messages: messages }); }

function lineApiRequest(apiPath, payload) {
  return new Promise(function(resolve, reject) {
    const body = JSON.stringify(payload);
    const req = https.request({
      hostname: "api.line.me",
      path: apiPath,
      method: "POST",
      headers: { Authorization: "Bearer " + LINE_TOKEN, "Content-Type": "application/json", "Content-Length": Buffer.byteLength(body) },
    }, function(res) {
      let data = "";
      res.on("data", function(chunk) { data += chunk; });
      res.on("end", function() {
        if (res.statusCode >= 200 && res.statusCode < 300) resolve(data ? JSON.parse(data) : {});
        else reject(new Error("LINE API error " + res.statusCode + ": " + data));
      });
    });
    req.on("error", reject);
    req.write(body);
    req.end();
  });
}
