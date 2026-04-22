// api/webhook.js
// LINE Bot — 個人管理工具 v5.1.0
// 可直接覆蓋版
//
// 主要修正：
// 1) 修正檔案備份流程，不再重複使用同一個 replyToken 兩次
// 2) 補上「待辦清單」主入口，直接顯示操作方式
// 3) 收藏成功後，明確顯示分類按鈕（個人 / 工作 / 家庭）
// 4) 查詢最近收藏 / 分類收藏 / 最近檔案時改為以 userId 篩選
// 5) 支援 audio 備份到 Google Drive
//
// ============================================================
//  環境變數（Vercel）
//  LINE_CHANNEL_ACCESS_TOKEN   → LINE Developers 取得
//  GOOGLE_SERVICE_ACCOUNT_JSON → GCP 服務帳號 JSON（整份內容）
//  GOOGLE_SHEETS_ID            → Google Sheets 試算表 ID
//  GOOGLE_DRIVE_FOLDER_ID      → Drive 備份根資料夾 ID
//  GOOGLE_CALENDAR_ID          → 日曆 ID（或填 primary）
// ============================================================
//
// 建議工作表：
// 1. Threads收藏
//    A:id | B:url | C:platform | D:title | E:note | F:tags | G:category | H:userId | I:createdAt | J:updatedAt | K:rawText
// 2. 檔案備份
//    A:id | B:fileName | C:fileType | D:mimeType | E:size | F:driveFileId | G:driveUrl | H:tag1 | I:tag2 | J:userId | K:lineMessageId | L:createdAt
// 3. 日曆待辦
//    A:id | B:title | C:date | D:time | E:isAllDay | F:googleEventId | G:rawText | H:source | I:userId | J:createdAt
// 4. 待辦清單
//    清單列：[A:listId | B:listName | C:userId | D:createdAt | E:active]
//    項目列：[A:listId | B:listName | C:itemId | D:itemText | E:pending/done | F:userId | G:createdAt]

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
    return res.status(200).json({ status: "ok", version: "5.1.0" });
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

    await replyText(replyToken, "目前支援文字、圖片、影片、音訊與一般檔案。\n輸入「說明」可查看功能。\n");
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

  if (action === "todo_done") {
    await handleTodoDone(replyToken, params.get("listId"), params.get("itemId"), userId, sheets);
    return;
  }

  if (action === "set_category") {
    await handleSetCategory(replyToken, params.get("id"), params.get("cat"), userId, sheets);
    return;
  }

  if (action === "view_category") {
    await handleQueryCategory(replyToken, params.get("cat"), userId, sheets);
    return;
  }

  await replyText(replyToken, "⚠️ 未知的操作。\n");
}

async function handleTextMessage(replyToken, text, userId, clients) {
  const { sheets, calendar } = clients;
  const lower = text.toLowerCase();

  // Rich Menu / 主入口
  if (text === "功能說明" || /^(說明|help|使用說明|功能|如何使用)$/i.test(text)) {
    await replyFlex(replyToken, buildMainMenuFlex());
    return;
  }

  if (text === "待辦清單") {
    await replyFlex(replyToken, buildTodoEntryFlex());
    return;
  }

  if (/^(新增待辦|新增事項|新增項目)$/.test(text)) {
    await replyText(
      replyToken,
      "請用這個格式新增待辦：\n\n1. 先建立清單\n   例如：清單 購物清單\n\n2. 再加入項目\n   例如：新增 買牛奶 到 購物清單\n\n你也可以直接點選下方主選單中的待辦清單。"
    );
    return;
  }

  // 待辦清單：建立清單
  if (/^(清單|建立清單|新增清單)\s+\S+/.test(text)) {
    const listName = text.replace(/^(清單|建立清單|新增清單)\s+/, "").trim();
    await handleCreateTodoList(replyToken, listName, userId, sheets);
    return;
  }

  // 待辦清單：查清單列表
  if (/^(我的清單|查清單|清單列表|所有清單)$/.test(text)) {
    await handleListTodos(replyToken, userId, sheets);
    return;
  }

  // 待辦清單：看清單
  if (/^(看清單|打開清單|查看清單)\s+\S+/.test(text)) {
    const listName = text.replace(/^(看清單|打開清單|查看清單)\s+/, "").trim();
    await handleViewTodoList(replyToken, listName, userId, sheets);
    return;
  }

  // 待辦清單：新增項目
  if (/^(新增|加入|加)\s+.+\s+(到|進)\s+\S+$/.test(text)) {
    const match = text.match(/^(新增|加入|加)\s+(.+)\s+(到|進)\s+(\S+)$/);
    if (match) {
      const itemText = match[2].trim();
      const listName = match[4].trim();
      await handleAddTodoItem(replyToken, listName, itemText, userId, sheets);
      return;
    }
  }

  // 待辦清單：完成項目
  if (/^(完成|done|勾選)\s+\S+/.test(text)) {
    const itemText = text.replace(/^(完成|done|勾選)\s+/, "").trim();
    await handleCompleteTodoByText(replyToken, itemText, userId, sheets);
    return;
  }

  // 分類查詢
  for (const cat of CATEGORIES) {
    if (text === `${cat}收藏` || text === cat) {
      await handleQueryCategory(replyToken, cat, userId, sheets);
      return;
    }
  }

  // 分類指定（將最近一筆改分類）
  if (/^(分類|設為|歸類|標記為)\s*(個人|工作|家庭|未分類)/.test(text)) {
    const catMatch = text.match(/(個人|工作|家庭|未分類)/);
    if (catMatch) {
      await handleSetCategory(replyToken, "LAST", catMatch[1], userId, sheets);
      return;
    }
  }

  // 連結收藏
  const urls = extractSupportedUrls(text);
  if (urls.length > 0) {
    await handleLinkSave(replyToken, text, urls, userId, sheets);
    return;
  }

  // 標籤
  if (/^(標籤|加標籤|幫我加標籤)\s+\S+/.test(text)) {
    await handleAddTag(replyToken, text, userId, sheets);
    return;
  }
  if (/^(改標籤|修改標籤|更改標籤)\s+\S+/.test(text)) {
    const tagStr = text.replace(/^(改標籤|修改標籤|更改標籤)\s+/, "").trim();
    await handleAddTag(replyToken, `標籤 ${tagStr}`, userId, sheets);
    return;
  }

  // 查標籤
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

  // 搜尋收藏
  if (/^(搜尋|搜索|找)\s+\S+/.test(text) && !text.includes("檔案")) {
    const keyword = text.replace(/^(搜尋|搜索|找)\s+/, "").trim();
    await handleSearchLinks(replyToken, keyword, userId, sheets);
    return;
  }

  // 最近收藏
  if (/最近.*收藏|最近.*threads|最近.*thr/i.test(lower) || text === "最近收藏") {
    await handleRecentLinks(replyToken, userId, sheets);
    return;
  }

  // 刪除最近一筆收藏
  if (/^(刪除|delete)\s*(最新|最近一筆|last)?$/.test(text.trim())) {
    await handleDeleteLast(replyToken, userId, sheets);
    return;
  }

  // 最近檔案
  if (/最近.*檔案|檔案.*最近/.test(text) || text === "最近檔案") {
    await handleRecentFiles(replyToken, userId, sheets);
    return;
  }

  // 搜尋檔案
  if (/^(搜尋檔案|找檔案|搜檔案)\s+\S+/.test(text)) {
    const keyword = text.replace(/^(搜尋檔案|找檔案|搜檔案)\s+/, "").trim();
    await handleSearchFiles(replyToken, keyword, userId, sheets);
    return;
  }

  // 日曆
  if (/明天|行程|提醒|會議|今天|下午|上午|早上|明後天|後天|下週|下周|下星期|週一|週二|週三|週四|週五|週六|週日/.test(text)) {
    await handleCalendar(replyToken, text, userId, calendar, sheets);
    return;
  }

  await replyFlex(replyToken, buildMainMenuFlex());
}

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
    altText: `${platform.emoji} ${platform.name} 已收藏`,
    contents: {
      type: "bubble",
      styles: { header: { backgroundColor: platform.color } },
      header: {
        type: "box",
        layout: "horizontal",
        paddingAll: "16px",
        contents: [
          { type: "text", text: platform.emoji, size: "xl", flex: 0, gravity: "center" },
          {
            type: "box",
            layout: "vertical",
            flex: 1,
            paddingStart: "10px",
            contents: [
              { type: "text", text: `${platform.name} 已收藏`, color: "#FFFFFF", weight: "bold", size: "md" },
              { type: "text", text: "請直接點下方分類按鈕分類收藏", color: "#FFFFFF88", size: "xs" },
            ],
          },
        ],
      },
      body: {
        type: "box",
        layout: "vertical",
        paddingAll: "16px",
        spacing: "md",
        contents: [
          note ? makeRow("備註", note) : null,
          makeRow("分類", "未分類"),
          makeRow("時間", now),
        ].filter(Boolean),
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
            contents: ["個人", "工作", "家庭"].map((cat) => ({
              type: "button",
              style: "secondary",
              height: "sm",
              flex: 1,
              action: {
                type: "postback",
                label: `${CATEGORY_EMOJI[cat]} ${cat}`,
                data: `action=set_category&id=LAST&cat=${cat}`,
                displayText: `設為${cat}`,
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
                style: "primary",
                height: "sm",
                flex: 2,
                color: platform.color || "#1A1A2E",
                action: { type: "uri", label: `開啟 ${platform.name}`, uri: url },
              },
              {
                type: "button",
                style: "secondary",
                height: "sm",
                flex: 1,
                action: { type: "message", label: "加標籤", text: "加標籤 " },
              },
            ],
          },
        ],
      },
    },
  };
}

async function handleSetCategory(replyToken, recordId, category, userId, sheets) {
  const { row: lastRow, index: lastIndex } = await findLastUserRow(sheets, SHEET.THREADS, userId);
  if (lastIndex === -1) {
    await replyText(replyToken, "找不到最近的收藏。\n");
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

  if (results.length === 0) {
    await replyFlex(
      replyToken,
      buildEmptyStateFlex(`${CATEGORY_EMOJI[category]} ${category}還沒有收藏`, "收藏連結後，點成功卡片下方的分類按鈕即可歸類", [
        { label: "查看最近收藏", text: "最近收藏" },
      ])
    );
    return;
  }

  await replyFlex(replyToken, buildLinksCarousel(results, `${CATEGORY_EMOJI[category]} ${category} — ${all.length} 筆`));
}

async function handleCreateTodoList(replyToken, listName, userId, sheets) {
  const listId = generateId("L");
  const now = formatDateTime(new Date());
  const row = [listId, listName, userId, now, "active"];
  await appendRow(sheets, SHEET.TODOS, row);
  await replyFlex(replyToken, buildTodoListCreatedFlex(listName));
}

function buildTodoListCreatedFlex(listName) {
  return {
    type: "flex",
    altText: `✅ 清單「${listName}」已建立`,
    contents: {
      type: "bubble",
      styles: { header: { backgroundColor: "#2D3561" } },
      header: {
        type: "box",
        layout: "vertical",
        paddingAll: "16px",
        contents: [
          { type: "text", text: "✅ 清單已建立", color: "#FFFFFF", weight: "bold", size: "md" },
          { type: "text", text: listName, color: "#FFFFFFCC", size: "sm" },
        ],
      },
      body: {
        type: "box",
        layout: "vertical",
        paddingAll: "16px",
        spacing: "md",
        contents: [
          { type: "text", text: "💡 下一步怎麼新增項目？", weight: "bold", size: "sm", color: "#333333" },
          { type: "text", text: `請傳送：新增 買牛奶 到 ${listName}`, size: "sm", color: "#555555", wrap: true },
        ],
      },
      footer: {
        type: "box",
        layout: "vertical",
        paddingAll: "12px",
        contents: [
          { type: "button", style: "primary", height: "sm", color: "#2D3561", action: { type: "message", label: "查看我的清單", text: "我的清單" } },
        ],
      },
    },
  };
}

function buildTodoEntryFlex() {
  return {
    type: "flex",
    altText: "待辦清單操作",
    contents: {
      type: "bubble",
      styles: { header: { backgroundColor: "#2D3561" } },
      header: {
        type: "box",
        layout: "vertical",
        paddingAll: "16px",
        contents: [
          { type: "text", text: "📋 待辦清單", color: "#FFFFFF", weight: "bold", size: "md" },
          { type: "text", text: "先建立清單，再加入項目", color: "#FFFFFFCC", size: "xs" },
        ],
      },
      body: {
        type: "box",
        layout: "vertical",
        paddingAll: "16px",
        spacing: "md",
        contents: [
          { type: "text", text: "建立清單", weight: "bold", size: "sm", color: "#333333" },
          { type: "text", text: "例如：清單 購物清單", size: "sm", color: "#666666", wrap: true },
          { type: "separator" },
          { type: "text", text: "新增項目", weight: "bold", size: "sm", color: "#333333" },
          { type: "text", text: "例如：新增 買牛奶 到 購物清單", size: "sm", color: "#666666", wrap: true },
          { type: "separator" },
          { type: "text", text: "查看清單", weight: "bold", size: "sm", color: "#333333" },
          { type: "text", text: "例如：我的清單 或 看清單 購物清單", size: "sm", color: "#666666", wrap: true },
        ],
      },
      footer: {
        type: "box",
        layout: "horizontal",
        spacing: "sm",
        paddingAll: "12px",
        contents: [
          { type: "button", style: "secondary", height: "sm", flex: 1, action: { type: "message", label: "建立清單", text: "清單 " } },
          { type: "button", style: "primary", height: "sm", flex: 1, color: "#2D3561", action: { type: "message", label: "我的清單", text: "我的清單" } },
        ],
      },
    },
  };
}

async function handleAddTodoItem(replyToken, listName, itemText, userId, sheets) {
  const todoData = await getSheetData(sheets, SHEET.TODOS);
  const listRow = todoData
    .slice(1)
    .find((r) => String(r[1] || "") === listName && String(r[2] || "") === userId && String(r[4] || "") === "active");

  if (!listRow) {
    await replyFlex(replyToken, buildEmptyStateFlex(`找不到清單「${listName}」`, `請先建立清單：「清單 ${listName}」`, [{ label: "查看我的清單", text: "我的清單" }]));
    return;
  }

  const listId = String(listRow[0]);
  const itemId = generateId("I");
  const now = formatDateTime(new Date());
  const row = [listId, listName, itemId, itemText, "pending", userId, now];
  await appendRow(sheets, SHEET.TODOS, row);

  const progress = await getTodoProgress(sheets, listId);
  await replyFlex(replyToken, buildTodoItemAddedFlex(listName, itemText, progress));
}

function buildTodoItemAddedFlex(listName, itemText, progress) {
  return {
    type: "flex",
    altText: `➕ 已加入：${itemText}`,
    contents: {
      type: "bubble",
      body: {
        type: "box",
        layout: "vertical",
        paddingAll: "16px",
        spacing: "md",
        contents: [
          {
            type: "box",
            layout: "horizontal",
            spacing: "md",
            contents: [
              { type: "text", text: "➕", size: "xl", flex: 0, gravity: "center" },
              {
                type: "box",
                layout: "vertical",
                flex: 1,
                contents: [
                  { type: "text", text: itemText, weight: "bold", size: "md", color: "#333333", wrap: true },
                  { type: "text", text: `加入「${listName}」`, size: "xs", color: "#888888" },
                ],
              },
            ],
          },
          { type: "separator" },
          buildProgressBar(progress.done, progress.total),
        ],
      },
      footer: {
        type: "box",
        layout: "horizontal",
        spacing: "sm",
        paddingAll: "12px",
        contents: [
          { type: "button", style: "primary", height: "sm", flex: 1, color: "#2D3561", action: { type: "message", label: "查看清單", text: `看清單 ${listName}` } },
        ],
      },
    },
  };
}

function buildProgressBar(done, total) {
  const pct = total > 0 ? Math.round((done / total) * 100) : 0;
  const filled = Math.round(pct / 10);
  const bar = "█".repeat(filled) + "░".repeat(10 - filled);
  return {
    type: "box",
    layout: "vertical",
    spacing: "xs",
    contents: [
      {
        type: "box",
        layout: "horizontal",
        contents: [
          { type: "text", text: "進度", size: "xs", color: "#888888", flex: 1 },
          { type: "text", text: `${done}/${total} (${pct}%)`, size: "xs", color: "#555555", flex: 0 },
        ],
      },
      { type: "text", text: bar, size: "sm", color: pct === 100 ? "#0F9D58" : "#2D3561", letterSpacing: "-2px" },
    ],
  };
}

async function handleListTodos(replyToken, userId, sheets) {
  const todoData = await getSheetData(sheets, SHEET.TODOS);
  const lists = todoData
    .slice(1)
    .filter((r) => String(r[2] || "") === userId && String(r[4] || "") === "active");

  if (lists.length === 0) {
    await replyFlex(replyToken, buildEmptyStateFlex("還沒有待辦清單", "傳送「清單 清單名稱」來建立第一個清單", [{ label: "建立購物清單", text: "清單 購物清單" }]));
    return;
  }

  const listsWithProgress = await Promise.all(
    lists.map(async (r) => {
      const listId = String(r[0]);
      const progress = await getTodoProgress(sheets, listId);
      return { listId, name: String(r[1]), progress };
    })
  );

  await replyFlex(replyToken, buildTodoListsFlex(listsWithProgress));
}

function buildTodoListsFlex(lists) {
  const items = lists.slice(0, 8).map(({ name, progress }) => {
    const pct = progress.total > 0 ? Math.round((progress.done / progress.total) * 100) : 0;
    return {
      type: "box",
      layout: "horizontal",
      paddingTop: "8px",
      contents: [
        {
          type: "box",
          layout: "vertical",
          flex: 1,
          contents: [
            { type: "text", text: name, size: "sm", weight: "bold", color: "#333333" },
            { type: "text", text: `${progress.done}/${progress.total} 完成`, size: "xs", color: "#888888" },
          ],
        },
        {
          type: "box",
          layout: "vertical",
          flex: 0,
          gravity: "center",
          contents: [{ type: "text", text: `${pct}%`, size: "sm", color: pct === 100 ? "#0F9D58" : "#2D3561", weight: "bold" }],
        },
        { type: "button", style: "link", height: "sm", flex: 0, action: { type: "message", label: "查看", text: `看清單 ${name}` } },
      ],
    };
  });

  return {
    type: "flex",
    altText: "我的待辦清單",
    contents: {
      type: "bubble",
      styles: { header: { backgroundColor: "#2D3561" } },
      header: {
        type: "box",
        layout: "vertical",
        paddingAll: "16px",
        contents: [
          { type: "text", text: "📋 我的待辦清單", color: "#FFFFFF", weight: "bold", size: "md" },
          { type: "text", text: `共 ${lists.length} 個清單`, color: "#FFFFFFCC", size: "xs" },
        ],
      },
      body: {
        type: "box",
        layout: "vertical",
        paddingAll: "12px",
        spacing: "none",
        contents: items.length ? items.flatMap((item, i) => (i < items.length - 1 ? [item, { type: "separator" }] : [item])) : [{ type: "text", text: "還沒有清單", size: "sm", color: "#888888", align: "center" }],
      },
      footer: {
        type: "box",
        layout: "vertical",
        paddingAll: "12px",
        contents: [{ type: "button", style: "secondary", height: "sm", action: { type: "message", label: "➕ 建立新清單", text: "清單 " } }],
      },
    },
  };
}

async function handleViewTodoList(replyToken, listName, userId, sheets) {
  const todoData = await getSheetData(sheets, SHEET.TODOS);
  const listRow = todoData
    .slice(1)
    .find((r) => String(r[1] || "") === listName && String(r[2] || "") === userId && String(r[4] || "") === "active");

  if (!listRow) {
    await replyFlex(replyToken, buildEmptyStateFlex(`找不到清單「${listName}」`, "請確認清單名稱，或建立新清單", [{ label: "查看所有清單", text: "我的清單" }]));
    return;
  }

  const listId = String(listRow[0]);
  const items = todoData
    .slice(1)
    .filter((r) => String(r[0] || "") === listId && String(r[2] || "").startsWith("I"));

  await replyFlex(replyToken, buildTodoDetailFlex(listId, listName, items));
}

function buildTodoDetailFlex(listId, listName, items) {
  const done = items.filter((r) => String(r[4]) === "done").length;
  const total = items.length;

  const itemRows = items.slice(0, 10).map((r) => {
    const isDone = String(r[4]) === "done";
    const itemText = String(r[3] || "");
    const itemId = String(r[2] || "");
    return {
      type: "box",
      layout: "horizontal",
      spacing: "md",
      paddingTop: "8px",
      contents: [
        { type: "text", text: isDone ? "✅" : "⬜", size: "sm", flex: 0, gravity: "center" },
        { type: "text", text: itemText, size: "sm", flex: 1, gravity: "center", color: isDone ? "#AAAAAA" : "#333333", decoration: isDone ? "line-through" : "none", wrap: true },
        isDone
          ? null
          : {
              type: "button",
              style: "link",
              height: "sm",
              flex: 0,
              action: {
                type: "postback",
                label: "完成",
                data: `action=todo_done&listId=${listId}&itemId=${itemId}`,
                displayText: `完成：${itemText}`,
              },
            },
      ].filter(Boolean),
    };
  });

  return {
    type: "flex",
    altText: `📋 ${listName}`,
    contents: {
      type: "bubble",
      styles: { header: { backgroundColor: "#2D3561" } },
      header: {
        type: "box",
        layout: "vertical",
        paddingAll: "16px",
        contents: [
          { type: "text", text: `📋 ${listName}`, color: "#FFFFFF", weight: "bold", size: "md" },
          { type: "text", text: `${done}/${total} 完成`, color: "#FFFFFFBB", size: "xs" },
        ],
      },
      body: {
        type: "box",
        layout: "vertical",
        paddingAll: "12px",
        spacing: "none",
        contents: itemRows.length
          ? itemRows.flatMap((item, i) => (i < itemRows.length - 1 ? [item, { type: "separator" }] : [item]))
          : [{ type: "text", text: "清單是空的，快新增第一件事！", size: "sm", color: "#888888", align: "center" }],
      },
      footer: {
        type: "box",
        layout: "vertical",
        paddingAll: "12px",
        contents: [{ type: "button", style: "secondary", height: "sm", action: { type: "message", label: "➕ 新增項目", text: `新增  到 ${listName}` } }],
      },
    },
  };
}

async function handleTodoDone(replyToken, listId, itemId, userId, sheets) {
  const todoData = await getSheetData(sheets, SHEET.TODOS);
  const rowIndex = todoData.findIndex((r) => String(r[0]) === listId && String(r[2]) === itemId && String(r[5] || "") === userId);

  if (rowIndex === -1) {
    await replyText(replyToken, "找不到這個項目。\n");
    return;
  }

  const rowNum = rowIndex + 1;
  const itemText = String(todoData[rowIndex][3] || "");
  const listName = String(todoData[rowIndex][1] || "");

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEETS_ID,
    range: `${SHEET.TODOS}!E${rowNum}`,
    valueInputOption: "RAW",
    requestBody: { values: [["done"]] },
  });

  const progress = await getTodoProgress(sheets, listId);
  await replyFlex(replyToken, {
    type: "flex",
    altText: `✅ 完成：${itemText}`,
    contents: {
      type: "bubble",
      body: {
        type: "box",
        layout: "vertical",
        paddingAll: "20px",
        spacing: "md",
        contents: [
          { type: "text", text: "✅ 已完成！", weight: "bold", size: "lg", color: "#0F9D58" },
          { type: "text", text: itemText, size: "md", color: "#333333", wrap: true },
          { type: "separator" },
          buildProgressBar(progress.done, progress.total),
        ],
      },
      footer: {
        type: "box",
        layout: "vertical",
        paddingAll: "12px",
        contents: [{ type: "button", style: "primary", height: "sm", color: "#2D3561", action: { type: "message", label: `查看${listName}`, text: `看清單 ${listName}` } }],
      },
    },
  });
}

async function handleCompleteTodoByText(replyToken, itemText, userId, sheets) {
  const todoData = await getSheetData(sheets, SHEET.TODOS);
  const rowIndex = todoData.findIndex(
    (r) => String(r[3] || "").includes(itemText) && String(r[5] || "") === userId && String(r[4]) === "pending"
  );

  if (rowIndex === -1) {
    await replyText(replyToken, `找不到待辦項目「${itemText}」`);
    return;
  }

  const listId = String(todoData[rowIndex][0]);
  const itemId = String(todoData[rowIndex][2]);
  await handleTodoDone(replyToken, listId, itemId, userId, sheets);
}

async function getTodoProgress(sheets, listId) {
  const todoData = await getSheetData(sheets, SHEET.TODOS);
  const items = todoData.slice(1).filter((r) => String(r[0] || "") === listId && String(r[2] || "").startsWith("I"));
  const done = items.filter((r) => String(r[4]) === "done").length;
  return { done, total: items.length };
}

async function handleAddTag(replyToken, text, userId, sheets) {
  const tagStr = text.replace(/^(標籤|加標籤|幫我加標籤)\s+/, "").trim();
  const tags = tagStr.split(/[\s,，、]+/).filter((t) => t.length > 0);

  const { row: lastRow, index: lastIndex } = await findLastUserRow(sheets, SHEET.THREADS, userId);
  if (lastIndex === -1) {
    await replyText(replyToken, "找不到最近的收藏，請先貼上連結再加標籤。\n");
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
    await replyFlex(replyToken, buildEmptyStateFlex("找不到符合的內容", `目前沒有標籤「${tag}」的收藏資料`, [{ label: "查看最近收藏", text: "最近收藏" }]));
    return;
  }

  await replyFlex(replyToken, buildLinksCarousel(results, `🏷 標籤「${tag}」— ${all.length} 筆`));
}

async function handleRecentLinks(replyToken, userId, sheets) {
  const data = await getSheetData(sheets, SHEET.THREADS);
  const results = data.slice(1).filter((r) => String(r[7] || "") === userId).reverse().slice(0, 10);

  if (!results.length) {
    await replyFlex(replyToken, buildEmptyStateFlex("還沒有收藏", "貼上連結就會自動收藏", [{ label: "使用說明", text: "說明" }]));
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
    await replyFlex(replyToken, buildEmptyStateFlex("找不到符合的內容", `沒有找到包含「${keyword}」的收藏`, [{ label: "查看最近收藏", text: "最近收藏" }]));
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
    await replyText(replyToken, "找不到你最近的收藏。\n");
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

async function handleFileMessage(replyToken, message, type, userId, clients) {
  const { sheets, drive } = clients;

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
}

function downloadLineContent(messageId) {
  return new Promise((resolve, reject) => {
    https
      .get(
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
      )
      .on("error", reject);
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
        contents: [makeRow("檔案名", fileName), makeRow("類型", `${emoji[fileType] || "📁"} ${fileType.toUpperCase()}`), makeRow("時間", now)],
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
    await replyFlex(replyToken, buildEmptyStateFlex("還沒有備份檔案", "傳送圖片、PDF、影片、語音或其他檔案就會自動備份", [{ label: "使用說明", text: "說明" }]));
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
    await replyFlex(replyToken, buildEmptyStateFlex("找不到符合的內容", `沒有找到包含「${keyword}」的檔案`, [{ label: "查看最近檔案", text: "最近檔案" }]));
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
    const row = [id, parsed.title, formatDate(parsed.start), parsed.isAllDay ? "" : formatTime(parsed.start), parsed.isAllDay ? "TRUE" : "FALSE", event.id, text, "text", userId, now];
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
    reminders: { useDefault: false, overrides: parsed.isAllDay ? [] : [{ method: "popup", minutes: 10 }] },
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
  const timeStr = parsed.isAllDay ? `${formatDate(parsed.start)}（全天）` : `${formatDate(parsed.start)} ${formatTime(parsed.start)}`;
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
        contents: [makeRow("行程", parsed.title), makeRow("時間", timeStr), makeRow("提醒", parsed.isAllDay ? "（全天，無提醒）" : "開始前 10 分鐘")],
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
          menuItem("📋", "待辦清單", "建立清單、新增項目、追蹤進度"),
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
              action: { type: "postback", label: `${CATEGORY_EMOJI[cat]}${cat}`, data: `action=view_category&cat=${cat}`, displayText: `${cat}收藏` },
            })),
          },
          {
            type: "box",
            layout: "horizontal",
            spacing: "sm",
            contents: [
              { type: "button", style: "secondary", height: "sm", flex: 1, action: { type: "message", label: "最近收藏", text: "最近收藏" } },
              { type: "button", style: "secondary", height: "sm", flex: 1, action: { type: "message", label: "我的清單", text: "我的清單" } },
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
              contents: buttons.map((b) => ({ type: "button", style: "secondary", height: "sm", action: { type: "message", label: b.label, text: b.text } })),
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
  const rnd = Math.floor(Math.random() * 1000)
    .toString()
    .padStart(3, "0");
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
