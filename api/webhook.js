// api/webhook.js
// LINE Bot - v7.2.0
// bot state + any URL bookmark + todo direct input + OAuth Drive

const { google } = require(“googleapis”);
const https = require(“https”);
const path = require(“path”);
const { Readable } = require(“stream”);

const LINE_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN;
const GOOGLE_CREDENTIALS = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
const SHEETS_ID = process.env.GOOGLE_SHEETS_ID;
const DRIVE_FOLDER_ID = process.env.GOOGLE_DRIVE_FOLDER_ID;
const CALENDAR_ID = process.env.GOOGLE_CALENDAR_ID || “primary”;

const GOOGLE_CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const GOOGLE_CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const GOOGLE_REFRESH_TOKEN = process.env.GOOGLE_REFRESH_TOKEN;
const REDIRECT_URI = “https://2060422.vercel.app/api/auth/callback”;

const SHEET = {
THREADS: “Threads\u6536\u85CF”,
FILES: “\u6A94\u6848\u5099\u4EFD”,
CALENDAR: “\u65E5\u66C6\u5F85\u8FA6”,
TODOS: “\u5F85\u8FA6\u6E05\u55AE”,
BOT_STATE: “\u6A5F\u5668\u4EBA\u72C0\u614B”,
};

const CATEGORIES = [”\u672A\u5206\u985E”, “\u500B\u4EBA”, “\u5DE5\u4F5C”, “\u5BB6\u5EAD”];
const CATEGORY_COLORS = {
“\u672A\u5206\u985E”: “#888888”,
“\u500B\u4EBA”: “#FF6B6B”,
“\u5DE5\u4F5C”: “#4ECDC4”,
“\u5BB6\u5EAD”: “#45B7D1”,
};

// ============================================================
//  Google Clients
// ============================================================

async function getGoogleClients() {
const credentials = JSON.parse(GOOGLE_CREDENTIALS);
const auth = new google.auth.GoogleAuth({
credentials,
scopes: [
“https://www.googleapis.com/auth/spreadsheets”,
“https://www.googleapis.com/auth/calendar”,
],
});
const authClient = await auth.getClient();
return {
sheets: google.sheets({ version: “v4”, auth: authClient }),
calendar: google.calendar({ version: “v3”, auth: authClient }),
};
}

function getOAuthDriveClient() {
const oauth2Client = new google.auth.OAuth2(GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, REDIRECT_URI);
oauth2Client.setCredentials({ refresh_token: GOOGLE_REFRESH_TOKEN });
return google.drive({ version: “v3”, auth: oauth2Client });
}

// ============================================================
//  Bot State
// ============================================================

async function getBotState(sheets, userId) {
const data = await getSheetData(sheets, SHEET.BOT_STATE);
const row = data.slice(1).find(function(r) { return String(r[0] || “”) === userId; });
return row ? String(row[1] || “”) : “”;
}

async function setBotState(sheets, userId, state) {
const data = await getSheetData(sheets, SHEET.BOT_STATE);
const now = formatDateTime(new Date());
const rowIndex = data.slice(1).findIndex(function(r) { return String(r[0] || “”) === userId; });

if (rowIndex === -1) {
await appendRow(sheets, SHEET.BOT_STATE, [userId, state, now]);
} else {
const rowNum = rowIndex + 2;
await sheets.spreadsheets.values.batchUpdate({
spreadsheetId: SHEETS_ID,
requestBody: {
valueInputOption: “RAW”,
data: [
{ range: SHEET.BOT_STATE + “!B” + rowNum, values: [[state]] },
{ range: SHEET.BOT_STATE + “!C” + rowNum, values: [[now]] },
],
},
});
}
}

async function clearBotState(sheets, userId) {
await setBotState(sheets, userId, “”);
}

// ============================================================
//  Main Entry
// ============================================================

module.exports = async function(req, res) {
if (req.method === “GET”) {
return res.status(200).json({ status: “ok”, version: “7.2.0” });
}
if (req.method !== “POST”) {
return res.status(405).json({ error: “Method not allowed” });
}

try {
const body = typeof req.body === “string” ? JSON.parse(req.body) : req.body;
if (!body || !body.events || !body.events.length) {
return res.status(200).json({ status: “ok” });
}
const clients = await getGoogleClients();
for (let i = 0; i < body.events.length; i++) {
await handleEvent(body.events[i], clients);
}
} catch (err) {
console.error(“Global error:”, err);
}

return res.status(200).json({ status: “ok” });
};

async function handleEvent(event, clients) {
const replyToken = event.replyToken;
const type = event.type;
const message = event.message;
const source = event.source;
const postback = event.postback;
const userId = source && source.userId ? source.userId : “unknown”;

try {
if (type === “follow”) {
await replyFlex(replyToken, buildMainMenuFlex());
return;
}
if (type === “postback” && postback && postback.data) {
await handlePostback(replyToken, postback.data, userId, clients);
return;
}
if (type !== “message” || !message) return;

```
if (message.type === "text") {
  await handleTextMessage(replyToken, message.text.trim(), userId, clients);
  return;
}
if (message.type === "image" || message.type === "video" || message.type === "file" || message.type === "audio") {
  await handleFileMessage(replyToken, message, message.type, userId, clients);
  return;
}
```

} catch (err) {
console.error(“handleEvent error:”, err);
if (replyToken) {
await replyText(replyToken, “ERROR: “ + err.message);
}
}
}

// ============================================================
//  Postback
// ============================================================

async function handlePostback(replyToken, data, userId, clients) {
const sheets = clients.sheets;
const params = new URLSearchParams(data);
const action = params.get(“action”);

if (action === “set_category”) {
await handleSetCategory(replyToken, params.get(“cat”), userId, sheets);
return;
}
if (action === “view_category”) {
await handleQueryCategory(replyToken, params.get(“cat”), userId, sheets);
return;
}
if (action === “todo_new”) {
await setBotState(sheets, userId, “awaiting_todo_input”);
await replyText(replyToken, “\u8ACB\u76F4\u63A5\u8F38\u5165\u5F85\u8FA6\u5167\u5BB9\uFF1A\n\n\u4F8B\u5982\uFF1A\u6574\u7406\u5831\u50F9\u55AE\n\u6216\uFF1A\u660E\u5929\u4E0B\u53483\u9EDE\u63D0\u9192\u6211\u56DE\u8986\u5BA2\u6236”);
return;
}
if (action === “todo_view”) {
await handleShowTodos(replyToken, userId, sheets);
return;
}
if (action === “todo_complete”) {
await handlePromptCompleteTodo(replyToken, userId, sheets);
return;
}
if (action === “todo_delete”) {
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

if (state === “awaiting_todo_input”) {
await clearBotState(sheets, userId);
await handleCreateTodo(replyToken, text, userId, sheets, calendar);
return;
}

if (state === “awaiting_todo_complete”) {
const num = parseInt(text.trim(), 10);
if (!isNaN(num)) {
await clearBotState(sheets, userId);
await handleCompleteTodoByIndex(replyToken, num, userId, sheets);
return;
}
await replyText(replyToken, “\u8ACB\u8F38\u5165\u6578\u5B57\u5E8F\u865F\uFF0C\u4F8B\u5982\uFF1A1\n\n\u6216\u8F38\u5165\u300C\u53D6\u6D88\u300D\u96E2\u958B”);
return;
}

if (state === “awaiting_todo_delete”) {
const num = parseInt(text.trim(), 10);
if (!isNaN(num)) {
await clearBotState(sheets, userId);
await handleDeleteTodoByIndex(replyToken, num, userId, sheets);
return;
}
await replyText(replyToken, “\u8ACB\u8F38\u5165\u6578\u5B57\u5E8F\u865F\uFF0C\u4F8B\u5982\uFF1A1\n\n\u6216\u8F38\u5165\u300C\u53D6\u6D88\u300D\u96E2\u958B”);
return;
}

if (text === “\u53D6\u6D88”) {
await clearBotState(sheets, userId);
await replyText(replyToken, “\u5DF2\u53D6\u6D88\u64CD\u4F5C\u3002”);
return;
}

if (text === “\u529F\u80FD\u8AAA\u660E” || text === “\u8AAA\u660E” || text === “help”) {
await replyFlex(replyToken, buildMainMenuFlex());
return;
}

if (text === “\u5F85\u8FA6\u6E05\u55AE”) {
await replyFlex(replyToken, buildTodoMenuFlex());
return;
}
if (text === “\u67E5\u770B\u5F85\u8FA6”) {
await handleShowTodos(replyToken, userId, sheets);
return;
}
if (text === “\u65B0\u589E\u5F85\u8FA6”) {
await setBotState(sheets, userId, “awaiting_todo_input”);
await replyText(replyToken, “\u8ACB\u76F4\u63A5\u8F38\u5165\u5F85\u8FA6\u5167\u5BB9\uFF1A\n\n\u4F8B\u5982\uFF1A\u6574\u7406\u5831\u50F9\u55AE\n\u6216\uFF1A\u660E\u5929\u4E0B\u53483\u9EDE\u63D0\u9192\u6211\u56DE\u8986\u5BA2\u6236”);
return;
}
if (text.indexOf(”\u65B0\u589E\u5F85\u8FA6 “) === 0) {
const content = text.slice(5).trim();
await handleCreateTodo(replyToken, content, userId, sheets, calendar);
return;
}
if (text === “\u5B8C\u6210\u5F85\u8FA6”) {
await handlePromptCompleteTodo(replyToken, userId, sheets);
return;
}
if (/^\u5B8C\u6210\u5F85\u8FA6\s+\d+$/.test(text)) {
const index = parseInt(text.replace(/^\u5B8C\u6210\u5F85\u8FA6\s+/, “”), 10);
await handleCompleteTodoByIndex(replyToken, index, userId, sheets);
return;
}
if (text === “\u522A\u9664\u5F85\u8FA6”) {
await handlePromptDeleteTodo(replyToken, userId, sheets);
return;
}
if (/^\u522A\u9664\u5F85\u8FA6\s+\d+$/.test(text)) {
const index = parseInt(text.replace(/^\u522A\u9664\u5F85\u8FA6\s+/, “”), 10);
await handleDeleteTodoByIndex(replyToken, index, userId, sheets);
return;
}

for (let i = 0; i < CATEGORIES.length; i++) {
const cat = CATEGORIES[i];
if (text === cat + “\u6536\u85CF” || text === cat) {
await handleQueryCategory(replyToken, cat, userId, sheets);
return;
}
}

const urls = extractUrls(text);
if (urls.length > 0) {
await handleLinkSave(replyToken, text, urls, userId, sheets);
return;
}

if (text.indexOf(”\u6A19\u7C64 “) === 0 || text.indexOf(”\u52A0\u6A19\u7C64 “) === 0) {
await handleAddTag(replyToken, text, userId, sheets);
return;
}
if (text.indexOf(”\u67E5\u6A19\u7C64 “) === 0) {
const tag = text.slice(4).trim();
await handleQueryTag(replyToken, tag, userId, sheets);
return;
}
if (text.charAt(0) === “#”) {
const tag = text.slice(1).trim();
await handleQueryTag(replyToken, tag, userId, sheets);
return;
}

if (text.indexOf(”\u641C\u5C0B “) === 0 || text.indexOf(”\u627E “) === 0) {
const keyword = text.indexOf(”\u641C\u5C0B “) === 0 ? text.slice(3).trim() : text.slice(2).trim();
await handleSearchLinks(replyToken, keyword, userId, sheets);
return;
}

if (text === “\u6700\u8FD1\u6536\u85CF”) {
await handleRecentLinks(replyToken, userId, sheets);
return;
}
if (text === “\u6700\u8FD1\u6A94\u6848”) {
await handleRecentFiles(replyToken, userId, sheets);
return;
}

if (/\u660E\u5929|\u884C\u7A0B|\u63D0\u9192|\u6703\u8B70|\u4ECA\u5929|\u4E0B\u5348|\u4E0A\u5348|\u65E9\u4E0A|\u5F8C\u5929|\u4E0B\u9031|\u4E0B\u5468|\u9031\u4E00|\u9031\u4E8C|\u9031\u4E09|\u9031\u56DB|\u9031\u4E94|\u9031\u516D|\u9031\u65E5/.test(text)) {
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
type: “flex”,
altText: “\u5F85\u8FA6\u529F\u80FD\u9078\u55AE”,
contents: {
type: “bubble”,
styles: { header: { backgroundColor: “#2D3561” } },
header: {
type: “box”,
layout: “vertical”,
paddingAll: “16px”,
contents: [
{ type: “text”, text: “\u5F85\u8FA6\u529F\u80FD\u9078\u55AE”, color: “#FFFFFF”, weight: “bold”, size: “md” },
{ type: “text”, text: “\u9EDE\u4E0B\u65B9\u6309\u9215\u64CD\u4F5C”, color: “#FFFFFFCC”, size: “xs” },
],
},
body: {
type: “box”,
layout: “vertical”,
paddingAll: “16px”,
spacing: “md”,
contents: [
{ type: “text”, text: “\u65B0\u589E\u5F85\u8FA6\u7BC4\u4F8B”, weight: “bold”, size: “sm”, color: “#333333” },
{ type: “text”, text: “\u6574\u7406\u5831\u50F9\u55AE”, size: “sm”, color: “#666666”, wrap: true },
{ type: “text”, text: “\u660E\u5929\u4E0B\u53483\u9EDE\u63D0\u9192\u6211\u56DE\u8986\u5BA2\u6236”, size: “sm”, color: “#666666”, wrap: true },
{ type: “text”, text: “\u9EDE\u300C\u65B0\u589E\u5F85\u8FA6\u300D\u5F8C\u76F4\u63A5\u8F38\u5165\u5167\u5BB9\u5373\u53EF”, size: “xs”, color: “#999999”, wrap: true },
],
},
footer: {
type: “box”,
layout: “vertical”,
spacing: “sm”,
paddingAll: “12px”,
contents: [
{
type: “box”,
layout: “horizontal”,
spacing: “sm”,
contents: [
{ type: “button”, style: “secondary”, height: “sm”, flex: 1, action: { type: “postback”, label: “\u67E5\u770B\u5F85\u8FA6”, data: “action=todo_view”, displayText: “\u67E5\u770B\u5F85\u8FA6” } },
{ type: “button”, style: “primary”, height: “sm”, flex: 1, color: “#2D3561”, action: { type: “postback”, label: “\u65B0\u589E\u5F85\u8FA6”, data: “action=todo_new”, displayText: “\u65B0\u589E\u5F85\u8FA6” } },
],
},
{
type: “box”,
layout: “horizontal”,
spacing: “sm”,
contents: [
{ type: “button”, style: “secondary”, height: “sm”, flex: 1, action: { type: “postback”, label: “\u5B8C\u6210\u5F85\u8FA6”, data: “action=todo_complete”, displayText: “\u5B8C\u6210\u5F85\u8FA6” } },
{ type: “button”, style: “secondary”, height: “sm”, flex: 1, action: { type: “postback”, label: “\u522A\u9664\u5F85\u8FA6”, data: “action=todo_delete”, displayText: “\u522A\u9664\u5F85\u8FA6” } },
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
.filter(function(r) { return String(r[6] || “”) === “todo” && String(r[5] || “”) === userId; })
.filter(function(r) { return String(r[2] || “pending”) !== “deleted”; });
}

async function handleCreateTodo(replyToken, content, userId, sheets, calendar) {
const id = generateId(“TD”);
const now = formatDateTime(new Date());

const parsed = parseCalendarInput(content);
let remindText = “”;
let googleEventId = “”;

if (parsed) {
remindText = parsed.isAllDay ? formatDate(parsed.start) : formatDate(parsed.start) + “ “ + formatTime(parsed.start);
try {
const event = await createCalendarEvent(calendar, parsed);
googleEventId = event.id || “”;
} catch (e) {
console.error(“Calendar create error:”, e);
}
}

const row = [id, content, “pending”, now, remindText, userId, “todo”, googleEventId];
await appendRow(sheets, SHEET.TODOS, row);

const bodyContents = [
{ type: “text”, text: “\u5DF2\u65B0\u589E\u5F85\u8FA6”, weight: “bold”, size: “md”, color: “#0F9D58” },
{ type: “text”, text: content, size: “sm”, color: “#333333”, wrap: true },
];
if (remindText) bodyContents.push(makeRow(”\u63D0\u9192”, remindText));
if (googleEventId) bodyContents.push({ type: “text”, text: “\u5DF2\u540C\u6B65\u5230 Google \u65E5\u66C6”, size: “xs”, color: “#0F9D58” });

await replyFlex(replyToken, {
type: “flex”,
altText: “\u5DF2\u65B0\u589E\u5F85\u8FA6”,
contents: {
type: “bubble”,
body: { type: “box”, layout: “vertical”, paddingAll: “16px”, spacing: “md”, contents: bodyContents },
footer: {
type: “box”, layout: “horizontal”, spacing: “sm”, paddingAll: “12px”,
contents: [
{ type: “button”, style: “secondary”, height: “sm”, flex: 1, action: { type: “postback”, label: “\u67E5\u770B\u5F85\u8FA6”, data: “action=todo_view”, displayText: “\u67E5\u770B\u5F85\u8FA6” } },
{ type: “button”, style: “primary”, height: “sm”, flex: 1, color: “#2D3561”, action: { type: “postback”, label: “\u518D\u65B0\u589E\u4E00\u7B46”, data: “action=todo_new”, displayText: “\u65B0\u589E\u5F85\u8FA6” } },
],
},
},
});
}

async function handleShowTodos(replyToken, userId, sheets) {
const rows = (await getTodoRows(sheets, userId)).filter(function(r) { return String(r[2] || “pending”) === “pending”; });
if (!rows.length) {
await replyText(replyToken, “\u76EE\u524D\u6C92\u6709\u5F85\u8FA6\u3002\n\n\u8F38\u5165\u300C\u65B0\u589E\u5F85\u8FA6\u300D\u65B0\u589E\u7B2C\u4E00\u7B46”);
return;
}
const lines = rows.slice(0, 10).map(function(r, i) {
const content = String(r[1] || “”);
const remind = String(r[4] || “”);
return remind ? (i + 1) + “. “ + content + “\n   [” + remind + “]” : (i + 1) + “. “ + content;
});
await replyText(replyToken, “\u76EE\u524D\u5F85\u8FA6\uFF08” + rows.length + “ \u7B46\uFF09\n\n” + lines.join(”\n\n”) + “\n\n\u5B8C\u6210\u8ACB\u8F38\u5165\uFF1A\u5B8C\u6210\u5F85\u8FA6 \u5E8F\u865F”);
}

async function handlePromptCompleteTodo(replyToken, userId, sheets) {
const rows = (await getTodoRows(sheets, userId)).filter(function(r) { return String(r[2] || “pending”) === “pending”; });
if (!rows.length) { await replyText(replyToken, “\u76EE\u524D\u6C92\u6709\u53EF\u5B8C\u6210\u7684\u5F85\u8FA6\u3002”); return; }
const lines = rows.slice(0, 10).map(function(r, i) {
const remind = String(r[4] || “”);
return remind ? (i + 1) + “. “ + String(r[1] || “”) + “ (” + remind + “)” : (i + 1) + “. “ + String(r[1] || “”);
});
await replyText(replyToken, “\u8ACB\u8F38\u5165\u8981\u5B8C\u6210\u7684\u5E8F\u865F\uFF1A\n\n” + lines.join(”\n”) + “\n\n\u4F8B\u5982\u8F38\u5165\uFF1A\u5B8C\u6210\u5F85\u8FA6 1”);
}

async function handleCompleteTodoByIndex(replyToken, index, userId, sheets) {
const allData = await getSheetData(sheets, SHEET.TODOS);
const pendingRows = allData.map(function(row, idx) { return { row: row, idx: idx }; }).filter(function(item) {
return String(item.row[6] || “”) === “todo” && String(item.row[5] || “”) === userId && String(item.row[2] || “pending”) === “pending”;
});

if (!index || index < 1 || index > pendingRows.length) {
await replyText(replyToken, “\u5E8F\u865F\u7121\u6548\uFF0C\u8ACB\u8F38\u5165 1~” + pendingRows.length + “ \u4E4B\u9593\u7684\u6578\u5B57\u3002”);
return;
}

const target = pendingRows[index - 1];
const rowNum = target.idx + 1;
const content = String(target.row[1] || “”);

await sheets.spreadsheets.values.update({
spreadsheetId: SHEETS_ID,
range: SHEET.TODOS + “!C” + rowNum,
valueInputOption: “RAW”,
requestBody: { values: [[“done”]] },
});

await replyText(replyToken, “\u5DF2\u5B8C\u6210\u5F85\u8FA6\n” + content);
}

async function handlePromptDeleteTodo(replyToken, userId, sheets) {
const rows = (await getTodoRows(sheets, userId)).filter(function(r) { return String(r[2] || “pending”) === “pending”; });
if (!rows.length) { await replyText(replyToken, “\u76EE\u524D\u6C92\u6709\u53EF\u522A\u9664\u7684\u5F85\u8FA6\u3002”); return; }
const lines = rows.slice(0, 10).map(function(r, i) {
const remind = String(r[4] || “”);
return remind ? (i + 1) + “. “ + String(r[1] || “”) + “ (” + remind + “)” : (i + 1) + “. “ + String(r[1] || “”);
});
await replyText(replyToken, “\u8ACB\u8F38\u5165\u8981\u522A\u9664\u7684\u5E8F\u865F\uFF1A\n\n” + lines.join(”\n”) + “\n\n\u4F8B\u5982\u8F38\u5165\uFF1A\u522A\u9664\u5F85\u8FA6 1”);
}

async function handleDeleteTodoByIndex(replyToken, index, userId, sheets) {
const allData = await getSheetData(sheets, SHEET.TODOS);
const pendingRows = allData.map(function(row, idx) { return { row: row, idx: idx }; }).filter(function(item) {
return String(item.row[6] || “”) === “todo” && String(item.row[5] || “”) === userId && String(item.row[2] || “pending”) === “pending”;
});

if (!index || index < 1 || index > pendingRows.length) {
await replyText(replyToken, “\u5E8F\u865F\u7121\u6548\uFF0C\u8ACB\u8F38\u5165 1~” + pendingRows.length + “ \u4E4B\u9593\u7684\u6578\u5B57\u3002”);
return;
}

const target = pendingRows[index - 1];
const rowNum = target.idx + 1;
const content = String(target.row[1] || “”);

await sheets.spreadsheets.values.update({
spreadsheetId: SHEETS_ID,
range: SHEET.TODOS + “!C” + rowNum,
valueInputOption: “RAW”,
requestBody: { values: [[“deleted”]] },
});

await replyText(replyToken, “\u5DF2\u522A\u9664\u5F85\u8FA6\n” + content);
}

// ============================================================
//  Bookmarks
// ============================================================

function extractUrls(text) {
const regex = /https?://[^\s]+/gi;
return text.match(regex) || [];
}

function detectPlatform(url) {
if (/threads.(net|com)/i.test(url)) return { name: “Threads”, color: “#000000” };
if (/facebook.com|fb.com/i.test(url)) return { name: “Facebook”, color: “#1877F2” };
if (/youtube.com|youtu.be/i.test(url)) return { name: “YouTube”, color: “#FF0000” };
if (/instagram.com/i.test(url)) return { name: “Instagram”, color: “#E1306C” };
if (/twitter.com|x.com/i.test(url)) return { name: “X”, color: “#000000” };
if (/linkedin.com/i.test(url)) return { name: “LinkedIn”, color: “#0077B5” };
if (/github.com/i.test(url)) return { name: “GitHub”, color: “#333333” };
if (/notion.so/i.test(url)) return { name: “Notion”, color: “#000000” };
if (/medium.com/i.test(url)) return { name: “Medium”, color: “#000000” };
return { name: “Link”, color: “#555555” };
}

async function handleLinkSave(replyToken, rawText, urls, userId, sheets) {
let note = rawText;
urls.forEach(function(u) { note = note.replace(u, “”).trim(); });
note = note.replace(/^[\s-]+/, “”).trim();

const saved = [];
for (let i = 0; i < urls.length; i++) {
const url = urls[i];
const isDup = await checkDuplicateThreads(url, userId, sheets);
if (isDup) { saved.push({ url: url, duplicate: true }); continue; }
const platform = detectPlatform(url);
const id = generateId(“T”);
const now = formatDateTime(new Date());
const row = [id, url, platform.name, “”, note, “”, “\u672A\u5206\u985E”, userId, now, now, rawText];
await appendRow(sheets, SHEET.THREADS, row);
saved.push({ url: url, platform: platform, duplicate: false, note: note });
}

if (saved.length === 1 && !saved[0].duplicate) {
await replyFlex(replyToken, buildLinkSavedFlex(saved[0].url, saved[0].platform, saved[0].note));
return;
}
if (saved.length === 1 && saved[0].duplicate) {
await replyText(replyToken, “\u9019\u5247\u9023\u7D50\u4E4B\u524D\u5DF2\u7D93\u6536\u85CF\u904E\u3002”);
return;
}
const lines = saved.map(function(s, i) {
return s.duplicate ? (i + 1) + “. \u5DF2\u6536\u85CF\u904E” : (i + 1) + “. [” + s.platform.name + “] \u5DF2\u6536\u85CF”;
});
await replyText(replyToken, “\u6536\u85CF “ + urls.length + “ \u5247\u9023\u7D50\n\n” + lines.join(”\n”));
}

function buildLinkSavedFlex(url, platform, note) {
const now = formatDateTime(new Date());
return {
type: “flex”,
altText: “\u5DF2\u6536\u85CF\u6210\u529F”,
contents: {
type: “bubble”,
size: “mega”,
header: {
type: “box”, layout: “vertical”, paddingAll: “16px”, backgroundColor: “#1A1A2E”,
contents: [
{ type: “text”, text: “\u5DF2\u6536\u85CF\u6210\u529F”, color: “#FFFFFF”, weight: “bold”, size: “md” },
{ type: “text”, text: platform.name, color: “#FFFFFFCC”, size: “sm” },
],
},
body: {
type: “box”, layout: “vertical”, paddingAll: “16px”, spacing: “md”,
contents: [
makeRow(”\u5099\u8A3B”, note || “(\u672A\u586B\u5BEB)”),
makeRow(”\u5206\u985E”, “\u672A\u5206\u985E”),
makeRow(”\u6642\u9593”, now),
{ type: “separator” },
{ type: “text”, text: “\u8ACB\u9078\u64C7\u6536\u85CF\u5206\u985E”, size: “sm”, color: “#555555”, weight: “bold” },
{
type: “box”, layout: “horizontal”, spacing: “sm”,
contents: [”\u500B\u4EBA”, “\u5DE5\u4F5C”, “\u5BB6\u5EAD”].map(function(cat) {
return { type: “button”, style: “secondary”, height: “sm”, flex: 1, action: { type: “postback”, label: cat, data: “action=set_category&id=LAST&cat=” + cat, displayText: “\u8A2D\u70BA” + cat } };
}),
},
],
},
footer: {
type: “box”, layout: “horizontal”, spacing: “sm”, paddingAll: “12px”,
contents: [
{ type: “button”, style: “primary”, flex: 2, color: “#1A1A2E”, action: { type: “uri”, label: “\u958B\u555F\u9023\u7D50”, uri: url } },
{ type: “button”, style: “secondary”, flex: 1, action: { type: “message”, label: “\u52A0\u6A19\u7C64”, text: “\u52A0\u6A19\u7C64 “ } },
],
},
},
};
}

async function handleSetCategory(replyToken, category, userId, sheets) {
const result = await findLastUserRow(sheets, SHEET.THREADS, userId);
const lastRow = result.row;
const lastIndex = result.index;
if (lastIndex === -1) { await replyText(replyToken, “\u627E\u4E0D\u5230\u6700\u8FD1\u7684\u6536\u85CF\u3002”); return; }

const rowNum = lastIndex + 1;
const now = formatDateTime(new Date());
await sheets.spreadsheets.values.batchUpdate({
spreadsheetId: SHEETS_ID,
requestBody: {
valueInputOption: “RAW”,
data: [
{ range: SHEET.THREADS + “!G” + rowNum, values: [[category]] },
{ range: SHEET.THREADS + “!J” + rowNum, values: [[now]] },
],
},
});

const url = String(lastRow[1] || “”);
await replyFlex(replyToken, {
type: “flex”,
altText: “\u5DF2\u8A2D\u70BA” + category,
contents: {
type: “bubble”,
body: {
type: “box”, layout: “vertical”, paddingAll: “20px”, spacing: “md”,
contents: [
{ type: “text”, text: “\u5DF2\u6B78\u985E\u5230\u300C” + category + “\u300D”, weight: “bold”, size: “md”, color: “#333333” },
{ type: “text”, text: “\u4E4B\u5F8C\u53EF\u7528\u300C” + category + “\u6536\u85CF\u300D\u67E5\u8A62”, size: “xs”, color: “#888888” },
],
},
footer: {
type: “box”, layout: “horizontal”, spacing: “sm”, paddingAll: “12px”,
contents: [
{ type: “button”, style: “primary”, height: “sm”, flex: 1, color: CATEGORY_COLORS[category] || “#1A1A2E”, action: { type: “postback”, label: “\u67E5\u770B” + category, data: “action=view_category&cat=” + category, displayText: category + “\u6536\u85CF” } },
{ type: “button”, style: “secondary”, height: “sm”, flex: 1, action: { type: “uri”, label: “\u958B\u555F\u9023\u7D50”, uri: url } },
],
},
},
});
}

async function handleQueryCategory(replyToken, category, userId, sheets) {
const data = await getSheetData(sheets, SHEET.THREADS);
const all = data.slice(1).filter(function(row) { return String(row[7] || “”) === userId && String(row[6] || “\u672A\u5206\u985E”) === category; }).reverse();
const results = all.slice(0, 10);

if (!results.length) {
await replyText(replyToken, category + “ \u9084\u6C92\u6709\u6536\u85CF\u3002\n\n\u6536\u85CF\u9023\u7D50\u5F8C\uFF0C\u9EDE\u6210\u529F\u5361\u7247\u4E0B\u65B9\u7684\u5206\u985E\u6309\u9215\u5373\u53EF\u6B78\u985E\u3002”);
return;
}
await replyFlex(replyToken, buildLinksCarousel(results, category + “ - “ + all.length + “ \u7B46”));
}

async function handleAddTag(replyToken, text, userId, sheets) {
const tagStr = text.replace(/^(\u6A19\u7C64|\u52A0\u6A19\u7C64)\s+/, “”).trim();
const tags = tagStr.split(/[\s,]+/).filter(function(t) { return t.length > 0; });
const result = await findLastUserRow(sheets, SHEET.THREADS, userId);
if (result.index === -1) { await replyText(replyToken, “\u627E\u4E0D\u5230\u6700\u8FD1\u7684\u6536\u85CF\u3002”); return; }

const tagValue = tags.join(”\u3001”);
const rowNum = result.index + 1;
const now = formatDateTime(new Date());
await sheets.spreadsheets.values.batchUpdate({
spreadsheetId: SHEETS_ID,
requestBody: {
valueInputOption: “RAW”,
data: [
{ range: SHEET.THREADS + “!F” + rowNum, values: [[tagValue]] },
{ range: SHEET.THREADS + “!J” + rowNum, values: [[now]] },
],
},
});
await replyText(replyToken, “\u6A19\u7C64\u5DF2\u66F4\u65B0\n” + tags.map(function(t) { return “#” + t; }).join(” “));
}

async function handleQueryTag(replyToken, tag, userId, sheets) {
if (!tag) { await replyText(replyToken, “\u8ACB\u544A\u8A34\u6211\u8981\u67E5\u8A62\u7684\u6A19\u7C64”); return; }
const data = await getSheetData(sheets, SHEET.THREADS);
const all = data.slice(1).filter(function(row) { return String(row[7] || “”) === userId && String(row[5] || “”).includes(tag); });
const results = all.reverse().slice(0, 10);
if (!results.length) { await replyText(replyToken, “\u6C92\u6709\u6A19\u7C64\u300C” + tag + “\u300D\u7684\u6536\u85CF\u3002”); return; }
await replyFlex(replyToken, buildLinksCarousel(results, “\u6A19\u7C64 “ + tag + “ - “ + all.length + “ \u7B46”));
}

async function handleRecentLinks(replyToken, userId, sheets) {
const data = await getSheetData(sheets, SHEET.THREADS);
const results = data.slice(1).filter(function(r) { return String(r[7] || “”) === userId; }).reverse().slice(0, 10);
if (!results.length) { await replyText(replyToken, “\u9084\u6C92\u6709\u6536\u85CF\u3002\n\n\u8CBC\u4E0A\u4EFB\u610F\u9023\u7D50\u5C31\u6703\u81EA\u52D5\u6536\u85CF\u3002”); return; }
await replyFlex(replyToken, buildLinksCarousel(results, “\u6700\u8FD1 “ + results.length + “ \u5247\u6536\u85CF”));
}

async function handleSearchLinks(replyToken, keyword, userId, sheets) {
const data = await getSheetData(sheets, SHEET.THREADS);
const results = data.slice(1)
.filter(function(row) { return String(row[7] || “”) === userId; })
.filter(function(row) { return [row[1], row[4], row[5], row[6], row[10]].join(” “).includes(keyword); })
.reverse().slice(0, 8);
if (!results.length) { await replyText(replyToken, “\u6C92\u6709\u627E\u5230\u300C” + keyword + “\u300D\u7684\u6536\u85CF\u3002”); return; }
await replyFlex(replyToken, buildLinksCarousel(results, keyword + “ - “ + results.length + “ \u7B46”));
}

function buildLinksCarousel(rows, headerTitle) {
if (rows.length <= 5) {
return { type: “flex”, altText: headerTitle, contents: { type: “carousel”, contents: rows.map(function(r) { return buildLinkCardBubble(r); }) } };
}
return {
type: “flex”, altText: headerTitle,
contents: {
type: “bubble”, size: “mega”,
header: { type: “box”, layout: “vertical”, backgroundColor: “#1A1A2E”, paddingAll: “14px”, contents: [{ type: “text”, text: headerTitle, color: “#FFFFFF”, weight: “bold”, size: “sm” }] },
body: { type: “box”, layout: “vertical”, paddingAll: “12px”, spacing: “sm”, contents: rows.slice(0, 8).map(function(r, i) { return buildLinkListRow(r, i); }) },
footer: { type: “box”, layout: “vertical”, paddingAll: “10px”, contents: [{ type: “button”, style: “link”, height: “sm”, action: { type: “message”, label: “\u67E5\u770B\u6700\u8FD1\u6536\u85CF”, text: “\u6700\u8FD1\u6536\u85CF” } }] },
},
};
}

function buildLinkCardBubble(r) {
const url = String(r[1] || “”);
const platform = detectPlatform(url);
const note = String(r[4] || “”).substring(0, 40) || “(\u672A\u586B\u5099\u8A3B)”;
const tags = String(r[5] || “”);
const category = String(r[6] || “\u672A\u5206\u985E”);
const date = String(r[8] || r[7] || “”).substring(0, 10);

const bodyContents = [
{ type: “text”, text: note, size: “sm”, color: “#333333”, wrap: true, maxLines: 3 },
makeRow(”\u65E5\u671F”, date),
];
if (tags) bodyContents.splice(1, 0, { type: “text”, text: “#” + tags, size: “xs”, color: “#888888” });

return {
type: “bubble”, size: “kilo”,
styles: { header: { backgroundColor: platform.color } },
header: { type: “box”, layout: “horizontal”, paddingAll: “10px”, contents: [
{ type: “text”, text: platform.name, size: “xs”, flex: 1, color: “#FFFFFF”, gravity: “center” },
{ type: “text”, text: category, size: “xs”, flex: 0, color: “#FFFFFF99” },
]},
body: { type: “box”, layout: “vertical”, paddingAll: “12px”, spacing: “sm”, contents: bodyContents },
footer: { type: “box”, layout: “vertical”, paddingAll: “10px”, contents: [{ type: “button”, style: “primary”, height: “sm”, color: platform.color || “#1A1A2E”, action: { type: “uri”, label: “\u958B\u555F\u9023\u7D50”, uri: url } }] },
};
}

function buildLinkListRow(r, i) {
const url = String(r[1] || “”);
const platform = detectPlatform(url);
const note = String(r[4] || “”).substring(0, 30) || “(\u672A\u586B\u5099\u8A3B)”;
const category = String(r[6] || “\u672A\u5206\u985E”);
const date = String(r[8] || r[7] || “”).substring(0, 10);
return {
type: “box”, layout: “vertical”, spacing: “xs”, paddingTop: i === 0 ? “0px” : “8px”,
contents: [
{ type: “box”, layout: “horizontal”, contents: [
{ type: “text”, text: “[” + platform.name + “] “ + note, size: “sm”, color: “#333333”, flex: 1, wrap: true, maxLines: 2 },
{ type: “button”, style: “link”, height: “sm”, flex: 0, action: { type: “uri”, label: “\u958B\u555F”, uri: url } },
]},
{ type: “text”, text: category + “  “ + date, size: “xs”, color: “#888888” },
{ type: “separator” },
],
};
}

async function checkDuplicateThreads(url, userId, sheets) {
const data = await getSheetData(sheets, SHEET.THREADS);
return data.slice(1).some(function(row) { return String(row[1] || “”) === url && String(row[7] || “”) === userId; });
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
const id = generateId(“F”);
const now = formatDateTime(new Date());
const row = [id, fileName, fileType, mimeType, fileBuffer.length, driveFile.id, driveFile.webViewLink, “”, “”, userId, message.id, now];
await appendRow(sheets, SHEET.FILES, row);
await replyFlex(replyToken, buildFileSavedFlex(fileName, fileType, driveFile.webViewLink, now));
} catch (err) {
const msg = String(err && err.message ? err.message : err || “”);
console.error(“File upload error:”, msg);
await replyText(replyToken, “\u6A94\u6848\u5099\u4EFD\u5931\u6557\uFF1A” + msg);
}
}

function downloadLineContent(messageId) {
return new Promise(function(resolve, reject) {
https.get({
hostname: “api-data.line.me”,
path: “/v2/bot/message/” + messageId + “/content”,
method: “GET”,
headers: { Authorization: “Bearer “ + LINE_TOKEN },
}, function(res) {
if (res.statusCode && res.statusCode >= 400) { reject(new Error(“LINE content download failed: “ + res.statusCode)); return; }
const chunks = [];
res.on(“data”, function(chunk) { chunks.push(chunk); });
res.on(“end”, function() { resolve(Buffer.concat(chunks)); });
res.on(“error”, reject);
}).on(“error”, reject);
});
}

async function uploadToDrive(drive, fileName, mimeType, buffer, fileType) {
const folderId = await getOrCreateFolder(drive, fileType);
const stream = Readable.from(buffer);
const result = await drive.files.create({
requestBody: { name: fileName, parents: [folderId] },
media: { mimeType: mimeType, body: stream },
fields: “id,webViewLink,name”,
});
return result.data;
}

async function getOrCreateFolder(drive, folderName) {
if (!DRIVE_FOLDER_ID) {
const folder = await drive.files.create({
requestBody: { name: “LINE Bot \u5099\u4EFD”, mimeType: “application/vnd.google-apps.folder” },
fields: “id”,
});
return folder.data.id;
}

const res = await drive.files.list({
q: “’” + DRIVE_FOLDER_ID + “’ in parents and name=’” + folderName + “’ and mimeType=‘application/vnd.google-apps.folder’ and trashed=false”,
fields: “files(id)”,
});
if (res.data.files.length > 0) return res.data.files[0].id;

const folder = await drive.files.create({
requestBody: { name: folderName, mimeType: “application/vnd.google-apps.folder”, parents: [DRIVE_FOLDER_ID] },
fields: “id”,
});
return folder.data.id;
}

function buildFileSavedFlex(fileName, fileType, driveUrl, now) {
return {
type: “flex”, altText: “\u6A94\u6848\u5099\u4EFD\u5B8C\u6210”,
contents: {
type: “bubble”,
styles: { header: { backgroundColor: “#1a73e8” } },
header: { type: “box”, layout: “vertical”, paddingAll: “16px”, contents: [{ type: “text”, text: “\u6A94\u6848\u5099\u4EFD\u5B8C\u6210”, color: “#ffffff”, weight: “bold”, size: “md” }] },
body: { type: “box”, layout: “vertical”, paddingAll: “16px”, spacing: “md”, contents: [makeRow(”\u6A94\u6848\u540D”, fileName), makeRow(”\u985E\u578B”, fileType.toUpperCase()), makeRow(”\u6642\u9593”, now)] },
footer: { type: “box”, layout: “vertical”, paddingAll: “12px”, contents: [{ type: “button”, style: “primary”, height: “sm”, color: “#1a73e8”, action: { type: “uri”, label: “\u67E5\u770B Google Drive”, uri: driveUrl } }] },
},
};
}

async function handleRecentFiles(replyToken, userId, sheets) {
const data = await getSheetData(sheets, SHEET.FILES);
const results = data.slice(1).filter(function(r) { return String(r[9] || “”) === userId; }).reverse().slice(0, 10);
if (!results.length) { await replyText(replyToken, “\u9084\u6C92\u6709\u5099\u4EFD\u6A94\u6848\u3002”); return; }
const lines = results.map(function(r, i) {
return (i + 1) + “. “ + String(r[1] || “”) + “\n   “ + String(r[2] || “”) + “  “ + String(r[11] || “”).substring(0, 10) + “\n   “ + String(r[6] || “”);
});
await replyText(replyToken, “\u6700\u8FD1 “ + results.length + “ \u7B46\u6A94\u6848\n\n” + lines.join(”\n\n”));
}

// ============================================================
//  Calendar
// ============================================================

async function handleCalendar(replyToken, text, userId, calendar, sheets) {
const parsed = parseCalendarInput(text);
if (!parsed) {
await replyText(replyToken, “\u6211\u627E\u4E0D\u5230\u65E5\u671F\uFF0C\u7121\u6CD5\u5EFA\u7ACB\u884C\u7A0B\u3002\n\n\u8ACB\u7528\u9019\u6A23\u7684\u683C\u5F0F\uFF1A\n\u660E\u5929 \u4E0B\u53483\u9EDE \u958B\u6703\n4/25 10:00 \u5BA2\u6236\u7C21\u5831”);
return;
}
try {
const event = await createCalendarEvent(calendar, parsed);
const id = generateId(“C”);
const now = formatDateTime(new Date());
const row = [id, parsed.title, formatDate(parsed.start), parsed.isAllDay ? “” : formatTime(parsed.start), parsed.isAllDay ? “TRUE” : “FALSE”, event.id, text, “text”, userId, now];
await appendRow(sheets, SHEET.CALENDAR, row);
await replyFlex(replyToken, buildCalendarFlex(parsed, event));
} catch (err) {
console.error(“Calendar error:”, err);
await replyText(replyToken, “\u5EFA\u7ACB\u884C\u7A0B\u5931\u6557\uFF1A” + err.message);
}
}

function parseCalendarInput(text) {
const now = new Date();
now.setHours(0, 0, 0, 0);
let targetDate = null;
let targetTime = null;
let title = text;

if (/\u4ECA\u5929|\u4ECA\u65E5/.test(text)) {
targetDate = new Date(now);
title = title.replace(/\u4ECA\u5929|\u4ECA\u65E5/g, “”);
} else if (/\u660E\u5929|\u660E\u65E5/.test(text)) {
targetDate = new Date(now);
targetDate.setDate(targetDate.getDate() + 1);
title = title.replace(/\u660E\u5929|\u660E\u65E5/g, “”);
} else if (/\u5F8C\u5929/.test(text)) {
targetDate = new Date(now);
targetDate.setDate(targetDate.getDate() + 2);
title = title.replace(/\u5F8C\u5929/g, “”);
} else {
const weekMatch = text.match(/\u4E0B[\u9031\u5468]?(\u4E00|\u4E8C|\u4E09|\u56DB|\u4E94|\u516D|\u65E5|\u5929)?/);
if (weekMatch && weekMatch[1]) {
const dayMap = { “\u4E00”: 1, “\u4E8C”: 2, “\u4E09”: 3, “\u56DB”: 4, “\u4E94”: 5, “\u516D”: 6, “\u65E5”: 0, “\u5929”: 0 };
const target = dayMap[weekMatch[1]];
targetDate = new Date(now);
let diff = (target - targetDate.getDay() + 7) % 7;
if (diff === 0) diff = 7;
diff += 7;
targetDate.setDate(targetDate.getDate() + diff);
title = title.replace(weekMatch[0], “”);
} else {
const mdMatch = text.match(/(\d{1,2})[/\u6708](\d{1,2})[\u65E5]?/);
if (mdMatch) {
targetDate = new Date(now.getFullYear(), parseInt(mdMatch[1], 10) - 1, parseInt(mdMatch[2], 10));
title = title.replace(mdMatch[0], “”);
}
}
}

if (!targetDate) return null;

const isPM = /\u4E0B\u5348|\u665A\u4E0A|pm/i.test(text);
const isAM = /\u4E0A\u5348|\u65E9\u4E0A|am/i.test(text);
const timeMatch = text.match(/(\d{1,2})[:點時](\d{0,2})/);

if (timeMatch) {
let hour = parseInt(timeMatch[1], 10);
const min = parseInt(timeMatch[2] || “0”, 10);
if (isPM && hour < 12) hour += 12;
if (isAM && hour === 12) hour = 0;
targetTime = { hour: hour, min: min };
title = title.replace(timeMatch[0], “”);
}

title = title.replace(/\u4E0A\u5348|\u65E9\u4E0A|\u4E0B\u5348|\u665A\u4E0A|pm|am/gi, “”).replace(/\s+/g, “ “).trim();
if (!title) title = “\u884C\u7A0B\u5F85\u8FA6”;

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
reminders: { useDefault: false, overrides: parsed.isAllDay ? [] : [{ method: “popup”, minutes: 10 }] },
};
if (parsed.isAllDay) {
body.start = { date: formatDate(parsed.start) };
body.end = { date: formatDate(parsed.end) };
} else {
body.start = { dateTime: parsed.start.toISOString(), timeZone: “Asia/Taipei” };
body.end = { dateTime: parsed.end.toISOString(), timeZone: “Asia/Taipei” };
}
const result = await calendar.events.insert({ calendarId: CALENDAR_ID, requestBody: body });
return result.data;
}

function buildCalendarFlex(parsed, event) {
const timeStr = parsed.isAllDay ? formatDate(parsed.start) + “(\u5168\u5929)” : formatDate(parsed.start) + “ “ + formatTime(parsed.start);
return {
type: “flex”, altText: “\u5DF2\u65B0\u589E\uFF1A” + parsed.title,
contents: {
type: “bubble”,
styles: { header: { backgroundColor: “#0F9D58” } },
header: { type: “box”, layout: “vertical”, paddingAll: “16px”, contents: [{ type: “text”, text: “\u5DF2\u65B0\u589E\u5230 Google \u65E5\u66C6”, color: “#ffffff”, weight: “bold”, size: “md” }] },
body: { type: “box”, layout: “vertical”, paddingAll: “16px”, spacing: “md”, contents: [makeRow(”\u884C\u7A0B”, parsed.title), makeRow(”\u6642\u9593”, timeStr), makeRow(”\u63D0\u9192”, parsed.isAllDay ? “(\u5168\u5929\uFF0C\u7121\u63D0\u9192)” : “\u958B\u59CB\u524D 10 \u5206\u9418”)] },
footer: event.htmlLink ? { type: “box”, layout: “vertical”, paddingAll: “12px”, contents: [{ type: “button”, style: “primary”, height: “sm”, color: “#0F9D58”, action: { type: “uri”, label: “\u67E5\u770B Google \u65E5\u66C6”, uri: event.htmlLink } }] } : undefined,
},
};
}

// ============================================================
//  Main Menu
// ============================================================

function buildMainMenuFlex() {
return {
type: “flex”,
altText: “\u4F60\u597D\uFF01\u6211\u53EF\u4EE5\u5E6B\u4F60\u505A\u9019\u4E9B\u4E8B”,
contents: {
type: “bubble”, size: “mega”,
styles: { header: { backgroundColor: “#1A1A2E” } },
header: {
type: “box”, layout: “vertical”, paddingAll: “20px”,
contents: [
{ type: “text”, text: “\u4F60\u597D\uFF01\u6211\u53EF\u4EE5\u5E6B\u4F60\u505A\u9019\u4E9B\u4E8B”, color: “#FFFFFF”, weight: “bold”, size: “md” },
{ type: “text”, text: “\u9078\u64C7\u529F\u80FD\u6216\u76F4\u63A5\u8F38\u5165”, color: “#FFFFFF88”, size: “xs” },
],
},
body: {
type: “box”, layout: “vertical”, paddingAll: “16px”, spacing: “lg”,
contents: [
menuItem(”[LINK]”, “\u8CBC\u9023\u7D50\u6536\u85CF”, “\u4EFB\u610F\u7DB2\u5740\u81EA\u52D5\u6536\u85CF\uFF0C\u53EF\u76F4\u63A5\u5206\u985E”),
{ type: “separator” },
menuItem(”[LIST]”, “\u5F85\u8FA6\u6E05\u55AE”, “\u67E5\u770B\u3001\u65B0\u589E\u3001\u5B8C\u6210\u3001\u522A\u9664\u5F85\u8FA6”),
{ type: “separator” },
menuItem(”[FILE]”, “\u50B3\u5716\u7247/\u6A94\u6848”, “\u81EA\u52D5\u5099\u4EFD\u5230 Google Drive”),
{ type: “separator” },
menuItem(”[CAL]”, “\u8AAA\u5F85\u8FA6\u4E8B\u9805”, “\u542B\u6642\u9593\u5247\u65B0\u589E\u5230 Google \u65E5\u66C6”),
],
},
footer: {
type: “box”, layout: “vertical”, paddingAll: “12px”, spacing: “sm”,
contents: [
{
type: “box”, layout: “horizontal”, spacing: “sm”,
contents: CATEGORIES.filter(function(c) { return c !== “\u672A\u5206\u985E”; }).map(function(cat) {
return { type: “button”, style: “secondary”, height: “sm”, flex: 1, action: { type: “postback”, label: cat, data: “action=view_category&cat=” + cat, displayText: cat + “\u6536\u85CF” } };
}),
},
{
type: “box”, layout: “horizontal”, spacing: “sm”,
contents: [
{ type: “button”, style: “secondary”, height: “sm”, flex: 1, action: { type: “message”, label: “\u6700\u8FD1\u6536\u85CF”, text: “\u6700\u8FD1\u6536\u85CF” } },
{ type: “button”, style: “secondary”, height: “sm”, flex: 1, action: { type: “message”, label: “\u67E5\u770B\u5F85\u8FA6”, text: “\u67E5\u770B\u5F85\u8FA6” } },
{ type: “button”, style: “primary”, height: “sm”, flex: 1, color: “#1A1A2E”, action: { type: “message”, label: “\u5F85\u8FA6\u6E05\u55AE”, text: “\u5F85\u8FA6\u6E05\u55AE” } },
],
},
],
},
},
};
}

function menuItem(icon, title, desc) {
return {
type: “box”, layout: “horizontal”, spacing: “md”,
contents: [
{ type: “text”, text: icon, size: “sm”, flex: 0, gravity: “center”, color: “#1A1A2E” },
{ type: “box”, layout: “vertical”, flex: 1, contents: [
{ type: “text”, text: title, size: “sm”, weight: “bold”, color: “#1A1A2E” },
{ type: “text”, text: desc, size: “xs”, color: “#888888”, wrap: true },
]},
],
};
}

// ============================================================
//  Utilities
// ============================================================

async function getSheetData(sheets, sheetName) {
const result = await sheets.spreadsheets.values.get({ spreadsheetId: SHEETS_ID, range: sheetName + “!A:Z” });
return result.data.values || [];
}

async function appendRow(sheets, sheetName, row) {
await sheets.spreadsheets.values.append({
spreadsheetId: SHEETS_ID,
range: sheetName + “!A:Z”,
valueInputOption: “RAW”,
requestBody: { values: [row] },
});
}

async function findLastUserRow(sheets, sheetName, userId) {
const data = await getSheetData(sheets, sheetName);
for (let i = data.length - 1; i >= 1; i–) {
const row = data[i];
if (String(row[7] || row[9] || “”) === userId) return { row: row, index: i };
}
return { row: null, index: -1 };
}

function generateId(prefix) {
const ts = new Date().toISOString().replace(/[:-T.Z]/g, “”).substring(0, 14);
const rnd = Math.floor(Math.random() * 1000).toString().padStart(3, “0”);
return prefix + ts + rnd;
}

function formatDateTime(date) {
return new Date(date).toLocaleString(“zh-TW”, { timeZone: “Asia/Taipei”, year: “numeric”, month: “2-digit”, day: “2-digit”, hour: “2-digit”, minute: “2-digit” });
}

function formatDate(date) {
const d = new Date(date);
return d.getFullYear() + “-” + String(d.getMonth() + 1).padStart(2, “0”) + “-” + String(d.getDate()).padStart(2, “0”);
}

function formatTime(date) {
const d = new Date(date);
return String(d.getHours()).padStart(2, “0”) + “:” + String(d.getMinutes()).padStart(2, “0”);
}

function generateFileName(type, messageId) {
const ext = { image: “jpg”, video: “mp4”, audio: “m4a”, file: “bin” };
const now = new Date().toISOString().substring(0, 10).replace(/-/g, “”);
return type + “*” + now + “*” + messageId + “.” + (ext[type] || “bin”);
}

function getMimeType(type, fileName) {
const ext = path.extname(fileName || “”).toLowerCase();
const map = {
“.jpg”: “image/jpeg”, “.jpeg”: “image/jpeg”, “.png”: “image/png”, “.gif”: “image/gif”, “.webp”: “image/webp”,
“.pdf”: “application/pdf”, “.doc”: “application/msword”,
“.docx”: “application/vnd.openxmlformats-officedocument.wordprocessingml.document”,
“.xls”: “application/vnd.ms-excel”, “.xlsx”: “application/vnd.openxmlformats-officedocument.spreadsheetml.sheet”,
“.mp4”: “video/mp4”, “.mov”: “video/quicktime”, “.m4a”: “audio/m4a”, “.mp3”: “audio/mpeg”,
“.wav”: “audio/wav”, “.txt”: “text/plain”, “.zip”: “application/zip”,
};
if (map[ext]) return map[ext];
if (type === “image”) return “image/jpeg”;
if (type === “video”) return “video/mp4”;
if (type === “audio”) return “audio/m4a”;
return “application/octet-stream”;
}

function classifyFileType(type, fileName) {
const ext = path.extname(fileName || “”).toLowerCase();
if (type === “image”) return “image”;
if (type === “video”) return “video”;
if (type === “audio”) return “audio”;
if (ext === “.pdf”) return “pdf”;
if ([”.xls”, “.xlsx”, “.csv”].includes(ext)) return “excel”;
if ([”.doc”, “.docx”].includes(ext)) return “word”;
return “other”;
}

function makeRow(label, value) {
return {
type: “box”, layout: “horizontal”,
contents: [
{ type: “text”, text: label, size: “sm”, color: “#888888”, flex: 1 },
{ type: “text”, text: String(value || “”), size: “sm”, color: “#333333”, flex: 3, wrap: true },
],
};
}

function replyText(replyToken, text) { return replyMessages(replyToken, [{ type: “text”, text: text }]); }
function replyFlex(replyToken, flexMessage) { return replyMessages(replyToken, [flexMessage]); }
function replyMessages(replyToken, messages) { return lineApiRequest(”/v2/bot/message/reply”, { replyToken: replyToken, messages: messages }); }

function lineApiRequest(apiPath, payload) {
return new Promise(function(resolve, reject) {
const body = JSON.stringify(payload);
const req = https.request({
hostname: “api.line.me”,
path: apiPath,
method: “POST”,
headers: { Authorization: “Bearer “ + LINE_TOKEN, “Content-Type”: “application/json”, “Content-Length”: Buffer.byteLength(body) },
}, function(res) {
let data = “”;
res.on(“data”, function(chunk) { data += chunk; });
res.on(“end”, function() {
if (res.statusCode >= 200 && res.statusCode < 300) resolve(data ? JSON.parse(data) : {});
else reject(new Error(“LINE API error “ + res.statusCode + “: “ + data));
});
});
req.on(“error”, reject);
req.write(body);
req.end();
});
}
