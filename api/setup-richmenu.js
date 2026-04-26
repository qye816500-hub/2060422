// api/setup-richmenu.js
// 打開這個網址一次就會自動設定 Rich Menu
// https://2060422.vercel.app/api/setup-richmenu

const https = require("https");
const fs = require("fs");
const path = require("path");

const LINE_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN;
const LIFF_URL = "https://liff.line.me/2009891497-Bd5P0goB";

function lineRequest(method, apiPath, data, contentType) {
  return new Promise(function(resolve, reject) {
    const isBuffer = Buffer.isBuffer(data);
    const body = isBuffer ? data : (data ? JSON.stringify(data) : null);
    const headers = {
      Authorization: "Bearer " + LINE_TOKEN,
      "Content-Type": contentType || "application/json",
    };
    if (body) headers["Content-Length"] = body.length;

    const req = https.request({
      hostname: "api.line.me",
      path: apiPath,
      method: method,
      headers: headers,
    }, function(res) {
      const chunks = [];
      res.on("data", function(c) { chunks.push(c); });
      res.on("end", function() {
        const text = Buffer.concat(chunks).toString();
        try { resolve({ status: res.statusCode, body: JSON.parse(text) }); }
        catch(e) { resolve({ status: res.statusCode, body: text }); }
      });
    });
    req.on("error", reject);
    if (body) req.write(body);
    req.end();
  });
}

module.exports = async function(req, res) {
  if (req.method !== "GET") {
    return res.status(405).json({ error: "GET only" });
  }

  const log = [];

  try {
    // Step 1: Delete existing rich menus
    log.push("Step 1: Deleting existing rich menus...");
    const listRes = await lineRequest("GET", "/v2/bot/richmenu/list");
    if (listRes.body && listRes.body.richmenus) {
      for (const rm of listRes.body.richmenus) {
        await lineRequest("DELETE", "/v2/bot/richmenu/" + rm.richMenuId);
        log.push("Deleted: " + rm.richMenuId);
      }
    }

    // Step 2: Create rich menu
    log.push("Step 2: Creating rich menu...");
    const richMenuBody = {
      size: { width: 2500, height: 843 },
      selected: true,
      name: "EE\u5C0F\u52A9\u7406\u9078\u55AE",
      chatBarText: "\u9078\u55AE",
      areas: [
        {
          bounds: { x: 0, y: 0, width: 625, height: 843 },
          action: { type: "message", text: "\u5F85\u8FA6\u6E05\u55AE" }
        },
        {
          bounds: { x: 625, y: 0, width: 625, height: 843 },
          action: { type: "message", text: "\u6700\u8FD1\u6536\u85CF" }
        },
        {
          bounds: { x: 1250, y: 0, width: 625, height: 843 },
          action: {
            type: "uri",
            uri: LIFF_URL + "?tab=calendar"
          }
        },
        {
          bounds: { x: 1875, y: 0, width: 625, height: 843 },
          action: {
            type: "uri",
            uri: LIFF_URL
          }
        }
      ]
    };

    const createRes = await lineRequest("POST", "/v2/bot/richmenu", richMenuBody);
    log.push("Created rich menu: " + JSON.stringify(createRes.body));

    if (!createRes.body || !createRes.body.richMenuId) {
      return res.status(500).json({ error: "Failed to create rich menu", log: log, detail: createRes.body });
    }

    const richMenuId = createRes.body.richMenuId;

    // Step 3: Upload image from GitHub raw
    log.push("Step 3: Uploading image...");

    const imgBuffer = await new Promise(function(resolve, reject) {
      https.get("https://raw.githubusercontent.com/qye816500-hub/2060422/main/richmenu.png", function(imgRes) {
        const chunks = [];
        imgRes.on("data", function(c) { chunks.push(c); });
        imgRes.on("end", function() { resolve(Buffer.concat(chunks)); });
        imgRes.on("error", reject);
      }).on("error", reject);
    });

    log.push("Image downloaded, size: " + imgBuffer.length + " bytes");

    const uploadRes = await lineRequest(
      "POST",
      "/v2/bot/richmenu/" + richMenuId + "/content",
      imgBuffer,
      "image/png"
    );
    log.push("Upload result: " + JSON.stringify(uploadRes.body) + " status: " + uploadRes.status);

    // Step 4: Set as default
    log.push("Step 4: Setting as default...");
    const defaultRes = await lineRequest("POST", "/v2/bot/richmenu/default/" + richMenuId);
    log.push("Set default result: " + JSON.stringify(defaultRes.body));

    return res.status(200).json({
      success: true,
      richMenuId: richMenuId,
      message: "Rich Menu \u8A2D\u5B9A\u5B8C\u6210\uFF01\u53BB LINE \u6E2C\u8A66\u770B\u770B\u5427\uFF01",
      log: log
    });

  } catch(e) {
    log.push("Error: " + e.message);
    return res.status(500).json({ error: e.message, log: log });
  }
};
