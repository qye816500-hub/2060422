// api/og.js
// GET /api/og?url=https://...
// Fetches OG image from a URL server-side to avoid CORS issues

const https = require("https");
const http = require("http");

function fetchHtml(targetUrl) {
  return new Promise(function(resolve, reject) {
    const mod = targetUrl.startsWith("https") ? https : http;
    const req = mod.get(targetUrl, {
      headers: {
        "User-Agent": "Mozilla/5.0 (compatible; EEBot/1.0)",
        "Accept": "text/html",
      },
      timeout: 5000,
    }, function(res) {
      // Follow redirects
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        return fetchHtml(res.headers.location).then(resolve).catch(reject);
      }
      let body = "";
      res.setEncoding("utf8");
      res.on("data", function(chunk) {
        body += chunk;
        if (body.length > 50000) { req.destroy(); resolve(body); } // limit
      });
      res.on("end", function() { resolve(body); });
    });
    req.on("error", reject);
    req.on("timeout", function() { req.destroy(); reject(new Error("Timeout")); });
  });
}

module.exports = async function(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  if (req.method === "OPTIONS") return res.status(200).end();

  const targetUrl = req.query && req.query.url;
  if (!targetUrl) return res.status(400).json({ error: "Missing url" });

  try {
    const html = await fetchHtml(decodeURIComponent(targetUrl));

    // Extract OG image
    const ogMatch = html.match(/<meta[^>]*property=["']og:image["'][^>]*content=["']([^"']+)["']/i)
                 || html.match(/<meta[^>]*content=["']([^"']+)["'][^>]*property=["']og:image["']/i);
    // Extract OG title
    const titleMatch = html.match(/<meta[^>]*property=["']og:title["'][^>]*content=["']([^"']+)["']/i)
                    || html.match(/<meta[^>]*content=["']([^"']+)["'][^>]*property=["']og:title["']/i)
                    || html.match(/<title[^>]*>([^<]+)<\/title>/i);

    return res.status(200).json({
      success: true,
      image: ogMatch ? ogMatch[1] : null,
      title: titleMatch ? titleMatch[1].trim() : null,
    });
  } catch (e) {
    console.error("OG API error:", e.message);
    return res.status(200).json({ success: false, image: null, title: null });
  }
};
