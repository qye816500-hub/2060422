// api/og.js
// GET /api/og?url=https://...
// Fetches OG image from a URL server-side to avoid CORS issues

const https = require("https");
const http = require("http");

function fetchJson(url) {
  return new Promise(function(resolve, reject) {
    https.get(url, {
      headers: { "User-Agent": "Mozilla/5.0 (compatible; EEBot/1.0)" },
      timeout: 6000,
    }, function(res) {
      let body = "";
      res.setEncoding("utf8");
      res.on("data", function(c) { body += c; });
      res.on("end", function() {
        try { resolve(JSON.parse(body)); } catch(e) { reject(e); }
      });
    }).on("error", reject).on("timeout", function(r) { r.destroy(); reject(new Error("Timeout")); });
  });
}

function fetchHtml(targetUrl, redirects) {
  redirects = redirects || 0;
  if (redirects > 4) return Promise.reject(new Error("Too many redirects"));
  return new Promise(function(resolve, reject) {
    const mod = targetUrl.startsWith("https") ? https : http;
    const req = mod.get(targetUrl, {
      headers: {
        "User-Agent": "facebookexternalhit/1.1 (+http://www.facebook.com/externalhit_uatext.php)",
        "Accept": "text/html,application/xhtml+xml",
      },
      timeout: 6000,
    }, function(res) {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        return fetchHtml(res.headers.location, redirects + 1).then(resolve).catch(reject);
      }
      let body = "";
      res.setEncoding("utf8");
      res.on("data", function(chunk) {
        body += chunk;
        if (body.length > 80000) { req.destroy(); resolve(body); }
      });
      res.on("end", function() { resolve(body); });
    });
    req.on("error", reject);
    req.on("timeout", function() { req.destroy(); reject(new Error("Timeout")); });
  });
}

function extractOg(html) {
  const ogImage = html.match(/<meta[^>]*property=["']og:image["'][^>]*content=["']([^"']+)["']/i)
               || html.match(/<meta[^>]*content=["']([^"']+)["'][^>]*property=["']og:image["']/i);
  const ogTitle = html.match(/<meta[^>]*property=["']og:title["'][^>]*content=["']([^"']+)["']/i)
               || html.match(/<meta[^>]*content=["']([^"']+)["'][^>]*property=["']og:title["']/i)
               || html.match(/<title[^>]*>([^<]+)<\/title>/i);
  return {
    image: ogImage ? ogImage[1] : null,
    title: ogTitle ? ogTitle[1].trim() : null,
  };
}

module.exports = async function(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  if (req.method === "OPTIONS") return res.status(200).end();

  const targetUrl = req.query && req.query.url;
  if (!targetUrl) return res.status(400).json({ error: "Missing url" });

  const decoded = decodeURIComponent(targetUrl);

  try {
    // ── Threads: use oEmbed API ──────────────────────────
    if (decoded.includes("threads.net") || decoded.includes("threads.com")) {
      try {
        const oembedUrl = "https://www.threads.net/oembed/?url=" + encodeURIComponent(decoded);
        const data = await fetchJson(oembedUrl);
        // oEmbed gives thumbnail_url and title
        return res.status(200).json({
          success: true,
          image: data.thumbnail_url || null,
          title: data.title || data.author_name || null,
        });
      } catch(e) {
        // oEmbed failed, fall through to HTML fetch
      }
    }

    // ── Instagram: use oEmbed API ────────────────────────
    if (decoded.includes("instagram.com")) {
      try {
        const oembedUrl = "https://graph.facebook.com/v18.0/instagram_oembed?url=" + encodeURIComponent(decoded) + "&maxwidth=400";
        const data = await fetchJson(oembedUrl);
        if (data.thumbnail_url) {
          return res.status(200).json({
            success: true,
            image: data.thumbnail_url,
            title: data.title || null,
          });
        }
      } catch(e) {
        // oEmbed failed, fall through
      }
    }

    // ── General: fetch HTML and extract OG tags ──────────
    const html = await fetchHtml(decoded);
    const og = extractOg(html);

    return res.status(200).json({ success: true, ...og });

  } catch (e) {
    console.error("OG API error:", e.message);
    return res.status(200).json({ success: false, image: null, title: null });
  }
};
