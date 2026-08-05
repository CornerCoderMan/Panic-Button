"use strict";

/* ============================================================
   Report This! — local development static host
   ============================================================
   Serves the add-in's static files over HTTP/HTTPS for local
   sideloading. The add-in runs entirely client-side (Office.js),
   so there is no backend API — this server only hosts files.
   ============================================================ */

const fs = require("node:fs");
const http = require("node:http");
const https = require("node:https");
const os = require("node:os");
const path = require("node:path");

const ROOT = __dirname;
const PORT = Number(process.env.PORT || 3000);
const USE_HTTPS = process.argv.includes("--https");

const PUBLIC_FILES = new Map([
  ["/", "taskpane.html"],
  ["/taskpane.html", "taskpane.html"],
  ["/taskpane.js", "taskpane.js"],
  ["/taskpane.css", "taskpane.css"],
  ["/commands.html", "commands.html"],
  ["/support.html", "support.html"],
  ["/assets/icon-16.png", "assets/icon-16.png"],
  ["/assets/icon-32.png", "assets/icon-32.png"],
  ["/assets/icon-80.png", "assets/icon-80.png"],
]);

const CONTENT_TYPES = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".png": "image/png",
};

function securityHeaders(contentType) {
  return {
    "Cache-Control": "no-store",
    "Content-Type": contentType,
    "Referrer-Policy": "no-referrer",
    "X-Content-Type-Options": "nosniff",
  };
}

function contentTypeFor(filePath) {
  return CONTENT_TYPES[path.extname(filePath)] || "application/octet-stream";
}

function sendJson(response, status, payload) {
  response.writeHead(status, securityHeaders("application/json; charset=utf-8"));
  response.end(JSON.stringify(payload));
}

function serveFile(requestPath, response) {
  const relativePath = PUBLIC_FILES.get(requestPath);
  if (!relativePath) {
    sendJson(response, 404, { error: "Not found." });
    return;
  }

  const filePath = path.join(ROOT, relativePath);
  fs.readFile(filePath, (error, content) => {
    if (error) {
      sendJson(response, 404, { error: "Not found." });
      return;
    }
    response.writeHead(200, securityHeaders(contentTypeFor(filePath)));
    response.end(content);
  });
}

function handleRequest(request, response) {
  const requestUrl = new URL(request.url, `http://${request.headers.host || "localhost"}`);

  if ((request.method === "GET" || request.method === "HEAD") && PUBLIC_FILES.has(requestUrl.pathname)) {
    serveFile(requestUrl.pathname, response);
    return;
  }

  sendJson(response, 404, { error: "Not found." });
}

function createServer() {
  if (!USE_HTTPS) return http.createServer(handleRequest);

  const certificateDirectory = path.join(os.homedir(), ".office-addin-dev-certs");
  const certificatePath = process.env.SSL_CERT_FILE || path.join(certificateDirectory, "localhost.crt");
  const keyPath = process.env.SSL_KEY_FILE || path.join(certificateDirectory, "localhost.key");
  return https.createServer({
    cert: fs.readFileSync(certificatePath),
    key: fs.readFileSync(keyPath),
  }, handleRequest);
}

if (require.main === module) {
  createServer().listen(PORT, () => {
    const protocol = USE_HTTPS ? "https" : "http";
    console.log(`Report This! dev host running at ${protocol}://localhost:${PORT}`);
  });
}

module.exports = { PUBLIC_FILES, CONTENT_TYPES, contentTypeFor };
