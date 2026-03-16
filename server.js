// server.js — local HTTP server for the Word Add-in
// Works with ngrok which provides the HTTPS layer Word requires.
//
// Usage:
//   1. node server.js
//   2. In a second terminal: ngrok http 3000
//   3. Copy the https://xxxx.ngrok-free.app URL ngrok gives you
//   4. Run: node update-manifest.js https://xxxx.ngrok-free.app
//   5. Load the manifest in Word (see README.md)

const http = require("http");
const fs   = require("fs");
const path = require("path");

const PORT = 3000;

const MIME = {
  ".html": "text/html",
  ".js":   "application/javascript",
  ".css":  "text/css",
  ".png":  "image/png",
  ".jpg":  "image/jpeg",
  ".xml":  "application/xml",
  ".json": "application/json",
};

const server = http.createServer((req, res) => {
  let filePath = path.join(__dirname, req.url === "/" ? "/taskpane.html" : req.url);

  // Serve a simple 1x1 pixel PNG as icon if icon.png is missing
  if (req.url === "/icon.png" && !fs.existsSync(filePath)) {
    const pixel = Buffer.from(
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==",
      "base64"
    );
    res.writeHead(200, { "Content-Type": "image/png" });
    return res.end(pixel);
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      return res.end("Not found");
    }
    const ext = path.extname(filePath);
    res.writeHead(200, {
      "Content-Type": MIME[ext] || "text/plain",
      "Access-Control-Allow-Origin": "*",
      "ngrok-skip-browser-warning": "true",
    });
    res.end(data);
  });
});

server.listen(PORT, () => {
  console.log(`✅ Server running at http://localhost:${PORT}`);
  console.log(`\nNext step: open a second terminal and run:`);
  console.log(`   ngrok http 3000`);
  console.log(`\nThen copy the https://xxxx.ngrok-free.app URL and run:`);
  console.log(`   node update-manifest.js https://xxxx.ngrok-free.app`);
});