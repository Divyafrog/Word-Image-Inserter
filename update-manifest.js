// update-manifest.js
// Updates manifest.xml with your ngrok URL automatically.
//
// Usage:
//   node update-manifest.js https://xxxx.ngrok-free.app

const fs   = require("fs");
const path = require("path");

const newUrl = process.argv[2];

if (!newUrl) {
  console.error("Usage: node update-manifest.js https://xxxx.ngrok-free.app");
  process.exit(1);
}

if (!newUrl.startsWith("https://")) {
  console.error("URL must start with https://");
  process.exit(1);
}

const manifestPath = path.join(__dirname, "manifest.xml");
let manifest = fs.readFileSync(manifestPath, "utf8");

// Replace all occurrences of localhost:3000 with the ngrok URL
manifest = manifest.replace(/https:\/\/localhost:3000/g, newUrl);

fs.writeFileSync(manifestPath, manifest);
console.log(`✅ manifest.xml updated with: ${newUrl}`);
console.log(`\nNow reload the add-in in Word:`);
console.log(`  Insert → My Add-ins → Shared Folder → Image Inserter`);
console.log(`  (If it was already loaded, close Word and reopen it)`);