const fs = require("fs");
const path = require("path");

const FILE_PATH = path.join(__dirname, "tokenStore.json");

function saveTokens(tokens) {
  fs.writeFileSync(FILE_PATH, JSON.stringify(tokens, null, 2));
}

function loadTokens() {
  if (!fs.existsSync(FILE_PATH)) return null;
  const data = fs.readFileSync(FILE_PATH);
  return JSON.parse(data);
}

module.exports = { saveTokens, loadTokens };
