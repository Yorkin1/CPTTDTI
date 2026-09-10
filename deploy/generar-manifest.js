// Genera src/appsscript.json (ignorado por git, solo existe en el runner para `clasp push`).
const fs = require("fs");
const path = require("path");

const manifest = {
  timeZone: "America/Santo_Domingo",
  dependencies: {},
  exceptionLogging: "STACKDRIVER",
  runtimeVersion: "V8",
  oauthScopes: [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/script.send_mail",
    "https://www.googleapis.com/auth/script.container.ui",
    "https://www.googleapis.com/auth/userinfo.email",
  ],
  webapp: {
    executeAs: "USER_DEPLOYING",
    access: "ANYONE",
  },
  executionApi: {
    access: "MYSELF",
  },
};

const dest = path.join(__dirname, "..", "src", "appsscript.json");
fs.writeFileSync(dest, JSON.stringify(manifest, null, 2) + "\n");
console.log("Manifiesto escrito en: " + dest);
