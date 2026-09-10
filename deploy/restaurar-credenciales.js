// Restaura ~/.clasprc.json a partir del secret CLASP_CREDENTIALS.
// Soporta formato V3 (tokens.default), V1 local (token + oauth2ClientSettings),
// V1 con solo token, y V1 global (access_token suelto). Usado por ambos workflows.
const fs = require("fs");
const os = require("os");
const path = require("path");

const rawSecret = process.env.CLASP_CREDENTIALS;
if (!rawSecret) {
  console.error("Falta la variable CLASP_CREDENTIALS");
  process.exit(1);
}

let raw;
try {
  raw = JSON.parse(rawSecret);
} catch (e) {
  console.error("CLASP_CREDENTIALS no es JSON válido: " + e.message);
  process.exit(1);
}

const DEFAULT_CLIENT_ID =
  "1072944905499-vm2v2i5dvn0a0d2o4ca36i1vge8cvbn0.apps.googleusercontent.com";
const DEFAULT_CLIENT_SECRET = "v6V3fKV_zWU7iw1DrpO1rknX";

let output;
if (raw.tokens && raw.tokens.default) {
  output = raw;
  console.log("Formato V3 detectado (tokens.default)");
} else if (raw.token && raw.oauth2ClientSettings) {
  output = raw;
  console.log("Formato V1 local detectado");
} else if (raw.token) {
  output = {
    token: raw.token,
    oauth2ClientSettings: {
      clientId: raw.token.client_id || DEFAULT_CLIENT_ID,
      clientSecret: raw.token.client_secret || DEFAULT_CLIENT_SECRET,
    },
  };
  console.log("Formato V1 sin oauth2ClientSettings, completado");
} else if (raw.access_token) {
  output = {
    tokens: {
      default: {
        access_token: raw.access_token,
        refresh_token: raw.refresh_token,
        token_type: raw.token_type || "Bearer",
        client_id: raw.client_id || DEFAULT_CLIENT_ID,
        client_secret: raw.client_secret || DEFAULT_CLIENT_SECRET,
      },
    },
  };
  console.log("Formato V1 global, convertido a V3");
} else {
  output = raw;
  console.log("Formato desconocido, usando tal cual");
}

const dest = path.join(os.homedir(), ".clasprc.json");
fs.writeFileSync(dest, JSON.stringify(output, null, 2));
console.log("Credenciales escritas en: " + dest);
