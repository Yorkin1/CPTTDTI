// Detecta spreadsheetId/deploymentId de un cliente vía Apps Script API.
// Entradas por env (evita interpolar ${{ inputs }} dentro de JS): INPUT_SLUG, INPUT_SCRIPT_ID.
const { execSync } = require("child_process");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { google } = require("googleapis");

const DEFAULT_CLIENT_ID =
  "1072944905499-vm2v2i5dvn0a0d2o4ca36i1vge8cvbn0.apps.googleusercontent.com";
const DEFAULT_CLIENT_SECRET = "v6V3fKV_zWU7iw1DrpO1rknX";

(async () => {
  const slug = (process.env.INPUT_SLUG || "").trim();
  let scriptIdInput = (process.env.INPUT_SCRIPT_ID || "").trim();
  if (!slug || !scriptIdInput) {
    console.error("Faltan INPUT_SLUG o INPUT_SCRIPT_ID");
    process.exit(1);
  }

  if (scriptIdInput.includes("script.google.com")) {
    const matchMacro = scriptIdInput.match(/\/macros\/s\/([a-zA-Z0-9_-]+)/);
    const matchProject = scriptIdInput.match(
      /\/projects\/([a-zA-Z0-9_-]+)/
    );
    if (matchMacro) {
      scriptIdInput = matchMacro[1];
    } else if (matchProject) {
      scriptIdInput = matchProject[1];
    } else {
      console.error("No se pudo extraer el Script ID de la URL del editor");
      process.exit(1);
    }
  } else if (
    scriptIdInput.includes("docs.google.com") ||
    scriptIdInput.includes("spreadsheets")
  ) {
    console.error("ERROR: Pusiste la URL del spreadsheet, no del Apps Script.");
    console.error(
      "Necesitas ir a Extensiones > Apps Script en la hoja, y copiar el ID de la URL del editor."
    );
    process.exit(1);
  }

  const scriptId = scriptIdInput;
  console.log("Script ID: " + scriptId);

  const rawSecret = process.env.CLASP_CREDENTIALS || "";
  const rawPath = "/tmp/raw-clasprc.json";
  fs.writeFileSync(
    rawPath,
    rawSecret || fs.readFileSync(path.join(os.homedir(), ".clasprc.json"), "utf8")
  );
  const clasprc = JSON.parse(fs.readFileSync(rawPath, "utf8"));
  const tokenObj =
    (clasprc.tokens && clasprc.tokens.default) || clasprc.token || clasprc;
  const oauthSettings = clasprc.oauth2ClientSettings || {};
  const clientId =
    tokenObj.client_id ||
    tokenObj.clientId ||
    oauthSettings.clientId ||
    oauthSettings.client_id ||
    DEFAULT_CLIENT_ID;
  const clientSecret =
    tokenObj.client_secret ||
    tokenObj.clientSecret ||
    oauthSettings.clientSecret ||
    oauthSettings.client_secret ||
    DEFAULT_CLIENT_SECRET;
  const accessToken = tokenObj.access_token;
  const refreshToken = tokenObj.refresh_token;

  if (!accessToken || !refreshToken || !clientId || !clientSecret) {
    console.error("Faltan campos en las credenciales");
    console.error("Estructura:", JSON.stringify(Object.keys(clasprc)));
    process.exit(1);
  }

  const oauth2 = new google.auth.OAuth2(
    clientId,
    clientSecret,
    "https://developers.google.com/oauthplayground"
  );
  oauth2.setCredentials({
    access_token: accessToken,
    refresh_token: refreshToken,
  });

  const script = google.script({ version: "v1", auth: oauth2 });

  const projectRes = await script.projects.get({ scriptId });
  const spreadsheetId = projectRes.data.parentId || "";
  console.log("Spreadsheet ID: " + spreadsheetId);

  const deployRes = await script.projects.deployments.list({ scriptId });
  const deployments = deployRes.data.deployments || [];

  const versionedDeployments = deployments.filter(
    (d) =>
      d.deploymentConfig &&
      d.deploymentConfig.versionNumber &&
      parseInt(d.deploymentConfig.versionNumber, 10) > 0
  );
  let deploymentId;

  if (versionedDeployments.length > 0) {
    deploymentId = versionedDeployments[0].deploymentId;
    console.log("Usando implementación versionada: " + deploymentId);
  } else {
    console.log("No hay implementaciones versionadas. Creando una nueva...");
    const repoRoot = path.join(__dirname, "..");
    fs.writeFileSync(
      path.join(repoRoot, ".clasp.json"),
      JSON.stringify(
        {
          scriptId: scriptId,
          rootDir: "./src",
          scriptExtensions: [".js", ".gs"],
          htmlExtensions: [".html"],
          jsonExtensions: [".json"],
          filePushOrder: [],
          skipSubdirectories: false,
        },
        null,
        2
      )
    );
    const output = execSync('clasp deploy -d "deploy inicial"', {
      encoding: "utf8",
      stdio: "pipe",
      cwd: repoRoot,
    });
    console.log(output);
    const match = output.match(
      /https:\/\/script\.google\.com\/macros\/s\/([a-zA-Z0-9_-]+)\/exec/
    );
    if (match && match[1]) {
      deploymentId = match[1];
      console.log("Nueva implementación creada: " + deploymentId);
    } else {
      console.error("No se pudo extraer el deploymentId");
      process.exit(1);
    }
  }

  console.log("Deployment ID: " + deploymentId);

  fs.writeFileSync(
    "/tmp/cliente-detectado.json",
    JSON.stringify(
      { slug, scriptId, deploymentId, spreadsheetId },
      null,
      2
    )
  );
  console.log("Cliente detectado correctamente");
})().catch((err) => {
  console.error("Error detectando cliente: " + err.message);
  process.exit(1);
});
