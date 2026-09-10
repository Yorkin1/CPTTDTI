// Despliega src/ a cada cliente activo (push + versión + redeploy vía API). Cada cliente falla aislado.
const { execSync } = require("child_process");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { google } = require("googleapis");

const DEFAULT_CLIENT_ID =
  "1072944905499-vm2v2i5dvn0a0d2o4ca36i1vge8cvbn0.apps.googleusercontent.com";
const DEFAULT_CLIENT_SECRET = "v6V3fKV_zWU7iw1DrpO1rknX";

function sanitizarDescripcion(texto) {
  return String(texto || "")
    .replace(/[\r\n]+/g, " ")
    .replace(/"/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 120);
}

function resolverCredenciales(clasprc) {
  const tokenObj =
    (clasprc.tokens && clasprc.tokens.default) || clasprc.token || clasprc;
  const oauthSettings = clasprc.oauth2ClientSettings || {};
  return {
    clientId:
      tokenObj.client_id ||
      tokenObj.clientId ||
      oauthSettings.clientId ||
      oauthSettings.client_id ||
      DEFAULT_CLIENT_ID,
    clientSecret:
      tokenObj.client_secret ||
      tokenObj.clientSecret ||
      oauthSettings.clientSecret ||
      oauthSettings.client_secret ||
      DEFAULT_CLIENT_SECRET,
    accessToken: tokenObj.access_token || tokenObj.accessToken,
    refreshToken: tokenObj.refresh_token || tokenObj.refreshToken,
  };
}

(async () => {
  const repoRoot = path.join(__dirname, "..");
  const clientesPath = path.join(repoRoot, "deploy", "clientes.json");
  const clientes = require(clientesPath);

  const sha = (process.env.GITHUB_SHA || "").slice(0, 7);
  let commitMsg = "actualizar";
  try {
    commitMsg = execSync(
      "git log -1 --format=%s " + (process.env.GITHUB_SHA || "HEAD"),
      { encoding: "utf8" }
    ).trim();
  } catch (e) {
    console.log("No se pudo leer el mensaje del commit, usando genérico.");
  }
  const fecha = new Date().toISOString().slice(0, 10);
  const desc = sanitizarDescripcion(fecha + ": " + commitMsg);

  const fallos = [];
  const exitosos = [];
  let clientesModificados = false;

  const clasprcPath = path.join(os.homedir(), ".clasprc.json");
  const clasprc = JSON.parse(fs.readFileSync(clasprcPath, "utf8"));
  const creds = resolverCredenciales(clasprc);
  if (!creds.accessToken || !creds.refreshToken) {
    console.error("Credenciales incompletas en " + clasprcPath);
    process.exit(1);
  }
  const oauth2 = new google.auth.OAuth2(
    creds.clientId,
    creds.clientSecret,
    "https://developers.google.com/oauthplayground"
  );
  oauth2.setCredentials({
    access_token: creds.accessToken,
    refresh_token: creds.refreshToken,
  });
  const scriptApi = google.script({ version: "v1", auth: oauth2 });

  for (const c of clientes) {
    if (!c.activo) {
      console.log("⊘ " + c.slug + " (inactivo, saltando)");
      continue;
    }
    console.log("→ " + c.slug);
    fs.writeFileSync(
      path.join(repoRoot, ".clasp.json"),
      JSON.stringify(
        {
          scriptId: c.scriptId,
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

    // 1) Subir código. Si falla, no se intenta versionar.
    try {
      execSync("clasp push -f", { stdio: "inherit", cwd: repoRoot });
    } catch (pushErr) {
      console.error("✗ Falló " + c.slug + " en push: " + pushErr.message);
      fallos.push(c.slug);
      continue;
    }

    // 2) Versionar y (re)desplegar vía API, con catch propio por cliente.
    //    Este bloque tenía el bug: `try { ... }` sin `catch` (SyntaxError en [eval]:108).
    try {
      const versionRes = await scriptApi.projects.versions.create({
        scriptId: c.scriptId,
        requestBody: { description: desc },
      });
      const versionNumber = versionRes.data.versionNumber;
      console.log("  Versión " + versionNumber + " creada");

      const deployRes = await scriptApi.projects.deployments.list({
        scriptId: c.scriptId,
      });
      const existing = (deployRes.data.deployments || []).find(function (d) {
        return d.deploymentId === c.deploymentId;
      });

      let deployOk = false;

      if (existing) {
        try {
          await scriptApi.projects.deployments.update({
            scriptId: c.scriptId,
            deploymentId: c.deploymentId,
            requestBody: {
              deploymentConfig: {
                versionNumber: versionNumber,
                manifestFileName: "appsscript",
                description: desc,
              },
            },
          });
          console.log(
            "  Implementación existente activada con versión " + versionNumber
          );
          deployOk = true;
        } catch (updateErr) {
          console.log(
            "  Implementación existente archivada/no editable; creando una nueva activa..."
          );
        }
      }

      if (!deployOk) {
        const newDeploy = await scriptApi.projects.deployments.create({
          scriptId: c.scriptId,
          requestBody: {
            versionNumber: versionNumber,
            manifestFileName: "appsscript",
            description: desc,
          },
        });
        const newId = newDeploy.data.deploymentId;
        const webApp = (newDeploy.data.entryPoints || []).some(function (
          entry
        ) {
          return !!entry.webApp;
        });
        if (!newId || !webApp) {
          throw new Error(
            "La nueva implementación no quedó activa como aplicación web"
          );
        }
        c.deploymentId = newId;
        clientesModificados = true;
        console.log("  Nueva implementación web activa: " + newId);
        deployOk = true;
      }

      if (deployOk) {
        console.log("✓ " + c.slug + " actualizado");
        exitosos.push(c.slug);
      } else {
        throw new Error("No se pudo activar ninguna implementación");
      }
    } catch (apiErr) {
      console.error(
        "✗ Falló " + c.slug + " en versionado/despliegue: " + apiErr.message
      );
      fallos.push(c.slug);
    }
  }

  if (clientesModificados) {
    fs.writeFileSync(
      clientesPath,
      JSON.stringify(clientes, null, 2) + "\n"
    );
    console.log("\nArchivos de clientes actualizados.");
  }

  const ahora = new Date();
  const fechaLegible = ahora.toLocaleString("es-DO", {
    timeZone: "America/Santo_Domingo",
  });
  let md = "# Reporte de Clientes\n\n";
  md += "**Fecha:** " + fechaLegible + "  \n";
  md += "**Commit:** " + sha + "\n\n";
  md += "## Resumen\n\n";
  md += "- **Total:** " + clientes.length + "\n";
  md += "- **Exitosos:** " + exitosos.length + "\n";
  md += "- **Fallidos:** " + fallos.length + "\n\n";
  md += "## Estado\n\n";
  md += "| Cliente | Estado | URL |\n";
  md += "|---------|--------|-----|\n";
  for (const c of clientes) {
    const ok = exitosos.includes(c.slug);
    const fallo = fallos.includes(c.slug);
    const estado = fallo ? "error" : ok ? "ok" : "saltado";
    const url =
      "https://script.google.com/macros/s/" + c.deploymentId + "/exec";
    md += "| " + c.slug + " | " + estado + " | " + url + " |\n";
  }
  md += "\n---\n\n";
  md += "## Detalle\n\n";
  for (const c of clientes) {
    md += "### " + c.slug + "\n";
    md += "- **Script ID:** " + c.scriptId + "\n";
    md += "- **Deployment ID:** " + c.deploymentId + "\n";
    md += "- **Spreadsheet ID:** " + c.spreadsheetId + "\n";
    md +=
      "- **URL:** https://script.google.com/macros/s/" +
      c.deploymentId +
      "/exec\n\n";
  }
  fs.writeFileSync(path.join(repoRoot, "deploy", "REPORT.md"), md);
  console.log("\nReporte: deploy/REPORT.md");

  console.log("");
  console.log(
    "Resumen: " + exitosos.length + " exitosos, " + fallos.length + " fallidos"
  );
  if (fallos.length) {
    console.error("Fallaron: " + fallos.join(", "));
    process.exit(1);
  }
})().catch(function (err) {
  console.error("Error fatal del actualizador: " + err.message);
  process.exit(1);
});
