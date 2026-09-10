// Agrega el cliente detectado (/tmp/cliente-detectado.json) a deploy/clientes.json.
const fs = require("fs");
const path = require("path");

const detectadoPath = "/tmp/cliente-detectado.json";
const clientesPath = path.join(__dirname, "clientes.json");

const cliente = JSON.parse(fs.readFileSync(detectadoPath, "utf8"));

let clientes = [];
if (fs.existsSync(clientesPath)) {
  clientes = JSON.parse(fs.readFileSync(clientesPath, "utf8"));
}

if (clientes.some((c) => c.slug === cliente.slug)) {
  console.error("Ya existe un cliente con el slug: " + cliente.slug);
  process.exit(1);
}

clientes.push({
  slug: cliente.slug,
  scriptId: cliente.scriptId,
  deploymentId: cliente.deploymentId,
  spreadsheetId: cliente.spreadsheetId,
  activo: true,
});

fs.writeFileSync(clientesPath, JSON.stringify(clientes, null, 2) + "\n");
console.log("Cliente agregado: " + cliente.slug);
