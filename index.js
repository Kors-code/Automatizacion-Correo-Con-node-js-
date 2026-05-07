const fs = require("fs");
const path = require("path");
const cron = require("node-cron");
const { downloadLatestZip } = require("./gmail.service");
const { processZipAndExtractExcel } = require("./zip.service");
const { uploadExcelToLaravel } = require("./hostinger.service");
const STATE_DIR = path.join(__dirname, "storage", "state");
const PROCESSED_PATH = path.join(STATE_DIR, "processed.json");

function ensureState() {
  fs.mkdirSync(STATE_DIR, { recursive: true });

  if (!fs.existsSync(PROCESSED_PATH)) {
    fs.writeFileSync(PROCESSED_PATH, JSON.stringify([], null, 2));
  }
}

function loadProcessed() {
  try {
    return JSON.parse(fs.readFileSync(PROCESSED_PATH, "utf8"));
  } catch {
    return [];
  }
}

function saveProcessed(list) {
  fs.writeFileSync(PROCESSED_PATH, JSON.stringify(list, null, 2));
}

async function runOnce() {
  const processed = loadProcessed();

  const result = await downloadLatestZip(processed);
  if (!result) {
    console.log("No hay ZIP nuevos que coincidan.");
    return;
  }

  console.log("Correo detectado:");
  console.log("Asunto:", result.subject);
  console.log("Remitente:", result.from);
  console.log("Tipo:", result.rule.type);
  console.log("ZIP:", result.zipPath);

  const finalExcel = processZipAndExtrractExcel(
    result.zipPath,
    result.rule.type,
    new Date()
  );

  if (finalExcel.finalPath) {
  console.log("Excel final:", finalExcel.finalPath);

  const response = await uploadExcelToLaravel(
    finalExcel.finalPath,
    4
  );

  console.log("Respuesta Laravel:");
  console.log(response);
}

  processed.push(result.messageId);
  saveProcessed(processed);
}

async function main() {
  ensureState();

  await runOnce();

  cron.schedule("*/5 * * * *", async () => {
    try {
      await runOnce();
    } catch (error) {
      console.error("Error en la tarea programada:", error.message);
    }
  });

  console.log("Proceso activo. Revisando cada 5 minutos.");
}

main().catch((error) => {
  console.error("Error general:", error.message);
});