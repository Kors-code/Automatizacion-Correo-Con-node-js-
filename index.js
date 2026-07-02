const fs = require("fs");
const path = require("path");
const cron = require("node-cron");
const { downloadLatestZip } = require("./gmail.service");
const { processZipAndExtractExcel } = require("./zip.service");
const {
  uploadInventoryExcelToLaravel,
  uploadSalesExcelToLaravel,
  uploadCatalogExcelToLaravel,
} = require("./hostinger.service");
const { uploadToOneDrive } = require("./onedrive.service");

const STATE_DIR = path.join(__dirname, "storage", "state");
const PROCESSED_PATH = path.join(STATE_DIR, "processed.json");
const LOCK_PATH = path.join(STATE_DIR, "import.lock.json");
const STALE_LOCK_MS = 1000 * 60 * 30;
const DEFAULT_ONEDRIVE_SALES_FOLDER = "Documents/test/ventas";
const SERVER_LOG_DIR = path.join(__dirname, "storage", "server-logs");
const STATE_WRITE_RETRIES = 5;
const STATE_WRITE_RETRY_MS = 500;
const CATALOG_CONTINUE_ON_FAILURE =
  String(process.env.CATALOG_CONTINUE_ON_FAILURE || "true").toLowerCase() ===
  "true";
let activeRun = false;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

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

function writeProcessedFile(list) {
  fs.mkdirSync(STATE_DIR, { recursive: true });

  const payload = JSON.stringify(list, null, 2);
  const tempPath = `${PROCESSED_PATH}.${process.pid}.tmp`;

  if (fs.existsSync(PROCESSED_PATH)) {
    try {
      fs.chmodSync(PROCESSED_PATH, 0o666);
    } catch {
      // Windows puede negar chmod si otro proceso tiene el archivo abierto.
    }
  }

  fs.writeFileSync(tempPath, payload);

  try {
    fs.renameSync(tempPath, PROCESSED_PATH);
  } catch (error) {
    if (error.code !== "EEXIST" && error.code !== "EPERM") {
      throw error;
    }

    if (fs.existsSync(PROCESSED_PATH)) {
      fs.unlinkSync(PROCESSED_PATH);
    }

    fs.renameSync(tempPath, PROCESSED_PATH);
  }
}

async function saveProcessed(list) {
  let lastError;

  for (let attempt = 1; attempt <= STATE_WRITE_RETRIES; attempt += 1) {
    try {
      writeProcessedFile(list);
      return;
    } catch (error) {
      lastError = error;

      if (!["EPERM", "EACCES", "EBUSY"].includes(error.code)) {
        throw error;
      }

      console.error(
        `No se pudo escribir processed.json intento ${attempt}/${STATE_WRITE_RETRIES}: ${error.code}`
      );

      if (attempt < STATE_WRITE_RETRIES) {
        await sleep(STATE_WRITE_RETRY_MS * attempt);
      }
    }
  }

  throw lastError;
}

function sanitizeLogName(value) {
  return String(value || "sin-nombre")
    .replace(/[<>:"/\\|?*\s]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function writeServerErrorLog(serviceName, error, context = {}) {
  fs.mkdirSync(SERVER_LOG_DIR, { recursive: true });

  const status = error.response?.status || null;
  const payload = {
    createdAt: new Date().toISOString(),
    service: serviceName,
    status,
    message: error.message,
    code: error.code || null,
    context,
    response: {
      headers: error.response?.headers || null,
      data: error.response?.data || null,
    },
  };

  const fileName = [
    Date.now(),
    sanitizeLogName(serviceName),
    status || "sin_status",
    sanitizeLogName(context.reportType),
    sanitizeLogName(context.company),
  ].join("-");

  const logPath = path.join(SERVER_LOG_DIR, `${fileName}.json`);
  fs.writeFileSync(logPath, JSON.stringify(payload, null, 2));

  console.error(`${serviceName} fallo. Log guardado en: ${logPath}`);

  if (status) {
    console.error(`Status ${serviceName}: ${status}`);
  }

  if (error.response?.data) {
    console.error(`Respuesta ${serviceName}:`);
    console.error(error.response.data);
  }

  return logPath;
}

function isProcessAlive(pid) {
  const numericPid = Number(pid);

  if (!numericPid) {
    return false;
  }

  try {
    process.kill(numericPid, 0);
    return true;
  } catch (error) {
    return error.code === "EPERM";
  }
}

function acquireRunLock() {
  fs.mkdirSync(STATE_DIR, { recursive: true });

  const lock = {
    token: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
    startedAt: Date.now(),
    pid: process.pid,
  };

  try {
    fs.writeFileSync(LOCK_PATH, JSON.stringify(lock, null, 2), { flag: "wx" });
    return lock;
  } catch (error) {
    if (error.code !== "EEXIST") {
      throw error;
    }
  }

  try {
    const current = JSON.parse(fs.readFileSync(LOCK_PATH, "utf8"));
    const age = Date.now() - Number(current.startedAt || 0);

    if (isProcessAlive(current.pid)) {
      console.log(`Candado activo por proceso ${current.pid}. Se omite este ciclo.`);
      return null;
    }

    if (current.pid) {
      console.log(`Candado huerfano del proceso ${current.pid}. Se reemplaza.`);
      fs.unlinkSync(LOCK_PATH);
      fs.writeFileSync(LOCK_PATH, JSON.stringify(lock, null, 2), { flag: "wx" });
      return lock;
    }

    if (age < STALE_LOCK_MS) {
      return null;
    }

    console.log("Candado anterior vencido o huerfano. Se reemplaza.");
    fs.unlinkSync(LOCK_PATH);
    fs.writeFileSync(LOCK_PATH, JSON.stringify(lock, null, 2), { flag: "wx" });
    return lock;
  } catch (error) {
    if (fs.existsSync(LOCK_PATH)) {
      fs.unlinkSync(LOCK_PATH);
    }

    fs.writeFileSync(LOCK_PATH, JSON.stringify(lock, null, 2), { flag: "wx" });
    return lock;
  }
}

function releaseRunLock(lock) {
  if (!lock || !fs.existsSync(LOCK_PATH)) {
    return;
  }

  try {
    const current = JSON.parse(fs.readFileSync(LOCK_PATH, "utf8"));

    if (current.token === lock.token) {
      fs.unlinkSync(LOCK_PATH);
    }
  } catch {
    // Si el lock se corrompe, no lo borramos a ciegas.
  }
}

function buildOneDriveSalesFolder(rule) {
  const baseFolder =
    process.env.ONEDRIVE_SALES_FOLDER || DEFAULT_ONEDRIVE_SALES_FOLDER;
  const company = String(rule.company || rule.folder || "general").toUpperCase();

  return `${baseFolder.replace(/[\\/]+$/g, "")}/${company}`;
}

async function uploadSalesToOneDrive(localPath, rule) {
  if (rule.reportType !== "SALES") {
    return null;
  }

  const fileName = path.basename(localPath);
  const folderPath = buildOneDriveSalesFolder(rule);

  console.log("Subiendo venta a OneDrive...");
  console.log("Carpeta OneDrive:", folderPath);

  const response = await uploadToOneDrive(localPath, fileName, folderPath);

  console.log("Archivo subido a OneDrive:");
  console.log(response.webUrl || response.name || response.id);

  return response;
}

async function uploadToLaravelSafely(uploadFn, localPath, result) {
  try {
    const response = await uploadFn();

    console.log("Respuesta Laravel:");
    console.log(response);

    return response;
  } catch (error) {
    writeServerErrorLog("laravel", error, {
      subject: result.subject,
      messageId: result.messageId,
      reportType: result.rule.reportType,
      company: result.rule.company,
      ruleKey: result.rule.key,
      excelPath: localPath,
    });

    return null;
  }
}

async function uploadToOneDriveSafely(localPath, result) {
  try {
    return await uploadSalesToOneDrive(localPath, result.rule);
  } catch (error) {
    writeServerErrorLog("onedrive", error, {
      subject: result.subject,
      messageId: result.messageId,
      reportType: result.rule.reportType,
      company: result.rule.company,
      ruleKey: result.rule.key,
      excelPath: localPath,
    });

    return null;
  }
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
  console.log("Tipo:", result.rule.reportType || result.rule.type);
  console.log("ZIP:", result.zipPath);

  const finalExcel = processZipAndExtractExcel(
    result.zipPath,
    result.rule.reportType,
    new Date(),
    result.rule.company,
    result.rule.type
  );

  if (finalExcel.finalPath) {
    console.log("Excel final:", finalExcel.finalPath);
    let importCompleted = false;

    if (result.rule.reportType === "INVENTORY") {
      console.log("Procesando INVENTARIO...");
      const laravelResponse = await uploadToLaravelSafely(
        () =>
          uploadInventoryExcelToLaravel(finalExcel.finalPath, result.rule.storeId),
        finalExcel.finalPath,
        result
      );

      importCompleted = Boolean(laravelResponse);
    } else if (result.rule.reportType === "SALES") {
      console.log("Procesando VENTAS...");
      const laravelResponse = await uploadToLaravelSafely(
        () => uploadSalesExcelToLaravel(finalExcel.finalPath, result.rule.storeId),
        finalExcel.finalPath,
        result
      );
      const oneDriveResponse = laravelResponse
        ? await uploadToOneDriveSafely(finalExcel.finalPath, result)
        : null;

      importCompleted = Boolean(laravelResponse && oneDriveResponse);
    } else if (result.rule.reportType === "CATALOG") {
      console.log("Procesando CATALOGO...");
      const laravelResponse = await uploadToLaravelSafely(
        () => uploadCatalogExcelToLaravel(finalExcel.finalPath),
        finalExcel.finalPath,
        result
      );

      importCompleted = Boolean(laravelResponse);
    } else {
      throw new Error(`Tipo de reporte no soportado: ${result.rule.reportType}`);
    }

    if (!importCompleted) {
      if (result.rule.reportType === "CATALOG" && CATALOG_CONTINUE_ON_FAILURE) {
        console.log(
          "El catalogo no se completo. Se marca como procesado para no bloquear ventas e inventario; revisa el log guardado."
        );
      } else {
        console.log(
          "La importacion no se completo. El correo no se marca como procesado para reintentar en el proximo ciclo."
        );
        return;
      }
    }
  }

  processed.push(result.messageId);
  await saveProcessed(processed);
}

async function runSafely() {
  if (activeRun) {
    console.log("Ya hay una importacion activa en este proceso. Se omite este ciclo.");
    return;
  }

  activeRun = true;
  const lock = acquireRunLock();

  if (!lock) {
    activeRun = false;
    console.log("Ya hay una importacion en proceso. Se omite este ciclo.");
    return;
  }

  try {
    await runOnce();
  } finally {
    releaseRunLock(lock);
    activeRun = false;
  }
}

async function main() {
  ensureState();

  await runSafely();

  cron.schedule("*/5 * * * *", async () => {
    try {
      await runSafely();
    } catch (error) {
      console.error("Error en la tarea programada:", error.message);
    }
  });

  console.log("Proceso activo. Revisando cada 5 minutos.");
}

main().catch((error) => {
  console.error("Error general:", error.message);

  if (error.response) {
    writeServerErrorLog("general", error);
  }
});
