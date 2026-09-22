const fs = require("fs");
const path = require("path");
const AdmZip = require("adm-zip");

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

function getDefaultFinalBaseDir() {
  if (process.env.FINAL_BASE_DIR) {
    return process.env.FINAL_BASE_DIR;
  }

  return process.platform === "win32"
    ? "C:\\Sells"
    : path.join(__dirname, "storage", "final");
}

const FINAL_BASE_DIR = getDefaultFinalBaseDir();

const FINAL_PATHS = {
  SALES: {
    DFP: path.join(FINAL_BASE_DIR, "DFP"),
    LDC: path.join(FINAL_BASE_DIR, "LDC"),
  },
  INVENTORY: {
    DFP: path.join(FINAL_BASE_DIR, "Inventory", "DFP"),
    LDC: path.join(FINAL_BASE_DIR, "Inventory", "LDC"),
    ZF: path.join(FINAL_BASE_DIR, "Inventory", "ZF"),
    DEFAULT: path.join(FINAL_BASE_DIR, "Inventory"),
  },
  CATALOG: {
    DFP: path.join(FINAL_BASE_DIR, "Catalog", "DFP"),
    LDC: path.join(FINAL_BASE_DIR, "Catalog", "LDC"),
    DEFAULT: path.join(FINAL_BASE_DIR, "Catalog"),
  },
};

function getMonthNameES(monthIndex) {
  const months = [
    "Enero",
    "Febrero",
    "Marzo",
    "Abril",
    "Mayo",
    "Junio",
    "Julio",
    "Agosto",
    "Septiembre",
    "Octubre",
    "Noviembre",
    "Diciembre",
  ];

  return months[monthIndex];
}

function buildFinalExcelName(reportType, company, date = new Date()) {
  const monthName = getMonthNameES(date.getMonth());
  const year = date.getFullYear();

  return `${reportType} ${company} COLOMBIA ${monthName} ${year}.xlsx`;
}

function unzipFile(zipPath) {
  const extractedRoot = path.join(
    __dirname,
    "storage",
    "extracted",
    path.basename(zipPath, path.extname(zipPath))
  );

  ensureDir(extractedRoot);

  const zip = new AdmZip(zipPath);
  zip.extractAllTo(extractedRoot, true);

  return extractedRoot;
}

function walkFiles(dir, collected = []) {
  const entries = fs.readdirSync(dir, { withFileTypes: true });

  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);

    if (entry.isDirectory()) {
      walkFiles(fullPath, collected);
    } else {
      collected.push(fullPath);
    }
  }

  return collected;
}

function findExcelFile(folderPath) {
  const files = walkFiles(folderPath);

  return (
    files.find((file) => {
      const lower = file.toLowerCase();
      return (
        lower.endsWith(".xlsx") ||
        lower.endsWith(".xls") ||
        lower.endsWith(".csv")
      );
    }) || null
  );
}

function resolveFinalDir(reportType, ruleType = null, company = null) {
  const normalizedReportType = String(reportType || "").toUpperCase();
  const normalizedCompany = String(company || ruleType || "").toUpperCase();

  if (FINAL_PATHS[normalizedReportType]) {
    const byCompany = FINAL_PATHS[normalizedReportType][normalizedCompany];
    if (byCompany) {
      return byCompany;
    }

    return FINAL_PATHS[normalizedReportType].DEFAULT || null;
  }

  return null;
}

function processZipAndExtractExcel(zipPath, reportType, date = new Date(), company = null, ruleType = null
, mode = null) {
  const extractedFolder = unzipFile(zipPath);
  const excelFile = findExcelFile(extractedFolder);

  if (!excelFile) {
    console.log("No se encontró ningún Excel dentro del ZIP.");
    return {
      extractedFolder,
      excelFile: null,
      finalPath: null,
    };
  }

  const finalDir = resolveFinalDir(reportType, ruleType, company);

  if (!finalDir) {
    throw new Error(`No existe ruta configurada para tipo ${reportType}`);
  }

  fs.mkdirSync(finalDir, { recursive: true });

  const finalName = buildFinalExcelName(
    reportType,
    company,
    date,
    mode
  );
  const finalPath = path.join(finalDir, finalName);

  fs.copyFileSync(excelFile, finalPath);

  console.log("Excel encontrado:", excelFile);
  console.log("Excel guardado como:", finalPath);

  return {
    extractedFolder,
    excelFile,
    finalPath,
  };
}
function getPreviousMonthDate(date = new Date()) {
  return new Date(date.getFullYear(), date.getMonth() - 1, 1);
}

function buildFinalExcelName(reportType, company, date = new Date(), mode = null) {
  const targetDate = mode === "FINAL_MONTH" ? getPreviousMonthDate(date) : date;
  const monthName = getMonthNameES(targetDate.getMonth());
  const year = targetDate.getFullYear();
  const suffix = mode === "FINAL_MONTH" ? " FINAL" : "";

  return `${reportType} ${company} COLOMBIA ${monthName} ${year}${suffix}.xlsx`;
}

module.exports = {
  processZipAndExtractExcel,
};
