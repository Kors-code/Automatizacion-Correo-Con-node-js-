const fs = require("fs");
const path = require("path");
const AdmZip = require("adm-zip");

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

const FINAL_PATHS = {
  SALES: {
    DFP: "C:\\Sells\\DFP",
    LDC: "C:\\Sells\\LDC",
  },
  INVENTORY: {
    DFP: "C:\\Sells\\Inventory\\DFP",
    LDC: "C:\\Sells\\Inventory\\LDC",
    ZF: "C:\\Sells\\Inventory\\ZF",
    DEFAULT: "C:\\Sells\\Inventory",
  },
  CATALOG: {
    DFP: "C:\\Sells\\Catalog\\DFP",
    LDC: "C:\\Sells\\Catalog\\LDC",
    DEFAULT: "C:\\Sells\\Catalog",
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

function processZipAndExtractExcel(zipPath, reportType, date = new Date(), company = null, ruleType = null) {
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
    date
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

module.exports = {
  processZipAndExtractExcel,
};
