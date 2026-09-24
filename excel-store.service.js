const path = require("path");
const AdmZip = require("adm-zip");
const ExcelJS = require("exceljs");

function normalizeHeader(value) {
  return String(value || "")
    .trim()
    .toLowerCase();
}

function getCellText(cell) {
  if (!cell || cell.value == null) {
    return "";
  }

  if (typeof cell.value === "object") {
    if (cell.value.text) {
      return String(cell.value.text);
    }

    if (Array.isArray(cell.value.richText)) {
      return cell.value.richText.map((part) => part.text || "").join("");
    }
  }

  return String(cell.value);
}

function findStoreColumn(worksheet) {
  const maxHeaderRows = Math.min(10, worksheet.rowCount || 10);

  for (let rowNumber = 1; rowNumber <= maxHeaderRows; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);

    for (let colNumber = 1; colNumber <= row.cellCount; colNumber += 1) {
      const header = normalizeHeader(getCellText(row.getCell(colNumber)));

      if (header === "store") {
        return {
          rowNumber,
          colNumber,
        };
      }
    }
  }

  return null;
}

function getTodayDateOnly() {
  const timeZone = process.env.REPORT_DATE_TIMEZONE || "America/Bogota";
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(new Date());

  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));

  return new Date(
    Number(values.year),
    Number(values.month) - 1,
    Number(values.day)
  );
}

function getTodayDateText() {
  const date = getTodayDateOnly();
  const year = String(date.getFullYear());
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");

  return `${year}-${month}-${day}`;
}

function rowHasDataOutsideColumns(row, ignoredColumns) {
  let hasData = false;

  row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    if (ignoredColumns.has(colNumber)) {
      return;
    }

    if (getCellText(cell).trim() !== "") {
      hasData = true;
    }
  });

  return hasData;
}

function columnNumberToName(columnNumber) {
  let name = "";
  let current = columnNumber;

  while (current > 0) {
    const remainder = (current - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    current = Math.floor((current - 1) / 26);
  }

  return name;
}

function columnNameToNumber(columnName) {
  return String(columnName || "")
    .toUpperCase()
    .split("")
    .reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0);
}

function escapeXml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function getCellStyleAttribute(cellXml) {
  const styleMatch = String(cellXml || "").match(/\ss="[^"]*"/);
  return styleMatch ? styleMatch[0] : "";
}

function buildInlineStringCell(cellRef, value, previousCellXml = "") {
  const style = getCellStyleAttribute(previousCellXml);
  return `<c r="${cellRef}"${style} t="inlineStr"><is><t>${escapeXml(
    value
  )}</t></is></c>`;
}

function upsertCell(rowXml, rowNumber, colNumber, value) {
  const columnName = columnNumberToName(colNumber);
  const cellRef = `${columnName}${rowNumber}`;
  const cellRegex = /<c\b[^>]*\br="([A-Z]+)(\d+)"[^>]*(?:\/>|>[\s\S]*?<\/c>)/g;
  let existingCellXml = "";

  const replaced = rowXml.replace(cellRegex, (cellXml, colName, cellRow) => {
    if (colName === columnName && Number(cellRow) === rowNumber) {
      existingCellXml = cellXml;
      return buildInlineStringCell(cellRef, value, cellXml);
    }

    return cellXml;
  });

  if (existingCellXml) {
    return replaced;
  }

  const newCellXml = buildInlineStringCell(cellRef, value);
  let insertAt = -1;

  replaced.replace(cellRegex, (cellXml, colName, cellRow, offset) => {
    if (insertAt === -1 && Number(cellRow) === rowNumber) {
      const currentColumnNumber = columnNameToNumber(colName);

      if (currentColumnNumber > colNumber) {
        insertAt = offset;
      }
    }

    return cellXml;
  });

  if (insertAt !== -1) {
    return `${replaced.slice(0, insertAt)}${newCellXml}${replaced.slice(
      insertAt
    )}`;
  }

  return replaced.replace("</row>", `${newCellXml}</row>`);
}

function updateWorksheetXml(worksheetXml, updates) {
  return worksheetXml.replace(
    /<row\b[^>]*\br="(\d+)"[^>]*>[\s\S]*?<\/row>/g,
    (rowXml, rowNumberText) => {
      const rowNumber = Number(rowNumberText);
      const rowUpdate = updates.get(rowNumber);

      if (!rowUpdate) {
        return rowXml;
      }

      let updatedRowXml = rowXml;

      for (const cellUpdate of rowUpdate) {
        updatedRowXml = upsertCell(
          updatedRowXml,
          rowNumber,
          cellUpdate.colNumber,
          cellUpdate.value
        );
      }

      return updatedRowXml;
    }
  );
}

function writeWorksheetXml(xlsxPath, updates) {
  const zip = new AdmZip(xlsxPath);
  const worksheetEntry = zip.getEntry("xl/worksheets/sheet1.xml");

  if (!worksheetEntry) {
    throw new Error("No se encontro xl/worksheets/sheet1.xml en el Excel.");
  }

  const worksheetXml = worksheetEntry.getData().toString("utf8");
  const updatedWorksheetXml = updateWorksheetXml(worksheetXml, updates);

  zip.updateFile("xl/worksheets/sheet1.xml", Buffer.from(updatedWorksheetXml));
  zip.writeZip(xlsxPath);
}

async function fillStoreColumn(xlsxPath, storeValue) {
  if (!storeValue || path.extname(xlsxPath).toLowerCase() !== ".xlsx") {
    return false;
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(xlsxPath);

  const worksheet = workbook.worksheets[0];

  if (!worksheet) {
    throw new Error("No se encontro una hoja en el Excel.");
  }

  const storeColumn = findStoreColumn(worksheet);

  if (!storeColumn) {
    throw new Error("No se encontro la columna Store en el Excel.");
  }

  let updatedRows = 0;
  const dateColumnNumber = storeColumn.colNumber + 1;
  const reportDate = getTodayDateText();
  const ignoredColumns = new Set([storeColumn.colNumber, dateColumnNumber]);
  const updates = new Map();

  for (
    let rowNumber = storeColumn.rowNumber + 1;
    rowNumber <= worksheet.rowCount;
    rowNumber += 1
  ) {
    const row = worksheet.getRow(rowNumber);

    if (!rowHasDataOutsideColumns(row, ignoredColumns)) {
      continue;
    }

    updates.set(rowNumber, [
      { colNumber: storeColumn.colNumber, value: storeValue },
      { colNumber: dateColumnNumber, value: reportDate },
    ]);

    updatedRows += 1;
  }

  if (updatedRows > 0) {
    writeWorksheetXml(xlsxPath, updates);
  }

  return updatedRows;
}

module.exports = {
  fillStoreColumn,
};
