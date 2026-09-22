const path = require("path");
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

function rowHasDataOutsideStore(row, storeColumnNumber) {
  let hasData = false;

  row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    if (colNumber === storeColumnNumber) {
      return;
    }

    if (getCellText(cell).trim() !== "") {
      hasData = true;
    }
  });

  return hasData;
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

  for (
    let rowNumber = storeColumn.rowNumber + 1;
    rowNumber <= worksheet.rowCount;
    rowNumber += 1
  ) {
    const row = worksheet.getRow(rowNumber);

    if (!rowHasDataOutsideStore(row, storeColumn.colNumber)) {
      continue;
    }

    row.getCell(storeColumn.colNumber).value = storeValue;
    row.commit();
    updatedRows += 1;
  }

  await workbook.xlsx.writeFile(xlsxPath);

  return updatedRows;
}

module.exports = {
  fillStoreColumn,
};
