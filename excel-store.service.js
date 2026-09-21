const path = require("path");
const AdmZip = require("adm-zip");

function decodeXml(value) {
  return String(value || "")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

function escapeXml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function columnNameToNumber(columnName) {
  return String(columnName || "")
    .toUpperCase()
    .split("")
    .reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0);
}

function parseCellRef(ref) {
  const match = String(ref || "").match(/^([A-Z]+)(\d+)$/i);

  if (!match) {
    return null;
  }

  return {
    column: match[1].toUpperCase(),
    row: Number(match[2]),
  };
}

function readSharedStrings(zip) {
  const entry = zip.getEntry("xl/sharedStrings.xml");

  if (!entry) {
    return [];
  }

  const xml = entry.getData().toString("utf8");
  const strings = [];
  const siRegex = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
  let match;

  while ((match = siRegex.exec(xml))) {
    const textParts = [];
    const textRegex = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
    let textMatch;

    while ((textMatch = textRegex.exec(match[1]))) {
      textParts.push(decodeXml(textMatch[1]));
    }

    strings.push(textParts.join(""));
  }

  return strings;
}

function getCellText(cellXml, sharedStrings) {
  const type = cellXml.match(/\bt="([^"]+)"/)?.[1];

  if (type === "inlineStr") {
    const inlineText = cellXml.match(/<t\b[^>]*>([\s\S]*?)<\/t>/)?.[1];
    return decodeXml(inlineText || "");
  }

  const rawValue = cellXml.match(/<v>([\s\S]*?)<\/v>/)?.[1] || "";

  if (type === "s") {
    return sharedStrings[Number(rawValue)] || "";
  }

  return decodeXml(rawValue);
}

function findStoreColumn(sheetXml, sharedStrings) {
  const rowRegex = /<row\b[^>]*\br="(\d+)"[^>]*>[\s\S]*?<\/row>/g;
  let rowMatch;

  while ((rowMatch = rowRegex.exec(sheetXml))) {
    const rowNumber = Number(rowMatch[1]);

    if (rowNumber > 10) {
      break;
    }

    const cellRegex = /<c\b[^>]*\br="([^"]+)"[^>]*(?:>[\s\S]*?<\/c>|\/>)/g;
    let cellMatch;

    while ((cellMatch = cellRegex.exec(rowMatch[0]))) {
      const ref = parseCellRef(cellMatch[1]);

      if (!ref) {
        continue;
      }

      const cellText = getCellText(cellMatch[0], sharedStrings).trim().toLowerCase();

      if (cellText === "store") {
        return ref.column;
      }
    }
  }

  return null;
}

function getLastRowNumber(sheetXml) {
  let lastRow = 1;
  const rowRegex = /<row\b[^>]*\br="(\d+)"/g;
  let match;

  while ((match = rowRegex.exec(sheetXml))) {
    lastRow = Math.max(lastRow, Number(match[1]));
  }

  return lastRow;
}

function makeInlineStringCell(cellRef, value) {
  return `<c r="${cellRef}" t="inlineStr"><is><t>${escapeXml(value)}</t></is></c>`;
}

function upsertCellInRow(rowXml, columnName, rowNumber, value) {
  const cellRef = `${columnName}${rowNumber}`;
  const newCell = makeInlineStringCell(cellRef, value);
  const existingCellRegex = new RegExp(
    `<c\\b[^>]*\\br="${cellRef}"[^>]*(?:>[\\s\\S]*?<\\/c>|\\/>)`
  );

  if (existingCellRegex.test(rowXml)) {
    return rowXml.replace(existingCellRegex, newCell);
  }

  const cellRegex = /<c\b[^>]*\br="([^"]+)"[^>]*(?:>[\s\S]*?<\/c>|\/>)/g;
  let insertAt = -1;
  let match;
  const targetColumnNumber = columnNameToNumber(columnName);

  while ((match = cellRegex.exec(rowXml))) {
    const ref = parseCellRef(match[1]);

    if (ref && columnNameToNumber(ref.column) > targetColumnNumber) {
      insertAt = match.index;
      break;
    }
  }

  if (insertAt >= 0) {
    return `${rowXml.slice(0, insertAt)}${newCell}${rowXml.slice(insertAt)}`;
  }

  return rowXml.replace("</row>", `${newCell}</row>`);
}

function fillStoreColumn(xlsxPath, storeValue) {
  if (!storeValue || path.extname(xlsxPath).toLowerCase() !== ".xlsx") {
    return false;
  }

  const zip = new AdmZip(xlsxPath);
  const sheetEntry = zip.getEntry("xl/worksheets/sheet1.xml");

  if (!sheetEntry) {
    throw new Error("No se encontro xl/worksheets/sheet1.xml en el Excel.");
  }

  const sharedStrings = readSharedStrings(zip);
  const originalXml = sheetEntry.getData().toString("utf8");
  const storeColumn = findStoreColumn(originalXml, sharedStrings);

  if (!storeColumn) {
    throw new Error("No se encontro la columna Store en el Excel.");
  }

  const lastRow = getLastRowNumber(originalXml);
  const rowRegex = /<row\b[^>]*\br="(\d+)"[^>]*>[\s\S]*?<\/row>/g;
  const updatedXml = originalXml.replace(rowRegex, (rowXml, rowNumberText) => {
    const rowNumber = Number(rowNumberText);

    if (rowNumber <= 1 || rowNumber > lastRow) {
      return rowXml;
    }

    return upsertCellInRow(rowXml, storeColumn, rowNumber, storeValue);
  });

  zip.updateFile("xl/worksheets/sheet1.xml", Buffer.from(updatedXml, "utf8"));
  zip.writeZip(xlsxPath);

  return true;
}

module.exports = {
  fillStoreColumn,
};
