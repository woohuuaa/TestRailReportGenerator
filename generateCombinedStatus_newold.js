require('dotenv').config();
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

// 📅 Format date to yyyymmdd
function formatDate(date = new Date()) {
  return `${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, '0')}${String(date.getDate()).padStart(2, '0')}`;
}

// 📁 Ensure directory exists
function ensureDir(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

// ⏱️ Helper: Get time spent in seconds
function getElapsedSeconds(startTime) {
  return ((new Date() - startTime) / 1000).toFixed(2);
}

// 📝 Logging setup (new file per run)
const logDateStr = formatDate();
const logDir = path.join(__dirname, 'logs');
ensureDir(logDir);
const logPath = path.join(logDir, `export_log_${logDateStr}.txt`);

function logToFile(message) {
  const timestamp = new Date().toISOString();
  const fullMessage = `[${timestamp}] ${message}\n`;
  fs.appendFileSync(logPath, fullMessage);
  console.log(message);
}

// 📤 Export Combined Status only (.xlsx)
async function exportToCombineStatus(data, filePath) {
  const start = new Date();
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Combined Status');

  worksheet.columns = Object.keys(data[0]).map(key => ({
    header: key,
    key: key,
    width: 20,
  }));

  data.forEach(row => worksheet.addRow(row));

  await workbook.xlsx.writeFile(filePath);
  logToFile(`✅ CombinedStatus exported to: ${filePath}`);
  logToFile(`⏱️ CombinedStatus export took ${getElapsedSeconds(start)} seconds`);
}

// 📤 Export using TestRailMetrics template (.xlsm)
async function exportToMetrics(combinedData, headerMap, prevHeaderMap, rowsCurr, rowsPrev, outputPath) {
  const start = new Date();
  const templatePath = path.join(__dirname, 'template', 'TestRailMetrics_template.xlsm');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);

  function writeSheet(sheetName, headerDef, rawRows) {
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) {
      logToFile(`⚠️ Sheet '${sheetName}' not found. Skipped.`);
      return;
    }

    while (sheet.rowCount > 0) {
      sheet.spliceRows(1, 1);
    }

    sheet.columns = Object.keys(headerDef).map(key => ({
      header: key,
      key: key,
      width: 20,
    }));

    rawRows.forEach(row => {
      const obj = {};
      for (const col in headerDef) {
        obj[col] = row[headerDef[col]];
      }
      sheet.addRow(obj);
    });
  }

  // CombinedStatus sheet
  let combinedSheet = workbook.getWorksheet('CombinedStatus');
  if (!combinedSheet) {
    combinedSheet = workbook.addWorksheet('CombinedStatus');
  } else {
    while (combinedSheet.rowCount > 0) {
      combinedSheet.spliceRows(1, 1);
    }
  }

  combinedSheet.columns = Object.keys(combinedData[0]).map(key => ({
    header: key,
    key: key,
    width: 20,
  }));

  combinedData.forEach(row => combinedSheet.addRow(row));

  // Write current & previous sprints
  writeSheet('currentSprint', headerMap, rowsCurr);
  writeSheet('previousSprint', prevHeaderMap, rowsPrev);

  await workbook.xlsx.writeFile(outputPath);
  logToFile(`✅ TestRailMetrics exported to: ${outputPath}`);
  logToFile(`⏱️ TestRailMetrics export took ${getElapsedSeconds(start)} seconds`);
}

// 🔄 Find latest and previous .xlsx filenames
function getBiweeklyFilenames(folder, prefix = 'brink_pos_') {
  const files = fs.readdirSync(folder).filter(file => file.startsWith(prefix) && file.endsWith('.xlsx'));
  const sortedDates = files
    .map(f => f.match(/(\d{8})/)?.[1])
    .filter(Boolean)
    .sort();

  if (sortedDates.length < 2) {
    logToFile('⚠️ Not enough bi-weekly files to compare.');
    return null;
  }

  return {
    latestDateStr: sortedDates[sortedDates.length - 1],
    latestPath: path.join(folder, `${prefix}${sortedDates[sortedDates.length - 1]}.xlsx`),
    lastWeekPath: path.join(folder, `${prefix}${sortedDates[sortedDates.length - 2]}.xlsx`)
  };
}

// 🧠 MAIN
(async () => {
  const scriptStart = new Date();
  try {
    logToFile('🚀 Export script started');

    const rawFolder = path.join(__dirname, 'raw_data');
    const combinedFolder = path.join(__dirname, 'CombinedStatus');
    const metricsFolder = path.join(__dirname, 'TestRailMetrics');

    ensureDir(combinedFolder);
    ensureDir(metricsFolder);

    const biweekly = getBiweeklyFilenames(rawFolder);
    if (!biweekly) return;

    const { latestPath, lastWeekPath, latestDateStr } = biweekly;

    const workbookPrev = new ExcelJS.Workbook();
    const workbookCurr = new ExcelJS.Workbook();
    await workbookPrev.xlsx.readFile(lastWeekPath);
    await workbookCurr.xlsx.readFile(latestPath);

    const sheetPrev = workbookPrev.getWorksheet(1);
    const sheetCurr = workbookCurr.getWorksheet(1);
    const rowsPrev = sheetPrev.getSheetValues().slice(2);
    const rowsCurr = sheetCurr.getSheetValues().slice(2);

    const headerMap = {};
    sheetCurr.getRow(1).eachCell((cell, colNumber) => {
      headerMap[cell.value] = colNumber;
    });

    const prevHeaderMap = {};
    sheetPrev.getRow(1).eachCell((cell, colNumber) => {
      prevHeaderMap[cell.value] = colNumber;
    });

    const prevMap = {};
    rowsPrev.forEach(row => {
      const id = row[1];
      if (id) {
        prevMap[id] = {
          'Previous Automation Status': row[prevHeaderMap['Automation Status']],
          'Previous Sprint Test Case Status': row[prevHeaderMap['Test Case Status']],
        };
      }
    });

    const combinedData = rowsCurr.map(row => {
      const id = row[1];
      const prev = Object.prototype.hasOwnProperty.call(prevMap, id)
        ? prevMap[id]
        : {
            'Previous Automation Status': 'N/A',
            'Previous Sprint Test Case Status': 'N/A',
          };
      return {
        ID: id,
        Title: row[headerMap['Title']],
        'Automation Area': row[headerMap['Automation Area']],
        'Automation Effort': row[headerMap['Automation Effort']],
        'Automation Priority': row[headerMap['Automation Priority']],
        'Previous Automation Status': prev['Previous Automation Status'],
        'Current Automation Status': row[headerMap['Automation Status']],
        'Team Name': row[headerMap['Team Name']],
        'Previous Sprint Test Case Status': prev['Previous Sprint Test Case Status'],
        'Current Sprint Test Case Status': row[headerMap['Test Case Status']],
        Type: row[headerMap['Type']],
      };
    });

    const outputCombined = path.join(combinedFolder, `CombinedStatus_${latestDateStr}.xlsx`);
    const outputMetrics = path.join(metricsFolder, `TestRailMetrics_${latestDateStr}.xlsx`);

    await exportToCombineStatus(combinedData, outputCombined);

    await exportToMetrics(
      combinedData,
      Object.fromEntries(Object.entries(headerMap)),
      Object.fromEntries(Object.entries(prevHeaderMap)),
      rowsCurr,
      rowsPrev,
      outputMetrics
    );

    logToFile(`✅ Script completed successfully in ${getElapsedSeconds(scriptStart)} seconds`);
  } catch (err) {
    const errorMsg = `🚨 Script Error: ${err.message}`;
    console.error(errorMsg);
    logToFile(errorMsg);
  }
})();