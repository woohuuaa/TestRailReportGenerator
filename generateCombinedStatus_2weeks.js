require('dotenv').config();
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

// 📅 Format date to yyyymmdd
function formatDate(date) {
  return `${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, '0')}${String(date.getDate()).padStart(2, '0')}`;
}

// 🧮 Calculate date difference in days
function dateDiffInDays(a, b) {
  const msPerDay = 24 * 60 * 60 * 1000;
  return Math.abs((a - b) / msPerDay);
}

// 📤 Export JSON data to Excel
async function exportToExcel(data, filename) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Combined Status');

  if (data.length === 0) {
    console.warn('⚠️ No data to export.');
    return;
  }

  worksheet.columns = Object.keys(data[0]).map(key => ({
    header: key,
    key: key,
    width: 20,
  }));

  data.forEach(row => worksheet.addRow(row));

  await workbook.xlsx.writeFile(filename);
  console.log(`✅ Combined file saved: ${filename}`);
}

// 🔄 Get most recent bi-weekly file and its previous one
function getBiweeklyFilenames(folder, patternPrefix = 'brink_pos_', startDate = new Date('2025-04-11')) {
  const files = fs.readdirSync(folder).filter(file => file.startsWith(patternPrefix) && file.endsWith('.xlsx'));
  const sortedDates = files
    .map(f => f.match(/(\d{8})/)?.[1])
    .filter(Boolean)
    .sort();

  if (sortedDates.length < 2) {
    console.warn('⚠️ Not enough bi-weekly files to compare.');
    return null;
  }

  const latestDateStr = sortedDates[sortedDates.length - 1];
  const prevDateStr = sortedDates[sortedDates.length - 2];

  const latestDate = new Date(
    parseInt(latestDateStr.slice(0, 4)),
    parseInt(latestDateStr.slice(4, 6)) - 1,
    parseInt(latestDateStr.slice(6, 8))
  );

  const prevDate = new Date(
    parseInt(prevDateStr.slice(0, 4)),
    parseInt(prevDateStr.slice(4, 6)) - 1,
    parseInt(prevDateStr.slice(6, 8))
  );

  const diffDays = dateDiffInDays(latestDate, prevDate);
  if (diffDays !== 14) {
    console.warn(`⚠️ Expected 14 days between reports, but found ${diffDays} days.`);
    return null;
  }

  return {
    latestPath: path.join(folder, `${patternPrefix}${latestDateStr}.xlsx`),
    lastWeekPath: path.join(folder, `${patternPrefix}${prevDateStr}.xlsx`),
    latestDateStr,
  };
}

(async () => {
  try {
    const folder = path.join(__dirname, 'raw_data');
    const biweekly = getBiweeklyFilenames(folder);

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

    const prevMap = {};
    rowsPrev.forEach(row => {
      if (row && row[1]) {
        prevMap[row[1]] = {
          'Previous Automation Status': row[headerMap['Automation Status']],
          'Previous Sprint Test Case Status': row[headerMap['Test Case Status']],
        };
      }
    });

    const combinedData = rowsCurr.map(row => {
      const id = row[1];
      const prev = prevMap[id] || {
        'Previous Automation Status': '',
        'Previous Sprint Test Case Status': '',
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

    const outputCombined = path.join(folder, `CombinedStatus_${latestDateStr}.xlsx`);
    await exportToExcel(combinedData, outputCombined);
  } catch (err) {
    console.error('🚨 Script Error:', err.message);
  }
})();
