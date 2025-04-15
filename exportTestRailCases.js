require('dotenv').config();
const axios = require('axios');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

// Setup TestRail client
const testrail = axios.create({
  baseURL: `${process.env.TESTRAIL_URL}/index.php?/api/v2/`,
  auth: {
    username: process.env.TESTRAIL_USER,
    password: process.env.TESTRAIL_API_KEY,
  },
  headers: {
    'Content-Type': 'application/json',
  },
});

// Export JSON data to Excel
async function exportToExcel(data, filename = 'output.xlsx') {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Test Cases');

  if (data.length === 0) {
    console.warn('⚠️ No data to export.');
    return;
  }

  worksheet.columns = Object.keys(data[0]).map(key => ({
    header: key,
    key: key,
    width: 10,
  }));

  data.forEach(row => worksheet.addRow(row));
  worksheet.getRow(1).font = { bold: true };

  await workbook.xlsx.writeFile(filename);
  console.log(`✅ Exported to ${filename}`);
}

// Today's date (yyyymmdd)
function getTodayDateString() {
  const now = new Date();
  return `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
}

// Get project ID
async function getProjectIdByName(name) {
  const response = await testrail.get('get_projects');
  const projects = response.data.projects || response.data;
  const project = projects.find(p => p.name.toLowerCase() === name.toLowerCase());
  return project ? project.id : null;
}

// Get all test cases with pagination
async function getTestCases(projectId) {
  const allCases = [];
  let offset = 0;
  const limit = 250;

  try {
    while (true) {
      const response = await testrail.get(`get_cases/${projectId}&limit=${limit}&offset=${offset}`);
      const cases = response.data.cases || response.data;

      if (!Array.isArray(cases)) {
        console.error('❌ Unexpected test case format:', response.data);
        break;
      }

      allCases.push(...cases);

      if (cases.length < limit) break; // Last page
      offset += limit;
    }

    return allCases;
  } catch (err) {
    console.error('🚨 Failed to get test cases:', err.response?.data || err.message);
    return [];
  }
}

// Get field definitions
async function getCaseFields() {
  try {
    const response = await testrail.get('get_case_fields');
    return response.data;
  } catch (err) {
    console.error('🚨 Failed to fetch case fields:', err.response?.data || err.message);
    return [];
  }
}

// Get case type names
async function getCaseTypes() {
  try {
    const response = await testrail.get('get_case_types');
    return response.data.reduce((map, type) => {
      map[type.id] = type.name;
      return map;
    }, {});
  } catch (err) {
    console.error('🚨 Failed to fetch case types:', err.response?.data || err.message);
    return {};
  }
}

// Build dropdown value map for a field
function buildFieldValueMap(field) {
  const map = {};
  const rawItems = field?.configs?.[0]?.options?.items || '';
  rawItems.split('\n').forEach(line => {
    const [key, value] = line.split(',').map(s => s.trim());
    if (key && value) map[key] = value;
  });
  return map;
}

// Format test cases with mapped dropdown labels
function formatTestCases(cases, fieldMaps, typeMap) {
  return cases.map(test => ({
    ID: `C${test.id}`,
    Title: test.title?.replace(/"/g, '""'),
    'Automation Area': Array.isArray(test.custom_automationarea)
      ? test.custom_automationarea.map(id => fieldMaps.automation_area[id]).filter(Boolean).join('\n')
      : (fieldMaps.automation_area[test.custom_automationarea] || ''),
    'Automation Effort': fieldMaps.automation_effort[test.custom_automationeffort]
      ? ' ' + fieldMaps.automation_effort[test.custom_automationeffort]
      : '',
    'Automation Priority': fieldMaps.automation_priority[test.custom_autopriority]
      ? ' ' + fieldMaps.automation_priority[test.custom_autopriority]
      : '',
    'Automation Status': fieldMaps.automation_status[test.custom_automation_status] === 'Retired'
      ? ' ' + fieldMaps.automation_status[test.custom_automation_status]
      : (fieldMaps.automation_status[test.custom_automation_status] || ''),
    'Team Name': fieldMaps.team_name[test.custom_teamname]
      ? ' ' + fieldMaps.team_name[test.custom_teamname]
      : '',
    'Test Case Status': fieldMaps.test_case_status[test.custom_test_case_status]
      ? ' ' + fieldMaps.test_case_status[test.custom_test_case_status]
      : '',
    Type: typeMap[test.type_id] || '',
  }));
}

// Get a unique filename (append _1, _2 if needed)
// This function will create a unique filename by appending _1, _2, etc. if the file already exists
function getUniqueFilename(folder, baseName, extension = 'xlsx') {
  let counter = 0;
  let filename;
  do {
    const suffix = counter === 0 ? '' : `_${counter}`;
    filename = path.join(folder, `${baseName}${suffix}.${extension}`);
    counter++;
  } while (fs.existsSync(filename));
  return filename;
}

// Main runner
// It will fetch test cases, format them, and export them to an Excel file
// It will create a folder named 'raw_data' if it doesn't exist
// It will create a unique filename based on the current date and the base name 'brink_pos'
// It will save the Excel file in the 'raw_data' folder
(async () => {
  try {
    const projectId = await getProjectIdByName('Brink POS');
    if (!projectId) return console.log('❌ Project not found.');

    const [testCases, caseFields, typeMap] = await Promise.all([
      getTestCases(projectId),
      getCaseFields(),
      getCaseTypes()
    ]);

    if (testCases.length === 0) return console.log('⚠️ No test cases found.');

    // Build field maps
    const getMap = name => buildFieldValueMap(caseFields.find(f => f.system_name === name));
    const fieldMaps = {
      automation_area: getMap('custom_automationarea'),
      automation_effort: getMap('custom_automationeffort'),
      automation_priority: getMap('custom_autopriority'),
      automation_status: getMap('custom_automation_status'),
      team_name: getMap('custom_teamname'),
      test_case_status: getMap('custom_test_case_status'),
    };    

    const filtered = formatTestCases(testCases, fieldMaps, typeMap);
    const folder = path.join(__dirname, 'raw_data');
    if (!fs.existsSync(folder)) {
      fs.mkdirSync(folder, { recursive: true });
    }
    
    const baseName = `brink_pos_${getTodayDateString()}`;
    const filename = getUniqueFilename(folder, baseName);
    await exportToExcel(filtered, filename);    
    
  } catch (err) {
    console.error('🚨 Script Error:', err.message);
  }
})();