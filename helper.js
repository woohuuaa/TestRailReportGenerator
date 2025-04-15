const fs = require('fs');
const path = require('path');

// 🧠 Build dropdown value map for a single field
function buildFieldValueMap(field) {
  const map = {};
  const rawItems = field?.configs?.[0]?.options?.items || '';
  rawItems.split('\n').forEach(line => {
    const [key, value] = line.split(',').map(s => s.trim());
    if (key && value) map[key] = value;
  });
  return map;
}

// 🧠 Build all dropdown maps from caseFields
function buildAllFieldMaps(fields) {
  const maps = {};
  fields.forEach(field => {
    const isDropdown = field?.configs?.[0]?.options?.items;
    if (isDropdown) {
      maps[field.system_name] = buildFieldValueMap(field);
    }
  });
  return maps;
}

// 📅 Get today's date in yyyymmdd format
function getTodayDateString() {
  const now = new Date();
  return `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
}

// 🧾 Get a unique filename (append _1, _2 if needed)
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

module.exports = {
  buildFieldValueMap,
  buildAllFieldMaps,
  getTodayDateString,
  getUniqueFilename,
};

// Only run this block if called directly: `node helper.js`
if (require.main === module) {
    require('dotenv').config();
    const axios = require('axios');
  
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
  
    // Fetch and print all field maps
    async function printAllFieldMaps() {
      try {
        const res = await testrail.get('get_case_fields');
        const fields = res.data;
        const fieldMaps = buildAllFieldMaps(fields);
  
        console.log('📋 All Dropdown Field Maps:');
        console.log(JSON.stringify(fieldMaps, null, 2));
      } catch (err) {
        console.error('🚨 Failed to fetch fields:', err.response?.data || err.message);
      }
    }
  
    printAllFieldMaps();
  }
  
