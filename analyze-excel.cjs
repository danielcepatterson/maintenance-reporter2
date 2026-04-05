const XLSX = require('xlsx');

// Try reading with different options
const workbook = XLSX.readFile('/Users/danielpatterson/Downloads/excel_maintenance_work_order-3.xls', {
  cellStyles: true,
  cellDates: true,
  sheetStubs: true
});

console.log('Sheet names:', workbook.SheetNames);
console.log('');

// Try each sheet
workbook.SheetNames.forEach((sheetName, idx) => {
  console.log('=== Sheet:', sheetName, '===');
  const sheet = workbook.Sheets[sheetName];
  const range = sheet['!ref'];
  console.log('Range:', range);
  
  if (range) {
    const decoded = XLSX.utils.decode_range(range);
    console.log('Rows:', decoded.e.r + 1, 'Cols:', decoded.e.c + 1);
    
    // Convert to JSON to see data
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    console.log('JSON rows:', json.length);
    
    // Show first 5 rows
    for (let i = 0; i < Math.min(5, json.length); i++) {
      console.log('Row', i + 1, ':', json[i].slice(0, 35));
    }
  }
  console.log('');
});
