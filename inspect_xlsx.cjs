const XLSX = require('xlsx');

try {
  const workbook = XLSX.readFile('public/images/login/Data/Smple_DPR.xlsx');
  console.log("Sheet names in Smple_DPR.xlsx:", workbook.SheetNames);
  
  workbook.SheetNames.forEach(sheetName => {
    if (sheetName.toLowerCase().includes('maint') || sheetName.toLowerCase().includes('complain')) {
      console.log(`\n--- Sheet Content: ${sheetName} ---`);
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      rows.forEach((row, i) => {
        const rowStr = row.map(cell => cell !== undefined && cell !== null ? String(cell).trim() : '').join(' | ');
        console.log(`Row ${i}: ${rowStr}`);
      });
    }
  });
} catch (err) {
  console.error("Error reading file:", err);
}
