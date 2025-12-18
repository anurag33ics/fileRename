const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const MAIN_FOLDER = path.join(__dirname);
const OUTPUT_EXCEL = path.join(__dirname, "file-records.xlsx");

let records = [];
let paintingNumber = 1;
let srNo = 1;

// Read main folder
const folders = fs.readdirSync(MAIN_FOLDER, { withFileTypes: true })
  .filter(dirent => dirent.isDirectory())
  .map(dirent => dirent.name);

folders.forEach(folderName => {
  const folderPath = path.join(MAIN_FOLDER, folderName);

  const files = fs.readdirSync(folderPath);

  files.forEach(file => {
    const ext = path.extname(file).toLowerCase();

    if ([".jpg", ".jpeg", ".png", ".pdf"].includes(ext)) {
      const oldPath = path.join(folderPath, file);
      const newFileName = `${paintingNumber}${ext}`;
      const newPath = path.join(MAIN_FOLDER, newFileName);

      // Copy file to main folder
      fs.copyFileSync(oldPath, newPath);

      // Add Excel record
      records.push({
        "Sr No": srNo++,
        "Folder Name": folderName,
        "Original File Name": file,
        "Painting Number": paintingNumber
      });

      paintingNumber++;
    }
  });
});

// Create Excel file
const worksheet = XLSX.utils.json_to_sheet(records);
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, "Records");

XLSX.writeFile(workbook, OUTPUT_EXCEL);

console.log("✅ Files copied & Excel created successfully!");
