const { app, BrowserWindow, ipcMain, shell } = require("electron");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

const createWindow = () => {
  const win = new BrowserWindow({
    width: 800,
    height: 600,
    center: true,
    autoHideMenuBar: true,
    webPreferences: {
      worldSafeExecuteJavascript: true,
      preload: path.join(__dirname, "preload.js"),
    },
  });

  win.loadFile("index.html");
};

app.whenReady().then(() => {
  createWindow();
});

async function processExcelWorkbook(excelPath) {
  const workbook = new ExcelJS.Workbook();

  try {
    await workbook.xlsx.readFile(excelPath);

    let sqlStatements = "USE crlauncher_data;";

    // Loop through each sheet in the workbook
    workbook.eachSheet((worksheet, sheetId) => {
      const tableName = worksheet.name;
      const columnNames = [];
      const records = [];

      // Iterate through row  to get column names
      const row1 = worksheet.getRow(1);
      row1.eachCell({ includeEmpty: false }, (cell) => {
        columnNames.push(cell.value);
      });

      // Iterate through the remaining rows for records
      for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber++) {
        const currentRow = worksheet.getRow(rowNumber);
        const record = {};

        row1.eachCell({ includeEmpty: false }, (cell, colNumber) => {
          record[columnNames[colNumber - 1]] =
            currentRow.getCell(colNumber).value;
        });

        records.push(record);
      }
      sqlStatements += `TRUNCATE TABLE \`${tableName}\`;`;
      sqlStatements += `INSERT INTO \`${tableName}\` (${columnNames
        .map((columnName) => `\`${columnName}\``)
        .join(", ")})`;
      sqlStatements += ` VALUES `;
      records.forEach((record, index) => {
        const values = columnNames
          .map((columnName) => {
            if (columnName === "created_at") {
              return "CURRENT_TIMESTAMP()";
            }

            if (columnName === "updated_at") {
              return "CURRENT_TIMESTAMP() ON UPDATE CURRENT_TIMESTAMP()";
            }
            return record[columnName] == "none" || record[columnName] == null
              ? `${null}`
              : `'${record[columnName]}'`;
          })
          .join(", ");
        sqlStatements += `(${values})`;
        if (index !== records.length - 1) {
          sqlStatements += ",";
        }
        sqlStatements += `\n`;
      });
      sqlStatements += ";";
      sqlStatements += `\n`;
    });

    // Write the SQL statements to a file
    fs.writeFileSync("data.sql", sqlStatements);
    console.log("SQL file generated successfully.");
  } catch (error) {
    console.error("Error reading the Excel file:", error);
  }
}

ipcMain.on("generate-sql", async (event, excelData) => {
  await processExcelWorkbook(excelData);
  event.reply("sql-generation-completed");
  shell.showItemInFolder(
    "C:/Users/user/Desktop/Data Exporter/dist/win-unpacked"
  );
  shell.openPath("data.sql");
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
