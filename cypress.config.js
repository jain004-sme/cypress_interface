const { defineConfig } = require("cypress");
const fs = require('fs')
const XLSX = require('xlsx')
const xlsx = require('xlsx')
const path = require('path');
const ExcelJs= require('exceljs');


module.exports = defineConfig({
  
  projectId: "2t6vp5",
  reporter: 'cypress-mochawesome-reporter',
  e2e: {
    defaultCommandTimeout: 40000,
    viewportHeight: 980,
    viewportWidth: 1480,
    setupNodeEvents(on) {
      
      require('cypress-mochawesome-reporter/plugin')(on);
 
      on('task',{
        writeExcelFile({fileName,data, sheetName}){
          const workbook = XLSX.utils.book_new()
          const worksheet = XLSX.utils.aoa_to_sheet(data)
          XLSX.utils.book_append_sheet(workbook,worksheet,sheetName)
          XLSX.writeFile(workbook,fileName)
         return null;
       
       
      },
      
      appendSheet({fileName,data,sheetName}){
           const workbook =fs.existsSync(fileName) ? XLSX.readFile(fileName) : XLSX.utils.book_new()
           const workSheet = XLSX.utils.aoa_to_sheet(data)
           XLSX.utils.book_append_sheet(workbook,workSheet,sheetName)
           XLSX.writeFile(workbook,fileName)
           return null;
      }
    })
   
      on('task', {
        
        readExcelFile() {
          // console.log(__dirname)
          const filePath = path.resolve(__dirname, "cypress/fixtures/cypress.xlsx");
          const workbook = xlsx.readFile(filePath);
          const sheetName = workbook.SheetNames[0];
          const workSheet = workbook.Sheets[sheetName];

          const jsonData = xlsx.utils.sheet_to_json(workSheet, { header: 1 });

          const result = {};
          jsonData.forEach(row => {
            console.log(row)
            if (row[0] && row[1]) {
              const key = row[0];
              const value = row[1];
              result[key] = value;
            }
          });
          return result;


        }
       
      })      
 
    }
  }

})

