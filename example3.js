var XLSX = require('../js-xlsx');
var workbook = { SheetNames: ['Sheet 1'], Sheets: { 'Sheet 1': {'!ref': 'A1:A1'}}};
workbook.Sheets['Sheet 1']['A1'] = {"v": "Hello Red Arial 24pt", "s": {font: {name: "Arial", sz: 24, color: {rgb: "FFFF0000"}}}}
XLSX.writeFile(workbook, '/tmp/wb.xlsx');

console.log(workbook);
var OUTFILE = '/tmp/wb.xlsx';
var OUTFILE1 = '/tmp/wb1.xlsx';

console.log("Results written to " + OUTFILE)
