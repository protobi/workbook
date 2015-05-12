var XLSX = require('../js-xlsx');
var Workbook = require('./workbook');

///credit http://daveaddey.com/?p=40
function JSDateToExcelDate(inDate) {
  return 25569.0 + ((inDate.getTime() - (inDate.getTimezoneOffset() * 60 * 1000)) / (1000 * 60 * 60 * 24));
}

var workbook = new Workbook();
workbook.setCell('Main',0,0,{v: "Hello"});
workbook.setCell('Main',1,1,{v: "Hello"});
workbook.setCell('Main',4,2,{v: "Hello"});

workbook.finalize();
console.log(workbook);
var OUTFILE = '/tmp/wb.xlsx';
var OUTFILE1 = '/tmp/wb1.xlsx';
console.log(workbook)
XLSX.writeFile(workbook, OUTFILE);
console.log("Results written to " + OUTFILE)

var workbook1 = XLSX.readFile(OUTFILE, {cellStyles: true, cellNF: true});
XLSX.writeFile(workbook1, OUTFILE1);
console.log("Results written to " + OUTFILE1)
