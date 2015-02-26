var XLSX = require('xlsx');
var Workbook = require('./workbook')(XLSX);

var workbook = new Workbook();
console.log(workbook)
workbook.addRowsToSheet("Main", [
  ["This is a merged cell"],
  [
    {"v": "Bold", "s": "1"},
    {"v": "Italic", "s": "1"},
    {"v": "Bold Italic", "s": "1"}
  ],
  [
    {"v": "red font", "s": "2"},
    {"v": "red fill", "s": "2"}
  ],
  [
    {"v": "Arial", "s": "3"},
    {"v": "Arial 18pt", "s": "3"}
  ],
  [0.618033989, {"v": 0.618033989, "t": "n", "z": "0.00"}, {"v": 0.618033989, "z": "0.00%"},{"v": 0.618033989, "z": "0.00%","s":"4"}],
  [
    {"v": 0.618033989, "s": "1"},
    {"v": 0.618033989, "s": "2"},
    {"v": 0.618033989, "s": "3"},
    {"v": 0.618033989, "s": "4"}
  ],
  [
    {"f": "=SUM(A5,C5)"}
  ],
    [(new Date()).toLocaleString()]
]);

workbook.mergeCells("Main", {
  "s": {"c": 0, "r": 0 },
  "e": {"c": 2, "r": 0 }
});

workbook.addStyles([
  {
    font: {name: 'Arial', sz: '12'},
    fill: { fgColor: { patternType: "none"}},
    border: {},
    numFmt: null
  },
  {
    font: {name: 'Arial', sz: '18'},
    fill: { fgColor: { patternType: "gray125"}},
    border: null,
    numFmt: null
  },{
    font: {name: 'Arial', sz: '18'},
    fill: { fgColor: { theme: "3", tint: "0.59999389629810485"}},
    border: null,
    numFmt: null
  },{
    font: {name: 'Arial', sz: '24'},
    fill: { fgColor: { theme: "0", tint: "-0.14999847407452621"}},
    border: null,
    numFmt: null
  },{
    font: {name: 'Arial', sz: '36'},
    fill: { fgColor: { rgb: "FFFFCC00"}},
    border: null,
    numFmt: null
  }
])

workbook.finalize();
console.log(workbook);
var fs = require('fs');
var OUTFILE = '/tmp/wb.xlsx';
XLSX.writeFile(workbook, OUTFILE);
