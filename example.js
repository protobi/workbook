var XLSX = require('xlsx');
var Workbook = require('./workbook')(XLSX);

var workbook = new Workbook();
console.log(workbook)
workbook.addRowsToSheet("Main", [
  ["This is a submerged cell"],
  [
    {"v": "Bold", "s": {
      font: {name: 'Arial', sz: '18'},
      fill: { fgColor: { rgb: "FFFF0000"}},
      border: null,
      numFmt: null
    }},
    {"v": "Italic", "s": {
      font: {name: 'Arial', sz: '18'},
      fill: { fgColor: { rgb: "FF00FF00"}},
      border: null,
      numFmt: null
    }},
    {"v": "Bold Italic", "s": {
      font: {name: 'Arial', sz: '18'},
      fill: { fgColor: { rgb: "FF0000FF"}},
      border: null,
      numFmt: null
    }}
  ],
  [
    {"v": "red font", "s": {
      font: {name: 'Arial', sz: '18'},
      fill: { fgColor: { theme: "3", tint: "0.59999389629810485"}},
      border: null,
      numFmt: null
    }},
    {"v": "red fill", "s": {
      font: {name: 'Arial', sz: '18'},
      fill: { fgColor: { theme: "3", tint: "0.59999389629810485"}},
      border: null,
      numFmt: null
    }}
  ],
  [
    {"v": "Arial", "s": {
      font: {name: 'Arial', sz: '24'},
      fill: { fgColor: { theme: "0", tint: "-0.14999847407452621"}},
      border: null,
      numFmt: null
    }},
    {"v": "Arial 18pt", "s": {
      font: {name: 'Arial', sz: '24'},
      fill: { fgColor: { theme: "0", tint: "-0.14999847407452621"}},
      border: null,
      numFmt: null
    }}
  ],
  [ 0.618033989,
    {"v": 0.618033989, "s": {"numFmt": "0"}},
    {"v": 0.618033989, "s": {"numFmt": "0.00%"}},
    {"v": 0.618033989, "s": {"numFmt": 10 }}, // equivalent to above "0.00%"
    {"v": 0.618033989, "t": "n", "s": {"numFmt": "0.00", font: {name: 'Calibri', sz: '36'}}},
    {"v": 0.618033989, "t": "n", "s": {"numFmt": "0.0%", font: {name: 'Georgia', sz: '24'}, fill: { fgColor: { theme: "3", tint: "+0.3"}}}},
    {"v": 0.618033989, "t": "n", "s": {"numFmt": 44, font: {name: 'Avenir Book', sz: '12'}, fill: { fgColor: { rgb: "FFFFCC00"}}}}
  ],
  [
    {"v": 0.618033989, "s": {"numFmt": "0"}},
    {"v": 0.618033989, "s": {"numFmt": "0.00"}},
    {"v": 0.618033989, "s": {"numFmt": "0.00%"}},
    {"v": 0.618033989, "s": {"numFmt": "0%"}}
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


workbook.finalize();
console.log(workbook);
var fs = require('fs');
var OUTFILE = '/tmp/wb.xlsx';
XLSX.writeFile(workbook, OUTFILE);
console.log("Results written to "+OUTFILE)
