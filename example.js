var XLSX = require('xlsx');
var Workbook = require('./workbook')(XLSX);

var workbook = new Workbook()
    .addRowsToSheet("Main", [
      ["This is a merged cell"],
      [ // fill colors
        {"v": "Blank"},
        {"v": "Red", "s": {fill: { fgColor: { rgb: "FFFF0000"}}}},
        {"v": "Green", "s": {fill: { fgColor: { rgb: "FF00FF00"}}}},
        {"v": "Blue", "s": {fill: { fgColor: { rgb: "FF0000FF"}}}}
      ],
      [ // fonts
        {"v": "Default"},
        {"v": "Arial", "s": {font: {name: "Arial", sz: 24}}},
        {"v": "Times New Roman", "s": {font: {name: "Times New Roman", sz: 16}}},
        {"v": "Courier New", "s": {font: {name: "Courier New", sz: 14}}}
      ],
      [ // built in formats
        0.618033989,
        {"v": 0.618033989},
        {"v": 0.618033989, "t": "n"},
        {"v": 0.618033989, "t": "n", "s": { "numFmt": "0.00%"}},
        {"v": 0.618033989, "t": "n", "s": { "numFmt": "0.00%"}, fill: { fgColor: { rgb: "FFFFCC00"}}}
      ],
      [ // cusotm formats
        {"v": 0.618033989, "t": "n", "s": { "numFmt": "0.00%;-0.00%;-;@"}, fill: { fgColor: { rgb: "FFFFCC00"}}},
        {"v": -0.618033989, "t": "n", "s": { "numFmt": "0.00%;-0.00%;-;@"}, fill: { fgColor: { rgb: "FFFFCC00"}}},
        {"v": 0, "t": "n", "s": { "numFmt": "0.00%;-0.00%;-;@"}, fill: { fgColor: { rgb: "FFFFCC00"}}},
        {"v": "n/a", "t": "n", "s": { "numFmt": "0.00%;-0.00%;-;@"}, fill: { fgColor: { rgb: "FFFFCC00"}}}
      ],
      [ // alignment
        {v: "left", "s": { alignment: {horizontal: "left"}}},
        {v: "left", "s": { alignment: {horizontal: "center"}}},
        {v: "left", "s": { alignment: {horizontal: "right"}}}
      ],[
        {v: "vertical", "s": { alignment: {vertical: "top"}}},
        {v: "vertical", "s": { alignment: {vertical: "center"}}},
        {v: "vertical", "s": { alignment: {vertical: "bottom"}}}
      ],[
        {v: "indent", "s": { alignment: {indent: "1"}}},
        {v: "indent", "s": { alignment: {indent: "2"}}},
        {v: "indent", "s": { alignment: {indent: "3"}}}
      ],
      [{
        v: "In publishing and graphic design, lorem ipsum is a filler text commonly used to demonstrate the graphic elements of a document or visual presentation. ",
        s: { alignment: { wrapText: 1, alignment: 'right', vertical: 'center', indent: 1}}
      }
      ]
    ]).mergeCells("Main", {
      "s": {"c": 0, "r": 0 },
      "e": {"c": 2, "r": 0 }
    }).finalize();

var OUTFILE = '/tmp/wb.xlsx';
XLSX.writeFile(workbook, OUTFILE);
console.log("Results written to " + OUTFILE)
