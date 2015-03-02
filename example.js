var XLSX = require('xlsx');
var Workbook = require('./workbook');

///credit http://daveaddey.com/?p=40
function JSDateToExcelDate(inDate) {
  return 25569.0 + ((inDate.getTime() - (inDate.getTimezoneOffset() * 60 * 1000)) / (1000 * 60 * 60 * 24));
}

var workbook = new Workbook(XLSX)
    .addRowsToSheet("Main", [
      ["This is a submerged cell"],
      [
        {"v": "Blank"},
        {"v": "Red", "s": {fill: { fgColor: { rgb: "FFFF0000"}}}},
        {"v": "Green", "s": {fill: { fgColor: { rgb: "FF00FF00"}}}},
        {"v": "Blue", "s": {fill: { fgColor: { rgb: "FF0000FF"}}}}
      ],
      [
        {"v": "Default"},
        {"v": "Arial", "s": {font: {name: "Arial", sz: 24}}},
        {"v": "Times New Roman", "s": {font: {name: "Times New Roman", sz: 16}}},
        {"v": "Courier New", "s": {font: {name: "Courier New", sz: 14}}}
      ],
      [
        0.618033989,
        {"v": 0.618033989},
        {"v": 0.618033989, "t": "n"},
        {"v": 0.618033989, "t": "n", "s": { "numFmt": "0.00%"}},
        {"v": 0.618033989, "t": "n", "s": { "numFmt": "0.00%"}, fill: { fgColor: { rgb: "FFFFCC00"}}}
      ],
      [
        {"v": 0.618033989, "t": "n", "s": { "numFmt": "0%"}},
        {"v": 0.618033989, "t": "n", "s": { "numFmt": "0.0%"}},
        {"v": 0.618033989, "t": "n", "s": { "numFmt": "0.00%"}},
        {"v": 0.618033989, "t": "n", "s": { "numFmt": "0.000%"}},
        {"v": 0.618033989, "t": "n", "s": { "numFmt": "0.0000%"}},
        {"v": 0, "t": "n", "s": { numFmt: "0.00%;\\(0.00%\\);\\-;@"}, fill: { fgColor: { rgb: "FFFFCC00"}}}
      ],
      [
        {v: (new Date()).toLocaleString()},
        {v: (new Date()), t: 'd'},
        {v: (new Date()),  s: {numFmt: 'd-mmm-yy'}}
      ]
      ,
      [
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
      ],
      [
        {v: 41684.35264774306, s: {numFmt: 'm/d/yy'}},
        {v: 41684.35264774306, s: {numFmt: 'd-mmm-yy'}},
        {v: 41684.35264774306, s: {numFmt: 'h:mm:ss AM/PM'}},
        {v: new Date(),  s: {numFmt: 'm/d/yy'}},
        {v: 42065.02247239584,  s: {numFmt: 'm/d/yy'}},
        {v: new Date(),  s: {numFmt: 'm/d/yy h:mm:ss AM/PM'}}
      ]
    ]).mergeCells("Main", {
      "s": {"c": 0, "r": 0 },
      "e": {"c": 2, "r": 0 }
    }).finalize();

var OUTFILE = '/tmp/wb.xlsx';
XLSX.writeFile(workbook, OUTFILE);
console.log("Results written to " + OUTFILE)
