# workbook
Wrapper for js-xlsx providing convenient way to accumulate sheets, rows, styles.


# Install

`npm install protobi/workbook`
`npm install protobi/js-xlsx` (fork of SheetJS/js-xlsx testing extensions to save cell styles)


# Use

```js
var XLSX = require('xlsx'),
     Workbook = require('workbook')(XLSX)

var wb = new Workbook(XLSX);
```

Add rows to worksheets as arrays of row arrays.  Values may be elementary (e.g. number, string, date)
or Common Spreadsheet Format (CSF) objects

```js
workbook.addRowsToSheet("Main", [
  ["This is a merged cell"],
  [
    {"v": "Gray pattern", "s": "1"},
    {"v": "Blue ", "s": "2"},
    {"v": "Gray  ", "s": "3"}
    {"v": "Gold ", "s": "4"}
  ],
  [0.618033989, {"v": 0.618033989, "t": "n", "z": "0.00"}, {"v": 0.618033989, "z": "0.00%"},{"v": 0.618033989, "z": "0.00%","s":"4"}],
  [(new Date()).toLocaleString()]
]);
```

Merge cells by specifying start and end cell.  This method may be called repeatedly.  No error checking for overlapping rows is done.

```js
workbook.mergeCells("Main", {
  "s": {"c": 0, "r": 0 },
  "e": {"c": 2, "r": 0 }
});

```

Styles can be added.  The method `addStyle` returns an integer index that can be referenced in `cell.s`.

```js
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
```

Note that the first style specifies the default cell style and the second must be `gray125` for some reason
(see http://stackoverflow.com/questions/11116176/cell-styles-in-openxml-spreadsheet-spreadsheetml)

# Finalize and save

```js
workbook.finalize();
console.log(workbook);
var fs = require('fs');
var OUTFILE = '/tmp/wb.xlsx';
XLSX.writeFile(workbook, OUTFILE);
```