// Support  CommonJS, AMD or independent user
(function (name, definition) {
  if (typeof define == 'function' && typeof define.amd == 'object') {
    define(definition);
  }
  else if (typeof module != 'undefined') module.exports = definition();
  else this[name] = definition();
}('Workbook', function () {

  // encode functions lifted from XLSX to eliminate dependency on XLSX so this can be serialized/deserialized to JSON
  function encode_row(row) { return "" + (row + 1); }
  function encode_col(col) { var s=""; for(++col; col; col=Math.floor((col-1)/26)) s = String.fromCharCode(((col-1)%26) + 65) + s; return s; }
  function encode_cell(cell) { return encode_col(cell.c) + encode_row(cell.r); }
  function encode_range(cs,ce) {
    if(ce === undefined || typeof ce === 'number') return encode_range(cs.s, cs.e);
    if(typeof cs !== 'string') cs = encode_cell(cs); if(typeof ce !== 'string') ce = encode_cell(ce);
    return cs == ce ? cs : cs + ":" + ce;
  }
  function datenum(v, date1904) {
    if (date1904) v += 1462;
    return (Date.parse(v) - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
  }

  return function Workbook() {

      var ranges = {}; //track  extent of each sheet
      var rows = {};   // accumulate data rows for each sheet

      return {
        SheetNames: [],
        Sheets: {},
        CellStyles: [],

        getSheet: function (sheetName) {
          if (!this.Sheets[sheetName]) {
            this.Sheets[sheetName] = {};

            this.SheetNames.push(sheetName);
            ranges[sheetName] = {s: {c: 10000000, r: 10000000}, e: {c: 0, r: 0 }};
            rows[sheetName] = [];
          }
          return this.Sheets[sheetName];
        },

        getSheetRows: function (sheetName) {
          var ws = this.getSheet(sheetName); // init if not previously called
          return rows[sheetName];
        },

        getRange: function (sheetName) {
          return ranges[sheetName];
        },

        getColumnProperties: function(sheetName) {
          var ws = this.getSheet(sheetName);
          if (!ws['!cols']) ws['!cols'] = [];
          return ws['!cols'];
        },
        setCell: function (sheetName, rowIdx, colIdx, val) {
          var rows = this.getSheetRows(sheetName);
          if (!rows[rowIdx]) {
            rows[rowIdx] = [];
          }
          rows[rowIdx][colIdx] = val;
          return this;
        },
        setColWidthChars: function(sheetName, colId, width) {
          var cps = this.getColumnProperties(sheetName);
          if (!cps[colId]) cps[colId] = {};
          cps[colId].wch = width;
          return this;
        },
        setColWidthPixels: function(sheetName, colId, width) {
          var cps = this.getColumnProperties(sheetName);
          if (!cps[colId]) cps[colId] = {};
          cps[colId].wpx = width;
          return this;
        },
        getCell: function (sheetName, rowIdx, colIdx, val) {
          var rows = this.getSheetRows(sheetName);
          if (!rows[rowIdx]) {
            rows[rowIdx] = [];
          }
          return rows[rowIdx][colIdx] = val;
        },


        addRowsToSheet: function (sheetName, rows) {
          var ws = this.getSheet(sheetName);
          var sheetRows = this.getSheetRows(sheetName);
          rows.forEach(function (row) {
            sheetRows.push(row);
          });
          return this;
        },

        finalize: function () {
          var self = this;
          this.SheetNames.forEach(function (sheetName) {
            self._finalizeSheet(sheetName);
          });
          return this;
        },

        // { s: { c: 0, r: 0}, e: {c: 2, r: 2}}
        mergeCells: function (sheetName, merge) {
          var sheet = this.getSheet(sheetName);
          sheet["!merges"] = sheet["!merges"] || [];
          sheet["!merges"].push(merge);
          return this;
        },

        // data is an array of row arrays
        // from https://gist.github.com/SheetJSDev/88a3ca3533adf389d13c
        _finalizeSheet: function (sheetName) {
          var ws = this.getSheet(sheetName), range = this.getRange(sheetName);

          var data = this.getSheetRows(sheetName) || [];

          for (var R = 0; R < data.length; ++R) {

            for (var C = 0; data[R] && C < data[R].length; ++C) {
              if (range.s.r > R) range.s.r = R;
              if (range.s.c > C) range.s.c = C;
              if (range.e.r < R) range.e.r = R;
              if (range.e.c < C) range.e.c = C;

              var cell = (typeof data[R][C] == 'object' ? data[R][C] : {v: data[R][C] });
              if (cell.v == null) continue;
              var cell_ref = encode_cell({c: C, r: R});

              if (typeof cell.v === 'number') cell.t = 'n';
              else if (typeof cell.v === 'boolean') cell.t = 'b';
              else if (cell.v instanceof Date) {
                cell.t = 'n';
                cell.z = 'm/d/yy';
                cell.v = datenum(cell.v);
              }
              else cell.t = 's';

              ws[cell_ref] = cell;
            }
          }
          if (range && range.s.c < 10000000) ws['!ref'] = encode_range(range);
          return ws;
        }
      }
    }
}));

