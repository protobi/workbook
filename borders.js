var XLSX = require('xlsx');

var cellRef = XLSX.utils.encode_cell({c: 1, r: 1});
var ws = {
  '!cols': [{
    wpx: 40
  }, {
    wpx: 40
  }, {
    wpx: 40
  }],
  '!merges': [{s: {c: 1, r: 1}, e: {c: 2, r: 2}}],
  '!ref': XLSX.utils.encode_range({s: {c: 0, r: 0}, e: {c: 2, r: 2}})
};
ws[XLSX.utils.encode_cell({c: 1, r: 1})] =
    ws[XLSX.utils.encode_cell({c: 1, r: 2})] =
        ws[XLSX.utils.encode_cell({c: 2, r: 1})] =
            ws[XLSX.utils.encode_cell({c: 2, r: 2})] ={
  v: 'test',
  t: 's',
  s: {
    alignment: {horizontal: 'center', vertical: 'center'},
    border: {
      left: {style: 'thin', color: {auto: 1}},
      right: {style: 'thin', color: {auto: 1}},
      top: {style: 'thin', color: {auto: 1}},
      bottom: {style: 'thin', color: {auto: 1}}
    }
  }
}


var wb = {
  SheetNames: ['test'],
  Sheets: {test: ws}
};
XLSX.writeFile(wb, '/tmp/borders.xlsx');