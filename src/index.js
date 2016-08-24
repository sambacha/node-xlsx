var XLSX = require('xlsx-style');
var helpers = require('./helpers');
var Workbook = require('./workbook');

function parse(mixed, options = {}) {
  const workSheet = XLSX[helpers.isString(mixed) ? 'readFile' : 'read'](mixed, options);
  return Object.keys(workSheet.Sheets).map(name => {
    const sheet = workSheet.Sheets[name];
    return {name, data: XLSX.utils.sheet_to_json(sheet, {header: 1, raw: true})};
  });
}

function build(worksheets, write_opts = {}) {
  const defaults = {
    bookType: 'xlsx',
    bookSST: false,
    type: 'binary'
  };
  const workBook = new Workbook();
  worksheets.forEach(worksheet => {
    const name = worksheet.name || 'Sheet';
    const data = helpers.buildSheetFromMatrix(worksheet.data || [], worksheet.options || {});
    workBook.SheetNames.push(name);
    workBook.Sheets[name] = data;
  });
  const excelData = XLSX.write(workBook, Object.assign({}, defaults, write_opts));
  return excelData instanceof Buffer ? excelData : new Buffer(excelData, 'binary');
}

module.exports = {
    parse: parse,
    build: build
}
