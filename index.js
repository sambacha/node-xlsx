var XLSX = require('./lib/xlsx-style');
var helpers = require('./helpers');
var Workbook = require('./workbook');
var _ = require('lodash');
var StyledContent = require('./styledcontent');

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
        type: 'buffer',
        compression: 'DEFLATE'
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

function style(definition, parentStyler) {
    var styleDefinition = _.defaultsDeep(
        {},
        _.cloneDeep(definition),
        _.cloneDeep(parentStyler ? parentStyler.definition : {})
    );

    var styler;

    styler = function(item) {
        if (_.isArray(item)) {
            return item.map(styler);
        }
        return new StyledContent(item, styleDefinition);
    };
    styler.definition = styleDefinition;
    return styler;
}

module.exports = {
    parse: parse,
    build: build,
    style: style
}
