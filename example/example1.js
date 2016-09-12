var fs = require('fs');
var XLSX = require('../');


var xCenter = XLSX.style({
    alignment: {horizontal: "center"}
});

var xTitle = XLSX.style({
    fill: {patternType: 'solid', fgColor: {rgb: '60497B'}},
    font: {color: {rgb: 'FFFFFF'}}
}, xCenter);

var xMoney = XLSX.style({
    numFmt: `_-£* #,##0.00_-;-£* #,##0.00_-;_-@_-`,
    alignment: {vertical: 'center'}
}, xCenter);

var rows = [
    ["SF9002", "Widget Blue",   xCenter(180),  xMoney(0.28)   ],
    ["SF9003", "Widget Green",  xCenter(180),  xMoney(0.28)   ],
    ["SF9004", "Widget Red",    xCenter(20),   xMoney(0.24)   ],
    ["SF9055", "Widget Gold",   xCenter(20),   xMoney(0.24)   ],
    ["SF9056", "Widget Silver", xCenter(null), xMoney(1337.29)]
];

var head = xTitle(["SKU","Description","Qty","Price"]);

var data = [];
data.push(head);
rows.forEach(r=>data.push(r));

const sheet1 = {
    name: "Prices",
    data: data,
    options:{
        '!cols': [
            {wch: 10}, // stock code
            {wch: 20}, // description
            {wch: 8}, // moq
            {wch: 10} // price
        ]
    }
};

let xlsxBuffer = XLSX.build([sheet1]);
fs.writeFileSync('example1.xlsx', xlsxBuffer);
