var xl = require('excel4node');
var wb = new xl.Workbook({
    defaultFont: {
        size: 12,
        name: 'Calibri',
        color: 'FFFFFFFF'
    },
    dateFormat: 'm/d/yy hh:mm:ss',
});
var newDate = new Date('2015-01-01T00:00:00.0000Z');
xl.getExcelTS(newDate);
var ws = wb.addWorksheet('Sheet 1');
ws.column(2).setWidth(15);
ws.column(1).setWidth(15);
ws.column(4).setWidth(15);
ws.column(5).setWidth(15);
ws.column(7).setWidth(15);
ws.column(8).setWidth(15);
var style = wb.createStyle({
    font: {
        color: '#FF0800',
        size: 12,
        bold: true
    },
    dateFormat: 'm/d/yy hh:mm:ss',
    alignment: {
        wrapText: true,
        horizontal: 'center'
    },
    border: {
        left: {
            style: 'medium', //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
            color: '110000' // HTML style hex value
        }, right: {
            style: 'medium',
            color: '110000'
        },
        top: {
            style: 'medium',
            color: '110000'
        },
        bottom: {
            style: 'medium',
            color: '110000'
        },
    }

});

let data = [
    {
        telname: 'Thẻ Cào Viettel',
        price: 'Mệnh giá',
        amount: '10,000',
        serialname: 'Serial',
        serial: '123456789',
        codename: 'Mã thẻ',
        code: '987654321',
        expiration: 'Ngày Hết Hạn',
        expired: new Date().getTime()
    },
    {
        telname: 'Thẻ Cào Viettel',
        price: 'Mệnh giá',
        amount: '20,000',
        serialname: 'Serial',
        serial: '123456789',
        codename: 'Mã thẻ',
        code: '987654321',
        expiration: 'Ngày Hết Hạn',
        expired1: new Date().getTime()
    },
    {
        telname: 'Thẻ Cào Viettel',
        price: 'Mệnh giá',
        amount: '30,000',
        serialname: 'Serial',
        serial: '123456789',
        codename: 'Mã thẻ',
        code: '987654321',
        expiration: 'Ngày Hết Hạn',
        expired: new Date().getTime()
    },
    {
        telname: 'Thẻ Cào Viettel',
        price: 'Mệnh giá',
        amount: '40,000',
        serialname: 'Serial',
        serial: '123456789',
        codename: 'Mã thẻ',
        code: '987654321',
        expiration: 'Ngày Hết Hạn',
        expired: new Date().getTime()
    },
    {
        telname: 'Thẻ Cào Viettel',
        price: 'Mệnh giá',
        amount: '50,000',
        serialname: 'Serial',
        serial: '123456789',
        codename: 'Mã thẻ',
        code: '987654321',
        expiration: 'Ngày Hết Hạn',
        expired: new Date().getTime()
    },
    {
        telname: 'Thẻ Cào Viettel',
        price: 'Mệnh giá',
        amount: '60,000',
        serialname: 'Serial',
        serial: '123456789',
        codename: 'Mã thẻ',
        code: '987654321',
        expiration: 'Ngày Hết Hạn',
        expired1: new Date().getTime()
    },
    {
        telname: 'Thẻ Cào Viettel',
        price: 'Mệnh giá',
        amount: '70,000',
        serialname: 'Serial',
        serial: '123456789',
        codename: 'Mã thẻ',
        code: '987654321',
        expiration: 'Ngày Hết Hạn',
        expired: new Date().getTime()
    }, {
        telname: 'Thẻ Cào Viettel',
        price: 'Mệnh giá',
        amount: '80,000',
        serialname: 'Serial',
        serial: '123456789',
        codename: 'Mã thẻ',
        code: '987654321',
        expiration: 'Ngày Hết Hạn',
        expired: new Date().getTime()
    }
];
row = 1;
col = 1;
for (let i = 0; i < data.length; i++) {
    let card = data[i];
    console.log(card.amount);
    col = 3 * (i % 3) + 1;

    if(i % 3 == 0){
        row = (i/3)*6 + 1;
    }

    console.log(row);
    ws.cell(row, col, row, col + 1, true).string(card.telname).style(style);
    ws.cell(row + 1, col).string(card.price).style(style);
    ws.cell(row + 1, col + 1).string(card.amount).style(style);
    ws.cell(row + 2, col).string(card.serialname).style(style);
    ws.cell(row + 2, col + 1).string(card.serial).style(style);
    ws.cell(row + 3, col).string(card.codename).style(style);
    ws.cell(row + 3, col + 1).string(card.code).style(style);
    ws.cell(row + 4, col).string(card.expiration).style(style);
    ws.cell(row + 4, col + 1).date(new Date()).style(style);
    xl.getExcelTS(5, 2);
    xl.getExcelTS(11, 2);
}


// row = 1;
// col = 1;
// flag = 1;
// let count = 0;
// for (let i = 0; i < data.length; i++) {
//     let card = data[i];
//     col = 3 * (i % 3) + 1;
//     if (count < 3) {
//         ws.cell(1, col, 1, col + 1, true).string(card.telname).style(style);
//         ws.cell(2, col).string(card.price).style(style);
//         ws.cell(2, col + 1).string(card.amount).style(style);
//         ws.cell(3, col).string(card.serialname).style(style);
//         ws.cell(3, col + 1).string(card.serial).style(style);
//         ws.cell(4, col).string(card.codename).style(style);
//         ws.cell(4, col + 1).string(card.code).style(style);
//         ws.cell(5, col).string(card.expiration).style(style);
//         ws.cell(5, col + 1).date(new Date()).style(style);
//     } else {
//         ws.cell(col + 1, col).string(card.price).style(style);
//         ws.cell(col + 1, col + 1).string(card.amount).style(style);
//         ws.cell(col + 2, col).string(card.serialname).style(style);
//         ws.cell(col + 2, col + 1).string(card.serial).style(style);
//         ws.cell(col + 3, col).string(card.codename).style(style);
//         ws.cell(col + 3, col + 1).string(card.code).style(style);
//         ws.cell(col + 4, col).string(card.expiration).style(style);
//         ws.cell(col + 4, col + 1).date(new Date()).style(style);
//     }
//     xl.getExcelTS(5, 2);
//     xl.getExcelTS(11, 2);
//     count++;
// }
// for(let a = 0; a < data.length; a++){
//     let array_two = data[a];
//     col2 = 3 * (a % 3) + 1;
//     ws.cell(7 , col2, 7, col2 + 1, true).string(array_two.telname).style(style);
//     ws.cell(8, col2).string(array_two.price).style(style);
//     ws.cell(8, col2 +1).string(array_two.amount).style(style);
//     ws.cell(9, col2).string(array_two.serialname).style(style);
//     ws.cell(9, col2 +1).string(array_two.serial).style(style);
//     ws.cell(10, col2).string(array_two.codename).style(style);
//     ws.cell(10, col2 +1).string(array_two.code).style(style);
//     ws.cell(11, col2).string(array_two.expiration).style(style);
//     ws.cell(11, col2 +1).date(new Date()).style(style);
// }
// for(let b =0; b < data.length; b++){
//     let array_three = data[b];
//     col3 = 3 * ( b %3 ) + 1;
//     ws.cell(13, col3, 13, col3 +1, true).string(array_three.telname).style(style);
//     ws.cell(14, col3).string(array_three.price).style(style);
//     ws.cell(14, col3+1).string(array_three.amount).style(style);
//     ws.cell(15, col3).string(array_three.serialname).style(style);
//     ws.cell(15, col3+1).string(array_three.serial).style(style);
//     ws.cell(16, col3).string(array_three.codename).style(style);
//     ws.cell(16, col3+1).string(array_three.code).style(style);
//     ws.cell(17, col3).string(array_three.expiration).style(style);
//     ws.cell(17, col3+1).date(new Date()).style(style);
// }
//
//
//
// ws.cell(2,2).number(50.000).style(style);
// ws.cell(3,1).string('Serial:').style(style);
// ws.cell(3,2).number(123456789).style(style);
// ws.cell(4,1).string('Mã Thẻ:').style(style);
// ws.cell(4,2).number(987654321).style(style);
// ws.cell(5,1).string('Ngày Hết Hạn:').style(style);
// ws.cell(5,2).date(new Date()).style(style);
// xl.getExcelTS(5,2);
// ws.cell(1, 1, 1, 2, true).string('Thẻ Cào Viettel').style(style);
// ws.cell(2, 1).string('Mệnh Giá:').style(style);


wb.write('Excel.xlsx');