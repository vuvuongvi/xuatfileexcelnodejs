/* require XLSX */
var excel = require('node-excel-export');
var XLSX = require('XLSX')
function datenum(v, date1904) {
    if(date1904) v+=1462;
    var epoch = Date.parse(v);
    return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}




function sheet_from_array_of_arrays(data, opts) {
    var ws = {};
    var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
    for(var R = 0; R != data.length; ++R) {
        for(var C = 0; C != data[R].length; ++C) {
            if(range.s.r > R) range.s.r = R;
            if(range.s.c > C) range.s.c = C;
            if(range.e.r < R) range.e.r = R;
            if(range.e.c < C) range.e.c = C;
            var cell = {v: data[R][C] };
            if(cell.v == null) continue;
            var cell_ref = XLSX.utils.encode_cell({c:C,r:R});

            if(typeof cell.v === 'number') cell.t = 'n';
            else if(typeof cell.v === 'boolean') cell.t = 'b';
            else if(cell.v instanceof Date) {
                cell.t = 'n'; cell.z = XLSX.SSF._table[14];
                cell.v = datenum(cell.v);
            }
            else cell.t = 's';
            ws[cell_ref] = cell;
        }
    }
    if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
    return ws;
}

var styles = {
    headerDark: {
        fill: {
            fgColor: {
                rgb: 'FF000000'
            }
        },
        font: {
            color: {
                rgb: 'FFFFFFFF'
            },
            sz: 14,
            bold: true,
            underline: true
        }
    },
    cellPink: {
        fill: {
            fgColor: {
                rgb: 'FFFFCCFF'
            }
        }
    },
    cellGreen: {
        fill: {
            fgColor: {
                rgb: 'FF00FF00'
            }
        }
    }
};



/* original data */
// var data = [[1,2,3],[true, false, null, "sheetjs"],["foo","bar",new Date("2014-02-19T14:30Z"), "0.3"], ["baz", null, "qux"]]
var ws_name = "SheetJS";
const heading = [
    [{value: 'a1', style: styles.headerDark}, {value: 'b1', style: styles.headerDark}, {value: 'c1', style: styles.headerDark}],
    ['a2', 'b2', 'c2'] // <-- It can be only values
];
const specification = {
    customer_name: { // <- the key should match the actual data key
        displayName: 'Customer', // <- Here you specify the column header
        headerStyle: styles.headerDark, // <- Header style
        cellStyle: function(value, row) { // <- style renderer function
            // if the status is 1 then color in green else color in red
            // Notice how we use another cell value to style the current one
            return (row.status_id == 1) ? styles.cellGreen : {fill: {fgColor: {rgb: 'FFFF0000'}}}; // <- Inline cell style is possible
        },
        width: 120 // <- width in pixels
    },
    status_id: {
        displayName: 'Status',
        headerStyle: styles.headerDark,
        cellFormat: function(value, row) { // <- Renderer function, you can access also any row.property
            return (value == 1) ? 'Active' : 'Inactive';
        },
        width: '10' // <- width in chars (when the number is passed as string)
    },
    note: {
        displayName: 'Description',
        headerStyle: styles.headerDark,
        cellStyle: styles.cellPink, // <- Cell style
        width: 220 // <- width in pixels
    }
}
const dataset = [
    {customer_name: 'IBM', status_id: 1, note: 'some note', misc: 'not shown'},
    {customer_name: 'HP', status_id: 0, note: 'some note'},
    {customer_name: 'MS', status_id: 0, note: 'some note', misc: 'not shown'}
]
const merges = [
    { start: { row: 1, column: 1 }, end: { row: 1, column: 10 } },
    { start: { row: 2, column: 1 }, end: { row: 2, column: 5 } },
    { start: { row: 2, column: 6 }, end: { row: 2, column: 10 } }
]
const report = excel.buildExport(
    [ // <- Notice that this is an array. Pass multiple sheets to create multi sheet report
        {
            name: 'Report', // <- Specify sheet name (optional)
            heading: heading, // <- Raw heading array (optional)
            merges: merges, // <- Merge cell ranges
            specification: specification, // <- Report specification
            data: dataset // <-- Report data
        }
    ]
);




function Workbook() {
    if(!(this instanceof Workbook)) return new Workbook();
    this.SheetNames = [];
    this.Sheets = {};
}

var wb = new Workbook()
// ws = sheet_from_array_of_arrays(data);
wp = sheet_from_array_of_arrays(dataset);

/* add worksheet to workbook */
wb.SheetNames.push(ws_name);
wb.Sheets[ws_name] = wp;
// wb.SheetNames.push(ws_name1);
wb.Sheets[ws_name] = wp;

/* write file */
XLSX.writeFile(wb, 'Report.xlsx');
// res.setHeader('Content-disposition', 'attachment; filename=Report.xlsx');
// res.send(report);