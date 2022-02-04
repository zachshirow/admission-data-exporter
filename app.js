const XLSX = require('xlsx');




var workbook = XLSX.readFile('excel-data/164373266889.xlsx');

var second_sheet_name = workbook.SheetNames[1];
var address_of_cell = 'A2';


/* Get worksheet */
var worksheet = workbook.Sheets[second_sheet_name];

/* Find desired cell */
var desired_cell = worksheet[address_of_cell];

/* Get the value */
var desired_value = (desired_cell ? desired_cell.v : undefined);

console.log(desired_value)

