var http = require('http');
var fs = require('fs');
var excel = require('xlsx.js');

var server = http.createServer(function (req, res) {
    var filename = "test.xlsx";
    var wb = excel.readFile(filename);
    var first_sheet_name = workbook.SheetNames[0];
    var address_of_cell = 'A1';

/* Get worksheet */
    var worksheet = workbook.Sheets[first_sheet_name];

/* Find desired cell */
    var desired_cell = worksheet[address_of_cell];

/* Get the value */
    var desired_value = desired_cell.v;

    console.log(desired_value);
}); 
    
server.listen(9090);
console.log('server running...')