var xlsx = require('./xlsx.js');
var shim = require('./shim.js');
var ode = require('./ode.js');
var fs = require("fs");
var cp = require('./cpexcel.js')

var workbook = XLSX.readFile('test.xlsx');
var worksheet = workbook.Sheets[workbook.SheetName[0]];
console.log(worksheet["A1"]);