#!/usr/bin/nodejs
// 2015-06-17 Create minimal test for creating xlsx
var xlsx = require('node-xlsx');
var fs = require('fs');
var Getopt = require('node-getopt');
var emapper = require('./excelmapper');

var data = [
  ["col1","col2","col3"],
  [1,2,3],
  [true, false, null, 'sheetjs'],
  ['foo','bar', new Date('2014-02-19T14:30Z'), '0.3'],
  ['baz', null, 'qux']
];
var buffer = xlsx.build([{name: "mySheetName", data: data}]); // returns a buffer
console.log("Generated XLSX (in-mem):" + buffer.length + " B");

// console.log(process.argv);
var aspec = [
  ['f','file=ARG','Filename'],
  ['s','sheet=ARG','Sheed Index (0-based)'],
  ['t','table=ARG','Destination Table'],
  ['','strict','Strict Mode'],
];
var getopt = new Getopt(aspec);
var opt = getopt.parseSystem();
var opts = opt.options;
if (!opts.file) { throw "No File passed "; }
xl_write(buffer, opts.file); // Debug
var stats = fs.lstatSync(opts.file);
if (!stats) {throw "No File Found ";}
if (!stats.isFile) {throw "File is not a regular file";}

// 
if (!opts.sheet) { throw "No Sheet (Name or Index) passed (by --sheet)"; }
console.log(opt);
// getopt.showHelp();
// process.exit(0);
// fs.writeSync(fd, data[, position[, encoding]]);
// var filename = "/tmp/file.xlsx";
// var filename = opts.file;

// opts from command line has "file" for filename, "sheet" for sheet addressing

var emap = new emapper.ExcelMapper(opts.file, {debug: 1});
var sheets = emap.listsheets();
console.log("List of sheets: " + JSON.stringify(sheets));
console.log("Set Sheet to: " + opts.sheet);
// Ok, we are storing return value and doing low level access to sheet here, but only
// for testing, I promise :-)
var sheet = emap.sheet(opts.sheet);
// Sheet has properties of node-xlsx single sheet ...
if (sheet.data && sheet.name) { console.log("Sheet looks cool, has name and data."); }
else { console.log("Something wrong with sheet ?!?"); process.exit(1); }
// var aoo = emap.to_aoh(3); // ERROR, must pass Array if any
var aoo = emap.to_aoh();
// Start import
// var sheet = getsheet(obj, opts.sheet);
// console.log(JSON.stringify(sh, null, 2));
// console.log(JSON.stringify(cols, null, 2));
// console.log(JSON.stringify(obj, null, 2));
console.log(JSON.stringify(aoo, null, 2));

// OLD: Read back in (round trip)

// Create test spreadsheet to FS to use as basis for testing.
function xl_write (buffer, filename) {
  fs.writeFileSync(filename, buffer); // data,[, options]
  console.log("Wrote XLSX: " + filename);
}

