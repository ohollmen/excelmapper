var xlsx = require('node-xlsx');
var fs = require('fs');

/* Construct Excel Spreadsheet Mapper.
* @param {string} filename - Filename for XLSX Spreadsheet
* @param {opts} opts - Options for ExcelMapper
*/
function ExcelMapper(filename, opts) {
   this.opts = opts || {};
   var xlsxcont = fs.readFileSync(filename); // try {...}
   if (opts.debug) {
     this.debug = 1;
     console.log(xlsxcont.length +  " B (of XLSX)");}
   var obj = this.book = xlsx.parse(xlsxcont);
   if (this.debug) { console.log(obj); }
   this.sh = null;
   this.cols = null;
}
/** List sheets of excel workbook.
* @return Array of Original sheet names.
*/
ExcelMapper.prototype.listsheets = function () {
  console.log("Listing sheets");
  return this.book.map(function (s) {return s.name});
}
/** Set Active sheet in Excel.
 * As a side effect:
 * - sets the column names to be used on AoO mapping to ones found
 * on the first row of spreadsheet.
 * - Sets the active sheet (for later data access)
 * @return Single sheet (in case low level access is needed)
 */
ExcelMapper.prototype.sheet = function (saddr) {
  if (!saddr) {throw "No Sheet Name or Index";}
  var si;
  this.sh = null; // Single sheet
  var obj = this.book; // Workbook (Whole spreadsheet, i.e. Set of sheets)
  if (saddr.match(/^\d+$/)) {
    si = parseInt(saddr);
    if (si < obj.length) {
      this.sh = obj[si];
      this.cols = this.cols_extract();
      return(this.sh);
    }
    throw "Could not access sheet by Index "+ si+ " (Number of sheets: "+obj.length+")";
  }
  else {
     sh = obj.filter(function (s) {return s.name === saddr;});
     if (!sh.length) {throw "Sheet by label '" + saddr + "' not in set of sheets";}
     this.sh = sh[0];
     this.cols = this.cols_extract();
     return(this.sh);
  }
  throw "Sheet "+ saddr +"not in boundaries of XLSX";
};

/** Extract and remove first row of sheet as columns.
 * Removes the first row completely from data.
 * @return columns as Array
 */
ExcelMapper.prototype.cols_extract = function ( opts) {
   opts = opts || {}; // self.opts
   var cols = this.sh.data.shift(); // self.cols;
   if (!Array.isArray(cols)) { throw "Cols Not an array"; }
   if (opts.nows) { // No whitespace
     cols.forEach(function (c) {
       if (c.match(/\s+/)) { throw "Whitespace in column name '" + c + "' !"; }
     });
   }
   return(cols);
};
/** Get existing column names or set columns explicitly.
 * In case of setting no check is done for the presence of columns
 * (You can perform a get call before setting if you want to do this.
 */
ExcelMapper.prototype.cols = function ( colnames) {
  if (colnames) {this.cols = colnames;}
  return(this.cols);
};

/** Convert (XLSX) AoA to more universal AoO based structure.
* Base conversion on column / property names spec in:
 * - attrs passed directly here
 * - the column names stored implicitly at the time of call to sheet().
 * @param {string} attrs - Column names to use in  conversion to AoO.
* @return data transformed into AoH format.
*/
ExcelMapper.prototype.to_aoh = function (attrs) {
   var d = this.sh.data;
   attrs = attrs || this.cols; // Default: Use cols extracted earlier
   if (!Array.isArray(d)) { throw "Sheet Data Not an array"; }
   if (!Array.isArray(attrs)) { throw "Property Names Not an array (Got: " + attrs+ ")"; }
   var out = d.map(function (r) {
      var i = 0;
      var e = {};
      // Strict ? Create check col. length(d) {}
      // if (d.length != attrs.length) {throw "Not corr. width";}
      attrs.forEach(function (an) {
         e[an] = r[i];
         i++;
      });
      return(e);
   });
   return(out);
}
/* TODO: Get rid of lines / rows
 * Count on this being so custom and on the other hand trivial in JS that
 * this method does not need to be there (at least first).
 * Also decide if this clenup is done on Ram AoO data or if this should be
 * a static method clening up the AoO
 * Tentatively strip:
 * - Empty records (i.e. {})
 * - Lines with single property that are heuristically not likely to be
 *   actual data lines
 */
function cleanup () {
  
}
module.exports.ExcelMapper = ExcelMapper;
