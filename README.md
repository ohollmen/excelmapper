# excelmapper

excelmapper is a Excel trasformation utility library for converting / mapping XLSX files into Array of Objects (AoO)
by an column attribute definition.

The purpose of module is convert the (ode-xlsx) Excel spreadsheet internal format to the more usable
Array of Objects format, that is more likely to be useed in various contexts (as output of a REST service,
format to pass to a view template, or format from which to do bulk inserts to database using an ORM toolkit).

# Usage

    // opts from command line has "file" for filename, "sheet" for sheet addressing

    var emap = new emapper.ExcelMapper(opts.file, { debug: 1 });
    var sheets = emap.listsheets();
    console.log("List of sheets: " + JSON.stringify(sheets));
    console.log("Set Sheet to: " + opts.sheet);
    // Set "active" sheet (by name or 0-based sheet index)
    // Ok, we are storing return value and doing low level access to sheet here, but only
    // for testing, I promise :-)
    // This *must* be called to select the "active" sheet.
    var sheet = emap.sheet(opts.sheet);
    // Sheet has properties of node-xlsx single sheet ...
    if (sheet.data && sheet.name) {console.log("Sheet looks cool, has name and data.");}
    else {console.log("Something wrong with sheet ?!?");process.exit(1);}
    var aoo = emap.to_aoh();
    // Dump the AoO (Array of Objects)
    console.log(JSON.stringify(aoo, null, 2));
    // .. Put AoH into real use ...

DB Import Usage with Sequelize:

    // ... map to AoO
    var aoo = emap.to_aoh();
    // Bulk insert to DB
    products.bulkCreate(aoo, {validate: true})
    .then(function () {
       ...
    });

# Notes and Warnings

Data import, especially from spreadsheet is always a "dirty job" and you will most likely be doing
various massaging operations, cleanup and data transformations plus validation before ending up with
a viewable / imprtable format. These processing steps are not the focus of excelmapper. Try the good-ol'
Javascript and maybe underscore for these ops.

