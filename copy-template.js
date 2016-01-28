var JSZip = require('jszip');
var fs = require('fs');
var cheerio = require('cheerio');

var INFILE = process.argv[2];
var OUTFILE = process.argv[3];

var infile = fs.readFileSync(INFILE);
var zip = new JSZip(infile);

var themeXml = zip.file("xl/theme/theme1.xml").asText();
var stylesXml = zip.file("xl/styles.xml").asText();

var outfile = fs.readFileSync(OUTFILE);
var outzip = new JSZip(outfile);
outzip.file("xs/theme/theme1.xml", themeXml);
outzip.file("xs/theme/styles.xml", stylesXml);

var buffer = outzip.generate({type: "nodebuffer"});
fs.writeFile(OUTFILE, buffer, function(err) { if (err) throw(err); });
