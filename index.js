'use strict'

// database config
let database = require("./database/config.js");

// required modules
const express = require("express");
const app = express();
const nodeSSPI = require('node-sspi');
const sql = require('mssql');
const excel = require('excel4node');

app.set('port', process.env.PORT || 3000);
app.use(express.static(__dirname + '/views')); // allows direct navigation to static files
app.use(require("body-parser").urlencoded({extended: true})); // parse form submissions

// requiring express-handlebars module and making Handlebars the template engine
let handlebars =  require("express-handlebars")
.create({ defaultLayout: "main"});
app.engine("handlebars", handlebars.engine);
app.set("view engine", "handlebars");

// windows authentication with node-sspi module
app.use(function (req, res, next) {
  var nodeSSPIObj = new nodeSSPI({
    retrieveGroups: true
  })
  nodeSSPIObj.authenticate(req, res, function(err){
    res.finished || next()
  })
})

// global variables
app.locals.currentYear = new Date().getFullYear(); // current year

// search page
app.get('/', function(req, res) { 
	
app.locals.currentUser = (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "");
	
  res.render('search', {title: "Search"}); 
});

// reports page
app.get('/reports', function(req,res) {
	
if(req.query.plant) {
	
	// return all plant names and numbers
new sql.ConnectionPool(database.config).connect().then(pool => {
    
let querycode = 'select * from tblConsumers where plant = ' + "'" + req.query.plant + "'" + ' and ReceiveDate > ' + "'" + "2020-08-01" + "'" + ' order by ReceiveDate desc';
     
return pool.query(querycode);
}).then(result => {   

	
	
	
	
	
	
	
	
	
	// output sql query into Excel file, create workbook and worksheet
var workbook = new excel.Workbook();
var worksheet = workbook.addWorksheet('Plant Report');

// create a reusable style
var style = workbook.createStyle({
  font: {
    color: '#000000',
    size: 11,
	bold: true
  },
});
	
	// create style 2
var style2 = workbook.createStyle({
  font: {
    color: '#000000',
    size: 10
  },
});
	
	// create style 3
var style3 = workbook.createStyle({
  font: {
    color: '#FF0800',
    size: 10,
	bold: true
  },
});

	// spreadsheet header
worksheet.cell(1,1).string('ProductName').style(style);
worksheet.cell(1,2).string('UPC').style(style);
worksheet.cell(1,3).string('Plant').style(style);
worksheet.cell(1,4).string('Customer ID').style(style);
worksheet.cell(1,5).string('Receive Date').style(style);
worksheet.cell(1,6).string('Report Code').style(style);
worksheet.cell(1,7).string('Sum').style(style);

	// loop through database object and parse out corresponding fields
		for(let i = 0, j = 2; i < result.recordset.length; i++, j++){
			
		// condition that makes cell red if value is greater than 500
			if(result.recordset[i].ReportCode < 3){
				worksheet.cell(j,6).string(result.recordset[i].ReportCode).style(style3);
			} else {
				worksheet.cell(j,6).string(result.recordset[i].ReportCode).style(style2);
			}
			
			worksheet.cell(j,1).string(result.recordset[i].ProductName).style(style2);
			worksheet.cell(j,2).string(result.recordset[i].UPC).style(style2);
			worksheet.cell(j,3).string(result.recordset[i].Plant).style(style2);
			worksheet.cell(j,4).number(result.recordset[i].CustomerID).style(style2);
			worksheet.cell(j,5).date(result.recordset[i].ReceiveDate).style(style2);
			worksheet.cell(j,7).formula("=sum(D:D)").style(style2);
	}
	
workbook.write('Report_for_plant_' + req.query.plant + '.xlsx', res);
	
}).catch(console.error);
    
	
	
	
	
	
	
	
	
	
	
	
} else {
	app.locals.currentUser = (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "");
	res.render('reports', {title: "Reports"}); 	
}
});


// define 404 handler
app.use(function(req,res) {
	
	app.locals.currentUser = (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "");
	
  res.render('404', {title: "404"}); 
});

app.listen(app.get('port'), function() {
 console.log('Node app has started at ' + new Date().toLocaleString() + ".");
});