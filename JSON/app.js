//By TRex22 2014
//This will take some test json data and convert to xlsx using desired styling and formatting
//https://github.com/natergj/excel4node

//to use a stream and not write to file, just use buildExcelDataFromJson and push where-ever its the base-64 string
//always save as xlsx the other formats may not work correctly.

"use strict";
console.log("Excel converter");
//requirements
var xl = require('excel4node');
var fs = require('fs'); //TODO JMC remove when done

/*Save directory for files*/
var dir = "data/";
var filename = "reports.xlsx";
var file = dir+filename;

/*Test Data*/
var reports = JSON.parse(fs.readFileSync(dir+"test-data.json"));
var style = JSON.parse(fs.readFileSync(dir+"test-styles.json"));
var debug = true;

makeExcelDocument(reports, style, file);

function makeExcelDocument(reports, style, file){
  var wb = new xl.WorkBook();

  var worksheets = getWorksheets(reports, wb);
  //if worksheets is null then cry

  var styles = getStyles(style, wb);
  //check if styles contains any style objects

  wb.write(file);
  log("file written.");
  //wb.write(file,res); //node response version
};

function getWorksheets(reports, wb){
  //check if null
  var worksheets = [];
  if (isEmptyObject(reports))
    return worksheets; //TODO JMC Error reporting

  for (var i=0; i<reports.length; i++){
    var ws = wb.WorkSheet(reports[i].name);
    log(reports[i].name);

    //insert data will be string for now. The typecasting will happen in styling
    log("checking data loop: "+reports[i].data.length);
    for (var j=0; j<reports[i].data.length; j++){
      var k = 0;
      var p = reports[i].data[j];
      for (prop in p) {
        if (!p.hasOwnProperty(prop)) {
          //The current property is not a direct property of p
          continue;
        }
        log(prop+" : "+p[k]);
        k++;
      }
    }
  }

};

function getStyles(styleObj, wb){
  //check if null
  var styles = [];
  if (isEmptyObject(styleObj))
    return styles; //If there is not style then the data should just be pushed to the file as plain text.

  //put in loop return array of styles
  var style = wb.Style();

  return styles;
};

function isEmptyObject(obj) {
  return !Object.keys(obj).length;
};

void function saveExcelFile(xlsBuf, outputFile) {
  // build file
  makeAFile(outputFile);
  fs.writeFileSync(outputFile, xlsBuf);
};

void function makeAFile(file) {
  fs.writeFile(file, "", function(err) {
    if (err) {
      console.log(err);
    } else {
      console.log("The file was created!");
    }
  });
};

function log(str){
  if (debug == true)
    console.log("log: "+str);
};
