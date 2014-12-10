//By TRex22 2014
//This will take some test json data and convert to xlsx using desired styling and formatting
//https://github.com/natergj/excel4node

//to use a stream and not write to file, just use buildExcelDataFromJson and push where-ever its the base-64 string
//always save as xlsx the other formats may not work correctly.

"use strict";
console.log("Excel converter");
//requirements
var xl = require('excel4node');

/*Test Data*/
var data = require('test-data.json');
var style = require('test-styles.json');

makeExcelDocument(data, style);

function makeExcelDocument(data, style){
  var worksheets = getWorksheets(data);
  //if worksheets is null then cry

  var styles = getStyles(style);
  //check if styles contains any style objects
}

function getWorksheets(){
  //check if null
  if (isEmptyObject(styleObj))
    return null;

}

function getStyles(styleObj){
  //check if null
  if (isEmptyObject(styleObj))
    return null; //If there is not style then the data should just be pushed to the file as plain text.
  var styles = [];

  //put in loop return array of styles
  var style = wb.Style();

  return null;
};

function isEmptyObject(obj) {
  return !Object.keys(obj).length;
};

function saveExcelFile(xlsBuf, outputFile) {
  // build file
  makeAFile(outputFile);
  fs.writeFileSync(outputFile, xlsBuf);
};

function makeAFile(file) {
  fs.writeFile(file, "", function(err) {
    if (err) {
      console.log(err);
    } else {
      console.log("The file was created!");
    }
  });
};
