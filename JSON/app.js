/*By Jason Chalom 2014
  This will take some test json data and convert to xlsx using desired styling and formatting
  https://github.com/natergj/excel4node

  to use a stream and not write to file, just use buildExcelDataFromJson and push where-ever its the base-64 string
  always save as xlsx the other formats may not work correctly.

  Styles must have headings
  Styles all have arrays for each main object. This corresponds to specific worksheets
  Sub-object arrays sometimes correspond to headings, or image number or cell locations.

  custStyles are linked directly at the cell level

  NB TODO: color is Colour in styles.json perhaps change this to American?

  **There are some todos
*/

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
var styleObj = JSON.parse(fs.readFileSync(dir+"test-styles.json"));
var debug = true;

console.log("Starting Conversion....");
makeExcelDocument(reports, styleObj, file);

function makeExcelDocument(reports, styleObj, file){
  var wb = new xl.WorkBook();
  //print headings first
  var custStyles = getCustomStyles(styleObj, wb);
  if (isEmptyObject(styleObj)){
      throw("Styles Object is Empty."); //TODO JMC add try-catches
  }
  var worksheets = getWorksheets(wb, reports, styleObj);
  if (isEmptyObject(styleObj)){
    throw("Reports Object is Empty.");
  }
  //if worksheets or styles is null then cry
  log("------------Raw Data Done------------")

  //check if styles contains any style objects

  //run stylizer


  wb.write(file);
  console.log("file written.");
  //wb.write(file,res); //node response version
};


function getCustomStyles(styleObj, wb){
  var styles = [];
  //If there is not style then the data should just be pushed to the file as plain text.
  //This will only get the custStyles object
  for (var i=0; i<styleObj.data.custStyles.length; i++){
    var p = styleObj.data.custStyles[i].data;
    var prop;
    var k = 0;
    for(prop in p){
      log("prop: "+prop);
      if (!p.hasOwnProperty(prop)) {
        //The current property is not a direct property of p
        log("alert! "+prop + " -> " + p[prop]);
        continue;
      }
      log(prop+" : "+p[prop]);

      var style = wb.Style();

      if (prop===("bold")){
        if (p[prop])
          style.Font.Bold();
      }
      if (prop===("italics")){
        if (p[prop])
          style.Font.Italics();
      }
      if (prop===("underline")){
        if (p[prop])
          style.Font.Underline();
      }

      if (prop===("font.family"))
        style.Font.Family(p[prop]);
      if (prop===("colour"))
        style.Font.Color(p[prop]);
      if (prop===("size"))
        style.Font.Size(p[prop]);
      if (prop===("alignmentVertical"))
        style.Font.Alignment.Vertical(p[prop]);
      if (prop===("alignmentHorizontal"))
        style.Font.Alignment.Horizontal(p[prop]);
      if (prop===("wrapText"))
        style.Font.WrapText(p[prop]);
      //TODO JMC add more styles
    }
    styles.push(style);
  }
  log("------------Custom Styles Done------------")
  return styles;
};

function getWorksheets(wb, reports, styleObj){
  //check if null
  var worksheets = [];
  if (isEmptyObject(reports))
    return worksheets; //TODO JMC Error reporting

    for (var i=0; i<reports.length; i++){
      var ws = wb.WorkSheet(reports[i].name);
      log(reports[i].name);

      var p = styleObj.data.headings[i];
      var prop;
      var k = 0;
      for(prop in p){
        log("prop: "+prop);
        if (!p.hasOwnProperty(prop)) {
          //The current property is not a direct property of p
          log("alert! "+prop + " -> " + p[prop]);
          continue;
        }
        log(prop+" : "+p[prop]);
        if (i == 0){//j+2 to leave space for headings
          ws.Cell(1,i+k+1).String(""+p[prop]); //fix for r c going from 1,1
        }
        else{
          ws.Cell(1,i+k).String(""+p[prop]);
        }
        k++;
      }

      //insert data will be string for now. The typecasting will happen in styling
      for (var j=0; j<reports[i].data.length; j++){
        var p = reports[i].data[j];
        var prop; //property in p
        var k = 0;
        for (prop in p) {
          log("prop: "+prop);
          if (!p.hasOwnProperty(prop)) {
            //The current property is not a direct property of p
            log("alert! "+prop + " -> " + p[prop]);
            continue;
          }
          log(prop+" : "+p[prop]);
          if (i == 0){//j+2 to leave space for headings
            ws.Cell(j+2,i+k+1).String(""+p[prop]); //fix for r c going from 1,1
          }
          else{
            ws.Cell(j+2,i+k).String(""+p[prop]);
          }
          k++;
        }
      }
      worksheets.push(ws);
    }
    return worksheets;
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

function log(p, prop){
  if (debug)
    log(prop+" : "+p[prop]);
}
function log(str){
  if (debug)
    console.log("log: "+str);
};

function isEmptyObject(obj) {
  return !Object.keys(obj).length;
};
