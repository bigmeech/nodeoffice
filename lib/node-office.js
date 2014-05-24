/**
 * Author           -  Larry Eliemenye
 * Description      -  Read and extract Data from Office Files, Microsoft word, Powerpoint, Spreadsheet etc
 *
 *
 * WordProcessingML Spec - http://officeopenxml.com/anatomyofOOXML.php
 * ====================================================================
 *  w:p = paragraph
 *  w:r = runs
 *
 *
 * **/

var NodeOffice = (function () {

  var fs = require("fs"),
      async = require("async"),
      xml2js = require("xml2js"),
      zip = require("adm-zip");
  path = require("path");
  parser = xml2js.Parser({xmlns: "w"});


  var xml_data = null,
      EXTRACT_FOLDER = "./extracts",
      ext = ['.docx', '.xlsx', '.pptx'],
      RAW_XMLPATH = "./extracts/word/document.xml",
      xml_json = null,
      xml_obj = null,
      xml_body = null,
      content = null,
      err = null;

  parser.addListener("end", function (result) {
    xml_data = result;
    xml_json = JSON.stringify(xml_data);
  })

  //reads file and returns of the file
  var readFile = function (file, next) {
    //extract content of file, first test for open office extension
    fs.exists(file, function (exist) {
      if (exist) {
        if (HasSupportedExtension(file)) {
          var zipFile = new zip(file);
          var entries = zipFile.getEntries()
          entries.forEach(function (e) {
            console.log(e.entryName);
          });
          zipFile.extractAllTo(EXTRACT_FOLDER, true);
          parseDocument(RAW_XMLPATH, function (data) {
            xml_obj = JSON.parse(data);
            xml_body = xml_obj["w:document"]["w:body"];
            next(err, getBodyObject);
          });
          return xml_body
        }
      }
      else {
        var err = new Error("cannot find file: " + file);
        next(err, getBodyObject);
      }
    })
  }

  //returns paragraphs(w:p) as an array
  var getParagraphs = function(){
    var body = xml_body[0]
    var paragraphs = [];
    for(var element in body){
      if(element === "w:p" && typeof(element) === "string"){
        paragraphs = body[element];
      }
    }
    return paragraphs;
  }

  //returns runs(w:r) as an array from which to get textual content.
  var getRuns = function(paragraph){
    var runs;
    for(var element in paragraph){
      if(element === "w:r" && typeof(element) == "string"){
        runs = paragraph[element];
      }
    }
    return runs
  }
  var getRunContent = function(run){
    var content = "";
    for (var key in run){
      if(key === "w:t" && typeof(key) === "string"){
        var contentArray = run[key];
        for(var textIndex in contentArray){
          content += contentArray[textIndex]._;
        }
      }
    }
    return content
  }
  var containsRichFormatting = function (para) {
    for (var key in para) {
      console.log(key);
    }
  }

  //parse content xml document
  var parseDocument = function (rel, next) {
    fs.exists(rel, function (exist) {
      if (exist) {
        fs.readFile(rel, function (err, data) {
          parser.parseString(data);
          next(JSON.stringify(xml_data));
        })
      } else throw "File not found at specified path: " + rel
    })
  }

  var trimWhitespace = function(){

  }

  //returns tables
  var getTableData = function()
  {
    //for()
  }

  var getContent = function(){

  }

  var hasTables = function(){
    return false
  }

  var hasMedia = function(){
    return false
  }


  //Utlity functions
  var HasSupportedExtension = function (file) {
    for (var key in ext) {
      if (ext[key] === path.extname(file)) {
        return true
      } else throw "Unsupported File format, Must be an Open office XML format of either .docx,.xlsx or .pptx"
    }
  }
  //API to return after a call to readfile
  var getBodyObject = {
    getParagraphs: getParagraphs,
    containRichFormatting: containsRichFormatting,
    getTableData:getTableData,
    getRuns:getRuns,
    getRunContent:getRunContent,
    hasMedia:hasMedia,
    hasTables:hasTables
  }
  //returns API
  return{
    readFile: readFile
  }

})()

module.exports = NodeOffice;