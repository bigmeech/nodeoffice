/**
 * Created by laggie on 08/05/14.
 * Just a test for the Library
 */

var NodeOffice = require("./lib/node-office");
NodeOffice.readFile("loremipsum.docx",function(err,bodyObject){


    var paras = bodyObject.getParagraphs();
    console.log(paras.length);
    //console.log("using content")
});


