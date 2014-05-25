/**
 * Created by laggie on 08/05/14.
 * Example Usage
 */

var NodeOffice = require("./lib/node-office");
NodeOffice.readFile("Loremipsum.docx", function (err, bodyObject) {
  /*var paras = bodyObject.getParagraphs();
  var runs = [];
  var content = ""
  //for each paragraph
  for (var paraIndex in paras) {
    var paragraph = paras[paraIndex];
    var runs = bodyObject.getRuns(paragraph);
    for (var runIndex in runs){
      var run = runs[runIndex];
      content += bodyObject.getRunContent(run)+"\n";
    }
  }
  console.log(content)*/

  var media = bodyObject.getMedia();
});


