/**
 * Created by laggie on 08/05/14.
 * Just a test for the Library
 */

var no = require("./lib/node-office");
new no.readFile("loremipsum.docx",function(err, result){
    if(err) throw err
    else console.log(result)
});