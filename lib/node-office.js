/**
 * Author           -  Larry Eliemenye
 * Description      -  Read and extract Data from Office Files, Microsoft word, Powerpoint, Spreadsheet etc
 *
 * **/

 var NodeOffice = (function(){

    var fs          =  require("fs"),
        async       =  require("async"),
        xml2js      =  require("xml2js"),
        zip         =  require("adm-zip");
        path        =  require("path");

    var targetFile = null,
        EXTRACT_FOLDER = "./extracts",
        ext = ['.docx','.xlsx','.pptx'],
        RAW_XMLPATH = "./extracts/word/document.xml"

    var NodeOffice = function(file){
        //extract content of file, first test for open office extension
        fs.exists(file,function(exist){
            if(exist){
                if(HasSupportedExtension(file)){
                    var zipFile = new zip(file);
                    var entries = zipFile.getEntries()
                    entries.forEach(function(e){
                        console.log(e.entryName);
                    });
                    zipFile.extractAllTo(EXTRACT_FOLDER,true);
                    parseDocument(RAW_XMLPATH, function(){

                    });
                    //return api to access word documents
                    return{
                        getPage:getPage
                    }
                }
            }
            else{
                throw "cannot find file: "+file
            }
        })
    }

    //DOCX APIS
    var getPage =  function(number){

    }

    var parseDocument = function(rel,next){
        fs.exists(rel,function(exist){
            if(exist){

            }else throw "File not found at specified path: "+rel
        })
    }

    //Utlity functions
    var HasSupportedExtension = function(file){
        for(var key in ext){
            if(ext[key] === path.extname(file)){
                return true
            }else throw "Unsupported File format, Must be an Open office XML format of either .docx,.xlsx or .pptx"
        }
    }
    //returns API
     return{
         NodeOffice:NodeOffice
     }

})()

module.exports = NodeOffice;