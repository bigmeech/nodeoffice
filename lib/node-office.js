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
        parser      =  xml2js.Parser({xmlns:"w"});


    var xml_data        = null,
        EXTRACT_FOLDER  = "./extracts",
        ext             = ['.docx','.xlsx','.pptx'],
        RAW_XMLPATH     = "./extracts/word/document.xml",
        xml_json        = null,
        xml_obj         = null,
        xml_body        = null,
        content         = null,
        err             = null;

    parser.addListener("end",function(result){
        xml_data = result;
        xml_json = JSON.stringify(xml_data);
    })

    //reads file and returns of the file
    var readFile = function(file,next){
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
                    parseDocument(RAW_XMLPATH, function(data){
                        xml_obj     = JSON.parse(data);
                        xml_body    = xml_obj["w:document"]["w:body"];
                        next(err,getBodyObject);
                    });
                    return xml_body
                }
            }
            else{
                var err = new Error("cannot find file: "+file);
                next(err,getBodyObject);
            }
        })
    }

    //DOCX APIS
    var getParagraphs=  function(){
        //returns a list of paragraphs as an Array of paragraph "w:p" objects
        var para = []
        for(var key in xml_body){
            for(var key1 in xml_body[key]["w:p"]);{
                for(var key2 in xml_body[key]["w:p"][key1]){
                    for(var key3 in xml_body[key]["w:p"][key1][key2]){
                        for(var key4 in xml_body[key]["w:p"][key1][key2][key3]){
                            if(typeof xml_body[key]["w:p"][key1][key2][key3]["w:t"] === 'object' && key4 === "w:t"){
                                //console.log(key4)
                                for(var key5 in xml_body[key]["w:p"][key1][key2][key3]["w:t"]){
                                    for( var key6 in xml_body[key]["w:p"][key1][key2][key3]["w:t"][key5]){
                                        var content = xml_body[key]["w:p"][key1][key2][key3]["w:t"][key5]["_"];
                                        if(content !== undefined
                                            && content !== ". "
                                            && content !== ", "
                                            && content !== '.'){
                                            para.push(xml_body[key]["w:p"][key1][key2][key3]["w:t"][key5]["_"])
                                        }
                                    }
                                }

                            }
                        }

                    }
                }

            }
        }
        return para;
    }

    var containsRichFormatting = function(para){
        for (var key in para){
            console.log(key);
        }
    }

    var getFullTextContent = function(){
        console.log("Getting Content from xml body")
        for(var key in xml_body){
            for(var key1 in xml_body[key]["w:p"]){
                for(var key2 in xml_body[key]["w:p"][key1]["w:r"]){ //
                    for(var key3 in xml_body[key]["w:p"][key1]["w:r"][key2]["w:t"]){
                        var data = xml_body[key]["w:p"][key1]["w:r"][key2]["w:t"][key3]["_"]
                        if(data !== null && data !== undefined)
                            content += data !== undefined ? data: ""; //if content somehow contains undefined even after check, turn it to empty strings
                    }
                }
            }
        }
        return content;
    }

    //parse content xml document
    var parseDocument = function(rel,next){
        fs.exists(rel,function(exist){
            if(exist){
                fs.readFile(rel,function(err,data){
                    parser.parseString(data);
                    next(JSON.stringify(xml_data));
                })
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
    //API to return after a call to readfile
    var getBodyObject = {
        getFullTextContent:getFullTextContent,
        getParagraphs:getParagraphs,
        containRichFormatting:containsRichFormatting
    }
    //returns API
     return{
         readFile:readFile
     }

})()

module.exports = NodeOffice;