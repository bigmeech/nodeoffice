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
                        xml_body    = xml_obj["w:document"]["w:body"]; //get body element

                        //transverse the xml tree and fetch relivant content only
                        for(var key in xml_body){
                            for(var key1 in xml_body[key]["w:p"]){
                                for(var key2 in xml_body[key]["w:p"][key1]["w:r"]){ //
                                    for(var key3 in xml_body[key]["w:p"][key1]["w:r"][key2]["w:t"]){
                                        var data = xml_body[key]["w:p"][key1]["w:r"][key2]["w:t"][key3]["_"]
                                        if(data !== null && data !== undefined)
                                            content +="\n"+ data;
                                        //console.log(content);
                                    }
                                }
                            }
                        }
                        next(err,content);

                    });
                }
            }
            else{
                var err = new Error("cannot find file: "+file);
                next(err,content);
            }
        })
    }

    //DOCX APIS
    var getPage =  function(number){

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
    //returns API
     return{
         readFile:readFile
     }

})()

module.exports = NodeOffice;