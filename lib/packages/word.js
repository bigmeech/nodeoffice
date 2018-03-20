const parser = require('xml2json');
const jsonpath = require('jsonpath');

/**
 *
 * @param resolve
 * @param entry
 * @returns {Promise.<*>}
 */
exports.retriever = async function (resolve, entry){
    if(entry.path === 'word/document.xml'){
        const content = await entry.buffer();
        const contentJson = parser.toJson(content.toString(), {
            object: true
        });
        return resolve(contentJson);
    }
    entry.autodrain();
};

/**
 *
 * @param docContent
 * @returns {{getParagraphs: (function())}}
 */
exports.parseDocument = function(docContent){
    return {
        getText: function() {
            return jsonpath.query(docContent, "$..['w:t']");
        },
        getParagraphs: function(){
            return jsonpath.query(docContent, "$..['w:p']");
        },
        getParagraphIds: function() {
            return jsonpath.query(docContent, "$..['w14:paraId']");
        },
        getTextByParaId: function(paraId){
            const para = this.getParagraphs()[0].find(para => {
                return (para['w14:paraId'] === paraId);
            });
            return jsonpath.query(para, "$['w:r']['w:t']")[0];
        }
    }
};