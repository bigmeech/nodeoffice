const fs = require('fs');
const path = require('path');
const Promise = require('bluebird');
const unzipper = require('unzipper');
const etl = require('etl');

/**
 *
 * @param ext
 * @returns {object | null}
 */
function getPackageByExtension(ext){
    switch (ext){
        case '.docx':
            return require('./packages/word');
        default:
            return null;
    }
}

/**
 *
 * @param pathToFile
 * @param options
 * @returns {Promise.<TResult>}
 * @constructor
 */
function Importer (pathToFile, options) {
    const extension = path.extname(pathToFile);
    if(!['.docx', '.xlsx', '.pptx'].includes(extension)) {
        throw new Error('Unrecognized file type');
    }

    const { retriever, parseDocument } = getPackageByExtension(extension);
    return readOfficeFile(pathToFile, retriever)
        .then(doc => parseDocument(doc));
}

/**
 *
 * @param filePath
 */
function existsAsync (filePath) {
    return new Promise((resolve, reject) => {
        fs.exists(filePath, (exists) => resolve(exists));
    })
}

/**
 *
 * @param filePath
 * @returns {Promise.<TResult>}
 */
async function readOfficeFile (filePath, retriever) {
    return existsAsync(filePath).then(exists => {
        if(exists) {
            return new Promise((resolve, reject) => {
                fs.createReadStream(filePath)
                    .pipe(unzipper.Parse())
                    .pipe(etl.map(retriever.bind(null, resolve)))
                    .on('error', (err) => reject(err));
            });
        }
    }).catch(e => { throw e })
}

module.exports = Importer;