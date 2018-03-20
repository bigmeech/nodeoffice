const fs = require('fs');
const path = require('path');
const Promise = require('bluebird');

const WordPackage = require('../lib/packages/word');

function Importer (pathToFile, options) {
    const extension = path.extname(pathToFile);

    switch(extension) {
        case 'docx':
            return 'import docx file';
        default:
            throw new Error(`unknown file type: ${extension}`);
    }
}

exports.Importer = Importer;