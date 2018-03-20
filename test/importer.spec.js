const { expect } = require('chai');
const Importer = require('../lib/importer');
describe('importer', function() {
    it('should throw an error for an unrecognised file type', () => {
        const someFileType = './path/to/docs.mov';
        expect(Importer.bind(null, someFileType)).to.throw();
    });

    it('should not throw for a docx file', () => {
        const someFileType = 'sample.docx';
        expect(Importer.bind(null, someFileType)).to.not.throw();
    });

    it('should return an array of paragraphs', async () => {
        const someFileType = 'sample.docx';
        const word = await Importer(someFileType);
        expect(word.getParagraphs()).to.be.an('array');
    });

    it('should return an array of paragraph text', async () => {
        const someFileType = 'sample.docx';
        const word = await Importer(someFileType);
        expect(word.getText()).to.be.an('array');
    });

    it('should return an array of paragraph ids', async () => {
        const someFileType = 'sample.docx';
        const word = await Importer(someFileType);
        expect(word.getParagraphIds()).to.be.an('array');
    });

    it('should return a paragraph by id', async () => {
        const someFileType = 'sample.docx';
        const word = await Importer(someFileType);
        expect(word.getParagraphIds()).to.be.an('array');
    });

    it('should return a text by paragraph id', async () => {
        const someFileType = 'sample.docx';
        const word = await Importer(someFileType);
        expect(word.getTextByParaId('53547906')).to.be.equal('What happened');
    });
});