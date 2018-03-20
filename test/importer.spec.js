const { expect } = require('chai');
const NodeOffice = require('../lib/index');
describe('importer', function() {
    it('should throw an error for an unrecognised file type', () => {
        const someFileType = './path/to/docs.mov';
        expect(NodeOffice.Importer.bind(null, someFileType)).to.throw();
    })
});