const path = require('path');
const fs = require('fs').promises;
const libre = require('libreoffice-convert');
libre.convertAsync = require('util').promisify(libre.convert);
async function main() {
    const ext = 'pdf'; // Output extension.
	const inputPath = path.join(__dirname, '/ReportTCD2.docx');
	const outputPath = path.join(__dirname, `/ReportTCD2.${ext}`);
    const convertOptions = {
        tmpDir: 'E:\react\testnode2',
      };
	// Read the input file.
	const docxBuf = await fs.readFile(inputPath);
	const pdfBuf = await fs.readFile(outputPath);
    // let pdfBuf = await libre.convertAsync(docxBuf, ext, convertOptions);
    console.log('docxBuf',docxBuf)
    console.log('pdfBuf',pdfBuf)
    // console.log('pdfBuf',pdfBuf)
}
main()