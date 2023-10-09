const { load } = require('@pspdfkit/nodejs');
const fs = require('node:fs');

async function convertToPDF() {
	const docx = fs.readFileSync('ReportTCD2.docx');
    console.log('docxdocx', docx);
	const instance = await load({
        document: docx,
	});
    console.log('instance', instance);
	// const buffer = await instance.exportPDF();
	// fs.writeFileSync('converted.pdf', Buffer.from(buffer));
	await instance.close();
}

convertToPDF();