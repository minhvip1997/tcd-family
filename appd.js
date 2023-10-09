const fs = require('fs');
const MammothConvert = require('mammoth-convert');
const PDFDocument = require('pdfkit');

async function convertBufferToPdf() {
  // Read the DOCX buffer
  const docxBuffer = fs.readFileSync('ReportTCD2.docx');

  // Create a new PDF document
  const pdfDoc = new PDFDocument();
  const pdfBuffer = new Promise((resolve, reject) => {
    const pdfStream = fs.createWriteStream('output.pdf');

    pdfStream.on('error', reject);
    pdfDoc.on('error', reject);

    pdfDoc.pipe(pdfStream);

    // Use MammothConvert to convert DOCX buffer to PDF
    const mammothConvert = new MammothConvert();
    mammothConvert.convertToPdf(docxBuffer, pdfDoc);

    pdfDoc.end();

    pdfStream.on('finish', () => {
      resolve();
    });
  });

  // Wait for the PDF buffer to be generated
  await pdfBuffer;

  console.log('PDF conversion complete.');
}

convertBufferToPdf().catch((err) => {
  console.error('Error:', err);
});
