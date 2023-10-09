const fs = require('fs');
const mammoth = require('mammoth');
const PDFDocument = require('pdfkit');

async function convertDocxToPdf() {
  // Read the DOCX file as a buffer
  const docxBuffer = fs.readFileSync('ReportTCD2.docx');

  // Convert the DOCX buffer to plain text using Mammoth
  const { value } = await mammoth.extractRawText({ buffer: docxBuffer });

  // Create a new PDF document
  const pdfDoc = new PDFDocument();
  const pdfBuffer = new Promise((resolve, reject) => {
    const pdfStream = fs.createWriteStream('output.pdf');

    pdfStream.on('error', reject);
    pdfDoc.on('error', reject);

    pdfDoc.pipe(pdfStream);

    // Add the plain text content to the PDF
    pdfDoc.text(value);

    pdfDoc.end();

    pdfStream.on('finish', () => {
      resolve();
    });
  });

  // Wait for the PDF buffer to be generated
 const minh = await pdfBuffer;
 console.log('minh', minh);
//   console.log('PDF conversion complete.');
}

convertDocxToPdf().catch((err) => {
  console.error('Error:', err);
});
