const fs = require('fs');
const pandoc = require('pandoc');

async function convertDocxToPdf() {
  // Read the DOCX file as a buffer
  const docxBuffer = fs.readFileSync('ReportTCD2.docx');

  // Convert the DOCX buffer directly to PDF
  const pdfBuffer = await pandoc(docxBuffer, '-f docx -t pdf');

  // Save the PDF buffer to a file (optional)
  fs.writeFileSync('output.pdf', pdfBuffer);

  console.log('PDF conversion complete.');
}

convertDocxToPdf().catch((err) => {
  console.error('Error:', err);
});
