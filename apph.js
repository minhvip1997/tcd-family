const fs = require('fs');
const pdf = require('pdf-parse');
const pdf2html = require('html-pdf'); // OR const puppeteer = require('puppeteer');

const pdfPath = 'ReportTCD2.pdf'; // Replace with your PDF file's path
const htmlPath = 'outputt.html'; // Output HTML file path

fs.readFile(pdfPath, (err, data) => {
  if (err) {
    console.error(`Error reading PDF file: ${err}`);
  } else {
    pdf(data).then(data => {
      const pdfText = data.text;
    //   console.log('pdfText', pdfText);
    // fs.writeFileSync(htmlPath, `<div>${pdfText}</div>`);
      console.log(`<div>${pdfText}</div>`);

      // You can further process the extracted text if needed

      // Create an HTML file from the extracted text
      pdf2html.create(pdfText, { format: 'A4' }).toFile(htmlPath, (err, res) => {
        if (err) {
          console.error(`Error converting to HTML: ${err}`);
        } else {
          console.log(`PDF converted to HTML in ${htmlPath}`);
        }
      });

    });
  }
});
