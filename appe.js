const fs = require('fs');
// const docxBuffer = fs.readFileSync('TCD.docx');
// const docxBuffer = fs.readFileSync('ReportTCD2.docx');
// const docxBuffer = fs.readFileSync('ReportTCD1.docx');
const docxBuffer = fs.readFileSync('ReportTCD3.docx');
var mammoth = require("mammoth");
var {Window} = require("happy-dom");
// const officegen = require('officegen')('docx');
const htmlPdf = require('html-pdf');
mammoth.convertToHtml({buffer: docxBuffer})
    .then(function(result){
        var html = result.value; // The generated HTML

        const window = new Window({
            url: 'https://localhost:8080',
            width: 1024,
            height: 768
          });
        const document = window.document;
        document.body.innerHTML = html

        // console.log('htmlhtml', html);

        const images = document?.querySelectorAll('img');
        // const extractedHTML = Array.from(images).map((p) => p.outerHTML).join('');
        const extractedImages = Array.from(images);
        const imageRows = [];
        const imagesPerRow = 3
        for (const image of extractedImages) {
            image.style.maxWidth = '350px';
            image.style.marginLeft  = '10px';
            image.style.marginTop  = '5px';
          }
        for (let i = 0; i < extractedImages.length; i += imagesPerRow) {
            const rowImages = extractedImages.slice(i, i + imagesPerRow);
            const rowHtml = rowImages.map(img => img.outerHTML).join('');
            imageRows.push(rowHtml);
        }
        
        const extractedHTML = imageRows.map(row => `<div>${row}</div>`).join('');

        const pdfOptions = {
            format: 'A3', // You can customize the PDF format as needed
        };

        htmlPdf.create(extractedHTML, pdfOptions).toFile('ExtractedContent3.pdf', function (err, res) {
            if (err) {
                console.error(err);
            } else {
                console.log('Extracted content written to ExtractedContent3.pdf');
            }
        });
        
    })
    .catch(function(error) {
        console.error(error);
    });