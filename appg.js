const fs = require('fs');
const mammoth = require("mammoth");
const { Window } = require("happy-dom");
const htmlPdf = require('html-pdf');

const docxBuffer = fs.readFileSync('TCD.docx');
// const docxBuffer = fs.readFileSync('ReportTCD2.docx');
// const docxBuffer = fs.readFileSync('ReportTCD1.docx');
// const docxBuffer = fs.readFileSync('ReportTCD3.docx');

mammoth.convertToHtml({ buffer: docxBuffer })
    .then(function(result) {
        var html = result.value; // The generated HTML

        const window = new Window({
            url: 'https://localhost:8080',
            width: 1024,
            height: 768
        });
        const document = window.document;
        document.body.innerHTML = html;

        // console.log('htmlhtml', html);

        const images = document?.querySelectorAll('img');
        const nonImageContent = document.body.cloneNode(true);

        // Remove all images from the non-image content
        images.forEach(img => nonImageContent.querySelector(`img[src="${img.getAttribute('src')}"]`).remove());

        const extractedImages = Array.from(images);
        const imagesPerRow = 3;

        // Calculate the number of rows needed for images
        const numRows = Math.ceil(extractedImages.length / imagesPerRow);

        const imageRows = [];
        for (const image of extractedImages) {
            image.style.maxWidth = 350;
          }
        for (let i = 0; i < numRows; i++) {
            const startIndex = i * imagesPerRow;
            const endIndex = Math.min(startIndex + imagesPerRow, extractedImages.length);
            const rowImages = extractedImages.slice(startIndex, endIndex);
            const rowHtml = rowImages.map(img => img.outerHTML).join('');
            imageRows.push(rowHtml);
        }

        const nonImageContentHtml = nonImageContent.innerHTML;
        const extractedHTML = imageRows.map(row => `<div>${row}</div>`).join('') + nonImageContentHtml;

        const pdfOptions = {
            format: 'A3', // You can customize the PDF format as needed
        };

        htmlPdf.create(extractedHTML, pdfOptions).toFile('ExtractedContent3.pdf', function(err, res) {
            if (err) {
                console.error(err);
            } else {
                console.log('Extracted content with images and stable content written to ExtractedContent3.pdf');
            }
        });

    })
    .catch(function(error) {
        console.error(error);
    });
