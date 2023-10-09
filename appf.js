const fs = require('fs');
const mammoth = require("mammoth");
const { Window } = require("happy-dom");
const htmlPdf = require('html-pdf');

// const docxBuffer = fs.readFileSync('ReportTCD2.docx');
const docxBuffer = fs.readFileSync('TCD.docx');

mammoth.convertToHtml({ buffer: docxBuffer })
    .then(function(result){
        var html = result.value; // The generated HTML
        // console.log('htmlhtml', html);

        const window = new Window({
            url: 'https://localhost:8080',
            width: 1024,
            height: 768
        });
        const document = window.document;
        document.body.innerHTML = html;
        
        const images = document?.querySelectorAll('img');
            const pdfOptions = {
                format: 'A3', // You can customize the PDF format as needed
            };

            fs.writeFileSync('GeneratedContent.html', html, 'utf8', function(err) {
                if (err) {
                    console.error(err);
                } else {
                    console.log('Generated HTML content saved to GeneratedContent.html');
                }
            });
            // return html

        
    })
    .catch(function(error) {
        console.error(error);
    });
