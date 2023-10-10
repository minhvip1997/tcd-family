const fs = require('fs');
// const docxBuffer = fs.readFileSync('TCD.docx');
const docxBuffer = fs.readFileSync('ReportTCD3.docx');
var mammoth = require("mammoth");
var {Window} = require("happy-dom");
// const officegen = require('officegen')('docx');
const htmlPdf = require('html-pdf');

let meanValues = []
let peakValues = []
let sbiValues = []
let piValues = []
let diasValues = []
let stiValues = []
let riValues = []
let sdValues = []

let lmcaValues = []
let lacaValues = []
let lpcaValues = []
let rmcaValues = []
let racaValues = []
let rpcaValues = []

let imageHeader = ''
let imageFooter = ''


mammoth.extractRawText({buffer: docxBuffer})
.then(function(result){
    const string = result.value;
        
        
    const lines = string.split('\r'); // Split by carriage return
    let outText = '';
  
    lines.forEach(thisLine => {
      if (thisLine.includes('\x00') || thisLine.length === 0) {
        // Skip lines with null character or empty lines
      } else {
        outText += thisLine + ' ';
      }
    });
  
    outText = outText.replace(/[^a-zA-Z0-9\s,\.\-\n\r\t@/_\(\)]/g, '');

    // console.log('outTextoutText', outText);

     meanValues = extractMeanValues(outText, 'Mean')
     peakValues = extractMeanValues(outText, 'Peak')
     sbiValues = extractMeanValues(outText, 'SBI')
     piValues = extractMeanValues(outText, 'PI')
     diasValues = extractMeanValues(outText, 'Dias')
     stiValues = extractMeanValues(outText, 'STI')
     riValues = extractMeanValues(outText, 'RI')
     sdValues = extractMeanValues(outText, 'S/D')
    
     lmcaValues = extractMeanValues(outText, 'LMCA')
     lacaValues = extractMeanValues(outText, 'LACA')
     lpcaValues = extractMeanValues(outText, 'LPCA')
     rmcaValues = extractMeanValues(outText, 'RMCA')
     racaValues = extractMeanValues(outText, 'RACA')
     rpcaValues = extractMeanValues(outText, 'RPCA')

})
.catch(function(error) {
    console.error(error);
});

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
            image.style.maxWidth = '450px';
            image.style.marginLeft  = '10px';
            image.style.marginTop  = '5px';
            image.style.width  = '230px';
          }
        for (let i = 0; i < extractedImages.length; i += imagesPerRow) {
            const rowImages = extractedImages.slice(i, i + imagesPerRow);
            const rowHtml = rowImages.map(img => `<td>${img.outerHTML}</td>`).join('');
            
            

            imageRows.push(['<table style="margin-top:30px">'+'<tr>'+rowHtml+'<tr>'+'</table>']);
            imageHeader = ''
            imageFooter = ''

        }

        const extractedHTML = imageRows.map(row => `<div>${row}</div>`).join('');

        fs.writeFileSync('GeneratedContent2.html', extractedHTML, 'utf8', function(err) {
            if (err) {
                console.error(err);
            } else {
                console.log('Generated HTML content saved to GeneratedContent.html');
            }
        });

        const pdfOptions = {
            format: 'A4', // You can customize the PDF format as needed
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

function extractMeanValues(input, measurement) {
    const meanRegex =new RegExp(`${measurement}\\s+([\\d.-]+)`, 'g');
    const meanMatches = [...input.matchAll(meanRegex)];
  
    const meanValues = meanMatches.map(match => match[1]);
  
    return meanValues;
  }

   

   