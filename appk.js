const fs = require('fs');
const docxBuffer = fs.readFileSync('TCD.docx');
// const docxBuffer = fs.readFileSync('ReportTCD2.docx');
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
            
            if(lmcaValues && i===0){
                imageHeader += '<tr><td><span style="margin-right:50px; margin-left:10px">LMCA:</span> '+ lmcaValues.toString()+'mm' + '<span style="margin-left:50px">+</span></td>';
            }
            if(lacaValues && i===0){
                imageHeader += '<td><span style="margin-right:50px; margin-left:10px">LACA:</span> '+ lmcaValues.toString()+'mm' + '<span style="margin-left:50px">-</span></td>';
            }
            if(lpcaValues && i===0){
                imageHeader += '<td><span style="margin-right:50px; margin-left:10px">LPCA:</span> '+ lmcaValues.toString()+'mm' + '<span style="margin-left:50px">+</span></td></tr>';
            }
            if(i===0){
                imageFooter += `<tr><td style="padding-left: 10px;"><span style="float:left">Mean: ${meanValues[0]}</span> <span style="margin-left: 20px;">Peak: ${peakValues[0]}</span> <span style="float:right">SBI: ${sbiValues[0]}</span></td>
                <td style="padding-left: 10px;"><span style="float:left">Mean: ${meanValues[1]}</span> <span style="margin-left: 15px;">Peak: ${peakValues[1]}</span> <span style="float:right">SBI: ${sbiValues[1]}</span></td>
                <td style="padding-left: 10px;"><span style="float:left">Mean: ${meanValues[2]}</span> <span style="margin-left: 20px;">Peak: ${peakValues[2]}</span> <span style="float:right">SBI: ${sbiValues[2]}</span></td>
                </tr>
                <tr>
                <td style="padding-left: 10px;"><span style="float:left">PI: ${piValues[0]}</span> <span style="margin-left: 40px;">Dias: ${diasValues[0]}</span> <span style="float:right">STI: ${stiValues[0]}</span></td>
                <td style="padding-left: 10px;"><span style="float:left">PI: ${piValues[1]}</span> <span style="margin-left: 40px;">Dias: ${diasValues[1]}</span> <span style="float:right">STI: ${stiValues[1]}</span></td>
                <td style="padding-left: 10px;"><span style="float:left">PI: ${piValues[2]}</span> <span style="margin-left: 40px;">Dias: ${diasValues[2]}</span> <span style="float:right">STI: ${stiValues[2]}</span></td>
                </tr>
                <tr>
                <td style="padding-left: 10px;"><span style="float:left">RI: ${riValues[0]}</span> <span style="margin-left: 40px;">S/D: ${sdValues[0]}</span></td>
                <td style="padding-left: 10px;"><span style="float:left">RI: ${riValues[1]}</span> <span style="margin-left: 40px;">S/D: ${sdValues[1]}</span></td>
                <td style="padding-left: 10px;"><span style="float:left">RI: ${riValues[2]}</span> <span style="margin-left: 40px;">S/D: ${sdValues[2]}</span></td>
                </tr>`;
            }
            
            if(rmcaValues && i===3){
                imageHeader += '<tr><td><span style="margin-right:50px; margin-left:10px">RMCA:</span> '+ rmcaValues.toString()+'mm' + '<span style="margin-left:50px">+</span></td>';
            }
            if(racaValues && i===3){
                imageHeader += '<td><span style="margin-right:50px; margin-left:10px">RACA:</span> '+ racaValues.toString()+'mm' + '<span style="margin-left:50px">+</span></td>';
            }
            if(rpcaValues && i===3){
                imageHeader += '<td><span style="margin-right:50px; margin-left:10px">RPCA:</span> '+ rpcaValues.toString()+'mm' + '<span style="margin-left:50px">+</span></td></tr>';
            }
            if(i===3){
                imageFooter += `<tr><td style="padding-left: 10px;"><span style="float:left">Mean: ${meanValues[3]}</span> <span style="margin-left: 20px;">Peak: ${peakValues[3]}</span> <span style="float:right">SBI: ${sbiValues[3]}</span></td>
                <td style="padding-left: 10px;"><span style="float:left">Mean: ${meanValues[4]} </span><span style="margin-left: 15px;">Peak: ${peakValues[4]}</span> <span style="float:right">SBI: ${sbiValues[4]}</span></td>
                <td style="padding-left: 10px;"><span style="float:left">Mean: ${meanValues[5]} </span><span style="margin-left: 20px;">Peak: ${peakValues[5]}</span> <span style="float:right">SBI: ${sbiValues[5]}</span></td>
                </tr>
                <tr>
                <td style="padding-left: 10px;"><span style="float:left">PI: ${piValues[3]}</span> <span style="margin-left:40px">Dias: ${diasValues[3]}</span> <span style="float:right">STI: ${stiValues[3]}</span></td>
                <td style="padding-left: 10px;"><span style="float:left">PI: ${piValues[4]}</span> <span style="margin-left:40px">Dias: ${diasValues[4]}</span> <span style="float:right">STI: ${stiValues[4]}</span></td>
                <td style="padding-left: 10px;"><span style="float:left">PI: ${piValues[5]}</span> <span style="margin-left:40px">Dias: ${diasValues[5]}</span> <span style="float:right">STI: ${stiValues[5]}</span></td>
                </tr>
                <tr>
                <td style="padding-left: 10px;"><span style="float:left">RI: ${riValues[3]}</span> <span style="margin-left:40px">S/D: ${sdValues[3]}</span></td>
                <td style="padding-left: 10px;"><span style="float:left">RI: ${riValues[4]}</span> <span style="margin-left:40px">S/D: ${sdValues[4]}</span></td>
                <td style="padding-left: 10px;"><span style="float:left">RI: ${riValues[5]}</span> <span style="margin-left:40px">S/D: ${sdValues[5]}</span></td>
                </tr>`;
            }

            imageRows.push(['<table style="margin-top:30px">'+imageHeader+'<tr>'+rowHtml+'<tr>'+imageFooter+'</table>']);
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

   

   