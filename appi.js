const fs = require('fs');
const PDFDocument = require('pdf-lib').PDFDocument;
const pdf = require('pdf-parse');
async function extractTextFromPDF(pdfFilePath) {
    // Read the PDF file
    const dataBuffer = fs.readFileSync(pdfFilePath);
  
    // Create a PDFDocument instance (from pdf-lib) for further processing
    const pdfDoc = await PDFDocument.load(dataBuffer);
  
    // Initialize pdf-parse
    const data = await pdf(dataBuffer);
  
    // Extracted text
    const text = data.text;
  
    return text;
  }

function textToHtml(text) {
    // Split the text into lines
    const lines = text.split('\n');
    
    // Initialize the HTML content
    let html = '<div>';
  
    // Process each line
    for (const line of lines) {
      if (line) {
        // You can add style processing logic here, e.g., detect bold, italic, or other styles
        const styledLine = `<p style="font-weight: bold; font-style: italic">${line}</p>`;
        html += styledLine;
      }
    }
  
    // Close the HTML container
    html += '</div>';
  
    return html;
  }
  
  async function main() {
    try {
      const pdfFilePath = 'ReportTCD2.pdf'; // Replace with the path to your PDF file
      const extractedText = await extractTextFromPDF(pdfFilePath);
      
      // Convert extracted text to HTML with styles
      const htmlContent = textToHtml(extractedText);
      
      // You can save the HTML content to a file or use it as needed
      console.log(htmlContent);
    } catch (error) {
      console.error('Error:', error);
    }
  }
  
  main();
  