const fs = require('fs');
const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js'); // Use legacy for node environment

async function extractText(pdfPath) {
    const data = new Uint8Array(fs.readFileSync(pdfPath));
    const doc = await pdfjsLib.getDocument(data).promise;
    
    let text = "";
    for (let pageNum = 1; pageNum <= doc.numPages; pageNum++) {
        const page = await doc.getPage(pageNum);
        const textContent = await page.getTextContent();
        
        // Items are in textContent.items
        for (const item of textContent.items) {
           text += item.str + " ";
           if (item.hasEOL) text += "\n";
        }
        text += "\n---PAGE BREAK---\n";
    }
    
    fs.writeFileSync('output_pdfjs.txt', text, 'utf-8');
    console.log("Done");
}

extractText('../assets/20260310 MWC Daily Service Update Report (1).pdf').catch(console.error);
