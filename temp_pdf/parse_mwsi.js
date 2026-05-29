const fs = require('fs');
const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js');

const dataBuffer = new Uint8Array(fs.readFileSync('../assets/MAYNILAD WATER SERVICE UPDATE REPORT-March 29, 2026.pdf'));

(async () => {
    try {
        const pdf = await pdfjsLib.getDocument(dataBuffer).promise;
        let fullText = '';
        
        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const tc = await page.getTextContent();
            
            // Group by Y coordinate (row)
            const linesObj = {};
            tc.items.forEach(item => {
                const x = item.transform[4];
                const y = Math.round(item.transform[5] / 2) * 2;
                if (!linesObj[y]) linesObj[y] = [];
                linesObj[y].push({ x, str: item.str });
            });
            
            const sortedY = Object.keys(linesObj).map(Number).sort((a,b) => b - a);
            sortedY.forEach(y => {
                const lineItems = linesObj[y].sort((a,b) => a.x - b.x);
                const lineStr = lineItems.map(item => item.str).join(' ').replace(/\s+/g, ' ').trim();
                if (lineStr) fullText += lineStr + '\n';
            });
            
            fullText += '\n--- PAGE BREAK ---\n\n';
        }
        
        fs.writeFileSync('mwsi_output.txt', fullText, 'utf8');
        console.log("Done! Output written to mwsi_output.txt");
    } catch(err) {
        console.error('ERROR:', err);
    }
})();
