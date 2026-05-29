const fs = require('fs');

function parseData() {
    const text = fs.readFileSync('output.txt', 'utf-8');
    const fullTextLines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);

    let reportYearStr = new Date().getFullYear().toString();
    for (let i = 0; i < Math.min(20, fullTextLines.length); i++) {
        const line = fullTextLines[i];
        if (line.includes("WATER SERVICE UPDATE REPORT")) {
           let dateLine = fullTextLines[i+1] || "";
           const d = new Date(dateLine);
           if (!isNaN(d.getTime())) {
             reportYearStr = d.getFullYear().toString();
             break;
           }
        }
    }

    const monthMap = { 'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06', 
                       'Jul': '07', 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12' };

    let importedDays = [];

    const getLineValues = (line, prefix) => {
        const str = line.substring(prefix.length).trim();
        const matches = str.split(/\s+/);
        return matches;
    };

    const prefixDates = "DAM LEVELS (as of 6AM)";
    const dateLine = fullTextLines.find(l => l.startsWith(prefixDates));
    if (dateLine) {
        const rawDates = getLineValues(dateLine, prefixDates);
        importedDays = rawDates.map(rd => {
            // rd format: D-Mon e.g. 4-Mar
            const parts = rd.split('-');
            if (parts.length === 2) {
                const dd = parts[0].padStart(2, '0');
                const mm = monthMap[parts[1]] || '01';
                return { dateStr: `${reportYearStr}-${mm}-${dd}` };
            }
            return null;
        }).filter(d => d !== null);
    }

    const fillMetric = (prefix, prop) => {
        const line = fullTextLines.find(l => l.startsWith(prefix));
        if (line) {
            const vals = getLineValues(line, prefix);
            for (let i = 0; i < Math.min(importedDays.length, vals.length); i++) {
                const num = parseFloat(vals[i].replace(/,/g, ''));
                importedDays[i][prop] = isNaN(num) ? 0 : num;
            }
        } else {
             for (let i = 0; i < importedDays.length; i++) importedDays[i][prop] = 0;
        }
    };

    fillMetric("Angat Dam (m)", "angat");
    fillMetric("Ipo Dam (m)", "ipo");
    fillMetric("La Mesa Dam (m)", "laMesa");
    fillMetric("Deepwell Production", "deepwell");
    fillMetric("Cross Border Flow", "crossBorder");

    for (const d of importedDays) {
        d.supplyAug = Math.round((d.deepwell + d.crossBorder) * 100) / 100;
        d.plants = [];
        d.rawInflowsMap = {};
        d.productionMap = {};
    }

    const plantMap = [
        { pdfName: "Balara WTP 1", elName: "Balara WTP 1" },
        { pdfName: "Balara WTP 2", elName: "Balara WTP 2" },
        { pdfName: "East La Mesa WTP", elName: "East LMTP" },
        { pdfName: "Luzon WTP", elName: "Luzon WTP" },
        { pdfName: "Cardona WTP", elName: "Cardona WTP" },
        { pdfName: "Calawis WTP", elName: "Calawis TP" },
        { pdfName: "East Bay Phase 1 WTP", elName: "East Bay TP" }
    ];

    let inRawSection = false;
    let inProdSection = false;

    fullTextLines.forEach(line => {
        if (line.includes("RAW WATER INFLOWS")) { inRawSection = true; inProdSection = false; return; }
        if (line.includes("DAILY PRODUCTION")) { inRawSection = false; inProdSection = true; return; }
        if (line.includes("SUPPLY AUGMENTATION")) { inRawSection = false; inProdSection = false; return; }
        
        if (inRawSection || inProdSection) {
            for (let p of plantMap) {
                if (line.startsWith(p.pdfName)) {
                    const vals = getLineValues(line, p.pdfName);
                    for (let i = 0; i < Math.min(importedDays.length, vals.length); i++) {
                        const num = parseFloat(vals[i].replace(/,/g, ''));
                        const val = isNaN(num) ? 0 : num;
                        if (inRawSection) importedDays[i].rawInflowsMap[p.elName] = val;
                        if (inProdSection) importedDays[i].productionMap[p.elName] = val;
                    }
                }
            }
        }
    });

    for (const d of importedDays) {
        for (let p of plantMap) {
            const raw = d.rawInflowsMap[p.elName] || 0;
            const prod = d.productionMap[p.elName] || 0;
            d.plants.push({ name: p.elName, inflow: raw, production: prod });
        }
        d.inflows = Math.round(d.plants.reduce((sum, p) => sum + p.inflow, 0) * 100) / 100;
        d.production = Math.round(d.plants.reduce((sum, p) => sum + p.production, 0) * 100) / 100;
        
        delete d.deepwell;
        delete d.crossBorder;
        delete d.rawInflowsMap;
        delete d.productionMap;
    }

    console.log(JSON.stringify(importedDays, null, 2));
}

parseData();
