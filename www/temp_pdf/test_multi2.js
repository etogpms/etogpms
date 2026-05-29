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

    const parseDatesFromLine = (line, prefix) => {
        const str = line.substring(prefix.length).trim();
        const matches = str.split(/\s+/);
        return matches.map(rd => {
            const parts = rd.split('-');
            if (parts.length === 2) {
                const dd = parts[0].padStart(2, '0');
                const mm = monthMap[parts[1]] || '01';
                return `${reportYearStr}-${mm}-${dd}`;
            }
            return null;
        });
    };

    const getLineValues = (line, prefix) => {
        const str = line.substring(prefix.length).trim();
        return str.split(/\s+/);
    };

    let damDates = [];
    let inflowDates = [];
    let prodDates = [];
    let augDates = [];

    fullTextLines.forEach(line => {
        if (line.startsWith("DAM LEVELS (as of 6AM)")) damDates = parseDatesFromLine(line, "DAM LEVELS (as of 6AM)");
        if (line.startsWith("RAW WATER INFLOWS")) inflowDates = parseDatesFromLine(line, "RAW WATER INFLOWS (MLD)");
        if (line.startsWith("DAILY PRODUCTION")) prodDates = parseDatesFromLine(line, "DAILY PRODUCTION (MLD)");
        if (line.startsWith("SUPPLY AUGMENTATION")) augDates = parseDatesFromLine(line, "SUPPLY AUGMENTATION (MLD)");
    });

    // Master map of date -> payload
    const recordsByDate = {};
    const getRecord = (dateStr) => {
        if (!dateStr) return null;
        if (!recordsByDate[dateStr]) {
            recordsByDate[dateStr] = {
                dateStr,
                angat: 0, ipo: 0, laMesa: 0,
                deepwell: 0, crossBorder: 0,
                rawInflowsMap: {}, productionMap: {},
                hasDam: false, hasPlants: false, hasAug: false
            };
        }
        return recordsByDate[dateStr];
    };

    const fillMetric = (prefix, datesArr, prop, flag) => {
        const line = fullTextLines.find(l => l.startsWith(prefix));
        if (line) {
            const vals = getLineValues(line, prefix);
            for (let i = 0; i < Math.min(datesArr.length, vals.length); i++) {
                const dateStr = datesArr[i];
                if (dateStr) {
                    const rec = getRecord(dateStr);
                    const num = parseFloat(vals[i].replace(/,/g, ''));
                    rec[prop] = isNaN(num) ? 0 : num;
                    if (flag) rec[flag] = true;
                }
            }
        }
    };

    fillMetric("Angat Dam (m)", damDates, "angat", "hasDam");
    fillMetric("Ipo Dam (m)", damDates, "ipo", "hasDam");
    fillMetric("La Mesa Dam (m)", damDates, "laMesa", "hasDam");

    fillMetric("Deepwell Production", augDates, "deepwell", "hasAug");
    fillMetric("Cross Border Flow", augDates, "crossBorder", "hasAug");

    const plantMap = [
        { pdfName: "Balara WTP 1", elName: "Balara WTP 1" },
        { pdfName: "Balara WTP 2", elName: "Balara WTP 2" },
        { pdfName: "East La Mesa WTP", elName: "East LMTP" },
        { pdfName: "Luzon WTP", elName: "Luzon WTP" },
        { pdfName: "Cardona WTP", elName: "Cardona WTP" },
        { pdfName: "Calawis WTP", elName: "Calawis TP" },
        { pdfName: "East Bay Phase 1 WTP", elName: "East Bay TP" }
    ];

    let currentSectionConfig = { prefix: "", dates: [] };

    fullTextLines.forEach(line => {
        if (line.startsWith("RAW WATER INFLOWS")) { currentSectionConfig = { prefix: "inflow", dates: inflowDates }; return; }
        if (line.startsWith("DAILY PRODUCTION")) { currentSectionConfig = { prefix: "prod", dates: prodDates }; return; }
        if (line.startsWith("SUPPLY AUGMENTATION")) { currentSectionConfig = { prefix: "", dates: [] }; return; }
        if (line.startsWith("DAM LEVELS")) { currentSectionConfig = { prefix: "", dates: [] }; return; }
        
        if (currentSectionConfig.prefix) {
            for (let p of plantMap) {
                if (line.startsWith(p.pdfName)) {
                    const vals = getLineValues(line, p.pdfName);
                    for (let i = 0; i < Math.min(currentSectionConfig.dates.length, vals.length); i++) {
                        const dateStr = currentSectionConfig.dates[i];
                        if (dateStr) {
                            const rec = getRecord(dateStr);
                            const num = parseFloat(vals[i].replace(/,/g, ''));
                            const val = isNaN(num) ? 0 : num;
                            if (currentSectionConfig.prefix === "inflow") rec.rawInflowsMap[p.elName] = val;
                            if (currentSectionConfig.prefix === "prod") rec.productionMap[p.elName] = val;
                            rec.hasPlants = true;
                        }
                    }
                }
            }
        }
    });

    const finalRecords = [];
    // Sort keys just in case
    Object.keys(recordsByDate).sort().forEach(dateStr => {
        const d = recordsByDate[dateStr];
        d.supplyAug = Math.round((d.deepwell + d.crossBorder) * 100) / 100;
        d.plants = [];
        
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
        
        finalRecords.push(d);
    });

    console.log(JSON.stringify(finalRecords, null, 2));
}

parseData();
