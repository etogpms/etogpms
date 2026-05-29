const fs = require('fs');

const monthMap = { 'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06',
                   'Jul': '07', 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12' };

const extractValues = (line, prefix) => {
  let rest = line.substring(prefix.length).trim();
  rest = rest.replace(/(\d)\s+\./g, '$1.').replace(/\.\s+(\d)/g, '.$1');
  rest = rest.replace(/(\d)\s+(\d)/g, '$1$2');
  
  return rest.split(/\s+/).map(v => {
    const cleaned = v.replace(/[*]/g, '').replace(/,/g, '').trim();
    if (!cleaned) return null;
    const num = parseFloat(cleaned);
    return isNaN(num) ? null : num;
  });
};

const line1 = 'La Mesa Water Treatment Plant 1* 1351.760 1375.531 1402.488 1398.524 1411.398 1399.594';
const line2 = 'La Mesa Water Treatment Plant 1 1251.53 * 1273.84 * 1269.99 * 1274.46 * 1281.06 * 1281.68 *';

let fullTextLines = [
  '23 - Mar 24 - Mar 25 - Mar 26 - Mar 2 7 - Mar 2 8 - Mar 2 9 - Mar'
];
let reportYearStr = '2026';
let columnDates = [];
for (let i = 0; i < fullTextLines.length; i++) {
  const line = fullTextLines[i].replace(/(\d)\s+(\d)/g, '$1$2');
  const datePattern = /(\d{1,2})\s*-\s*(\w{3})/g;
  let match;
  const found = [];
  while ((match = datePattern.exec(line)) !== null) {
    found.push({ day: match[1], mon: match[2] });
  }
  if (found.length >= 3) {
    columnDates = found.map(f => {
      const dd = f.day.padStart(2, '0');
      const mm = monthMap[f.mon] || '01';
      return `${reportYearStr}-${mm}-${dd}`;
    });
    break;
  }
}

const out = {
  line1: extractValues(line1, 'La Mesa Water Treatment Plant 1'),
  line2: extractValues(line2, 'La Mesa Water Treatment Plant 1'),
  columnDates
};

fs.writeFileSync('out.json', JSON.stringify(out, null, 2), 'utf8');
