const fs = require('fs');

const extractValues = (line, prefix) => {
  let rest = line.substring(prefix.length).trim();
  // Pre-clean PDF artifacts: collapse split numbers like "8 .02" -> "8.02", "17. 55" -> "17.55"
  rest = rest.replace(/(\d)\s+\./g, '$1.').replace(/\.\s+(\d)/g, '.$1');
  
  const tokens = rest.split(/\s+/);
  const parsedValues = [];
  for (const t of tokens) {
    if (t === '*') continue; // ignore standalone single asterisk (detached annotation)
    
    // For anything else (numbers, or "**"), clean it
    const cleaned = t.replace(/[*]/g, '').replace(/,/g, '').trim();
    if (!cleaned) {
      if (t.includes('**') || t === '-' || t === '') {
        parsedValues.push(null);
      }
      else {
        // If it was just "*", wait, we already continue'd if t === '*'.
        // What if it was "***"?
        parsedValues.push(null);
      }
    } else {
      const num = parseFloat(cleaned);
      parsedValues.push(isNaN(num) ? null : num);
    }
  }
  return parsedValues;
};

const line1 = 'La Mesa Water Treatment Plant 1* 1351.760 1375.531 1402.488 1398.524 1411.398 1399.594';
const line2 = 'La Mesa Water Treatment Plant 1 1251.53 * 1273.84 * 1269.99 * 1274.46 * 1281.06 * 1281.68 *';
const line3 = 'San Martin de Porres 1.67 ** ** ** ** **';

const out = {
  line1: extractValues(line1, 'La Mesa Water Treatment Plant 1'), // Note: the "*" after Plant 1 is part of prefix or not?
  // Wait, the prefix passed in is 'La Mesa Water Treatment Plant 1', so the rest string will be '* 1351.760...' 
  line2: extractValues(line2, 'La Mesa Water Treatment Plant 1'),
  line3: extractValues(line3, 'San Martin de Porres')
};

fs.writeFileSync('out2.json', JSON.stringify(out, null, 2), 'utf8');
