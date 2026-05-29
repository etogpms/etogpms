const fs = require('fs');
const pdf = require('pdf-parse');

let dataBuffer = fs.readFileSync('../assets/20260310 MWC Daily Service Update Report (1).pdf');

pdf(dataBuffer).then(function(data) {
    console.log("--- START TEXT ---");
    console.log(data.text);
    console.log("--- END TEXT ---");
}).catch(console.error);
