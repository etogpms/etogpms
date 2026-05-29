const fs = require('fs');
const files = ['script.js'];

files.forEach(f => {
  if (fs.existsSync(f)) {
    let content = fs.readFileSync(f, 'utf8');
    
    // Replace mangled bullet points
    content = content.replace(/â€¢/g, '-');
    
    // Replace mangled em-dashes
    content = content.replace(/â€”/g, '-');
    
    fs.writeFileSync(f, content, 'utf8');
    console.log('Fixed encoding issues in ' + f);
  }
});
