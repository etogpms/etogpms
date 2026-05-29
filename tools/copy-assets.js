const fs = require('fs');
const path = require('path');

const srcDir = path.resolve(__dirname, '..');
const destDir = path.resolve(__dirname, '../www');

// List of folders and files to copy from the root directory to www/
const itemsToCopy = [
  'index.html',
  '404.html',
  'deepwell_summary.html',
  'ipcr-form.html',
  'opcr-form-snippet.html',
  'opcr-form.html',
  'styles.css',
  'style-modern.css',
  'script.js',
  'docreg.js',
  'export.js',
  'ipcr_data.json',
  'assets',
  'images',
  'js',
  'storage',
  'temp_pdf'
];

console.log('Cleaning target directory (www)...');
try {
  if (fs.existsSync(destDir)) {
    fs.rmSync(destDir, { recursive: true, force: true });
    console.log('Existing www directory removed.');
  }
  fs.mkdirSync(destDir, { recursive: true });
  console.log('Created empty www directory.');
} catch (err) {
  console.error('Error cleaning www directory:', err);
  process.exit(1);
}

console.log('Copying assets...');
let successCount = 0;
let errorCount = 0;

itemsToCopy.forEach(item => {
  const srcPath = path.join(srcDir, item);
  const destPath = path.join(destDir, item);

  if (!fs.existsSync(srcPath)) {
    console.warn(`Warning: Source item does not exist, skipping: ${item}`);
    return;
  }

  try {
    const stats = fs.statSync(srcPath);
    if (stats.isDirectory()) {
      fs.cpSync(srcPath, destPath, { recursive: true });
      console.log(`Copied directory: ${item}`);
    } else {
      fs.copyFileSync(srcPath, destPath);
      console.log(`Copied file: ${item}`);
    }
    successCount++;
  } catch (err) {
    console.error(`Error copying ${item}:`, err);
    errorCount++;
  }
});

console.log(`\nCopy process finished. Success: ${successCount}, Errors: ${errorCount}`);
if (errorCount > 0) {
  process.exit(1);
} else {
  console.log('Web assets build completed successfully.');
}
