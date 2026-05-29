const fs = require('fs');
const path = require('path');

const rootDir = path.resolve(__dirname, '..');

// Helper to normalize line endings to perform safe matching
function normalizeText(text) {
  return text.replace(/\r\n/g, '\n');
}

function patchFile(filename, search, replace) {
  const filePath = path.join(rootDir, filename);
  if (!fs.existsSync(filePath)) {
    console.error(`File not found: ${filePath}`);
    process.exit(1);
  }

  const originalContent = fs.readFileSync(filePath, 'utf8');
  const normalizedOriginal = normalizeText(originalContent);
  const normalizedSearch = normalizeText(search);
  const normalizedReplace = normalizeText(replace);

  // If already replaced, skip silently
  if (normalizedOriginal.includes(normalizedReplace)) {
    console.log(`Already patched or present in ${filename}, skipping.`);
    return;
  }

  if (!normalizedOriginal.includes(normalizedSearch)) {
    console.error(`Could not find search pattern in ${filename}.`);
    process.exit(1);
  }

  // Replace using the normalized content
  const updatedContent = normalizedOriginal.replace(normalizedSearch, normalizedReplace);

  // Preserve the original file's line endings by detecting what it used
  const usesCrlf = originalContent.includes('\r\n');
  const finalContent = usesCrlf ? updatedContent.replace(/\n/g, '\r\n') : updatedContent;

  fs.writeFileSync(filePath, finalContent, 'utf8');
  console.log(`Successfully patched ${filename}`);
}

// 1. Patch export.js
console.log('Patching export.js...');
const searchPdfWantOpen = `      const filename = \`\${(p.name || 'site_inspection_report').replace(/[^a-z0-9\\- _()+]/gi,'_')}.pdf\`;
      const wantOpen = !!(opts && opts.open);
      if(wantOpen){`;

const replacePdfWantOpen = `      const filename = \`\${(p.name || 'site_inspection_report').replace(/[^a-z0-9\\- _()+]/gi,'_')}.pdf\`;
      const isMobile = /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent);
      const wantOpen = isMobile ? false : !!(opts && opts.open);
      if(wantOpen){`;

const searchDocxWantOpen = `      const filename = \`\${(p.name || 'site_inspection_report').replace(/[^a-z0-9\\- _()+]/gi,'_')}.docx\`;
      const wantOpen = !!(opts && opts.open);
      if(wantOpen && window.navigator && typeof window.navigator.msSaveOrOpenBlob === 'function'){`;

const replaceDocxWantOpen = `      const filename = \`\${(p.name || 'site_inspection_report').replace(/[^a-z0-9\\- _()+]/gi,'_')}.docx\`;
      const isMobile = /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent);
      const wantOpen = isMobile ? false : !!(opts && opts.open);
      if(wantOpen && window.navigator && typeof window.navigator.msSaveOrOpenBlob === 'function'){`;

patchFile('export.js', searchPdfWantOpen, replacePdfWantOpen);
patchFile('export.js', searchDocxWantOpen, replaceDocxWantOpen);

// 2. Patch docreg.js
console.log('Patching docreg.js...');
const searchDocregBlob = `  // Download blob as file
  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 100);
  }`;

const replaceDocregBlob = `  // Download blob as file
  function downloadBlob(blob, filename) {
    if (window.saveAs) {
      window.saveAs(blob, filename);
    } else {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 100);
    }
  }`;

patchFile('docreg.js', searchDocregBlob, replaceDocregBlob);

// 3. Patch script.js
console.log('Patching script.js...');
const searchScriptWord = `  function downloadTreatmentPlantExpandedWord() {
    const series = treatmentPlantsExpandedSeries;
    if (!series) return;
    const reportHtml = buildTreatmentPlantReportHtml(series, getTreatmentPlantExpandedChartImage());
    const blob = new Blob(['\\ufeff' + reportHtml], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const provider = String(series.provider || 'Treatment').replace(/[^a-z0-9_-]+/gi, '_');
    a.href = url;
    a.download = \`Treatment_Plants_\${provider}_\${series.year}.doc\`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }`;

const replaceScriptWord = `  function downloadTreatmentPlantExpandedWord() {
    const series = treatmentPlantsExpandedSeries;
    if (!series) return;
    const reportHtml = buildTreatmentPlantReportHtml(series, getTreatmentPlantExpandedChartImage());
    const blob = new Blob(['\\ufeff' + reportHtml], { type: 'application/msword' });
    const provider = String(series.provider || 'Treatment').replace(/[^a-z0-9_-]+/gi, '_');
    const filename = \`Treatment_Plants_\${provider}_\${series.year}.doc\`;
    if (window.saveAs) {
      window.saveAs(blob, filename);
    } else {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    }
  }`;

const searchScriptCsv = `  function exportPresentationsCsv() {
    const list = PresentationsFeatureInstance.getPresentations();
    if (list.length === 0) { alert('No data to export.'); return; }
    const headers = ['Date', 'Subject', 'Time', 'Venue', 'Presenter', 'Remarks'];
    const rows = list.map(p => [
      p.date || '',
      p.subject || '',
      p.time || '',
      p.venue || '',
      p.presenter || '',
      p.remarks || ''
    ]);
    const csv = [headers.join(','), ...rows.map(r => r.map(v => \`"\${String(v).replace(/"/g, '""')}"\`).join(','))].join('\\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = \`product_presentations_\${todayYmd()}.csv\`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }`;

const replaceScriptCsv = `  function exportPresentationsCsv() {
    const list = PresentationsFeatureInstance.getPresentations();
    if (list.length === 0) { alert('No data to export.'); return; }
    const headers = ['Date', 'Subject', 'Time', 'Venue', 'Presenter', 'Remarks'];
    const rows = list.map(p => [
      p.date || '',
      p.subject || '',
      p.time || '',
      p.venue || '',
      p.presenter || '',
      p.remarks || ''
    ]);
    const csv = [headers.join(','), ...rows.map(r => r.map(v => \`"\${String(v).replace(/"/g, '""')}"\`).join(','))].join('\\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const filename = \`product_presentations_\${todayYmd()}.csv\`;
    if (window.saveAs) {
      window.saveAs(blob, filename);
    } else {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = filename;
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }
  }`;

patchFile('script.js', searchScriptWord, replaceScriptWord);
patchFile('script.js', searchScriptCsv, replaceScriptCsv);

// 4. Patch index.html
console.log('Patching index.html...');
const searchIndexFilters = `          <!-- Filters - All in one row -->
          <div class="card p-3 mb-4 shadow-sm">
            <div class="d-flex flex-wrap align-items-end gap-2 mb-3">
              <div>
                <label class="form-label mb-0 small" for="presentationStartDateFilter">Start Date</label>
                <input type="date" id="presentationStartDateFilter" class="form-control form-control-sm"
                  style="width: 130px;" />
              </div>
              <div>
                <label class="form-label mb-0 small" for="presentationEndDateFilter">End Date</label>
                <input type="date" id="presentationEndDateFilter" class="form-control form-control-sm"
                  style="width: 130px;" />
              </div>
              <div>
                <label class="form-label mb-0 small" for="presentationYearFilter">Year</label>
                <select id="presentationYearFilter" class="form-select form-select-sm" style="width: 90px;">
                  <option value="">Current</option>
                  <option value="ALL">All years</option>
                </select>
              </div>
              <div>
                <label class="form-label mb-0 small" for="presentationSubjectFilter">Subject</label>
                <input id="presentationSubjectFilter" type="text" class="form-control form-control-sm"
                  placeholder="Search subject..." style="width: 150px;" />
              </div>
              <div>
                <label class="form-label mb-0 small" for="presentationPresenterFilter">Presenter</label>
                <input id="presentationPresenterFilter" type="text" class="form-control form-control-sm"
                  placeholder="Search presenter..." style="width: 150px;" />
              </div>
              <div class="ms-auto d-flex flex-wrap gap-1">
                <button class="btn btn-sm btn-primary" id="addPresentationBtn">
                  <i class="fa-solid fa-plus me-1"></i>Add Presentation
                </button>
              </div>
            </div>
          </div>`;

const replaceIndexFilters = `          <!-- Filters - Responsive Layout -->
          <div class="card p-3 mb-4 shadow-sm">
            <div class="row g-2 mb-1 align-items-end">
              <div class="col-6 col-md-auto">
                <label class="form-label mb-0 small" for="presentationStartDateFilter">Start Date</label>
                <input type="date" id="presentationStartDateFilter" class="form-control form-control-sm w-100" />
              </div>
              <div class="col-6 col-md-auto">
                <label class="form-label mb-0 small" for="presentationEndDateFilter">End Date</label>
                <input type="date" id="presentationEndDateFilter" class="form-control form-control-sm w-100" />
              </div>
              <div class="col-4 col-md-auto">
                <label class="form-label mb-0 small" for="presentationYearFilter">Year</label>
                <select id="presentationYearFilter" class="form-select form-select-sm w-100">
                  <option value="">Current</option>
                  <option value="ALL">All years</option>
                </select>
              </div>
              <div class="col-8 col-md">
                <label class="form-label mb-0 small" for="presentationSubjectFilter">Subject</label>
                <input id="presentationSubjectFilter" type="text" class="form-control form-control-sm w-100"
                  placeholder="Search subject..." />
              </div>
              <div class="col-12 col-md">
                <label class="form-label mb-0 small" for="presentationPresenterFilter">Presenter</label>
                <input id="presentationPresenterFilter" type="text" class="form-control form-control-sm w-100"
                  placeholder="Search presenter..." />
              </div>
              <div class="col-12 col-md-auto ms-md-auto mt-2 mt-md-0">
                <button class="btn btn-sm btn-primary w-100" id="addPresentationBtn">
                  <i class="fa-solid fa-plus me-1"></i>Add Presentation
                </button>
              </div>
            </div>
          </div>`;

patchFile('index.html', searchIndexFilters, replaceIndexFilters);

console.log('All patches completed successfully!');
