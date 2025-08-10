// export.js - handles DOCX generation from a Word template using Docxtemplater
// Requires: PizZip, Docxtemplater, FileSaver (already included in index.html)

(function(){
  'use strict';
  try{ window.EXPORT_JS_VERSION = 7; console.info('[export] script v7 loaded'); }catch(_){/* ignore */}

  // Primary on-disk location of the Word template
  const TEMPLATE_PATH = 'assets/site_inspection_template.docx';

  // Dynamically load a script if a CDN is blocked or not yet loaded
  function loadScript(src){
    return new Promise((resolve, reject)=>{
      const s = document.createElement('script');
      s.src = src;
      s.async = true;
      s.onload = ()=> resolve();
      s.onerror = ()=> reject(new Error('Failed to load script ' + src));
      document.head.appendChild(s);
    });
  }

  async function ensureDocxDeps(){
    const DocxTemplater = window.Docxtemplater || window.docxtemplater;
    if(window.PizZip && DocxTemplater) return; // already present
    // Try jsDelivr fallbacks
    if(!window.PizZip){
      // Local vendor first
      try{ await loadScript('assets/vendor/pizzip.min.js'); console.info('[export] PizZip loaded from local vendor'); }catch(_){ /* try CDN */ }
      if(!window.PizZip){
        try{ await loadScript('https://cdn.jsdelivr.net/npm/pizzip@3.1.7/dist/pizzip.min.js'); }catch(e){ /* ignore; try unpkg */}
        if(!window.PizZip){ await loadScript('https://unpkg.com/pizzip@3.1.7/dist/pizzip.min.js'); }
      }
    }
    if(!(window.Docxtemplater || window.docxtemplater)){
      // Local vendor first
      try{ await loadScript('assets/vendor/docxtemplater.js'); console.info('[export] Docxtemplater loaded from local vendor'); }catch(_){ /* try CDN */ }
      if(!(window.Docxtemplater || window.docxtemplater)){
        try{ await loadScript('https://cdn.jsdelivr.net/npm/docxtemplater@3.22.2/build/docxtemplater.js'); }catch(e){ /* ignore; try unpkg */}
        if(!(window.Docxtemplater || window.docxtemplater)){
          await loadScript('https://unpkg.com/docxtemplater@3.22.2/build/docxtemplater.js');
        }
      }
    }
  }

  async function ensureImageModule(){ /* no-op: image module not used with alt-text replacement */ }

  // Replace placeholder images by matching Alt Text Description (wp:docPr/@descr)
  // with keys like 'ProjectPhoto1', 'ProjectPhoto2', 'ProjectPhoto3'.
  async function replaceAltTextImagesInZip(zip, photos){
    try{
      const docXmlPath = 'word/document.xml';
      const relsPath = 'word/_rels/document.xml.rels';
      let docXml = zip.file(docXmlPath).asText();
      const relsXml = zip.file(relsPath).asText();

      function findRidFor(tag){
        const idx = docXml.indexOf(`descr="${tag}"`);
        if(idx === -1) return null;
        const fwd = docXml.slice(idx, idx + 5000);
        let m = fwd.match(/r:embed=\"(rId[0-9]+)\"/);
        if(m) return m[1];
        const back = docXml.slice(Math.max(0, idx - 5000), idx);
        const allBack = [...back.matchAll(/r:embed=\"(rId[0-9]+)\"/g)];
        if(allBack.length) return allBack[allBack.length-1][1];
        return null;
      }

      function targetFor(rid){
        const re = new RegExp(`<Relationship[^>]+Id=\"${rid}\"[^>]+Target=\"([^\"]+)\"`, 'i');
        const m = relsXml.match(re);
        return m ? m[1] : null;
      }

      async function valueToPng(tagValue){
        if(!tagValue) return null;
        if(typeof tagValue === 'string' && tagValue.startsWith('data:image/')){
          const isPng = /^data:image\/png/i.test(tagValue);
          if(isPng) return base64DataURLToArrayBuffer(tagValue);
          return await dataUrlToPngArrayBuffer(tagValue);
        }
        if(typeof tagValue === 'string' && (/^https?:\/\//i.test(tagValue) || tagValue.startsWith('blob:'))){
          return await urlToPngArrayBuffer(tagValue);
        }
        return null;
      }

      const tagKeys = ['ProjectPhoto1','ProjectPhoto2','ProjectPhoto3','Project Photo 1','Project Photo 2','Project Photo 3'];
      let xmlChanged = false;
      for(const key of tagKeys){
        const val = photos && (photos[key] || photos[key.replace(/\s+/g,'')]);
        // If a value is provided, embed image; otherwise replace the drawing with literal 'n/a' text
        if(val){
          const ab = await valueToPng(val);
          if(!ab) continue;
          const rid = findRidFor(key);
          if(!rid){ try{ console.warn('[export] Alt-text not found for', key); }catch(_){ } continue; }
          const target = targetFor(rid);
          if(!target){ try{ console.warn('[export] Relationship target not found for', rid); }catch(_){ } continue; }
          const mediaPath = ('word/' + target).replace(/\\/g,'/');
          zip.file(mediaPath, new Uint8Array(ab));
          try{ console.debug('[export] Replaced image for', key, '->', mediaPath); }catch(_){ }
        }else{
          // Replace the enclosing <w:drawing> element (with this descr) by a simple text run
          const marker = `descr="${key}"`;
          // Try to replace all occurrences for this key (usually one)
          let attempts = 0;
          while(attempts < 3){
            const descrIdx = docXml.indexOf(marker);
            if(descrIdx === -1) break;
            const drawingStart = docXml.lastIndexOf('<w:drawing', descrIdx);
            const drawingEnd = docXml.indexOf('</w:drawing>', descrIdx);
            if(drawingStart === -1 || drawingEnd === -1) { break; }
            const drawingClose = drawingEnd + '</w:drawing>'.length;
            docXml = docXml.slice(0, drawingStart) + '<w:t>n/a</w:t>' + docXml.slice(drawingClose);
            xmlChanged = true;
            attempts++;
          }
          if(attempts === 0){ try{ console.warn('[export] Could not replace drawing with text for', key); }catch(_){ } }
        }
      }
      if(xmlChanged){
        zip.file(docXmlPath, docXml);
        try{ console.debug('[export] Updated document.xml with literal text for missing photos'); }catch(_){ }
      }
    }catch(e){
      console.warn('[export] Alt-text image replace failed:', e && e.message ? e.message : e);
    }
  }

  function base64DataURLToArrayBuffer(dataURL){
    const base64Regex = /^data:image\/(png|jpe?g|gif|webp|svg|svg\+xml);base64,/i;
    if(!dataURL || !base64Regex.test(dataURL)) return null;
    const stringBase64 = dataURL.replace(base64Regex, '');
    const binaryString = atob(stringBase64);
    const len = binaryString.length;
    const bytes = new Uint8Array(len);
    for(let i=0;i<len;i++){ bytes[i] = binaryString.charCodeAt(i); }
    return bytes.buffer;
  }

  // 1x1 transparent PNG in base64 (safe placeholder to avoid module errors)
  const TRANSPARENT_PNG_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGNgYAAAAAMAAWgmWQ0AAAAASUVORK5CYII=';
  function transparentPngArrayBuffer(){
    return base64DataURLToArrayBuffer('data:image/png;base64,' + TRANSPARENT_PNG_BASE64);
  }

  // Convert any dataURL (jpeg/webp/gif/svg/etc.) to PNG ArrayBuffer
  function dataUrlToPngArrayBuffer(dataURL){
    return new Promise((resolve,reject)=>{
      try{
        const img = new Image();
        img.onload = ()=>{
          const canvas = document.createElement('canvas');
          const w = img.naturalWidth || img.width;
          const h = img.naturalHeight || img.height;
          canvas.width = w; canvas.height = h;
          const ctx = canvas.getContext('2d');
          ctx.drawImage(img,0,0,w,h);
          const pngDataUrl = canvas.toDataURL('image/png');
          const ab = base64DataURLToArrayBuffer(pngDataUrl);
          resolve(ab);
        };

        img.onerror = reject;
        img.src = dataURL;
      }catch(e){ reject(e); }
    });
  }

  // Generate a simple PNG placeholder with centered 'N/A' text
  function createTextPlaceholderPng(text){
    return new Promise((resolve)=>{
      try{
        const canvas = document.createElement('canvas');
        // Reasonable default size; Word will scale to template's frame
        canvas.width = 800; canvas.height = 450;
        const ctx = canvas.getContext('2d');
        // Background
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        // Border
        ctx.strokeStyle = '#cccccc';
        ctx.lineWidth = 4;
        ctx.strokeRect(2, 2, canvas.width - 4, canvas.height - 4);
        // Text
        ctx.fillStyle = '#888888';
        ctx.font = 'bold 96px Arial, sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(String(text || 'N/A'), canvas.width/2, canvas.height/2);
        const dataUrl = canvas.toDataURL('image/png');
        const ab = base64DataURLToArrayBuffer(dataUrl);
        resolve(ab || transparentPngArrayBuffer());
      }catch(_){
        resolve(transparentPngArrayBuffer());
      }
    });
  }

  // Convert DOCX Blob to PDF by calling a configurable HTTP endpoint.
  // Set window.DOCX_PDF_ENDPOINT to your converter URL (e.g., a local Gotenberg/LibreOffice service).
  async function convertDocxBlobToPdfUsingEndpoint(docxBlob){
    const endpoint = (typeof window !== 'undefined' && typeof window.DOCX_PDF_ENDPOINT === 'string') ? window.DOCX_PDF_ENDPOINT : '';
    if(!endpoint){
      throw new Error('DOCX_PDF_ENDPOINT is not configured');
    }
    const type = (typeof window !== 'undefined' && window.DOCX_PDF_ENDPOINT_TYPE) ? String(window.DOCX_PDF_ENDPOINT_TYPE).toLowerCase() : '';
    const isGotenberg = type === 'gotenberg' || /\/forms\/libreoffice\/convert/i.test(endpoint) || /\/convert\/office/i.test(endpoint);
    let res;
    if(isGotenberg){
      // Gotenberg v7 expects multipart form with field name 'files'
      const fd = new FormData();
      fd.append('files', docxBlob, 'document.docx');
      res = await fetch(endpoint, { method: 'POST', body: fd });
    }else{
      // Generic raw POST of DOCX
      res = await fetch(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
        body: docxBlob
      });
    }
    if(!res.ok){
      throw new Error('PDF conversion failed: ' + res.status + ' ' + res.statusText);
    }
    const pdfBlob = await res.blob();
    const ct = (res.headers && typeof res.headers.get === 'function') ? (res.headers.get('Content-Type') || '') : '';
    if(!(pdfBlob && pdfBlob.size > 0)){
      throw new Error('PDF conversion returned empty blob');
    }
    return pdfBlob;
  }

  // Export as PDF by first generating DOCX, then converting through the configured endpoint
  window.exportProjectPdf = async function(p, opts){
    try{
      await ensureDocxDeps();

      // Build data object expected by the Word template placeholders (same as DOCX export)
      const latest = latestAccomplishment(p);
      const durationText = `${p.originalDuration || 0} days${p.timeExtension ? ' +' + p.timeExtension : ''}`;
      const photos = (p.photos && p.photos.length) ? p.photos : (p.sCurveDataUrl ? [p.sCurveDataUrl] : []);
      try{ console.info('[export] Photos available:', (photos||[]).length); }catch(_){/* ignore */}
      const altImages = {
        ProjectPhoto1: photos[0] || '',
        ProjectPhoto2: photos[1] || '',
        ProjectPhoto3: photos[2] || '',
        'Project Photo 1': photos[0] || '',
        'Project Photo 2': photos[1] || '',
        'Project Photo 3': photos[2] || ''
      };
      const data = {
        ProjectName: valOrNA(p.name),
        ImplementingAgency: valOrNA(p.implementingAgency),
        Contractor: valOrNA(p.contractor),
        Location: valOrNA(p.location),
        ContractAmount: (p.contractAmount !== undefined && p.contractAmount !== null && String(p.contractAmount) !== '') ? currency(p.contractAmount) : 'n/a',
        RevisedContractAmount: (p.revisedContractAmount !== undefined && p.revisedContractAmount !== null && String(p.revisedContractAmount) !== '') ? currency(p.revisedContractAmount) : 'n/a',
        Status: determineStatus(p),
        NTP: valOrNA(p.ntpDate),
        Duration: durationText,
        TargetCompletion: valOrNA(p.revisedCompletion || p.originalCompletion),
        TimeExtension: valOrNA(p.timeExtension),
        OriginalTargetCompletion: valOrNA(p.originalCompletion),
        RevisedTargetCompletion: valOrNA(p.revisedCompletion),
        PercentToDate: (latest.percent ?? 0).toFixed(2) + '%',
        PercentPlanned: (latest.plannedPercent ?? 0).toFixed(2) + '%',
        PercentPrevious: (latest.prevPercent ?? 0).toFixed(2) + '%',
        AsOfDate: valOrNA(latest.date),
        Issues: valOrNA(p.issues),
        Issue: valOrNA(p.issues),
        Actions: valOrNA(latest.action || p.action),
        ActionTaken: valOrNA(latest.action || p.action),
        Remarks: valOrNA(p.remarks),
        OtherDetails: valOrNA(p.otherDetails),
        OtherProjectDetails: valOrNA(p.otherDetails),
        Activities: valOrNA(latest.activities),
        ProjectPhoto1: photos[0] || '',
        ProjectPhoto2: photos[1] || '',
        ProjectPhoto3: photos[2] || '',
        'Project Photo 1': photos[0] || '',
        'Project Photo 2': photos[1] || '',
        'Project Photo 3': photos[2] || '',
        Variance: (((latest.percent ?? 0) - (latest.plannedPercent ?? 0)).toFixed(2)) + '%',
        accomplishments: (p.accomplishments || []).map(a=>({
          date: valOrNA(a.date),
          plannedPercent: (a.plannedPercent ?? 0).toFixed(2) + '%',
          prevPercent: (a.prevPercent ?? 0).toFixed(2) + '%',
          percent: (a.percent ?? 0).toFixed(2) + '%',
          variance: ((a.variance ?? ((a.percent ?? 0) - (a.plannedPercent ?? 0))).toFixed(2)) + '%',
          activities: valOrNA(a.activities),
          issue: valOrNA(a.issue || p.issues),
          action: valOrNA(a.action),
          remarks: valOrNA(a.remarks),
        }))
      };

      const docxBlob = await generateFromTemplate(data, altImages);
      let pdfBlob;
      try{
        pdfBlob = await convertDocxBlobToPdfUsingEndpoint(docxBlob);
      }catch(e){
        console.error('[export] PDF conversion error:', e);
        alert('PDF conversion is not configured. Set window.DOCX_PDF_ENDPOINT to a converter URL (e.g., a local LibreOffice/Gotenberg service). Exporting DOCX instead.');
        const fallbackName = `${(p.name || 'site_inspection_report').replace(/[^a-z0-9\- _()+]/gi,'_')}.docx`;
        saveAs(docxBlob, fallbackName);
        return;
      }

      const filename = `${(p.name || 'site_inspection_report').replace(/[^a-z0-9\- _()+]/gi,'_')}.pdf`;
      const wantOpen = !!(opts && opts.open);
      if(wantOpen){
        try{
          const url = URL.createObjectURL(pdfBlob);
          const w = window.open(url, '_blank', 'noopener');
          if(!w){ saveAs(pdfBlob, filename); }
          // Revoke later to allow the tab to load
          setTimeout(()=>{ try{ URL.revokeObjectURL(url); }catch(_){/* ignore */} }, 5000);
        }catch(_){
          saveAs(pdfBlob, filename);
        }
      }else{
        saveAs(pdfBlob, filename);
      }
    }catch(err){
      console.error(err);
      alert('Unable to export PDF. Details: ' + (err.message || err));
    }
  };

  async function urlToPngArrayBuffer(url){
    const res = await fetch(url, { cache: 'no-store' });
    if(!res.ok) throw new Error('Failed to fetch image: ' + res.status + ' ' + res.statusText);
    const blob = await res.blob();
    const dataURL = await new Promise((resolve,reject)=>{ const fr=new FileReader(); fr.onload=()=>resolve(fr.result); fr.onerror=reject; fr.readAsDataURL(blob); });
    return await dataUrlToPngArrayBuffer(String(dataURL));
  }

  function currency(amount){
    const n = Number(amount||0);
    return '₱' + n.toLocaleString(undefined,{ minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  // Return 'n/a' for empty/blank strings and null/undefined values; leave numbers/objects as-is
  function valOrNA(v){
    if(v === null || v === undefined) return 'n/a';
    if(typeof v === 'string'){
      const t = v.trim();
      return t ? v : 'n/a';
    }
    return v;
  }

  function determineStatus(p){
    try{
      const accomplishments = p.accomplishments || [];
      const latest = accomplishments.length > 0
        ? accomplishments.slice().sort((a,b)=> new Date(b.date) - new Date(a.date))[0]
        : null;
      if(latest){
        if((latest.percent ?? 0) >= 100) return 'Completed';
        if((latest.percent ?? 0) < (latest.plannedPercent ?? 0)) return 'Delayed';
      }
      const today = new Date().toISOString().split('T')[0];
      if(p.revisedCompletion && today > p.revisedCompletion) return 'Delayed';
      if(!p.revisedCompletion && today > p.originalCompletion) return 'Delayed';
      return 'On-going';
    }catch(_){ return 'On-going'; }
  }

  function latestAccomplishment(p){
    const arr = (p.accomplishments || []).slice().sort((a,b)=> new Date(b.date) - new Date(a.date));
    return arr[0] || {};
  }

  async function loadTemplateArrayBuffer(){
    const candidates = [];
    try{
      if(typeof window !== 'undefined' && typeof window.DOCX_TEMPLATE_URL === 'string' && window.DOCX_TEMPLATE_URL){
        candidates.push(window.DOCX_TEMPLATE_URL);
      }
    }catch(_){/* ignore */}
    // Primary configured path
    candidates.push(TEMPLATE_PATH);
    // Common alternate location (legacy)
    candidates.push('assets/templates/site_inspection_template.docx');
    // Fallbacks for common layouts
    candidates.push('site_inspection_template.docx');
    candidates.push('site_inspection_template.docx/site_inspection_template.docx');

    const tried = [];
    for(const url of candidates){
      try{
        const res = await fetch(url, { cache: 'no-store' });
        if(res.ok){
          console.debug('[export] Using template:', url);
          return await res.arrayBuffer();
        }
        tried.push(url + ' -> ' + res.status + ' ' + res.statusText);
      }catch(e){
        tried.push(url + ' -> ' + (e && e.message ? e.message : String(e)));
      }
    }
    throw new Error('Failed to load template from any path. Tried: ' + tried.join(' | '));
  }

  async function generateFromTemplate(data, altImages){
    const DocxTemplater = window.Docxtemplater || window.docxtemplater || (window.docxtemplater && window.docxtemplater.default);
    if(!(window.PizZip && DocxTemplater)){
      throw new Error('Docxtemplater or PizZip not loaded');
    }
    const content = await loadTemplateArrayBuffer();
    const zip = new window.PizZip(content);
    const doc = new DocxTemplater(zip, {
          paragraphLoop: true,
          linebreaks: true,
          // If a tag is missing in data, return 'n/a' (except for image tags which must be empty)
          nullGetter(part){
            try{
              const tag = (part && (part.tag || part.name)) || '';
              const imageTags = new Set(['ProjectPhoto1','ProjectPhoto2','ProjectPhoto3','Project Photo 1','Project Photo 2','Project Photo 3']);
              if(imageTags.has(tag)) return '';
            }catch(_){ /* ignore */ }
            return 'n/a';
          },
          modules: []
        });
    if(typeof doc.compile === 'function'){
      doc.compile();
    }
    try{
      if(typeof doc.resolveData === 'function'){
        await doc.resolveData(data);
      }else{
        doc.setData(data);
      }
      doc.render();
    }catch(err){
      let summary = '';
      try{
        const props = err && err.properties ? err.properties : {};
        const errors = Array.isArray(props.errors) ? props.errors : [];
        const summarized = errors.map((sub)=>{
          const sp = sub && sub.properties ? sub.properties : {};
          const tag = sp.tag || sp.name || sp.id || '';
          const exp = sp.explanation || sub.message || (err && err.message) || 'Unknown template error';
          const file = sp.file || props.file || 'document';
          return `[${file}] ${tag} -> ${exp}`;
        }).join(' | ');
        summary = summarized || '';
        console.error('Docxtemplater Render Error', err, summarized);
      }catch(_){
        console.error('Docxtemplater Render Error', err);
      }
      throw new Error(summary || (err && err.message) || 'Template render failed');
    }
    // Replace images by alt-text after rendering
    try{ await replaceAltTextImagesInZip(doc.getZip(), altImages || {}); }catch(_){/* non-fatal */}
    return doc.getZip().generate({ type: 'blob' });
  }

  // Expose global function used by script.js in the Project Details modal
  window.exportProjectDocx = async function(p, opts){
    try{
      // Ensure dependencies are present (handles slow CDN or first use)
      await ensureDocxDeps();
      // Alt-text replacement approach (no ImageModule/xmldom required)

      // Build data object expected by your Word template placeholders
      const latest = latestAccomplishment(p);
      const durationText = `${p.originalDuration || 0} days${p.timeExtension ? ' +' + p.timeExtension : ''}`;
      const photos = (p.photos && p.photos.length) ? p.photos : (p.sCurveDataUrl ? [p.sCurveDataUrl] : []);
      try{ console.info('[export] Photos available:', (photos||[]).length); }catch(_){/* ignore */}
      const altImages = {
        ProjectPhoto1: photos[0] || '',
        ProjectPhoto2: photos[1] || '',
        ProjectPhoto3: photos[2] || '',
        'Project Photo 1': photos[0] || '',
        'Project Photo 2': photos[1] || '',
        'Project Photo 3': photos[2] || ''
      };
      const data = {
        ProjectName: valOrNA(p.name),
        ImplementingAgency: valOrNA(p.implementingAgency),
        Contractor: valOrNA(p.contractor),
        Location: valOrNA(p.location),
        ContractAmount: (p.contractAmount !== undefined && p.contractAmount !== null && String(p.contractAmount) !== '') ? currency(p.contractAmount) : 'n/a',
        RevisedContractAmount: (p.revisedContractAmount !== undefined && p.revisedContractAmount !== null && String(p.revisedContractAmount) !== '') ? currency(p.revisedContractAmount) : 'n/a',
        Status: determineStatus(p),
        NTP: valOrNA(p.ntpDate),
        Duration: durationText,
        TargetCompletion: valOrNA(p.revisedCompletion || p.originalCompletion),
        // Aliases requested by user
        TimeExtension: valOrNA(p.timeExtension),
        OriginalTargetCompletion: valOrNA(p.originalCompletion),
        RevisedTargetCompletion: valOrNA(p.revisedCompletion),
        // Optional extras if your template has these
        PercentToDate: (latest.percent ?? 0).toFixed(2) + '%',
        PercentPlanned: (latest.plannedPercent ?? 0).toFixed(2) + '%',
        PercentPrevious: (latest.prevPercent ?? 0).toFixed(2) + '%',
        AsOfDate: valOrNA(latest.date),
        Issues: valOrNA(p.issues),
        Issue: valOrNA(p.issues), // alias
        Actions: valOrNA(latest.action || p.action),
        ActionTaken: valOrNA(latest.action || p.action), // alias
        Remarks: valOrNA(p.remarks),
        OtherDetails: valOrNA(p.otherDetails),
        OtherProjectDetails: valOrNA(p.otherDetails), // alias
        Activities: valOrNA(latest.activities),
        // Photos (kept in data only for potential text placeholders)
        ProjectPhoto1: photos[0] || '',
        ProjectPhoto2: photos[1] || '',
        ProjectPhoto3: photos[2] || '',
        'Project Photo 1': photos[0] || '',
        'Project Photo 2': photos[1] || '',
        'Project Photo 3': photos[2] || '',
        // Provide single snapshot variance as well if needed
        Variance: (((latest.percent ?? 0) - (latest.plannedPercent ?? 0)).toFixed(2)) + '%',
        // If you designed a table loop in Word like {#accomplishments}{date}{percent}{/accomplishments}
        accomplishments: (p.accomplishments || []).map(a=>({
          date: valOrNA(a.date),
          plannedPercent: (a.plannedPercent ?? 0).toFixed(2) + '%',
          prevPercent: (a.prevPercent ?? 0).toFixed(2) + '%',
          percent: (a.percent ?? 0).toFixed(2) + '%',
          variance: ((a.variance ?? ((a.percent ?? 0) - (a.plannedPercent ?? 0))).toFixed(2)) + '%',
          activities: valOrNA(a.activities),
          issue: valOrNA(a.issue || p.issues),
          action: valOrNA(a.action),
          remarks: valOrNA(a.remarks),
        }))
      };

      const blob = await generateFromTemplate(data, altImages);
      const filename = `${(p.name || 'site_inspection_report').replace(/[^a-z0-9\- _()+]/gi,'_')}.docx`;
      const wantOpen = !!(opts && opts.open);
      if(wantOpen && window.navigator && typeof window.navigator.msSaveOrOpenBlob === 'function'){
        // Best experience on Windows: prompts Open/Save and can open directly in Word
        window.navigator.msSaveOrOpenBlob(blob, filename);
      }else if(wantOpen){
        // Fallback: trigger download via blob URL
        try{
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.rel = 'noopener';
          // Not setting download may let the handler decide; most browsers will still download
          a.download = filename;
          document.body.appendChild(a);
          a.click();
          setTimeout(()=>{ try{ URL.revokeObjectURL(url); a.remove(); }catch(_){/* ignore */} }, 1500);
        }catch(_){
          saveAs(blob, filename);
        }
      }else{
        // Default: save as download
        saveAs(blob, filename);
      }
    }catch(err){
      console.error(err);
      alert('Unable to generate Word document. Please ensure the template exists at\n' + TEMPLATE_PATH + '\n\nDetails: ' + (err.message || err));
    }
  };
})();
