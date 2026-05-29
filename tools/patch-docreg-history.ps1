# Patch docreg.js to add Edit History write-and-render across modules
param(
  [string]$FilePath = "c:\Users\johnlowel.fradejas\CascadeProjects\construction-dashboard\docreg.js"
)

if(!(Test-Path $FilePath)){
  Write-Error "File not found: $FilePath"; exit 1
}

$orig = Get-Content -Raw -LiteralPath $FilePath
$back = "$FilePath.bak_$(Get-Date -Format yyyyMMdd_HHmmss)"
$orig | Set-Content -LiteralPath $back -Encoding UTF8

# 1) Replace generic Firestore set(...) during submit with history append logic
# Pattern: await ref.set({ id: ref.id, ...payload }, { merge: true });
$patSet = 'await\s+ref\.set\(\{\s*id:\s*ref\.id,\s*\.\.\.payload\s*\},\s*\{\s*merge:\s*true\s*\}\s*\);'
$repSet = @"
let __existingHistory=[]; try{ const snap = await ref.get(); __existingHistory = snap.exists ? ((snap.data()||{}).history || []) : []; }catch(_){}
const __user = firebase.auth().currentUser;
const __email = __user?.email || 'unknown';
const __full = (__user?.displayName || '').trim();
const __action = id ? 'edit' : 'create';
const __history = [...__existingHistory, { email: __email, fullName: __full, timestamp: new Date().toISOString(), action: __action }];
await ref.set({ id: ref.id, history: __history, ...payload }, { merge: true });
"@

$stage1 = [regex]::Replace($orig, $patSet, $repSet, 'Singleline')

# 2) Append history UI in each details view container assignment
function Add-HistoryUI {
  param([string]$content, [string]$container, [string]$afterAssign, [switch]$HasDataAttr)
  $assignPattern = [regex]::Escape("if($container) $afterAssign")
  $assignPattern = $assignPattern.Replace('\ ','\\s*')
  $assignPattern = $assignPattern.Replace('= html;','= html;')
  $histBlock = if($HasDataAttr){
    @"
if($container){ $afterAssign try{ const __hist = (it.history||[]).slice().sort((a,b)=> new Date(b.timestamp)-new Date(a.timestamp)); if(__hist.length){ const __rows = __hist.map(h=>{ const who=(h.fullName||'').trim() || (h.email||'unknown'); const when = (new Date(h.timestamp)).toLocaleString(); const act = h.action || 'update'; return `<div class=\"d-flex align-items-start gap-2 py-1\"><i class=\"fa-regular fa-clock mt-1 text-muted\"></i><div><div><strong>${act}</strong> by ${escapeHtml(who)}</div><div class=\"text-muted small\">${escapeHtml(when)}</div></div></div>`; }).join(''); $container.insertAdjacentHTML('beforeend', `<div class=\"small text-muted mt-3\">Edit History</div><div class=\"border rounded p-2 bg-light small\">${__rows}</div>`); } }catch(_){ } $container.setAttribute('data-current-id', it.id); }
"@
  } else {
    @"
if($container){ $afterAssign try{ const __hist = (it.history||[]).slice().sort((a,b)=> new Date(b.timestamp)-new Date(a.timestamp)); if(__hist.length){ const __rows = __hist.map(h=>{ const who=(h.fullName||'').trim() || (h.email||'unknown'); const when = (new Date(h.timestamp)).toLocaleString(); const act = h.action || 'update'; return `<div class=\"d-flex align-items-start gap-2 py-1\"><i class=\"fa-regular fa-clock mt-1 text-muted\"></i><div><div><strong>${act}</strong> by ${escapeHtml(who)}</div><div class=\"text-muted small\">${escapeHtml(when)}</div></div></div>`; }).join(''); $container.insertAdjacentHTML('beforeend', `<div class=\"small text-muted mt-3\">Edit History</div><div class=\"border rounded p-2 bg-light small\">${__rows}</div>`); } }catch(_){ } }
"@
  }
  $content -replace $assignPattern, [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $histBlock }
}

$stage2 = $stage1
# Outgoing
$stage2 = Add-HistoryUI -content $stage2 -container 'els.detailsBody' -afterAssign 'els.detailsBody.innerHTML = html;' 
# Incoming (has data-current-id attribute set in the same line)
$stage2 = $stage2 -replace [regex]::Escape("if(els.ilDetailsBody){ els.ilDetailsBody.innerHTML = html; els.ilDetailsBody.setAttribute('data-current-id', it.id); }"), (
  "if(els.ilDetailsBody){ els.ilDetailsBody.innerHTML = html; try{ const __hist = (it.history||[]).slice().sort((a,b)=> new Date(b.timestamp)-new Date(a.timestamp)); if(__hist.length){ const __rows = __hist.map(h=>{ const who=(h.fullName||'').trim() || (h.email||'unknown'); const when = (new Date(h.timestamp)).toLocaleString(); const act = h.action || 'update'; return `<div class=\"d-flex align-items-start gap-2 py-1\"><i class=\"fa-regular fa-clock mt-1 text-muted\"></i><div><div><strong>${act}</strong> by ${escapeHtml(who)}</div><div class=\"text-muted small\">${escapeHtml(when)}</div></div></div>`; }).join(''); els.ilDetailsBody.insertAdjacentHTML('beforeend', `<div class=\"small text-muted mt-3\">Edit History</div><div class=\"border rounded p-2 bg-light small\">${__rows}</div>`); } }catch(_){ } els.ilDetailsBody.setAttribute('data-current-id', it.id); }"
)
# Memo
$stage2 = Add-HistoryUI -content $stage2 -container 'els.moDetailsBody' -afterAssign 'els.moDetailsBody.innerHTML = html;'
# Office Order
$stage2 = Add-HistoryUI -content $stage2 -container 'els.ooDetailsBody' -afterAssign 'els.ooDetailsBody.innerHTML = html;'
# Travel Order
$stage2 = Add-HistoryUI -content $stage2 -container 'els.toDetailsBody' -afterAssign 'els.toDetailsBody.innerHTML = html;'
# Minutes of Meeting
$stage2 = Add-HistoryUI -content $stage2 -container 'els.miDetailsBody' -afterAssign 'els.miDetailsBody.innerHTML = html;'
# Notice of Meeting
$stage2 = Add-HistoryUI -content $stage2 -container 'els.noDetailsBody' -afterAssign 'els.noDetailsBody.innerHTML = html;'

# Write back
if($stage2 -ne $orig){
  $stage2 | Set-Content -LiteralPath $FilePath -Encoding UTF8
  Write-Host "Patched docreg.js. Backup saved to $back"
}else{
  Write-Warning "No changes applied. Patterns may not have matched. Backup: $back"
}
