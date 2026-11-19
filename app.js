/* ECIarc — app.js (generado desde web/js; no editar a mano) */
(function(){
// ---- 00-core.js ----
var ALL_ROWS = [];           // filas planas para buscador/escaner (derivadas de catalog.json)
var SHEETS_BY_SLUG = {};     // slug -> hoja de catalog.json
var APPLE_ON_THIS_DAY = {};  // efemerides de events.txt
var IMAGE_MANIFEST = {};
var panels = document.querySelectorAll('.panel');
var sheetSearch = document.getElementById('sheetSearch');
var sheetSearchClear = document.getElementById('sheetSearchClear');
var level1Buttons = document.querySelectorAll('.btn-row.level1 .nav-btn');
var lobRows = document.querySelectorAll('.btn-row.level2');
var leafButtons = document.querySelectorAll('.nav-btn-leaf');
var mainEl = document.querySelector('.main');
function setModalOpen(modalEl, open){
if(!modalEl) return;
modalEl.classList.toggle('open', open);
modalEl.setAttribute('aria-hidden', open ? 'false' : 'true');
}
var ESC_HANDLERS = [];
function onEscape(modalEl, closeFn){
if(modalEl){ ESC_HANDLERS.push({ el: modalEl, fn: closeFn }); }
}
document.addEventListener('keydown', function(e){
if(e.key !== 'Escape') return;
ESC_HANDLERS.forEach(function(h){ if(h.el.classList.contains('open')) h.fn(); });
});
function closeOnBackdrop(modalEl, closeFn){
if(!modalEl) return;
modalEl.addEventListener('click', function(e){ if(e.target === modalEl) closeFn(); });
}
var _scriptPromises = {};
function loadScriptOnce(url){
if(_scriptPromises[url]) return _scriptPromises[url];
_scriptPromises[url] = new Promise(function(resolve, reject){
var s = document.createElement('script');
s.src = url;
s.async = true;
s.onload = function(){ resolve(); };
s.onerror = function(){ delete _scriptPromises[url]; reject(new Error('No se pudo cargar ' + url)); };
document.head.appendChild(s);
});
return _scriptPromises[url];
}

// ---- 01-util.js ----
function escHtml(s){
return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function escAttr(s){
return escHtml(s).replace(/"/g, '&quot;').replace(/'/g, '&#x27;');
}

// ---- 10-catalog.js ----
var MONO_LABELS = {};
var ACC_MARKER_RE = /^(accesor|access)/;
var EAN_TEXT_RE = /^\d{8,14}$/;
function isAccessoryText(v){
return !!v && ACC_MARKER_RE.test(v.trim().toLowerCase());
}
function famStyleAttr(o){
return o.grad ? ' style="--fam-grad:' + o.grad + ';--fam-wash:' + o.wash + '"' : '';
}
function cellStyleAttr(c){
return c.bg ? ' style="background:' + c.bg + ';color:' + (c.fg || '#111318') + ';"' : '';
}
function rowEanForClick(row){
if(row.length < 4) return '';
var v = (row[3].v || '').trim();
return EAN_TEXT_RE.test(v) ? v : '';
}
function renderCells(row, monoIdx, clickEan, tdClassFn){
var out = '', colPos = 0;
for(var i = 0; i < row.length; i++){
var c = row[i];
var cls = tdClassFn ? tdClassFn(c, colPos) : (monoIdx[colPos] ? ' class="mono"' : '');
var span = c.span ? ' colspan="' + c.span + '"' : '';
var spanCls = (clickEan && i < 4) ? 'cell-text ean-clickable" title="Ver código de barras' : 'cell-text';
out += '<td' + cls + cellStyleAttr(c) + span + '><span class="' + spanCls + '">' +
(c.v ? escAttr(c.v) : '') + '</span></td>';
colPos += c.span || 1;
}
return out;
}
function renderGroup(grp, gi){
var parts = [];
var monoIdx = {};
grp.cols.forEach(function(c, i){ if(MONO_LABELS[c.trim().toLowerCase()]) monoIdx[i] = true; });
if(gi > 0){
parts.push('    <div class="group-divider"><span class="group-divider-text">' + escAttr(grp.title) + '</span></div>');
}
parts.push('    <div class="table-wrap"><table>');
parts.push('      <thead><tr>');
grp.cols.forEach(function(c){ parts.push('<th>' + escAttr(c) + '</th>'); });
parts.push('      </tr></thead>');
parts.push('      <tbody>');
grp.rows.forEach(function(row){
var c1v = row.length ? (row[0].v || '') : '';
var cls = isAccessoryText(c1v) ? ' class="accessory-marker-row"' : '';
var ean = rowEanForClick(row);
parts.push('<tr' + cls + (ean ? ' data-ean="' + ean + '"' : '') + '>');
parts.push(renderCells(row, monoIdx, !!ean));
parts.push('</tr>');
});
parts.push('      </tbody>');
parts.push('    </table></div>');
return parts.join('\n');
}
function renderCard(card){
var acc = card.acc ? ' accessory-card' : '';
var fam = card.grad ? ' fam-card' : '';
var pair = card.pair ? ' data-pair-card="' + card.pair + '"' : '';
var parts = [
'<div class="card' + acc + fam + '" id="' + card.id + '-card"' + famStyleAttr(card) + pair + '>',
'  <div class="card-header" data-toggle="' + card.id + '">',
'    <div class="card-title-wrap">',
'      <span class="chevron">▶</span>',
'      <h3 class="card-title' + (card.acc ? ' card-title-accessory' : '') + '">' + escAttr(card.title) + '</h3>',
'    </div>',
'  </div>',
'  <div class="card-body" id="' + card.id + '">'
];
card.groups.forEach(function(g, gi){ parts.push(renderGroup(g, gi)); });
parts.push('  </div>');
parts.push('</div>');
return parts.join('\n');
}
function renderMatrix(sheet){
var parts = [
'<div class="card matrix-card">',
'  <div class="table-wrap"><table class="matrix-table">',
'    <tbody>'
];
sheet.rows.forEach(function(row){
var anyVal = row.some(function(c){ return !!c.v; });
if(!anyVal){ parts.push('<tr class="matrix-spacer"><td colspan="99"></td></tr>'); return; }
var ean = rowEanForClick(row);
parts.push('<tr' + (ean ? ' data-ean="' + ean + '"' : '') + '>');
parts.push(renderCells(row, {}, !!ean, function(c){ return ' class="' + (c.bold ? 'th-cell' : 'td-cell') + '"'; }));
parts.push('</tr>');
});
parts.push('    </tbody>');
parts.push('  </table></div>');
parts.push('</div>');
return parts.join('\n');
}
function renderModelRow(models){
return '<div class="btn-row level3">' + models.map(function(m){
return '<button type="button" class="nav-btn nav-btn-model' + (m.grad ? ' nav-btn-fam' : '') + '"' +
famStyleAttr(m) + ' data-model-target="' + m.target + '">' + escAttr(m.label) + '</button>';
}).join('') + '</div>';
}
var MATRIX_HINT = '<div class="panel-toolbar"><p class="matrix-hint">Tabla de compatibilidad Apple Pencil / Magic Keyboard × modelo de iPad.</p></div>';
function renderSheet(sheet){
if(sheet.type === 'matrix') return { before: MATRIX_HINT, body: renderMatrix(sheet) };
return {
before: sheet.models ? renderModelRow(sheet.models) : '',
body: sheet.cards.map(renderCard).join('\n')
};
}
function ensurePanelRendered(slug){
var sheet = SHEETS_BY_SLUG[slug];
if(!sheet) return;
var panel = document.getElementById(slug);
if(!panel || panel.dataset.rendered) return;
var content = panel.querySelector('.panel-content');
if(!content) return;
var html = renderSheet(sheet);
if(html.before) content.insertAdjacentHTML('beforebegin', html.before);
content.innerHTML = html.body;
panel.dataset.rendered = '1';
}
function buildFlatRows(catalog){
var rows = [];
catalog.sheets.forEach(function(s){
if(s.type !== 'sections') return;
s.cards.forEach(function(card){
var main = card.groups[0];
var cols = main.cols;
var monoIdx = {};
cols.forEach(function(c, i){ if(MONO_LABELS[c.trim().toLowerCase()]) monoIdx[i] = true; });
card.groups.forEach(function(g){
g.rows.forEach(function(row){
var cells = [], parts = [s.folder, s.title, main.title], colPos = 0;
row.forEach(function(c){
var cell = {
label: colPos < cols.length ? cols[colPos] : '',
v: c.v || '',
bg: c.bg || null,
fg: c.fg || null,
mono: !!monoIdx[colPos]
};
if(c.span) cell.span = c.span;
cells.push(cell);
if(c.v) parts.push(c.v);
colPos += c.span || 1;
});
rows.push({
sheet: s.title,
sheetFolder: s.folder,
slug: s.slug,
section: main.title,
isAccessory: !!card.acc,
cells: cells,
search: parts.join(' ').toLowerCase(),
searchnum: parts.map(function(p){ return p.replace(/\D/g, ''); })
.filter(function(d){ return d.length >= 6; }).join(' ')
});
});
});
});
});
return rows;
}
function setCatalog(catalog){
MONO_LABELS = {};
(catalog.monoLabels || []).forEach(function(l){ MONO_LABELS[l] = true; });
SHEETS_BY_SLUG = {};
catalog.sheets.forEach(function(s){ SHEETS_BY_SLUG[s.slug] = s; });
ALL_ROWS = buildFlatRows(catalog);
setSearchTerms(catalog.terms || {}, catalog.termMaxWords || 1);
}

// ---- 20-ui.js ----
var homeTodayDate = document.getElementById('homeTodayDate');
var eventsToggle = document.getElementById('eventsToggle');
var eventsModal = document.getElementById('eventsModal');
var eventsCloseBtn = document.getElementById('eventsCloseBtn');
var eventsModalDate = document.getElementById('eventsModalDate');
var eventsModalBody = document.getElementById('eventsModalBody');
function openEventsModal(){ setModalOpen(eventsModal, true); }
function closeEventsModal(){ setModalOpen(eventsModal, false); }
var today = new Date();
var todayDD = String(today.getDate()).padStart(2, '0');
var todayMM = String(today.getMonth() + 1).padStart(2, '0');
var todayLabel = (function(){
var weekday = today.toLocaleDateString('es-ES', { weekday: 'long' });
weekday = weekday.charAt(0).toUpperCase() + weekday.slice(1);
return weekday + ' ' + todayDD + '/' + todayMM + '/' + today.getFullYear();
})();
if(homeTodayDate){ homeTodayDate.textContent = todayLabel; }
function initTodayEvents(){
var todayEvents = APPLE_ON_THIS_DAY[todayMM + '-' + todayDD];
if(!(todayEvents && todayEvents.length)){
if(eventsToggle){ eventsToggle.style.display = 'none'; }
return;
}
if(eventsModalDate){ eventsModalDate.textContent = todayLabel; }
if(eventsModalBody){
var items = todayEvents.map(function(ev){
return '<li><strong>' + todayDD + '/' + todayMM + '/' + ev.y + '</strong>: ' + escHtml(ev.t) + '</li>';
}).join('');
eventsModalBody.innerHTML = '<ul class="events-modal-list">' + items + '</ul>';
}
if(eventsToggle){
eventsToggle.style.display = '';
eventsToggle.addEventListener('click', function(){
if(eventsModal && eventsModal.classList.contains('open')) closeEventsModal();
else openEventsModal();
});
}
}
function parseEvents(raw){
var events = {};
raw.split('\n').forEach(function(line){
line = line.trim();
if(!line || line.charAt(0) === '#') return;
var parts = line.split(' | ');
if(parts.length < 3) return;
var date = parts[0].trim();
var year = parseInt(parts[1].trim(), 10);
var text = parts.slice(2).join(' | ').trim();
if(!events[date]) events[date] = [];
events[date].push({ y: year, t: text });
});
return events;
}
if(eventsCloseBtn){ eventsCloseBtn.addEventListener('click', closeEventsModal); }
closeOnBackdrop(eventsModal, closeEventsModal);
onEscape(eventsModal, closeEventsModal);
var homeHelpLink = document.getElementById('homeHelpLink');
var helpModal = document.getElementById('helpModal');
var helpCloseBtn = document.getElementById('helpCloseBtn');
function openHelpModal(){ setModalOpen(helpModal, true); }
function closeHelpModal(){ setModalOpen(helpModal, false); }
if(homeHelpLink){
homeHelpLink.addEventListener('click', function(e){ e.preventDefault(); openHelpModal(); });
}
if(helpCloseBtn){ helpCloseBtn.addEventListener('click', closeHelpModal); }
closeOnBackdrop(helpModal, closeHelpModal);
onEscape(helpModal, closeHelpModal);
var EMAILJS_URL = 'https://cdn.jsdelivr.net/npm/@emailjs/browser@4/dist/email.min.js';
var EMAILJS_PUBLIC_KEY = 'LUJwdk3MQpdADb5Nm';
var EMAILJS_SERVICE_ID = 'service_jqp3osl';
var EMAILJS_TEMPLATE_ID = 'template_1146g92';
var emailjsReady = null;
function ensureEmailJs(){
if(emailjsReady) return emailjsReady;
emailjsReady = loadScriptOnce(EMAILJS_URL).then(function(){
if(window.emailjs && typeof window.emailjs.init === 'function'){
window.emailjs.init({ publicKey: EMAILJS_PUBLIC_KEY });
}
return window.emailjs;
}).catch(function(err){ emailjsReady = null; throw err; });
return emailjsReady;
}
var homeIncidentLink = document.getElementById('homeIncidentLink');
var incidentModal = document.getElementById('incidentModal');
var incidentCloseBtn = document.getElementById('incidentCloseBtn');
var incidentText = document.getElementById('incidentText');
var incidentSendBtn = document.getElementById('incidentSendBtn');
var incidentStatus = document.getElementById('incidentStatus');
function setIncidentStatus(text, cls){
if(!incidentStatus) return;
incidentStatus.textContent = text;
incidentStatus.className = 'incident-status' + (cls ? ' ' + cls : '');
}
function openIncidentModal(){
setIncidentStatus('');
if(incidentSendBtn){ incidentSendBtn.disabled = false; incidentSendBtn.textContent = 'Enviar'; }
setModalOpen(incidentModal, true);
ensureEmailJs().catch(function(){});
if(incidentText){ setTimeout(function(){ incidentText.focus(); }, 50); }
}
function closeIncidentModal(){
setModalOpen(incidentModal, false);
if(incidentText){ incidentText.value = ''; }
}
if(homeIncidentLink){
homeIncidentLink.addEventListener('click', function(e){ e.preventDefault(); openIncidentModal(); });
}
if(incidentCloseBtn){ incidentCloseBtn.addEventListener('click', closeIncidentModal); }
closeOnBackdrop(incidentModal, closeIncidentModal);
onEscape(incidentModal, closeIncidentModal);
if(incidentSendBtn){
incidentSendBtn.addEventListener('click', function(){
var msg = incidentText ? incidentText.value.trim() : '';
if(!msg){ setIncidentStatus('Escribe algo antes de enviar.', 'err'); return; }
incidentSendBtn.disabled = true;
incidentSendBtn.textContent = 'Enviando...';
setIncidentStatus('');
ensureEmailJs().then(function(emailjs){
if(!emailjs) throw new Error('sin emailjs');
return emailjs.send(EMAILJS_SERVICE_ID, EMAILJS_TEMPLATE_ID, { message: msg });
}).then(function(){
setIncidentStatus('Enviado. ¡Gracias!', 'ok');
incidentSendBtn.textContent = 'Enviado';
if(incidentText){ incidentText.value = ''; }
setTimeout(closeIncidentModal, 1400);
}).catch(function(){
setIncidentStatus('No se pudo enviar. Intentalo de nuevo.', 'err');
incidentSendBtn.disabled = false;
incidentSendBtn.textContent = 'Enviar';
});
});
}
var themeToggle = document.getElementById('themeToggle');
var statusBarMeta = document.getElementById('statusBarMeta');
var themeColorMeta = document.getElementById('themeColorMeta');
var THEME_KEY = 'catalogTheme';
function applyTheme(theme){
document.body.classList.toggle('light-mode', theme === 'light');
if(themeToggle){ themeToggle.textContent = theme === 'light' ? '🌙' : '☀️'; }
if(statusBarMeta){ statusBarMeta.setAttribute('content', theme === 'light' ? 'default' : 'black'); }
if(themeColorMeta){ themeColorMeta.setAttribute('content', theme === 'light' ? '#ffffff' : '#131620'); }
}
(function(){
var saved = null;
try{ saved = localStorage.getItem(THEME_KEY); }catch(e){}
applyTheme(saved === 'dark' ? 'dark' : 'light');
})();
if(themeToggle){
themeToggle.addEventListener('click', function(){
var next = document.body.classList.contains('light-mode') ? 'dark' : 'light';
applyTheme(next);
try{ localStorage.setItem(THEME_KEY, next); }catch(e){}
});
}
var FONT_SCALE_KEY = 'fontScale';
var FONT_SCALE_MIN = 0.8, FONT_SCALE_MAX = 1.4, FONT_SCALE_STEP = 0.1;
var fontDecBtn = document.getElementById('fontDecBtn');
var fontIncBtn = document.getElementById('fontIncBtn');
var fontScale = 1;
try{
var storedScale = parseFloat(localStorage.getItem(FONT_SCALE_KEY));
if(!isNaN(storedScale)){ fontScale = storedScale; }
}catch(e){}
function applyFontScale(){
document.documentElement.style.setProperty('--ref-zoom', fontScale);
if(fontIncBtn){ fontIncBtn.disabled = fontScale >= FONT_SCALE_MAX - 1e-9; }
if(fontDecBtn){ fontDecBtn.disabled = fontScale <= FONT_SCALE_MIN + 1e-9; }
}
function setFontScale(next){
fontScale = Math.min(FONT_SCALE_MAX, Math.max(FONT_SCALE_MIN, Math.round(next * 100) / 100));
try{ localStorage.setItem(FONT_SCALE_KEY, fontScale); }catch(e){}
applyFontScale();
}
applyFontScale();
if(fontIncBtn){ fontIncBtn.addEventListener('click', function(){ setFontScale(fontScale + FONT_SCALE_STEP); }); }
if(fontDecBtn){ fontDecBtn.addEventListener('click', function(){ setFontScale(fontScale - FONT_SCALE_STEP); }); }
function toggleFontButtons(show){
var disp = show ? 'flex' : 'none';
if(fontIncBtn){ fontIncBtn.style.display = disp; }
if(fontDecBtn){ fontDecBtn.style.display = disp; }
}

// ---- 30-nav.js ----
var searchToggle = document.getElementById('searchToggle');
var searchBarRow = document.getElementById('searchBarRow');
searchToggle.addEventListener('click', function(){
var open = searchBarRow.classList.toggle('open');
if(open){ sheetSearch.focus(); }
});
function setActiveLevel1(el){
level1Buttons.forEach(function(b){ b.classList.remove('active'); });
if(el){ el.classList.add('active'); }
}
function setActiveLeaf(slug){
leafButtons.forEach(function(b){ b.classList.toggle('active', b.dataset.target === slug); });
}
function hideAllLobRows(){
lobRows.forEach(function(r){ r.classList.remove('open'); });
}
function showLobRow(lobSlug){
lobRows.forEach(function(r){ r.classList.toggle('open', r.id === 'lobrow-' + lobSlug); });
}
function findLobForSlug(slug){
var leaf = document.querySelector('.nav-btn-leaf[data-target="' + slug + '"]');
if(leaf){
var row = leaf.closest('.btn-row.level2');
if(row){ return row.id.replace('lobrow-', ''); }
}
return null;
}
function syncNav(slug){
if(slug === 'home' || slug === 'search-results'){
setActiveLevel1(null);
hideAllLobRows();
setActiveLeaf(null);
return;
}
var directBtn = document.querySelector('.btn-row.level1 .nav-btn[data-target="' + slug + '"]');
if(directBtn){
setActiveLevel1(directBtn);
hideAllLobRows();
setActiveLeaf(null);
return;
}
var lobSlug = findLobForSlug(slug);
if(lobSlug){
var lobBtn = document.querySelector('.btn-row.level1 .nav-btn[data-lob="' + lobSlug + '"]');
setActiveLevel1(lobBtn);
showLobRow(lobSlug);
setActiveLeaf(slug);
}
}
function resetToHome(opts){
var keepSearchOpen = !!(opts && opts.keepSearchOpen);
panels.forEach(function(p){ p.classList.toggle('active', p.id === 'home'); });
level1Buttons.forEach(function(b){ b.classList.remove('active'); });
lobRows.forEach(function(r){ r.classList.remove('open'); });
leafButtons.forEach(function(b){ b.classList.remove('active'); });
collapseAllCards();
if(!keepSearchOpen){
if(searchDebounceTimer){ clearTimeout(searchDebounceTimer); searchDebounceTimer = null; }
if(sheetSearch){ sheetSearch.value = ''; sheetSearch.blur(); }
if(sheetSearchClear){ sheetSearchClear.style.display = 'none'; }
if(searchBarRow){ searchBarRow.classList.remove('open'); }
}
clearSearchResults();
toggleFontButtons(false);
if(history.replaceState){ history.replaceState(null, '', '#home'); }
window.scrollTo({top: 0, behavior: 'smooth'});
}
function activate(slug, opts){
if(slug === 'home'){ resetToHome(opts); return; }
ensurePanelRendered(slug);
collapseAllCards();
panels.forEach(function(p){ p.classList.toggle('active', p.id === slug); });
syncNav(slug);
toggleFontButtons(true);
if(history.replaceState){ history.replaceState(null, '', '#' + slug); }
}
function activePanelId(){
var p = document.querySelector('.panel.active');
return p ? p.id : null;
}
function collapseAllCards(){
document.querySelectorAll('.card.open').forEach(function(c){ c.classList.remove('open'); });
document.querySelectorAll('.nav-btn-model.active').forEach(function(b){ b.classList.remove('active'); });
}
function deactivateAllPanels(){
panels.forEach(function(p){ p.classList.remove('active'); });
}
level1Buttons.forEach(function(btn){
btn.addEventListener('click', function(e){
e.preventDefault();
var wasActive = btn.classList.contains('active');
if(btn.dataset.lob){
if(wasActive && !activePanelId()){ resetToHome(); return; }
collapseAllCards();
setActiveLevel1(btn);
showLobRow(btn.dataset.lob);
setActiveLeaf(null);
deactivateAllPanels();
if(history.replaceState){ history.replaceState(null, '', '#'); }
} else if(btn.dataset.target){
if(wasActive){ resetToHome(); return; }
hideAllLobRows();
activate(btn.dataset.target);
window.scrollTo({top:0, behavior:'smooth'});
}
});
});
var brandHome = document.getElementById('brandHome');
if(brandHome){
brandHome.addEventListener('click', function(e){ e.preventDefault(); resetToHome(); });
brandHome.addEventListener('keydown', function(e){
if(e.key === 'Enter' || e.key === ' ' || e.key === 'Spacebar'){ e.preventDefault(); resetToHome(); }
});
}
leafButtons.forEach(function(btn){
btn.addEventListener('click', function(e){
e.preventDefault();
if(btn.dataset.target){
activate(btn.dataset.target);
window.scrollTo({top:0, behavior:'smooth'});
}
});
});
function syncModelButtonActive(card){
if(!card) return;
var btn = document.querySelector('.nav-btn-model[data-model-target="' + card.id + '"]');
if(btn){ btn.classList.toggle('active', card.classList.contains('open')); }
}
function setCardOpen(card, open){
if(!card) return;
card.classList.toggle('open', open);
syncModelButtonActive(card);
if(!open && card.dataset.pairCard){
var partner = document.getElementById(card.dataset.pairCard);
if(partner){
partner.classList.remove('open');
syncModelButtonActive(partner);
}
}
}
mainEl.addEventListener('click', function(e){
var modelBtn = e.target.closest('.nav-btn-model');
if(modelBtn){
e.preventDefault();
var card = document.getElementById(modelBtn.dataset.modelTarget);
if(!card) return;
if(card.classList.contains('open')){ setCardOpen(card, false); return; }
setCardOpen(card, true);
var topbar = document.querySelector('.topbar');
var headerOffset = topbar ? topbar.offsetHeight + 12 : 60;
var top = card.getBoundingClientRect().top + window.pageYOffset - headerOffset;
window.scrollTo({top: Math.max(0, top), behavior:'smooth'});
return;
}
var header = e.target.closest('[data-toggle]');
if(header){
var c = header.closest('.card');
if(c){ setCardOpen(c, !c.classList.contains('open')); }
}
});

// ---- 40-search.js ----
var TERM_SYNONYMS = {};
var TERM_MAX_WORDS = 1;
var TERM_KEYS = [];
function setSearchTerms(terms, maxWords){
TERM_SYNONYMS = terms;
TERM_MAX_WORDS = maxWords;
TERM_KEYS = Object.keys(terms);
}
var TERM_PREFIX_MIN = 2;
var STOPWORDS = {
'de':1,'del':1,'con':1,'y':1,'o':1,'u':1,'e':1,'para':1,'por':1,
'el':1,'la':1,'los':1,'las':1,'un':1,'una':1,'unos':1,'unas':1,
'en':1,'al':1,'a':1,'sin':1,'su':1,'sus':1,'lo':1,
'the':1,'of':1,'with':1,'and':1,'for':1
};
function stripAccents(s){
return String(s).toLowerCase()
.replace(/[áàäâã]/g,'a').replace(/[éèëê]/g,'e')
.replace(/[íìïî]/g,'i').replace(/[óòöôõ]/g,'o')
.replace(/[úùüû]/g,'u').replace(/ñ/g,'n').replace(/ç/g,'c');
}
function foldES(s){
var words = stripAccents(s).split(/\s+/), out = [];
for(var i=0;i<words.length;i++){
var w = words[i];
if(!w){ continue; }
if(/^[a-z]+$/.test(w)){
if(w.length > 3 && w.charAt(w.length-1) === 's'){ w = w.slice(0,-1); }
if(w.length > 3 && w.charAt(w.length-1) === 'a'){ w = w.slice(0,-1) + 'o'; }
}
out.push(w);
}
return out.join(' ');
}
var _altReCache = {};
function altMatches(hay, alt){
var re = _altReCache[alt];
if(!re){
re = new RegExp('(^|[^a-z0-9])' + escapeRegExp(alt) + '([^a-z0-9]|$)');
_altReCache[alt] = re;
}
return re.test(hay);
}
function buildGroups(raw){
var endsWithSpace = /\s$/.test(raw);
var lit = raw.toLowerCase().split(/\s+/).filter(Boolean);
var acc = stripAccents(raw).split(/\s+/).filter(Boolean);
var fold = foldES(raw).split(/\s+/).filter(Boolean);
var groups = [], i = 0;
while(i < fold.length){
var hit = null, end = i + 1;
var maxJ = Math.min(fold.length, i + TERM_MAX_WORDS);
for(var j = maxJ; j > i; j--){
var phrase = fold.slice(i, j).join(' ');
var vals = TERM_SYNONYMS[phrase];
if(vals){ hit = vals; end = j; break; }
var phraseAcc = acc.slice(i, j).join(' ');
var atTail = (j === fold.length) && !endsWithSpace;
if(atTail && phraseAcc.length >= TERM_PREFIX_MIN){
var merged = null;
for(var k = 0; k < TERM_KEYS.length; k++){
var tk = TERM_KEYS[k];
if(tk.indexOf(phrase) === 0 || tk.indexOf(phraseAcc) === 0){
merged = merged || [];
var cand = TERM_SYNONYMS[tk];
for(var m = 0; m < cand.length; m++){
if(merged.indexOf(cand[m]) === -1){ merged.push(cand[m]); }
}
}
}
if(merged){ hit = merged; end = j; break; }
}
}
if(hit === null && end === i + 1){
var sw = acc.slice(i, end).join(' ');
var tail = (end === fold.length) && !endsWithSpace;
if(!tail && STOPWORDS[sw]){ i = end; continue; }
}
groups.push({
lit: lit.slice(i, end).join(' '),
acc: acc.slice(i, end).join(' '),
fold: fold.slice(i, end).join(' '),
alts: hit
});
i = end;
}
return groups;
}
function escapeRegExp(s){ return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
function buildHighlightRegex(tokens){
if(!tokens || !tokens.length){ return null; }
var sorted = tokens.slice().sort(function(a, b){ return b.length - a.length; });
var pattern = sorted.map(escapeRegExp).join('|');
return pattern ? new RegExp('(' + pattern + ')', 'ig') : null;
}
function highlightText(text, re){
if(!re){ return text; }
return String(text).replace(re, '<mark>$1</mark>');
}
var resultsBody = document.getElementById('searchResultsBody');
var resultsCount = document.getElementById('searchResultsCount');
function clearSearchResults(){
if(resultsBody){ resultsBody.innerHTML = ''; }
if(resultsCount){ resultsCount.textContent = ''; }
var existingEmpty = document.querySelector('#search-results .no-results');
if(existingEmpty){ existingEmpty.remove(); }
}
function renderResults(term, tokens){
resultsBody.innerHTML = '';
var groups = buildGroups(term);
var hlTokens = [];
groups.forEach(function(g){
if(g.lit && hlTokens.indexOf(g.lit) === -1){ hlTokens.push(g.lit); }
if(g.alts){
g.alts.forEach(function(a){ if(hlTokens.indexOf(a) === -1){ hlTokens.push(a); } });
}
});
var hlRe = buildHighlightRegex(hlTokens);
var matches = ALL_ROWS.filter(function(r){
var hay = r._sa || (r._sa = stripAccents(r.search));
return groups.every(function(g){
if (hay.indexOf(g.acc) !== -1) return true;
if (g.fold) {
var hayF = r._sf || (r._sf = foldES(r.search));
if (hayF.indexOf(g.fold) !== -1) return true;
}
if (g.alts) {
for (var i = 0; i < g.alts.length; i++) {
if (altMatches(hay, stripAccents(g.alts[i]))) return true;
}
}
if (/^[\d.\-\s·,/]+$/.test(g.lit)) {
var td = g.lit.replace(/\D/g, '');
if (td.length >= 6 && r.searchnum.indexOf(td) !== -1) return true;
}
return false;
});
});
resultsCount.textContent = matches.length + ' referencia' + (matches.length === 1 ? '' : 's') + ' encontrada' + (matches.length === 1 ? '' : 's') + ' para "' + term + '"';
var existingEmpty = document.querySelector('#search-results .no-results');
if(existingEmpty){ existingEmpty.remove(); }
if(matches.length === 0){
var emptyRow = document.createElement('div');
emptyRow.className = 'no-results';
emptyRow.textContent = 'Sin resultados para "' + term + '"';
resultsBody.parentElement.parentElement.appendChild(emptyRow);
return;
}
matches.forEach(function(r){
var tr = document.createElement('tr');
if(r.isAccessory){ tr.classList.add('result-accessory'); }
tr.dataset.sheetFolder = r.sheetFolder || '';
var eanCell = r.cells.filter(function(c){ return c.label && c.label.trim().toUpperCase() === 'EAN'; })[0];
var rowEan = (eanCell && eanCell.v && EAN_TEXT_RE.test(eanCell.v.trim())) ? eanCell.v.trim() : '';
tr.dataset.ean = rowEan;
var eciCell = r.cells.filter(function(c){ return c.label && c.label.trim().toUpperCase() === 'ECI'; })[0];
tr.dataset.eci = (eciCell && eciCell.v) ? eciCell.v.trim() : '';
var tdSheet = document.createElement('td');
tdSheet.className = 'result-sheet-cell';
var link = document.createElement('a');
link.href = '#' + r.slug;
link.className = 'result-sheet-link';
link.innerHTML = highlightText(r.sheet, hlRe);
var secSpan = document.createElement('div');
secSpan.className = 'result-section-name';
secSpan.innerHTML = (r.isAccessory ? '🔌 ' : '') + highlightText(r.section, hlRe);
link.addEventListener('click', function(e){
e.preventDefault();
activate(r.slug);
window.scrollTo({top:0, behavior:'smooth'});
});
tdSheet.appendChild(link);
tdSheet.appendChild(secSpan);
tr.appendChild(tdSheet);
r.cells.forEach(function(c){
var td = document.createElement('td');
if(c.span){ td.colSpan = c.span; }
if(c.mono){ td.classList.add('mono'); }
if(c.bg){
td.style.background = c.bg;
td.style.color = c.fg || '#111318';
}
var wrap = document.createElement('div');
wrap.className = 'result-cell-inner';
var labelSpan = document.createElement('span');
labelSpan.className = 'result-cell-label';
labelSpan.textContent = c.label;
var valueSpan = document.createElement('span');
valueSpan.className = 'result-cell-value';
valueSpan.innerHTML = c.v ? highlightText(c.v, hlRe) : '';
if (rowEan) { valueSpan.classList.add('ean-clickable'); valueSpan.title = 'Ver código de barras'; }
wrap.appendChild(labelSpan);
wrap.appendChild(valueSpan);
td.appendChild(wrap);
tr.appendChild(td);
});
resultsBody.appendChild(tr);
});
}
var searchDebounceTimer = null;
sheetSearch.addEventListener('input', function(){
if(sheetSearchClear){ sheetSearchClear.style.display = sheetSearch.value ? 'flex' : 'none'; }
if(searchDebounceTimer){ clearTimeout(searchDebounceTimer); }
searchDebounceTimer = setTimeout(runSearch, 130);
});
function runSearch(){
var raw = sheetSearch.value.trim();
var term = raw.toLowerCase();
var tokens = term.split(/\s+/).filter(Boolean);
if(!tokens.length){
var existingEmpty = document.querySelector('#search-results .no-results');
if(existingEmpty){ existingEmpty.remove(); }
if(document.getElementById('search-results').classList.contains('active')){
activate('home');
}
return;
}
renderResults(raw, tokens);
activate('search-results');
}
if(sheetSearchClear){
sheetSearchClear.addEventListener('click', function(){
if(searchDebounceTimer){ clearTimeout(searchDebounceTimer); searchDebounceTimer = null; }
sheetSearch.value = '';
sheetSearchClear.style.display = 'none';
var existingEmpty = document.querySelector('#search-results .no-results');
if(existingEmpty){ existingEmpty.remove(); }
if(document.getElementById('search-results').classList.contains('active')){
activate('home', {keepSearchOpen:true});
}
if(searchBarRow){ searchBarRow.classList.add('open'); }
sheetSearch.focus();
});
}

// ---- 50-barcode.js ----
function parseEAN13(code) {
var raw = String(code).replace(/\D/g, '');
var s = raw;
if (s.length === 12) { s = '0' + s; }
if (s.length !== 13) return null;
var sum = 0;
for (var i = 0; i < 12; i++) {
sum += parseInt(s[i], 10) * (i % 2 === 0 ? 1 : 3);
}
var check = (10 - (sum % 10)) % 10;
if (check !== parseInt(s[12], 10)) { s = s.slice(0, 12) + check; }
var L = ["0001101","0011001","0010011","0111101","0100011","0110001","0101111","0111011","0110111","0001011"];
var G = ["0100111","0110011","0011011","0100001","0011101","0111001","0000101","0010001","0001001","0010111"];
var R = ["1110010","1100110","1101100","1000010","1011100","1001110","1010000","1000100","1001000","1110100"];
var P = ["LLLLLL","LLGLGG","LLGGLG","LLGGGL","LGLLGG","LGGLLG","LGGGLL","LGLGLG","LGLGGL","LGGLGL"];
var first = parseInt(s[0], 10);
var pattern = P[first];
var leftBits = "";
for (var j = 1; j <= 6; j++) {
var dL = parseInt(s[j], 10);
leftBits += (pattern[j - 1] === 'L') ? L[dL] : G[dL];
}
var rightBits = "";
for (var k = 7; k <= 12; k++) {
var dR = parseInt(s[k], 10);
rightBits += R[dR];
}
return { firstDigit: s[0], leftDigits: s.slice(1, 7), rightDigits: s.slice(7, 13), leftBits: leftBits, rightBits: rightBits, isUPCA: raw.length === 12 };
}
function renderBarcodeSVG(rawCode) {
var svg = document.getElementById('barcodeSvg');
if (!svg) return;
var clean = String(rawCode).trim();
var ean = parseEAN13(clean);
var moduleWidth = 2.4;
var quietLeft = 11 * moduleWidth;
var quietRight = 9 * moduleWidth;
var barcodeModules = 95;
var totalWidth = quietLeft + (barcodeModules * moduleWidth) + quietRight;
var totalHeight = 115;
var barHeightNormal = 75;
var barHeightGuard = 87;
svg.setAttribute('viewBox', '0 0 ' + totalWidth + ' ' + totalHeight);
svg.setAttribute('width', '100%');
svg.setAttribute('height', 'auto');
svg.style.maxWidth = '340px';
if (!ean) {
svg.innerHTML = '<rect width="' + totalWidth + '" height="' + totalHeight + '" fill="#ffffff"/>' +
'<text x="' + (totalWidth / 2) + '" y="' + (totalHeight / 2) + '" font-family="monospace" font-size="14" text-anchor="middle" fill="#111318">' + clean + '</text>';
return;
}
var out = '<rect width="' + totalWidth + '" height="' + totalHeight + '" fill="#ffffff"/>';
var currentX = quietLeft;
function drawModules(bits, isGuard) {
var h = isGuard ? barHeightGuard : barHeightNormal;
for (var i = 0; i < bits.length; i++) {
if (bits[i] === '1') {
out += '<rect x="' + currentX.toFixed(2) + '" y="8" width="' + moduleWidth.toFixed(2) + '" height="' + h + '" fill="#000000"/>';
}
currentX += moduleWidth;
}
}
drawModules("101", true);
drawModules(ean.leftBits, false);
drawModules("01010", true);
drawModules(ean.rightBits, false);
drawModules("101", true);
var textY = 104;
var fontSize = 14;
var monoFont = "'SF Mono', 'Menlo', 'Consolas', monospace";
out += '<text x="' + (quietLeft - (moduleWidth * 3.5)).toFixed(2) + '" y="' + textY + '" font-family="' + monoFont + '" font-size="' + fontSize + '" font-weight="700" text-anchor="middle" fill="#000000">' + ean.firstDigit + '</text>';
var leftGroupCenter = quietLeft + (3 * moduleWidth) + (21 * moduleWidth);
out += '<text x="' + leftGroupCenter.toFixed(2) + '" y="' + textY + '" font-family="' + monoFont + '" font-size="' + fontSize + '" font-weight="600" letter-spacing="2" text-anchor="middle" fill="#000000">' + ean.leftDigits + '</text>';
var rightGroupCenter = quietLeft + ((3 + 42 + 5) * moduleWidth) + (21 * moduleWidth);
out += '<text x="' + rightGroupCenter.toFixed(2) + '" y="' + textY + '" font-family="' + monoFont + '" font-size="' + fontSize + '" font-weight="600" letter-spacing="2" text-anchor="middle" fill="#000000">' + ean.rightDigits + '</text>';
out += '<text x="' + (totalWidth - (moduleWidth * 3)).toFixed(2) + '" y="' + textY + '" font-family="' + monoFont + '" font-size="12" font-weight="700" text-anchor="middle" fill="#000000">&gt;</text>';
svg.innerHTML = out;
}
var barcodeModal = document.getElementById('barcodeModal');
var barcodeCloseBtn = document.getElementById('barcodeCloseBtn');
var barcodeShareBtn = document.getElementById('barcodeShareBtn');
var barcodeModalTitle = document.getElementById('barcodeModalTitle');
var barcodeProductImgWrap = document.getElementById('barcodeProductImgWrap');
var barcodeProductImg = document.getElementById('barcodeProductImg');
function hasProductImage(sheetFolder, eanCode) {
if (!sheetFolder || !eanCode) return false;
var bucket = IMAGE_MANIFEST[sheetFolder];
return !!bucket && bucket[eanCode] === 1;
}
function setImageManifest(raw){
IMAGE_MANIFEST = {};
Object.keys(raw || {}).forEach(function(folder){
var set = {};
(raw[folder] || []).forEach(function(e){ set[e] = 1; });
IMAGE_MANIFEST[folder] = set;
});
}
function productImageUrl(sheetFolder, eanCode) {
return 'pictures/' + encodeURIComponent(sheetFolder) + '/' + encodeURIComponent(eanCode) + '.webp';
}
function findProductImageUrl(sheetFolder, eanCode) {
if (hasProductImage(sheetFolder, eanCode)) return productImageUrl(sheetFolder, eanCode);
var folders = Object.keys(IMAGE_MANIFEST);
for (var i = 0; i < folders.length; i++) {
if (IMAGE_MANIFEST[folders[i]][eanCode] === 1) return productImageUrl(folders[i], eanCode);
}
return null;
}
var IMG_CACHE = {};
function prefetchProductImage(sheetFolder, eanCode) {
if (!hasProductImage(sheetFolder, eanCode)) return null;
var url = productImageUrl(sheetFolder, eanCode);
if (IMG_CACHE[url]) return IMG_CACHE[url];
var img = new Image();
IMG_CACHE[url] = img;
img.decoding = 'async';
img.src = url;
if (img.decode) { img.decode().catch(function(){}); }
return img;
}
function updateBarcodeProductImage(sheetFolder, eanCode) {
if (!barcodeProductImgWrap || !barcodeProductImg) return;
if (!hasProductImage(sheetFolder, eanCode)) {
barcodeProductImgWrap.style.display = 'none';
barcodeProductImg.removeAttribute('src');
return;
}
var url = productImageUrl(sheetFolder, eanCode);
var cached = prefetchProductImage(sheetFolder, eanCode);
if (cached && cached.complete && cached.naturalWidth > 0) {
barcodeProductImg.onload = null;
barcodeProductImg.onerror = null;
barcodeProductImg.src = url;
barcodeProductImgWrap.style.display = 'flex';
return;
}
barcodeProductImgWrap.style.display = 'none';
barcodeProductImg.onload = function() { barcodeProductImgWrap.style.display = 'flex'; };
barcodeProductImg.onerror = function() {
barcodeProductImgWrap.style.display = 'none';
barcodeProductImg.removeAttribute('src');
};
barcodeProductImg.src = url;
}
function openBarcodeModal(eanCode, sheetFolder, eciValue) {
if (!eanCode || !barcodeModal) return;
if (barcodeModalTitle) {
barcodeModalTitle.textContent = eciValue || '';
}
renderBarcodeSVG(eanCode);
updateBarcodeProductImage(sheetFolder, eanCode);
setModalOpen(barcodeModal, true);
}
function svgToCanvas(svgEl, targetWidth, cb) {
if (!svgEl) { cb(null); return; }
var vb = (svgEl.getAttribute('viewBox') || '0 0 300 150').split(' ');
var w = parseFloat(vb[2]) || 300;
var h = parseFloat(vb[3]) || 150;
var scale = targetWidth / w;
var clone = svgEl.cloneNode(true);
clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
clone.setAttribute('width', w);
clone.setAttribute('height', h);
var svgStr = new XMLSerializer().serializeToString(clone);
var svgBlob = new Blob([svgStr], { type: 'image/svg+xml;charset=utf-8' });
var url = URL.createObjectURL(svgBlob);
var img = new Image();
img.onload = function() {
var canvas = document.createElement('canvas');
canvas.width = w * scale;
canvas.height = h * scale;
var ctx = canvas.getContext('2d');
ctx.fillStyle = '#ffffff';
ctx.fillRect(0, 0, canvas.width, canvas.height);
ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
URL.revokeObjectURL(url);
cb(canvas);
};
img.onerror = function() { URL.revokeObjectURL(url); cb(null); };
img.src = url;
}
function buildShareCanvas(cb) {
var title = (barcodeModalTitle && barcodeModalTitle.textContent) || '';
var hasImg = !!(barcodeProductImgWrap && barcodeProductImgWrap.style.display !== 'none' &&
barcodeProductImg && barcodeProductImg.complete && barcodeProductImg.naturalWidth > 0);
var cardWidth = (function() {
var card = barcodeModal && barcodeModal.querySelector('.barcode-modal-card');
return (card && card.clientWidth) || 340;
})();
var exportScale = 3;
var padding = Math.round(22 * exportScale);
var contentWidth = Math.round((cardWidth - 44) * exportScale);
svgToCanvas(document.getElementById('barcodeSvg'), contentWidth, function(barcodeCanvas) {
if (!barcodeCanvas) { cb(null); return; }
var titleFontSize = Math.round(20 * exportScale);
var titleHeight = title ? Math.round(titleFontSize * 1.5) : 0;
var imgMaxHeight = hasImg ? Math.round(220 * exportScale) : 0;
var gap = hasImg ? Math.round(14 * exportScale) : 0;
var canvas = document.createElement('canvas');
canvas.width = contentWidth + padding * 2;
canvas.height = padding * 2 + titleHeight + imgMaxHeight + gap + barcodeCanvas.height;
var ctx = canvas.getContext('2d');
ctx.fillStyle = '#ffffff';
ctx.fillRect(0, 0, canvas.width, canvas.height);
var y = padding;
if (title) {
ctx.fillStyle = '#000000';
ctx.font = '700 ' + titleFontSize + 'px -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif';
ctx.textAlign = 'center';
ctx.textBaseline = 'top';
ctx.fillText(title, canvas.width / 2, y);
y += titleHeight;
}
function drawBarcodeAndFinish() {
var bx = (canvas.width - barcodeCanvas.width) / 2;
ctx.drawImage(barcodeCanvas, bx, y);
cb(canvas);
}
if (hasImg) {
var pImg = new Image();
pImg.onload = function() {
var maxW = contentWidth;
var maxH = imgMaxHeight;
var ratio = Math.min(maxW / pImg.naturalWidth, maxH / pImg.naturalHeight, 1);
var dw = pImg.naturalWidth * ratio;
var dh = pImg.naturalHeight * ratio;
ctx.drawImage(pImg, (canvas.width - dw) / 2, y + (imgMaxHeight - dh) / 2, dw, dh);
y += imgMaxHeight + gap;
drawBarcodeAndFinish();
};
pImg.onerror = drawBarcodeAndFinish;
pImg.src = barcodeProductImg.src;
} else {
drawBarcodeAndFinish();
}
});
}
function shareBarcodeModal() {
if (!barcodeShareBtn || barcodeShareBtn.disabled) return;
barcodeShareBtn.disabled = true;
buildShareCanvas(function(canvas) {
if (!canvas) { barcodeShareBtn.disabled = false; return; }
canvas.toBlob(function(blob) {
barcodeShareBtn.disabled = false;
if (!blob) return;
var eciText = (barcodeModalTitle && barcodeModalTitle.textContent) || 'referencia';
var fileName = (eciText.replace(/[^a-z0-9]+/gi, '-').toLowerCase() || 'referencia') + '.png';
var file;
try { file = new File([blob], fileName, { type: 'image/png' }); } catch (err) { file = null; }
if (file && navigator.canShare && navigator.canShare({ files: [file] })) {
navigator.share({ files: [file] }).catch(function(){});
return;
}
var url = URL.createObjectURL(blob);
var a = document.createElement('a');
a.href = url;
a.download = fileName;
document.body.appendChild(a);
a.click();
document.body.removeChild(a);
setTimeout(function(){ URL.revokeObjectURL(url); }, 4000);
}, 'image/png');
});
}
if (barcodeShareBtn) barcodeShareBtn.addEventListener('click', shareBarcodeModal);
function closeBarcodeModal() {
setModalOpen(barcodeModal, false);
}
if (barcodeCloseBtn) barcodeCloseBtn.addEventListener('click', closeBarcodeModal);
if (barcodeModal) {
barcodeModal.addEventListener('click', function(e) {
if (e.target === barcodeModal) closeBarcodeModal();
});
}
document.addEventListener('keydown', function(e) {
if (e.key === 'Escape' && barcodeModal && barcodeModal.classList.contains('open')) {
closeBarcodeModal();
}
});

// ---- 51-row-clicks.js ----
function rowContext(tr){
var panelEl = tr.closest('.panel');
if(panelEl && panelEl.id === 'search-results'){
return { ean: tr.dataset.ean || '', folder: tr.dataset.sheetFolder || null, eci: tr.dataset.eci || '' };
}
var cells = tr.querySelectorAll('td');
var eciCell = cells[2];
var eciSpan = eciCell ? (eciCell.querySelector('.cell-text') || eciCell) : null;
return {
ean: tr.dataset.ean || '',
folder: panelEl ? panelEl.dataset.sheetFolder : null,
eci: eciSpan ? eciSpan.textContent.trim() : ''
};
}
function prefetchFromEvent(e){
var tr = e.target.closest ? e.target.closest('tr[data-ean]') : null;
if(!tr || tr.dataset.prefetched) return;
tr.dataset.prefetched = '1';
var ctx = rowContext(tr);
prefetchProductImage(ctx.folder, ctx.ean);
}
mainEl.addEventListener('mouseover', prefetchFromEvent);
mainEl.addEventListener('touchstart', prefetchFromEvent, { passive: true });
mainEl.addEventListener('click', function(e){
var hit = e.target.closest('.ean-clickable') ||
(function(){
var inner = e.target.closest('.result-cell-inner');
return inner && inner.querySelector('.result-cell-value.ean-clickable') ? inner : null;
})();
if(!hit) return;
var tr = hit.closest('tr');
if(!tr) return;
var ctx = rowContext(tr);
if(!ctx.ean || !EAN_TEXT_RE.test(ctx.ean)) return;
e.stopPropagation();
openBarcodeModal(ctx.ean, ctx.folder, ctx.eci);
});

// ---- 60-scan-refs.js ----
var REF_TYPES = {
EAN: { label: 'EAN', hint: '12 dígitos' },
PN:  { label: 'P/N', hint: 'p.ej. MHFA4Y/A' }
};
var REF_TYPE_ORDER = ['EAN', 'PN'];
var PN_RE = /^([A-Z0-9]{6,7})\/([A-Z])$/;
var PN_BARE_RE = /^[A-Z0-9]{7,8}$/;
function pnFold(s){
return s.replace(/O/g, '0').replace(/I/g, '1').replace(/S/g, '5').replace(/B/g, '8');
}
function canonEan(d){
if(d.length === 13 && d.charAt(0) === '0') return d.slice(1);
return d;
}
function eanCheckDigit(first12){
var sum = 0;
for(var i = 0; i < 12; i++) sum += (+first12.charAt(i)) * (i % 2 === 0 ? 1 : 3);
return (10 - (sum % 10)) % 10;
}
function eanChecksumOk(d){
if(!/^\d{12,13}$/.test(d)) return false;
var d13 = d.length === 12 ? '0' + d : d;
return eanCheckDigit(d13) === +d13.charAt(12);
}
function classifyRef(raw){
var s = String(raw == null ? '' : raw).trim().toUpperCase();
if(!s) return null;
if(/^[\d.\s\-·,]+$/.test(s)){
var d = s.replace(/\D/g, '');
if(d.length === 12 || d.length === 13) return { type: 'EAN', key: canonEan(d), valid: eanChecksumOk(d) };
return null;
}
var compact = s.replace(/\s+/g, '');
var m = PN_RE.exec(compact);
if(m) return { type: 'PN', key: pnFold(m[1] + m[2]), valid: true };
if(PN_BARE_RE.test(compact) && /[A-Z]/.test(compact) && /\d/.test(compact) && /[A-Z]$/.test(compact)){
return { type: 'PN', key: pnFold(compact), valid: true, bare: true };
}
return null;
}
var REF_INDEX = null;
function buildRefIndex(rows){
var idx = { EAN: new Map(), PN: new Map() };
rows.forEach(function(r){
var seenInRow = {};
r.cells.forEach(function(c){
var cl = classifyRef(c.v);
if(!cl) return;
var k = cl.type + ':' + cl.key;
if(seenInRow[k]) return;
seenInRow[k] = true;
var list = idx[cl.type].get(cl.key);
if(!list){ list = []; idx[cl.type].set(cl.key, list); }
list.push({ row: r, type: cl.type, ref: c.v.trim() });
});
});
return idx;
}
function getRefIndex(){
if(!REF_INDEX) REF_INDEX = buildRefIndex(ALL_ROWS);
return REF_INDEX;
}
function substitutionVariants(d){
var out = [];
for(var i = 0; i < d.length; i++){
var orig = d.charAt(i);
for(var n = 0; n <= 9; n++){
var ch = String(n);
if(ch !== orig) out.push(d.slice(0, i) + ch + d.slice(i + 1));
}
}
return out;
}
function deletionVariants(d){            // la lectura trae un digito DE MAS
var out = [];
for(var i = 0; i < d.length; i++) out.push(d.slice(0, i) + d.slice(i + 1));
return out;
}
function transpositionVariants(d){
var out = [];
for(var i = 0; i + 1 < d.length; i++){
if(d.charAt(i) !== d.charAt(i + 1)) out.push(d.slice(0, i) + d.charAt(i + 1) + d.charAt(i) + d.slice(i + 2));
}
return out;
}
function lookupVariants(map, variants, filter){
var hits = [], seen = {};
variants.forEach(function(v){
if(seen[v] || (filter && !filter(v))) return;
seen[v] = true;
var list = map.get(v);
if(list) list.forEach(function(e){ hits.push({ entry: e, key: v }); });
});
return hits;
}
function rowKey(row){ return row.slug + '|' + row.section + '|' + row.search; }
function uniqueRows(entries){
var seen = {}, out = [];
entries.forEach(function(e){
var k = rowKey(e.row);
if(seen[k]) return;
seen[k] = true;
out.push(e);
});
return out;
}
function toMatches(entries, read, dist){
return uniqueRows(entries).map(function(e){ return { row: e.row, type: e.type, ref: e.ref, read: read, dist: dist }; });
}
function settle(base, hits, dist){
var rows = hits.map(function(h){ return h.entry; });
if(!rows.length) return false;
var keys = {};
hits.forEach(function(h){ keys[h.entry.type + ':' + h.key] = h.key; });
var distinct = Object.keys(keys);
base.status = distinct.length === 1 ? 'corrected' : 'ambiguous';
if(base.status === 'corrected'){ base.type = hits[0].entry.type; base.key = keys[distinct[0]]; }
base.matches = toMatches(rows, base.read, dist);
return true;
}
function resolveRef(cl, index){
index = index || getRefIndex();
var map = index[cl.type];
var base = { type: cl.type, readType: cl.type, key: cl.key, read: cl.key, matches: [], level: cl.level };
var exact = map.get(cl.key);
if(exact){
base.status = 'exact';
base.matches = toMatches(exact, cl.key, 0);
return base;
}
if(cl.exactOnly){ base.status = 'unknown'; return base; }
if(cl.type === 'EAN'){
if(cl.valid){ base.status = 'unknown'; return base; }
var d = cl.key;
var v = substitutionVariants(d).concat(transpositionVariants(d));
var hits = lookupVariants(map, v, function(x){ return eanChecksumOk(x); });
if(!hits.length && d.length === 13) hits = lookupVariants(map, deletionVariants(d), function(x){ return eanChecksumOk(x); });
if(hits.length && settle(base, hits, 1)) return base;
base.status = 'invalid';
return base;
}
base.status = 'unknown';
return base;
}
function formatRefKey(type, key){
if(type === 'PN' && key.length >= 7) return key.slice(0, -1) + '/' + key.slice(-1);
return key;
}
function refReadLabel(r){
var rt = r.readType || r.type;
var s = REF_TYPES[rt].label + ' ' + formatRefKey(rt, r.read || r.key);
if(r.status === 'corrected' && r.key !== r.read) s += ' → ' + REF_TYPES[r.type].label + ' ' + formatRefKey(r.type, r.key);
return s;
}

// ---- 61-scan-ocr.js ----
function ocrDigitFix(s){
return s
.replace(/[OoQD]/g, '0')
.replace(/[IilL|!\]\[]/g, '1')
.replace(/[Zz]/g, '2')
.replace(/[Ss]/g, '5')
.replace(/[Gg]/g, '6')
.replace(/[Tt]/g, '7')
.replace(/[Bb]/g, '8')
.replace(/[gq]/g, '9');
}
function isGarbageRead(v){ return /(.)\1{4,}/.test(v); }
var RE_DIGIT_RUN = /\d[\d .,\-·]*\d/g;
var RE_PN_SLASH = /([A-Z0-9]{6,7})\s*\/\s*([A-Z])(?![A-Z0-9])/g;
var RE_PN_BARE = /[A-Z0-9]{7,8}/g;
var LEVEL_PATTERN = 1, LEVEL_RUN = 2, LEVEL_WINDOW = 3;
function digitRunCandidates(d, add, exactOnly){
var n = d.length, i;
if(exactOnly){
if(n === 12 || n === 13) add(d, LEVEL_WINDOW, { exactOnly: true });
for(i = 0; i + 12 <= n; i++) if(eanChecksumOk(d.substr(i, 12))) add(d.substr(i, 12), LEVEL_WINDOW, { exactOnly: true });
for(i = 0; i + 13 <= n; i++) if(eanChecksumOk(d.substr(i, 13))) add(d.substr(i, 13), LEVEL_WINDOW, { exactOnly: true });
return;
}
if(n === 12 || n === 13) add(d, LEVEL_RUN);
else if(n > 13) digitRunCandidates(d, add, true);
}
function extractCandidates(rawText){
var out = [];
var seen = {};
function add(raw, level, opts){
var cl = classifyRef(raw);
if(!cl) return;
if(opts && opts.exactOnly) cl.exactOnly = true;
if(isGarbageRead(cl.key)) return;
var k = cl.type + ':' + cl.key;
if(seen[k]){ if(seen[k].level > level) seen[k].level = level; return; }
cl.level = level;
seen[k] = cl;
out.push(cl);
}
var upper = String(rawText || '').toUpperCase();
var lines = upper.split(/\n+/);
lines.forEach(function(line){
if(!line.trim()) return;
[line, ocrDigitFix(line)].forEach(function(text){
var m;
RE_DIGIT_RUN.lastIndex = 0;
while((m = RE_DIGIT_RUN.exec(text)) !== null) digitRunCandidates(m[0].replace(/\D/g, ''), add, false);
});
var p;
RE_PN_SLASH.lastIndex = 0;
while((p = RE_PN_SLASH.exec(line)) !== null) add(p[1] + '/' + p[2], LEVEL_PATTERN);
RE_PN_BARE.lastIndex = 0;
while((p = RE_PN_BARE.exec(line)) !== null) add(p[0], LEVEL_RUN);
});
if(lines.length > 1){
var joined = lines.join(' ');
[joined, ocrDigitFix(joined)].forEach(function(text){
var m;
RE_DIGIT_RUN.lastIndex = 0;
while((m = RE_DIGIT_RUN.exec(text)) !== null){
var d = m[0].replace(/\D/g, '');
if(d.length >= 12) digitRunCandidates(d, add, true);
}
});
}
return out;
}
function evaluateOcrText(rawText, index){
var cands = extractCandidates(rawText);
var result = { status: 'none', matches: [], reads: [], candidates: cands };
if(!cands.length) return result;
var resolved = [];
[LEVEL_PATTERN, LEVEL_RUN, LEVEL_WINDOW].some(function(level){
var batch = cands.filter(function(c){ return c.level === level; })
.map(function(c){ var r = resolveRef(c, index); r.level = level; return r; });
resolved = resolved.concat(batch);
return batch.some(function(r){ return r.status === 'exact'; });
});
function sortByType(a, b){ return REF_TYPE_ORDER.indexOf(a.type) - REF_TYPE_ORDER.indexOf(b.type); }
var exact = resolved.filter(function(r){ return r.status === 'exact'; }).sort(sortByType);
if(exact.length){
result.status = 'exact';
result.matches = mergeMatches(exact);
result.reads = exact.map(refReadLabel);
return result;
}
var corrected = resolved.filter(function(r){ return r.status === 'corrected'; }).sort(sortByType);
var ambiguous = resolved.filter(function(r){ return r.status === 'ambiguous'; }).sort(sortByType);
if(corrected.length || ambiguous.length){
var all = mergeMatches(corrected.concat(ambiguous));
result.status = (all.length === 1 && !ambiguous.length) ? 'corrected' : 'ambiguous';
result.matches = all;
result.reads = corrected.concat(ambiguous).map(refReadLabel);
return result;
}
var wellFormed = resolved.filter(function(r){ return r.status === 'unknown' && r.level < LEVEL_WINDOW; }).sort(sortByType);
if(wellFormed.length){
result.status = 'unknown';
result.reads = wellFormed.slice(0, 3).map(refReadLabel);
}
return result;
}
function mergeMatches(results){
var seen = {}, out = [];
results.forEach(function(r){
r.matches.forEach(function(m){
var k = rowKey(m.row);
if(seen[k]) return;
seen[k] = true;
out.push(m);
});
});
return out;
}

// ---- 62-scan-camera.js ----
function detectNativeZoom() {
liveNativeZoom = null;
if (!liveTrack || !liveTrack.getCapabilities) return;
try {
var caps = liveTrack.getCapabilities();
if (caps && caps.zoom && typeof caps.zoom.max === 'number' && caps.zoom.max > caps.zoom.min) {
liveNativeZoom = caps.zoom;
}
} catch (e) { liveNativeZoom = null; }
}
function detectCaptureCapabilities() {
liveImageCapture = null;
if (!liveTrack) return;
try {
if (typeof window.ImageCapture === 'function') {
liveImageCapture = new window.ImageCapture(liveTrack);
}
} catch (e) { liveImageCapture = null; }
}
function maximiseTrackResolution() {
if (!liveTrack || !liveTrack.getCapabilities || !liveTrack.applyConstraints) return;
var caps;
try { caps = liveTrack.getCapabilities(); } catch (e) { return; }
if (!caps || !caps.width || typeof caps.width.max !== 'number') return;
var current = 0;
try { current = (liveTrack.getSettings() || {}).width || 0; } catch (e) {}
if (caps.width.max <= current * 1.2) return;
try {
liveTrack.applyConstraints({ width: { ideal: caps.width.max } })
.then(function() { detectNativeZoom(); })
.catch(function() {});
} catch (e) {}
}
function grabStillSource() {
if (liveImageCapture && liveImageCapture.grabFrame) {
return liveImageCapture.grabFrame().catch(function() { return scannerVideo; });
}
return Promise.resolve(scannerVideo);
}
function applyZoom(z) {
liveZoom = z;
scannerZoomValue.textContent = z.toFixed(1) + 'x';
if (liveNativeZoom && liveTrack) {
var span = liveNativeZoom.max - liveNativeZoom.min;
var frac = Math.min(1, (z - 1) / 3);
var target = liveNativeZoom.min + span * frac;
try { liveTrack.applyConstraints({ advanced: [{ zoom: target }] }); } catch (e) {}
}
var wPct = LIVE_CROP_W_PCT / z;
var hPct = LIVE_CROP_H_PCT / z;
scannerLiveFrame.style.width = wPct + '%';
scannerLiveFrame.style.height = hPct + '%';
scannerLiveFrame.style.left = ((100 - wPct) / 2) + '%';
scannerLiveFrame.style.top = ((100 - hPct) / 2) + '%';
}
function stopLiveScan(keepRestart) {
liveActive = false;
liveBusy = false;
stopBarcodeFastPath();
if (liveTimer) { clearInterval(liveTimer); liveTimer = null; }
if (ocrStartTimer) { clearTimeout(ocrStartTimer); ocrStartTimer = null; }
if (liveStream) {
liveStream.getTracks().forEach(function(t) { t.stop(); });
liveStream = null;
}
liveTrack = null;
liveImageCapture = null;
if (scannerVideo) scannerVideo.srcObject = null;
if (scannerLiveWrap) { scannerLiveWrap.style.display = 'none'; scannerLiveWrap.classList.remove('camera-ready'); }
if (scannerLiveStopBtn) scannerLiveStopBtn.style.display = keepRestart ? 'block' : 'none';
if (scannerZoomRow) scannerZoomRow.classList.remove('visible');
if (scannerLiveCrop) scannerLiveCrop.textContent = '';
}
function grabLiveCrop(src, sw, sh) {
if (!sw || !sh) return null;
var cropW = Math.round(sw * (LIVE_CROP_W_PCT / 100) / liveZoom);
var cropH = Math.round(sh * (LIVE_CROP_H_PCT / 100) / liveZoom);
var cropX = Math.round((sw - cropW) / 2);
var cropY = Math.round((sh - cropH) / 2);
if (cropW < 16 || cropH < 8) return null;
var TARGET_W = 1400;
var MAX_SCALE = 3;
var scale = Math.min(MAX_SCALE, Math.max(1, TARGET_W / cropW));
var outW = Math.round(cropW * scale);
var outH = Math.round(cropH * scale);
var canvas = document.createElement('canvas');
canvas.width = outW;
canvas.height = outH;
var ctx = canvas.getContext('2d');
ctx.imageSmoothingEnabled = true;
if ('imageSmoothingQuality' in ctx) ctx.imageSmoothingQuality = 'high';
ctx.drawImage(src, cropX, cropY, cropW, cropH, 0, 0, outW, outH);
preprocessCrop(ctx, outW, outH);
scannerLiveCrop.textContent = outW + 'x' + outH;
return canvas;
}
function preprocessCrop(ctx, outW, outH) {
try {
var img = ctx.getImageData(0, 0, outW, outH);
var d = img.data;
var i;
var lo = 255, hi = 0;
for (i = 0; i < d.length; i += 4) {
var g = (d[i] * 0.299 + d[i + 1] * 0.587 + d[i + 2] * 0.114) | 0;
d[i] = d[i + 1] = d[i + 2] = g;
if (g < lo) lo = g;
if (g > hi) hi = g;
}
var range = hi - lo;
if (range > 20) {
var mul = 255 / range;
for (i = 0; i < d.length; i += 4) {
var v = (d[i] - lo) * mul;
v = v < 0 ? 0 : (v > 255 ? 255 : v);
d[i] = d[i + 1] = d[i + 2] = v;
}
}
ctx.putImageData(img, 0, 0);
} catch (e) { /* tainted canvas should not happen with a local stream */ }
}

// ---- 63-scan-live.js ----
var scanBtn = document.getElementById('scanToggle');
var scannerModal = document.getElementById('scannerModal');
var scannerCloseBtn = document.getElementById('scannerCloseBtn');
var scannerStatus = document.getElementById('scannerStatus');
var scannerResults = document.getElementById('scannerResults');
var scannerDetected = document.getElementById('scannerDetected');
var scannerDetectedText = document.getElementById('scannerDetectedText');
var scannerLiveWrap = document.getElementById('scannerLiveWrap');
var scannerVideo = document.getElementById('scannerVideo');
var scannerLiveOverlay = document.getElementById('scannerLiveOverlay');
var scannerLiveStopBtn = document.getElementById('scannerLiveStopBtn');
var scannerLiveFrame = document.getElementById('scannerLiveFrame');
var scannerLiveCrop = document.getElementById('scannerLiveCrop');
var scannerBarcodeOverlay = document.getElementById('scannerBarcodeOverlay');
var scannerZoomRow = document.getElementById('scannerZoomRow');
var scannerZoomSlider = document.getElementById('scannerZoomSlider');
var scannerZoomValue = document.getElementById('scannerZoomValue');
var scannerTypes = document.getElementById('scannerTypes');
function isMobileDevice() {
var ua = navigator.userAgent || '';
if (/Android|iPhone|iPad|iPod|Mobile/i.test(ua)) return true;
if (/Macintosh/i.test(ua) && navigator.maxTouchPoints > 1) return true;
return false;
}
function hasGetUserMedia() {
return !!(navigator.mediaDevices && navigator.mediaDevices.getUserMedia);
}
if (scanBtn && isMobileDevice() && hasGetUserMedia()) {
scanBtn.style.display = 'flex';
setTimeout(function() { ensureBarcodeDetector(); }, 2000);
}
var liveStream = null;
var liveTrack = null;
var liveTimer = null;
var liveBusy = false;
var liveActive = false;
var liveZoom = 1.0;              // zoom digital aplicado al recorte
var liveNativeZoom = null;       // {min,max,step} si la camara expone zoom nativo
var LIVE_INTERVAL_MS = 700;
var OCR_START_DELAY_MS = 1200;   // el OCR (P/N) entra solo si el codigo de barras no ha dado nada
var ocrStartTimer = null;
var liveImageCapture = null;     // ImageCapture: existe en Chrome/Android, no en Safari/iOS
var liveCaptureCanvas = null;    // canvas reutilizado para el detect del barcode (downscale)
var LIVE_CROP_W_PCT = 100;
var LIVE_CROP_H_PCT = 65;
var liveProfileToggle = false;
var liveTextWindow = [];         // ultimas lecturas OCR (los dos alfabetos se evaluan juntos)
var liveVotes = {};              // confirmacion de correcciones entre fotogramas
var liveTypeLock = null;         // {type, until}: tipo detectado que guia el alfabeto OCR
var TYPE_LOCK_MS = 2500;
var VOTES_TO_CONFIRM = 2;
var TESSERACT_URL = 'https://cdn.jsdelivr.net/npm/tesseract.js@4/dist/tesseract.min.js';
var ocrWorker = null;
var ocrWorkerPromise = null;
var ocrWorkerMode = null;
function getOcrWorker() {
if (ocrWorkerPromise) return ocrWorkerPromise;
ocrWorkerPromise = loadScriptOnce(TESSERACT_URL).then(function() {
if (typeof Tesseract === 'undefined' || !Tesseract.createWorker) throw new Error('Tesseract no disponible');
return Tesseract.createWorker();
}).then(function(w) {
return w.loadLanguage('eng').then(function() { return w.initialize('eng'); })
.then(function() { ocrWorker = w; return w; });
}).catch(function(err) {
ocrWorkerPromise = null;
throw err;
});
return ocrWorkerPromise;
}
var OCR_PROFILES = {
digits: '0123456789. -',
pn: '0123456789/ ACDEFGHJKLMNPQRTUVWXYZ'
};
function setOcrProfile(worker, profile, pageseg) {
var key = profile + '|' + pageseg;
if (ocrWorkerMode === key) return Promise.resolve();
ocrWorkerMode = key;
return worker.setParameters({
tessedit_char_whitelist: OCR_PROFILES[profile],
tessedit_pageseg_mode: pageseg
}).catch(function() {});
}
function nextOcrProfile() {
if (liveTypeLock && Date.now() < liveTypeLock.until) {
return liveTypeLock.type === 'PN' ? 'pn' : 'digits';
}
liveTypeLock = null;
liveProfileToggle = !liveProfileToggle;
return liveProfileToggle ? 'digits' : 'pn';
}
function setTypeIndicator(types, state) {
if (!scannerTypes) return;
scannerTypes.querySelectorAll('[data-type]').forEach(function(el) {
var t = el.dataset.type;
el.classList.toggle('detected', !!(types && types[t]));
el.classList.toggle('found', state === 'found' && !!(types && types[t]));
el.classList.toggle('locked', !!(liveTypeLock && liveTypeLock.type === t && state !== 'found'));
});
}
function openScanner() {
if (!scannerModal) return;
setModalOpen(scannerModal, true);
resetScannerUI();
if (hasGetUserMedia()) {
startLiveScan();
} else {
setScanStatus('Este navegador no permite acceso a la camara. Usa el buscador de la pagina.', 'error');
}
}
function closeScanner() {
if (!scannerModal) return;
setModalOpen(scannerModal, false);
stopLiveScan();
}
function setScanStatus(text, cls) {
scannerStatus.textContent = text;
scannerStatus.className = 'scanner-status' + (cls ? ' ' + cls : '');
}
function showDetected(text) {
scannerDetectedText.textContent = text;
scannerDetected.style.display = text ? 'block' : 'none';
}
function resetScannerUI() {
scannerResults.innerHTML = '';
setScanStatus('');
showDetected('');
setTypeIndicator(null);
}
function restartLiveScan() {
resetScannerUI();
stopLiveScan();
startLiveScan();
}
function startOcrLoop() {
ocrStartTimer = null;
if (!liveActive || liveTimer) return;
maximiseTrackResolution();
scannerLiveOverlay.textContent = 'Encuadra el EAN o el P/N (preparando lectura de texto)...';
getOcrWorker().then(function() {
if (!liveActive || liveTimer) return;
scannerLiveOverlay.textContent = 'Encuadra el EAN o el P/N...';
liveTimer = setInterval(liveScanTick, LIVE_INTERVAL_MS);
}).catch(function() {
if (!liveActive) return;
scannerLiveOverlay.textContent = 'Solo codigo de barras: no se pudo cargar el OCR.';
});
}
function startLiveScan() {
if (liveActive) return;
if (!hasGetUserMedia()) {
scannerLiveOverlay.textContent = 'Camara no disponible en este navegador.';
return;
}
liveActive = true;
liveTextWindow = [];
liveVotes = {};
liveTypeLock = null;
scannerLiveWrap.style.display = 'block';
scannerLiveStopBtn.style.display = 'block';
scannerZoomRow.classList.add('visible');
scannerLiveOverlay.textContent = 'Iniciando camara...';
getRefIndex();   // el indice se construye una vez; mejor ahora que en el primer acierto
var constraints = {
video: {
facingMode: { ideal: 'environment' },
width: { ideal: 1920 },
height: { ideal: 1080 },
focusMode: { ideal: 'continuous' }
}
};
navigator.mediaDevices.getUserMedia(constraints)
.catch(function() {
return navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' } });
})
.then(function(stream) {
if (!liveActive) { stream.getTracks().forEach(function(t){ t.stop(); }); return null; }
liveStream = stream;
liveTrack = stream.getVideoTracks()[0] || null;
detectNativeZoom();
detectCaptureCapabilities();
scannerVideo.srcObject = stream;
return scannerVideo.play();
}).then(function() {
if (!liveActive) return;
startBarcodeFastPath();
scannerLiveWrap.classList.add('camera-ready');
applyZoom(liveZoom);
scannerLiveOverlay.textContent = 'Apunta al codigo de barras EAN...';
ocrStartTimer = setTimeout(startOcrLoop, OCR_START_DELAY_MS);
}).catch(function() {
scannerLiveOverlay.textContent = 'No se pudo iniciar la camara. Revisa los permisos.';
stopLiveScan();
});
}
var ZXING_URL = 'vendor/zxing-browser.min.js';
var ZXING_TICK_MS = 50;
var barcodeEngine = null;        // 'native' | 'zxing' | null (sin lector)
var barcodeDetector = null;      // BarcodeDetector nativo
var zxingReader = null;
var barcodeRafId = null;
var barcodeTimerId = null;
var barcodeBoxesShown = false;
var barcodeReadyPromise = null;
var BARCODE_DETECT_W = 1280;     // las barras se decodifican de sobra a este ancho
function ensureBarcodeDetector() {
if (barcodeReadyPromise) return barcodeReadyPromise;
if (typeof window.BarcodeDetector !== 'undefined') {
barcodeReadyPromise = Promise.resolve('native');
return barcodeReadyPromise;
}
barcodeReadyPromise = loadScriptOnce(ZXING_URL).then(function() {
var ok = window.ZXingBrowser && window.ZXingBrowser.BrowserMultiFormatOneDReader;
if (!ok) throw new Error('ZXing no disponible');
return 'zxing';
}).catch(function() {
barcodeReadyPromise = null;   // permite reintentar en el siguiente escaneo
return null;
});
return barcodeReadyPromise;
}
function initBarcodeDetector() {
return ensureBarcodeDetector().then(function(engine) {
if (!engine) return false;
try {
if (engine === 'native') {
barcodeDetector = new window.BarcodeDetector({ formats: ['ean_13', 'upc_a', 'code_128', 'code_39'] });
} else if (!zxingReader) {
var BF = window.ZXingBrowser.BarcodeFormat;
var hints = new Map();
if (BF) hints.set(2, [BF.EAN_13, BF.UPC_A, BF.CODE_128, BF.CODE_39]);
zxingReader = new window.ZXingBrowser.BrowserMultiFormatOneDReader(hints);
}
barcodeEngine = engine;
return true;
} catch (e) { return false; }
});
}
function drawBarcodeBoxes(boxes) {
if (!scannerBarcodeOverlay) return;
var vw = scannerVideo.videoWidth, vh = scannerVideo.videoHeight;
if (!vw || !vh) return;
if (scannerBarcodeOverlay.width !== vw) scannerBarcodeOverlay.width = vw;
if (scannerBarcodeOverlay.height !== vh) scannerBarcodeOverlay.height = vh;
var ctx = scannerBarcodeOverlay.getContext('2d');
ctx.clearRect(0, 0, vw, vh);
ctx.lineWidth = Math.max(2, vw * 0.004);
ctx.strokeStyle = 'rgba(52,199,89,.95)';
boxes.forEach(function(b) { if (b) ctx.strokeRect(b.x, b.y, b.width, b.height); });
barcodeBoxesShown = boxes.length > 0;
}
function barcodeDetectSource(needCanvas) {
var vw = scannerVideo.videoWidth, vh = scannerVideo.videoHeight;
if (!vw) return null;
if (vw <= BARCODE_DETECT_W && !needCanvas) return { src: scannerVideo, k: 1 };
var dw = Math.min(vw, BARCODE_DETECT_W);
var dh = Math.max(1, Math.round(vh * dw / vw));
if (!liveCaptureCanvas) liveCaptureCanvas = document.createElement('canvas');
if (liveCaptureCanvas.width !== dw || liveCaptureCanvas.height !== dh) {
liveCaptureCanvas.width = dw;
liveCaptureCanvas.height = dh;
}
try {
liveCaptureCanvas.getContext('2d').drawImage(scannerVideo, 0, 0, dw, dh);
} catch (e) { return needCanvas ? null : { src: scannerVideo, k: 1 }; }
return { src: liveCaptureCanvas, k: vw / dw };
}
function barcodeValueCandidates(raw) {
var out = [], seen = {};
function add(v) {
var cl = v && classifyRef(v);
if (!cl) return;
var k = cl.type + ':' + cl.key;
if (seen[k]) return;
seen[k] = true;
out.push(cl);
}
var up = String(raw || '').toUpperCase().trim();
add(up);
add(up.replace(/\D/g, ''));
var m = /([A-Z0-9]{6,})\s*\/\s*([A-Z])(?![A-Z0-9])/.exec(up);
if (m) {
add(m[1].slice(-7) + '/' + m[2]);
add(m[1].slice(-6) + '/' + m[2]);
}
var runs = up.match(/\d{12,13}/g) || [];
runs.forEach(function(r) { if (eanChecksumOk(r)) add(r); });
return out;
}
function handleBarcodeCodes(codes) {
var accepted = false;
var unknown = null, unrecognized = '';
for (var i = 0; i < codes.length; i++) {
var raw = String(codes[i].rawValue || '').trim();
var found = false;
barcodeValueCandidates(raw).forEach(function(cl) {
if (accepted) return;
if (cl.type === 'EAN' && !cl.valid) return;
found = true;
var res = resolveRef(cl);
if (res.status === 'exact') {
acceptResult(res.matches, [REF_TYPES[cl.type].label + ' ' + formatRefKey(cl.type, cl.key)], typesOf([cl]), 'barcode');
accepted = true;
} else {
unknown = cl;
}
});
if (accepted) return true;
if (!found && raw) unrecognized = raw;
}
if (!unknown && unrecognized) {
var umsg = 'Codigo leido: ' + unrecognized.slice(0, 24) + ' — no es un EAN ni un P/N conocido';
if (scannerStatus.textContent !== umsg) setScanStatus(umsg);
}
if (unknown) {
var msg = REF_TYPES[unknown.type].label + ' ' + formatRefKey(unknown.type, unknown.key) + ' leido — no esta en el catalogo';
if (scannerStatus.textContent !== msg) setScanStatus(msg);
setTypeIndicator(typesOf([unknown]));
}
return false;
}
function barcodeFastTick() {
if (!liveActive || !barcodeEngine || !scannerVideo.videoWidth) return;
var frame = barcodeDetectSource(false);
if (!frame) { scheduleBarcodeTick(); return; }
barcodeDetector.detect(frame.src).then(function(codes) {
if (!liveActive) return;
var k = frame.k;
var boxes = codes.map(function(c) {
var b = c.boundingBox;
if (!b || k === 1) return b;
return { x: b.x * k, y: b.y * k, width: b.width * k, height: b.height * k };
});
if (boxes.length || barcodeBoxesShown) drawBarcodeBoxes(boxes);
if (handleBarcodeCodes(codes)) return;
scheduleBarcodeTick();
}).catch(function() {
scheduleBarcodeTick();
});
}
function zxingBox(result, k) {
var pts = (result.getResultPoints && result.getResultPoints()) || [];
if (!pts.length) return null;
var x0 = Infinity, x1 = -Infinity, ySum = 0;
pts.forEach(function(p) {
var x = p.getX(), y = p.getY();
if (x < x0) x0 = x;
if (x > x1) x1 = x;
ySum += y;
});
var w = x1 - x0;
if (!(w > 0)) return null;
var h = w * 0.4;
var cy = ySum / pts.length;
return { x: x0 * k, y: (cy - h / 2) * k, width: w * k, height: h * k };
}
function zxingFastTick() {
barcodeTimerId = null;
if (!liveActive || !barcodeEngine) return;
var frame = barcodeDetectSource(true);
var result = null;
if (frame) {
try { result = zxingReader.decodeFromCanvas(frame.src); } catch (e) { result = null; }   // NotFoundException: frame sin codigo
}
if (result) {
var box = zxingBox(result, frame.k);
var boxes = box ? [box] : [];
drawBarcodeBoxes(boxes);
if (handleBarcodeCodes([{ rawValue: result.getText() }])) return;
} else if (barcodeBoxesShown) {
drawBarcodeBoxes([]);
}
scheduleBarcodeTick();
}
function scheduleBarcodeTick() {
if (!liveActive || !barcodeEngine) return;
if (barcodeEngine === 'zxing') barcodeTimerId = setTimeout(zxingFastTick, ZXING_TICK_MS);
else barcodeRafId = requestAnimationFrame(barcodeFastTick);
}
function startBarcodeFastPath() {
initBarcodeDetector().then(function(ok) {
if (!ok || !liveActive) return;
drawBarcodeBoxes([]);
scheduleBarcodeTick();
});
}
function stopBarcodeFastPath() {
if (barcodeRafId) { cancelAnimationFrame(barcodeRafId); barcodeRafId = null; }
if (barcodeTimerId) { clearTimeout(barcodeTimerId); barcodeTimerId = null; }
drawBarcodeBoxes([]);
}
function liveScanTick() {
if (liveBusy || !liveActive || !ocrWorker || !scannerVideo.videoWidth) return;
liveBusy = true;
grabStillSource().then(function(src) {
if (!liveActive) { liveBusy = false; return; }
var sw = src.videoWidth || src.width || 0;
var sh = src.videoHeight || src.height || 0;
var canvas = grabLiveCrop(src, sw, sh);
if (src && src.close) src.close();   // ImageBitmap: liberar el buffer
if (!canvas) { liveBusy = false; return; }
return runOcrOnCrop(canvas);
}).catch(function() {
liveBusy = false;
});
}
function typesOf(list) {
var t = {};
(list || []).forEach(function(c) { if (c && c.type) t[c.type] = true; });
return t;
}
function runOcrOnCrop(canvas) {
var profile = nextOcrProfile();
return setOcrProfile(ocrWorker, profile, '6').then(function() {
return ocrWorker.recognize(canvas);
}).then(function(result) {
liveBusy = false;
if (!liveActive) return;
var text = (result && result.data && result.data.text) || '';
liveTextWindow.push(text);
if (liveTextWindow.length > 3) liveTextWindow.shift();
handleOcrOutcome(evaluateOcrText(liveTextWindow.join('\n')));
}).catch(function() {
liveBusy = false;
});
}
function handleOcrOutcome(outcome) {
var detectedTypes = typesOf(outcome.candidates);
var strong = outcome.candidates.filter(function(c) { return c.level <= LEVEL_RUN; })
.sort(function(a, b) { return a.level - b.level || REF_TYPE_ORDER.indexOf(a.type) - REF_TYPE_ORDER.indexOf(b.type); });
if (strong.length && outcome.status !== 'exact') {
liveTypeLock = { type: strong[0].type, until: Date.now() + TYPE_LOCK_MS };
}
if (outcome.status === 'exact') {
acceptResult(outcome.matches, outcome.reads, detectedTypes, 'ocr');
return;
}
if (outcome.status === 'corrected') {
var key = 'C:' + rowKey(outcome.matches[0].row);
var votes = (liveVotes[key] = (liveVotes[key] || 0) + 1);
if (votes >= VOTES_TO_CONFIRM) {
acceptResult(outcome.matches, outcome.reads, detectedTypes, 'corrected');
return;
}
scannerLiveOverlay.textContent = 'Posible: ' + outcome.reads[0] + ' (1 digito dudoso) — mantén el encuadre para confirmar';
setTypeIndicator(detectedTypes);
return;
}
if (outcome.status === 'ambiguous') {
scannerLiveOverlay.textContent = 'Leido: ' + outcome.reads.slice(0, 2).join(' | ') + ' — varias coincidencias posibles';
setScanStatus(dedupeCount(outcome.matches) + ' posibles coincidencias — confirma visualmente cual es');
showDetected(outcome.reads.slice(0, 4).join(' | '));
setTypeIndicator(detectedTypes);
displayScanResults(outcome.matches, 'approx');
return;
}
setTypeIndicator(detectedTypes);
if (outcome.status === 'unknown') {
var ukey = 'U:' + outcome.reads[0];
liveVotes[ukey] = (liveVotes[ukey] || 0) + 1;
scannerLiveOverlay.textContent = 'Leido: ' + outcome.reads.slice(0, 2).join(' | ') +
(liveVotes[ukey] >= VOTES_TO_CONFIRM ? ' — no esta en el catalogo' : ' — comprobando...');
return;
}
if (liveTypeLock) {
scannerLiveOverlay.textContent = 'Parece ' + REF_TYPES[liveTypeLock.type].label + ' (' + REF_TYPES[liveTypeLock.type].hint + '), leyendo...';
} else {
scannerLiveOverlay.textContent = 'Buscando referencia... usa el zoom si los digitos se ven pequenos';
}
}
function acceptResult(matches, reads, types, how) {
stopLiveScan(true);
var n = dedupeCount(matches);
var msg = n + ' producto' + (n > 1 ? 's' : '') + ' encontrado' + (n > 1 ? 's' : '');
if (how === 'corrected') msg += ' (1 digito corregido, confirmado en 2 lecturas)';
setScanStatus(msg, 'found');
showDetected((reads || []).slice(0, 4).join(' | '));
setTypeIndicator(types, 'found');
displayScanResults(matches, how === 'corrected' ? 'approx' : 'exact');
}
function dedupeCount(matches) {
var seen = {};
matches.forEach(function(m) { seen[rowKey(m.row)] = 1; });
return Object.keys(seen).length;
}
var SCAN_RESULTS_MAX = 10;
function displayScanResults(matches, mode) {
scannerResults.innerHTML = '';
matches.slice(0, SCAN_RESULTS_MAX).forEach(function(m) {
var r = m.row;
var item = document.createElement('div');
item.className = 'scanner-result-item';
var body = item;
var photoUrl = m.type === 'EAN' ? findProductImageUrl(r.sheetFolder, m.ref) : null;
if (photoUrl) {
var photo = document.createElement('img');
photo.className = 'scanner-result-photo';
photo.src = photoUrl;
photo.alt = 'Foto: ' + r.section;
photo.decoding = 'async';
photo.onerror = function() { this.remove(); item.classList.remove('with-photo'); };
item.classList.add('with-photo');
item.appendChild(photo);
body = document.createElement('div');
body.className = 'scanner-result-text';
item.appendChild(body);
}
var titleDiv = document.createElement('div');
titleDiv.className = 'scanner-result-title';
var badge = document.createElement('span');
badge.className = 'scanner-type-badge';
badge.textContent = REF_TYPES[m.type].label;
titleDiv.appendChild(badge);
titleDiv.appendChild(document.createTextNode(r.sheet + ' — ' + r.section));
body.appendChild(titleDiv);
if (mode === 'approx' && m.dist) {
var note = document.createElement('div');
note.className = 'scanner-approx-note';
note.textContent = REF_TYPES[m.type].label + ' ' + m.ref + ' (leido ' + formatRefKey(m.type, m.read) + ')';
body.appendChild(note);
}
var detailDiv = document.createElement('div');
detailDiv.className = 'scanner-result-detail';
r.cells.forEach(function(c) {
if (!c.v) return;
var chip = document.createElement('span');
chip.className = 'scanner-detail-chip';
if (c.v.trim() === m.ref) chip.classList.add('scanner-detail-match');
chip.textContent = (c.label ? c.label + ': ' : '') + c.v;
detailDiv.appendChild(chip);
});
body.appendChild(detailDiv);
item.addEventListener('click', function() {
var cellOf = function(label) {
var c = r.cells.filter(function(x) { return x.label && x.label.trim().toUpperCase() === label; })[0];
return c && c.v ? c.v.trim() : '';
};
var ean = cellOf('EAN');
closeScanner();
if (ean && EAN_TEXT_RE.test(ean)) {
try {
openBarcodeModal(ean, r.sheetFolder, cellOf('ECI'));
return;
} catch (err) { /* si la ficha falla, se cae a la busqueda */ }
}
if (sheetSearch) {
sheetSearch.value = m.ref;
sheetSearch.dispatchEvent(new Event('input'));
}
});
scannerResults.appendChild(item);
});
if (matches.length > SCAN_RESULTS_MAX) {
var more = document.createElement('div');
more.className = 'scanner-no-match';
more.textContent = '... y ' + (matches.length - SCAN_RESULTS_MAX) + ' mas. Toca un resultado para ver en el catalogo.';
scannerResults.appendChild(more);
}
}
if (scanBtn) {
scanBtn.addEventListener('click', openScanner);
}
if (scannerCloseBtn) scannerCloseBtn.addEventListener('click', closeScanner);
if (scannerLiveStopBtn) scannerLiveStopBtn.addEventListener('click', restartLiveScan);
if (scannerZoomSlider) {
scannerZoomSlider.addEventListener('input', function() {
applyZoom(parseInt(scannerZoomSlider.value, 10) / 10);
});
}
function reportScannerError(msg) {
if (scannerModal && scannerModal.classList.contains('open')) {
setScanStatus('Error del escaner: ' + String(msg).slice(0, 120), 'error');
}
}
window.addEventListener('error', function(e) { reportScannerError(e.message); });
window.addEventListener('unhandledrejection', function(e) {
reportScannerError(e.reason && e.reason.message ? e.reason.message : e.reason);
});
closeOnBackdrop(scannerModal, closeScanner);
onEscape(scannerModal, closeScanner);

// ---- 90-main.js ----
function assetUrl(name, attr){
var v = document.body.getAttribute(attr);
return v ? name + '?v=' + v : name;
}
function fetchJson(url){
return fetch(url).then(function(r){
if(!r.ok) throw new Error(url + ': HTTP ' + r.status);
return r.json();
});
}
function showInitialPanel(){
var initial = (location.hash || '').replace('#', '') || 'home';
if(!document.getElementById(initial)){ initial = 'home'; }
activate(initial);
}
function renderPendingPanels(){
var id = activePanelId();
if(id && SHEETS_BY_SLUG[id]){ ensurePanelRendered(id); }
if(sheetSearch && sheetSearch.value.trim()){ runSearch(); }
}
showInitialPanel();
Promise.all([
fetchJson(assetUrl('catalog.json', 'data-catalog-v')),
fetch('events.txt').then(function(r){ return r.ok ? r.text() : ''; }).catch(function(){ return ''; }),
fetchJson(assetUrl('images.json', 'data-images-v')).catch(function(){ return {}; })
]).then(function(results){
setCatalog(results[0]);
APPLE_ON_THIS_DAY = parseEvents(results[1]);
setImageManifest(results[2]);
initTodayEvents();
renderPendingPanels();
}).catch(function(err){
console.error('Error cargando datos externos:', err);
var box = document.createElement('div');
box.setAttribute('role', 'alert');
box.style.cssText = 'margin:12px;padding:12px 14px;border:1px solid var(--err);border-radius:10px;' +
'color:var(--err);background:var(--panel);font-size:13px;line-height:1.5;';
box.textContent = 'No se han podido cargar los datos del catálogo (' + (err && err.message ? err.message : err) +
'). Si abres index.html con doble clic (file://), el navegador bloquea la carga: usa un servidor local, ' +
'por ejemplo "python3 -m http.server" dentro de la carpeta WebReferencias 2.1.';
var main = document.querySelector('.main');
if (main) main.insertBefore(box, main.firstChild);
});
})();
