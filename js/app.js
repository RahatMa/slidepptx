/* app.js — wires the page together: files in, preview, .pptx out. Runs fully in the browser. */
(function () {
  'use strict';
  const $ = (s, r) => (r || document).querySelector(s);
  const $$ = (s, r) => Array.from((r || document).querySelectorAll(s));
  const STORE_KEY = 'qbank-slides-settings-v1';

  const state = {
    workbookName: '',
    items: [],            // all parsed items
    parseWarnings: [],
    imageWarnings: [],
    images: new Map(),    // normalized file name -> dataURL
    overlay: null,
    selectedExams: null,  // Set of exam ids, null = all
    deck: null
  };

  // ------------------------------------------------------------ settings form
  function readSettings() {
    const v = (id) => $('#' + id).value;
    const c = (id) => $('#' + id).checked;
    const num = (id, d) => { const n = parseFloat(v(id)); return isFinite(n) ? n : d; };
    const fontSize = Math.max(12, Math.min(200, num('fontSize', 72)));
    return {
      fonts: { bn: v('fontBn').trim() || 'Kalpurush', lat: v('fontLat').trim() || 'Hind Siliguri', math: v('fontMath').trim() || 'Hind Siliguri', sym: 'Cambria Math' },
      fontSize,
      minFontSize: Math.max(10, Math.min(fontSize, num('minFontSize', 44))),
      lineSpacing: Math.max(0.8, Math.min(3, num('lineSpacing', 1.5))),
      slides: { question: c('slQuestion'), answer: c('slAnswer'), solution: c('slSolution') },
      notes: c('slNotes'),
      headerLabel: c('headerLabel'),
      optionLetters: v('optionLetters'),
      theme: v('theme'),
      layout: v('layoutDefault'),
      contentTop: num('contentTop', 4.9),
      contentBottom: num('contentBottom', 28.8),
      labels: { mcq: v('lblMcq') || 'MCQ', problem: v('lblProblem') || 'Problem', solution: v('lblSolution') || 'Solution', answer: v('lblAnswer'), cont: '(চলবে)' },
      splitByExam: c('splitByExam'),
      nativeEquations: c('nativeEq')
    };
  }

  function saveSettings() {
    const data = {};
    $$('[data-setting]').forEach((el) => { data[el.id] = el.type === 'checkbox' ? el.checked : el.value; });
    try { localStorage.setItem(STORE_KEY, JSON.stringify(data)); } catch (e) { /* private mode */ }
  }
  function loadSettings() {
    let data = null;
    try { data = JSON.parse(localStorage.getItem(STORE_KEY) || 'null'); } catch (e) { data = null; }
    if (!data) return;
    for (const [id, val] of Object.entries(data)) {
      const el = document.getElementById(id);
      if (!el) continue;
      if (el.type === 'checkbox') el.checked = !!val; else el.value = val;
    }
  }

  // ------------------------------------------------------------ text measuring
  const canvasCtx = document.createElement('canvas').getContext('2d');
  const fontAvail = {};
  function hasFont(font) {
    if (font in fontAvail) return fontAvail[font];
    const probe = 'mwআকখগঘ১২৩ ABCxyz';
    let differs = false;
    for (const fb of ['monospace', 'serif']) {
      canvasCtx.font = `72px ${fb}`;
      const a = canvasCtx.measureText(probe).width;
      canvasCtx.font = `72px "${font}", ${fb}`;
      const b = canvasCtx.measureText(probe).width;
      if (Math.abs(a - b) > 0.5) differs = true;
    }
    fontAvail[font] = differs;
    return differs;
  }
  function makeMeasurer() {
    const heur = QBLayout.heuristicMeasurer();
    return function (text, st) {
      if (!hasFont(st.font)) return heur(text, st);
      canvasCtx.font = `${st.italic ? 'italic ' : ''}${st.bold ? '700 ' : ''}100px "${st.font}"`;
      return canvasCtx.measureText(text).width / 100 * st.sizePt / 72;
    };
  }
  async function prepareFonts(S) {
    const fams = [...new Set([S.fonts.bn, S.fonts.lat, S.fonts.math])];
    await Promise.all(fams.map((f) => document.fonts.load(`40px "${f}"`, 'অa').catch(() => null)));
    for (const f of fams) delete fontAvail[f];
    const missing = fams.filter((f) => !hasFont(f));
    const note = $('#fontNote');
    if (missing.length) {
      note.hidden = false;
      note.textContent = `${missing.join(' and ')} ${missing.length > 1 ? 'are' : 'is'} not installed on this computer, so the preview uses a substitute and line fitting is estimated. The PowerPoint file still uses ${missing.length > 1 ? 'these fonts' : 'this font'}.`;
    } else note.hidden = true;
  }

  // ------------------------------------------------------------ files
  const normName = (s) => String(s).trim().split(/[\\/]/).pop().split('?')[0].toLowerCase();
  const stripExt = (s) => s.replace(/\.[a-z0-9]+$/i, '');

  function readAsDataURL(blob) {
    return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = () => rej(r.error); r.readAsDataURL(blob); });
  }
  function loadImg(src) {
    return new Promise((res, rej) => { const im = new Image(); im.onload = () => res(im); im.onerror = () => rej(new Error('Not an image')); im.src = src; });
  }
  async function toSupportedDataURL(dataURL) {
    // PowerPoint is happiest with PNG/JPEG; convert anything else (webp, bmp, svg…) to PNG
    const im = await loadImg(dataURL);
    if (/^data:image\/(png|jpe?g)/i.test(dataURL)) return { data: dataURL, w: im.naturalWidth, h: im.naturalHeight };
    const c = document.createElement('canvas');
    c.width = im.naturalWidth || 1200; c.height = im.naturalHeight || 800;
    const g = c.getContext('2d');
    g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
    g.drawImage(im, 0, 0, c.width, c.height);
    return { data: c.toDataURL('image/png'), w: c.width, h: c.height };
  }

  function addImageEntry(name, dataURL) {
    const n = normName(name);
    state.images.set(n, dataURL);
    if (!state.images.has(stripExt(n))) state.images.set(stripExt(n), dataURL);
  }

  async function handleFiles(fileList) {
    const files = Array.from(fileList || []);
    if (!files.length) return;
    let excel = null;
    for (const f of files) {
      const n = f.name.toLowerCase();
      if (/\.(xlsx|xlsm|xls|csv|ods)$/.test(n)) excel = f;
      else if (/\.zip$/.test(n)) await readZip(f);
      else if (/^image\//.test(f.type) || /\.(png|jpe?g|gif|webp|bmp|svg)$/.test(n)) addImageEntry(f.name, await readAsDataURL(f));
      else if (/\.pdf$/.test(n)) showPdfHint();
    }
    renderFileList();
    if (excel) await loadWorkbook(excel);
    else if (state.items.length) { await resolveImages(); rebuild(); }
  }

  async function readZip(file) {
    if (!window.JSZip) { setStatus('Zip files need JSZip, which did not load. Unzip and select the images instead.', 'warn'); return; }
    const zip = await window.JSZip.loadAsync(file);
    const jobs = [];
    zip.forEach((path, entry) => {
      if (entry.dir || /(^|\/)(__MACOSX|\.)/.test(path)) return;
      const lower = path.toLowerCase();
      if (/\.(png|jpe?g|gif|webp|bmp)$/.test(lower)) {
        jobs.push(entry.async('base64').then((b64) => {
          const ext = lower.split('.').pop().replace('jpg', 'jpeg');
          addImageEntry(path, `data:image/${ext};base64,${b64}`);
        }));
      } else if (/\.(xlsx|xls|csv)$/.test(lower)) {
        jobs.push(entry.async('arraybuffer').then((buf) => { state.pendingWorkbook = { name: path.split('/').pop(), buf }; }));
      }
    });
    await Promise.all(jobs);
    if (state.pendingWorkbook) {
      const wb = state.pendingWorkbook; state.pendingWorkbook = null;
      await loadWorkbookBuffer(wb.name, wb.buf);
    }
  }

  async function loadWorkbook(file) {
    const buf = await file.arrayBuffer();
    await loadWorkbookBuffer(file.name, buf);
  }

  async function loadWorkbookBuffer(name, buf) {
    let wb;
    try { wb = XLSX.read(buf, { type: 'array' }); } catch (e) { setStatus(`Couldn't read ${name}: ${e.message}`, 'error'); return; }
    let rows = [];
    const used = [];
    for (const sn of wb.SheetNames) {
      if (/^(instructions?|readme|help)$/i.test(sn)) continue;
      const r = XLSX.utils.sheet_to_json(wb.Sheets[sn], { defval: '', raw: false });
      if (!r.length) continue;
      const map = QBLayout.mapHeaders(Object.keys(r[0]));
      if (!map.question) continue;
      rows = rows.concat(r);
      used.push(sn);
    }
    if (!used.length) { setStatus(`No sheet in ${name} has a "Question" column. Start from the template.`, 'error'); return; }
    const { items, warnings } = QBLayout.normalizeRows(rows);
    state.workbookName = name.replace(/\.[^.]+$/, '');
    state.items = items;
    state.parseWarnings = warnings;
    state.selectedExams = null;
    state.urlCache = state.urlCache || new Map();
    $('#workbookInfo').textContent = `${name} · ${items.length} question${items.length === 1 ? '' : 's'} from ${used.join(', ')}`;
    renderExamFilter();
    await resolveImages();
    rebuild();
  }

  async function urlToDataURL(url) {
    let u = url;
    const gd = u.match(/drive\.google\.com\/file\/d\/([^/]+)/);
    if (gd) u = `https://drive.google.com/uc?export=download&id=${gd[1]}`;
    const resp = await fetch(u, { mode: 'cors' });
    if (!resp.ok) throw new Error('HTTP ' + resp.status);
    return readAsDataURL(await resp.blob());
  }

  async function resolveImages() {
    const warn = [];
    state.urlCache = state.urlCache || new Map();
    for (const it of state.items) {
      it.image = null;
      const ref = it.imageRef;
      if (!ref) continue;
      if (/^need_image$/i.test(ref)) { warn.push(`Row ${it.row}: image marked NEED_IMAGE — add the picture and put its file name in ImageURL.`); continue; }
      let src = state.images.get(normName(ref)) || state.images.get(stripExt(normName(ref)));
      if (!src && /^data:image\//i.test(ref)) src = ref;
      if (!src && /^https?:\/\//i.test(ref)) {
        if (state.urlCache.has(ref)) src = state.urlCache.get(ref);
        else {
          try { src = await urlToDataURL(ref); state.urlCache.set(ref, src); }
          catch (e) { state.urlCache.set(ref, null); }
        }
        if (!src) warn.push(`Row ${it.row}: couldn't download ${ref}. The site blocks it — save the image and add it as a file instead.`);
      } else if (!src) warn.push(`Row ${it.row}: image "${ref}" not found. Add it with the Excel file (or inside a .zip).`);
      if (src) {
        try { it.image = await toSupportedDataURL(src); }
        catch (e) { warn.push(`Row ${it.row}: "${ref}" is not a readable image.`); }
      }
    }
    state.imageWarnings = warn;
  }

  // ------------------------------------------------------------ UI pieces
  function renderFileList() {
    const names = new Set();
    for (const k of state.images.keys()) if (/\.[a-z0-9]+$/.test(k)) names.add(k);
    $('#imageInfo').textContent = names.size ? `${names.size} image${names.size === 1 ? '' : 's'} added` : '';
  }

  function examGroups(items) {
    const g = new Map();
    for (const it of items) { const k = it.examId || '(no ExamID)'; if (!g.has(k)) g.set(k, []); g.get(k).push(it); }
    return g;
  }

  function renderExamFilter() {
    const box = $('#examFilter');
    box.innerHTML = '';
    const groups = examGroups(state.items);
    $('#examFilterWrap').hidden = groups.size < 2;
    $('#splitWrap').hidden = groups.size < 2;
    for (const [id, list] of groups) {
      const lab = document.createElement('label');
      lab.className = 'chip';
      const cb = document.createElement('input');
      cb.type = 'checkbox'; cb.checked = true; cb.value = id;
      cb.addEventListener('change', () => {
        state.selectedExams = new Set($$('#examFilter input:checked').map((x) => x.value));
        rebuild();
      });
      lab.append(cb, document.createTextNode(` ${id} `));
      const n = document.createElement('span'); n.className = 'chip-n'; n.textContent = list.length;
      lab.appendChild(n);
      box.appendChild(lab);
    }
  }

  function selectedItems() {
    if (!state.selectedExams) return state.items;
    return state.items.filter((it) => state.selectedExams.has(it.examId || '(no ExamID)'));
  }

  function setStatus(msg, kind) {
    const el = $('#status');
    el.textContent = msg || '';
    el.dataset.kind = kind || '';
  }

  function showPdfHint() {
    $('#claudeCard').open = true;
    $('#claudeCard').scrollIntoView({ behavior: 'smooth', block: 'center' });
    setStatus('PDFs and photos need to become an Excel file first — see "Have photos or PDFs?" above.', 'warn');
  }

  // ------------------------------------------------------------ build + preview
  let rebuildTimer = null;
  function rebuild() { clearTimeout(rebuildTimer); rebuildTimer = setTimeout(doRebuild, 120); }

  async function doRebuild() {
    const items = selectedItems();
    const S = readSettings();
    saveSettings();
    await prepareFonts(S);
    if (!items.length) {
      state.deck = null;
      $('#preview').innerHTML = '';
      $('#emptyState').hidden = false;
      $('#generate').disabled = true;
      renderWarnings([]);
      $('#deckInfo').textContent = '';
      return;
    }
    const deck = QBLayout.buildDeck(items, S, makeMeasurer());
    state.deck = deck;
    $('#emptyState').hidden = true;
    $('#generate').disabled = false;
    const counts = deck.slides.reduce((a, s) => { a[s.kind] = (a[s.kind] || 0) + 1; return a; }, {});
    $('#deckInfo').textContent = `${deck.slides.length} slides · ${items.length} questions` +
      (counts.solution ? ` · ${counts.solution} solution slide${counts.solution === 1 ? '' : 's'}` : '');
    renderWarnings([...state.parseWarnings, ...state.imageWarnings, ...deck.warnings]);
    renderPreview(deck);
  }

  function renderWarnings(list) {
    const box = $('#warnings');
    box.innerHTML = '';
    box.hidden = !list.length;
    if (!list.length) return;
    const h = document.createElement('p');
    h.className = 'warn-title';
    h.textContent = `${list.length} thing${list.length === 1 ? '' : 's'} to check`;
    const ul = document.createElement('ul');
    list.slice(0, 60).forEach((w) => { const li = document.createElement('li'); li.textContent = w; ul.appendChild(li); });
    if (list.length > 60) { const li = document.createElement('li'); li.textContent = `…and ${list.length - 60} more`; ul.appendChild(li); }
    box.append(h, ul);
  }

  const KIND_LABEL = { mcq: 'Question', 'mcq-answer': 'Answer', problem: 'Question', solution: 'Solution' };

  function renderPreview(deck) {
    const grid = $('#preview');
    grid.innerHTML = '';
    const avail = Math.max(200, grid.clientWidth);
    const cols = Math.max(1, Math.floor((avail + 16) / (300 + 16)));
    const width = Math.floor((avail - 16 * (cols - 1)) / cols);
    grid.style.gridTemplateColumns = `repeat(${cols}, ${width}px)`;
    const frag = document.createDocumentFragment();
    deck.slides.forEach((s, i) => {
      const fig = document.createElement('figure');
      fig.className = 'thumb';
      fig.tabIndex = 0;
      fig.setAttribute('role', 'button');
      fig.setAttribute('aria-label', `Open slide ${i + 1} larger`);
      fig.appendChild(QBPreview.renderSlide(s, deck, width, state.overlay));
      const cap = document.createElement('figcaption');
      cap.textContent = `${i + 1} · ${KIND_LABEL[s.kind] || s.kind}${s.item.qno ? ' · Q' + s.item.qno : ''} · ${s.size} pt`;
      fig.appendChild(cap);
      const open = () => openLarge(i);
      fig.addEventListener('click', open);
      fig.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); open(); } });
      frag.appendChild(fig);
    });
    grid.appendChild(frag);
  }

  function openLarge(i) {
    const deck = state.deck;
    if (!deck) return;
    const dlg = $('#large');
    const body = $('#largeBody');
    let idx = i;
    const draw = () => {
      body.innerHTML = '';
      const w = Math.min(window.innerWidth - 64, 1400);
      body.appendChild(QBPreview.renderSlide(deck.slides[idx], deck, w, state.overlay));
      $('#largeCap').textContent = `Slide ${idx + 1} of ${deck.slides.length}`;
    };
    dlg.onkeydown = (e) => {
      if (e.key === 'ArrowRight' && idx < deck.slides.length - 1) { idx++; draw(); }
      if (e.key === 'ArrowLeft' && idx > 0) { idx--; draw(); }
    };
    $('#largePrev').onclick = () => { if (idx > 0) { idx--; draw(); } };
    $('#largeNext').onclick = () => { if (idx < deck.slides.length - 1) { idx++; draw(); } };
    draw();
    dlg.showModal();
  }

  // ------------------------------------------------------------ download
  const safeName = (s) => String(s).replace(/[\\/:*?"<>|]+/g, '-').replace(/\s+/g, '_').slice(0, 80) || 'slides';

  function downloadBlob(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = fileName;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }

  async function generate() {
    const items = selectedItems();
    if (!items.length) return;
    const S = readSettings();
    const btn = $('#generate');
    btn.disabled = true;
    setStatus('Building PowerPoint…');
    try {
      const groups = S.splitByExam ? [...examGroups(items).entries()] : [[state.workbookName || 'slides', items]];
      for (const [name, list] of groups) {
        const deck = QBLayout.buildDeck(list, S, makeMeasurer());
        const pres = QBWriter.writeDeck(deck, window.PptxGenJS, { overlay: state.overlay, title: name, nativeEquations: S.nativeEquations });
        const blob = await QBWriter.toFile(pres, window.JSZip, 'blob');
        downloadBlob(blob, safeName(name) + '.pptx');
        await new Promise((r) => setTimeout(r, 400)); // let the browser start each download
      }
      setStatus(groups.length > 1 ? `Downloaded ${groups.length} PowerPoint files.` : 'Downloaded the PowerPoint file.', 'ok');
    } catch (e) {
      console.error(e);
      setStatus('Something went wrong while writing the file: ' + e.message, 'error');
    } finally {
      btn.disabled = false;
    }
  }

  // ------------------------------------------------------------ init
  function init() {
    loadSettings();
    $('#claudePrompt').value = window.QB_CLAUDE_PROMPT || '';
    $('#copyPrompt').addEventListener('click', async () => {
      const t = $('#claudePrompt');
      try { await navigator.clipboard.writeText(t.value); }
      catch (e) { t.select(); document.execCommand('copy'); }
      $('#copyPrompt').textContent = 'Copied';
      setTimeout(() => { $('#copyPrompt').textContent = 'Copy prompt'; }, 1800);
    });

    const input = $('#fileInput');
    input.addEventListener('change', () => { handleFiles(input.files); input.value = ''; });
    const drop = $('#drop');
    ['dragenter', 'dragover'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
    drop.addEventListener('drop', (e) => handleFiles(e.dataTransfer.files));

    $('#overlayInput').addEventListener('change', async (e) => {
      const f = e.target.files[0];
      state.overlay = f ? (await toSupportedDataURL(await readAsDataURL(f))).data : null;
      $('#overlayClear').hidden = !state.overlay;
      rebuild();
    });
    $('#overlayClear').addEventListener('click', () => { state.overlay = null; $('#overlayInput').value = ''; $('#overlayClear').hidden = true; rebuild(); });

    $$('[data-setting]').forEach((el) => el.addEventListener(el.type === 'checkbox' || el.tagName === 'SELECT' ? 'change' : 'input', rebuild));
    $('#resetSettings').addEventListener('click', () => { try { localStorage.removeItem(STORE_KEY); } catch (e) { /* ignore */ } location.reload(); });
    $('#generate').addEventListener('click', generate);
    $('#largeClose').addEventListener('click', () => $('#large').close());
    let resizeT = null;
    window.addEventListener('resize', () => { clearTimeout(resizeT); resizeT = setTimeout(() => { if (state.deck) renderPreview(state.deck); }, 200); });

    $('#loadExample').addEventListener('click', async (e) => {
      e.preventDefault();
      setStatus('Loading the example…');
      try {
        const [xl, img] = await Promise.all([fetch('template/question_template.xlsx'), fetch('examples/images/mcq24.png')]);
        if (img.ok) addImageEntry('mcq24.png', await readAsDataURL(await img.blob()));
        renderFileList();
        await loadWorkbookBuffer('question_template.xlsx', await xl.arrayBuffer());
        setStatus('');
      } catch (err) { setStatus('Could not load the example: ' + err.message, 'error'); }
    });

    if (!window.XLSX || !window.PptxGenJS) setStatus('The spreadsheet or PowerPoint library did not load. Check your internet connection and reload.', 'error');
  }

  document.addEventListener('DOMContentLoaded', init);
})();
