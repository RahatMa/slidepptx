/*
 * layout.js — turns spreadsheet rows into slide models (positions in inches).
 *
 * The same model feeds the PPTX writer (pptx-writer.js) and the HTML preview,
 * so what you see in the browser is what lands in PowerPoint.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./latex.js'));
  else root.QBLayout = factory(root.QBLatex);
})(typeof self !== 'undefined' ? self : this, function (QBLatex) {
  'use strict';

  // ------------------------------------------------------------------ settings
  const DEFAULTS = {
    slideW: 56, slideH: 31.5,             // inches (PowerPoint's maximum, 16:9)
    marginX: 3,                           // left/right content margin
    contentTop: 4.9,                      // below the header band (slide master)
    contentBottom: 28.8,                  // above the footer band
    fontSize: 72, minFontSize: 44, lineSpacing: 1.5,
    lineFactor: 1.62,                     // font line height / font size (Kalpurush ≈1.58, Hind Siliguri ≈1.62)
    fonts: { bn: 'Kalpurush', lat: 'Hind Siliguri', math: 'Hind Siliguri', sym: 'Cambria Math' },
    theme: 'dark',
    optionLetters: 'en',                  // 'en' → A B C D, 'bn' → ক খ গ ঘ
    slides: { question: true, answer: true, solution: true },
    notes: true,
    headerLabel: true,
    labels: { mcq: 'MCQ', problem: 'Problem', solution: 'Solution', answer: 'উত্তর', cont: '(চলবে)' },
    padNumber: 2,
    layout: 'auto',                       // auto | side | grid | stack (a row's Layout column wins)
    nativeEquations: true                 // formulas as PowerPoint equation objects (Insert ▸ Equation)
  };

  const THEMES = {
    dark: {
      bg: '000000', text: 'FFFFFF', optFill: 'FFFFFF', optText: '000000', labelFill: 'E6E6E6', labelText: '000000',
      badge: '8C5CFF', badgeText: 'FFFFFF', badgeNumFill: 'FFFFFF', badgeNumText: '8C5CFF',
      correctFill: '16A34A', correctText: 'FFFFFF', correctLabelFill: 'FFFFFF', correctLabelText: '15803D',
      answerText: '4ADE80'
    },
    light: {
      bg: 'FFFFFF', text: '111111', optFill: 'F1F1F6', optText: '111111', labelFill: 'FFFFFF', labelText: '111111',
      badge: '6D4AFF', badgeText: 'FFFFFF', badgeNumFill: 'FFFFFF', badgeNumText: '6D4AFF',
      correctFill: '16A34A', correctText: 'FFFFFF', correctLabelFill: 'FFFFFF', correctLabelText: '15803D',
      answerText: '15803D'
    }
  };

  function mergeSettings(s) {
    const out = JSON.parse(JSON.stringify(DEFAULTS));
    if (!s) return out;
    for (const k of Object.keys(s)) {
      if (s[k] && typeof s[k] === 'object' && !Array.isArray(s[k]) && out[k] && typeof out[k] === 'object') Object.assign(out[k], s[k]);
      else if (s[k] !== undefined) out[k] = s[k];
    }
    return out;
  }

  // ------------------------------------------------------------------ rows → items
  const ALIASES = {
    examId: ['examid', 'exam', 'examcode', 'setid', 'chapter'],
    qno: ['qno', 'questionno', 'qnumber', 'no', 'sl', 'serial', 'number'],
    type: ['type', 'questiontype', 'qtype', 'kind'],
    question: ['question', 'q', 'stem', 'problem', 'questiontext'],
    image: ['imageurl', 'image', 'img', 'imagefile', 'picture', 'figure', 'diagram'],
    a: ['optiona', 'opta', 'a', 'option1', 'optionka'],
    b: ['optionb', 'optb', 'b', 'option2', 'optionkha'],
    c: ['optionc', 'optc', 'c', 'option3', 'optionga'],
    d: ['optiond', 'optd', 'd', 'option4', 'optiongha'],
    answer: ['answer', 'ans', 'correct', 'correctanswer', 'key'],
    solution: ['solution', 'explanation', 'sol', 'solve', 'details'],
    layout: ['layout', 'optionlayout', 'arrangement']
  };

  function keyOf(h) { return String(h).toLowerCase().replace(/[^a-z0-9]/g, ''); }

  function mapHeaders(headers) {
    const map = {};
    for (const h of headers) {
      const k = keyOf(h);
      for (const field of Object.keys(ALIASES)) {
        if (!map[field] && ALIASES[field].includes(k)) map[field] = h;
      }
    }
    return map;
  }

  const BN_DIGITS = '০১২৩৪৫৬৭৮৯';
  function toLatinDigits(s) { return String(s).replace(/[০-৯]/g, d => String(BN_DIGITS.indexOf(d))); }

  function answerIndex(v) {
    const s = toLatinDigits(String(v == null ? '' : v)).trim().replace(/^[(\[]|[)\].:]$/g, '').trim().toLowerCase().replace(/^option\s*/, '');
    if (!s) return -1;
    const table = { a: 0, b: 1, c: 2, d: 3, '1': 0, '2': 1, '3': 2, '4': 3, 'ক': 0, 'খ': 1, 'গ': 2, 'ঘ': 3, 'i': 0, 'ii': 1, 'iii': 2, 'iv': 3 };
    return table[s] !== undefined ? table[s] : -1;
  }

  function cell(row, map, field) {
    const h = map[field];
    if (!h) return '';
    const v = row[h];
    return v === null || v === undefined ? '' : String(v);
  }

  /** rows: array of objects keyed by header (SheetJS sheet_to_json with defval:''). */
  function normalizeRows(rows) {
    const warnings = [];
    if (!rows.length) return { items: [], warnings: ['The sheet has no rows.'] };
    const headers = Object.keys(rows.reduce((acc, r) => Object.assign(acc, r), {}));
    const map = mapHeaders(headers);
    if (!map.question) return { items: [], warnings: ['No "Question" column found. Use the template headers.'] };

    const items = [];
    rows.forEach((row, i) => {
      const excelRow = (row.__rowNum__ !== undefined ? row.__rowNum__ + 1 : i + 2);
      const q = cell(row, map, 'question');
      const opts = ['a', 'b', 'c', 'd'].map(k => cell(row, map, k));
      if (!q.trim() && !opts.some(o => o.trim())) return;
      let type = keyOf(cell(row, map, 'type'));
      if (['mcq', 'mc', 'objective', 'bohunirbachoni'].includes(type)) type = 'mcq';
      else if (['problem', 'cq', 'written', 'math', 'creative', 'srijonshil', 'short', 'broad'].includes(type)) type = 'problem';
      else type = opts.some(o => o.trim()) ? 'mcq' : 'problem';

      const answerRaw = cell(row, map, 'answer').trim();
      const ai = answerIndex(answerRaw);
      const layoutRaw = keyOf(cell(row, map, 'layout'));
      const item = {
        row: excelRow,
        examId: cell(row, map, 'examId').trim(),
        qno: toLatinDigits(cell(row, map, 'qno')).trim(),
        type,
        question: QBLatex.parseRich(q),
        options: opts.map(o => QBLatex.parseRich(o)),
        answerRaw,
        answerIndex: ai,
        solution: QBLatex.parseRich(cell(row, map, 'solution')),
        imageRef: cell(row, map, 'image').trim(),
        layout: ['side', 'grid', 'stack', 'auto'].includes(layoutRaw) ? layoutRaw : '',
        image: null
      };
      if (type === 'mcq') {
        const filled = item.options.filter(o => o.length).length;
        if (filled < 2) warnings.push(`Row ${excelRow}: MCQ has fewer than 2 options.`);
        if (ai < 0 && answerRaw) warnings.push(`Row ${excelRow}: answer "${answerRaw}" not understood (use A–D or ক–ঘ).`);
        if (ai < 0 && !answerRaw) warnings.push(`Row ${excelRow}: no answer — the answer slide is skipped.`);
      }
      items.push(item);
    });
    return { items, warnings };
  }

  // ------------------------------------------------------------------ measuring
  /** Fallback width estimate (em units) when no canvas is available. */
  function heuristicEm(ch, kind) {
    const c = ch.codePointAt(0);
    if (ch === ' ') return kind === 'bn' ? 0.33 : 0.25;
    if (c >= 0x0980 && c <= 0x09FF) {
      if (c === 0x09CD || c === 0x0981 || c === 0x09BC || (c >= 0x09C1 && c <= 0x09C4) || c === 0x09E2 || c === 0x09E3) return 0;
      if (c === 0x09BE || c === 0x09BF || c === 0x09C0) return 0.24;
      if (c === 0x09C7 || c === 0x09C8) return 0.31;
      if (c === 0x09CB || c === 0x09CC) return 0.55;
      if (c === 0x0982 || c === 0x0983) return 0.3;
      if (c >= 0x09E6 && c <= 0x09EF) return 0.5;
      return 0.6;
    }
    if (c === 0x0964 || c === 0x0965) return 0.3;
    if (c >= 0x0300 && c <= 0x036F) return 0;
    if (c >= 0x20D0 && c <= 0x20FF) return 0;
    if (/[a-z]/.test(ch)) return /[mw]/.test(ch) ? 0.78 : /[ijlft]/.test(ch) ? 0.3 : 0.5;
    if (/[A-Z]/.test(ch)) return /[MW]/.test(ch) ? 0.85 : /[IJ]/.test(ch) ? 0.32 : 0.64;
    if (/[0-9]/.test(ch)) return 0.5;
    if (/[.,:;'!|]/.test(ch)) return 0.26;
    if (/[()\[\]{}]/.test(ch)) return 0.33;
    if (c > 0x2000) return 0.7;
    return 0.55;
  }

  function heuristicMeasurer() {
    return function measure(text, st) {
      let em = 0;
      for (const ch of Array.from(text)) em += heuristicEm(ch, st.kind);
      if (st.bold) em *= 1.06;
      return em * st.sizePt / 72; // inches
    };
  }

  function makeMetrics(settings, measure) {
    const S = settings;
    const fontFor = (kind) => S.fonts[kind] || S.fonts.lat;
    const scriptScale = 0.66;

    function runPieces(runs, sizePt) {
      const pieces = []; // {text, w, space}
      for (const r of runs) {
        for (const seg of QBLatex.scriptSegments(r)) {
          const parts = seg.text.split(/(\s+)/);
          for (const p of parts) {
            if (!p) continue;
            const st = { font: fontFor(seg.kind), kind: seg.kind, sizePt: (r.sup || r.sub) ? sizePt * scriptScale : sizePt, italic: r.italic, bold: r.bold };
            pieces.push({ text: p, w: measure(p, st), space: /^\s+$/.test(p) });
          }
        }
      }
      return pieces;
    }

    function wrapCount(runs, width, sizePt) {
      if (!runs || !runs.length) return 1;
      const pieces = runPieces(runs, sizePt);
      const words = [];
      let cur = null;
      for (const p of pieces) {
        if (p.space) { if (cur) { words.push(cur); cur = null; } words.push({ space: true, w: p.w }); }
        else { if (!cur) cur = { w: 0 }; cur.w += p.w; }
      }
      if (cur) words.push(cur);
      const max = width * 0.95; // safety for renderer differences
      let lines = 1, lineW = 0, pendingSpace = 0;
      for (const w of words) {
        if (w.space) { pendingSpace += w.w; continue; }
        if (lineW === 0) { lineW = w.w; }
        else if (lineW + pendingSpace + w.w <= max) { lineW += pendingSpace + w.w; }
        else { lines++; lineW = w.w; }
        pendingSpace = 0;
        if (lineW > max) { const extra = Math.ceil(lineW / max) - 1; lines += extra; lineW = lineW - extra * max; }
      }
      return lines;
    }

    function textWidth(runs, sizePt) { return runPieces(runs, sizePt).reduce((a, p) => a + p.w, 0); }
    function lineH(sizePt, spacing) { return sizePt / 72 * S.lineFactor * (spacing || S.lineSpacing); }
    // a stacked fraction (native equation) makes its line ~65% taller
    const fracExtra = (p) => (S.nativeEquations !== false && p.some(r => r.eqFrac) ? 0.65 : 0);
    function blockLines(paras, width, sizePt) { return (paras || []).reduce((a, p) => a + wrapCount(p, width, sizePt) + fracExtra(p), 0); }
    function blockH(paras, width, sizePt) { return blockLines(paras, width, sizePt) * lineH(sizePt); }
    return { wrapCount, textWidth, lineH, blockH, blockLines };
  }

  // ------------------------------------------------------------------ helpers
  function sizesToTry(S) {
    const out = [];
    for (let s = S.fontSize; s >= S.minFontSize; s -= 4) out.push(s);
    if (out[out.length - 1] !== S.minFontSize) out.push(S.minFontSize);
    return out;
  }

  function numberLabel(S, qno) {
    if (!qno) return '';
    return /^\d+$/.test(qno) ? qno.padStart(S.padNumber || 0, '0') : qno;
  }

  function letters(S) { return S.optionLetters === 'bn' ? ['ক', 'খ', 'গ', 'ঘ'] : ['A', 'B', 'C', 'D']; }

  function fitImage(img, maxW, maxH) {
    if (!img || maxW <= 0 || maxH <= 0) return null;
    const pxPerIn = 24; // don't blow small images up past ~24 px per inch
    const capW = Math.min(maxW, img.w / pxPerIn), capH = Math.min(maxH, img.h / pxPerIn);
    const scale = Math.min(capW / img.w, capH / img.h);
    return { w: img.w * scale, h: img.h * scale };
  }

  function headerElements(S, T, kind, qno) {
    if (!S.headerLabel) return [];
    const num = numberLabel(S, qno);
    if (kind === 'problem' || kind === 'problem-solution') {
      const word = kind === 'problem' ? S.labels.problem : S.labels.solution;
      return [{ kind: 'text', name: 'Header Label', x: 14, y: 0.6, w: 28, h: 2.6, size: 66, bold: true, color: T.text, align: 'center', valign: 'middle', spacing: 1, paragraphs: [[{ text: (word + ' ' + num).trim() }]] }];
    }
    const word = kind === 'mcq' ? S.labels.mcq : S.labels.solution;
    const x = 22.6, y = 1.0, w = 14.8, h = 2.0;
    const els = [{ kind: 'text', name: 'Header Badge', x, y, w, h, shape: 'roundRect', radius: 0.35, fill: T.badge, size: 54, bold: true, color: T.badgeText, align: 'center', valign: 'middle', spacing: 1, inset: [0, num ? 4.4 : 0, 0, 0], paragraphs: [[{ text: word }]] }];
    if (num) els.push({ kind: 'text', name: 'Header Number', x: x + w - 4.15, y: y + 0.18, w: 3.95, h: h - 0.36, shape: 'roundRect', radius: 0.3, fill: T.badgeNumFill, size: 54, bold: true, color: T.badgeNumText, align: 'center', valign: 'middle', spacing: 1, paragraphs: [[{ text: num }]] });
    return els;
  }

  // ------------------------------------------------------------------ MCQ
  function optionMetrics(S, size, w) {
    const k = size / 72;
    const L = 1.75 * k, padX = 0.4 * k + 0.1, gapLT = 0.5 * k + 0.1;
    const insetL = padX + L + gapLT, insetR = 0.35;
    return { L, padX, insetL, insetR, textW: w - insetL - insetR };
  }

  function optionElements(S, T, item, i, box, size, reveal) {
    const correct = reveal && i === item.answerIndex;
    const om = optionMetrics(S, size, box.w);
    const L = letters(S);
    return [
      { kind: 'text', name: `Option ${L[i]}`, x: box.x, y: box.y, w: box.w, h: box.h, shape: 'roundRect', radius: 0.3, fill: correct ? T.correctFill : T.optFill, size, color: correct ? T.correctText : T.optText, align: 'left', valign: 'middle', inset: [0, om.insetR, 0, om.insetL], paragraphs: item.options[i].length ? item.options[i] : [[{ text: ' ' }]] },
      { kind: 'text', name: `Option ${L[i]} Label`, x: box.x + om.padX, y: box.y + (box.h - om.L) / 2, w: om.L, h: om.L, shape: 'roundRect', radius: 0.15, fill: correct ? T.correctLabelFill : T.labelFill, size, color: correct ? T.correctLabelText : T.labelText, align: 'center', valign: 'middle', spacing: 1, paragraphs: [[{ text: L[i], bold: correct }]] }
    ];
  }

  function tryMCQ(S, M, item, size, mode) {
    const T = THEMES[S.theme] || THEMES.dark;
    const top = S.contentTop, bottom = S.contentBottom, left = S.marginX, fullW = S.slideW - 2 * S.marginX;
    const lh = M.lineH(size);
    const optIdx = [0, 1, 2, 3].filter(i => item.options[i].length || i < 2);
    const k = size / 72;
    const boxPadY = 0.25 * k + 0.2, minBoxH = 1.75 * k + 0.9 * k;
    const gap = 0.55;
    const els = [];
    let qW, oX, oW, cols;

    if (mode === 'side') { qW = 33.5; oX = left + qW + 2; oW = S.slideW - S.marginX - oX; cols = 1; }
    else if (mode === 'grid') { qW = fullW; oX = left; oW = (fullW - 1.2) / 2; cols = 2; }
    else { qW = fullW; oX = left; oW = fullW; cols = 1; }

    const maxLines = mode === 'side' ? 2 : mode === 'grid' ? 3 : 99;
    const om = optionMetrics(S, size, oW);
    const optLines = optIdx.map(i => M.blockLines(item.options[i].length ? item.options[i] : [[{ text: ' ' }]], om.textW, size));
    if (optLines.some(n => n > maxLines)) return null;
    const optH = optLines.map(n => Math.max(minBoxH, n * lh + 2 * boxPadY));

    const qLines = M.blockLines(item.question, qW, size);
    const qH = qLines * lh;
    let y = top + qH;

    // image under the question
    let imgEl = null;
    if (item.image) {
      const imgTop = y + 0.5;
      let availH;
      if (mode === 'side') availH = bottom - imgTop;
      else {
        // leave room for the option block under the image
        const rows = cols === 2 ? [Math.max(optH[0] || 0, optH[1] || 0), Math.max(optH[2] || 0, optH[3] || 0)].filter(Boolean) : optH;
        const block = rows.reduce((a, b) => a + b, 0) + gap * (rows.length - 1);
        availH = bottom - imgTop - 0.8 - block;
      }
      const fit = fitImage(item.image, qW, Math.min(availH, 16));
      if (!fit || fit.h < Math.min(4, item.image.h / 24)) return null;
      const ix = mode === 'side' ? left : left + (qW - fit.w) / 2;
      imgEl = { kind: 'image', name: 'Question Image', x: ix, y: imgTop, w: fit.w, h: fit.h, data: item.image.data };
      y = imgTop + fit.h;
    }
    if (y > bottom) return null;

    // options
    const optEls = [];
    if (mode === 'side') {
      const total = optH.reduce((a, b) => a + b, 0) + gap * (optH.length - 1);
      let oy = top + 2.7;
      if (oy + total > bottom) oy = top;
      if (oy + total > bottom) return null;
      optIdx.forEach((i, n) => { optEls.push({ i, box: { x: oX, y: oy, w: oW, h: optH[n] } }); oy += optH[n] + gap; });
    } else {
      let oy = y + 0.8;
      if (cols === 2) {
        for (let r = 0; r < optIdx.length; r += 2) {
          const h = Math.max(optH[r], optH[r + 1] || 0);
          optIdx.slice(r, r + 2).forEach((i, c) => optEls.push({ i, box: { x: oX + c * (oW + 1.2), y: oy, w: oW, h } }));
          oy += h + gap;
        }
      } else {
        optIdx.forEach((i, n) => { optEls.push({ i, box: { x: oX, y: oy, w: oW, h: optH[n] } }); oy += optH[n] + gap; });
      }
      if (oy - gap > bottom) return null;
    }

    const qBoxH = (imgEl || mode !== 'side') ? qH + 0.15 : Math.max(qH + 0.15, bottom - top);
    els.push({ kind: 'text', name: 'Question', x: left, y: top, w: qW, h: qBoxH, size, color: T.text, align: 'left', valign: 'top', paragraphs: item.question.length ? item.question : [[{ text: ' ' }]] });
    if (imgEl) els.push(imgEl);
    return { size, mode, els, optEls };
  }

  function layoutMCQ(S, M, item, reveal, forced) {
    const T = THEMES[S.theme] || THEMES.dark;
    const pref = forced || item.layout || S.layout || 'auto';
    const modes = pref === 'auto' ? ['side', 'grid', 'stack'] : [pref];
    let res = null, overflow = false;
    outer: for (const size of sizesToTry(S)) {
      for (const mode of modes) { res = tryMCQ(S, M, item, size, mode); if (res) break outer; }
    }
    if (!res) {
      overflow = true;
      for (const mode of [...modes, 'stack']) {
        res = tryMCQ(Object.assign({}, S, { contentBottom: 1e9 }), M, item, S.minFontSize, mode);
        if (res) break;
      }
    }
    const els = [...headerElements(S, T, 'mcq', item.qno), ...res.els];
    for (const o of res.optEls) els.push(...optionElements(S, T, item, o.i, o.box, res.size, reveal));
    return { els, size: res.size, mode: res.mode, overflow };
  }

  // ------------------------------------------------------------------ Problem
  function layoutProblem(S, M, item) {
    const T = THEMES[S.theme] || THEMES.dark;
    const top = S.contentTop, bottom = S.contentBottom, left = S.marginX, fullW = S.slideW - 2 * S.marginX;
    let chosen = null;
    for (const size of sizesToTry(S)) {
      const lh = M.lineH(size);
      const lines = M.blockLines(item.question, fullW, size);
      const qH = lines * lh;
      let ok = top + qH <= bottom, imgEl = null;
      if (ok && item.image) {
        const imgTop = top + qH + 0.6;
        const fit = fitImage(item.image, fullW, Math.min(bottom - imgTop, 18));
        if (!fit || fit.h < Math.min(4, item.image.h / 24)) ok = false;
        else imgEl = { kind: 'image', name: 'Question Image', x: left + (fullW - fit.w) / 2, y: imgTop, w: fit.w, h: fit.h, data: item.image.data };
      }
      if (ok) { chosen = { size, qH, lines, imgEl }; break; }
    }
    let overflow = false;
    if (!chosen) {
      overflow = true;
      const size = S.minFontSize;
      chosen = { size, qH: M.blockH(item.question, fullW, size), lines: M.blockLines(item.question, fullW, size), imgEl: null };
    }
    const center = item.question.length === 1 && chosen.lines <= 3;
    const els = [...headerElements(S, T, 'problem', item.qno)];
    els.push({ kind: 'text', name: 'Question', x: left, y: top, w: fullW, h: chosen.imgEl ? chosen.qH + 0.15 : Math.max(chosen.qH + 0.15, bottom - top), size: chosen.size, color: T.text, align: center ? 'center' : 'left', valign: 'top', paragraphs: item.question });
    if (chosen.imgEl) els.push(chosen.imgEl);
    return { els, size: chosen.size, overflow };
  }

  // ------------------------------------------------------------------ Solution
  function answerParagraph(S, item) {
    const label = S.labels.answer ? S.labels.answer + ': ' : '';
    if (item.type === 'mcq' && item.answerIndex >= 0) {
      const L = letters(S)[item.answerIndex];
      const opt = item.options[item.answerIndex] || [];
      const runs = [{ text: label + '(' + L + ') ', bold: true }];
      (opt[0] || []).forEach(r => runs.push(Object.assign({}, r, { bold: true })));
      return [runs];
    }
    if (item.type === 'problem' && item.answerRaw) {
      const parsed = QBLatex.parseRich(item.answerRaw);
      if (!parsed.length) return [];
      return [[{ text: label, bold: true }, ...parsed[0].map(r => Object.assign({}, r, { bold: true }))], ...parsed.slice(1)];
    }
    return [];
  }

  function layoutSolution(S, M, item) {
    const T = THEMES[S.theme] || THEMES.dark;
    const top = S.contentTop, bottom = S.contentBottom, left = S.marginX, fullW = S.slideW - 2 * S.marginX;
    const ans = answerParagraph(S, item);
    const sol = item.solution;
    const kind = item.type === 'mcq' ? 'mcq-solution' : 'problem-solution';

    const build = (size, ansParas, solParas, isCont) => {
      const lh = M.lineH(size);
      const els = [...headerElements(S, T, kind, item.qno)];
      let y = top;
      if (ansParas.length) {
        const h = M.blockLines(ansParas, fullW, size) * lh;
        els.push({ kind: 'text', name: 'Answer', x: left, y, w: fullW, h: h + 0.1, size, color: T.answerText, align: 'left', valign: 'top', paragraphs: ansParas });
        y += h + 0.4;
      }
      if (solParas.length) {
        const paras = isCont ? solParas : solParas;
        els.push({ kind: 'text', name: 'Solution', x: left, y, w: fullW, h: Math.max(0.5, bottom - y), size, color: T.text, align: 'left', valign: 'top', paragraphs: paras });
      }
      return els;
    };

    for (const size of sizesToTry(S)) {
      const lh = M.lineH(size);
      const h = (ans.length ? M.blockLines(ans, fullW, size) * lh + 0.4 : 0) + M.blockLines(sol, fullW, size) * lh;
      if (top + h <= bottom) return [{ els: build(size, ans, sol, false), size, overflow: false }];
    }

    // paginate at a readable size
    const size = Math.max(S.minFontSize, Math.round(S.fontSize * 0.83 / 2) * 2);
    const lh = M.lineH(size);
    const ansH = ans.length ? M.blockLines(ans, fullW, size) * lh + 0.4 : 0;
    const avail = bottom - top - (S.labels.cont ? lh : 0); // keep a line free for "(চলবে)"
    const heights = sol.map(p => M.blockLines([p], fullW, size) * lh);
    const pack = (limit) => {
      const pages = [];
      let cur = [], used = ansH;
      sol.forEach((p, i) => {
        const h = heights[i];
        if (cur.length && (used + h > avail || used + h > limit)) { pages.push(cur); cur = []; used = 0; }
        cur.push(p); used += h;
      });
      if (cur.length) pages.push(cur);
      return pages;
    };
    let pages = pack(Infinity);
    if (pages.length > 1) { // spread the lines evenly instead of leaving a near-empty last slide
      const total = ansH + heights.reduce((a, b) => a + b, 0);
      const balanced = pack(total / pages.length + lh * 0.5);
      if (balanced.length === pages.length) pages = balanced;
    }
    return pages.map((paras, idx) => {
      const ps = idx < pages.length - 1 && S.labels.cont ? [...paras, [{ text: S.labels.cont, italic: true }]] : paras;
      const tooTall = paras.some(p => M.wrapCount(p, fullW, size) * lh > avail);
      return { els: build(size, idx === 0 ? ans : [], ps, idx > 0), size, overflow: tooTall };
    });
  }

  // ------------------------------------------------------------------ notes
  function notesFor(S, item) {
    const parts = [];
    const L = letters(S);
    if (item.type === 'mcq' && item.answerIndex >= 0) parts.push(`${S.labels.answer || 'Answer'}: (${L[item.answerIndex]}) ${QBLatex.toPlain(item.options[item.answerIndex])}`);
    else if (item.answerRaw) parts.push(`${S.labels.answer || 'Answer'}: ${item.answerRaw}`);
    if (item.solution.length) parts.push(QBLatex.toPlain(item.solution));
    if (item.examId) parts.push(`[${item.examId}${item.qno ? ' · ' + item.qno : ''}]`);
    return parts.join('\n\n');
  }

  // ------------------------------------------------------------------ deck
  function buildDeck(items, settings, measure) {
    const S = mergeSettings(settings);
    const M = makeMetrics(S, measure || heuristicMeasurer());
    const slides = [];
    const warnings = [];
    const tag = (item) => `Row ${item.row}${item.qno ? ' (Q' + item.qno + ')' : ''}`;

    for (const item of items) {
      const notes = S.notes ? notesFor(S, item) : '';
      if (item.type === 'mcq') {
        const q = layoutMCQ(S, M, item, false);
        if (S.slides.question) slides.push({ kind: 'mcq', item, els: q.els, notes, size: q.size });
        if (q.overflow) warnings.push(`${tag(item)}: too much text to fit even at ${S.minFontSize} pt — check the slide.`);
        if (S.slides.answer && item.answerIndex >= 0) {
          const a = layoutMCQ(S, M, item, true, q.mode);
          slides.push({ kind: 'mcq-answer', item, els: a.els, notes, size: a.size });
        }
      } else {
        const p = layoutProblem(S, M, item);
        if (S.slides.question) slides.push({ kind: 'problem', item, els: p.els, notes, size: p.size });
        if (p.overflow) warnings.push(`${tag(item)}: question too long to fit — check the slide.`);
      }
      if (S.slides.solution && (item.solution.length || (item.type === 'problem' && item.answerRaw))) {
        const pages = layoutSolution(S, M, item);
        pages.forEach((pg) => {
          slides.push({ kind: 'solution', item, els: pg.els, notes, size: pg.size });
          if (pg.overflow) warnings.push(`${tag(item)}: a solution paragraph is taller than one slide — split it with line breaks.`);
        });
      }
    }
    return { slides, warnings, settings: S, theme: THEMES[S.theme] || THEMES.dark };
  }

  return { DEFAULTS, THEMES, mergeSettings, normalizeRows, buildDeck, heuristicMeasurer, answerIndex, mapHeaders };
});
