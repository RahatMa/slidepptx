/*
 * latex.js — turns cell text with inline LaTeX into editable rich-text runs.
 *
 * Input : "নিচের কোনটি \(mv^2/r\)?"   (math in \( \), \[ \], $ $ or $$ $$)
 * Output: paragraphs -> runs  [[{text, italic, bold, sup, sub, math, symbol}]]
 *
 * Nothing is rendered as an image: every formula ends up as normal PowerPoint
 * text with real superscript/subscript formatting, so it stays editable.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.QBLatex = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const GREEK = {
    alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', varepsilon: 'ε', zeta: 'ζ', eta: 'η',
    theta: 'θ', vartheta: 'ϑ', iota: 'ι', kappa: 'κ', lambda: 'λ', mu: 'μ', nu: 'ν', xi: 'ξ', pi: 'π',
    varpi: 'ϖ', rho: 'ρ', varrho: 'ϱ', sigma: 'σ', varsigma: 'ς', tau: 'τ', upsilon: 'υ', phi: 'ϕ',
    varphi: 'φ', chi: 'χ', psi: 'ψ', omega: 'ω', Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ',
    Xi: 'Ξ', Pi: 'Π', Sigma: 'Σ', Upsilon: 'Υ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω'
  };

  const SYMBOLS = {
    times: '×', cdot: '·', div: '÷', pm: '±', mp: '∓', le: '≤', leq: '≤', ge: '≥', geq: '≥', ne: '≠',
    neq: '≠', approx: '≈', sim: '∼', simeq: '≃', cong: '≅', equiv: '≡', propto: '∝', infty: '∞',
    to: '→', rightarrow: '→', leftarrow: '←', gets: '←', Rightarrow: '⇒', Leftarrow: '⇐',
    Leftrightarrow: '⇔', leftrightarrow: '↔', longrightarrow: '⟶', Longrightarrow: '⟹',
    implies: '⟹', iff: '⟺', uparrow: '↑', downarrow: '↓', rightleftharpoons: '⇌', degree: '°',
    circ: '∘', angle: '∠', perp: '⊥', parallel: '∥', partial: '∂', nabla: '∇', sum: '∑', prod: '∏',
    int: '∫', iint: '∬', oint: '∮', therefore: '∴', because: '∵', ldots: '…', cdots: '⋯', dots: '…',
    vdots: '⋮', prime: '′', hbar: 'ℏ', ell: 'ℓ', in: '∈', notin: '∉', subset: '⊂', supset: '⊃',
    cup: '∪', cap: '∩', forall: '∀', exists: '∃', neg: '¬', land: '∧', lor: '∨', gg: '≫', ll: '≪',
    star: '⋆', ast: '∗', bullet: '•', checkmark: '✓', triangle: '△', square: '□', AA: 'Å',
    lbrace: '{', rbrace: '}', langle: '⟨', rangle: '⟩', vert: '|', Vert: '‖', mid: '|', lvert: '|',
    rvert: '|', lVert: '‖', rVert: '‖', percent: '%', colon: ':', backslash: '\\', emptyset: '∅',
    odot: '⊙', otimes: '⊗', oplus: '⊕', dagger: '†', S: '§'
  };

  const SPACES = { quad: '  ', qquad: '    ', ',': ' ', ';': ' ', ':': ' ', '>': ' ', '!': '', ' ': ' ', enspace: ' ', thinspace: ' ', medspace: ' ', thickspace: ' ' };

  const FUNCS = new Set(['sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'cosec', 'arcsin', 'arccos', 'arctan',
    'sinh', 'cosh', 'tanh', 'log', 'ln', 'lg', 'exp', 'lim', 'max', 'min', 'sup', 'inf', 'det', 'deg',
    'dim', 'mod', 'gcd', 'arg']);

  const TEXT_CMDS = new Set(['text', 'textrm', 'mbox', 'textup', 'textnormal', 'hbox']);
  const UPRIGHT_CMDS = new Set(['mathrm', 'operatorname', 'mathsf', 'mathup', 'rm', 'mathtt', 'texttt', 'mathbb', 'mathcal', 'mathscr', 'mathfrak']);
  const BOLD_CMDS = new Set(['mathbf', 'textbf', 'boldsymbol', 'bm', 'bf', 'pmb']);
  const ITALIC_CMDS = new Set(['mathit', 'textit', 'emph', 'it']);
  const ACCENTS = {
    vec: '\u20D7', overrightarrow: '\u20D7', hat: '\u0302', widehat: '\u0302', bar: '\u0305',
    overline: '\u0305', dot: '\u0307', ddot: '\u0308', tilde: '\u0303', widetilde: '\u0303', underline: '\u0332'
  };
  const SKIP = new Set(['left', 'right', 'big', 'Big', 'bigg', 'Bigg', 'bigl', 'bigr', 'Bigl', 'Bigr',
    'biggl', 'biggr', 'displaystyle', 'textstyle', 'scriptstyle', 'limits', 'nolimits', 'middle',
    'nonumber', 'notag', 'label', 'tag', 'hfill', 'centering', 'noindent']);

  // Relations get a space on both sides; binary operators only when they sit between operands.
  const RELATIONS = new Set(['=', '<', '>', '≤', '≥', '≈', '≠', '→', '←', '⇒', '⇐', '⇔', '↔', '⟶', '⟹', '⟺', '∝', '≡', '≃', '≅', '∼', '⇌', '≫', '≪', '∴', '∵', '∈', '∉', '⊂', '⊃']);
  const BINARY = new Set(['+', '−', '×', '÷', '±', '∓', '·', '∗', '⊗', '⊕']);

  // Characters many text fonts lack; the writer sends them to the symbol font (Cambria Math).
  // Characters that Kalpurush / Hind Siliguri lack (Greek, arrows, most operators);
  // the writer sends these to the symbol font (Cambria Math by default).
  const TEXT_FONT_SAFE = new Set(Array.from('√≈≠≤≥−∞∓'));
  function isSymbolChar(ch) {
    if (TEXT_FONT_SAFE.has(ch)) return false;
    const c = ch.codePointAt(0);
    return (c >= 0x0370 && c <= 0x03FF) || (c >= 0x2032 && c <= 0x2037) ||
      (c >= 0x2190 && c <= 0x22FF) || (c >= 0x27C0 && c <= 0x27FF) || (c >= 0x2900 && c <= 0x2AFF) ||
      (c >= 0x20D0 && c <= 0x20FF) || (c >= 0x2100 && c <= 0x214F && c !== 0x2116);
  }

  // ---------- tokenizer / parser ----------
  function tokenize(s) {
    const t = [];
    let i = 0;
    while (i < s.length) {
      const c = s[i];
      if (c === '\\') {
        const n = s[i + 1] || '';
        if (/[a-zA-Z]/.test(n)) {
          let j = i + 1;
          while (j < s.length && /[a-zA-Z]/.test(s[j])) j++;
          t.push({ type: 'cmd', name: s.slice(i + 1, j) });
          i = j;
        } else { t.push({ type: 'cmd', name: n }); i += 2; }
      } else if (c === '{') { t.push({ type: 'open' }); i++; }
      else if (c === '}') { t.push({ type: 'close' }); i++; }
      else if (c === '^' || c === '_') { t.push({ type: c }); i++; }
      else if (c === '&') { i++; } // alignment marks
      else {
        const ch = String.fromCodePoint(s.codePointAt(i));
        t.push({ type: 'char', ch });
        i += ch.length;
      }
    }
    return t;
  }

  function parse(tokens) {
    let pos = 0;
    const EMPTY = { type: 'empty' };
    const isSpace = (tk) => tk && tk.type === 'char' && /\s/.test(tk.ch);

    function parseSeq(stopAtClose) {
      const nodes = [];
      while (pos < tokens.length) {
        const tk = tokens[pos];
        if (tk.type === 'close') { pos++; if (stopAtClose) return nodes; continue; }
        if (tk.type === '^' || tk.type === '_') {
          pos++;
          const arg = parseArg();
          const key = tk.type === '^' ? 'sup' : 'sub';
          let base = nodes.pop();
          while (base && base.type === 'char' && /\s/.test(base.ch)) base = nodes.pop();
          if (base && base.type === 'scripts' && !base[key]) { base[key] = arg; nodes.push(base); }
          else nodes.push({ type: 'scripts', base: base || EMPTY, sup: key === 'sup' ? arg : null, sub: key === 'sub' ? arg : null });
          continue;
        }
        nodes.push(parseAtom());
      }
      return nodes;
    }
    function parseArg() {
      while (isSpace(tokens[pos])) pos++;
      if (pos >= tokens.length) return [];
      return [parseAtom()];
    }
    function parseOptional() { // [ ... ]
      while (isSpace(tokens[pos])) pos++;
      if (!(tokens[pos] && tokens[pos].type === 'char' && tokens[pos].ch === '[')) return null;
      pos++;
      const out = [];
      while (pos < tokens.length && !(tokens[pos].type === 'char' && tokens[pos].ch === ']')) out.push(parseAtom());
      pos++;
      return out;
    }
    function parseAtom() {
      const tk = tokens[pos++];
      if (!tk) return EMPTY;
      if (tk.type === 'open') return { type: 'group', children: parseSeq(true) };
      if (tk.type === 'char') return { type: 'char', ch: tk.ch };
      if (tk.type === 'cmd') return parseCmd(tk.name);
      return EMPTY;
    }
    function parseCmd(name) {
      if (name === 'frac' || name === 'dfrac' || name === 'tfrac' || name === 'cfrac') return { type: 'frac', num: parseArg(), den: parseArg() };
      if (name === 'sqrt') { const index = parseOptional(); return { type: 'sqrt', index, arg: parseArg() }; }
      if (ACCENTS[name]) return { type: 'accent', mark: ACCENTS[name], arg: parseArg() };
      if (TEXT_CMDS.has(name)) return { type: 'style', upright: true, textMode: true, arg: parseArg() };
      if (UPRIGHT_CMDS.has(name)) return { type: 'style', upright: true, arg: parseArg() };
      if (BOLD_CMDS.has(name)) return { type: 'style', bold: true, arg: parseArg() };
      if (ITALIC_CMDS.has(name)) return { type: 'style', italic: true, arg: parseArg() };
      if (name === 'left' || name === 'right' || name === 'middle') {
        while (isSpace(tokens[pos])) pos++;
        const nx = tokens[pos];
        if (nx && nx.type === 'char' && nx.ch === '.') { pos++; return EMPTY; }
        return EMPTY;
      }
      if (name === 'begin' || name === 'end') { parseArg(); if (name === 'begin') parseOptional(); return EMPTY; }
      if (name === 'label' || name === 'tag') { parseArg(); return EMPTY; }
      if (SKIP.has(name)) return EMPTY;
      if (name === 'not') { const a = parseArg(); return { type: 'accent', mark: '\u0338', arg: a }; }
      if (name === 'pmod') return { type: 'group', children: [{ type: 'char', ch: '(' }, { type: 'func', name: 'mod' }, ...parseArg(), { type: 'char', ch: ')' }] };
      if (GREEK[name]) return { type: 'sym', ch: GREEK[name], greek: true };
      if (SYMBOLS[name]) return { type: 'sym', ch: SYMBOLS[name] };
      if (Object.prototype.hasOwnProperty.call(SPACES, name)) return { type: 'space', ch: SPACES[name] };
      if (FUNCS.has(name)) return { type: 'func', name: name === 'cosec' ? 'cosec' : name };
      if (name === '\\') return { type: 'newline' };
      if ('{}%$&#_|'.includes(name)) return { type: 'char', ch: name, literal: true };
      return { type: 'func', name }; // unknown command: print its name upright
    }
    return parseSeq(false);
  }

  // ---------- renderer ----------
  function lastChar(out) {
    for (let i = out.length - 1; i >= 0; i--) {
      const t = out[i].text.replace(/\s+$/, '');
      if (t) return { ch: t[t.length - 1], run: out[i] };
    }
    return null;
  }

  function push(out, text, ctx, extra) {
    if (!text) return;
    const run = {
      text,
      italic: !!ctx.italicForce,
      bold: !!ctx.bold,
      sup: !!ctx.sup,
      sub: !!ctx.sub,
      math: true
    };
    if (extra) Object.assign(run, extra);
    out.push(run);
  }

  function plain(runs) { return runs.map(r => r.text).join(''); }
  // "simple" = a single term (mv², r, 10, θ) that needs no brackets around it in a/b form
  function isSimple(runs) {
    if (!runs.length) return false;
    if (runs.some(r => /\s/.test(r.text))) return false;
    const base = plain(runs.filter(r => !r.sup && !r.sub));
    return /^[\p{L}\p{N}.′°\u0300-\u036F\u20D0-\u20FF]+$/u.test(base);
  }

  function renderNodes(nodes, ctx, out) { for (const n of nodes) renderNode(n, ctx, out); }

  function renderNode(n, ctx, out) {
    const inScript = ctx.sup || ctx.sub;
    switch (n.type) {
      case 'empty': return;
      case 'newline': out.push({ newline: true, text: '' }); return;
      case 'space': push(out, n.ch, ctx); return;
      case 'char': {
        let ch = n.ch;
        if (/\s/.test(ch)) { if (ctx.textMode) push(out, ' ', ctx); return; }
        if (ctx.textMode || n.literal) { push(out, ch, ctx, { italic: !!ctx.italicForce }); return; }
        if (ch === '-') ch = '−';
        if (ch === '*') ch = '∗';
        if (ch === "'") { push(out, '′', ctx); return; }
        if (RELATIONS.has(ch)) { push(out, inScript ? ch : ' ' + ch + ' ', ctx); return; }
        if (BINARY.has(ch)) {
          const prev = lastChar(out);
          const binary = prev && /[\p{L}\p{N})\]}′°!²³]/u.test(prev.ch);
          push(out, binary && !inScript ? ' ' + ch + ' ' : ch, ctx);
          return;
        }
        if (/[A-Za-z]/.test(ch)) {
          const italic = ctx.italicForce || (!ctx.upright && !ctx.bold);
          push(out, ch, ctx, { italic });
          return;
        }
        push(out, ch, ctx, { symbol: isSymbolChar(ch) });
        return;
      }
      case 'sym': {
        const ch = n.ch;
        if (RELATIONS.has(ch)) { push(out, inScript ? ch : ' ' + ch + ' ', ctx, { symbol: isSymbolChar(ch) }); return; }
        if (BINARY.has(ch)) {
          const prev = lastChar(out);
          const binary = prev && /[\p{L}\p{N})\]}′°!²³]/u.test(prev.ch);
          push(out, binary && !inScript ? ' ' + ch + ' ' : ch, ctx, { symbol: isSymbolChar(ch) });
          return;
        }
        push(out, ch, ctx, { symbol: isSymbolChar(ch), italic: n.greek && !ctx.upright ? false : !!ctx.italicForce });
        return;
      }
      case 'func': push(out, n.name + (inScript ? '' : ' '), ctx, { italic: false }); return;
      case 'group': renderNodes(n.children, ctx, out); return;
      case 'style': {
        const c2 = Object.assign({}, ctx);
        if (n.upright) { c2.upright = true; c2.italicForce = false; }
        if (n.textMode) c2.textMode = true;
        if (n.bold) { c2.bold = true; c2.upright = true; }
        if (n.italic) c2.italicForce = true;
        renderNodes(n.arg, c2, out);
        return;
      }
      case 'frac': {
        const a = [], b = [];
        renderNodes(n.num, ctx, a);
        renderNodes(n.den, ctx, b);
        const wrap = (rs) => (isSimple(rs) ? rs : [mk('(', ctx), ...rs, mk(')', ctx)]);
        out.push(...wrap(a));
        push(out, '/', ctx);
        out.push(...wrap(b));
        return;
      }
      case 'sqrt': {
        if (n.index && n.index.length) renderNodes(n.index, Object.assign({}, ctx, { sup: true, sub: false }), out);
        push(out, '√', ctx);
        const a = [];
        renderNodes(n.arg, ctx, a);
        const t = plain(a);
        if (isSimple(a) && (Array.from(t).length === 1 || /^[\d.]+$/.test(t))) out.push(...a);
        else out.push(mk('(', ctx), ...a, mk(')', ctx));
        return;
      }
      case 'accent': {
        const a = [];
        renderNodes(n.arg, ctx, a);
        // put the last character together with its combining mark in a symbol-font run
        for (let i = a.length - 1; i >= 0; i--) {
          const t = a[i].text;
          if (!t.trim()) continue;
          const chars = Array.from(t.replace(/\s+$/, ''));
          const last = chars.pop();
          const before = chars.join('');
          const runs = [];
          if (before) runs.push(Object.assign({}, a[i], { text: before }));
          runs.push(Object.assign({}, a[i], { text: last + n.mark, symbol: true }));
          a.splice(i, 1, ...runs);
          break;
        }
        out.push(...a);
        return;
      }
      case 'scripts': {
        renderNode(n.base, ctx, out);
        const scriptText = (arg) => {
          const tmp = [];
          renderNodes(arg, Object.assign({}, ctx, { sup: false, sub: false }), tmp);
          return plain(tmp).trim();
        };
        if (n.sub) renderNodes(n.sub, Object.assign({}, ctx, { sub: true, sup: false }), out);
        if (n.sup) {
          const st = scriptText(n.sup);
          if (st === '∘' || st === '°') push(out, '°', ctx, { italic: false });
          else if (st === '′' || st === '′′') push(out, st, ctx, { italic: false });
          else renderNodes(n.sup, Object.assign({}, ctx, { sup: true, sub: false }), out);
        }
        return;
      }
      default: return;
    }
  }

  function mk(text, ctx) { return { text, italic: false, bold: !!ctx.bold, sup: !!ctx.sup, sub: !!ctx.sub, math: true }; }

  /** Convert one LaTeX math string into an array of lines (each an array of runs). */
  function mathToLines(src) {
    const nodes = parse(tokenize(src));
    const out = [];
    renderNodes(nodes, {}, out);
    const lines = [[]];
    for (const r of out) {
      if (r.newline) { lines.push([]); continue; }
      lines[lines.length - 1].push(r);
    }
    return lines.map(l => tidy(l));
  }

  function sameStyle(a, b) {
    return a.italic === b.italic && a.bold === b.bold && a.sup === b.sup && a.sub === b.sub &&
      !!a.math === !!b.math && !!a.symbol === !!b.symbol && a.eqId === b.eqId;
  }

  function tidy(runs) {
    // merge neighbours with identical style; collapse double spaces created by operator padding
    const merged = [];
    for (const r of runs) {
      if (!r.text) continue;
      const p = merged[merged.length - 1];
      if (p && sameStyle(p, r)) p.text += r.text;
      else merged.push(Object.assign({}, r));
    }
    for (let i = 0; i < merged.length; i++) {
      merged[i].text = merged[i].text.replace(/ {2,}/g, (m) => (merged[i].text.trim() ? ' ' : m));
      if (i > 0 && / $/.test(merged[i - 1].text) && /^ /.test(merged[i].text)) merged[i].text = merged[i].text.replace(/^ +/, '');
    }
    if (merged.length) {
      merged[0].text = merged[0].text.replace(/^ +/, '');
      merged[merged.length - 1].text = merged[merged.length - 1].text.replace(/ +$/, '');
    }
    return merged.filter(r => r.text);
  }

  let eqCounter = 0;
  const MATH_RE = /\\\(([\s\S]*?)\\\)|\\\[([\s\S]*?)\\\]|\$\$([\s\S]*?)\$\$|\$([^$\n]+?)\$/g;

  /**
   * Parse a cell into paragraphs of runs.
   * Line breaks inside the cell (Alt+Enter) and \\ outside math start new paragraphs.
   */
  function parseRich(input) {
    if (input === null || input === undefined) return [];
    let s = String(input).replace(/\r\n?/g, '\n').replace(/\u00A0/g, ' ');
    if (!s.trim()) return [];
    const paras = [[]];
    const cur = () => paras[paras.length - 1];
    const newPara = () => { paras.push([]); };

    let afterDisplay = false;
    const addPlain = (text) => {
      if (afterDisplay) text = text.replace(/^[ \t]*\n/, '');
      afterDisplay = false;
      if (!text) return;
      const lines = text.replace(/\\\\/g, '\n').split('\n');
      lines.forEach((line, idx) => {
        if (idx > 0) newPara();
        const t = line.replace(/\\%/g, '%').replace(/\\&/g, '&').replace(/\\\$/g, '$').replace(/\\_/g, '_');
        if (t) cur().push({ text: t, italic: false, bold: false, sup: false, sub: false, math: false });
      });
    };

    let last = 0;
    let m;
    MATH_RE.lastIndex = 0;
    while ((m = MATH_RE.exec(s)) !== null) {
      addPlain(s.slice(last, m.index));
      const display = m[2] !== undefined || m[3] !== undefined;
      const body = m[1] !== undefined ? m[1] : m[2] !== undefined ? m[2] : m[3] !== undefined ? m[3] : m[4];
      // split aligned/multi-line math into one equation per line; every run remembers its LaTeX source
      const parts = body.replace(/\\begin\{[^}]*\}(\[[^\]]*\])?|\\end\{[^}]*\}/g, '').split(/\\\\/)
        .map(t => t.replace(/&/g, '').trim()).filter(Boolean);
      if (display && cur().length) newPara();
      parts.forEach((part, idx) => {
        if (idx > 0) newPara();
        const eqId = ++eqCounter;
        const frac = /\\[dtc]?frac/.test(part);
        const runs = mathToLines(part).reduce((a, l) => a.concat(l), []);
        runs.forEach(r => { r.eqId = eqId; r.eqSrc = part; if (frac) r.eqFrac = true; });
        cur().push(...runs);
      });
      if (display) { newPara(); afterDisplay = true; }
      last = m.index + m[0].length;
    }
    addPlain(s.slice(last));

    // trim each paragraph, drop leading/trailing empty paragraphs, keep inner blank lines
    const cleaned = paras.map(tidyPara);
    while (cleaned.length && !cleaned[0].length) cleaned.shift();
    while (cleaned.length && !cleaned[cleaned.length - 1].length) cleaned.pop();
    return cleaned;
  }

  function tidyPara(runs) {
    const merged = [];
    for (const r of runs) {
      if (!r.text) continue;
      const p = merged[merged.length - 1];
      if (p && sameStyle(p, r)) p.text += r.text;
      else merged.push(Object.assign({}, r));
    }
    if (merged.length) {
      merged[0].text = merged[0].text.replace(/^\s+/, '');
      merged[merged.length - 1].text = merged[merged.length - 1].text.replace(/\s+$/, '');
    }
    return merged.filter(r => r.text);
  }

  // ---------- plain-text export (speaker notes, previews of long text) ----------
  const SUP_MAP = { '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹', '+': '⁺', '-': '⁻', '−': '⁻', '=': '⁼', '(': '⁽', ')': '⁾', n: 'ⁿ', i: 'ⁱ' };
  const SUB_MAP = { '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄', '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉', '+': '₊', '-': '₋', '−': '₋', '=': '₌', '(': '₍', ')': '₎' };

  function runText(r) {
    if (!r.sup && !r.sub) return r.text;
    const map = r.sup ? SUP_MAP : SUB_MAP;
    const chars = Array.from(r.text);
    if (chars.every(c => map[c])) return chars.map(c => map[c]).join('');
    if (/^[\p{L}\p{N}]+$/u.test(r.text)) return (r.sup ? '^' : '_') + r.text;
    return (r.sup ? '^(' : '_(') + r.text + ')';
  }
  function toPlain(paragraphs) {
    return (paragraphs || []).map(p => p.map(runText).join('')).join('\n');
  }

  /** Split a run into font segments: Bengali text, symbols, everything else. */
  function scriptSegments(run) {
    const segs = [];
    let curKind = null, buf = '';
    const flush = () => { if (buf) segs.push({ text: buf, kind: curKind }); buf = ''; };
    for (const ch of Array.from(run.text)) {
      const c = ch.codePointAt(0);
      let kind;
      if (/\s/.test(ch)) kind = curKind || 'lat';
      else if ((c >= 0x0980 && c <= 0x09FF) || c === 0x0964 || c === 0x0965 || c === 0x200C || c === 0x200D) kind = 'bn';
      else if (run.symbol || isSymbolChar(ch)) kind = 'sym';
      else kind = run.math ? 'math' : 'lat';
      if (kind !== curKind) { flush(); curKind = kind; }
      buf += ch;
    }
    flush();
    return segs;
  }

  return { parseRich, mathToLines, toPlain, scriptSegments, isSymbolChar, parseLatex: (src) => parse(tokenize(src)), _tokenize: tokenize };
});
