/*
 * omml.js — LaTeX → native PowerPoint equations (Office Math Markup, the format of Insert ▸ Equation),
 * plus a matching HTML rendering for the browser preview.
 *
 * In PowerPoint each formula becomes a real equation object: stacked fractions, vector arrows,
 * radicals and scripts that can be edited with the Equation tab.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./latex.js'));
  else root.QBOmml = factory(root.QBLatex);
})(typeof self !== 'undefined' ? self : this, function (QBLatex) {
  'use strict';

  const NS_M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
  const MATH_FONT = 'Cambria Math';
  const SPACE_MAP = { '  ': '\u2003', '    ': '\u2003\u2003', ' ': '\u2005', ' ': '\u2009' };

  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const isBn = (ch) => { const c = ch.codePointAt(0); return (c >= 0x0980 && c <= 0x09FF) || c === 0x0964 || c === 0x0965; };

  function charKind(ch, ctx) {
    if (isBn(ch)) return 'bn';
    if (/[A-Za-z]/.test(ch) && !ctx.upright && !ctx.text && !ctx.bold) return ctx.italicForce === false ? 'op' : 'var';
    return 'op';
  }
  function normChar(ch, ctx) {
    if (ctx.text) return ch;
    if (ch === '-') return '−';
    if (ch === '*') return '∗';
    if (ch === "'") return '′';
    return ch;
  }

  /** Walk the LaTeX AST and hand pieces to a target (OMML or HTML). */
  function walk(nodes, ctx, T) {
    let out = '';
    let buf = null;
    const flush = () => { if (buf) { out += T.text(buf.text, buf.ctx, buf.kind); buf = null; } };
    const add = (text, c, kind) => {
      const key = [kind, !!c.upright, !!c.text, !!c.bold].join('|');
      if (buf && buf.key === key) buf.text += text;
      else { flush(); buf = { text, ctx: c, kind, key }; }
    };
    const plainOf = (arg) => { const r = []; (arg || []).forEach(n => collect(n, r)); return r.join('').trim(); };
    function collect(n, r) {
      if (!n) return;
      if (n.type === 'char' || n.type === 'sym') r.push(n.ch);
      else if (n.type === 'group') n.children.forEach(c => collect(c, r));
      else if (n.type === 'style') n.arg.forEach(c => collect(c, r));
      else r.push('\u0000');
    }

    for (const n of nodes) {
      switch (n.type) {
        case 'empty': case 'newline': break;
        case 'space': add(SPACE_MAP[n.ch] || '\u2005', Object.assign({}, ctx, { upright: true }), 'op'); break;
        case 'char': {
          if (/\s/.test(n.ch)) { if (ctx.text) add(' ', ctx, 'op'); break; }
          const ch = normChar(n.ch, ctx);
          add(ch, ctx, charKind(ch, ctx));
          break;
        }
        case 'sym': add(n.ch, ctx, 'op'); break;
        case 'func': flush(); out += T.text(n.name, Object.assign({}, ctx, { upright: true }), 'op'); break;
        case 'group': flush(); out += walk(n.children, ctx, T); break;
        case 'style': {
          flush();
          const c2 = Object.assign({}, ctx);
          if (n.upright) c2.upright = true;
          if (n.textMode) c2.text = true;
          if (n.bold) c2.bold = true;
          if (n.italic) { c2.italicForce = true; c2.upright = false; }
          out += walk(n.arg, c2, T);
          break;
        }
        case 'frac': flush(); out += T.frac(walk(n.num, ctx, T), walk(n.den, ctx, T), ctx.inScript); break;
        case 'sqrt': flush(); out += T.sqrt(walk(n.arg, ctx, T), n.index && n.index.length ? walk(n.index, Object.assign({}, ctx, { inScript: true }), T) : ''); break;
        case 'accent': flush(); out += T.accent(walk(n.arg, ctx, T), n.mark); break;
        case 'scripts': {
          flush();
          const base = n.base && n.base.type !== 'empty' ? walk([n.base], ctx, T) : '';
          const sc = Object.assign({}, ctx, { inScript: true });
          const supPlain = n.sup ? plainOf(n.sup) : '';
          if (n.sup && (supPlain === '∘' || supPlain === '°')) {
            out += n.sub ? T.sub(base, walk(n.sub, sc, T)) : base;
            out += T.text('°', Object.assign({}, ctx, { upright: true }), 'op');
          } else if (n.sup && n.sub) out += T.subsup(base, walk(n.sub, sc, T), walk(n.sup, sc, T));
          else if (n.sup) out += T.sup(base, walk(n.sup, sc, T));
          else if (n.sub) out += T.sub(base, walk(n.sub, sc, T));
          else out += base;
          break;
        }
        default: break;
      }
    }
    flush();
    return out;
  }

  // ------------------------------------------------------------------ OMML
  function ommlTarget(o) {
    const rPr = (font, italic, lang) =>
      `<a:rPr lang="${lang}" sz="${Math.round(o.sizePt * 100)}"${o.bold ? ' b="1"' : ''}${italic ? ' i="1"' : ' i="0"'} dirty="0">` +
      `<a:solidFill><a:srgbClr val="${o.color}"/></a:solidFill>` +
      `<a:latin typeface="${esc(font)}" charset="0"/><a:ea typeface="${esc(font)}" charset="0"/><a:cs typeface="${esc(font)}" charset="0"/></a:rPr>`;
    return {
      text(text, ctx, kind) {
        let mr = '';
        if (ctx.text || kind === 'bn') mr = '<m:rPr><m:nor/></m:rPr>';
        else if (ctx.bold || o.bold) mr = `<m:rPr><m:sty m:val="${kind === 'var' ? 'bi' : 'b'}"/></m:rPr>`;
        else if (kind !== 'var' && /[A-Za-z]/.test(text)) mr = '<m:rPr><m:sty m:val="p"/></m:rPr>';
        const font = kind === 'bn' ? o.fonts.bn : MATH_FONT;
        return `<m:r>${mr}${rPr(font, kind === 'var', kind === 'bn' ? 'bn-BD' : 'en-US')}<m:t xml:space="preserve">${esc(text)}</m:t></m:r>`;
      },
      frac: (num, den, lin) => `<m:f>${lin ? '<m:fPr><m:type m:val="lin"/></m:fPr>' : ''}<m:num>${num}</m:num><m:den>${den}</m:den></m:f>`,
      sqrt: (arg, idx) => idx
        ? `<m:rad><m:deg>${idx}</m:deg><m:e>${arg}</m:e></m:rad>`
        : `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${arg}</m:e></m:rad>`,
      accent: (arg, mark) => `<m:acc><m:accPr><m:chr m:val="${esc(mark)}"/></m:accPr><m:e>${arg}</m:e></m:acc>`,
      sup: (b, s) => `<m:sSup><m:e>${b}</m:e><m:sup>${s}</m:sup></m:sSup>`,
      sub: (b, s) => `<m:sSub><m:e>${b}</m:e><m:sub>${s}</m:sub></m:sSub>`,
      subsup: (b, s, p) => `<m:sSubSup><m:e>${b}</m:e><m:sub>${s}</m:sub><m:sup>${p}</m:sup></m:sSubSup>`
    };
  }

  /** One inline PowerPoint equation (goes inside an <a:p>). */
  function toOMML(src, o) {
    const body = walk(QBLatex.parseLatex(src), {}, ommlTarget(o));
    return `<a14:m><m:oMath xmlns:m="${NS_M}">${body}</m:oMath></a14:m>`;
  }

  // ------------------------------------------------------------------ HTML (preview)
  function htmlTarget(o) {
    const mathFam = `'Cambria Math','STIX Two Math','Latin Modern Math','Times New Roman',serif`;
    return {
      text(text, ctx, kind) {
        const fam = kind === 'bn' ? `'${o.fonts.bn}',sans-serif` : mathFam;
        const st = `font-family:${fam};font-style:${kind === 'var' ? 'italic' : 'normal'};${ctx.bold || o.bold ? 'font-weight:700;' : ''}`;
        let t = esc(text).replace(/ /g, '&nbsp;');
        // PowerPoint spaces operators automatically; mimic that in the preview
        if (!ctx.inScript && !ctx.text) t = t.replace(/[=+−×÷±≈≠≤≥→←⇒⟹⟶∝·]|&lt;|&gt;/g, '<span style="margin:0 .24em">$&</span>');
        return `<span style="${st}">${t}</span>`;
      },
      frac: (num, den, lin) => lin
        ? `${num}/${den}`
        : `<span style="display:inline-flex;flex-direction:column;align-items:center;vertical-align:middle;line-height:1.15;margin:0 .1em">` +
          `<span style="padding:0 .1em">${num}</span><span style="border-top:.06em solid currentColor;padding:0 .1em;align-self:stretch;text-align:center">${den}</span></span>`,
      sqrt: (arg, idx) => `${idx ? `<sup style="font-size:.6em">${idx}</sup>` : ''}<span style="font-family:${mathFam}">√</span><span style="border-top:.06em solid currentColor">${arg}</span>`,
      accent: (arg, mark) => {
        const sym = mark === '\u20D7' ? '→' : mark === '\u0302' ? '^' : mark === '\u0303' ? '~' : mark === '\u0307' ? '·' : mark === '\u0308' ? '¨' : '‾';
        return `<span style="position:relative;display:inline-block">${arg}<span style="position:absolute;left:0;right:0;top:-.62em;text-align:center;font-size:.75em;font-style:normal;font-family:${mathFam}">${sym}</span></span>`;
      },
      sup: (b, s) => `${b}<sup style="font-size:.66em;line-height:0">${s}</sup>`,
      sub: (b, s) => `${b}<sub style="font-size:.66em;line-height:0">${s}</sub>`,
      subsup: (b, s, p) => `${b}<span style="display:inline-flex;flex-direction:column;vertical-align:middle;font-size:.66em;line-height:1"><span>${p}</span><span>${s}</span></span>`
    };
  }

  function toHTML(src, o) {
    return walk(QBLatex.parseLatex(src), {}, htmlTarget(o));
  }

  return { toOMML, toHTML, NS_M, MATH_FONT };
});
