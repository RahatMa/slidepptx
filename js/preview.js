/*
 * preview.js — draws slide models as HTML so you can check them before downloading.
 * Uses the same coordinates as the .pptx, scaled to the preview width.
 */
(function (root) {
  'use strict';
  const L = root.QBLatex;

  function esc(s) {
    return String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  }

  function runsHtml(paragraphs, S, k, el) {
    const fs = el.size / 72 * k;
    const lh = fs * S.lineFactor * (el.spacing || S.lineSpacing);
    const paras = paragraphs && paragraphs.length ? paragraphs : [[{ text: ' ' }]];
    return paras.map((runs) => {
      const list = runs.length ? runs : [{ text: '\u00A0' }];
      const parts = [];
      for (let i = 0; i < list.length;) {
        const r = list[i];
        if (r.eqSrc && S.nativeEquations !== false && root.QBOmml) {
          const id = r.eqId;
          while (i < list.length && list[i].eqId === id) i++;
          parts.push(`<span style="white-space:nowrap">${root.QBOmml.toHTML(r.eqSrc, { fonts: S.fonts, bold: !!(el.bold || r.bold) })}</span>`);
          continue;
        }
        i++;
        parts.push(L.scriptSegments(r).map((seg) => {
          const font = S.fonts[seg.kind] || S.fonts.lat;
          let st = `font-family:'${font}','Hind Siliguri','Noto Sans Bengali',sans-serif;`;
          if (r.italic) st += 'font-style:italic;';
          if (r.bold || el.bold) st += 'font-weight:700;';
          const t = esc(seg.text).replace(/ {2}/g, ' &nbsp;');
          if (r.sup) return `<sup style="${st}font-size:.62em;line-height:0">${t}</sup>`;
          if (r.sub) return `<sub style="${st}font-size:.62em;line-height:0">${t}</sub>`;
          return `<span style="${st}">${t}</span>`;
        }).join(''));
      }
      const inner = parts.join('') || '&nbsp;';
      return `<p style="margin:0;font-size:${fs}px;line-height:${lh}px;text-align:${el.align || 'left'}">${inner}</p>`;
    }).join('');
  }

  function renderSlide(slide, deck, widthPx, overlay) {
    const S = deck.settings, T = deck.theme;
    const k = widthPx / S.slideW;
    const wrap = document.createElement('div');
    wrap.className = 'slide';
    wrap.style.cssText = `position:relative;width:${widthPx}px;height:${S.slideH * k}px;background:#${T.bg};overflow:hidden;color:#${T.text}`;
    if (overlay) {
      const o = document.createElement('img');
      o.src = overlay;
      o.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;pointer-events:none';
      wrap.appendChild(o);
    }
    for (const el of slide.els) {
      if (el.kind === 'image') {
        const im = document.createElement('img');
        im.src = el.data;
        im.alt = '';
        im.style.cssText = `position:absolute;left:${el.x * k}px;top:${el.y * k}px;width:${el.w * k}px;height:${el.h * k}px`;
        wrap.appendChild(im);
        continue;
      }
      const d = document.createElement('div');
      const inset = el.inset || [0, 0, 0, 0];
      const jc = el.valign === 'middle' ? 'center' : el.valign === 'bottom' ? 'flex-end' : 'flex-start';
      d.style.cssText = `position:absolute;box-sizing:border-box;left:${el.x * k}px;top:${el.y * k}px;width:${el.w * k}px;height:${el.h * k}px;` +
        `padding:${inset[0] * k}px ${inset[1] * k}px ${inset[2] * k}px ${inset[3] * k}px;display:flex;flex-direction:column;justify-content:${jc};` +
        `color:#${el.color};overflow:visible;white-space:normal;word-wrap:break-word;` +
        (el.shape ? `background:#${el.fill};border-radius:${(el.radius || 0.25) * k}px;` : '');
      d.innerHTML = runsHtml(el.paragraphs, S, k, el);
      wrap.appendChild(d);
    }
    return wrap;
  }

  root.QBPreview = { renderSlide };
})(typeof self !== 'undefined' ? self : this);
