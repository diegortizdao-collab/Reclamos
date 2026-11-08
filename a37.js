/* Genera el A.37 "Informe de reclamo" a partir de la plantilla en blanco (A37.docx) y los datos del reclamo.
   Funciona en el navegador (JSZip + DOMParser) y en Node (con @xmldom/xmldom). */
(function (g) {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const dmy = (iso) => (iso ? String(iso).slice(8, 10) + '/' + String(iso).slice(5, 7) + '/' + String(iso).slice(2, 4) : '');
  const dmyFull = (iso) => (iso ? String(iso).slice(8, 10) + '/' + String(iso).slice(5, 7) + '/' + String(iso).slice(0, 4) : '');
  const kids = (n, name) => [...n.childNodes].filter((c) => c.nodeType === 1 && c.localName === name);

  function setCell(doc, tc, text) {
    const ps = kids(tc, 'p');
    const first = ps[0];
    ps.slice(1).forEach((p) => tc.removeChild(p));
    [...first.childNodes].forEach((c) => { if (c.localName !== 'pPr') first.removeChild(c); });
    const lines = String(text == null ? '' : text).split(/\r?\n/);
    let p = first;
    lines.forEach((ln, i) => {
      if (i > 0) { p = first.cloneNode(false); [...first.childNodes].forEach((c) => c.localName === 'pPr' && p.appendChild(c.cloneNode(true))); tc.appendChild(p); }
      if (ln === '') return;
      const r = doc.createElementNS(W, 'w:r');
      const rPr = doc.createElementNS(W, 'w:rPr');
      ['sz', 'szCs'].forEach((t) => { const e = doc.createElementNS(W, 'w:' + t); e.setAttributeNS(W, 'w:val', '22'); rPr.appendChild(e); });
      r.appendChild(rPr);
      const t = doc.createElementNS(W, 'w:t'); t.setAttribute('xml:space', 'preserve'); t.appendChild(doc.createTextNode(ln)); r.appendChild(t);
      p.appendChild(r);
    });
  }
  const rowsOf = (tbl) => kids(tbl, 'tr');
  const cellsOf = (tr) => kids(tr, 'tc');
  const si = (v) => (v === true ? true : v === false ? false : null);

  // pone el óvalo sobre "Sí" (228.2pt) o "No" (+27.8pt); sin dato, lo deja sobre "Sí" quitando nada
  function ovalo(doc, idx, valor) {
    const o = doc.getElementsByTagName('v:oval')[idx];
    if (!o) return;
    let s = o.getAttribute('style') || '';
    const left = valor === false ? '253pt' : '228.2pt';
    s = s.replace(/margin-left:[^;]+/, 'margin-left:' + left);
    o.setAttribute('style', s);
    if (valor == null) o.setAttribute('style', s.replace('margin-left:' + left, 'margin-left:-9999pt'));
  }

  async function llenarA37(plantillaBytes, r, JSZip, DP, XS) {
    const a = r.a37 || {};
    const zip = await JSZip.loadAsync(plantillaBytes);
    const xml = await zip.file('word/document.xml').async('string');
    const doc = new DP().parseFromString(xml, 'application/xml');
    const tbls = [...doc.getElementsByTagNameNS(W, 'tbl')];
    const T = tbls.map(rowsOf);

    // T0 datos generales
    const dg = [r.numero, r.cliente, r.producto, r.cantidad, r.oti, dmy(r.fecha)];
    dg.forEach((v, i) => setCell(doc, cellsOf(T[0][i + 1])[1], v == null ? '' : v));
    // T1 motivo
    setCell(doc, cellsOf(T[1][1])[0], r.motivo || '');
    // T2 NC asociada
    const ncs = r.nc_numeros || [];
    ovalo(doc, 0, ncs.length ? true : false);
    setCell(doc, cellsOf(T[2][0])[3], ncs.join(', '));
    // T3 acciones de Silver, T4 del cliente
    [['nota_debito', 2], ['nota_credito', 3], ['reposicion', 4]].forEach(([k, i]) => {
      const c = cellsOf(T[3][i]); setCell(doc, c[1], a[k + '_nro'] || ''); setCell(doc, c[2], dmyFull(a[k + '_fecha']) || '');
    });
    { const c = cellsOf(T[4][2]); setCell(doc, c[1], a.nd_cliente_nro || ''); setCell(doc, c[2], dmyFull(a.nd_cliente_fecha) || ''); }
    // T5 factura
    const tieneFac = !!a.factura;
    ovalo(doc, 1, tieneFac);
    setCell(doc, cellsOf(T[5][0])[3], a.factura || '');
    // T6 novedades
    const nov = (a.novedades && a.novedades.length ? a.novedades : r.accion_inmediata ? [{ texto: r.accion_inmediata, fecha: r.fecha }] : []);
    const filaBase = T[6][T[6].length - 1];
    nov.forEach((n, i) => {
      let tr = filaBase;
      if (i > 0) { tr = filaBase.cloneNode(true); filaBase.parentNode.appendChild(tr); }
      const c = cellsOf(tr);
      setCell(doc, c[0], n.texto || ''); setCell(doc, c[1], dmyFull(n.fecha) || '');
      if (i > 0) { const h = kids(tr, 'trPr')[0]; if (h) tr.removeChild(h); }
    });
    if (!nov.length) { const c = cellsOf(filaBase); setCell(doc, c[0], ''); setCell(doc, c[1], ''); }

    zip.file('word/document.xml', new XS().serializeToString(doc));
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }
  g.llenarA37 = llenarA37;
  if (typeof module !== 'undefined') module.exports = { llenarA37 };
})(typeof window !== 'undefined' ? window : globalThis);
