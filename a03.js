/* Lee la planilla "A.03 Reclamos de Cliente" (.xlsm/.xlsx, hoja "Planilla") y la convierte en reclamos listos para
   guardar. Todo ocurre en el navegador: el archivo no se sube a ningún lado. Requiere SheetJS (XLSX). */
(function (g) {
  const sinAcentos = (s) => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '');
  const txt = (v) => { if (v == null) return null; const t = String(v).replace(/\r/g, '').trim(); return t === '' ? null : t; };

  function isoDeSerial(n) { return new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000).toISOString().slice(0, 10); }
  function fechaCelda(c) {
    if (!c) return null;
    if (c.t === 'n' && c.v > 20000 && c.v < 80000) return isoDeSerial(c.v);
    if (c.t === 'd' && c.v instanceof Date) return c.v.toISOString().slice(0, 10);
    const m = /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/.exec(String(c.v || '').trim());
    if (m) { const y = +m[3] < 100 ? 2000 + +m[3] : +m[3]; return `${y}-${String(+m[2]).padStart(2, '0')}-${String(+m[1]).padStart(2, '0')}`; }
    return null;
  }
  const esFecha = (c) => c && c.t === 'n' && c.z && /[dmy]/i.test(String(c.z).replace(/\[[^\]]*\]/g, '').replace(/"[^"]*"/g, '')) && c.v > 20000 && c.v < 80000;
  const dmy = (iso) => iso.slice(8, 10) + '/' + iso.slice(5, 7) + '/' + iso.slice(0, 4);

  function tri(v) {
    const t = sinAcentos(txt(v) || '').toLowerCase().replace(/[^a-z]/g, '');
    if (!t) return null;
    if (t.startsWith('s')) return true;
    if (t.startsWith('n')) return false;
    return null;
  }
  function flag(v) { const n = Number(String(v).trim()); return n === 1 ? true : n === 0 && String(v).trim() !== '' ? false : null; }

  // Números de NC mencionados en texto libre ("Se genero NC 918", "Ver NC 913", "NC Nº49 y 50", "No conformidad 917")
  function ncDeTexto(t) {
    const out = new Set();
    const s = sinAcentos(t || '');
    const re = /(?:\bN\.?\s?C\.?\b|no\s*conformidad(?:es)?)[^0-9\n]{0,14}((?:\d{1,4})(?:\s*(?:,|y|e|\/|-)\s*(?:N\.?\s?C\.?\s*)?(?:N[o°º.]*\s*)?\d{1,4})*)/gi;
    let m;
    while ((m = re.exec(s))) for (const d of m[1].match(/\d{1,4}/g) || []) { const n = +d; if (n >= 1 && n <= 2000) out.add(n); }
    return [...out];
  }

  function parseA03(buffer, XLSX) {
    const wb = XLSX.read(buffer, { type: 'array', sheets: 'Planilla', cellNF: true, cellDates: false });
    const ws = wb.Sheets['Planilla'];
    if (!ws) throw new Error('El archivo no tiene la hoja "Planilla". ¿Es el A.03 Reclamos de Cliente?');
    const rango = XLSX.utils.decode_range(ws['!ref']);
    const C = (r, c) => ws[XLSX.utils.encode_cell({ r, c })];
    const hdr = [...Array(17).keys()].map((c) => sinAcentos(txt(C(2, c) && C(2, c).v) || '').toLowerCase());
    if (!/^n/.test(hdr[0]) || !/fecha/.test(hdr[1]) || !/cliente/.test(hdr[4]) || !/motivo/.test(hdr[6])) {
      throw new Error('No reconozco las columnas de la hoja Planilla (se esperaba Nº, Fecha, …, Cliente/Proveedor, …, Motivo).');
    }
    const reclamos = [], avisos = [], vistos = new Set();
    for (let r = 3; r <= rango.e.r; r++) {
      const a = C(r, 0);
      if (!a || a.v == null || String(a.v).trim() === '') continue;
      let numero = typeof a.v === 'number' ? String(Math.round(a.v)) : String(a.v).trim().toLowerCase().replace(/\s+/g, '');
      if (!/^\d{1,5}(-[a-z])?$/.test(numero)) { avisos.push(`Fila ${r + 1}: número "${a.v}" no reconocido, se omite.`); continue; }
      if (vistos.has(numero)) { avisos.push(`Fila ${r + 1}: el Nº ${numero} está repetido, se omite la segunda.`); continue; }
      const fecha = fechaCelda(C(r, 1));
      const cliente = txt(C(r, 4) && C(r, 4).v);
      const motivo = txt(C(r, 6) && C(r, 6).v);
      if (!fecha) { avisos.push(`Nº ${numero}: fecha ilegible, se omite.`); continue; }
      if (!cliente) { avisos.push(`Nº ${numero}: sin cliente, se omite.`); continue; }
      if (!motivo) { avisos.push(`Nº ${numero}: sin motivo, se omite.`); continue; }
      vistos.add(numero);
      const otiC = C(r, 5);
      const oti = !otiC || otiC.v == null ? null : esFecha(otiC) ? dmy(isoDeSerial(otiC.v)) : txt(otiC.v);
      const accion = txt(C(r, 9) && C(r, 9).v), obs = txt(C(r, 10) && C(r, 10).v), medio = txt(C(r, 11) && C(r, 11).v);
      const tipoTxt = sinAcentos(txt(C(r, 12) && C(r, 12).v) || '').toLowerCase();
      const dev = C(r, 13) ? flag(C(r, 13).v) : null;
      reclamos.push({
        numero, fecha, cliente,
        tipo: /prov/.test(tipoTxt) ? 'Proveedor' : 'Cliente',
        oti: oti === '-' ? null : oti, motivo,
        pertinente: tri(C(r, 7) && C(r, 7).v), significativo: tri(C(r, 8) && C(r, 8).v),
        accion_inmediata: accion, observaciones: obs, medio,
        devolucion: dev, cant_devuelta: txt(C(r, 16) && C(r, 16).v),
        nc_numeros: ncDeTexto([obs, accion].filter(Boolean).join(' . ')),
      });
    }
    return { reclamos, avisos };
  }

  g.parseA03 = parseA03; g.ncDeTexto = ncDeTexto;
  if (typeof module !== 'undefined') module.exports = { parseA03, ncDeTexto };
})(typeof window !== 'undefined' ? window : globalThis);
