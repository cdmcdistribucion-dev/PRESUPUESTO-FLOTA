/* Lee todos los .xlsx de la carpeta base_excels y genera datos/datos_presupuesto.json
   Uso:  node actualizar_datos.js   (requiere Node 18 o superior) */
const fs = require('fs'), path = require('path');
const XL = require('./lib/xl-lib.js');
const DIR = path.join(__dirname, 'base_excels'), OUT = path.join(__dirname, 'datos', 'datos_presupuesto.json');

(async () => {
  if (!fs.existsSync(DIR)) { console.error('No existe la carpeta base_excels.'); process.exit(1); }
  const files = fs.readdirSync(DIR).filter(f => /\.xlsx$/i.test(f) && !f.startsWith('~$')).sort();
  if (!files.length) { console.error('No hay archivos .xlsx en base_excels.'); process.exit(1); }
  const model = XL.emptyModel(); let ok = 0;
  for (const f of files) {
    try {
      const b = fs.readFileSync(path.join(DIR, f));
      const r = XL.parseWorkbook(await XL.readXlsx(b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength)), { year: 2026 });
      if (!r.kind) { console.log('  [omitido] ' + f + ': formato no reconocido'); continue; }
      let ols = r.ols;
      if (r.needsOL) {
        const ol = XL.guessOL(f);
        if (!ol || XL.OLS.indexOf(ol) < 0) { console.log('  [omitido] ' + f + ': no se pudo saber a qué operador corresponde (poné la sigla RDV, TL, TOL o ST en el nombre)'); continue; }
        ols = {}; ols[ol] = r.single;
      }
      const st = XL.applyImport(model, ols);
      (r.warnings || []).forEach(w => console.log('  [aviso] ' + f + ': ' + w));
      console.log('  [ok] ' + f + ' → ' + st.cells + ' celdas (' + Object.keys(ols).join(', ') + ')'); ok++;
    } catch (e) { console.log('  [error] ' + f + ': ' + e.message); }
  }
  if (!ok) { console.error('No se pudo leer ningún archivo. No se modificó el JSON.'); process.exit(1); }
  model.meta = { savedAt: new Date().toISOString(), savedBy: 'actualizar_datos.js', version: ok, files: files };
  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, JSON.stringify(model));
  console.log('Listo: ' + path.relative(__dirname, OUT) + ' generado con ' + ok + ' archivo(s).');
})();
