/* Табель Т-12 в формате шаблона: собираем xlsx из частей шаблона (стили, шапка, подвал) */
(function (root) {
  const ESC = (t) => String(t == null ? '' : t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  function colN(s) { let n = 0; for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n; }
  function colL(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  const serial = (y, m, d) => Math.round(Date.UTC(y, m - 1, d) / 86400000) + 25569;
  const MONTHS = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
  const dimOf = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();
  const istr = (ref, s, v, text) => '<c r="' + ref + '" s="' + s + '" t="inlineStr"><is><t xml:space="preserve">' + ESC(text) + '</t></is></c>';
  const empty = (ref, s) => '<c r="' + ref + '" s="' + s + '"/>';
  const num = (ref, s, v) => '<c r="' + ref + '" s="' + s + '"><v>' + v + '</v></c>';
  const fml = (ref, s, f, v, bool) => '<c r="' + ref + '" s="' + s + '"' + (bool ? ' t="b"' : '') + '><f>' + ESC(f) + '</f><v>' + v + '</v></c>';

  function sheetXml(A, sh, spec, idx) {
    const y = sh.year, m = sh.month, dim = dimOf(y, m), P = sh.people, N = Math.max(1, P.length);
    let rowsXml = '';
    // верх: строки 1..15 как в шаблоне
    A.rows_top.forEach((r, i) => {
      const rn = i + 1;
      if (rn === 1) r = r.replace(/<c r="A1"[^>]*?(?:\/>|>.*?<\/c>)/, istr('A1', 0, 0, spec.org).replace('s="0"', 's="' + (/<c r="A1" s="(\d+)"/.exec(r) || [0, 0])[1] + '"'));
      if (rn === 3) r = r.replace(/<c r="A3"[^>]*?(?:\/>|>.*?<\/c>)/, istr('A3', 0, 0, spec.dept).replace('s="0"', 's="' + (/<c r="A3" s="(\d+)"/.exec(r) || [0, 0])[1] + '"'));
      if (rn === 8) {
        const last = serial(y, m, dim), first = serial(y, m, 1);
        r = r.replace(/(<c r="AJ8" s="\d+">)<v>\d+<\/v>/, '$1<v>' + last + '</v>').replace(/(<c r="AM8" s="\d+">)<v>\d+<\/v>/, '$1<v>' + first + '</v>').replace(/(<c r="AO8" s="\d+"><f>AJ8<\/f>)<v>\d+<\/v>/, '$1<v>' + last + '</v>');
      }
      if (rn === 12) {
        [29, 30, 31].forEach((d) => {
          const L = colL(d <= 15 ? 3 + d : 4 + d); // D=4 -> день1; день16 в T(20)
          const ref = L + '12';
          const st = (new RegExp('<c r="' + ref + '" s="(\\d+)"').exec(r) || [0, 55])[1];
          const cell = d <= dim ? num(ref, st, d) : '<c r="' + ref + '" s="' + st + '" t="s"><v>' + A.xChar + '</v></c>';
          r = r.replace(new RegExp('<c r="' + ref + '"[^>]*?(?:/>|>.*?</c>)'), cell);
        });
      }
      rowsXml += r;
    });
    // сотрудники
    const merges = A.merges_top.slice();
    P.forEach((p, k) => {
      const r1 = 16 + 2 * k, r2 = r1 + 1, S = k === 0 ? A.pair.first : A.pair.next;
      const sty1 = {}, sty2 = {}; S.r1.forEach((x) => { sty1[x[0]] = x[1]; }); S.r2.forEach((x) => { sty2[x[0]] = x[1]; });
      const days = p.days || [];
      let c1 = '', c2 = '', h1 = 0, h2 = 0, n1 = 0, n2 = 0;
      const dayCol = (d) => colL(d <= 15 ? 3 + d : 4 + d);
      const helperCol = (d) => colL(d <= 15 ? 43 + d : 59 + (d - 15)); // AR=44 -> день1 ; BH=60 -> день16
      c1 += istr('A' + r1, sty1.A, 0, String(k + 1)) + empty('B' + r1, sty1.B) + istr('C' + r1, sty1.C, 0, p.fio + (p.role ? ' (' + p.role + ')' : ''));
      c2 += empty('A' + r2, sty2.A) + empty('B' + r2, sty2.B) + empty('C' + r2, sty2.C);
      for (let d = 1; d <= 31; d++) {
        const L = dayCol(d), x = d <= dim ? (days[d - 1] || { c: 'в', h: 0 }) : null;
        if (d === 16) {
          c1 += fml('S' + r1, sty1.S, 'SUM(AR' + r2 + ':BF' + r2 + ')', n1);
          c2 += fml('S' + r2, sty2.S, 'SUM(D' + r2 + ':R' + r2 + ')', h1);
        }
        if (x) {
          c1 += istr(L + r1, sty1[L], 0, x.c);
          const isW = x.c === 'я/д' || x.c === 'я/н' || x.c === 'РВ';
          if (isW) { if (d <= 15) { n1++; h1 += x.h; } else { n2++; h2 += x.h; } }
          c2 += x.h ? num(L + r2, sty2[L], x.h) : empty(L + r2, sty2[L]);
        } else { c1 += empty(L + r1, sty1[L]); c2 += empty(L + r2, sty2[L]); }
      }
      // итоги (после дня 31): AJ, AK, AL, AM..AP
      c1 += fml('AJ' + r1, sty1.AJ, 'SUM(BH' + r2 + ':BW' + r2 + ')', n2) + fml('AK' + r1, sty1.AK, 'AJ' + r1 + '+S' + r1, n1 + n2) + fml('AL' + r1, sty1.AL, 'S' + r2 + '+AJ' + r2, h1 + h2);
      c2 += fml('AJ' + r2, sty2.AJ, 'SUM(T' + r2 + ':AI' + r2 + ')', h2) + empty('AK' + r2, sty2.AK) + empty('AL' + r2, sty2.AL);
      ['AM', 'AN', 'AO', 'AP'].forEach((L) => { c1 += empty(L + r1, sty1[L]); c2 += empty(L + r2, sty2[L]); });
      // скрытые вспомогательные ячейки
      for (let d = 1; d <= 31; d++) {
        const H = helperCol(d), L = dayCol(d), x = d <= dim ? (days[d - 1] || { c: 'в' }) : { c: '' };
        const w = x.c === 'я/д' || x.c === 'я/н' || x.c === 'РВ';
        if (d === 16) { c1 += empty('BG' + r1, sty1.BG); c2 += empty('BG' + r2, sty2.BG); }
        c1 += fml(H + r1, sty1[H], 'OR(' + L + r1 + '="я/д",' + L + r1 + '="я/н",' + L + r1 + '="РВ")', w ? 1 : 0, true);
        c2 += fml(H + r2, sty2[H], '_xlfn.IFS(' + H + r1 + '=TRUE,1,' + H + r1 + '=FALSE,0)', w ? 1 : 0);
      }
      rowsXml += '<row r="' + r1 + '" spans="1:76" ht="30" customHeight="1">' + c1 + '</row><row r="' + r2 + '" spans="1:76" ht="30" customHeight="1">' + c2 + '</row>';
      ['A', 'C', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP'].forEach((L) => merges.push(L + r1 + ':' + L + r2));
    });
    // подвал
    const off = 2 * (N - 3), shiftRow = (xml) => xml.replace(/(<row r=")(\d+)(")/, (a, b, c, d) => b + (+c + off) + d).replace(/(<c r="[A-Z]+)(\d+)(")/g, (a, b, c, d) => b + (+c + off) + d);
    A.footer.forEach((r, i) => {
      let rr = r;
      if (i === 1) {
        const rn = 23;
        const rep = (ref, text) => { const mm = new RegExp('<c r="' + ref + rn + '" s="(\\d+)"[^>]*?(?:/>|>.*?</c>)').exec(rr); if (mm) rr = rr.replace(mm[0], istr(ref + rn, mm[1], 0, text)); };
        rep('D', spec.leaderPos); rep('J', spec.leaderShort); rep('P', spec.leaderShort);
      }
      rowsXml += shiftRow(rr);
    });
    A.merges_foot.forEach((mg) => { merges.push(mg.replace(/([A-Z]+)(\d+)/g, (a, b, c) => b + (+c + off))); });
    const lastRow = 24 + off;
    let head = A.head.replace(/<dimension ref="[^"]*"\/>/, '<dimension ref="A1:BX' + lastRow + '"/>');
    if (idx === 0) head = head.replace('<sheetView ', '<sheetView tabSelected="1" ');
    // убираем ссылку на принтерные настройки и старые комментарии
    const tail = '<mergeCells count="' + merges.length + '">' + merges.map((x) => '<mergeCell ref="' + x + '"/>').join('') + '</mergeCells><pageMargins left="0.37" right="0.57" top="0.75" bottom="0.17" header="0.3" footer="0.3"/><pageSetup paperSize="9" scale="56" orientation="landscape"/></worksheet>';
    return head + '<sheetData>' + rowsXml + '</sheetData>' + tail;
  }

  async function build(JSZip, spec, A) {
    const zip = new JSZip();
    const sheets = spec.sheets;
    let ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>';
    let rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    let wbs = '';
    sheets.forEach((sh, i) => {
      zip.file('xl/worksheets/sheet' + (i + 1) + '.xml', sheetXml(A, sh, spec, i));
      ct += '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
      rels += '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>';
      wbs += '<sheet name="' + ESC(sh.name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>';
    });
    const n = sheets.length;
    rels += '<Relationship Id="rId' + (n + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId' + (n + 2) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId' + (n + 3) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>';
    ct += '</Types>';
    zip.file('[Content_Types].xml', ct);
    zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
    zip.file('xl/_rels/workbook.xml.rels', rels);
    zip.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="0"/></bookViews><sheets>' + wbs + '</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>');
    zip.file('xl/styles.xml', A.styles); zip.file('xl/theme/theme1.xml', A.theme); zip.file('xl/sharedStrings.xml', A.sst);
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }
  const api = { build: build, MONTHS: MONTHS, dimOf: dimOf };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.TabelX = api;
})(typeof window !== 'undefined' ? window : globalThis);
