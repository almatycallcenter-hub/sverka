/* ═════════ Отчёт за месяц: запись .xlsx со стилями и графиками ═════════
   Свой маленький писатель, без библиотек: страница работает и без интернета
   к CDN, а файл должен открываться в Excel как сделанный руками — с цветами
   статусов, шириной колонок, закреплёнными заголовками и графиками.        */
var XW = (function(){
  function esc(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
  function colName(i){ var s = ''; i++; while (i > 0){ var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = (i - m - 1) / 26; } return s; }
  function ref(c, r){ return colName(c - 1) + r; }               // c, r — с единицы
  function aref(sh, c, r){ return "'" + sh.replace(/'/g, "''") + "'!$" + colName(c - 1) + '$' + r; }

  /* стили: шрифты, заливки, границы и сами xf */
  var FILL = { GRN:'E2EFDA', RED:'F8CBAD', YEL:'FFF2CC', GRY:'EDEDED', ORG:'FCE4D6', LGR:'EAF4E4', BLU:'DDEBF7', DGR:'D9D9D9' };
  var FILLS = ['HDR'].concat(Object.keys(FILL));                  // id заливки = индекс + 2
  var ST = { def:0, head:1, body:2, title:3, h2:4, total:5, wrap:6, bold:7 };
  var nBase = 8;
  Object.keys(FILL).forEach(function(k, i){ ST[k] = nBase + i; });
  function stylesXml(){
    var fills = '<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/></patternFill></fill>' +
      Object.keys(FILL).map(function(k){ return '<fill><patternFill patternType="solid"><fgColor rgb="FF' + FILL[k] + '"/></patternFill></fill>'; }).join('');
    var bd = '<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right>' +
             '<top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>';
    var top = '<alignment wrapText="1" vertical="top"/>';
    var xfs = [
      '<xf xfId="0" fontId="0" fillId="0" borderId="0"/>',
      '<xf xfId="0" fontId="1" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment wrapText="1" vertical="center"/></xf>',
      '<xf xfId="0" fontId="0" fillId="0" borderId="1" applyBorder="1" applyAlignment="1">' + top + '</xf>',
      '<xf xfId="0" fontId="2" fillId="0" borderId="0" applyFont="1"/>',
      '<xf xfId="0" fontId="3" fillId="0" borderId="0" applyFont="1"/>',
      '<xf xfId="0" fontId="4" fillId="' + (2 + FILLS.indexOf('BLU')) + '" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>',
      '<xf xfId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1">' + top + '</xf>',
      '<xf xfId="0" fontId="4" fillId="0" borderId="1" applyFont="1" applyBorder="1"/>'
    ];
    Object.keys(FILL).forEach(function(k){
      xfs.push('<xf xfId="0" fontId="0" fillId="' + (2 + FILLS.indexOf(k)) + '" borderId="1" applyFill="1" applyBorder="1" applyAlignment="1">' + top + '</xf>');
    });
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<fonts count="5"><font><sz val="10"/><name val="Arial"/></font>' +
      '<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>' +
      '<font><b/><sz val="14"/><color rgb="FF1F4E79"/><name val="Arial"/></font>' +
      '<font><b/><sz val="11"/><color rgb="FF1F4E79"/><name val="Arial"/></font>' +
      '<font><b/><sz val="10"/><name val="Arial"/></font></fonts>' +
      '<fills count="' + (2 + FILLS.length) + '">' + fills + '</fills>' +
      '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' + bd + '</borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="' + xfs.length + '">' + xfs.join('') + '</cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  }

  function Sheet(name){
    this.name = name; this.cells = {}; this.maxR = 0; this.maxC = 0; this.widths = {};
    this.heights = {}; this.merges = []; this.freeze = null; this.filter = null; this.charts = []; this.grid = true;
  }
  Sheet.prototype.set = function(r, c, v, s){
    if (!this.cells[r]) this.cells[r] = {};
    this.cells[r][c] = { v: v, s: s || 0 };
    if (r > this.maxR) this.maxR = r; if (c > this.maxC) this.maxC = c;
  };
  Sheet.prototype.width = function(c, w){ if (!this.widths[c] || this.widths[c] < w) this.widths[c] = w; };
  /* Таблица: шапка, строки, итог. chk — индексы колонок, которые красим по
     статусу (fill(значение) → ключ заливки), fills — заливка на каждую ячейку. */
  Sheet.prototype.table = function(r0, head, data, o){
    o = o || {}; var c0 = o.c0 || 1, self = this;
    if (o.title){ this.set(r0, c0, o.title, ST.h2); r0++; }
    head.forEach(function(h, j){ self.set(r0, c0 + j, h, ST.head); });
    this.heights[r0] = 42;
    data.forEach(function(row, i){
      row.forEach(function(v, j){
        var s = ST.body;
        if (o.chk && o.chk.indexOf(j) >= 0 && o.fill){ var f = o.fill(v); if (f) s = ST[f]; }
        if (o.fills && o.fills[i] && o.fills[i][j]) s = ST[o.fills[i][j]];
        self.set(r0 + 1 + i, c0 + j, v, s);
      });
    });
    var end = r0 + data.length;
    if (o.total){ o.total.forEach(function(v, j){ self.set(end + 1, c0 + j, v, ST.total); }); end++; }
    (o.widths || []).forEach(function(w, j){ self.width(c0 + j, w); });
    return [r0, end];
  };
  Sheet.prototype.text = function(r, c, v, s){ this.set(r, c, v, s === undefined ? ST.wrap : s); };
  Sheet.prototype.merge = function(r1, c1, r2, c2){ this.merges.push(ref(c1, r1) + ':' + ref(c2, r2)); };
  /* График: kind bar|col|line, grouping clustered|stacked|percentStacked;
     series — [{nameCell:[c,r], val:[c, r1, r2]}], cat — [c, r1, r2]; at — [col, row] (с единицы), size — [колонок, строк]. */
  Sheet.prototype.chart = function(o){ this.charts.push(o); };

  function cellXml(r, c, cell){
    var v = cell.v, s = cell.s, a = ref(c, r);
    if (v === '' || v === null || v === undefined) return s ? '<c r="' + a + '" s="' + s + '"/>' : '';
    if (typeof v === 'number' && isFinite(v)) return '<c r="' + a + '" s="' + s + '"><v>' + v + '</v></c>';
    return '<c r="' + a + '" s="' + s + '" t="inlineStr"><is><t xml:space="preserve">' + esc(v) + '</t></is></c>';
  }
  function sheetXml(sh, hasDrawing){
    var x = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'];
    x.push('<sheetViews><sheetView workbookViewId="0"' + (sh.grid ? '' : ' showGridLines="0"') + '>');
    if (sh.freeze){
      var fc = sh.freeze[0], fr = sh.freeze[1];   // первая незакреплённая колонка/строка (с единицы)
      x.push('<pane' + (fc > 1 ? ' xSplit="' + (fc - 1) + '"' : '') + (fr > 1 ? ' ySplit="' + (fr - 1) + '"' : '') +
             ' topLeftCell="' + ref(fc, fr) + '" activePane="' + (fc > 1 && fr > 1 ? 'bottomRight' : fr > 1 ? 'bottomLeft' : 'topRight') + '" state="frozen"/>');
    }
    x.push('</sheetView></sheetViews><sheetFormatPr defaultRowHeight="13"/>');
    var ws = Object.keys(sh.widths).map(Number).sort(function(a, b){ return a - b; });
    if (ws.length) x.push('<cols>' + ws.map(function(c){ return '<col min="' + c + '" max="' + c + '" width="' + sh.widths[c] + '" customWidth="1"/>'; }).join('') + '</cols>');
    x.push('<sheetData>');
    for (var r = 1; r <= sh.maxR; r++){
      var row = sh.cells[r]; if (!row && !sh.heights[r]) continue;
      var cs = '';
      if (row) Object.keys(row).map(Number).sort(function(a, b){ return a - b; }).forEach(function(c){ cs += cellXml(r, c, row[c]); });
      x.push('<row r="' + r + '"' + (sh.heights[r] ? ' ht="' + sh.heights[r] + '" customHeight="1"' : '') + '>' + cs + '</row>');
    }
    x.push('</sheetData>');
    if (sh.filter) x.push('<autoFilter ref="' + sh.filter + '"/>');
    if (sh.merges.length) x.push('<mergeCells count="' + sh.merges.length + '">' + sh.merges.map(function(m){ return '<mergeCell ref="' + m + '"/>'; }).join('') + '</mergeCells>');
    if (hasDrawing) x.push('<drawing r:id="rId1"/>');
    x.push('</worksheet>');
    return x.join('');
  }

  var CH = 'http://schemas.openxmlformats.org/drawingml/2006/chart', DA = 'http://schemas.openxmlformats.org/drawingml/2006/main',
      RN = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  function cellVal(sh, c, r){ var row = sh.cells[r]; return row && row[c] ? row[c].v : ''; }
  var PAL = ['2E75B6', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '264478', '9E480E'];
  function chartXml(sh, o){
    function strCache(c, r1, r2){
      var pts = ''; for (var r = r1; r <= r2; r++) pts += '<c:pt idx="' + (r - r1) + '"><c:v>' + esc(cellVal(sh, c, r)) + '</c:v></c:pt>';
      return '<c:strCache><c:ptCount val="' + (r2 - r1 + 1) + '"/>' + pts + '</c:strCache>';
    }
    function numCache(c, r1, r2){
      var pts = ''; for (var r = r1; r <= r2; r++){ var v = cellVal(sh, c, r); if (typeof v === 'number') pts += '<c:pt idx="' + (r - r1) + '"><c:v>' + v + '</c:v></c:pt>'; }
      return '<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="' + (r2 - r1 + 1) + '"/>' + pts + '</c:numCache>';
    }
    var cat = o.cat, line = o.kind === 'line';
    var ser = o.series.map(function(s, i){
      var nm = s.nameCell ? '<c:tx><c:strRef><c:f>' + aref(sh.name, s.nameCell[0], s.nameCell[1]) + '</c:f>' + strCache(s.nameCell[0], s.nameCell[1], s.nameCell[1]) + '</c:strRef></c:tx>' : '';
      var col = (s.color || (o.colors && o.colors[i]) || PAL[i % PAL.length]);
      var sp = line ? '<c:spPr><a:ln w="22225"><a:solidFill><a:srgbClr val="' + col + '"/></a:solidFill></a:ln></c:spPr>'
                    : '<c:spPr><a:solidFill><a:srgbClr val="' + col + '"/></a:solidFill></c:spPr>';
      return '<c:ser><c:idx val="' + i + '"/><c:order val="' + i + '"/>' + nm + sp +
        (line ? '<c:marker><c:symbol val="circle"/><c:size val="4"/><c:spPr><a:solidFill><a:srgbClr val="' + col + '"/></a:solidFill></c:spPr></c:marker>' : '<c:invertIfNegative val="0"/>') +
        '<c:cat><c:strRef><c:f>' + aref(sh.name, cat[0], cat[1]) + ':$' + colName(cat[0] - 1) + '$' + cat[2] + '</c:f>' + strCache(cat[0], cat[1], cat[2]) + '</c:strRef></c:cat>' +
        '<c:val><c:numRef><c:f>' + aref(sh.name, s.val[0], s.val[1]) + ':$' + colName(s.val[0] - 1) + '$' + s.val[2] + '</c:f>' + numCache(s.val[0], s.val[1], s.val[2]) + '</c:numRef></c:val>' +
        (line ? '<c:smooth val="0"/>' : '') + '</c:ser>';
    }).join('');
    var g = o.grouping || 'clustered';
    var plot = line
      ? '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>' + ser + '<c:marker val="1"/><c:axId val="1001"/><c:axId val="1002"/></c:lineChart>'
      : '<c:barChart><c:barDir val="' + (o.kind === 'bar' ? 'bar' : 'col') + '"/><c:grouping val="' + g + '"/><c:varyColors val="0"/>' + ser +
        '<c:gapWidth val="80"/>' + (g !== 'clustered' ? '<c:overlap val="100"/>' : '') + '<c:axId val="1001"/><c:axId val="1002"/></c:barChart>';
    var horiz = o.kind === 'bar';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<c:chartSpace xmlns:c="' + CH + '" xmlns:a="' + DA + '" xmlns:r="' + RN + '"><c:roundedCorners val="0"/><c:chart>' +
      '<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1200" b="1"/></a:pPr><a:r><a:rPr lang="ru-RU" sz="1200" b="1"/><a:t>' + esc(o.title || '') + '</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>' +
      '<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>' + plot +
      '<c:catAx><c:axId val="1001"/><c:scaling><c:orientation val="' + (horiz ? 'maxMin' : 'minMax') + '"/></c:scaling><c:delete val="0"/><c:axPos val="' + (horiz ? 'l' : 'b') + '"/>' +
      '<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="1002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>' +
      '<c:valAx><c:axId val="1002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="' + (horiz ? 'b' : 'l') + '"/><c:majorGridlines/>' +
      '<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="1001"/><c:crosses val="' + (horiz ? 'max' : 'autoZero') + '"/><c:crossBetween val="between"/></c:valAx>' +
      '</c:plotArea>' + (o.legend === false ? '' : '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>') +
      '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>';
  }
  function drawingXml(sh, firstId){
    var a = sh.charts.map(function(o, i){
      var c1 = o.at[0] - 1, r1 = o.at[1] - 1, c2 = c1 + (o.size ? o.size[0] : 8), r2 = r1 + (o.size ? o.size[1] : 16);
      return '<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>' + c1 + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + r1 + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
        '<xdr:to><xdr:col>' + c2 + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + r2 + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
        '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="' + (i + 2) + '" name="Chart ' + (i + 1) + '"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>' +
        '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="' + CH + '">' +
        '<c:chart xmlns:c="' + CH + '" xmlns:r="' + RN + '" r:id="rId' + (i + 1) + '"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>';
    }).join('');
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="' + DA + '">' + a + '</xdr:wsDr>';
  }

  /* zip: deflate через CompressionStream, где он есть; иначе без сжатия */
  var CRC = (function(){ var t = []; for (var n = 0; n < 256; n++){ var c = n; for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  function crc32(u8){ var c = 0xFFFFFFFF; for (var i = 0; i < u8.length; i++) c = CRC[(c ^ u8[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  async function deflate(u8){
    if (typeof CompressionStream !== 'function') return null;
    try{ var b = await new Response(new Blob([u8]).stream().pipeThrough(new CompressionStream('deflate-raw'))).arrayBuffer(); return new Uint8Array(b); }
    catch(e){ return null; }
  }
  async function zip(files){
    var enc = new TextEncoder(), chunks = [], central = [], off = 0;
    for (var i = 0; i < files.length; i++){
      var f = files[i], data = typeof f.data === 'string' ? enc.encode(f.data) : f.data;
      var nameB = enc.encode(f.name), crc = crc32(data), z = await deflate(data), method = z ? 8 : 0, body = z || data;
      var lh = new Uint8Array(30 + nameB.length), dv = new DataView(lh.buffer);
      dv.setUint32(0, 0x04034b50, true); dv.setUint16(4, 20, true); dv.setUint16(6, 0x0800, true); dv.setUint16(8, method, true);
      dv.setUint32(14, crc, true); dv.setUint32(18, body.length, true); dv.setUint32(22, data.length, true); dv.setUint16(26, nameB.length, true);
      lh.set(nameB, 30); chunks.push(lh, body);
      var ch = new Uint8Array(46 + nameB.length), cv = new DataView(ch.buffer);
      cv.setUint32(0, 0x02014b50, true); cv.setUint16(4, 20, true); cv.setUint16(6, 20, true); cv.setUint16(8, 0x0800, true); cv.setUint16(10, method, true);
      cv.setUint32(16, crc, true); cv.setUint32(20, body.length, true); cv.setUint32(24, data.length, true); cv.setUint16(28, nameB.length, true); cv.setUint32(42, off, true);
      ch.set(nameB, 46); central.push(ch);
      off += lh.length + body.length;
    }
    var cd = central.reduce(function(a, b){ return a + b.length; }, 0), eo = new Uint8Array(22), ev = new DataView(eo.buffer);
    ev.setUint32(0, 0x06054b50, true); ev.setUint16(8, files.length, true); ev.setUint16(10, files.length, true); ev.setUint32(12, cd, true); ev.setUint32(16, off, true);
    return new Blob(chunks.concat(central, [eo]), { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  function Book(){ this.sheets = []; }
  Book.prototype.sheet = function(name){ var s = new Sheet(name); this.sheets.push(s); return s; };
  Book.prototype.blob = async function(){
    var files = [], ct = [], rel = [], st = [], nChart = 0, nDraw = 0;
    this.sheets.forEach(function(sh, i){
      var n = i + 1, hasD = sh.charts.length > 0;
      st.push('<sheet name="' + esc(sh.name.slice(0, 31)) + '" sheetId="' + n + '" r:id="rId' + n + '"/>');
      rel.push('<Relationship Id="rId' + n + '" Type="' + RN + '/worksheet" Target="worksheets/sheet' + n + '.xml"/>');
      ct.push('<Override PartName="/xl/worksheets/sheet' + n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');
      files.push({ name: 'xl/worksheets/sheet' + n + '.xml', data: sheetXml(sh, hasD) });
      if (hasD){
        nDraw++;
        files.push({ name: 'xl/worksheets/_rels/sheet' + n + '.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="' + RN + '/drawing" Target="../drawings/drawing' + nDraw + '.xml"/></Relationships>' });
        files.push({ name: 'xl/drawings/drawing' + nDraw + '.xml', data: drawingXml(sh) });
        ct.push('<Override PartName="/xl/drawings/drawing' + nDraw + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>');
        var dr = [];
        sh.charts.forEach(function(o, j){
          nChart++;
          dr.push('<Relationship Id="rId' + (j + 1) + '" Type="' + RN + '/chart" Target="../charts/chart' + nChart + '.xml"/>');
          files.push({ name: 'xl/charts/chart' + nChart + '.xml', data: chartXml(sh, o) });
          ct.push('<Override PartName="/xl/charts/chart' + nChart + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>');
        });
        files.push({ name: 'xl/drawings/_rels/drawing' + nDraw + '.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + dr.join('') + '</Relationships>' });
      }
    });
    var n = this.sheets.length;
    files.unshift(
      { name: '[Content_Types].xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' + ct.join('') + '</Types>' },
      { name: '_rels/.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="' + RN + '/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
      { name: 'xl/workbook.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="' + RN + '">' +
        '<bookViews><workbookView/></bookViews><sheets>' + st.join('') + '</sheets></workbook>' },
      { name: 'xl/_rels/workbook.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + rel.join('') +
        '<Relationship Id="rId' + (n + 1) + '" Type="' + RN + '/styles" Target="styles.xml"/></Relationships>' },
      { name: 'xl/styles.xml', data: stylesXml() });
    return zip(files);
  };
  return { Book: Book, ST: ST, colName: colName };
})();
/* ═════════ Отчёт за месяц: скидка и списание по срокам ═════════
   Тот же расчёт, что в ноябрьском отчёте за октябрь (Claude, 03.11.2026),
   только прямо в странице и за любой период. Правила — в листе «Как
   считалось» книги; здесь только то, что нужно, чтобы цифры совпали.        */
var MON = (function(){
  var EPS = 1e-9;
  var NZ = 'таңқурай нәзіктігі торты';
  var IN_K = ['prihod_gp', 'prihod_tort', 'prihod_other', 'prigotovlenie', 'inv_in'];
  var OUT_SITE = ['realiz', 'spisanie', 'vozvrat_srok', 'vozvrat_kach', 'perem_50', 'perem_dost', 'ap_rashod', 'inv_out', 'perem_out'];
  var OUT_K = ['other'].concat(OUT_SITE);
  var MOVK = ['other', 'prihod_gp', 'prihod_other', 'prihod_tort', 'prigotovlenie', 'realiz', 'perem_50', 'spisanie', 'vozvrat_srok',
              'vozvrat_kach', 'perem_dost', 'perem_out', 'perem_in', 'inv_in', 'inv_out'];
  var VES = /печенье вупи|бауырсақ|рогалик|көкнәр/i;
  var KSRE = /круасс|сэндвич|сендвич/i;
  var BRN = { 1:'ТТ1 Атакент', 2:'ТТ2 Сарыарка', 3:'ТТ3 Ламия', 4:'ТТ4 Ньютон', 6:'ТТ6 Автодом', 8:'ТТ8 Каратау', 9:'ТТ9 Бесагаш', 10:'ТТ10 Шугыла', 11:'ТТ11 Алма-Сити' };

  function fx(v){ if (typeof v !== 'number') return v; var r = Math.round(v * 1e6) / 1e6; return r === 0 ? 0 : r; }
  function nz(v){ return v ? fx(v) : ''; }
  function pct(a, b){ return b ? Math.round(a / b * 10000) / 100 : ''; }
  function sp_(x){ var s = String(x).split('.'); return s[0].replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + (s[1] ? ',' + s[1] : ''); }
  function days(a, b){ var p = a.split('-'), q = b.split('-'); return Math.round((new Date(+q[0], +q[1] - 1, +q[2]) - new Date(+p[0], +p[1] - 1, +p[2])) / 864e5); }
  function rng(from, to){ var o = [], d = from; while (d <= to){ o.push(d); d = shift(d, 1); } return o; }
  function brName(wh){
    var n = whNum(wh);
    return n && BRN[n] ? BRN[n] : String(wh).replace(/\s*склад\s*$/i, '').trim();
  }
  function cnt(){ var o = {}; o.add = function(k, v){ o[k] = (o[k] || 0) + (v === undefined ? 1 : v); }; return o; }
  function g(o, k){ return o[k] || 0; }
  function stripId(v){ return typeof v === 'string' ? v.replace(/\s*\/?\s*\d{9,}/g, '').trim() : v; }

  /* название → [ключ справочника, доля]: кусочек = 1/8 своего торта,
     «Таңқурайлы торт кусочек» — кусочек «Таңқурай нәзіктігі торты» */
  var _bo = {};
  function bo(name){
    if (_bo[name]) return _bo[name];
    var t = nrm(name), r;
    if (t === 'таңқурайлы торт кусочек' || t === 'таңқурайлы торт (кусочек)') r = { k: NZ, f: 1 / KUSOK };
    else r = baseOf(t);
    return (_bo[name] = r);
  }
  function isKs(k){ return KSRE.test(k); }
  function sprOf(k){ var s = SPR[k]; if (!s) return null;
    return isKs(k) ? Object.assign({}, s, { srok: 24, ks24: true }) : s; }   // 03.11: круассаны и сэндвичи — 24 ч

  /* ---------- подготовка данных ---------- */
  function prep(from, to){
    var mfrom = shift(from, -45), mto = shift(to, 2);
    var WHS = whList().filter(function(w){
      return !(typeof isDelivery === 'function' && isDelivery(w)) && !(typeof isClosed === 'function' && isClosed(w));
    });
    var wIdx = {}; WHS.forEach(function(w, i){ wIdx[w] = i; });
    var MV = {}, HAS_TT = {}, seen = {};
    DB.mov.forEach(function(r){
      var wi = wIdx[r[1]]; if (wi === undefined) return;
      var kd = r[3];
      if (kd === 'other' && r[6] === 'in') return;          // нераспознанный приход — не трогаем, как сайт
      var b = bo(r[2]); if (!SPR[b.k]) return;
      var q = N(r[4]) * b.f;
      if (kd === 'prihod_tort' && q > 0) HAS_TT[wi + '|' + b.k] = 1;
      if (r[0] < mfrom || r[0] > mto) return;
      var key = wi + '|' + b.k + '|' + r[0] + '|' + kd;
      MV[key] = (MV[key] || 0) + q; seen[wi] = 1;
    });
    var S20 = {}, S20SUM = {}, S20T = [];
    DB.s20.forEach(function(r){
      if (r[0] < mfrom || r[0] > mto) return;
      var b = bo(r[2]); if (!SPR[b.k]) return;
      for (var i = 0; i < WHS.length; i++) if (sameBranch(r[1], WHS[i])){
        var key = i + '|' + b.k + '|' + r[0];
        S20[key] = (S20[key] || 0) + N(r[3]) * b.f; S20SUM[key] = (S20SUM[key] || 0) + N(r[4]);
        if (isKs(b.k) && r[0] >= from && r[0] <= to){
          var h = parseInt(S(r[5]).slice(0, 2), 10);
          S20T.push({ wi: i, k: b.k, d: r[0], hr: isNaN(h) ? -1 : h, q: N(r[3]) * b.f });
        }
        break;
      }
    });
    var LIST = {};
    DB.list.forEach(function(r){
      if (r[0] < from || r[0] > to) return;
      var wi = wIdx[r[1]]; if (wi === undefined) return;
      var b = bo(r[4] || r[2]); if (!SPR[b.k]) return;
      var key = wi + '|' + b.k + '|' + r[0];
      if (S(r[3]) === ''){ if (!(key in LIST)) LIST[key] = null; }
      else LIST[key] = (LIST[key] || 0) + N(r[3]) * b.f;
    });
    return { WHS: WHS, MV: MV, HAS_TT: HAS_TT, S20: S20, S20SUM: S20SUM, S20T: S20T, LIST: LIST, used: seen, from: from, to: to };
  }

  /* ---------- FIFO ---------- */
  function engine(P){
    var MV = P.MV;
    function mv(wi, k, d, kd){ return MV[wi + '|' + k + '|' + d + '|' + kd] || 0; }
    function lotsAt(wi, k, T, dd, keep){
      var d = shift(T, -(dd * 3 + keep)), lots = [];
      while (d < T){
        var oth = mv(wi, k, d, 'prihod_other'), tort = mv(wi, k, d, 'prihod_tort'), tot = 0;
        IN_K.forEach(function(kd){ var v = mv(wi, k, d, kd); if (v > 0) tot += v; });
        if (tot > EPS) lots.push({ d: d, q: tot, oth: Math.max(oth, 0), tort: Math.max(tort, 0) });
        var dd0 = d;
        lots = lots.filter(function(L){ return shift(L.d, keep - 1) >= dd0; });
        var need = 0; OUT_K.forEach(function(kd){ need += Math.max(mv(wi, k, d, kd), 0); });
        for (var i = 0; i < lots.length && need > EPS; i++){ var t = Math.min(lots[i].q, need); lots[i].q -= t; need -= t; }
        lots = lots.filter(function(L){ return L.q > EPS; });
        d = shift(d, 1);
      }
      return lots;
    }
    function siteLeft(wh, k, dd, T){ try{ return expLeft(expMov(wh), k, dd, T); }catch(e){ return null; } }
    function frozen(wi, k, T, dd){
      var s = SPR[k];
      if (frozenByName(s.name) || P.TORT[k] || P.HAS_TT[wi + '|' + k]) return true;
      for (var b = 0; b < dd; b++) if (mv(wi, k, shift(T, -b), 'prihod_tort') > 0) return true;
      return false;
    }
    function lotsTxt(used){ return used.map(function(u){ return ru(u[0]) + ' — ' + fx(u[1]); }).join('; '); }
    function classify(wi, k, T, Pd, dd, c, sumfact, kind){
      var s = sprOf(k), ks = !!s.ks24;
      var gp = mv(wi, k, Pd, 'prihod_gp'), oth = mv(wi, k, Pd, 'prihod_other'), prig = mv(wi, k, Pd, 'prigotovlenie');
      var arr = ks ? gp + prig + oth : gp;
      var det = { gpP: gp, prigP: prig, othP: oth };
      if (dualSrok(s)) return ['Невозможно отследить', '', 'обе формы хранения', det];
      if (frozen(wi, k, T, dd)) return ['Невозможно отследить', '', 'заморозка (Таугуль торт)', det];
      var lots = lotsAt(wi, k, T, dd, 2 * dd), need = c, used = [], anyOth = false;
      for (var i = 0; i < lots.length && need > EPS; i++){ var t = Math.min(lots[i].q, need); used.push([lots[i].d, t]); if (lots[i].oth > EPS) anyOth = true; need -= t; }
      det.lots = lotsTxt(used);
      det.left = kind === 'disc' ? siteLeft(P.WHS[wi], k, dd, T) : null;
      if (!ks && (oth > EPS || anyOth)) return ['Невозможно отследить', '', 'перемещение из другого филиала', det];
      if (need > EPS){
        if (arr <= EPS && !used.length) return ['Нет прихода по сроку', '', '', det];
        if (arr <= EPS) return ['Нет прихода по сроку', '', 'приход был в другие дни: ' + det.lots, det];
        if (arr < c - EPS) return [kind === 'disc' ? 'Приход меньше скидки' : 'Приход меньше списания', '', '', det];
        if (sumfact >= c - EPS) return ['Превышен остаток FIFO (предупреждение)', '', 'по факту (20% + склад 50% + возврат + списание) ушло не меньше списка', det];
        return ['Превышен остаток FIFO', '', '', det];
      }
      var old = used.reduce(function(a, u){ return u[0] < a ? u[0] : a; }, '9'), nw = used.reduce(function(a, u){ return u[0] > a ? u[0] : a; }, '0');
      if (old < Pd) return ['Позже срока', days(old, Pd), det.lots, det];
      if (nw > Pd) return ['Раньше срока', -days(Pd, nw), det.lots, det];
      if (arr <= EPS) return ['Нет прихода по сроку', '', 'на эту дату нет прихода (Цех ГП, акт приготовления), есть только излишек по инвентаризации', det];
      return ['В срок', 0, '', det];
    }
    /* круассаны и сэндвичи: со скидкой — только вечером в день прихода,
       на следующий день их списывают. Остаток прошлого дня на скидке = «позже». */
    function classifyKs(wi, k, T, c){
      var arr = mv(wi, k, T, 'prihod_gp') + mv(wi, k, T, 'prigotovlenie') + mv(wi, k, T, 'prihod_other');
      var old = lotsAt(wi, k, T, 1, 3), need = c, used = [];
      var det = { gpP: mv(wi, k, T, 'prihod_gp'), prigP: mv(wi, k, T, 'prigotovlenie'), othP: mv(wi, k, T, 'prihod_other'), left: null };
      for (var i = 0; i < old.length && need > EPS; i++){ var t = Math.min(old[i].q, need); used.push([old[i].d, t]); need -= t; }
      var fromT = Math.min(need, arr); need -= fromT; if (fromT > EPS) used.push([T, fromT]);
      det.lots = lotsTxt(used);
      var olds = used.filter(function(u){ return u[0] < T; });
      if (olds.length){ var o = olds.reduce(function(a, u){ return u[0] < a ? u[0] : a; }, '9');
        return ['Позже срока', days(o, T), 'на скидку ушла партия прошлого дня — её нужно было списать: ' + det.lots, det]; }
      if (need > EPS) return arr <= EPS ? ['Нет прихода по сроку', '', 'в день скидки прихода не было', det] : ['Приход меньше скидки', '', '', det];
      return ['В срок', 0, '', det];
    }
    function kindOf(k){
      var s = sprOf(k);
      if (isKus[k]) return 'kus';
      if (isPizza(s.name)) return 'pizza';
      if (!s.srok) return 'nosrok';
      return 'ok';
    }
    return { mv: mv, lotsAt: lotsAt, classify: classify, classifyKs: classifyKs, kindOf: kindOf };
  }

  function compute(from, to){
    var P = prep(from, to);
    P.TORT = (typeof tortPositions === 'function') ? tortPositions() : {};
    var E = engine(P), mv = E.mv, DATES = rng(from, to);
    var inPer = {}; DATES.forEach(function(d){ inPer[d] = 1; });
    var nm = function(k){ return SPR[k].name; };

    /* ---- сроки скидки ---- */
    var DISC = [];
    Object.keys(P.LIST).map(function(key){ var p = key.split('|'); return { wi: +p[0], k: p[1], T: p[2], c: P.LIST[key] }; })
      .sort(function(a, b){ return a.wi - b.wi || (a.T < b.T ? -1 : a.T > b.T ? 1 : (nm(a.k) < nm(b.k) ? -1 : 1)); })
      .forEach(function(x){
        var wi = x.wi, k = x.k, T = x.T, s = sprOf(k), kd = E.kindOf(k);
        var s20 = P.S20[wi + '|' + k + '|' + T] || 0, p50 = mv(wi, k, T, 'perem_50'), vz = mv(wi, k, shift(T, 1), 'vozvrat_srok'), sp1 = mv(wi, k, shift(T, 1), 'spisanie');
        var fact = s20 + p50 + vz + sp1, noq = x.c === null, cc = noq ? fact : x.c;
        var row = { wi: wi, k: k, T: T, list: x.c, noq: noq, s20: s20, p50: p50, vz: vz, sp1: sp1, fact: fact, kd: kd, dd: '', P: '', dev: '', why: '', det: {} };
        if (kd !== 'ok' || s.skidka === 'Нет'){
          row.st = kd === 'pizza' ? 'Не проверяется (пицца)' : s.skidka === 'Нет' ? 'Не проверяется (в Справочнике скидка «Нет»)' : 'Не проверяется (нет срока)';
          DISC.push(row); return;
        }
        var dd = Math.max(2, Math.round(s.srok / 24)), Pd = shift(T, -(dd - 1));
        if (s.ks24){ dd = 1; Pd = T; }
        row.dd = dd; row.P = Pd;
        if (cc <= EPS){ row.st = 'Нет количества'; row.why = 'в списке без количества и по факту 0'; row.det = { gpP: mv(wi, k, Pd, 'prihod_gp'), othP: mv(wi, k, Pd, 'prihod_other') }; DISC.push(row); return; }
        var res = s.ks24 ? E.classifyKs(wi, k, T, cc) : E.classify(wi, k, T, Pd, dd, cc, fact, 'disc');
        row.st = res[0]; row.dev = res[1]; row.why = res[2]; row.det = res[3];
        DISC.push(row);
      });

    /* ---- сроки списания: документы iiko ---- */
    var SPW = [];
    var spk = [];
    Object.keys(P.MV).forEach(function(key){
      var p = key.split('|'); if (p[3] !== 'spisanie' && p[3] !== 'vozvrat_srok') return;
      if (!inPer[p[2]] || P.MV[key] <= EPS) return;
      spk.push({ wi: +p[0], k: p[1], d: p[2], kd: p[3], q: P.MV[key] });
    });
    spk.sort(function(a, b){ return a.wi - b.wi || (a.d < b.d ? -1 : a.d > b.d ? 1 : (nm(a.k) < nm(b.k) ? -1 : nm(a.k) > nm(b.k) ? 1 : (a.kd < b.kd ? -1 : 1))); });
    spk.forEach(function(x){
      var s = sprOf(x.k), kk = E.kindOf(x.k);
      var row = { wi: x.wi, k: x.k, T: x.d, typ: x.kd === 'spisanie' ? 'Списание' : 'Возврат по сроку', q: x.q, kd: kk, dd: '', P: '', dev: '', why: '', det: {} };
      if (kk !== 'ok'){ row.st = kk === 'pizza' ? 'Не проверяется (пицца)' : 'Не проверяется (нет срока)'; SPW.push(row); return; }
      var dd = Math.max(1, Math.round(s.srok / 24)), Pd = shift(x.d, -dd);
      var res = E.classify(x.wi, x.k, x.d, Pd, dd + 1, x.q, 0, 'sp');
      row.dd = dd; row.P = Pd; row.st = res[0]; row.dev = res[1]; row.why = res[2]; row.det = res[3];
      SPW.push(row);
    });

    /* ---- движение по дням ---- */
    var DAY = {};
    function dv(wi, k, d){ var key = wi + '|' + k + '|' + d; return DAY[key] || (DAY[key] = { wi: wi, k: k, d: d }); }
    Object.keys(P.MV).forEach(function(key){
      var p = key.split('|'); if (!inPer[p[2]] || MOVK.indexOf(p[3]) < 0) return;
      var o = dv(+p[0], p[1], p[2]); o[p[3]] = (o[p[3]] || 0) + P.MV[key];
    });
    Object.keys(P.S20).forEach(function(key){
      var p = key.split('|'); if (!inPer[p[2]]) return;
      var o = dv(+p[0], p[1], p[2]); o.s20 = (o.s20 || 0) + P.S20[key]; o.s20sum = (o.s20sum || 0) + (P.S20SUM[key] || 0);
    });
    Object.keys(P.LIST).forEach(function(key){
      var p = key.split('|'); var o = dv(+p[0], p[1], p[2]); o.list = (o.list || 0) + (P.LIST[key] || 0);
    });

    /* ---- круассаны и сэндвичи: из какой партии продали со скидкой ---- */
    var KS = P.S20T.map(function(x){
      var a = mv(x.wi, x.k, x.d, 'prihod_gp') + mv(x.wi, x.k, x.d, 'prigotovlenie') + mv(x.wi, x.k, x.d, 'prihod_other');
      var old = E.lotsAt(x.wi, x.k, x.d, 1, 3).reduce(function(s, L){ return s + L.q; }, 0);
      var c = a <= 0 && old > 0 ? 'Прихода в этот день не было — продана партия прошлого дня'
            : a <= 0 ? 'Прихода в этот день не было, остатка по учёту тоже'
            : old > 0 ? 'Был приход этого дня и остаток прошлого дня' : 'Только приход этого дня (как положено)';
      return { wi: x.wi, k: x.k, d: x.d, hr: x.hr, q: x.q, cls: c, arr: a, old: old, inlist: (x.wi + '|' + x.k + '|' + x.d) in P.LIST };
    });

    /* ---- склад 50% ---- */
    var M50 = [];
    try{
      var m = match50();
      (m.rows || []).forEach(function(r){ var iso = r.iso || ''; if (iso >= from && iso <= to) M50.push(r.slice(0, 12).map(stripId)); });
    }catch(e){ M50 = null; }

    return { P: P, DISC: DISC, SPW: SPW, DAY: DAY, KS: KS, M50: M50, DATES: DATES, mv: mv };
  }

  return { prep: prep, compute: compute, fx: fx, nz: nz, pct: pct, sp_: sp_, brName: brName, sprOf: sprOf, isKs: isKs, VES: VES, cnt: cnt, g: g, days: days, rng: rng, stripId: stripId };
})();
/* ═════════ Отчёт за месяц: книга «Скидка и списание» ═════════ */
function monBookMain(R, from, to){
  var fx = MON.fx, nz = MON.nz, pct = MON.pct, sp_ = MON.sp_, VES = MON.VES;
  var WHS = R.P.WHS, BRN = WHS.map(MON.brName), NW = WHS.length;
  var DISC = R.DISC, SPW = R.SPW, DATES = R.DATES;
  function S_(k){ return MON.sprOf(k); }
  function ves(k){ return VES.test(SPR[k].name); }
  var per = ru(from) + '–' + ru(to);
  var STCOL = { 'В срок':'GRN', 'Раньше срока':'ORG', 'Позже срока':'RED', 'Нет прихода по сроку':'RED', 'Приход меньше скидки':'RED', 'Приход меньше списания':'RED',
    'Превышен остаток FIFO':'RED', 'Превышен остаток FIFO (предупреждение)':'YEL', 'Невозможно отследить':'GRY', 'Нет количества':'GRY',
    'оплачено':'GRN', 'оплачено (возможно, в другой день)':'YEL', 'не оплачено':'RED', 'есть бронь':'GRN', 'брони нет':'YEL' };
  function fill(v){ return STCOL[v] || null; }
  var BAD_D = ['Раньше срока', 'Позже срока', 'Нет прихода по сроку', 'Приход меньше скидки', 'Превышен остаток FIFO'];
  var EVAL_D = BAD_D.concat(['В срок']);
  var BAD_S = ['Раньше срока', 'Позже срока', 'Нет прихода по сроку', 'Приход меньше списания', 'Превышен остаток FIFO'];
  var EVAL_S = BAD_S.concat(['В срок']);
  var CHK_D = DISC.filter(function(r){ return r.st.indexOf('Не проверяется') !== 0; });
  var CHK_S = SPW.filter(function(r){ return r.st.indexOf('Не проверяется') !== 0; });

  /* ---- агрегаты ---- */
  var FLD = ['list', 'prihod_gp', 'prihod_other', 'prihod_tort', 'prigotovlenie', 'realiz', 's20', 's20sum', 'perem_50', 'spisanie', 'vozvrat_srok'];
  var BR = [], BRV = [], PR = {}, DY = {};
  for (var i = 0; i < NW; i++){ BR.push({}); BRV.push({}); }
  var LV_ = { sp: 0, p50: 0, s20: 0 };
  Object.keys(R.DAY).forEach(function(key){
    var v = R.DAY[key], k = v.k, wi = v.wi, price = SPR[k].price || 0;
    var loss = ((v.spisanie || 0) + (v.vozvrat_srok || 0)) * price + (v.perem_50 || 0) * price * 0.5 + (v.s20 || 0) * price * 0.2;
    LV_.sp += ((v.spisanie || 0) + (v.vozvrat_srok || 0)) * price; LV_.p50 += (v.perem_50 || 0) * price * 0.5; LV_.s20 += (v.s20 || 0) * price * 0.2;
    var tgt = ves(k) ? BRV[wi] : BR[wi], p = PR[k] || (PR[k] = {});
    FLD.forEach(function(f){ tgt[f] = (tgt[f] || 0) + (v[f] || 0); p[f] = (p[f] || 0) + (v[f] || 0); });
    BR[wi].loss = (BR[wi].loss || 0) + loss; p.loss = (p.loss || 0) + loss;
    var y = DY[v.d] || (DY[v.d] = {});
    if (!ves(k)) ['s20', 'spisanie', 'vozvrat_srok', 'perem_50', 'list'].forEach(function(f){ y[f] = (y[f] || 0) + (v[f] || 0); });
    y.loss = (y.loss || 0) + loss;
  });
  var TOT_LOSS = BR.reduce(function(s, b){ return s + (b.loss || 0); }, 0);
  function stc(rows, keyf){ var o = {}; rows.forEach(function(r){ var k = keyf(r), c = o[k] || (o[k] = {}); c[r.st] = (c[r.st] || 0) + 1; }); return o; }
  var CD_B = stc(CHK_D, function(r){ return r.wi; }), CS_B = stc(CHK_S, function(r){ return r.wi; });
  var CD_P = stc(CHK_D, function(r){ return r.k; }), CS_P = stc(CHK_S, function(r){ return r.k; });
  function G(c, s){ return (c && c[s]) || 0; }
  function sum(c){ return c ? Object.keys(c).reduce(function(a, k){ return a + c[k]; }, 0) : 0; }
  function share(c, ev){ var e = ev.reduce(function(a, s){ return a + G(c, s); }, 0); return e ? pct(G(c, 'В срок'), e) : ''; }
  var ca = {}, sa = {};
  for (i = 0; i < NW; i++){ Object.keys(CD_B[i] || {}).forEach(function(s){ ca[s] = (ca[s] || 0) + CD_B[i][s]; }); Object.keys(CS_B[i] || {}).forEach(function(s){ sa[s] = (sa[s] || 0) + CS_B[i][s]; }); }

  var wb = new XW.Book(), ST = XW.ST;
  var wsD = wb.sheet('Дашборд');
  var W = {};
  ['Сроки скидки', 'Сроки списания', 'Круассаны и сэндвичи 20%', 'Сравнение филиалов', 'Сравнение продукции', 'Скидка и списание', 'Склад 50%', 'Движение товара', 'Как считалось']
    .forEach(function(n){ W[n] = wb.sheet(n); });

  /* ---- Сроки скидки ---- */
  var hd = ['Филиал', 'Дата скидки', 'Продукция', 'Категория', 'Срок на витрине, ч', 'Срок, дней (правило уценки)', 'Ожидаемая дата прихода',
    'Приход Цех ГП в эту дату', 'Акт приготовления в эту дату', 'Перемещение из филиала в эту дату', 'В списке филиала', 'Продано со скидкой 20%',
    'Склад 50%', 'Возврат по сроку (след. день)', 'Списание (след. день)', 'Остаток партии по FIFO (как на сайте)', 'Из каких партий по FIFO (дата — шт)',
    'Статус', 'Отклонение, дней (+ позже / − раньше)', 'Пояснение', 'Весовая (кг)'];
  var rows = DISC.map(function(r){
    var s = S_(r.k), d = r.det || {};
    return [BRN[r.wi], ru(r.T), s.name, s.kat, s.srok, r.dd, r.P ? ru(r.P) : '', nz(d.gpP), nz(d.prigP), nz(d.othP),
      r.noq ? 'не указано' : fx(r.list), nz(r.s20), nz(r.p50), nz(r.vz), nz(r.sp1), (d.left !== undefined && d.left !== null) ? fx(d.left) : '', d.lots || '',
      r.st, r.dev, r.why, ves(r.k) ? 'да' : ''];
  });
  var ws = W['Сроки скидки'];
  ws.table(1, hd, rows, { widths: [14, 11, 30, 12, 9, 9, 11, 9, 9, 10, 9, 9, 8, 9, 9, 10, 22, 22, 11, 34, 7], chk: [17], fill: fill });
  ws.freeze = [4, 2]; ws.filter = 'A1:' + XW.colName(hd.length - 1) + (rows.length + 1);

  /* ---- Сроки списания ---- */
  var hs = ['Филиал', 'Дата списания', 'Продукция', 'Категория', 'Документ iiko', 'Количество', 'Срок на витрине, ч', 'Срок, дней (правило списания)', 'Ожидаемая дата прихода',
    'Приход Цех ГП в эту дату', 'Акт приготовления в эту дату', 'Перемещение из филиала в эту дату', 'Из каких партий по FIFO (дата — шт)', 'Статус', 'Отклонение, дней (+ позже / − раньше)', 'Пояснение', 'Весовая (кг)'];
  rows = SPW.map(function(r){
    var s = S_(r.k), d = r.det || {};
    return [BRN[r.wi], ru(r.T), s.name, s.kat, r.typ, fx(r.q), s.srok, r.dd, r.P ? ru(r.P) : '', nz(d.gpP), nz(d.prigP), nz(d.othP), d.lots || '', r.st, r.dev, r.why, ves(r.k) ? 'да' : ''];
  });
  ws = W['Сроки списания'];
  ws.table(1, hs, rows, { widths: [14, 11, 30, 12, 14, 9, 9, 9, 11, 9, 9, 10, 22, 22, 11, 34, 7], chk: [13], fill: fill });
  ws.freeze = [4, 2]; ws.filter = 'A1:' + XW.colName(hs.length - 1) + (rows.length + 1);

  /* ---- Сравнение филиалов ---- */
  ws = W['Сравнение филиалов'];
  var hb = ['Филиал', 'Приход Цех ГП, шт', 'Перемещения из филиалов, шт', 'Приход Таугуль торт, шт', 'Продажи всего (iiko), шт', 'В списках на скидку, шт', 'Продано со скидкой 20%, шт',
    'Сумма продаж со скидкой 20% (по полной цене), ₸', 'Склад 50%, шт', 'Списание, шт', 'Возврат по сроку, шт', 'Скидка 20% от прихода ГП, %', 'Списание + возврат от прихода ГП, %', 'Потери выручки, ₸', 'Доля в потерях сети, %'];
  var tot = {};
  rows = BR.map(function(b, i){
    FLD.concat(['loss']).forEach(function(f){ tot[f] = (tot[f] || 0) + (b[f] || 0); });
    var sp = (b.spisanie || 0) + (b.vozvrat_srok || 0);
    return [BRN[i], fx(b.prihod_gp || 0), fx(b.prihod_other || 0), fx(b.prihod_tort || 0), fx(b.realiz || 0), fx(b.list || 0), fx(b.s20 || 0), fx(b.s20sum || 0), fx(b.perem_50 || 0),
      fx(b.spisanie || 0), fx(b.vozvrat_srok || 0), pct(b.s20 || 0, b.prihod_gp || 0), pct(sp, b.prihod_gp || 0), fx(b.loss || 0), pct(b.loss || 0, TOT_LOSS)];
  });
  var T = ['Итого (шт.)', fx(tot.prihod_gp || 0), fx(tot.prihod_other || 0), fx(tot.prihod_tort || 0), fx(tot.realiz || 0), fx(tot.list || 0), fx(tot.s20 || 0), fx(tot.s20sum || 0), fx(tot.perem_50 || 0),
    fx(tot.spisanie || 0), fx(tot.vozvrat_srok || 0), pct(tot.s20 || 0, tot.prihod_gp || 0), pct((tot.spisanie || 0) + (tot.vozvrat_srok || 0), tot.prihod_gp || 0), fx(tot.loss || 0), 100];
  var re = ws.table(1, hb, rows, { widths: [15, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12], title: 'Объёмы за ' + per + ' (штучная продукция; весовая — в таблице ниже, в кг; потери ₸ — по всей продукции)', total: T });
  rows = BRV.map(function(b, i){ return [BRN[i], fx(b.prihod_gp || 0), fx(b.prihod_other || 0), fx(b.realiz || 0), fx(b.s20 || 0), fx(b.perem_50 || 0), fx(b.spisanie || 0), fx(b.vozvrat_srok || 0)]; });
  re = ws.table(re[1] + 3, ['Филиал', 'Приход Цех ГП, кг', 'Перемещения из филиалов, кг', 'Продажи всего, кг', 'Продано со скидкой 20%, кг', 'Склад 50%, кг', 'Списание, кг', 'Возврат по сроку, кг'], rows,
    { title: 'Весовая продукция (Печенье Вупи, Бауырсақ, рогалик, көкнәр), кг' });
  var hq = ['Филиал', 'Скидка: проверено строк', 'В срок', 'Раньше срока', 'Позже срока', 'Нет прихода по сроку', 'Приход меньше скидки', 'Превышен остаток FIFO', 'Превышен остаток (предупр.)', 'Невозможно отследить',
    'Нет количества', 'Скидка в срок, % (от оценённых)', 'Списание: проверено строк', 'В срок', 'Раньше срока', 'Позже срока', 'Нет прихода по сроку', 'Превышен остаток FIFO', 'Невозможно отследить', 'Списание в срок, % (от оценённых)', 'Место по скидке в срок'];
  rows = BRN.map(function(b, i){
    var c = CD_B[i] || {}, s = CS_B[i] || {};
    return [b, sum(c), G(c, 'В срок'), G(c, 'Раньше срока'), G(c, 'Позже срока'), G(c, 'Нет прихода по сроку'), G(c, 'Приход меньше скидки'), G(c, 'Превышен остаток FIFO'),
      G(c, 'Превышен остаток FIFO (предупреждение)'), G(c, 'Невозможно отследить'), G(c, 'Нет количества'), share(c, EVAL_D),
      sum(s), G(s, 'В срок'), G(s, 'Раньше срока'), G(s, 'Позже срока'), G(s, 'Нет прихода по сроку'), G(s, 'Превышен остаток FIFO'), G(s, 'Невозможно отследить'), share(s, EVAL_S), 0];
  });
  rows.map(function(r, i){ return i; }).sort(function(a, b){ return (rows[b][11] || 0) - (rows[a][11] || 0); }).forEach(function(i, p){ rows[i][20] = p + 1; });
  T = ['Итого', sum(ca), G(ca, 'В срок'), G(ca, 'Раньше срока'), G(ca, 'Позже срока'), G(ca, 'Нет прихода по сроку'), G(ca, 'Приход меньше скидки'), G(ca, 'Превышен остаток FIFO'),
    G(ca, 'Превышен остаток FIFO (предупреждение)'), G(ca, 'Невозможно отследить'), G(ca, 'Нет количества'), share(ca, EVAL_D),
    sum(sa), G(sa, 'В срок'), G(sa, 'Раньше срока'), G(sa, 'Позже срока'), G(sa, 'Нет прихода по сроку'), G(sa, 'Превышен остаток FIFO'), G(sa, 'Невозможно отследить'), share(sa, EVAL_S), ''];
  var bq = ws.table(re[1] + 3, hq, rows, { title: 'Сроки: число строк по статусам (строка = филиал × продукция × дата)', total: T });
  ws.freeze = [2, 3];
  ws.chart({ kind: 'col', grouping: 'stacked', title: 'Скидка 20%: статусы сроков по филиалам', colors: ['70AD47', 'ED7D31', 'C00000', 'FF0000', 'FFC000', '7F7F7F'], cat: [1, bq[0] + 1, bq[0] + NW],
    series: [3, 4, 5, 6, 7, 8].map(function(c){ return { nameCell: [c, bq[0]], val: [c, bq[0] + 1, bq[0] + NW] }; }), at: [2, bq[1] + 3], size: [12, 20] });

  /* ---- Сравнение продукции ---- */
  ws = W['Сравнение продукции'];
  var hp = ['Продукция', 'Категория', 'Ед.', 'Срок на витрине, ч', 'Форма хранения', 'Цена, ₸', 'Приход Цех ГП', 'Перемещения из филиалов', 'Продажи всего (iiko)', 'В списках на скидку', 'Продано со скидкой 20%',
    'Склад 50%', 'Списание', 'Возврат по сроку', 'Скидка 20% от прихода ГП, %', 'Списание + возврат от прихода ГП, %', 'Потери выручки, ₸',
    'Скидка: строк проверено', 'Скидка: в срок', 'Скидка: раньше', 'Скидка: позже', 'Скидка: нет прихода / меньше / превышен остаток', 'Скидка: невозможно отследить', 'Скидка в срок, %',
    'Списание: строк', 'Списание: в срок', 'Списание: раньше', 'Списание: позже', 'Списание: прочие нарушения', 'Списание: невозможно отследить'];
  var PROD = [];
  Object.keys(PR).forEach(function(k){
    var p = PR[k];
    if (!['prihod_gp', 'realiz', 's20', 'spisanie', 'vozvrat_srok', 'perem_50', 'list'].some(function(f){ return p[f]; })) return;
    var s = S_(k), c = CD_P[k] || {}, q = CS_P[k] || {}, g = function(f){ return p[f] || 0; };
    PROD.push([s.name, s.kat, ves(k) ? 'кг' : 'шт', s.srok, s.forma, s.price, fx(g('prihod_gp')), fx(g('prihod_other')), fx(g('realiz')), fx(g('list')), fx(g('s20')),
      fx(g('perem_50')), fx(g('spisanie')), fx(g('vozvrat_srok')), pct(g('s20'), g('prihod_gp')), pct(g('spisanie') + g('vozvrat_srok'), g('prihod_gp')), fx(g('loss')),
      sum(c), G(c, 'В срок'), G(c, 'Раньше срока'), G(c, 'Позже срока'), G(c, 'Нет прихода по сроку') + G(c, 'Приход меньше скидки') + G(c, 'Превышен остаток FIFO'), G(c, 'Невозможно отследить'),
      share(c, EVAL_D), sum(q), G(q, 'В срок'), G(q, 'Раньше срока'), G(q, 'Позже срока'), G(q, 'Нет прихода по сроку') + G(q, 'Превышен остаток FIFO'), G(q, 'Невозможно отследить')]);
  });
  PROD.sort(function(a, b){ return (b[16] || 0) - (a[16] || 0); });
  re = ws.table(1, hp, PROD, { widths: [30, 12, 5, 9, 12, 8].concat(new Array(24).fill(10)), title: 'Продукция за ' + per + ', по всем филиалам (сортировка — по потерям выручки)' });
  ws.freeze = [2, 3]; ws.filter = 'A2:' + XW.colName(hp.length - 1) + re[1];
  if (PROD.length) ws.chart({ kind: 'bar', title: 'Топ-15 продукции по потерям выручки, ₸', legend: false, cat: [1, 3, 2 + Math.min(15, PROD.length)],
    series: [{ nameCell: [17, 2], val: [17, 3, 2 + Math.min(15, PROD.length)] }], at: [32, 3], size: [10, 22] });

  /* ---- Скидка и списание ---- */
  ws = W['Скидка и списание'];
  var LV = {};
  function lv(wi, k){ var key = wi + '|' + k; return LV[key] || (LV[key] = { wi: wi, k: k, days: 0, list: 0, s20: 0, p50: 0, vz: 0, sp1: 0, bad: 0, nld: 0, nlq: 0 }); }
  DISC.forEach(function(r){ var o = lv(r.wi, r.k); o.days++; o.list += r.list || 0; o.s20 += r.s20; o.p50 += r.p50; o.vz += r.vz; o.sp1 += r.sp1; if (BAD_D.indexOf(r.st) >= 0) o.bad++; });
  Object.keys(R.DAY).forEach(function(key){ var v = R.DAY[key]; if (v.s20 && !((v.wi + '|' + v.k + '|' + v.d) in R.P.LIST)){ var o = lv(v.wi, v.k); o.nld++; o.nlq += v.s20; } });
  var hl = ['Филиал', 'Продукция', 'Ед.', 'Дней в списке на скидку', 'Заявлено на скидку', 'Продано со скидкой 20% (дни списка)', 'Склад 50% (дни списка)', 'Возврат по сроку (след. день)', 'Списание (след. день)',
    'Ушло по факту всего', 'Разница: список − факт', 'Дней с нарушением срока', 'Дней продаж 20% без списка', 'Продано 20% без списка', 'Списание за период (iiko)', 'Возврат по сроку за период (iiko)'];
  rows = Object.keys(LV).map(function(key){
    var o = LV[key], fact = o.s20 + o.p50 + o.vz + o.sp1, spm = 0, vzm = 0;
    DATES.forEach(function(d){ spm += R.mv(o.wi, o.k, d, 'spisanie'); vzm += R.mv(o.wi, o.k, d, 'vozvrat_srok'); });
    return [BRN[o.wi], SPR[o.k].name, ves(o.k) ? 'кг' : 'шт', o.days, fx(o.list), fx(o.s20), fx(o.p50), fx(o.vz), fx(o.sp1), fx(fact), o.days ? fx(o.list - fact) : '',
      o.bad, o.nld, fx(o.nlq), fx(spm), fx(vzm)];
  }).sort(function(a, b){ return BRN.indexOf(a[0]) - BRN.indexOf(b[0]) || (a[1] < b[1] ? -1 : 1); });
  ws.table(1, hl, rows, { widths: [14, 30, 5].concat(new Array(13).fill(11)), title: 'Список на скидку и что ушло по факту — по филиалу и продукции, ' + per });
  ws.freeze = [3, 3]; ws.filter = 'A2:' + XW.colName(hl.length - 1) + (rows.length + 2);

  /* ---- Склад 50% ---- */
  ws = W['Склад 50%'];
  var S50 = { rows: 0, paid: 0, nopay: 0, bron: 0 };
  if (R.M50){
    var h5 = ['Дата накладной', 'Филиал', 'Позиция', 'Кол-во в накладной', 'Бронь №', 'Бронь — ФИО', 'Бронь — статус', 'Покупатель (ФИО)', 'Дата оплаты', 'Оплачено, шт', 'Статус оплаты', 'Совпадение с бронью', 'Примечание сайта', 'Подсказка (недостача в филиале)'];
    rows = R.M50.map(function(r){
      var paid = ['—', '', null, undefined].indexOf(r[7]) < 0;
      var stp = paid ? (/другой день/.test(S(r[9])) ? 'оплачено (возможно, в другой день)' : 'оплачено') : 'не оплачено';
      var wi = WHS.indexOf(r[1]);
      return [r[0], wi >= 0 ? BRN[wi] : r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], stp, r[4] ? 'есть бронь' : 'брони нет', r[10], r[11]];
    });
    re = ws.table(1, h5, rows, { widths: [11, 14, 28, 9, 9, 22, 10, 26, 11, 9, 16, 13, 34, 30], chk: [10, 11], fill: fill, title: 'Склад 50% за ' + per + ': накладная → покупатель → бронь (сопоставление сайта)' });
    ws.freeze = [4, 3];
    var c5 = {}, b5 = {}; rows.forEach(function(r){ c5[r[10]] = (c5[r[10]] || 0) + 1; b5[r[11]] = (b5[r[11]] || 0) + 1; });
    S50 = { rows: rows.length, paid: G(c5, 'оплачено') + G(c5, 'оплачено (возможно, в другой день)'), nopay: G(c5, 'не оплачено'), bron: G(b5, 'есть бронь') };
    ws.table(re[1] + 3, ['Показатель', 'Значение'], [['Строк', rows.length], ['Оплачено (строк)', G(c5, 'оплачено')], ['Оплачено, возможно в другой день (строк)', G(c5, 'оплачено (возможно, в другой день)')],
      ['Не оплачено (строк)', G(c5, 'не оплачено')], ['С бронью (строк)', G(b5, 'есть бронь')], ['Без брони (строк)', G(b5, 'брони нет')],
      ['Склад 50% по движениям iiko за период, шт (все филиалы)', fx((tot.perem_50 || 0) + BRV.reduce(function(s, b){ return s + (b.perem_50 || 0); }, 0))]], { title: 'Итого по складу 50%' });
  } else ws.text(1, 1, 'Склад 50%: сопоставление не посчиталось — загрузите «Продажа 50%» и откройте «Путь на склад 50%».');

  /* ---- Движение товара ---- */
  ws = W['Движение товара'];
  var hm = ['Дата', 'Филиал', 'Продукция', 'Категория', 'Ед.', 'В списке на скидку', 'Приход Цех ГП', 'Перемещение из филиала', 'Приход Таугуль торт', 'Акт приготовления',
    'Продажи всего (iiko)', 'Продано со скидкой 20%', 'Сумма продаж со скидкой 20% (по полной цене), ₸', 'Склад 50%', 'Списание', 'Возврат по сроку', 'Возврат по качеству', 'Перемещение в другой филиал',
    'Доставка', 'Прочая расходная накладная', 'Инвентаризация: излишек', 'Инвентаризация: недостача'];
  var MOVR = Object.keys(R.DAY).map(function(k){ return R.DAY[k]; })
    .sort(function(a, b){ return a.d < b.d ? -1 : a.d > b.d ? 1 : (a.wi - b.wi || (SPR[a.k].name < SPR[b.k].name ? -1 : 1)); })
    .filter(function(v){ return Object.keys(v).some(function(f){ return ['wi', 'k', 'd'].indexOf(f) < 0 && Math.abs(v[f]) > 1e-9; }); })
    .map(function(v){
      return [ru(v.d), BRN[v.wi], SPR[v.k].name, SPR[v.k].kat, ves(v.k) ? 'кг' : 'шт', nz(v.list), nz(v.prihod_gp), nz(v.prihod_other), nz(v.prihod_tort), nz(v.prigotovlenie),
        nz(v.realiz), nz(v.s20), nz(v.s20sum), nz(v.perem_50), nz(v.spisanie), nz(v.vozvrat_srok), nz(v.vozvrat_kach), nz(v.perem_out), nz(v.perem_dost), nz(v.other), nz(v.inv_in), nz(v.inv_out)];
    });
  ws.table(1, hm, MOVR, { widths: [11, 14, 30, 12, 5].concat(new Array(17).fill(10)) });
  ws.freeze = [4, 2]; ws.filter = 'A1:' + XW.colName(hm.length - 1) + (MOVR.length + 1);

  /* ---- Круассаны и сэндвичи 20% ---- */
  ws = W['Круассаны и сэндвичи 20%'];
  var KS = R.KS, CL = ['Только приход этого дня (как положено)', 'Был приход этого дня и остаток прошлого дня', 'Прихода в этот день не было — продана партия прошлого дня', 'Прихода в этот день не было, остатка по учёту тоже'];
  var CLF = {}; CLF[CL[0]] = 'GRN'; CLF[CL[1]] = 'YEL'; CLF[CL[2]] = 'RED'; CLF[CL[3]] = 'GRY';
  var tq = KS.reduce(function(s, x){ return s + x.q; }, 0);
  ws.text(1, 1, 'Круассаны и сэндвичи: продажи со скидкой 20% по времени чека, ' + per, ST.title);
  ws.text(2, 1, 'Правило: на следующий день после прихода их списывают; со скидкой — только вечером в день прихода. Время — из «Продаж по чекам» на сайте, приходы и остаток — iiko (FIFO по дням).');
  function qs(f){ return fx(KS.filter(f).reduce(function(s, x){ return s + x.q; }, 0)); }
  var sm = CL.map(function(c){ var q = qs(function(x){ return x.cls === c; }); return [c, q, pct(q, tq), qs(function(x){ return x.cls === c && x.hr < 16; }), qs(function(x){ return x.cls === c && x.hr >= 16; })]; });
  re = ws.table(4, ['Из какой партии продано со скидкой', 'Продано со скидкой 20%, шт', 'Доля, %', 'До 16:00, шт', 'С 16:00, шт'], sm,
    { widths: [58, 14, 10, 12, 12], chk: [0], fill: function(v){ return CLF[v]; }, title: 'Итого', total: ['Итого', fx(tq), tq ? 100 : '', qs(function(x){ return x.hr < 16; }), qs(function(x){ return x.hr >= 16; })] });
  var hrs = []; KS.forEach(function(x){ if (hrs.indexOf(x.hr) < 0) hrs.push(x.hr); }); hrs.sort(function(a, b){ return a - b; });
  var p2 = function(h){ return h < 0 ? 'без времени' : (h < 10 ? '0' : '') + h + ':00–' + (h < 10 ? '0' : '') + h + ':59'; };
  var hrr = hrs.map(function(h){ return [p2(h)].concat(CL.map(function(c){ return qs(function(x){ return x.hr === h && x.cls === c; }); })).concat([qs(function(x){ return x.hr === h; })]); });
  var hr0 = ws.table(re[1] + 3, ['Час чека'].concat(CL).concat(['Всего']), hrr, { title: 'По часам' });
  if (hrr.length) ws.chart({ kind: 'col', grouping: 'stacked', title: 'Продажи со скидкой 20% по часам, шт', colors: ['70AD47', 'FFC000', 'C00000', 'A5A5A5'], cat: [1, hr0[0] + 1, hr0[1]],
    series: [2, 3, 4, 5].map(function(c){ return { nameCell: [c, hr0[0]], val: [c, hr0[0] + 1, hr0[1]] }; }), at: [8, hr0[0] - 1], size: [10, 18] });
  var brr = BRN.map(function(b, i){ return [b].concat(CL.map(function(c){ return qs(function(x){ return x.wi === i && x.cls === c; }); })).concat([qs(function(x){ return x.wi === i; })]); });
  re = ws.table(Math.max(hr0[1], hr0[0] + 18) + 3, ['Филиал'].concat(CL).concat(['Всего']), brr, { title: 'По филиалам' });
  var det = KS.slice().sort(function(a, b){ return a.wi - b.wi || (a.d < b.d ? -1 : a.d > b.d ? 1 : a.hr - b.hr); })
    .map(function(x){ return [BRN[x.wi], ru(x.d), p2(x.hr), SPR[x.k].name, fx(x.q), nz(x.arr), nz(x.old), x.inlist ? 'да' : '', x.cls]; });
  ws.table(re[1] + 3, ['Филиал', 'Дата', 'Час чека', 'Продукция', 'Кол-во', 'Приход в этот день (ГП / акт / перемещение)', 'Остаток прошлых дней к утру (FIFO)', 'Есть в списке филиала', 'Из какой партии'],
    det, { widths: [16, 11, 13, 28, 8, 14, 14, 10, 50], chk: [8], fill: function(v){ return CLF[v]; }, title: 'Все чеки со скидкой 20%' });

  /* ---- проверки ---- */
  var QC = [];
  function ok(a, b){ return Math.abs((+a || 0) - (+b || 0)) < 1e-6 ? 'совпадает' : 'РАСХОЖДЕНИЕ'; }
  var nn = 0, mm = 0;
  QC.push(['Сумма строк по статусам в «Сравнении филиалов» + «не проверяется» = строк на листе «Сроки скидки»', sum(ca) + DISC.filter(function(r){ return r.st.indexOf('Не проверяется') === 0; }).length, DISC.length]);
  QC.push(['Сумма строк по статусам в «Сравнении филиалов» + «не проверяется» = строк на листе «Сроки списания»', sum(sa) + SPW.filter(function(r){ return r.st.indexOf('Не проверяется') === 0; }).length, SPW.length]);
  var mt = {};
  MOVR.forEach(function(r){ if (r[4] === 'шт') [[6, 'prihod_gp'], [11, 's20'], [13, 'perem_50'], [14, 'spisanie'], [15, 'vozvrat_srok']].forEach(function(p){ mt[p[1]] = (mt[p[1]] || 0) + (+r[p[0]] || 0); }); });
  [['prihod_gp', 'Приход Цех ГП'], ['s20', 'Продано со скидкой 20%'], ['perem_50', 'Склад 50%'], ['spisanie', 'Списание'], ['vozvrat_srok', 'Возврат по сроку']].forEach(function(p){
    QC.push(['Итог «' + p[1] + '» (шт) в «Сравнении филиалов» = сумма по «Движению товара»', fx(tot[p[0]] || 0), fx(mt[p[0]] || 0)]); });
  var pl = PROD.reduce(function(s, r){ return s + (r[16] || 0); }, 0);
  QC.push(['Потери выручки: сумма по продукции = сумма по филиалам', fx(pl), fx(TOT_LOSS)]);
  QC = QC.map(function(q){ return q.concat([ok(q[1], q[2])]); });

  /* ---- Дашборд ---- */
  ws = wsD;
  var evD = EVAL_D.reduce(function(a, s){ return a + G(ca, s); }, 0), evS = EVAL_S.reduce(function(a, s){ return a + G(sa, s); }, 0);
  var bp = BRN.map(function(b, i){ return [b, share(CD_B[i] || {}, EVAL_D), i]; }).filter(function(x){ return x[1] !== ''; }).sort(function(a, b){ return b[1] - a[1]; });
  var lossb = BRN.map(function(b, i){ return [b, BR[i].loss || 0]; }).sort(function(a, b){ return b[1] - a[1]; });
  function top(st, rr, n){ var c = {}; rr.forEach(function(r){ if (r.st === st) c[SPR[r.k].name] = (c[SPR[r.k].name] || 0) + 1; });
    return Object.keys(c).sort(function(a, b){ return c[b] - c[a]; }).slice(0, n).map(function(k){ return k + ' (' + c[k] + ')'; }).join(', ') || '—'; }
  var dE = DISC.filter(function(r){ return r.st === 'Раньше срока'; }), dL = DISC.filter(function(r){ return r.st === 'Позже срока'; });
  var ksL = dL.filter(function(r){ return MON.isKs(r.k); }).length;
  var nol = 0, inl = 0, nolD = 0, nolN = 0;
  Object.keys(R.DAY).forEach(function(key){ var v = R.DAY[key]; if (ves(v.k) || !v.s20) return;
    if ((v.wi + '|' + v.k + '|' + v.d) in R.P.LIST) inl += v.s20; else { nol += v.s20; if (dualSrok(SPR[v.k])) nolD += v.s20; if (SPR[v.k].skidka === 'Нет') nolN += v.s20; } });
  var P_ = function(a, b){ return b ? sp_(pct(a, b)) : '—'; };
  var CONC = [
    'Скидка 20% по срокам: из ' + sp_(evD) + ' строк «филиал × продукция × дата», где срок удалось оценить, в срок — ' + sp_(G(ca, 'В срок')) + ' (' + P_(G(ca, 'В срок'), evD) + '%). Раньше срока — ' +
      G(ca, 'Раньше срока') + ' (' + P_(G(ca, 'Раньше срока'), evD) + '%), позже срока — ' + G(ca, 'Позже срока') + ' (' + P_(G(ca, 'Позже срока'), evD) + '%), превышен остаток FIFO — ' +
      G(ca, 'Превышен остаток FIFO') + ', нет прихода по сроку — ' + G(ca, 'Нет прихода по сроку') + ', приход меньше скидки — ' + G(ca, 'Приход меньше скидки') + '.',
    bp.length ? 'Филиалы: лучше всех по срокам скидки ' + bp[0][0] + ' (' + sp_(bp[0][1]) + '% в срок)' + (bp[1] ? ' и ' + bp[1][0] + ' (' + sp_(bp[1][1]) + '%)' : '') + '; хуже всех ' +
      bp[bp.length - 1][0] + ' (' + sp_(bp[bp.length - 1][1]) + '%): у него ' + G(CD_B[bp[bp.length - 1][2]], 'Раньше срока') + ' случаев «раньше срока» и ' + G(CD_B[bp[bp.length - 1][2]], 'Превышен остаток FIFO') + ' — «превышен остаток FIFO».' : 'Филиалы: строк для оценки нет.',
    'Отклонения: «раньше срока» — ' + dE.length + ' (из них на 1 день — ' + dE.filter(function(r){ return r.dev === -1; }).length + '); «позже срока» — ' + dL.length + ' (на 1 день — ' +
      dL.filter(function(r){ return r.dev === 1; }).length + '). Чаще всего раньше срока: ' + top('Раньше срока', DISC, 4) + '.' + (ksL ? ' Из «позже срока» ' + ksL + ' — круассаны и сэндвичи: в списках на скидку у них остаток прошлого дня, а его нужно списывать.' : ''),
    'Списание (документы iiko «Списание» и «Возврат по сроку»): из ' + evS + ' оценённых строк в срок — ' + G(sa, 'В срок') + ' (' + P_(G(sa, 'В срок'), evS) + '%); раньше срока — ' + G(sa, 'Раньше срока') +
      ' (' + top('Раньше срока', SPW, 3) + '), позже — ' + G(sa, 'Позже срока') + ', превышен остаток FIFO — ' + G(sa, 'Превышен остаток FIFO') + ', нет прихода по сроку — ' + G(sa, 'Нет прихода по сроку') + '.',
    'Потери выручки — ' + sp_(fx(TOT_LOSS)) + ' ₸: скидка 20% — ' + sp_(fx(LV_.s20)) + ' ₸, списание и возврат по сроку — ' + sp_(fx(LV_.sp)) + ' ₸, склад 50% — ' + sp_(fx(LV_.p50)) + ' ₸. Больше всего у ' +
      lossb[0][0] + ' (' + sp_(fx(lossb[0][1])) + ' ₸)' + (lossb[1] ? ' и ' + lossb[1][0] + ' (' + sp_(fx(lossb[1][1])) + ' ₸)' : '') + (PROD[0] ? '; по продукции — ' + PROD[0][0] + ' (' + sp_(PROD[0][16]) + ' ₸)' : '') + '.',
    'Продажи со скидкой 20% без списка филиала: ' + sp_(fx(nol)) + ' шт штучной продукции продано по 20% в дни, когда позиции не было в списке (в дни списка — ' + sp_(fx(inl)) + ' шт). Из них ' +
      sp_(fx(nolD)) + ' шт — позиции с обеими формами хранения, ' + sp_(fx(nolN)) + ' шт — позиции, у которых в Справочнике скидка «Нет».',
    'Склад 50%: ' + S50.rows + ' строк «накладная → покупатель → бронь»; оплата найдена в ' + S50.paid + ', не найдена в ' + S50.nopay + '; бронь есть в ' + S50.bron + '. «Невозможно отследить» — ' +
      G(ca, 'Невозможно отследить') + ' строк скидки (' + DISC.filter(function(r){ return r.why === 'обе формы хранения'; }).length + ' из них — обе формы хранения).'
  ];
  ws.grid = false;
  [[1, 4], [2, 26]].forEach(function(p){ ws.width(p[0], p[1]); }); for (i = 3; i <= 14; i++) ws.width(i, 14);
  ws.text(1, 2, 'Скидка и списание — ' + per + ' (' + NW + ' филиалов)', ST.title);
  ws.text(2, 2, 'Главный акцент — сроки: уценка и списание сверены с датой прихода по FIFO. Собрано на сайте сверки ' + ru(new Date().toISOString().slice(0, 10)) + '; правила — на листе «Как считалось».', ST.def);
  ws.text(4, 2, 'Выводы для руководителя', ST.h2);
  var r = 5;
  CONC.forEach(function(t, j){ ws.text(r, 2, (j + 1) + '. ' + t); ws.merge(r, 2, r, 13); ws.heights[r] = 44; r++; });
  r++; ws.text(r, 2, 'Ключевые показатели', ST.h2); r++;
  var KPI = [['Строк скидки (филиал × продукция × дата)', DISC.length], ['Скидка: в срок, % от оценённых', evD ? pct(G(ca, 'В срок'), evD) : ''], ['Скидка: раньше срока, строк', G(ca, 'Раньше срока')],
    ['Скидка: позже срока, строк', G(ca, 'Позже срока')], ['Скидка: превышен остаток FIFO, строк', G(ca, 'Превышен остаток FIFO')], ['Скидка: невозможно отследить, строк', G(ca, 'Невозможно отследить')],
    ['Строк списания (iiko)', SPW.length], ['Списание: в срок, % от оценённых', evS ? pct(G(sa, 'В срок'), evS) : ''], ['Продано со скидкой 20%, шт', fx(tot.s20 || 0)], ['Сумма продаж со скидкой 20% (по полной цене), ₸', fx(tot.s20sum || 0)],
    ['Списание + возврат по сроку, шт', fx((tot.spisanie || 0) + (tot.vozvrat_srok || 0))], ['Склад 50%, шт', fx(tot.perem_50 || 0)], ['Потери выручки, ₸', fx(TOT_LOSS)]];
  var k0 = r;
  KPI.forEach(function(p, j){ var rr = k0 + Math.floor(j / 2), cc = 2 + (j % 2) * 6;
    ws.set(rr, cc, p[0], ST.BLU); ws.merge(rr, cc, rr, cc + 3); ws.set(rr, cc + 4, p[1], ST.bold); ws.merge(rr, cc + 4, rr, cc + 5); });
  r = k0 + Math.ceil(KPI.length / 2) + 2;
  var rk = bp.map(function(x, p){ var c = CD_B[x[2]] || {}; return [p + 1, x[0], x[1], G(c, 'Раньше срока'), G(c, 'Позже срока'), G(c, 'Превышен остаток FIFO'), G(c, 'Нет прихода по сроку') + G(c, 'Приход меньше скидки'), share(CS_B[x[2]] || {}, EVAL_S), fx(BR[x[2]].loss || 0)]; });
  var rkr = ws.table(r, ['Место', 'Филиал', 'Скидка в срок, %', 'Раньше срока', 'Позже срока', 'Превышен остаток FIFO', 'Нет прихода / приход меньше', 'Списание в срок, %', 'Потери выручки, ₸'], rk,
    { c0: 2, title: 'Рейтинг филиалов по срокам скидки (1 — лучший)' });
  if (rk.length){
    ws.chart({ kind: 'bar', title: 'Скидка в срок, % (по филиалам)', legend: false, cat: [3, rkr[0] + 1, rkr[1]], series: [{ nameCell: [4, rkr[0]], val: [4, rkr[0] + 1, rkr[1]] }], at: [12, rkr[0] - 1], size: [7, 15] });
    ws.chart({ kind: 'col', title: 'Потери выручки по филиалам, ₸', legend: false, cat: [3, rkr[0] + 1, rkr[1]], series: [{ nameCell: [10, rkr[0]], val: [10, rkr[0] + 1, rkr[1]] }], at: [12, rkr[1] + 6], size: [7, 15] });
  }
  var SS = ['В срок', 'Раньше срока', 'Позже срока', 'Нет прихода по сроку', 'Приход меньше скидки', 'Превышен остаток FIFO'];
  rows = BRN.map(function(b, i){ return [b].concat(SS.map(function(s){ return G(CD_B[i], s); })); });
  var sr = ws.table(rkr[1] + 3, ['Филиал'].concat(SS), rows, { c0: 2, title: 'Скидка 20%: строки по статусам сроков', total: ['Итого'].concat(SS.map(function(s){ return G(ca, s); })) });
  ws.chart({ kind: 'col', grouping: 'percentStacked', title: 'Скидка 20%: структура статусов по филиалам', colors: ['70AD47', 'ED7D31', 'C00000', 'FF0000', 'FFC000', '7F7F7F'], cat: [2, sr[0] + 1, sr[0] + NW],
    series: [3, 4, 5, 6, 7, 8].map(function(c){ return { nameCell: [c, sr[0]], val: [c, sr[0] + 1, sr[0] + NW] }; }), at: [2, sr[1] + 3], size: [11, 18] });
  r = sr[1] + 22;
  var tp = PROD.slice(0, 10).map(function(p, j){ return [j + 1, p[0], p[2], p[10], p[12], p[13], p[16]]; });
  var tr = ws.table(r, ['Место', 'Продукция', 'Ед.', 'Продано со скидкой 20%', 'Списание', 'Возврат по сроку', 'Потери выручки, ₸'], tp, { c0: 2, title: 'Топ-10 продукции по потерям выручки' });
  var dr = DATES.map(function(d){ var y = DY[d] || {}; return [ru(d), fx(y.s20 || 0), fx((y.spisanie || 0) + (y.vozvrat_srok || 0)), fx(y.perem_50 || 0), fx(y.list || 0), fx(y.loss || 0)]; });
  var dd2 = ws.table(tr[1] + 3, ['Дата', 'Продано со скидкой 20%, шт', 'Списание + возврат по сроку, шт', 'Склад 50%, шт', 'В списках на скидку, шт', 'Потери выручки, ₸'], dr,
    { c0: 2, title: 'По дням (все филиалы, штучная продукция; потери — вся)' });
  ws.chart({ kind: 'line', title: 'По дням: скидка 20%, список и списание, шт', cat: [2, dd2[0] + 1, dd2[1]],
    series: [3, 4, 6].map(function(c){ return { nameCell: [c, dd2[0]], val: [c, dd2[0] + 1, dd2[1]] }; }), at: [9, dd2[0]], size: [9, 18] });

  /* ---- Как считалось ---- */
  ws = W['Как считалось'];
  var TXT = [
    ['h', 'Как считался отчёт «Скидка и списание» за ' + per],
    ['', 'Филиалы: ' + BRN.join(', ') + '.'],
    ['h2', 'Источники'],
    ['', 'Сайт сверки: загруженные выгрузки iiko (товарная ведомость, товарный отчёт, продажи со скидкой 20% с временем чека, «Продажа 50%»), списки филиалов из бота («Списки филиалов»), Справочник (срок, форма хранения, цена, участие в скидке), брони сотрудников. Если выгрузки за часть периода не загружены, этих дней в отчёте нет.'],
    ['h2', 'Правила сроков'],
    ['', 'Срок — «Срок на витрине, ч» из Справочника.'],
    ['', 'Скидка 20%: срок в днях = ч / 24, но не меньше 2 (24 ч = 2 дня, список шлют накануне вечером). Партия для уценки в день T должна прийти в день T − (дни − 1). Пример: 72 ч, уценка 05.10 → приход 03.10.'],
    ['', 'Списание: приход = дата списания − срок (72 ч → 3 дня, 48 ч → 2, 24 ч → 1). Пример: 72 ч, приход 01.10 → списание 04.10. Проверяются документы iiko «Списание» и «Возврат по сроку».'],
    ['', '«Приход Цех ГП» — только поставки с Цеха. Перемещения из других филиалов приходом по сроку НЕ считаются: если в ожидаемую дату было перемещение или по FIFO в уценку/списание попадает партия с перемещением — «Невозможно отследить».'],
    ['', 'Обе формы хранения и заморозка — «Невозможно отследить». Заморозка — как на сайте: если позиция хоть раз приходила в этот филиал с Таугуль торт склада (или приходила оттуда в пределах срока), вся позиция в филиале не проверяется.'],
    ['', 'Кусочек торта = 1/8 целого и считается вместе со своим тортом. «Таңқурайлы торт кусочек» = кусочек «Таңқурай нәзіктігі торты».'],
    ['', 'Пицца не проверяется (её собирают в филиале).'],
    ['', 'Круассаны и сэндвичи: срок 24 ч. На следующий день после прихода их списывают; со скидкой 20% можно продать только вечером в день прихода. Для скидки ожидаемая дата прихода = день скидки, для списания = день списания − 1. Приход для них — Цех ГП, акт приготовления или перемещение из филиала, который их готовит (ТТ6, ТТ11). Лист «Круассаны и сэндвичи 20%» — по времени чеков.'],
    ['', 'Весовая продукция (Печенье Вупи, Бауырсақ, рогалик, көкнәр) — в кг, в штучные итоги не входит, показана отдельной таблицей; в «Потерях выручки» учтена.'],
    ['', 'Количества и суммы не округлялись (убрана только погрешность вычислений после 6-го знака). Доли в процентах — с двумя знаками после запятой.'],
    ['h2', 'FIFO и статусы'],
    ['', 'FIFO по дням: каждый приход (Цех ГП, перемещение, Таугуль, акт приготовления, излишек инвентаризации) заводит партию; любой расход (продажа, списание, возвраты, склад 50%, перемещение, доставка, прочая расходная накладная, недостача) съедает самую старую. Партия считается на витрине до двойного срока, дальше — списанной.'],
    ['', 'Количество в списке (для списания — количество в документе iiko) разносится по партиям, которые есть к утру дня, начиная с самой старой. Отсюда статус и отклонение в днях.'],
    ['', 'В срок — всё количество покрыто партией с ожидаемой даты прихода. Раньше срока — ушла более свежая партия (отклонение «−»). Позже срока — ушла партия старше положенной (отклонение «+»). Нет прихода по сроку — в ожидаемую дату прихода не было, и партий не хватает. Приход меньше скидки/списания — приход был, но меньше заявленного. Превышен остаток FIFO — приход был не меньше заявленного, но к утру по FIFO осталось меньше; «(предупреждение)» — по 20% + склад 50% + возврат + списание ушло не меньше списка.'],
    ['', '«Остаток партии по FIFO (как на сайте)» — расчёт «Отчёта по дням»: сайт смотрит только партию с ожидаемой даты.'],
    ['h2', 'Потери выручки, ₸'],
    ['', '(списание + возврат по сроку) × цена + склад 50% × цена × 0,5 + продано со скидкой 20% × цена × 0,2. Цена — из Справочника, по всей продукции.'],
    ['h2', 'Склад 50%'],
    ['', 'Связка накладная → оплата → бронь — как на странице «Путь на склад 50%». ФИО — только на листе «Склад 50%». ИИН в отчёт не переносится.']
  ];
  r = 1;
  TXT.forEach(function(t){ ws.text(r, 1, t[1], t[0] === 'h' ? ST.title : t[0] === 'h2' ? ST.h2 : ST.wrap); r += t[0] === 'h2' ? 2 : 1; });
  ws.width(1, 150);
  var bad = 0;
  wb.sheets.forEach(function(s){ Object.keys(s.cells).forEach(function(rr){ Object.keys(s.cells[rr]).forEach(function(cc){ var v = s.cells[rr][cc].v;
    if (typeof v === 'string' && /\d{9,}/.test(v)) bad++; if (typeof v === 'number' && v >= 1e11) bad++; }); }); });
  QC.push(['ИИН: в книге нет последовательностей из 9 и более цифр (все листы, все ячейки)', bad, 0, bad ? 'РАСХОЖДЕНИЕ' : 'совпадает']);
  ws.table(r + 1, ['Проверка качества', 'Этот отчёт', 'Сравнение', 'Итог'], QC, { title: 'Проверки качества', chk: [3], fill: function(v){ return v === 'совпадает' ? 'GRN' : 'RED'; } });
  ws.width(2, 14); ws.width(3, 16); ws.width(4, 14);
  return { wb: wb, CONC: CONC, QC: QC };
}
/* ═════════ Отчёт за месяц: работа филиалов с Telegram-ботом ═════════
   Время отправки списков, дедлайн, версии, кто отправил — по листам
   «Заявки», «Заявки — история», «Списания» (действие приёмника «audit»).
   Дедлайн: приём накануне с 20:00, до 23:00; после 23:30 отправка закрыта. */
function monBookBot(A, from, to){
  var fx = MON.fx, nz = MON.nz, pct = MON.pct, sp_ = MON.sp_, ST = XW.ST;
  var WHS = whList().filter(function(w){ return !(typeof isDelivery === 'function' && isDelivery(w)) && !(typeof isClosed === 'function' && isClosed(w)); });
  var BRN = WHS.map(MON.brName), NW = WHS.length, DATES = MON.rng(from, to), per = ru(from) + '–' + ru(to);
  function wiOf(b){ var i = WHS.indexOf(b); if (i >= 0) return i; for (i = 0; i < NW; i++) if (sameBranch(b, WHS[i]) || sameBranch(WHS[i], b)) return i; return -1; }
  /* время: «2026-10-12 20:19» (так отдаёт приёмник) или «12.10.2026 20:19» */
  function tm(s){
    s = S(s); var m = s.match(/^(\d{4})-(\d\d)-(\d\d)[ T](\d\d):(\d\d)/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]);
    m = s.match(/^(\d\d)\.(\d\d)\.(\d{4})[ ,]+(\d\d?):(\d\d)/);
    if (m) return new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5]);
    return null;
  }
  function eve(d){ var p = d.split('-'); return new Date(+p[0], +p[1] - 1, +p[2] - 1); }   // полночь накануне
  function hm(t){ return ('0' + t.getHours()).slice(-2) + ':' + ('0' + t.getMinutes()).slice(-2); }
  function dtx(t){ return ('0' + t.getDate()).slice(-2) + '.' + ('0' + (t.getMonth() + 1)).slice(-2) + '.' + t.getFullYear() + ' ' + hm(t); }
  function iso(t){ return t.getFullYear() + '-' + ('0' + (t.getMonth() + 1)).slice(-2) + '-' + ('0' + t.getDate()).slice(-2); }
  var STS = ['В срок (20:00–23:00)', 'Раньше 20:00 (до открытия приёма)', 'После дедлайна (23:00–23:30)', 'Поздно (после 23:30 или в день уценки)', 'Не прислали'];
  var ONT = [STS[0], STS[1]];
  var COLF = {}; COLF[STS[0]] = 'GRN'; COLF[STS[1]] = 'LGR'; COLF[STS[2]] = 'YEL'; COLF[STS[3]] = 'RED'; COLF[STS[4]] = 'DGR'; COLF['да'] = 'YEL';
  COLF['Был расход по 20% без списка'] = 'RED'; COLF['Было списание без списка'] = 'RED'; COLF['Расхода по 20% не было'] = 'GRN'; COLF['Списания не было'] = 'GRN'; COLF['Нет данных iiko'] = 'GRY';
  function stOf(d, t){
    var e = eve(d).getTime(), x = t.getTime(), h = 3600e3;
    if (x < e + 20 * h) return STS[1];
    if (x <= e + 23 * h) return STS[0];
    if (x <= e + 23.5 * h) return STS[2];
    return STS[3];
  }
  function mins(d, t){ return (t.getTime() - (eve(d).getTime() + 20 * 3600e3)) / 60000; }
  function clock(m){ if (m === null || m === undefined || isNaN(m)) return ''; var t = new Date(2026, 0, 1, 20, 0); t = new Date(t.getTime() + m * 60000); return hm(t) + (m < 240 ? '' : ' (след. день)'); }

  /* ---- сборка: тип (D — скидка, S — списание) × филиал × день ---- */
  var REC = {};
  function rec(typ, wi, d){ var k = typ + '|' + wi + '|' + d; return REC[k] || (REC[k] = { t: {}, h: {}, n: 0, q: 0, who: {}, pr: 0, acc: 0 }); }
  (A.lists || []).forEach(function(r){ var wi = wiOf(r.branch), t = tm(r.sent); if (wi < 0 || !t || r.day < from || r.day > to) return;
    var o = rec('D', wi, r.day); o.t[t.getTime()] = t; o.n++; o.q += N(r.qty); if (S(r.who)) o.who[MON.stripId(S(r.who))] = (o.who[MON.stripId(S(r.who))] || 0) + 1; });
  (A.hist || []).forEach(function(r){ var wi = wiOf(r.branch), t = tm(r.sent); if (wi < 0 || !t || r.day < from || r.day > to) return;
    var o = rec('D', wi, r.day); o.h[t.getTime()] = t; });
  (A.off || []).forEach(function(r){ var wi = wiOf(r.branch), t = tm(r.sent); if (wi < 0 || !t || r.day < from || r.day > to) return;
    var o = rec('S', wi, r.day); o.t[t.getTime()] = t; o.n++; o.q += N(r.qty); if (S(r.who)) o.who[MON.stripId(S(r.who))] = (o.who[MON.stripId(S(r.who))] || 0) + 1;
    if (S(r.brought) !== '') o.pr++; if (S(r.acc)) o.acc++; });
  var ROWS = [];
  ['D', 'S'].forEach(function(typ){ for (var wi = 0; wi < NW; wi++) DATES.forEach(function(d){
    var o = REC[typ + '|' + wi + '|' + d];
    var cur = o ? Object.keys(o.t).map(Number) : [];
    if (!o || !cur.length){ ROWS.push({ typ: typ, wi: wi, d: d, st: STS[4] }); return; }
    var all = cur.concat(Object.keys(o.h).map(Number)).filter(function(v, i, a){ return a.indexOf(v) === i; }).sort(function(a, b){ return a - b; });
    var first = new Date(all[0]), last = new Date(Math.max.apply(null, cur));
    var dl = eve(d).getTime() + 23 * 3600e3;
    ROWS.push({ typ: typ, wi: wi, d: d, st: stOf(d, first), m: mins(d, first), first: first, last: last, versions: all.length, items: o.n, qty: o.q,
      who: Object.keys(o.who), late_edit: all.length > 1 && last.getTime() > dl, pr: o.pr, acc: o.acc });
  }); });
  var TN = { D: 'Скидка 20%', S: 'Списание' };
  function med(a){ if (!a.length) return null; var s = a.slice().sort(function(x, y){ return x - y; }), h = s.length >> 1; return s.length % 2 ? s[h] : (s[h - 1] + s[h]) / 2; }
  function agg(rows){
    var c = {}; rows.forEach(function(r){ c[r.st] = (c[r.st] || 0) + 1; });
    var sent = rows.filter(function(r){ return r.st !== STS[4]; }), ms = sent.map(function(r){ return r.m; });
    var who = {}; sent.forEach(function(r){ r.who.forEach(function(w){ who[w] = (who[w] || 0) + 1; }); });
    return { c: c, n: rows.length, sent: sent.length, ontime: (c[STS[0]] || 0) + (c[STS[1]] || 0), med: med(ms),
      mean: ms.length ? ms.reduce(function(a, b){ return a + b; }, 0) / ms.length : null, mx: ms.length ? Math.max.apply(null, ms) : null,
      edits: sent.filter(function(r){ return r.versions > 1; }).length, vers: sent.reduce(function(a, r){ return a + r.versions; }, 0),
      late: sent.filter(function(r){ return r.late_edit; }).length, who: who };
  }
  function G(c, s){ return c[s] || 0; }
  var BA = {};
  ['D', 'S'].forEach(function(t){ for (var wi = 0; wi < NW; wi++) BA[t + wi] = agg(ROWS.filter(function(r){ return r.typ === t && r.wi === wi; }));
    BA[t + 'all'] = agg(ROWS.filter(function(r){ return r.typ === t; })); });

  var wb = new XW.Book(), wsD = wb.sheet('Дашборд'), W = {};
  ['По филиалам', 'Нет списка — что было в iiko', 'Календарь отправок', 'Все отправки', 'Как считалось'].forEach(function(n){ W[n] = wb.sheet(n); });

  /* ---- По филиалам ---- */
  var ws = W['По филиалам'], hb = ['Филиал', 'Дат в периоде', 'Прислали', 'Не прислали', 'В срок (20:00–23:00)', 'Раньше 20:00', 'После дедлайна (23:00–23:30)', 'Поздно (после 23:30 / в день уценки)',
    'В срок, % от дат', 'Медиана времени отправки', 'Среднее время отправки', 'Самая поздняя отправка', 'Дней с исправлениями', 'Всего версий', 'Исправляли после 23:00', 'Отправители (сколько раз)'];
  var end = -2;
  ['D', 'S'].forEach(function(t){
    function line(a, name){ return [name, a.n, a.sent, G(a.c, STS[4]), G(a.c, STS[0]), G(a.c, STS[1]), G(a.c, STS[2]), G(a.c, STS[3]), pct(a.ontime, a.n), clock(a.med), clock(a.mean), clock(a.mx),
      a.edits, a.vers, a.late, Object.keys(a.who).sort(function(x, y){ return a.who[y] - a.who[x]; }).map(function(w){ return w + ' (' + a.who[w] + ')'; }).join('; ')]; }
    var rows = BRN.map(function(b, wi){ return line(BA[t + wi], b); });
    var T = line(BA[t + 'all'], 'Итого'); T[15] = '';
    end = ws.table(end + 3, hb, rows, { widths: [16, 9, 9, 9, 10, 9, 11, 12, 9, 11, 11, 13, 10, 9, 10, 60],
      title: TN[t] + ': списки через Telegram-бот, дата ' + (t === 'D' ? 'уценки' : 'списания') + ' ' + per, total: T })[1];
  });
  ws.freeze = [2, 3];

  /* ---- Нет списка — что было в iiko ---- */
  ws = W['Нет списка — что было в iiko'];
  var P = MON.prep(from, shift(to, 1));
  var wmap = WHS.map(function(w){ return P.WHS.indexOf(w); });
  var maxD = ''; DB.mov.forEach(function(r){ if (r[0] > maxD) maxD = r[0]; });
  function items(wi, d, kinds){
    var c = {}, pw = wmap[wi];
    Object.keys(P.MV).forEach(function(k){ var p = k.split('|'); if (+p[0] === pw && p[2] === d && kinds.indexOf(p[3]) >= 0) c[p[1]] = (c[p[1]] || 0) + P.MV[k]; });
    if (kinds.indexOf('s20') >= 0) Object.keys(P.S20).forEach(function(k){ var p = k.split('|'); if (+p[0] === pw && p[2] === d) c[p[1]] = (c[p[1]] || 0) + P.S20[k]; });
    return c;
  }
  function tot(c, ves){ return Object.keys(c).reduce(function(s, k){ return s + (MON.VES.test(SPR[k].name) === ves ? c[k] : 0); }, 0); }
  function txt(c){ return Object.keys(c).sort(function(a, b){ return c[b] - c[a]; }).map(function(k){ return SPR[k].name + ' — ' + fx(c[k]) + (MON.VES.test(SPR[k].name) ? ' кг' : ''); }).join('; '); }
  var MD = [], MS = [];
  ROWS.forEach(function(r){
    if (r.st !== STS[4]) return;
    var d = r.d, d1 = shift(d, 1);
    if (r.typ === 'D'){
      if (d > maxD){ MD.push([BRN[r.wi], ru(d), '', '', '', '', '', '', 'Нет данных iiko', '']); return; }
      var c = items(r.wi, d, ['s20']), q = tot(c, false), qv = tot(c, true), sm = 0, p5 = tot(items(r.wi, d, ['perem_50']), false) + tot(items(r.wi, d, ['perem_50']), true);
      Object.keys(P.S20SUM).forEach(function(k){ var p = k.split('|'); if (+p[0] === wmap[r.wi] && p[2] === d) sm += P.S20SUM[k]; });
      MD.push([BRN[r.wi], ru(d), fx(q), fx(qv), Object.keys(c).length, fx(sm), fx(sm * 0.2), fx(p5), (q || qv) ? 'Был расход по 20% без списка' : 'Расхода по 20% не было', txt(c)]);
    } else {
      if (d > maxD){ MS.push([BRN[r.wi], ru(d), '', '', '', '', 'Нет данных iiko', '']); return; }
      var a1 = items(r.wi, d, ['spisanie']), a2 = items(r.wi, d, ['vozvrat_srok']);
      var A1 = tot(a1, false) + tot(a1, true), A2 = tot(a2, false) + tot(a2, true), B1 = '', B2 = '', tb = '';
      if (d1 <= maxD){ var b1 = items(r.wi, d1, ['spisanie']), b2 = items(r.wi, d1, ['vozvrat_srok']); B1 = tot(b1, false) + tot(b1, true); B2 = tot(b2, false) + tot(b2, true);
        if (B1 || B2) tb = ru(d1) + ': ' + txt(items(r.wi, d1, ['spisanie', 'vozvrat_srok'])); }
      var ta = (A1 || A2) ? ru(d) + ': ' + txt(items(r.wi, d, ['spisanie', 'vozvrat_srok'])) : '';
      MS.push([BRN[r.wi], ru(d), fx(A1), fx(A2), B1 === '' ? '' : fx(B1), B2 === '' ? '' : fx(B2), (A1 || A2 || B1 || B2) ? 'Было списание без списка' : 'Списания не было', [ta, tb].filter(Boolean).join('; ')]);
    }
  });
  ws.text(1, 1, 'Дни без списка в боте: был ли в iiko расход по скидке 20% или списание', ST.title);
  ws.text(2, 1, 'Скидка: продажи со скидкой 20% в день уценки. Списание: «Акт списания» и «Возврат по сроку» в день списания и на следующий день — накладную на склад по срокам оформляют на следующий день. «Акт списания» бывает и не по сроку (кофе бесплатно, ДР, съёмка) — смотрите позиции. Если два дня без списка подряд, документ следующего дня виден в обеих строках.');
  ws.merge(2, 1, 2, 10); ws.heights[2] = 56;
  var e1 = ws.table(4, ['Филиал', 'Дата уценки', 'Продано со скидкой 20%, шт', 'Весовая со скидкой 20%, кг', 'Позиций', 'Сумма по полной цене, ₸', 'Скидка 20%, ₸', 'Склад 50%', 'Итог', 'Что продали со скидкой 20%'],
    MD, { widths: [16, 12, 12, 12, 9, 13, 12, 9, 30, 90], chk: [8], fill: function(v){ return COLF[v]; }, title: 'Список на скидку не прислали' });
  var e2 = ws.table(e1[1] + 3, ['Филиал', 'Дата списания', 'Акт списания в этот день', 'Возврат по сроку в этот день', 'Акт списания на след. день', 'Возврат по сроку на след. день', 'Итог', 'Что списали (iiko)'],
    MS, { chk: [6], fill: function(v){ return COLF[v]; }, title: 'Список списания не прислали' });
  var cD = {}, cS = {}; MD.forEach(function(x){ cD[x[8]] = (cD[x[8]] || 0) + 1; }); MS.forEach(function(x){ cS[x[6]] = (cS[x[6]] || 0) + 1; });
  ws.table(e2[1] + 3, ['Итог', 'Дней'], Object.keys(cD).map(function(k){ return [k, cD[k]]; }).concat(Object.keys(cS).map(function(k){ return [k, cS[k]]; })),
    { chk: [0], fill: function(v){ return COLF[v]; }, title: 'Сводка' });
  ws.freeze = [3, 5];
  var SD = MD.reduce(function(s, x){ return s + (+x[2] || 0); }, 0), SDs = MD.reduce(function(s, x){ return s + (+x[5] || 0); }, 0);

  /* ---- Календарь ---- */
  ws = W['Календарь отправок'];
  var rr = 1;
  ['D', 'S'].forEach(function(t){
    var data = [], fills = [];
    BRN.forEach(function(b, wi){
      var row = [b], fr = [null];
      DATES.forEach(function(d){
        var r = ROWS.filter(function(x){ return x.typ === t && x.wi === wi && x.d === d; })[0];
        if (r.st === STS[4]) row.push('—');
        else row.push(hm(r.first) + (iso(r.first) === shift(d, -1) ? '' : ' (' + ru(iso(r.first)).slice(0, 5) + ')') + (r.versions > 1 ? ' ×' + r.versions : ''));
        fr.push(COLF[r.st]);
      });
      data.push(row); fills.push(fr);
    });
    rr = ws.table(rr, ['Филиал'].concat(DATES.map(function(d){ return ru(d).slice(0, 5); })), data, { widths: [16].concat(DATES.map(function(){ return 8; })), fills: fills,
      title: TN[t] + ': время первой отправки накануне дня ' + (t === 'D' ? 'уценки' : 'списания') + ' (×N — сколько было версий)' })[1] + 2;
  });
  ws.table(rr + 1, ['Цвет — статус'], STS.map(function(s){ return [s]; }), { widths: [40], chk: [0], fill: function(v){ return COLF[v]; }, title: 'Обозначения' });
  ws.freeze = [2, 3];

  /* ---- Все отправки ---- */
  ws = W['Все отправки'];
  var all = ROWS.slice().sort(function(a, b){ return (a.typ < b.typ ? -1 : a.typ > b.typ ? 1 : 0) || a.wi - b.wi || (a.d < b.d ? -1 : a.d > b.d ? 1 : 0); }).map(function(r){
    if (r.st === STS[4]) return [TN[r.typ], BRN[r.wi], ru(r.d), '', '', '', '', '', r.st, '', '', ''];
    var late = (r.first.getTime() - (eve(r.d).getTime() + 23 * 3600e3)) / 60000;
    return [TN[r.typ], BRN[r.wi], ru(r.d), dtx(r.first), dtx(r.last), r.versions, r.items, fx(r.qty), r.st, late > 0 ? fx(late) : '', r.late_edit ? 'да' : '', r.who.join('; ')];
  });
  ws.table(1, ['Список', 'Филиал', 'Дата уценки / списания', 'Первая отправка', 'Последняя версия', 'Версий', 'Позиций', 'Количество', 'Статус', 'Минут после 23:00 (если позже)', 'Исправляли после 23:00', 'Отправители'],
    all, { widths: [12, 16, 12, 17, 17, 8, 8, 10, 30, 11, 10, 40], chk: [8, 10], fill: function(v){ return COLF[v]; } });
  ws.freeze = [3, 2]; ws.filter = 'A1:L' + (all.length + 1);

  /* ---- Дашборд ---- */
  ws = wsD; ws.grid = false;
  ws.width(1, 3); ws.width(2, 18); for (var c = 3; c <= 14; c++) ws.width(c, 12);
  var aD = BA.Dall, aS = BA.Sall;
  var idx = BRN.map(function(b, i){ return i; });
  var rk = idx.slice().sort(function(a, b){ return (BA['D' + b].ontime - BA['D' + a].ontime) || ((BA['D' + a].med || 0) - (BA['D' + b].med || 0)); });
  function lst(f){ return idx.filter(function(i){ return f(i) > 0; }).sort(function(a, b){ return f(b) - f(a); }).map(function(i){ return BRN[i] + ' — ' + f(i); }).join(', '); }
  var lateL = lst(function(i){ return G(BA['D' + i].c, STS[2]) + G(BA['D' + i].c, STS[3]); }), missL = lst(function(i){ return G(BA['D' + i].c, STS[4]); });
  var editB = idx.slice().sort(function(a, b){ return BA['D' + b].vers - BA['D' + a].vers; })[0];
  var best = rk[0], worst = rk[rk.length - 1];
  var CONC = [
    'Список на скидку: из ' + aD.n + ' возможных списков (' + NW + ' филиалов × ' + DATES.length + ' дней) прислали ' + aD.sent + '; вовремя (до 23:00 накануне) — ' + aD.ontime + ' (' + sp_(pct(aD.ontime, aD.n)) +
      '%), после дедлайна 23:00–23:30 — ' + G(aD.c, STS[2]) + ', позже 23:30 или в день уценки — ' + G(aD.c, STS[3]) + ', не прислали — ' + G(aD.c, STS[4]) + '.',
    'Обычно список приходит в ' + clock(aD.med) + ' (медиана), в среднем — в ' + clock(aD.mean) + '; до 21:30 отправлено ' + ROWS.filter(function(x){ return x.typ === 'D' && x.st !== STS[4] && x.m <= 90; }).length + ' списков из ' + aD.sent + '.',
    'Лучше всех по дедлайну ' + BRN[best] + ' (вовремя ' + BA['D' + best].ontime + ' из ' + BA['D' + best].n + ', обычно в ' + clock(BA['D' + best].med) + '). Опоздания: ' + (lateL || 'нет') + '.',
    BRN[worst] + ' — хуже всех: обычно в ' + (clock(BA['D' + worst].med) || '—') + ', вовремя ' + BA['D' + worst].ontime + ' из ' + BA['D' + worst].n + '.',
    'Не прислали список на скидку: ' + (missL || 'все филиалы прислали каждый день') + '. В ' + G(cD, 'Был расход по 20% без списка') + ' таких дней по iiko была продажа со скидкой 20% — всего ' + sp_(fx(SD)) +
      ' шт штучной продукции, по полной цене ' + sp_(fx(SDs)) + ' ₸, скидка — ' + sp_(fx(SDs * 0.2)) + ' ₸. Из ' + MS.length + ' дней без списка списания в ' + G(cS, 'Было списание без списка') + ' по iiko было списание или возврат по сроку.',
    'Исправления: списков, отправленных больше одного раза, — ' + aD.edits + ' из ' + aD.sent + ' (версий всего — ' + aD.vers + '). Чаще всего исправляет ' + BRN[editB] + ': версий — ' + BA['D' + editB].vers + ' за ' + BA['D' + editB].sent + ' дней. Исправлений после 23:00 — ' + aD.late + '.',
    'Список списания: прислали ' + aS.sent + ' из ' + aS.n + ' (вовремя — ' + aS.ontime + ', ' + sp_(pct(aS.ontime, aS.n)) + '%; не прислали — ' + G(aS.c, STS[4]) + '). В приёмке цеха «Привезли» заполнено в ' +
      ROWS.filter(function(r){ return r.typ === 'S' && r.pr; }).length + ' списках.'
  ];
  ws.text(1, 2, 'Работа филиалов с Telegram-ботом: время и дедлайн отправки списков', ST.title);
  ws.text(2, 2, 'Дата уценки/списания ' + per + ', ' + NW + ' филиалов. Дедлайн — 23:00 накануне; приём с 20:00, после 23:30 отправка закрыта.', ST.def);
  ws.text(4, 2, 'Выводы', ST.h2);
  var r = 5;
  CONC.forEach(function(t, j){ ws.text(r, 2, (j + 1) + '. ' + t); ws.merge(r, 2, r, 13); ws.heights[r] = 32; r++; });
  r++;
  var rows = rk.map(function(i, p){ var a = BA['D' + i], s = BA['S' + i]; return [p + 1, BRN[i], a.ontime, G(a.c, STS[2]), G(a.c, STS[3]), G(a.c, STS[4]), pct(a.ontime, a.n), clock(a.med), a.vers, s.sent, pct(s.ontime, s.n)]; });
  var q = ws.table(r, ['Место', 'Филиал', 'Скидка: вовремя', 'После дедлайна (23:00–23:30)', 'Поздно (после 23:30)', 'Не прислали', 'Скидка вовремя, %', 'Обычное время (медиана)', 'Версий списка', 'Списание: прислали', 'Списание вовремя, %'],
    rows, { c0: 2, title: 'Рейтинг филиалов по дедлайну (список на скидку)' });
  var dist = BRN.map(function(b, i){ return [b].concat(STS.map(function(s){ return G(BA['D' + i].c, s); })); });
  var q2 = ws.table(q[1] + 3, ['Филиал'].concat(STS), dist, { c0: 2, title: 'Распределение по статусам' });
  ws.chart({ kind: 'col', grouping: 'stacked', title: 'Список на скидку: когда отправляют (по филиалам)', colors: ['70AD47', 'A9D18E', 'FFC000', 'C00000', 'A5A5A5'], cat: [2, q2[0] + 1, q2[1]],
    series: [3, 4, 5, 6, 7].map(function(cc){ return { nameCell: [cc, q2[0]], val: [cc, q2[0] + 1, q2[1]] }; }), at: [2, q2[1] + 2], size: [11, 17] });
  var bk = {};
  ROWS.forEach(function(x){ if (x.typ !== 'D' || x.st === STS[4]) return; var key = iso(x.first) === shift(x.d, -1) ? ('0' + x.first.getHours()).slice(-2) + ':00' : 'в день уценки'; bk[key] = (bk[key] || 0) + 1; });
  var hk = Object.keys(bk).filter(function(k){ return k !== 'в день уценки'; }).sort().concat(bk['в день уценки'] ? ['в день уценки'] : []);
  var q3 = ws.table(q2[1] + 20, ['Час отправки (накануне)', 'Списков'], hk.map(function(k){ return [k, bk[k]]; }), { c0: 2, title: 'Во сколько отправляют список на скидку' });
  if (hk.length) ws.chart({ kind: 'col', title: 'Списков на скидку по часу отправки', legend: false, cat: [2, q3[0] + 1, q3[1]], series: [{ nameCell: [3, q3[0]], val: [3, q3[0] + 1, q3[1]] }], at: [6, q3[0]], size: [7, 14] });
  var dr = DATES.map(function(d){
    var xs = ROWS.filter(function(x){ return x.typ === 'D' && x.d === d && x.st !== STS[4]; }).map(function(x){ return x.m; });
    var md = med(xs);
    return [ru(d), xs.length, ROWS.filter(function(x){ return x.typ === 'D' && x.d === d && ONT.indexOf(x.st) >= 0; }).length, md === null ? '' : Math.round((md / 60 + 20) * 100) / 100];
  });
  var q4 = ws.table(Math.max(q3[1], q3[0] + 14) + 3, ['Дата уценки', 'Прислали', 'Вовремя', 'Медиана времени, ч (20,5 = 20:30)'], dr, { c0: 2, title: 'По датам' });
  ws.chart({ kind: 'line', title: 'Медиана времени отправки по датам, ч', legend: false, cat: [2, q4[0] + 1, q4[1]], series: [{ nameCell: [5, q4[0]], val: [5, q4[0] + 1, q4[1]] }], at: [7, q4[0]], size: [8, 14] });

  /* ---- Как считалось ---- */
  ws = W['Как считалось'];
  var TXT = [
    'Источник: таблица «20-50% жеңілдік өнімдері», листы «Заявки» (текущий список на скидку: филиал, дата уценки, позиции, «Когда прислали», «Кто отправил»), «Заявки — история» (прежние версии) и «Списания» (список списания, «Привезли», «Когда приняли») — через приёмник, только чтение. Расход по 20% и списание — из выгрузок iiko на сайте.',
    'Период: дата уценки/списания ' + per + ' (' + DATES.length + ' дней × ' + NW + ' филиалов).',
    'Правила по времени: приём списков накануне с 20:00, дедлайн — 23:00, после 23:30 отправить уже нельзя.',
    'Статусы (по первой отправке): «В срок» — накануне 20:00–23:00; «Раньше 20:00» — до открытия приёма (засчитывается как вовремя); «После дедлайна» — 23:00–23:30; «Поздно» — после 23:30 или уже в день уценки; «Не прислали» — за эту дату нет ни одной версии.',
    'Версии: первая — самая ранняя «Когда прислали» среди текущего списка и истории; число версий — сколько раз филиал отправлял список на эту дату. «Исправляли после 23:00» — последняя версия отправлена после дедлайна.',
    '«Не прислали» не всегда нарушение: в какой-то день филиалу может быть нечего уценять или списывать — смотрите лист «Нет списка — что было в iiko».',
    '«Отправители» — имя аккаунта Telegram, с которого пришёл список. Списки должны идти с рабочего телефона филиала.',
    'Время показано как есть, без округления. Медиана — время, раньше которого филиал отправляет половину своих списков.'
  ];
  TXT.forEach(function(t, j){ ws.text(j + 1, 1, t); });
  ws.width(1, 140);
  var bad = 0;
  wb.sheets.forEach(function(s){ Object.keys(s.cells).forEach(function(r1){ Object.keys(s.cells[r1]).forEach(function(c1){ var v = s.cells[r1][c1].v; if (typeof v === 'string' && /\d{9,}/.test(v)) bad++; }); }); });
  var sent = ROWS.filter(function(r){ return r.st !== STS[4]; }).length;
  var QC = [['Сумма статусов на «По филиалам» = строк на «Все отправки»', BA.Dall.n + BA.Sall.n, ROWS.length], ['ИИН: нет последовательностей из 9 и более цифр', bad, 0]]
    .map(function(x){ return x.concat([x[1] === x[2] ? 'совпадает' : 'РАСХОЖДЕНИЕ']); });
  ws.table(TXT.length + 2, ['Проверка', 'Значение', 'Ожидается', 'Итог'], QC, { title: 'Проверки', chk: [3], fill: function(v){ return v === 'совпадает' ? 'GRN' : 'RED'; } });
  ws.width(2, 12); ws.width(3, 12); ws.width(4, 12);
  return { wb: wb, CONC: CONC, QC: QC, sent: sent };
}
/* ═════════ Отчёт за месяц: экран ═════════ */
(function(){
  function el(id){ return document.getElementById(id); }
  function iso(d){ return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
  /* по умолчанию — прошлый календарный месяц */
  var now = new Date(), f = new Date(now.getFullYear(), now.getMonth() - 1, 1), t = new Date(now.getFullYear(), now.getMonth(), 0);
  if (el('monFrom') && !el('monFrom').value) el('monFrom').value = iso(f);
  if (el('monTo') && !el('monTo').value) el('monTo').value = iso(t);
  function per(){
    var a = el('monFrom').value, b = el('monTo').value;
    if (!a || !b || a > b) throw new Error('Укажите период: дата «с» не позже даты «по».');
    return [a, b];
  }
  function save(blob, name){
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name.replace(/[\\/:*?"<>|]/g, '') + '.xlsx'; a.style.display = 'none';
    document.body.appendChild(a); a.click();
    setTimeout(function(){ URL.revokeObjectURL(a.href); a.remove(); }, 4000);
  }
  function busy(b, on){ el('btnMonMain').disabled = on; el('btnMonBot').disabled = on; }
  async function main(){
    try{
      var p = per();
      if (!DB.mov || !DB.mov.length) throw new Error('Нет выгрузок iiko — загрузите товарную ведомость и товарный отчёт.');
      busy(this, true); msg('mMon', 'Считаю «Скидка и списание» за ' + ru(p[0]) + '–' + ru(p[1]) + '…', 'wait');
      await new Promise(function(r){ setTimeout(r, 30); });
      var R = MON.compute(p[0], p[1]);
      var B = monBookMain(R, p[0], p[1]), akt = '';
      /* Листы по актам: строки фото списания — из приёмника, нужен ключ на чтение. */
      try{
        if (apiCfg().read){
          msg('mMon', 'Забираю из приёмника акты фото списания…', 'wait');
          var P = await AKT.load(p[0], p[1]);
          AKT.reasonsSheet(B.wb, P.rows || [], p[0], p[1]);
          akt = ' Листы «Акты по причинам» и «Акты — сверка с iiko» добавлены.';
        } else akt = ' Листов по актам нет — нужен ключ на чтение.';
      }catch(e){ akt = ' Листы по актам не добавлены: ' + e.message; }
      save(await B.wb.blob(), 'Отчёт скидка и списание ' + ru(p[0]) + ' — ' + ru(p[1]));
      var bad = B.QC.filter(function(q){ return q[3] !== 'совпадает'; }).length;
      msg('mMon', 'Готово: строк скидки — ' + R.DISC.length + ', списания — ' + R.SPW.length + '. Проверки: ' + (bad ? 'расхождений ' + bad + ' — см. лист «Как считалось»' : 'все сошлись') + '.' + akt, bad ? 'err' : 'ok');
    }catch(e){ msg('mMon', 'Не получилось: ' + e.message, 'err'); }
    busy(this, false);
  }
  async function bot(){
    try{
      var p = per(), c = apiCfg();
      if (!c.read) throw new Error('Нужен ключ на чтение — впишите его в «Списках филиалов» (Забрать из приёмника).');
      busy(this, true); msg('mMon', 'Забираю из приёмника время отправки списков…', 'wait');
      var A = await apiCall('audit', { key: c.read, from: p[0], to: p[1] });
      msg('mMon', 'Считаю «Работа с ботом»…', 'wait');
      await new Promise(function(r){ setTimeout(r, 30); });
      var B = monBookBot(A, p[0], p[1]);
      save(await B.wb.blob(), 'Отчёт работа с Telegram-ботом ' + ru(p[0]) + ' — ' + ru(p[1]));
      msg('mMon', 'Готово: отправок — ' + B.sent + '. Проверки: ' + (B.QC.every(function(q){ return q[3] === 'совпадает'; }) ? 'все сошлись' : 'есть расхождения — см. «Как считалось»') + '.', 'ok');
    }catch(e){ msg('mMon', 'Не получилось: ' + e.message, 'err'); }
    busy(this, false);
  }
  if (el('btnMonMain')) el('btnMonMain').onclick = main;
  if (el('btnMonBot')) el('btnMonBot').onclick = bot;
})();
/* ═════════ Акты списания: сверка с iiko и итог по причинам ═════════
   Строки «Фото списания» берём из приёмника (ph_rows, ключ на чтение),
   акты списания iiko — из загруженной товарной ведомости (вид «Акт списания»).
   Сверяем «Акт на месте» из приложения с актами iiko: день × филиал × позиция. */
var AKT = (function(){
  function el(id){ return document.getElementById(id); }
  function fx(v){ if (typeof v !== 'number') return v; var r = Math.round(v * 1e6) / 1e6; return r === 0 ? 0 : r; }
  function bkey(w){ var n = whNum(w); return n !== null ? 'n' + n : nrm(String(w).replace(/склад/i, '')); }
  function dkey(s){ return nrm(s).replace(/\s*\((ф|шт)\)\s*$/, '').replace(/(\d)\s+см/g, '$1см'); }
  /* Кусок пиццы в приложении — в iiko списывается долей целой пиццы 30 см (1 кусок = 1/8):
     «Пицца курица (кусок)» → «Пицца с курицей 30см». Слова сравниваем по первым 5 буквам. */
  var PIZ = null;
  function pizzaWhole(name){
    var k = nrm(name); if (!/^пицц/.test(k) || !/\(кус(ок|очек)\)\s*$/.test(k)) return null;
    if (!PIZ){ PIZ = []; Object.keys(typeof SPR === 'object' && SPR || {}).forEach(function(n){ if (/^пицц/.test(n) && /30\s*см/.test(n)) PIZ.push(n); }); }
    var words = k.replace(/\(кус(ок|очек)\)/, '').split(/\s+/).filter(function(w){ return w && !/^пицц/.test(w); });
    var hit = PIZ.filter(function(n){ var nw = n.split(/\s+/);
      return words.every(function(w){ var st = w.slice(0, 5); return nw.some(function(x){ return x.slice(0, 5) === st; }); }); });
    return hit.length === 1 ? (SPR[hit[0]] && SPR[hit[0]].name) || hit[0] : null;
  }
  async function load(from, to){
    var c = apiCfg();
    if (!c.read) throw new Error('Нужен ключ на чтение — впишите его в «Списках филиалов» (Забрать из приёмника).');
    var P = await apiCall('ph_rows', { key: c.read, from: from, to: to });
    if (!P || P.error) throw new Error((P && P.error) || 'Приёмник не ответил.');
    return P;
  }
  /* Сверка: строки с ключом день|филиал|позиция. */
  function compare(rows, from, to){
    var M = {}, names = {};
    function get(day, b, dish){
      var k = day + '|' + bkey(b) + '|' + dkey(dish);
      if (!M[k]) M[k] = { day: day, b: b, dish: dish, app: 0, iiko: 0, reasons: {}, kg: false, inApp: false, inIiko: false };
      if (!names[bkey(b)] || /склад/i.test(b)) names[bkey(b)] = b;
      return M[k];
    }
    /* Из актов iiko берём только продукцию: позиции справочника и то, что филиалы
       вообще отправляли через приложение. Тряпки, сливки, стаканы — мимо. */
    var known = {};
    Object.keys(typeof SPR === 'object' && SPR || {}).forEach(function(n){ known[dkey(n)] = 1; });
    rows.forEach(function(r){ known[dkey(r.dish)] = 1; var w = pizzaWhole(r.dish); if (w) known[dkey(w)] = 1; });
    rows.forEach(function(r){
      if (r.dest !== 'akt' || !(r.qty > 0)) return;
      var w = pizzaWhole(r.dish), q = w ? r.qty / 8 : r.qty;
      var x = get(r.day, r.branch, w || r.dish); x.app += q; x.inApp = true; x.reasons[r.reason + (w ? ' (кусков: ' + r.qty + ')' : '')] = 1; if (r.unit === 'кг') x.kg = true;
    });
    (DB.mov || []).forEach(function(r){
      if (r[3] !== 'spisanie' || r[0] < from || r[0] > to || !r[1]) return;
      if (typeof isDelivery === 'function' && isDelivery(r[1])) return;
      if (!known[dkey(r[2])]) return;
      var x = get(r[0], r[1], r[2]); x.iiko += Math.abs(N(r[4])); x.inIiko = true;
    });
    return Object.keys(M).map(function(k){
      var x = M[k], d = fx(x.iiko - x.app);
      x.b = names[bkey(x.b)] || x.b; x.app = fx(x.app); x.iiko = fx(x.iiko); x.diff = d;
      x.st = !x.inIiko ? 'Нет акта в iiko' : !x.inApp ? 'Нет в приложении' : Math.abs(d) < 1e-9 ? 'Совпадает' : d > 0 ? 'В iiko больше' : 'В iiko меньше';
      x.why = Object.keys(x.reasons).join(', ');
      return x;
    }).sort(function(a, b){ return a.day < b.day ? -1 : a.day > b.day ? 1 : (whNum(a.b) || 99) - (whNum(b.b) || 99) || (a.dish < b.dish ? -1 : 1); });
  }
  var COL = { 'Совпадает': 'GRN', 'Нет акта в iiko': 'RED', 'Нет в приложении': 'YEL', 'В iiko больше': 'ORG', 'В iiko меньше': 'ORG' };

  /* ── экран ── */
  var LAST = null;
  function iso(d){ return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
  var y = new Date(); y.setDate(y.getDate() - 1);
  if (el('aktFrom') && !el('aktFrom').value) el('aktFrom').value = iso(y);
  if (el('aktTo') && !el('aktTo').value) el('aktTo').value = iso(y);
  function draw(){
    if (!LAST) return;
    var only = el('aktDiff') && el('aktDiff').checked, L = LAST.lines.filter(function(x){ return !only || x.st !== 'Совпадает'; });
    var cnt = {}; LAST.lines.forEach(function(x){ cnt[x.st] = (cnt[x.st] || 0) + 1; });
    var head = '<div class="row" style="gap:8px;flex-wrap:wrap;margin:6px 0 10px">' + ['Совпадает', 'В iiko больше', 'В iiko меньше', 'Нет акта в iiko', 'Нет в приложении'].map(function(s){
      return '<span style="padding:4px 10px;border-radius:10px;background:var(--soft);font-size:13px">' + esc(s) + ': <b>' + (cnt[s] || 0) + '</b></span>'; }).join('') + '</div>';
    if (!L.length){ el('aktBody').innerHTML = head + '<p class="lead">' + (LAST.lines.length ? 'Расхождений нет.' : 'За эти дни нет ни актов в приложении, ни актов iiko.') + '</p>'; return; }
    var bg = { GRN: '#e2efda', RED: '#f8cbad', YEL: '#fff2cc', ORG: '#fce4d6' };
    var TD = 'style="padding:6px 8px;border-bottom:1px solid var(--bd)"', TN = 'style="padding:6px 8px;border-bottom:1px solid var(--bd);text-align:right"';
    el('aktBody').innerHTML = head + '<div class="card" style="overflow:auto"><table style="width:100%;border-collapse:collapse;font-size:14px"><thead><tr>' +
      ['Дата', 'Филиал', 'Позиция', 'Причина в приложении', 'В приложении', 'В iiko', 'Разница', 'Итог'].map(function(h, i){
        return '<th style="padding:6px 8px;text-align:' + (i >= 4 && i <= 6 ? 'right' : 'left') + ';border-bottom:2px solid var(--bd);color:var(--mut);font-weight:600">' + h + '</th>'; }).join('') +
      '</tr></thead><tbody>' +
      L.map(function(x){ return '<tr><td ' + TD + '>' + ru(x.day) + '</td><td ' + TD + '>' + esc(x.b) + '</td><td ' + TD + '>' + esc(x.dish) + '</td><td ' + TD + '>' + esc(x.why) + '</td>' +
        '<td ' + TN + '>' + (x.inApp ? x.app : '') + '</td><td ' + TN + '>' + (x.inIiko ? x.iiko : '') + '</td><td ' + TN + '>' + (x.diff ? x.diff : '') + '</td>' +
        '<td style="padding:6px 8px;border-bottom:1px solid var(--bd);color:#222;background:' + bg[COL[x.st]] + '">' + esc(x.st) + '</td></tr>'; }).join('') + '</tbody></table></div>';
  }
  async function run(){
    try{
      var a = el('aktFrom').value, b = el('aktTo').value;
      if (!a || !b || a > b) throw new Error('Укажите период: дата «с» не позже даты «по».');
      if (!DB.mov || !DB.mov.length) throw new Error('Нет товарной ведомости — загрузите её в «Выгрузках iiko».');
      var maxd = DB.mov.reduce(function(m, r){ return r[0] > m ? r[0] : m; }, '');
      msg('mAkt', 'Забираю из приёмника акты из приложения…', 'wait');
      var P = await load(a, b);
      LAST = { from: a, to: b, lines: compare(P.rows || [], a, b) };
      draw();
      var bad = LAST.lines.filter(function(x){ return x.st !== 'Совпадает'; }).length;
      msg('mAkt', 'Сверено строк: ' + LAST.lines.length + ', расхождений: ' + bad + '.' +
        (maxd < b ? ' Внимание: ведомость загружена только по ' + ru(maxd) + ' — акты iiko после этой даты не видны.' : ''), bad ? 'err' : 'ok');
    }catch(e){ msg('mAkt', 'Не получилось: ' + e.message, 'err'); }
  }
  async function xlsx(){
    if (!LAST) return msg('mAkt', 'Сначала нажмите «Сверить».', 'err');
    var wb = new XW.Book(), ws = wb.sheet('Акты — сверка с iiko');
    ws.set(1, 1, 'Акты списания: приложение против iiko, ' + ru(LAST.from) + ' — ' + ru(LAST.to), XW.ST.title);
    ws.table(3, ['Дата', 'Филиал', 'Позиция', 'Причина в приложении', 'В приложении', 'В iiko', 'Разница (iiko − приложение)', 'Итог'],
      LAST.lines.map(function(x){ return [ru(x.day), x.b, x.dish, x.why, x.inApp ? x.app : '', x.inIiko ? x.iiko : '', x.diff || '', x.st]; }),
      { chk: [7], fill: function(v){ return COL[v]; }, widths: [11, 22, 34, 26, 12, 10, 14, 18] });
    ws.freeze = [1, 4]; ws.filter = 'A3:H' + (3 + LAST.lines.length);
    var blob = await wb.blob(), aa = document.createElement('a');
    aa.href = URL.createObjectURL(blob); aa.download = ('Акты сверка с iiko ' + ru(LAST.from) + ' — ' + ru(LAST.to)).replace(/[\\/:*?"<>|]/g, '') + '.xlsx';
    document.body.appendChild(aa); aa.click(); setTimeout(function(){ URL.revokeObjectURL(aa.href); aa.remove(); }, 4000);
  }
  if (el('aktRun')) el('aktRun').onclick = run;
  if (el('aktDiff')) el('aktDiff').onchange = draw;
  if (el('aktXlsx')) el('aktXlsx').onclick = xlsx;

  /* ── лист «Акты по причинам» в книге «Скидка и списание» ── */
  function reasonsSheet(wb, rows, from, to){
    var ws = wb.sheet('Акты по причинам'), ST = XW.ST;
    ws.set(1, 1, 'Фото списания по причинам, ' + ru(from) + ' — ' + ru(to), ST.title);
    ws.text(2, 1, 'Источник — мини-приложение «Фото списания» (лист «Списание — фото»). Отправка = один акт филиала (одна карточка в группе). «По срокам» делится на «Склад по срокам» (расходная накладная) и «Акт на месте». Штуки и килограммы — отдельно.');
    ws.merge(2, 1, 2, 8);
    var R = {}, BR = {}, D = {}, brs = {}, rs = {};
    rows.forEach(function(r){
      var why = r.reason + (r.dest === 'sklad' ? ' — склад по срокам' : (r.reason === AKT_DAILY ? ' — акт на месте' : ''));
      var kg = r.unit === 'кг', send = r.day + '|' + r.branch + '|' + r.reason + '|' + r.sub;
      var x = R[why] || (R[why] = { sends: {}, branches: {}, pos: 0, sht: 0, kg: 0 });
      x.sends[send] = 1; x.branches[r.branch] = 1; x.pos++; if (kg) x.kg += r.qty; else x.sht += r.qty;
      var bk = r.branch + '|' + why, y = BR[bk] || (BR[bk] = { sends: {}, sht: 0, kg: 0 });
      y.sends[send] = 1; if (kg) y.kg += r.qty; else y.sht += r.qty;
      brs[r.branch] = 1; rs[why] = 1;
      var dk = why + '|' + r.dish, z = D[dk] || (D[dk] = { why: why, dish: r.dish, sht: 0, kg: 0, br: {} });
      if (kg) z.kg += r.qty; else z.sht += r.qty; z.br[r.branch] = 1;
    });
    var WHY = Object.keys(rs).sort(function(a, b){ return Object.keys(R[b].sends).length - Object.keys(R[a].sends).length; });
    var BRS = Object.keys(brs).sort(function(a, b){ return (whNum(a) || 99) - (whNum(b) || 99) || (a < b ? -1 : 1); });
    var r0 = 4;
    var t1 = ws.table(r0, ['Причина', 'Отправок', 'Филиалов', 'Строк (позиций)', 'Штук', 'Килограммов'],
      WHY.map(function(w){ var x = R[w]; return [w, Object.keys(x.sends).length, Object.keys(x.branches).length, x.pos, fx(x.sht), fx(x.kg)]; }),
      { title: 'Сеть: по причинам', total: ['Итого', WHY.reduce(function(s, w){ return s + Object.keys(R[w].sends).length; }, 0), BRS.length,
        WHY.reduce(function(s, w){ return s + R[w].pos; }, 0), fx(WHY.reduce(function(s, w){ return s + R[w].sht; }, 0)), fx(WHY.reduce(function(s, w){ return s + R[w].kg; }, 0))],
        widths: [40, 11, 10, 14, 10, 12] });
    r0 = t1[1] + 3;
    var t2 = ws.table(r0, ['Филиал'].concat(WHY.map(function(w){ return w + ', отправок / шт'; })),
      BRS.map(function(b){ return [b].concat(WHY.map(function(w){ var y = BR[b + '|' + w]; return y ? Object.keys(y.sends).length + ' / ' + fx(y.sht) + (y.kg ? ' + ' + fx(y.kg) + ' кг' : '') : ''; })); }),
      { title: 'Филиалы × причины (отправок / штук)' });
    r0 = t2[1] + 3;
    var DL = Object.keys(D).map(function(k){ return D[k]; }).sort(function(a, b){ return a.why < b.why ? -1 : a.why > b.why ? 1 : (b.sht + b.kg) - (a.sht + a.kg); });
    ws.table(r0, ['Причина', 'Позиция', 'Штук', 'Килограммов', 'Филиалов'],
      DL.map(function(z){ return [z.why, z.dish, fx(z.sht), fx(z.kg), Object.keys(z.br).length]; }), { title: 'Что списывали: позиции по причинам' });
    ws.freeze = [2, 5];
    /* сверка с iiko за весь период — итог по филиалам */
    var L = compare(rows, from, to), S2 = {};
    L.forEach(function(x){ var s = S2[x.b] || (S2[x.b] = { app: 0, iiko: 0, n: 0, bad: 0 }); s.app += x.inApp ? x.app : 0; s.iiko += x.inIiko ? x.iiko : 0; s.n++; if (x.st !== 'Совпадает') s.bad++; });
    var ws2 = wb.sheet('Акты — сверка с iiko');
    ws2.set(1, 1, 'Акты списания: приложение против iiko, ' + ru(from) + ' — ' + ru(to), ST.title);
    var tt = ws2.table(3, ['Филиал', 'Акт на месте в приложении, ед.', 'Акты списания iiko, ед.', 'Строк сверено', 'Расхождений'],
      Object.keys(S2).sort(function(a, b){ return (whNum(a) || 99) - (whNum(b) || 99); }).map(function(b){ var s = S2[b]; return [b, fx(s.app), fx(s.iiko), s.n, s.bad]; }),
      { title: 'По филиалам', widths: [24, 16, 16, 12, 12] });
    ws2.table(tt[1] + 3, ['Дата', 'Филиал', 'Позиция', 'Причина в приложении', 'В приложении', 'В iiko', 'Разница (iiko − приложение)', 'Итог'],
      L.map(function(x){ return [ru(x.day), x.b, x.dish, x.why, x.inApp ? x.app : '', x.inIiko ? x.iiko : '', x.diff || '', x.st]; }),
      { title: 'Построчно', chk: [7], fill: function(v){ return COL[v]; }, widths: [11, 24, 34, 26, 12, 10, 14, 18] });
  }
  var AKT_DAILY = 'По срокам';
  return { load: load, compare: compare, reasonsSheet: reasonsSheet };
})();
