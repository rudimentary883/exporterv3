/**
 * Name tags with Code 128 check-in barcodes in Google Slides (and a barcode column for Google Sheets).
 *
 * THE BARCODE: the same as the one the tab site draws with JsBarcode (format CODE128): a 6-digit number becomes Code 128 subset C =
 * start + 3 symbols + checksum + stop = 68 modules x 5 px = 340 px wide. Bars are 75 px high and the identifier is written under
 * them (monospace 20 px), so the picture is 340 x 100 px (= 255 x 75 pt on a slide). It is generated here, in the script - no
 * internet, no add-on - as a PNG and inserted into the slide.
 *
 * WHY PNG AND NOT SVG: Google Slides (insertImage) only accepts PNG, JPEG or GIF; an SVG cannot be inserted. A PNG is a picture of
 * fixed pixels, so for print raise PIXEL_SCALE (3 = 1020 x 300 px, still 255 x 75 pt on the slide) to keep the bars razor sharp.
 *
 * run()                 name tags in Slides. SOURCE: the first sheet of the spreadsheet this script is attached to, with a column for
 *                       the name and one for the six-digit identifier (extra columns such as role / institution are filled into
 *                       {{role}}, {{institution}} ... automatically):
 *                           name                 identifier   role      institution
 *                           Aadi Rohit Sardesai  356452       Debater   BINUS BBJ
 *                       TEMPLATE: one slide (the first one that contains {{barcode}}) with {{name}} {{role}} {{institution}}
 *                       {{identifier}} text boxes and a rectangle containing the text {{barcode}}: the barcode is drawn exactly in
 *                       that box (make it 255 x 75 pt = 340 x 100 px, ratio 3.4 : 1) and the box is removed.
 * fillBarcodeColumn()   Google SHEETS: puts the barcode picture IN each cell of a "barcode" column (Insert > Image > In cell).
 *                       Then File > Download > Microsoft Excel (.xlsx) gives the file Canva's Bulk create reads (Canva's help describes
 *                       exactly this route for Google Sheets).
 *
 * A long list is done in several runs: run() stops before Google's 6-minute limit and the next run() continues where it stopped
 * (AUTO_RESUME). resetProgress() starts again from the last row.
 */

var SLIDES_ID = 'PASTE_THE_ID_OF_YOUR_NAME_TAG_DECK_HERE';   // the long code in the deck's address: /presentation/d/<ID>/edit

// ----- settings ------------------------------------------------------------------------------------------------------
var TEMPLATE_SLIDE_INDEX = -1;         // -1 = automatic (first slide containing {{barcode}}); 0 = first slide, 1 = second ...
var BARCODE_PLACEHOLDER = '{{barcode}}';
var SHOW_NUMBER = true;                // the identifier under the bars (picture 340 x 100 px). false = bars only (340 x 75 px)
var PIXEL_SCALE = 1;                   // 1 = 340 x 100 px picture. 3 = 1020 x 300 px (same barcode, sharper when printed)
var QUIET_ZONE_PX = 0;                 // white margin left and right inside the picture (the tab site uses 10)
var BARCODE_WIDTH_PT = 255;            // used only when the template has no {{barcode}} box: 340 x 100 px = 255 x 75 pt
var BARCODE_HEIGHT_PT = 75;
var DEFAULT_LEFT_PT = 400;
var DEFAULT_TOP_PT = 30;
var SET_ALT_TEXT = false;              // true: title / description on each barcode picture (2 extra calls per slide)
var CLEAR_PREVIOUS_SLIDES = false;     // true: delete every slide except the template before building
var DELETE_TEMPLATE_WHEN_DONE = false; // true: delete the template slide when everything is finished
var FIRST_ROW = 2;                     // first sheet row to use (2 = right under the header)
var LAST_ROW = 0;                      // last sheet row to use (0 = all)
var AUTO_RESUME = true;                // continue after a time-limit stop (progress is kept in script properties)
var MAX_RUNTIME_MS = 5 * 60 * 1000;    // stop cleanly before the 6-minute limit
var ID_LENGTH = 6;
var BARCODE_COLUMN_HEADER = 'barcode'; // fillBarcodeColumn(): the header of the column that receives the pictures

var ID_HEADERS = ['identifier', 'identifiers', 'id', 'barcode', 'code', 'number', 'check-in', 'checkin'];
var NAME_HEADERS = ['name', 'full name', 'participant', 'speaker', 'adjudicator', 'person', 'debater'];
// placeholder -> columns that can fill it (first one found wins)
var ALIASES = {
  'name': NAME_HEADERS,
  'identifier': ID_HEADERS,
  'id': ID_HEADERS,
  'number': ID_HEADERS
};


function run() {
  var started = new Date().getTime();
  var props = PropertiesService.getScriptProperties();
  Logger.log('START');

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) { Logger.log('No data rows found.'); return; }

  var headers = values[0].map(function (h) { return String(h).toLowerCase().trim(); });
  var colIndex = {};
  headers.forEach(function (h, i) { if (h && colIndex[h] === undefined) colIndex[h] = i; });
  var idCol = firstColumn_(ID_HEADERS, colIndex);
  if (idCol === -1) { Logger.log('ERROR: no identifier column. Use a header such as: ' + ID_HEADERS.join(', ')); return; }
  Logger.log('Headers: ' + headers.join(', ') + ' | identifier column: ' + headers[idCol]);

  var pres = SlidesApp.openById(SLIDES_ID);
  var slides = pres.getSlides();
  var template = findTemplate_(slides);
  if (!template) { Logger.log('ERROR: template slide not found.'); return; }

  // text placeholders used on the template, and the box that marks the barcode position
  var tokens = findPlaceholders_(template);
  var plan = [];
  tokens.forEach(function (t) {
    if (t.token.toLowerCase() === BARCODE_PLACEHOLDER.toLowerCase()) return;
    var col = resolveColumn_(t.name, colIndex);
    if (col === -1) Logger.log('WARNING: ' + t.token + ' has no matching column (left as it is)');
    else plan.push({token: t.token, col: col});
  });
  var box = findBarcodeBox_(template);
  Logger.log(box ? 'Barcode box on the template: ' + box.width.toFixed(1) + ' x ' + box.height.toFixed(1) + ' pt'
                 : 'No ' + BARCODE_PLACEHOLDER + ' box on the template: using the default position and size');
  var geo = box || {left: DEFAULT_LEFT_PT, top: DEFAULT_TOP_PT, width: BARCODE_WIDTH_PT, height: BARCODE_HEIGHT_PT, index: -1};

  if (CLEAR_PREVIOUS_SLIDES && !props.getProperty('BARCODE_LAST_ROW')) {
    slides.forEach(function (s) { if (s.getObjectId() !== template.getObjectId()) s.remove(); });
    Logger.log('Previous slides removed');
  }

  var first = Math.max(FIRST_ROW, 2) - 1;                                   // index in values[]
  var lastRow = LAST_ROW > 0 ? Math.min(LAST_ROW, values.length) : values.length;
  var pending = AUTO_RESUME ? props.getProperty('BARCODE_LAST_ROW') : null;
  if (pending) {
    lastRow = Math.min(lastRow, parseInt(pending, 10));
    Logger.log('Continuing: rows ' + FIRST_ROW + ' to ' + lastRow);
  }
  var last = lastRow - 1;
  var count = 0, skipped = 0, stoppedAt = -1;

  // Bottom-up: every copy goes straight after the template, so the final order is the order of the sheet.
  for (var i = last; i >= first; i--) {
    if (new Date().getTime() - started > MAX_RUNTIME_MS) { stoppedAt = i; break; }
    var row = values[i];
    if (isBlankRow_(row)) continue;

    var ident = cleanIdentifier_(row[idCol]);
    var name = String(row[firstColumn_(NAME_HEADERS, colIndex)] === undefined ? '' : row[firstColumn_(NAME_HEADERS, colIndex)]).trim();
    if (!ident) {
      Logger.log('SKIPPED row ' + (i + 1) + ' (' + (name || 'no name') + '): "' + row[idCol] + '" is not a ' + ID_LENGTH + '-digit identifier');
      skipped++;
      continue;
    }

    var slide = template.duplicate();
    plan.forEach(function (p) {
      var text = p.col === idCol ? ident : cellText_(row[p.col]);
      slide.replaceAllText(p.token, text, false);
    });

    var holder = findBoxOn_(slide, geo.index);                 // the {{barcode}} placeholder on the copy
    if (holder) holder.remove();
    var blob = Utilities.newBlob(barcodePng_(ident, PIXEL_SCALE, QUIET_ZONE_PX, SHOW_NUMBER), 'image/png', ident + '.png');
    var image = slide.insertImage(blob, geo.left, geo.top, geo.width, geo.height);
    if (SET_ALT_TEXT) { image.setTitle(name || ident); image.setDescription('Check-in barcode ' + ident); }

    count++;
    if (count % 20 === 0) Logger.log('...' + count + ' slides');
  }

  if (stoppedAt !== -1) {
    props.setProperty('BARCODE_LAST_ROW', String(stoppedAt + 1));
    Logger.log('STOPPED at the time limit after ' + count + ' slides. Run run() again: it continues with rows ' + FIRST_ROW + ' to ' + (stoppedAt + 1) + '.');
    return;
  }
  props.deleteProperty('BARCODE_LAST_ROW');
  if (DELETE_TEMPLATE_WHEN_DONE) template.remove();
  Logger.log('DONE: ' + count + ' slides' + (skipped ? ', ' + skipped + ' rows skipped (see SKIPPED above)' : ''));
}

/** Forget a stopped run: the next run() starts again from the last row of the sheet. */
function resetProgress() {
  PropertiesService.getScriptProperties().deleteProperty('BARCODE_LAST_ROW');
  Logger.log('Progress cleared.');
}



// ===== Code 128 (subset C) exactly as JsBarcode draws it ====================================================================
// symbol patterns 0..105 (11 modules) and 106 = stop (13 modules): from the JsBarcode v3.11.5 source
var BARS_ = [
  '11011001100', '11001101100', '11001100110', '10010011000',
  '10010001100', '10001001100', '10011001000', '10011000100',
  '10001100100', '11001001000', '11001000100', '11000100100',
  '10110011100', '10011011100', '10011001110', '10111001100',
  '10011101100', '10011100110', '11001110010', '11001011100',
  '11001001110', '11011100100', '11001110100', '11101101110',
  '11101001100', '11100101100', '11100100110', '11101100100',
  '11100110100', '11100110010', '11011011000', '11011000110',
  '11000110110', '10100011000', '10001011000', '10001000110',
  '10110001000', '10001101000', '10001100010', '11010001000',
  '11000101000', '11000100010', '10110111000', '10110001110',
  '10001101110', '10111011000', '10111000110', '10001110110',
  '11101110110', '11010001110', '11000101110', '11011101000',
  '11011100010', '11011101110', '11101011000', '11101000110',
  '11100010110', '11101101000', '11101100010', '11100011010',
  '11101111010', '11001000010', '11110001010', '10100110000',
  '10100001100', '10010110000', '10010000110', '10000101100',
  '10000100110', '10110010000', '10110000100', '10011010000',
  '10011000010', '10000110100', '10000110010', '11000010010',
  '11001010000', '11110111010', '11000010100', '10001111010',
  '10100111100', '10010111100', '10010011110', '10111100100',
  '10011110100', '10011110010', '11110100100', '11110010100',
  '11110010010', '11011011110', '11011110110', '11110110110',
  '10101111000', '10100011110', '10001011110', '10111101000',
  '10111100010', '11110101000', '11110100010', '10111011110',
  '10111101110', '11101011110', '11110101110', '11010000100',
  '11010010000', '11010011100', '1100011101011'
];
var MODULE_PX_ = 5;        // jsbarcode-width="5"
var BAR_HEIGHT_PX_ = 75;   // bars (the tab site uses 50)
var LABEL_HEIGHT_PX_ = 25; // the number under the bars
var LABEL_FONT_PX_ = 20;   // monospace 20 px
var LABEL_BASELINE_PX_ = 95;

// ----- digit outlines: SVG-like paths in 1/512 em, y down, origin = left of the glyph on the baseline (DejaVu Sans Mono) -----
var GLYPH_EM_ = 512;
var GLYPH_ADVANCE_ = 308;
var GLYPHS_ = {
  '0': 'M121 -188Q121 -201 130 -211Q140 -221 154 -221Q168 -221 178 -211Q188 -201 188 -188Q188 -174 178 -164Q168 -154 154 -154Q140 -154 130 -164Q121 -173 121 -188ZM154 -340Q119 -340 101 -302Q84 -264 84 -186Q84 -109 101 -71Q119 -33 154 -33Q190 -33 207 -71Q224 -109 224 -186Q224 -264 207 -302Q190 -340 154 -340ZM154 -380Q214 -380 244 -331Q275 -282 275 -186Q275 -91 244 -42Q214 7 154 7Q94 7 64 -42Q33 -91 33 -186Q33 -282 64 -331Q94 -380 154 -380Z',
  '1': 'M68 -42L146 -42L146 -328L62 -309L62 -355L146 -373L196 -373L196 -42L274 -42L274 0L68 0Z',
  '2': 'M93 -42L265 -42L265 0L38 0L38 -42Q85 -92 120 -130Q155 -167 168 -183Q193 -213 202 -232Q210 -251 210 -271Q210 -302 192 -320Q174 -338 142 -338Q119 -338 94 -329Q69 -321 41 -304L41 -355Q67 -368 92 -374Q116 -380 141 -380Q196 -380 229 -351Q262 -322 262 -274Q262 -250 251 -226Q240 -202 215 -174Q201 -157 174 -128Q148 -100 93 -42Z',
  '3': 'M194 -200Q231 -190 250 -165Q270 -140 270 -103Q270 -52 235 -22Q200 7 139 7Q114 7 87 2Q60 -2 34 -11L34 -62Q60 -48 84 -42Q109 -35 134 -35Q175 -35 198 -54Q220 -73 220 -108Q220 -140 198 -160Q175 -179 137 -179L99 -179L99 -220L137 -220Q172 -220 192 -236Q211 -251 211 -278Q211 -307 193 -322Q175 -338 141 -338Q119 -338 95 -332Q72 -328 46 -318L46 -364Q76 -372 99 -376Q123 -380 141 -380Q196 -380 228 -353Q261 -325 261 -280Q261 -249 244 -229Q227 -208 194 -200Z',
  '4': 'M184 -327L66 -130L184 -130ZM176 -373L234 -373L234 -130L284 -130L284 -89L234 -89L234 0L184 0L184 -89L26 -89L26 -137Z',
  '5': 'M52 -373L241 -373L241 -331L98 -331L98 -239Q108 -243 119 -245Q130 -247 141 -247Q199 -247 233 -212Q267 -178 267 -120Q267 -61 232 -27Q196 7 134 7Q104 7 80 3Q55 -1 36 -9L36 -60Q59 -48 82 -41Q105 -35 130 -35Q171 -35 194 -57Q216 -79 216 -120Q216 -160 193 -182Q170 -204 128 -204Q108 -204 88 -200Q69 -195 52 -186Z',
  '6': 'M248 -365L248 -318Q232 -328 214 -333Q196 -338 177 -338Q129 -338 104 -301Q80 -265 80 -195Q92 -220 113 -233Q134 -247 162 -247Q216 -247 245 -214Q275 -180 275 -120Q275 -59 244 -26Q214 7 159 7Q94 7 64 -39Q33 -86 33 -186Q33 -281 70 -330Q106 -380 175 -380Q194 -380 212 -376Q230 -372 248 -365ZM158 -207Q126 -207 107 -184Q88 -161 88 -120Q88 -79 107 -56Q126 -32 158 -32Q191 -32 208 -54Q225 -76 225 -120Q225 -163 208 -185Q191 -207 158 -207Z',
  '7': 'M35 -373L270 -373L270 -352L136 0L84 0L214 -331L35 -331Z',
  '8': 'M154 -177Q120 -177 102 -158Q84 -140 84 -105Q84 -70 102 -51Q121 -32 154 -32Q188 -32 206 -51Q225 -70 225 -105Q225 -139 206 -158Q188 -177 154 -177ZM110 -198Q78 -206 60 -229Q42 -252 42 -283Q42 -328 72 -354Q102 -380 154 -380Q206 -380 236 -354Q267 -328 267 -283Q267 -252 249 -229Q230 -206 198 -198Q236 -190 256 -165Q276 -140 276 -100Q276 -50 243 -21Q211 7 154 7Q97 7 65 -21Q33 -50 33 -100Q33 -140 53 -165Q72 -190 110 -198ZM92 -278Q92 -248 108 -233Q124 -217 154 -217Q184 -217 200 -233Q216 -248 216 -278Q216 -309 201 -325Q185 -341 154 -341Q124 -341 108 -325Q92 -309 92 -278Z',
  '9': 'M149 -166Q181 -166 200 -189Q218 -212 218 -253Q218 -294 200 -317Q181 -340 149 -340Q116 -340 98 -318Q82 -296 82 -253Q82 -210 98 -188Q115 -166 149 -166ZM59 -8L59 -54Q75 -45 92 -40Q110 -35 130 -35Q178 -35 202 -71Q227 -108 227 -178Q215 -153 194 -139Q172 -126 145 -126Q91 -126 61 -159Q32 -192 32 -254Q32 -314 62 -347Q92 -380 148 -380Q213 -380 243 -333Q274 -286 274 -186Q274 -92 237 -42Q201 7 132 7Q113 7 95 3Q76 0 59 -8Z'
};

var GLYPH_CACHE_ = {};
function glyphContours_(ch) {
  if (GLYPH_CACHE_[ch]) return GLYPH_CACHE_[ch];
  var contours = [], current = [];
  var re = /([MLQZ])([^MLQZ]*)/g, m;
  while ((m = re.exec(GLYPHS_[ch])) !== null) {
    var nums = (m[2].match(/-?\d+/g) || []).map(Number);
    if (m[1] === 'M') {
      if (current.length) contours.push(current);
      current = [[nums[0], nums[1]]];
    } else if (m[1] === 'L') {
      current.push([nums[0], nums[1]]);
    } else if (m[1] === 'Q') {
      var x0 = current[current.length - 1][0], y0 = current[current.length - 1][1];
      for (var i = 1; i <= 8; i++) {
        var t = i / 8, u = 1 - t;
        current.push([u * u * x0 + 2 * u * t * nums[0] + t * t * nums[2], u * u * y0 + 2 * u * t * nums[1] + t * t * nums[3]]);
      }
    }
  }
  if (current.length) contours.push(current);
  GLYPH_CACHE_[ch] = contours;
  return contours;
}

/** '356452' -> module string for start C + 3 pairs + checksum + stop (68 modules). */
function code128c_(digits) {
  var values = [];
  for (var i = 0; i < digits.length; i += 2) values.push(parseInt(digits.substr(i, 2), 10));
  var sum = 105;
  for (var k = 0; k < values.length; k++) sum += values[k] * (k + 1);
  var symbols = [105].concat(values, [sum % 103, 106]);
  return symbols.map(function (s) { return BARS_[s]; }).join('');
}

/** The label (all digits) as polygons in picture pixels, centred. */
function labelPolygons_(text, scale, width) {
  var unit = LABEL_FONT_PX_ * scale / GLYPH_EM_;
  var left = (width - text.length * GLYPH_ADVANCE_ * unit) / 2;
  var base = LABEL_BASELINE_PX_ * scale;
  var polys = [];
  for (var i = 0; i < text.length; i++) {
    var ox = left + i * GLYPH_ADVANCE_ * unit;
    glyphContours_(text.charAt(i)).forEach(function (contour) {
      polys.push(contour.map(function (p) { return [ox + p[0] * unit, base + p[1] * unit]; }));
    });
  }
  return polys;
}

/** Anti-aliased coverage 0..1 (even-odd) for pixel rows yTop..yBottom-1: array of rows. */
function labelCoverage_(polys, width, yTop, yBottom) {
  var SUB = 4, edges = [];
  polys.forEach(function (poly) {
    for (var k = 0; k < poly.length; k++) {
      var a = poly[k], b = poly[(k + 1) % poly.length];
      var x1 = a[0], y1 = a[1], x2 = b[0], y2 = b[1];
      if (y1 !== y2) {
        if (y1 > y2) { var tx = x1, ty = y1; x1 = x2; y1 = y2; x2 = tx; y2 = ty; }
        edges.push([y1, y2, x1, (x2 - x1) / (y2 - y1)]);
      }
    }
  });
  var rows = [];
  for (var row = yTop; row < yBottom; row++) {
    var cov = [];
    for (var c = 0; c < width; c++) cov.push(0);
    for (var sub = 0; sub < SUB; sub++) {
      var ys = row + (sub + 0.5) / SUB;
      var xs = [];
      for (var e = 0; e < edges.length; e++) if (edges[e][0] <= ys && ys < edges[e][1]) xs.push(edges[e][2] + (ys - edges[e][0]) * edges[e][3]);
      xs.sort(function (p, q) { return p - q; });
      for (var j = 0; j + 1 < xs.length; j += 2) {
        var xa = xs[j], xb = xs[j + 1], px = Math.floor(xa);
        while (px < xb && px < width) {
          if (px >= 0) cov[px] += (Math.min(px + 1, xb) - Math.max(px, xa)) / SUB;
          px++;
        }
      }
    }
    rows.push(cov);
  }
  return rows;
}

/** PNG (2-bit grey: pure black bars, smooth digits) as an array of signed bytes for Utilities.newBlob(). 340 x 100 px at scale 1. */
function barcodePng_(digits, scale, quietPx, label) {
  var bits = code128c_(digits);
  scale = Math.max(parseInt(scale, 10) || 1, 1);
  var quiet = Math.max(parseInt(quietPx, 10) || 0, 0) * scale;
  var perModule = MODULE_PX_ * scale;
  var width = bits.length * perModule + 2 * quiet;
  var barH = BAR_HEIGHT_PX_ * scale;
  var height = (BAR_HEIGHT_PX_ + (label ? LABEL_HEIGHT_PX_ : 0)) * scale;
  var rowBytes = Math.ceil(width / 4);

  function pack(levels) {                                   // 2 bits per pixel, first pixel in the top bits; 3 = white, 0 = black
    var out = [0];                                          // PNG filter byte 0
    for (var b = 0; b < rowBytes; b++) out.push(0);
    for (var x = 0; x < levels.length; x++) out[1 + (x >> 2)] |= levels[x] << (6 - 2 * (x & 3));
    return out;
  }

  var barLevels = [];
  for (var x = 0; x < width; x++) barLevels.push(3);
  for (var m = 0; m < bits.length; m++) {
    if (bits.charAt(m) === '1') for (var px = quiet + m * perModule; px < quiet + (m + 1) * perModule; px++) barLevels[px] = 0;
  }
  var barRow = pack(barLevels);
  var raw = [];
  for (var y = 0; y < barH; y++) Array.prototype.push.apply(raw, barRow);
  if (label) {
    var cover = labelCoverage_(labelPolygons_(String(digits), scale, width), width, barH, height);
    cover.forEach(function (cov) {
      Array.prototype.push.apply(raw, pack(cov.map(function (c) { return 3 - Math.min(3, Math.floor(c * 3 + 0.5)); })));
    });
  }

  var z = [0x78, 0x01];                                     // zlib header + "stored" deflate blocks (no compression needed)
  for (var pos = 0; pos < raw.length; pos += 65535) {
    var len = Math.min(65535, raw.length - pos);
    var final = pos + len >= raw.length ? 1 : 0;
    z.push(final, len & 255, len >> 8, (~len) & 255, ((~len) >> 8) & 255);
    for (var q = pos; q < pos + len; q++) z.push(raw[q]);
  }
  var ad = adler32_(raw);
  z.push((ad >>> 24) & 255, (ad >>> 16) & 255, (ad >>> 8) & 255, ad & 255);

  var png = [137, 80, 78, 71, 13, 10, 26, 10];
  pushChunk_(png, 'IHDR', [].concat(be32_(width), be32_(height), [2, 0, 0, 0, 0]));   // 2-bit greyscale
  pushChunk_(png, 'IDAT', z);
  pushChunk_(png, 'IEND', []);
  return png.map(function (v) { return v > 127 ? v - 256 : v; });
}

function be32_(n) { return [(n >>> 24) & 255, (n >>> 16) & 255, (n >>> 8) & 255, n & 255]; }

function pushChunk_(out, type, data) {
  var body = [];
  for (var i = 0; i < 4; i++) body.push(type.charCodeAt(i));
  for (var d = 0; d < data.length; d++) body.push(data[d]);
  var crc = crc32_(body);
  Array.prototype.push.apply(out, be32_(data.length));
  for (var k = 0; k < body.length; k++) out.push(body[k]);
  Array.prototype.push.apply(out, be32_(crc));
}

var CRC_TABLE_ = null;
function crc32_(bytes) {
  if (!CRC_TABLE_) {
    CRC_TABLE_ = [];
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      CRC_TABLE_.push(c >>> 0);
    }
  }
  var crc = 0xFFFFFFFF;
  for (var i = 0; i < bytes.length; i++) crc = CRC_TABLE_[(crc ^ bytes[i]) & 255] ^ (crc >>> 8);
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

function adler32_(bytes) {
  var a = 1, b = 0;
  for (var i = 0; i < bytes.length; i++) { a = (a + bytes[i]) % 65521; b = (b + a) % 65521; }
  return ((b << 16) | a) >>> 0;
}


// ===== Google Sheets: barcode pictures inside the cells ======================================================================
/**
 * Fills a "barcode" column of the FIRST sheet with the barcode picture of every row (name + identifier columns as for run()).
 * The pictures are real "image in cell" pictures. Afterwards: File > Download > Microsoft Excel (.xlsx), and upload that file to
 * Canva's Bulk create (connect the barcode column to an image frame).
 */
function fillBarcodeColumn() {
  var started = new Date().getTime();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(function (h) { return String(h).toLowerCase().trim(); });
  var colIndex = {};
  headers.forEach(function (h, i) { if (h && colIndex[h] === undefined) colIndex[h] = i; });
  var idCol = firstColumn_(ID_HEADERS, colIndex);
  if (idCol === -1) { Logger.log('ERROR: no identifier column. Use a header such as: ' + ID_HEADERS.join(', ')); return; }
  var nameCol = firstColumn_(NAME_HEADERS, colIndex);

  var target = colIndex[BARCODE_COLUMN_HEADER.toLowerCase()];
  if (target === undefined) {                                  // no such column yet: add it after the last used column
    target = values[0].length;
    sheet.getRange(1, target + 1).setValue(BARCODE_COLUMN_HEADER);
  }
  var shownW = 340 + 2 * QUIET_ZONE_PX, shownH = SHOW_NUMBER ? 100 : 75;
  sheet.setColumnWidth(target + 1, shownW + 20);                // pixels; a little bigger than the picture
  if (values.length > 1) sheet.setRowHeights(2, values.length - 1, shownH + 10);

  var done = 0, skipped = 0;
  for (var i = 1; i < values.length; i++) {
    if (new Date().getTime() - started > MAX_RUNTIME_MS) { Logger.log('STOPPED at row ' + (i + 1) + ' (time limit). Run fillBarcodeColumn() again.'); break; }
    if (isBlankRow_(values[i])) continue;
    var ident = cleanIdentifier_(values[i][idCol]);
    if (!ident) { Logger.log('SKIPPED row ' + (i + 1) + ': "' + values[i][idCol] + '" is not a ' + ID_LENGTH + '-digit identifier'); skipped++; continue; }
    var name = nameCol === -1 ? '' : cellText_(values[i][nameCol]);
    var uri = 'data:image/png;base64,' + Utilities.base64Encode(barcodePng_(ident, PIXEL_SCALE, QUIET_ZONE_PX, SHOW_NUMBER));
    var image = SpreadsheetApp.newCellImage().setSourceUrl(uri).setAltTextTitle('Barcode ' + ident).setAltTextDescription(name).build();
    sheet.getRange(i + 1, target + 1).setValue(image);
    done++;
  }
  Logger.log('DONE: ' + done + ' barcode pictures' + (skipped ? ', ' + skipped + ' rows skipped' : ''));
}


// ===== helpers ===========================================================================================================

/** 356452 (number or text), 12345 (the sheet dropped a leading zero) -> '356452' / '012345'. '' if it is not a 6-digit number. */
function cleanIdentifier_(value) {
  if (value === null || value === undefined) return '';
  var text = (typeof value === 'number') ? String(Math.round(value)) : String(value).trim();
  text = text.replace(/\.0+$/, '');
  if (!/^\d+$/.test(text)) return '';
  while (text.length < ID_LENGTH) text = '0' + text;
  return text.length === ID_LENGTH ? text : '';
}

function cellText_(value) { return (value === null || value === undefined) ? '' : String(value).trim(); }

function isBlankRow_(row) {
  for (var c = 0; c < row.length; c++) if (row[c] !== null && row[c] !== undefined && String(row[c]).trim() !== '') return false;
  return true;
}

function firstColumn_(names, colIndex) {
  for (var k = 0; k < names.length; k++) if (colIndex[names[k]] !== undefined) return colIndex[names[k]];
  return -1;
}

function resolveColumn_(name, colIndex) {
  if (colIndex[name] !== undefined) return colIndex[name];
  return ALIASES[name] ? firstColumn_(ALIASES[name], colIndex) : -1;
}

function findTemplate_(slides) {
  if (TEMPLATE_SLIDE_INDEX >= 0) return slides[TEMPLATE_SLIDE_INDEX];
  for (var i = 0; i < slides.length; i++) {
    var texts = [];
    collectText_(slides[i].getPageElements(), texts);
    if (texts.join('\n').toLowerCase().indexOf(BARCODE_PLACEHOLDER.toLowerCase()) !== -1) return slides[i];
  }
  Logger.log('No slide contains ' + BARCODE_PLACEHOLDER + '; using the first slide as the template.');
  return slides[0];
}

function findPlaceholders_(slide) {
  var texts = [];
  collectText_(slide.getPageElements(), texts);
  var seen = {}, found = [];
  texts.join('\n').replace(/\{\{\s*([^{}]+?)\s*\}\}/g, function (token, name) {
    if (!seen[token.toLowerCase()]) { seen[token.toLowerCase()] = true; found.push({token: token, name: name.toLowerCase().trim()}); }
    return token;
  });
  return found;
}

function collectText_(elements, out) {
  for (var k = 0; k < elements.length; k++) {
    var el = elements[k], type = el.getPageElementType();
    if (type === SlidesApp.PageElementType.SHAPE) out.push(el.asShape().getText().asString());
    else if (type === SlidesApp.PageElementType.TABLE) {
      var table = el.asTable();
      for (var r = 0; r < table.getNumRows(); r++) for (var c = 0; c < table.getNumColumns(); c++) out.push(table.getCell(r, c).getText().asString());
    } else if (type === SlidesApp.PageElementType.GROUP) collectText_(el.asGroup().getChildren(), out);
  }
}

function isBarcodeBox_(el) {
  return el.getPageElementType() === SlidesApp.PageElementType.SHAPE &&
         el.asShape().getText().asString().toLowerCase().indexOf(BARCODE_PLACEHOLDER.toLowerCase()) !== -1;
}

/** Position and size of the {{barcode}} box on the template (top level shapes). */
function findBoxOn_(slide, hintIndex) {
  var els = slide.getPageElements();
  if (hintIndex >= 0 && hintIndex < els.length && isBarcodeBox_(els[hintIndex])) return els[hintIndex];
  for (var i = 0; i < els.length; i++) if (isBarcodeBox_(els[i])) return els[i];
  return null;
}

function findBarcodeBox_(template) {
  var els = template.getPageElements();
  for (var i = 0; i < els.length; i++) {
    if (isBarcodeBox_(els[i])) {
      return {left: els[i].getLeft(), top: els[i].getTop(), width: els[i].getWidth(), height: els[i].getHeight(), index: i};
    }
  }
  return null;
}
