/**
 * Name tags with Code 128 check-in barcodes in Google Slides.
 *
 * SOURCE: the first sheet of the spreadsheet this script is attached to. It needs a column with the name and a column with
 * the six-digit identifier (the number printed under the barcode on the Tabbycat check-in page). Extra columns such as
 * role or institution are optional and are filled into {{role}}, {{institution}} ... automatically.
 *
 *     name              identifier     role       institution
 *     Aadi Rohit Sardesai   356452     Debater    BINUS BBJ
 *
 * TEMPLATE: one slide of the deck (the first slide that contains {{barcode}}) with
 *     {{name}} {{role}} {{institution}} {{identifier}}   text boxes (any of them, any column name works as {{column}})
 *     a rectangle / text box containing the text {{barcode}}   <- the barcode is drawn exactly in this box, then the box is removed
 * Make that box 340 x 50 px (= 255 x 37.5 pt, ratio 6.8 : 1) to get the barcode at the size it has on the tab site.
 *
 * The barcode is the same as the one the tab site draws with JsBarcode (format CODE128): a 6-digit number becomes
 * Code 128 subset C = start + 3 symbols + checksum + stop = 68 modules x 5 px = 340 px wide and 50 px high.
 * The picture is generated here, in the script (no internet, no add-on), as a PNG and inserted into the slide.
 *
 * Run run(). A long list is done in several runs: the script stops before Google's 6-minute limit and the next run()
 * continues where it stopped (see AUTO_RESUME). resetProgress() starts again from the last row.
 */

var SLIDES_ID = 'PASTE_THE_ID_OF_YOUR_NAME_TAG_DECK_HERE';   // the long code in the deck's address: /presentation/d/<ID>/edit

// ----- settings ------------------------------------------------------------------------------------------------------
var TEMPLATE_SLIDE_INDEX = -1;         // -1 = automatic (first slide containing {{barcode}}); 0 = first slide, 1 = second ...
var BARCODE_PLACEHOLDER = '{{barcode}}';
var PIXEL_SCALE = 1;                   // 1 = 340 x 50 px picture. 3 = 1020 x 150 px (same barcode, sharper when printed)
var QUIET_ZONE_PX = 0;                 // white margin left and right inside the picture (the tab site uses 10)
var BARCODE_WIDTH_PT = 255;            // used only when the template has no {{barcode}} box: 340 x 50 px = 255 x 37.5 pt
var BARCODE_HEIGHT_PT = 37.5;
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
    var blob = Utilities.newBlob(barcodePng_(ident, PIXEL_SCALE, QUIET_ZONE_PX), 'image/png', ident + '.png');
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


// ===== Code 128 (subset C) exactly as JsBarcode draws it =========================================================
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
var MODULE_PX_ = 5;      // jsbarcode-width="5"
var BAR_HEIGHT_PX_ = 50; // jsbarcode-height="50"

/** '356452' -> module string for start C + 3 pairs + checksum + stop (68 modules). */
function code128c_(digits) {
  var values = [];
  for (var i = 0; i < digits.length; i += 2) values.push(parseInt(digits.substr(i, 2), 10));
  var sum = 105;
  for (var k = 0; k < values.length; k++) sum += values[k] * (k + 1);
  var symbols = [105].concat(values, [sum % 103, 106]);
  return symbols.map(function (s) { return BARS_[s]; }).join('');
}

/** PNG (1-bit, black bars on white) as an array of signed bytes, ready for Utilities.newBlob(). 340 x 50 px at scale 1. */
function barcodePng_(digits, scale, quietPx) {
  var bits = code128c_(digits);
  scale = Math.max(parseInt(scale, 10) || 1, 1);
  var quiet = Math.max(parseInt(quietPx, 10) || 0, 0) * scale;
  var perModule = MODULE_PX_ * scale;
  var width = bits.length * perModule + 2 * quiet;
  var height = BAR_HEIGHT_PX_ * scale;
  var rowBytes = Math.ceil(width / 8);

  var row = [];                                             // 1 = white, 0 = black
  for (var b = 0; b < rowBytes; b++) row.push(255);
  for (var m = 0; m < bits.length; m++) {
    if (bits.charAt(m) === '1') {
      for (var x = quiet + m * perModule; x < quiet + (m + 1) * perModule; x++) row[x >> 3] &= (~(0x80 >> (x & 7))) & 255;
    }
  }
  var raw = [];                                             // every line: filter byte 0 + the packed row
  for (var y = 0; y < height; y++) { raw.push(0); for (var r = 0; r < rowBytes; r++) raw.push(row[r]); }

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
  pushChunk_(png, 'IHDR', [].concat(be32_(width), be32_(height), [1, 0, 0, 0, 0]));   // 1-bit greyscale
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
