"""
barcodes.py - Code 128 check-in barcodes, drawn like the ones on the Tabbycat check-in print pages, with the number underneath.

The tab site draws them with JsBarcode v3.11.5 (MIT), format "auto" = CODE128:
    <svg class="barcode-placeholder" jsbarcode-value="356452" jsbarcode-width="5" jsbarcode-height="50">
A six-digit number is always Code 128 subset C (two digits per symbol):
    start C + 3 symbols + checksum + stop = 68 modules x 5 px = 340 px wide.
Here the bars are 75 px high and the number is written under them (monospace, 20 px): the picture is 340 x 100 px.
The symbol table is the one in JsBarcode's source (checked against the ISO/IEC 15417 table).
The digits are vector outlines of DejaVu Sans Mono (free licence, Bitstream Vera / DejaVu), so the PNG, the SVG and the Apps Script
all draw exactly the same label, at any size.

  * encode_code128c()  -> the 0/1 module string + the symbol values
  * barcode_png()      -> PNG (2-bit grey, anti-aliased label, pure python)      340 x 100 px at scale 1
  * barcode_svg()      -> SVG (bars + label as paths, no font needed)
  * parse_people()     -> name + identifier rows from an uploaded file or pasted text
  * build_xlsx()       -> XLSX with the barcode picture in each row ("anchored" inside the cell, or Excel 365 "in cell")
  * build_zip()        -> one <identifier>.png / .svg per person + names.csv
"""

import csv
import io
import re
import struct
import zipfile
import zlib

MODULE_PX = 5            # jsbarcode-width="5"
BAR_HEIGHT_PX = 75       # bars (the tab site uses 50)
LABEL_HEIGHT_PX = 25     # the number under the bars
LABEL_FONT_PX = 20       # monospace 20 px, like JsBarcode's label
LABEL_BASELINE_PX = 95   # baseline of the digits (5 px above the bottom edge)
ID_LENGTH = 6

BARS = [
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
    '11010010000', '11010011100', '1100011101011',
]
STOP = 106
START_C = 105

# ---- digit outlines: SVG-like paths in 1/512 em, y down, origin = left of the glyph on the baseline -----------------------
GLYPH_EM = 512
GLYPH_ADVANCE = 308
GLYPHS = {
    "0": "M121 -188Q121 -201 130 -211Q140 -221 154 -221Q168 -221 178 -211Q188 -201 188 -188Q188 -174 178 -164Q168 -154 154 -154Q140 -154 130 -164Q121 -173 121 -188ZM154 -340Q119 -340 101 -302Q84 -264 84 -186Q84 -109 101 -71Q119 -33 154 -33Q190 -33 207 -71Q224 -109 224 -186Q224 -264 207 -302Q190 -340 154 -340ZM154 -380Q214 -380 244 -331Q275 -282 275 -186Q275 -91 244 -42Q214 7 154 7Q94 7 64 -42Q33 -91 33 -186Q33 -282 64 -331Q94 -380 154 -380Z",
    "1": "M68 -42L146 -42L146 -328L62 -309L62 -355L146 -373L196 -373L196 -42L274 -42L274 0L68 0Z",
    "2": "M93 -42L265 -42L265 0L38 0L38 -42Q85 -92 120 -130Q155 -167 168 -183Q193 -213 202 -232Q210 -251 210 -271Q210 -302 192 -320Q174 -338 142 -338Q119 -338 94 -329Q69 -321 41 -304L41 -355Q67 -368 92 -374Q116 -380 141 -380Q196 -380 229 -351Q262 -322 262 -274Q262 -250 251 -226Q240 -202 215 -174Q201 -157 174 -128Q148 -100 93 -42Z",
    "3": "M194 -200Q231 -190 250 -165Q270 -140 270 -103Q270 -52 235 -22Q200 7 139 7Q114 7 87 2Q60 -2 34 -11L34 -62Q60 -48 84 -42Q109 -35 134 -35Q175 -35 198 -54Q220 -73 220 -108Q220 -140 198 -160Q175 -179 137 -179L99 -179L99 -220L137 -220Q172 -220 192 -236Q211 -251 211 -278Q211 -307 193 -322Q175 -338 141 -338Q119 -338 95 -332Q72 -328 46 -318L46 -364Q76 -372 99 -376Q123 -380 141 -380Q196 -380 228 -353Q261 -325 261 -280Q261 -249 244 -229Q227 -208 194 -200Z",
    "4": "M184 -327L66 -130L184 -130ZM176 -373L234 -373L234 -130L284 -130L284 -89L234 -89L234 0L184 0L184 -89L26 -89L26 -137Z",
    "5": "M52 -373L241 -373L241 -331L98 -331L98 -239Q108 -243 119 -245Q130 -247 141 -247Q199 -247 233 -212Q267 -178 267 -120Q267 -61 232 -27Q196 7 134 7Q104 7 80 3Q55 -1 36 -9L36 -60Q59 -48 82 -41Q105 -35 130 -35Q171 -35 194 -57Q216 -79 216 -120Q216 -160 193 -182Q170 -204 128 -204Q108 -204 88 -200Q69 -195 52 -186Z",
    "6": "M248 -365L248 -318Q232 -328 214 -333Q196 -338 177 -338Q129 -338 104 -301Q80 -265 80 -195Q92 -220 113 -233Q134 -247 162 -247Q216 -247 245 -214Q275 -180 275 -120Q275 -59 244 -26Q214 7 159 7Q94 7 64 -39Q33 -86 33 -186Q33 -281 70 -330Q106 -380 175 -380Q194 -380 212 -376Q230 -372 248 -365ZM158 -207Q126 -207 107 -184Q88 -161 88 -120Q88 -79 107 -56Q126 -32 158 -32Q191 -32 208 -54Q225 -76 225 -120Q225 -163 208 -185Q191 -207 158 -207Z",
    "7": "M35 -373L270 -373L270 -352L136 0L84 0L214 -331L35 -331Z",
    "8": "M154 -177Q120 -177 102 -158Q84 -140 84 -105Q84 -70 102 -51Q121 -32 154 -32Q188 -32 206 -51Q225 -70 225 -105Q225 -139 206 -158Q188 -177 154 -177ZM110 -198Q78 -206 60 -229Q42 -252 42 -283Q42 -328 72 -354Q102 -380 154 -380Q206 -380 236 -354Q267 -328 267 -283Q267 -252 249 -229Q230 -206 198 -198Q236 -190 256 -165Q276 -140 276 -100Q276 -50 243 -21Q211 7 154 7Q97 7 65 -21Q33 -50 33 -100Q33 -140 53 -165Q72 -190 110 -198ZM92 -278Q92 -248 108 -233Q124 -217 154 -217Q184 -217 200 -233Q216 -248 216 -278Q216 -309 201 -325Q185 -341 154 -341Q124 -341 108 -325Q92 -309 92 -278Z",
    "9": "M149 -166Q181 -166 200 -189Q218 -212 218 -253Q218 -294 200 -317Q181 -340 149 -340Q116 -340 98 -318Q82 -296 82 -253Q82 -210 98 -188Q115 -166 149 -166ZM59 -8L59 -54Q75 -45 92 -40Q110 -35 130 -35Q178 -35 202 -71Q227 -108 227 -178Q215 -153 194 -139Q172 -126 145 -126Q91 -126 61 -159Q32 -192 32 -254Q32 -314 62 -347Q92 -380 148 -380Q213 -380 243 -333Q274 -286 274 -186Q274 -92 237 -42Q201 7 132 7Q113 7 95 3Q76 0 59 -8Z",
}

_PATH_TOKEN = re.compile(r"([MLQZ])([^MLQZ]*)")
_FLATTEN_STEPS = 8
_GLYPH_CACHE = {}


def _glyph_contours(ch):
    """'5' -> list of contours, each a list of (x, y) in 1/512 em (curves flattened into short lines)."""
    if ch in _GLYPH_CACHE:
        return _GLYPH_CACHE[ch]
    contours, current = [], []
    for cmd, args in _PATH_TOKEN.findall(GLYPHS[ch]):
        nums = [float(n) for n in re.findall(r"-?\d+", args)]
        if cmd == "M":
            if current:
                contours.append(current)
            current = [(nums[0], nums[1])]
        elif cmd == "L":
            current.append((nums[0], nums[1]))
        elif cmd == "Q":
            x0, y0 = current[-1]
            cx, cy, x1, y1 = nums
            for i in range(1, _FLATTEN_STEPS + 1):
                t = i / _FLATTEN_STEPS
                u = 1 - t
                current.append((u * u * x0 + 2 * u * t * cx + t * t * x1, u * u * y0 + 2 * u * t * cy + t * t * y1))
    if current:
        contours.append(current)
    _GLYPH_CACHE[ch] = contours
    return contours


def encode_code128c(digits):
    """'356452' -> ([105, 35, 64, 52, 12, 106], '1101001110010...')  (Code 128, subset C)"""
    digits = str(digits)
    if not re.fullmatch(r"\d+", digits) or len(digits) % 2:
        raise ValueError(f"'{digits}' needs an even number of digits for Code 128 subset C")
    values = [int(digits[i:i + 2]) for i in range(0, len(digits), 2)]
    checksum = (START_C + sum(v * (i + 1) for i, v in enumerate(values))) % 103
    symbols = [START_C] + values + [checksum, STOP]
    return symbols, "".join(BARS[s] for s in symbols)


def _layout(identifier, scale, quiet_px, label):
    """Common geometry (in picture pixels at this scale)."""
    _, bits = encode_code128c(identifier)
    scale = max(int(scale), 1)
    quiet = max(int(quiet_px), 0) * scale
    px_module = MODULE_PX * scale
    width = len(bits) * px_module + 2 * quiet
    bar_h = BAR_HEIGHT_PX * scale
    height = (BAR_HEIGHT_PX + (LABEL_HEIGHT_PX if label else 0)) * scale
    return bits, scale, quiet, px_module, width, bar_h, height


def _label_polygons(text, scale, width):
    """The label as a list of polygons in picture pixels (all digits), centred on the picture."""
    unit = LABEL_FONT_PX * scale / GLYPH_EM
    total = len(text) * GLYPH_ADVANCE * unit
    left = (width - total) / 2.0
    base = LABEL_BASELINE_PX * scale
    polygons = []
    for i, ch in enumerate(text):
        ox = left + i * GLYPH_ADVANCE * unit
        for contour in _glyph_contours(ch):
            polygons.append([(ox + x * unit, base + y * unit) for x, y in contour])
    return polygons


_SUB_ROWS = 4


def _label_coverage(polygons, width, y_top, y_bottom):
    """Anti-aliased coverage 0..1 of the polygons (even-odd) for the pixel rows y_top..y_bottom-1: list of float rows."""
    edges = []
    for poly in polygons:
        n = len(poly)
        for k in range(n):
            (x1, y1), (x2, y2) = poly[k], poly[(k + 1) % n]
            if y1 != y2:
                if y1 > y2:
                    x1, y1, x2, y2 = x2, y2, x1, y1
                edges.append((y1, y2, x1, (x2 - x1) / (y2 - y1)))
    rows = []
    for row in range(y_top, y_bottom):
        cov = [0.0] * width
        for sub in range(_SUB_ROWS):
            ys = row + (sub + 0.5) / _SUB_ROWS
            xs = sorted(x1 + (ys - y1) * slope for (y1, y2, x1, slope) in edges if y1 <= ys < y2)
            for j in range(0, len(xs) - 1, 2):
                xa, xb = xs[j], xs[j + 1]
                px = int(xa // 1)
                while px < xb and px < width:
                    if px >= 0:
                        cov[px] += (min(px + 1, xb) - max(px, xa)) / _SUB_ROWS
                    px += 1
        rows.append(cov)
    return rows


def _chunk(kind, data):
    body = kind + data
    return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)


def barcode_png(identifier, scale=1, quiet_px=0, label=True):
    """
    PNG: black bars (75 px) on white with the identifier under them (25 px) = 340 x 100 px at scale 1.
    2-bit greyscale (4 shades) so the digits are smooth while the bars stay pure black / white.
    scale   : 2 -> 680 x 200 px (same picture, sharper when printed)
    quiet_px: white margin left/right (JsBarcode draws 10 px; 0 keeps the picture at exactly 340 px)
    label   : False -> bars only (340 x 75 px)
    """
    bits, scale, quiet, px_module, width, bar_h, height = _layout(identifier, scale, quiet_px, label)
    row_bytes = (width + 3) // 4

    def pack(levels):                                            # 2 bits per pixel, first pixel in the top bits
        out = bytearray(row_bytes)
        for x, level in enumerate(levels):
            out[x >> 2] |= level << (6 - 2 * (x & 3))
        return bytes(out)

    bar_levels = [3] * width                                    # 3 = white, 0 = black
    for m, bit in enumerate(bits):
        if bit == "1":
            for x in range(quiet + m * px_module, quiet + (m + 1) * px_module):
                bar_levels[x] = 0
    bar_row = b"\x00" + pack(bar_levels)
    blank_row = b"\x00" + pack([3] * width)
    raw = [bar_row] * bar_h
    if label:
        coverage = _label_coverage(_label_polygons(str(identifier), scale, width), width, bar_h, height)
        for cov in coverage:
            raw.append(b"\x00" + pack([3 - min(3, int(c * 3 + 0.5)) for c in cov]))
    ihdr = struct.pack(">IIBBBBB", width, height, 2, 0, 0, 0, 0)   # 2-bit greyscale
    return (b"\x89PNG\r\n\x1a\n" + _chunk(b"IHDR", ihdr) + _chunk(b"IDAT", zlib.compress(b"".join(raw), 9)) + _chunk(b"IEND", b""))


def _num(v):
    text = f"{v:.2f}".rstrip("0").rstrip(".")
    return "0" if text in ("", "-0") else text


def _label_svg_path(text, width):
    """The label as one SVG path (true quadratic curves from the font), centred, in viewBox pixels."""
    unit = LABEL_FONT_PX / GLYPH_EM
    left = (width - len(text) * GLYPH_ADVANCE * unit) / 2.0
    d = []
    for i, ch in enumerate(text):
        ox = left + i * GLYPH_ADVANCE * unit
        for cmd, args in _PATH_TOKEN.findall(GLYPHS[ch]):
            nums = [float(n) for n in re.findall(r"-?\d+", args)]
            pts = [f"{_num(ox + nums[k] * unit)} {_num(LABEL_BASELINE_PX + nums[k + 1] * unit)}" for k in range(0, len(nums), 2)]
            d.append(cmd + " ".join(pts))
    return "".join(d)


def barcode_svg(identifier, scale=1, quiet_px=0, label=True):
    """
    SVG (bytes): white background, the bars and the digits as filled paths (no font needed).
    viewBox is 340 x 100 (360 x 100 with the margin); width / height are multiplied by `scale` - it is a vector, so any size is sharp.
    """
    bits, _, quiet, px_module, width, bar_h, height = _layout(identifier, 1, quiet_px, label)
    scale = max(int(scale), 1)
    bars, m = [], 0
    while m < len(bits):
        if bits[m] == "1":
            n = m
            while n < len(bits) and bits[n] == "1":
                n += 1
            bars.append(f"M{quiet + m * px_module} 0h{(n - m) * px_module}v{bar_h}h-{(n - m) * px_module}z")
            m = n
        else:
            m += 1
    svg = [f'<svg xmlns="http://www.w3.org/2000/svg" version="1.1" width="{width * scale}" height="{height * scale}" viewBox="0 0 {width} {height}">',
           f'<rect width="{width}" height="{height}" fill="#ffffff"/>',
           f'<path d="{"".join(bars)}" fill="#000000" shape-rendering="crispEdges"/>']
    if label:
        svg.append(f'<path d="{_label_svg_path(str(identifier), width)}" fill="#000000" fill-rule="evenodd"/>')
    svg.append("</svg>")
    return "".join(svg).encode("utf-8")


def barcode_image(identifier, image_format="png", scale=1, quiet_px=0, label=True):
    return (barcode_svg if image_format == "svg" else barcode_png)(identifier, scale, quiet_px, label)


# ----------------------------------------------------------------------------------------------------------------
# people list: name + six-digit identifier
# ----------------------------------------------------------------------------------------------------------------
ID_HEADERS = ("identifier", "identifiers", "id", "barcode", "code", "number", "check-in", "checkin")
NAME_HEADERS = ("name", "participant", "speaker", "adjudicator", "person", "debater")


def clean_identifier(value):
    """356452, 356452.0, ' 356452 ', 12345 (Excel dropped a leading zero) -> '356452' / '012345'. '' if not a number."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".")[0]
    if not re.fullmatch(r"\d+", text):
        return ""
    return text.zfill(ID_LENGTH) if len(text) < ID_LENGTH else text


def _pick_columns(header):
    names = [str(h or "").strip().lower() for h in header]
    id_col = next((i for i, h in enumerate(names) if any(h == k or h.startswith(k) for k in ID_HEADERS)), None)
    name_col = next((i for i, h in enumerate(names) if i != id_col and any(k in h for k in NAME_HEADERS)), None)
    return name_col, id_col


def parse_people(rows):
    """
    rows: list of lists (a CSV / pasted table). The first row may be a header (name / identifier ...).
    Without a header the first column is the name and the second the identifier.
    Returns (people [(name, identifier)], problems [str], warnings [str]).
    """
    rows = [list(r) for r in rows if any(str(c or "").strip() for c in r)]
    if not rows:
        return [], ["The file / text is empty."], []
    first = rows[0]
    has_header = len(first) >= 2 and not clean_identifier(first[1]) and not clean_identifier(first[0])
    name_col, id_col = (0, 1)
    start = 0
    if has_header:
        picked_name, picked_id = _pick_columns(first)
        name_col = picked_name if picked_name is not None else 0
        id_col = picked_id if picked_id is not None else (1 if name_col == 0 else 0)
        start = 1
    people, problems, warnings, seen = [], [], [], {}
    for line_no, row in enumerate(rows[start:], start=start + 1):
        name = str(row[name_col]).strip() if name_col < len(row) and row[name_col] is not None else ""
        ident = clean_identifier(row[id_col]) if id_col < len(row) else ""
        if not ident or len(ident) != ID_LENGTH:
            problems.append(f"Row {line_no} ({name or 'no name'}): '{row[id_col] if id_col < len(row) else ''}' is not a {ID_LENGTH}-digit identifier")
            continue
        if ident in seen:
            warnings.append(f"Identifier {ident} is used by '{seen[ident]}' and '{name}'.")
        seen.setdefault(ident, name)
        people.append((name, ident))
    return people, problems, warnings


def rows_from_text(text):
    """Pasted text: tab separated (copied from a sheet) or comma separated, one person per line."""
    out = []
    for line in str(text or "").splitlines():
        if not line.strip():
            continue
        if "\t" in line:
            out.append([c.strip() for c in line.split("\t")])
        else:
            out.append([c.strip() for c in next(csv.reader([line]))])
    return out


# ----------------------------------------------------------------------------------------------------------------
# outputs
# ----------------------------------------------------------------------------------------------------------------
def build_xlsx(people, scale=1, quiet_px=0, style="anchored", label=True):
    """
    XLSX: name | identifier | barcode. The barcode PICTURES are PNG (spreadsheets cannot hold SVG pictures).
      style "anchored": a normal picture placed completely INSIDE the barcode cell (one cell, not spanning several) - works in
                        Excel of any version, in Google Sheets and is the layout other programs expect from Google-Sheets exports.
      style "incell"  : Excel 365 "Place in Cell" (rich value picture); other programs show #VALUE! in the cell.
    The picture is always shown 340 x 100 px (+ margin if asked); `scale` only adds pixels.
    """
    import xlsxwriter
    buffer = io.BytesIO()
    workbook = xlsxwriter.Workbook(buffer)          # NOT in_memory: that mode writes an invalid path for in-cell pictures
    sheet = workbook.add_worksheet("Barcodes")
    head = workbook.add_format({"bold": True})
    sheet.write_row(0, 0, ["name", "identifier", "barcode"], head)
    _, scale, quiet, _, width_px, _, height_px = _layout("000000", scale, quiet_px, label)
    shown_w, shown_h = width_px // scale, height_px // scale
    cell_w, cell_h = shown_w + 30, shown_h + 10                       # the cell is a little bigger than the picture
    sheet.set_landscape()
    sheet.fit_to_pages(1, 0)                                          # printing: all three columns on one page width
    sheet.set_column(0, 0, 36)
    sheet.set_column(1, 1, 14)
    sheet.set_column(2, 2, (cell_w - 5) / 7.0)                        # Excel: pixels = chars * 7 + 5
    for r, (name, ident) in enumerate(people, start=1):
        sheet.set_row(r, cell_h * 0.75)                               # points = pixels * 0.75
        sheet.write_string(r, 0, name)
        sheet.write_string(r, 1, ident)                                # text, so a leading zero is kept
        png = barcode_png(ident, scale, quiet_px, label)
        if style == "incell":
            sheet.embed_image(r, 2, f"{ident}.png", {"image_data": io.BytesIO(png), "description": f"Barcode {ident}"})
        else:
            sheet.insert_image(r, 2, f"{ident}.png", {
                "image_data": io.BytesIO(png), "x_offset": 15, "y_offset": 5, "x_scale": 1.0 / scale, "y_scale": 1.0 / scale,
                "object_position": 1, "description": f"Barcode {ident}"})
    workbook.close()
    return buffer.getvalue()


def build_zip(people, image_format="png", scale=1, quiet_px=0, label=True):
    ext = "svg" if image_format == "svg" else "png"
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        listing = io.StringIO()
        writer = csv.writer(listing)
        writer.writerow(["name", "identifier", "file"])
        for name, ident in people:
            archive.writestr(f"{ident}.{ext}", barcode_image(ident, ext, scale, quiet_px, label))
            writer.writerow([name, ident, f"{ident}.{ext}"])
        archive.writestr("names.csv", listing.getvalue())
    return buffer.getvalue()
