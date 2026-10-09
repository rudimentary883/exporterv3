"""
barcodes.py - Code 128 check-in barcodes, drawn exactly like the ones on the Tabbycat check-in print pages.

The tab site draws them with JsBarcode v3.11.5 (MIT), format "auto" = CODE128:
    <svg class="barcode-placeholder" jsbarcode-value="356452" jsbarcode-width="5" jsbarcode-height="50">
A six-digit number is always encoded in Code 128 subset C (two digits per symbol):
    start C + 3 symbols + checksum + stop = 68 modules x 5 px = 340 px wide, 50 px high.
The symbol table below is the one in JsBarcode's source (checked against the ISO/IEC 15417 table).

  * encode_code128c()  -> the 0/1 module string + the symbol values
  * barcode_png()      -> a PNG (1-bit, pure python, no Pillow needed), 340 x 50 px at scale 1
  * parse_people()     -> name + identifier rows from an uploaded file or pasted text
  * build_xlsx()       -> XLSX with the barcode PNG INSIDE each cell (what Canva Bulk Create accepts)
  * build_zip()        -> one <identifier>.png per person + names.csv
"""

import csv
import io
import re
import struct
import zipfile
import zlib

MODULE_PX = 5          # jsbarcode-width="5"
BAR_HEIGHT_PX = 50     # jsbarcode-height="50"
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


def encode_code128c(digits):
    """'356452' -> ([105, 35, 64, 52, 12, 106], '1101001110010...')  (Code 128, subset C)"""
    digits = str(digits)
    if not re.fullmatch(r"\d+", digits) or len(digits) % 2:
        raise ValueError(f"'{digits}' needs an even number of digits for Code 128 subset C")
    values = [int(digits[i:i + 2]) for i in range(0, len(digits), 2)]
    checksum = (START_C + sum(v * (i + 1) for i, v in enumerate(values))) % 103
    symbols = [START_C] + values + [checksum, STOP]
    return symbols, "".join(BARS[s] for s in symbols)


def _chunk(kind, data):
    body = kind + data
    return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)


def barcode_png(identifier, scale=1, quiet_px=0, bar_height=BAR_HEIGHT_PX, module_px=MODULE_PX):
    """
    PNG of the barcode: black bars on white, 340 x 50 px for a 6-digit identifier at scale 1.
    scale   : 2 -> 680 x 100 px (same picture, sharper when printed)
    quiet_px: white margin left/right (JsBarcode draws 10 px; 0 keeps the picture at exactly 340 px)
    """
    _, bits = encode_code128c(identifier)
    scale = max(int(scale), 1)
    quiet = max(int(quiet_px), 0) * scale
    px_per_module = module_px * scale
    width = len(bits) * px_per_module + 2 * quiet
    height = bar_height * scale
    row = bytearray(b"\xff" * ((width + 7) // 8))              # 1 = white
    for m, bit in enumerate(bits):
        if bit == "1":
            start = quiet + m * px_per_module
            for x in range(start, start + px_per_module):
                row[x >> 3] &= ~(0x80 >> (x & 7)) & 0xFF         # 0 = black
    raw = b"".join(b"\x00" + bytes(row) for _ in range(height))  # filter byte 0 + packed row, repeated
    ihdr = struct.pack(">IIBBBBB", width, height, 1, 0, 0, 0, 0)  # 1-bit greyscale
    return (b"\x89PNG\r\n\x1a\n" + _chunk(b"IHDR", ihdr) + _chunk(b"IDAT", zlib.compress(raw, 9)) + _chunk(b"IEND", b""))


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
    rows: list of lists (a CSV / XLSX / pasted table). The first row may be a header (name / identifier ...).
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
def build_xlsx(people, scale=1, quiet_px=0):
    """XLSX: name | identifier | barcode (the PNG sits inside the cell, as Excel's "Place in Cell")."""
    import xlsxwriter
    buffer = io.BytesIO()
    workbook = xlsxwriter.Workbook(buffer, {"in_memory": True})
    sheet = workbook.add_worksheet("Barcodes")
    head = workbook.add_format({"bold": True})
    sheet.write_row(0, 0, ["name", "identifier", "barcode"], head)
    sheet.set_column(0, 0, 36)
    sheet.set_column(1, 1, 14)
    sheet.set_column(2, 2, 48)                       # about 340 px
    for r, (name, ident) in enumerate(people, start=1):
        sheet.set_row(r, 38)                         # about 50 px
        sheet.write_string(r, 0, name)
        sheet.write_string(r, 1, ident)               # text, so a leading zero is kept
        png = barcode_png(ident, scale, quiet_px)
        sheet.embed_image(r, 2, f"{ident}.png", {"image_data": io.BytesIO(png), "description": f"Barcode {ident}"})
    workbook.close()
    return buffer.getvalue()


def build_zip(people, scale=1, quiet_px=0):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        listing = io.StringIO()
        writer = csv.writer(listing)
        writer.writerow(["name", "identifier", "file"])
        for name, ident in people:
            archive.writestr(f"{ident}.png", barcode_png(ident, scale, quiet_px))
            writer.writerow([name, ident, f"{ident}.png"])
        archive.writestr("names.csv", listing.getvalue())
    return buffer.getvalue()
