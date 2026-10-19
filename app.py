"""Formateador TXT local. Solo requiere Python 3.10 o superior."""

import csv
import io
import json
import re
import secrets
import threading
import unicodedata
import webbrowser
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from xml.etree.ElementTree import Element, SubElement, tostring

ROOT = Path(__file__).resolve().parent
MAX_BYTES = 20 * 1024 * 1024
TOKEN = secrets.token_urlsafe(32)


def normalize(value):
    value = unicodedata.normalize('NFKD', value.strip().lower())
    return ''.join(c for c in value if not unicodedata.combining(c))


ALIASES = {
    'name': 'Nombre', 'nombre': 'Nombre', 'full name': 'Nombre',
    'nombre completo': 'Nombre', 'email': 'Email', 'e-mail': 'Email',
    'correo': 'Email', 'correo electronico': 'Email',
    'phone': 'Teléfono', 'telefono': 'Teléfono', 'tel': 'Teléfono',
    'celular': 'Teléfono', 'mobile': 'Teléfono',
}
MULTI = r'nombre completo|full name|correo electr[oó]nico'
INLINE = re.compile(
    rf'(?<!\S)(?P<key>{MULTI}|[^\W\d][\w.\-]*)[ \t]*[:=](?!//)', re.I
)
START = re.compile(r'^\s*(?P<key>[^\W\d][\w.\- \t]{0,59}?)\s*[:=](?!//)', re.I)


def parse_text(text):
    """Parsea etiquetas y valores; conserva texto dudoso y reporta su posición."""
    columns, rows, warnings = [], [], []
    canonical = {}
    current = {}
    last_key = None

    def finish():
        nonlocal current, last_key
        if current:
            rows.append(current)
        current, last_key = {}, None

    for number, line in enumerate(text.splitlines(), 1):
        if not line.strip() or re.fullmatch(r'\s*[-_*]{3,}\s*', line):
            finish()
            continue
        matches = list(INLINE.finditer(line))
        initial = START.match(line)
        if initial:
            matches = [initial] + [m for m in matches if m.start() >= initial.end()]
        if not matches:
            if last_key is not None:
                current[last_key] += '\n' + line.strip()
                warnings.append(f'Línea {number}: texto sin etiqueta añadido a «{last_key}». Revísalo.')
            else:
                key = canonical.setdefault('texto sin etiqueta', 'Texto sin etiqueta')
                if key not in columns:
                    columns.append(key)
                current[key] = current.get(key, '') + ('\n' if key in current else '') + line.strip()
                last_key = key
                warnings.append(f'Línea {number}: no se detectó una etiqueta; se conservó como texto.')
            continue
        if line[:matches[0].start()].strip():
            warnings.append(f'Línea {number}: hay texto antes de la primera etiqueta; se conservó.')
            key = canonical.setdefault('texto sin etiqueta', 'Texto sin etiqueta')
            if key not in columns:
                columns.append(key)
            current[key] = current.get(key, '') + line[:matches[0].start()].strip()
        for i, match in enumerate(matches):
            raw = match.group('key').strip()
            normalized = normalize(raw)
            key = ALIASES.get(normalized) or canonical.setdefault(normalized, raw)
            if key not in columns:
                columns.append(key)
            if key in current:
                finish()
                warnings.append(f'Línea {number}: «{key}» se repite; se inició un nuevo registro.')
            end = matches[i + 1].start() if i + 1 < len(matches) else len(line)
            current[key] = line[match.end():end].strip()
            last_key = key
    finish()
    if not rows:
        raise ValueError('El archivo está vacío o no contiene datos.')
    if len(columns) > 16384 or len(rows) > 1048575:
        raise ValueError('El archivo excede el número de filas o columnas permitido por Excel.')
    return {'columns': columns, 'rows': rows, 'warnings': warnings,
            'note': 'Cada línea en blanco separa registros. Una etiqueta repetida inicia otro registro. '
                    'En una misma línea, usa etiquetas de una palabra o Nombre completo, Full name y Correo electrónico.'}


def decode_file(data):
    if data.startswith((b'\xff\xfe', b'\xfe\xff')):
        return data.decode('utf-16'), 'UTF-16'
    try:
        return data.decode('utf-8-sig'), 'UTF-8'
    except UnicodeDecodeError:
        return data.decode('cp1252'), 'Windows-1252'


def make_csv(columns, rows):
    if not columns:
        raise ValueError('Selecciona al menos una columna.')
    output = io.StringIO(newline='')
    writer = csv.writer(output, lineterminator='\r\n')
    writer.writerow(columns)
    writer.writerows([row.get(column, '') for column in columns] for row in rows)
    # BOM UTF-8 para reconocer las tildes también al abrirlo en Excel.
    # CSV conserva valores literales: importar como texto evita conversiones.
    return output.getvalue().encode('utf-8-sig')


def column_letter(index):
    result = ''
    while index:
        index, rem = divmod(index - 1, 26)
        result = chr(65 + rem) + result
    return result


def make_xlsx(columns, rows):
    if not columns:
        raise ValueError('Selecciona al menos una columna.')
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    sheet = Element('worksheet', xmlns=ns)
    views = SubElement(sheet, 'sheetViews')
    view = SubElement(views, 'sheetView', workbookViewId='0')
    SubElement(view, 'pane', ySplit='1', topLeftCell='A2', activePane='bottomLeft', state='frozen')
    widths = SubElement(sheet, 'cols')
    for i in range(1, len(columns) + 1):
        SubElement(widths, 'col', min=str(i), max=str(i), width='28', customWidth='1')
    data = SubElement(sheet, 'sheetData')
    for ri, values in enumerate([columns] + [[r.get(c, '') for c in columns] for r in rows], 1):
        row = SubElement(data, 'row', r=str(ri))
        for ci, value in enumerate(values, 1):
            value = str(value)
            if len(value) > 32767:
                raise ValueError(f'La celda {column_letter(ci)}{ri} supera los 32.767 caracteres de Excel.')
            # Se almacenan como texto, incluso si comienzan con =, +, - o @.
            cell = SubElement(row, 'c', r=f'{column_letter(ci)}{ri}', t='inlineStr', s='1' if ri == 1 else '0')
            content = SubElement(SubElement(cell, 'is'), 't', {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'})
            content.text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]', '', value)
    SubElement(sheet, 'autoFilter', ref=f'A1:{column_letter(len(columns))}{len(rows) + 1}')
    parts = {
        '[Content_Types].xml': '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
        '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': f'<workbook xmlns="{ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Datos" sheetId="1" r:id="rId1"/></sheets></workbook>',
        'xl/_rels/workbook.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
        'xl/styles.xml': f'<styleSheet xmlns="{ns}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF145E4D"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="49" fontId="1" fillId="2" borderId="0" xfId="0" applyFill="1" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>',
        'xl/worksheets/sheet1.xml': tostring(sheet, encoding='utf-8', xml_declaration=True),
    }
    output = io.BytesIO()
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, value in parts.items():
            archive.writestr(name, value)
    return output.getvalue()


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def respond(self, code, body, content_type='application/json; charset=utf-8'):
        if isinstance(body, dict):
            body = json.dumps(body, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path != '/':
            return self.respond(404, {'error': 'No encontrado.'})
        html = (ROOT / 'index.html').read_text(encoding='utf-8').replace('__TOKEN__', TOKEN)
        self.respond(200, html.encode('utf-8'), 'text/html; charset=utf-8')

    def do_POST(self):
        if self.headers.get('X-Local-Token') != TOKEN:
            return self.respond(403, {'error': 'Recarga la aplicación para continuar.'})
        try:
            length = int(self.headers.get('Content-Length', '0'))
            if not 0 < length <= MAX_BYTES:
                raise ValueError('El tamaño máximo por solicitud es 20 MB.')
            body = self.rfile.read(length)
            if self.path == '/api/parse':
                text, encoding = decode_file(body)
                if '\x00' in text:
                    raise ValueError('No parece un archivo de texto compatible. Usa UTF-8, UTF-16 con BOM o Windows-1252.')
                result = parse_text(text)
                result['encoding'] = encoding
                self.respond(200, result)
            elif self.path == '/api/export':
                payload = json.loads(body)
                if not isinstance(payload, dict):
                    raise ValueError('Datos de exportación inválidos.')
                export_format = payload.get('format', 'xlsx')
                if export_format not in ('csv', 'xlsx'):
                    raise ValueError('Elige CSV o Excel para exportar.')
                columns, rows = payload.get('columns'), payload.get('rows')
                if (not isinstance(columns, list) or not columns or len(columns) > 16384
                        or not all(isinstance(c, str) for c in columns)
                        or not isinstance(rows, list) or len(rows) > 1048575
                        or not all(isinstance(r, dict) for r in rows)):
                    raise ValueError('Datos de exportación inválidos.')
                if export_format == 'csv':
                    self.respond(200, make_csv(columns, rows), 'text/csv; charset=utf-8')
                else:
                    self.respond(200, make_xlsx(columns, rows),
                                 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            else:
                self.respond(404, {'error': 'No encontrado.'})
        except (ValueError, UnicodeError, TypeError) as error:
            self.respond(400, {'error': str(error)})


if __name__ == '__main__':
    server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
    address = f'http://127.0.0.1:{server.server_port}'
    print(f'Formateador TXT: {address}\nCierra esta ventana o pulsa Ctrl+C para detenerlo.')
    threading.Timer(0.5, webbrowser.open, args=(address,)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
