"""Prueba de carga y descarga por HTTP local, sin servicios externos."""
import csv
import io
import json
import threading
import urllib.error
import urllib.request
import zipfile
from xml.etree import ElementTree as ET

from app import Handler, TOKEN, ThreadingHTTPServer


server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
threading.Thread(target=server.serve_forever, daemon=True).start()
base = f'http://127.0.0.1:{server.server_port}'
opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def post(path, data, token=TOKEN):
    request = urllib.request.Request(base + path, data, {'X-Local-Token': token}, method='POST')
    return opener.open(request, timeout=10).read()


try:
    html = opener.open(base, timeout=10).read().decode('utf-8')
    assert TOKEN in html and '__TOKEN__' not in html
    parsed = json.loads(post('/api/parse', 'name: Angel Reyes email: AAA@gmail.com\ntelefono: 00123'.encode('utf-8')))
    assert parsed['rows'][0]['Email'] == 'AAA@gmail.com'
    csv_content = post('/api/export', json.dumps({'format': 'csv', 'columns': ['Teléfono', 'Nombre'], 'rows': parsed['rows']}).encode('utf-8'))
    csv_values = list(csv.reader(io.StringIO(csv_content.decode('utf-8-sig'), newline='')))
    assert csv_values == [['Teléfono', 'Nombre'], ['00123', 'Angel Reyes']]
    for payload in [[], {'format': 'pdf'}, {'format': 'csv', 'columns': [], 'rows': []}]:
        try:
            post('/api/export', json.dumps(payload).encode('utf-8'))
        except urllib.error.HTTPError as error:
            assert error.code == 400
        else:
            raise AssertionError('Se esperaba el rechazo de una exportación inválida.')
    content = post('/api/export', json.dumps({'columns': ['Nombre', 'Teléfono'], 'rows': parsed['rows']}).encode('utf-8'))
    with zipfile.ZipFile(io.BytesIO(content)) as archive:
        xml = archive.read('xl/worksheets/sheet1.xml').decode('utf-8')
        assert 'Email' not in xml and 'AAA@gmail.com' not in xml
        assert 'Angel Reyes' in xml and '00123' in xml
        for name in archive.namelist():
            ET.fromstring(archive.read(name))
    try:
        post('/api/parse', b'name: Test', token='invalid')
    except urllib.error.HTTPError as error:
        assert error.code == 403
    else:
        raise AssertionError('Se esperaba el rechazo de la solicitud sin token válido.')
    try:
        import openpyxl
    except ImportError:
        print('XLSX validado mediante su estructura XML; openpyxl no está instalado.')
    else:
        workbook = openpyxl.load_workbook(io.BytesIO(content))
        assert list(workbook.active.values) == [('Nombre', 'Teléfono'), ('Angel Reyes', '00123')]
        print('XLSX leído correctamente con openpyxl.')
    print('HTTP local: carga, detección, selección de columnas, CSV, Excel y control de acceso correctos.')
finally:
    server.shutdown()
    server.server_close()
