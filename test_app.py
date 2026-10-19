import csv
import io
import unittest
import zipfile
from xml.etree import ElementTree as ET

from app import decode_file, make_csv, make_xlsx, parse_text


class FormatterTests(unittest.TestCase):
    def test_inline_and_multiline_records(self):
        result = parse_text('name : Angel Reyes email : AAA@gmail.com\ntelefono: 00123\n\nname: Maria\nemail: maria@gmail.com')
        self.assertEqual(result['columns'], ['Nombre', 'Email', 'Teléfono'])
        self.assertEqual(result['rows'], [
            {'Nombre': 'Angel Reyes', 'Email': 'AAA@gmail.com', 'Teléfono': '00123'},
            {'Nombre': 'Maria', 'Email': 'maria@gmail.com'},
        ])

    def test_repeated_keys_split_records(self):
        result = parse_text('name: Uno\nemail: uno@example.com\nname: Dos\nemail: dos@example.com')
        self.assertEqual(len(result['rows']), 2)
        self.assertEqual(result['rows'][1]['Nombre'], 'Dos')
        self.assertEqual(len(result['warnings']), 1)

    def test_aliases_new_fields_and_urls(self):
        result = parse_text('Nombre completo: Ángel Reyes correo electrónico: a@example.com\nSitio web: https://example.com\nCódigo postal=0012')
        self.assertEqual(result['rows'][0], {
            'Nombre': 'Ángel Reyes', 'Email': 'a@example.com',
            'Sitio web': 'https://example.com', 'Código postal': '0012',
        })

    def test_unstructured_text_is_preserved(self):
        result = parse_text('Texto suelto\n\nnota: primera línea\nsegunda línea')
        self.assertEqual(result['rows'][0]['Texto sin etiqueta'], 'Texto suelto')
        self.assertEqual(result['rows'][1]['nota'], 'primera línea\nsegunda línea')
        self.assertEqual(len(result['warnings']), 2)

    def test_encodings_and_empty(self):
        for encoding in ['utf-8-sig', 'utf-16', 'cp1252']:
            self.assertEqual(decode_file('nombre: Ángel'.encode(encoding))[0], 'nombre: Ángel')
        with self.assertRaises(ValueError):
            parse_text(' \n\n')

    def test_export_only_selected_columns_and_literal_values(self):
        output = make_xlsx(['Nombre', 'Teléfono'], [
            {'Nombre': '=HYPERLINK("example.com")', 'Email': 'excluido@example.com', 'Teléfono': '00123'},
            {'Nombre': 'Ángel & Reyes'},
        ])
        ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        with zipfile.ZipFile(io.BytesIO(output)) as archive:
            for name in archive.namelist():
                ET.fromstring(archive.read(name))
            sheet = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
            rows = sheet.findall('s:sheetData/s:row', ns)
            values = [[cell.find('s:is/s:t', ns).text or '' for cell in row] for row in rows]
            self.assertEqual(values, [['Nombre', 'Teléfono'], ['=HYPERLINK("example.com")', '00123'], ['Ángel & Reyes', '']])
            self.assertEqual(sheet.findall('.//s:f', ns), [])
            self.assertTrue(all(c.attrib['t'] == 'inlineStr' for r in rows for c in r))

    def test_export_rejects_no_columns_and_oversized_cells(self):
        with self.assertRaises(ValueError):
            make_xlsx([], [])
        with self.assertRaises(ValueError):
            make_xlsx(['nota'], [{'nota': 'x' * 32768}])

    def test_csv_preserves_special_values_and_column_selection(self):
        columns = ['nota', 'Nombre', 'Teléfono']
        rows = [
            {'Nombre': 'Ángel, "Reyes"', 'Email': 'excluido@example.com',
             'Teléfono': '00123', 'nota': 'primera línea\nsegunda línea'},
            {'Nombre': '=texto literal', 'Teléfono': '+59170000000'},
        ]
        output = make_csv(columns, rows)
        self.assertTrue(output.startswith(b'\xef\xbb\xbf'))
        values = list(csv.reader(io.StringIO(output.decode('utf-8-sig'), newline='')))
        self.assertEqual(values, [columns,
            ['primera línea\nsegunda línea', 'Ángel, "Reyes"', '00123'],
            ['', '=texto literal', '+59170000000']])
        self.assertNotIn(b'excluido@example.com', output)

    def test_csv_rejects_no_columns(self):
        with self.assertRaises(ValueError):
            make_csv([], [])

    def test_text_values_survive_parsing_and_both_exports(self):
        originals = ['00123', '12345678901234567890', '1E10', '01/02',
                     '+59170000000', '1.234,56', '=1+2', '|dato;', ';dato|',
                     'Ángel, "Reyes"', '日本語']
        parsed = parse_text('\n\n'.join(f'valor: {value}' for value in originals))
        self.assertEqual([row['valor'] for row in parsed['rows']], originals)
        output = make_csv(['valor'], parsed['rows'])
        csv_rows = list(csv.reader(io.StringIO(output.decode('utf-8-sig'), newline='')))
        self.assertEqual(csv_rows, [['valor']] + [[value] for value in originals])
        ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        with zipfile.ZipFile(io.BytesIO(make_xlsx(['valor'], parsed['rows']))) as archive:
            sheet = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
            cells = sheet.findall('s:sheetData/s:row/s:c', ns)
            self.assertEqual([cell.find('s:is/s:t', ns).text for cell in cells], ['valor'] + originals)
            self.assertTrue(all(cell.attrib['t'] == 'inlineStr' for cell in cells))


if __name__ == '__main__':
    unittest.main()
