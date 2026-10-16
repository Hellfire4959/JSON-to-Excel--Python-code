import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
import pandas as pd

spec = importlib.util.spec_from_file_location('converter', Path(__file__).parents[1] / 'scripts/json_to_excel.py')
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)

class ConverterTests(unittest.TestCase):
    def test_combines_objects_arrays_and_utf8_bom(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / 'a.json').write_text(json.dumps([{'name':'Alex'},{'name':'Sam'}]), encoding='utf-8-sig')
            (root / 'b.json').write_text(json.dumps({'name':'Lee','profile':{'city':'Tokyo'}}))
            output = root / 'output/report.xlsx'
            self.assertEqual(module.convert_folder(root, output), 3)
            data = pd.read_excel(output)
            self.assertEqual(data['name'].tolist(), ['Alex','Sam','Lee'])
            self.assertEqual(data.loc[2, 'profile.city'], 'Tokyo')

    def test_rejects_scalar_input_without_writing_output(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory); (root / 'a.json').write_text('[1,2]')
            with self.assertRaises(ValueError): module.convert_folder(root, root / 'result.xlsx')
            self.assertFalse((root / 'result.xlsx').exists())

    def test_rejects_empty_input_directory(self):
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaisesRegex(ValueError, 'No .json'):
                module.convert_folder(directory, Path(directory) / 'result.xlsx')

    def test_invalid_later_file_does_not_write_partial_workbook(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / 'a.json').write_text('[{"name":"Alex"}]')
            (root / 'b.json').write_text('{broken')
            with self.assertRaises(ValueError): module.convert_folder(root, root / 'result.xlsx')
            self.assertFalse((root / 'result.xlsx').exists())

if __name__ == '__main__': unittest.main()
