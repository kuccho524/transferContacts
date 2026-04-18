import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / 'csvContactsFirst'))

from convert_contacts import generate_contact_csv, get_ordinal_suffix


class ConvertContactsOrdinalTests(unittest.TestCase):
    def test_ordinal_suffixes_use_standard_spelling(self):
        self.assertEqual(get_ordinal_suffix(3), 'Third')
        self.assertEqual(get_ordinal_suffix(4), 'Fourth')
        self.assertEqual(get_ordinal_suffix(5), 'Fifth')

    def test_generated_headers_follow_standard_spelling(self):
        records = [{
            'First Name': 'Taro',
            'Last Name': 'Suzuki',
            'Phonetic First Name': 'タロウ',
            'Phonetic Last Name': 'スズキ',
            'Organization Name': 'Example Corp',
            'Organization Department': 'Support',
            'E-mail 1 - Label': 'work',
            'E-mail 1 - Value': 'primary@example.com',
            'E-mail 2 - Label': 'home',
            'E-mail 2 - Value': 'secondary@example.com',
            'E-mail 3 - Label': 'other',
            'E-mail 3 - Value': 'third@example.com',
            'Phone 1 - Label': 'mobile',
            'Phone 1 - Value': '090-0000-0001',
            'Phone 2 - Label': 'work',
            'Phone 2 - Value': '03-0000-0002',
            'Phone 3 - Label': 'home',
            'Phone 3 - Value': '03-0000-0003',
            'Phone 4 - Label': 'other',
            'Phone 4 - Value': '03-0000-0004',
        }]

        output_path = PROJECT_ROOT / 'tests' / 'tmp_convert_contacts.csv'
        try:
            generate_contact_csv(records, output_path)
            header = output_path.read_text(encoding='utf-8').splitlines()[0]
        finally:
            if output_path.exists():
                output_path.unlink()

        self.assertIn('ThirdEmailAddress', header)
        self.assertIn('FourthPhoneNumber', header)
        self.assertNotIn('ThirdryEmailAddress', header)
        self.assertNotIn('ForthryPhoneNumber', header)


if __name__ == '__main__':
    unittest.main()

