import contextlib
import io
import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / 'csvContactsSecond'))

from csv_parser import parse_csv
from data_processor import (
    check_consistency,
    create_contact_map,
    create_label_map,
    transform_to_output_data,
)


class PublicSampleFlowTests(unittest.TestCase):
    def setUp(self):
        self.samples_dir = PROJECT_ROOT / 'samples'

    def test_sample_files_produce_label_output(self):
        export_data = parse_csv(str(self.samples_dir / 'export_contacts.csv'))
        registered_data = parse_csv(str(self.samples_dir / 'registered_contacts.csv'))
        contacts_data = parse_csv(str(self.samples_dir / 'contacts.csv'))
        label_data = parse_csv(str(self.samples_dir / 'contactgroups.csv'))

        with contextlib.redirect_stdout(io.StringIO()):
            consistency = check_consistency(export_data, registered_data)
        self.assertEqual(consistency['skip_list'], [])

        with contextlib.redirect_stdout(io.StringIO()):
            result = transform_to_output_data(
                export_data,
                registered_data,
                create_contact_map(contacts_data),
                create_label_map(label_data),
                consistency['skip_row_indexes'],
                'target@example.com',
            )

        self.assertEqual(result['max_label_count'], 3)
        self.assertEqual(sorted(result['grouped_data'].keys()), [2, 3])

        two_label_rows = result['grouped_data'][2]
        three_label_rows = result['grouped_data'][3]

        self.assertEqual(two_label_rows[0]['ContactID'], 'people/c002')
        self.assertEqual(two_label_rows[0]['TargetEmail'], 'target@example.com')
        self.assertEqual(three_label_rows[0]['ContactID'], 'people/c001')
        self.assertEqual(three_label_rows[0]['ThirdLabel'], 'contactGroups/vendor')


if __name__ == '__main__':
    unittest.main()
