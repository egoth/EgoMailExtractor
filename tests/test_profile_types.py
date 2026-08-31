import unittest

from egomail.profile import normalize_profile


class ProfileTypeTests(unittest.TestCase):
    def test_old_field_types_migrate_to_three_excel_types(self):
        p = {
            'schema_version': 1,
            'fields': [
                {'name': 'a', 'type': 'text'},
                {'name': 'b', 'type': 'date'},
                {'name': 'c', 'type': 'int'},
                {'name': 'd', 'type': 'money'},
            ],
        }
        normalize_profile(p)
        self.assertEqual([f['type'] for f in p['fields']], ['text', 'date', 'number', 'number'])
        self.assertEqual(p['fields'][3]['format'], 'currency_eur')
        self.assertEqual(p['schema_version'], 7)


    def test_old_criteria_migrate_to_match_rules(self):
        p = {
            'schema_version': 2,
            'fields': [],
            'mail_types': [{
                'name': 'x',
                'criteria': {
                    'subject_contains': 'Airbnb',
                    'sender_regex': r'@airbnb\.com',
                    'body_contains': 'Codice di conferma',
                },
                'rules': [],
            }],
        }
        normalize_profile(p)
        mt = p['mail_types'][0]
        self.assertNotIn('criteria', mt)
        self.assertEqual(len(mt['match_rules']), 3)
        self.assertEqual([r['source'] for r in mt['match_rules']], ['sender', 'subject', 'body'])


if __name__ == '__main__':
    unittest.main()
