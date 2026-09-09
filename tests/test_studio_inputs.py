import unittest

import pandas as pd

from studio_inputs import filter_exceptions, parse_domains


class WorkspaceInputTests(unittest.TestCase):
    def test_pasted_domains_are_normalised_without_silent_bad_rules(self):
        domains, invalid = parse_domains('@Example.ORG; example.org, second.org\nhttps://wrong.org\nuser@wrong.org\nfoo bar\n-bad.org\n@')
        self.assertEqual(domains, ['example.org', 'second.org'])
        self.assertEqual(len(invalid), 5)
        self.assertEqual(parse_domains('')[0], [])

    def test_issue_scope_and_literal_search_preserve_records(self):
        records = pd.DataFrame([
            {'Participant Key': 'ID:1', 'Missing Critical Fields': 'Full Name', 'Exact Repeat Review': '', 'Assignment Title': 'AML [2026]'},
            {'Participant Key': 'UNRESOLVED:2', 'Missing Critical Fields': 'User ID', 'Exact Repeat Review': '', 'Assignment Title': 'AML'},
            {'Participant Key': 'ID:3', 'Missing Critical Fields': '', 'Exact Repeat Review': 'Review', 'Assignment Title': 'AML'},
        ])
        self.assertEqual(len(filter_exceptions(records, '[2026]', 'Missing fields')), 1)
        self.assertEqual(len(filter_exceptions(records, '', 'Unresolved identity')), 1)
        self.assertEqual(len(filter_exceptions(records, '', 'Possible repeats')), 1)
        self.assertEqual(len(filter_exceptions(records, '[2026]', 'Possible repeats')), 0)
        self.assertEqual(len(records), 3)
