import unittest

from egomail.ui import matches_active_type_filters


class UiTypeFilterTests(unittest.TestCase):
    def test_no_filter_shows_everything(self):
        self.assertTrue(matches_active_type_filters(set(), set()))
        self.assertTrue(matches_active_type_filters({"airbnb"}, set()))

    def test_multiple_filters_use_or_between_mail_types(self):
        active = {"airbnb_promemoria", "airbnb_pagamento"}
        self.assertTrue(matches_active_type_filters({"airbnb_promemoria"}, active))
        self.assertTrue(matches_active_type_filters({"airbnb_pagamento"}, active))
        self.assertTrue(matches_active_type_filters({"airbnb_promemoria", "altro"}, active))
        self.assertFalse(matches_active_type_filters({"vrbo"}, active))
        self.assertFalse(matches_active_type_filters(set(), active))


if __name__ == "__main__":
    unittest.main()
