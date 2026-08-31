import unittest

from egomail.extraction import parse_it_date, transform
from egomail.mail import MailMessage


class MarchDateRegressionTests(unittest.TestCase):
    def test_march_without_weekday_is_not_stripped(self):
        self.assertEqual(parse_it_date("25 mar", 2026), "2026-03-25")

    def test_tuesday_in_march_keeps_month_mar(self):
        self.assertEqual(parse_it_date("mar 25 mar", 2026), "2026-03-25")

    def test_other_weekday_before_march(self):
        self.assertEqual(parse_it_date("mer 25 mar", 2026), "2026-03-25")

    def test_transform_uses_mail_year_for_march(self):
        mail = MailMessage(
            uid=2717,
            subject="Prenotazione confermata - Chen Hsuen Liang arriverà il 25 feb",
            sender="Airbnb <automated@airbnb.com>",
            date="Wed, 25 Feb 2026 11:19:34 +0000",
            body="",
            message_id="<test@example>",
        )
        self.assertEqual(transform("25 mar", "date_it_email_year", mail), "2026-03-25")


if __name__ == "__main__":
    unittest.main()
