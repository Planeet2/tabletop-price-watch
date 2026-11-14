import unittest
from datetime import date, datetime, timezone
from decimal import Decimal
from unittest.mock import Mock

from tabletop_price_watch.rates import cbr_rate, collect_rates, ecb_rate

ECB_XML = b"""
<Cube xmlns="http://www.ecb.int/vocabulary/2002-08-01/eurofxref">
  <Cube time="2026-11-13">
    <Cube currency="USD" rate="1.1052"/>
  </Cube>
</Cube>"""
CBR_XML = """<?xml version="1.0" encoding="windows-1251"?>
<ValCurs Date="14.11.2026">
  <Valute><CharCode>USD</CharCode><Nominal>1</Nominal><Name>Доллар США</Name><Value>82,1516</Value></Valute>
</ValCurs>""".encode("cp1251")


class RateTests(unittest.TestCase):
    def test_ecb_rate_is_dollars_per_euro(self):
        rate = ecb_rate(ECB_XML, "ECB")
        self.assertEqual(rate.effective_date, date(2026, 11, 13))
        self.assertEqual(rate.to_usd(Decimal("44.95")), Decimal("49.68"))

    def test_cbr_rate_inverts_rubles_per_dollar(self):
        rate = cbr_rate(CBR_XML, "CBR")
        self.assertEqual(rate.effective_date, date(2026, 11, 14))
        self.assertEqual(rate.to_usd(Decimal("5290")), Decimal("64.39"))
        quoted_per_hundred = CBR_XML.replace(b"<Nominal>1", b"<Nominal>100").replace(b"82,1516", b"8215,16")
        self.assertEqual(cbr_rate(quoted_per_hundred, "CBR").usd_per_unit, rate.usd_per_unit)

    def test_stale_or_unavailable_rate_is_absent(self):
        session = Mock()
        session.get.side_effect = [
            Mock(status_code=200, content=ECB_XML.replace(b"2026-11-13", b"2026-10-13")),
            Mock(status_code=503),
        ]
        with self.assertLogs("tabletop_price_watch.rates", level="WARNING"):
            rates = collect_rates(session, datetime(2026, 11, 14, tzinfo=timezone.utc))
        self.assertEqual(set(rates), {"USD"})
