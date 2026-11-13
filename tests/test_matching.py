import unittest
from decimal import Decimal

from tabletop_price_watch.cli import verify_identity
from tabletop_price_watch.retailers import Offer, parse_offer


class MatchingTests(unittest.TestCase):
    def test_rejects_different_edition_even_with_similar_title(self):
        offer = Offer("Azul Duel", Decimal("20.00"), "In stock", "826956621607", "UPC")
        with self.assertRaisesRegex(ValueError, "expected UPC"):
            verify_identity({"upc": "826956601609"}, offer)

    def test_accepts_distributor_prefix_for_same_part_number(self):
        offer = Offer("Azul Mini", Decimal("20.00"), "In stock", "ASMNMG60140EN", "MPN")
        self.assertEqual(verify_identity({"mpn": "NMG60140EN"}, offer), "MPN")

    def test_noble_knight_uses_new_sale_price_and_excludes_used_copy(self):
        html = b"""
        <h1>Test game</h1>
        <div class="info-line"><div class="label">MFG. Part #</div><div class="value">ABC123</div></div>
        <script type="application/ld+json">
          {"@type":"Product","offers":{"price":"70.00","priceCurrency":"USD"}}
        </script>
        <div class="condition-col in_stock">
          <span class="condition-value"><a>SW (MINT/New)</a></span>
          <input class="condition-quantity" aria-label="Edit quantity Condition SW (MINT/New) Price $62.00">
        </div>
        <div class="condition-col in_stock">
          <span class="condition-value"><a>VG/Used</a></span>
          <input class="condition-quantity" aria-label="Edit quantity Condition VG/Used Price $40.00">
        </div>"""
        offer = parse_offer("noble_knight", html)
        self.assertEqual(offer.price, Decimal("62.00"))
        self.assertEqual(offer.identity, "ABC123")
        self.assertEqual(offer.status, "In stock")
