import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from pointage import amount, norm, score_row


class MatchingTests(unittest.TestCase):
    def test_amount_parses_french_decimal(self):
        self.assertEqual(amount("1 234,50"), amount("1234.50"))

    def test_norm_removes_separators_and_accents(self):
        self.assertEqual(norm("Factu/2026/06/0028"), "FACTU2026060028")
        self.assertEqual(norm("Société SAMM"), "SOCIETESAMM")

    def test_generic_partner_does_not_add_partner_points(self):
        text = "FACTU-2026-01-0014 FA00012591 MALAGASY BEST SPICES TOTAL 12591"
        doc = {
            "text": text,
            "norm": norm(text),
            "tokens": {norm(tok) for tok in text.split() if len(norm(tok)) >= 3},
        }
        fields = {
            "numero": None,
            "reference": None,
            "reference_interne": None,
            "date": None,
            "ht": None,
            "tax": None,
            "ttc": None,
            "partner": "FOURNISSEUR DIVERS",
        }
        score, reasons = score_row(fields, doc)
        self.assertEqual(score, 0)
        self.assertNotIn("partenaire", reasons)
        self.assertNotIn("partenaire_approximatif", reasons)

    def test_specific_partner_requires_discriminating_match(self):
        text = "FACTURE MALAGASY BEST SPICES TOTAL 12591"
        doc = {
            "text": text,
            "norm": norm(text),
            "tokens": {norm(tok) for tok in text.split() if len(norm(tok)) >= 3},
        }
        fields = {
            "numero": None,
            "reference": None,
            "reference_interne": None,
            "date": None,
            "ht": None,
            "tax": None,
            "ttc": None,
            "partner": "MALAGASY BEST SPICES",
        }
        score, reasons = score_row(fields, doc)
        self.assertEqual(score, 14)
        self.assertIn("partenaire", reasons)


if __name__ == "__main__":
    unittest.main()
