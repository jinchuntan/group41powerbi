"""Check the financial definitions, source handling and image export contract."""

from decimal import Decimal
from pathlib import Path
import tempfile
import unittest
import xml.etree.ElementTree as ET

import matplotlib.pyplot as plt
from PIL import Image

from scripts.build_recommendation_figures import (
    calculate_financial_metrics,
    calculate_marketing_metrics,
    calculate_service_metrics,
    deduplicate_records,
    parse_discount,
    save_chart,
)


class CalculationTests(unittest.TestCase):
    def test_identical_duplicates_are_removed_but_conflicts_are_rejected(self):
        row = {"Order ID": "O1", "Quantity": 1}
        records, duplicate_count = deduplicate_records([row, dict(row)], "Order ID")
        self.assertEqual(len(records), 1)
        self.assertEqual(duplicate_count, 1)

        conflicting_row = {"Order ID": "O1", "Quantity": 2}
        with self.assertRaisesRegex(ValueError, "Conflicting records"):
            deduplicate_records([row, conflicting_row], "Order ID")

    def test_text_and_numeric_discounts_have_the_same_meaning(self):
        self.assertEqual(parse_discount("10%"), Decimal("0.10"))
        self.assertEqual(parse_discount(0.1), Decimal("0.10"))
        with self.assertRaises(ValueError):
            parse_discount("110%")

    def test_roas_uses_total_revenue_divided_by_total_spend(self):
        campaigns = {
            "M1": {"Campaign Type": "Email Retargeting", "Campaign Spend (RM)": 100,
                   "Revenue Generated (RM)": 500},
            "M2": {"Campaign Type": "Email Retargeting", "Campaign Spend (RM)": 900,
                   "Revenue Generated (RM)": 900},
        }
        rows, summary = calculate_marketing_metrics(campaigns)
        self.assertAlmostEqual(rows[0]["roas"], 1.4)
        self.assertEqual(summary["total_spend_rm"], 1000)

    def test_margin_is_weighted_and_missing_cost_orders_are_excluded(self):
        orders = {
            "O1": {"Product ID": "P1", "Order Status": "Completed", "Quantity": 1,
                   "Unit Price (RM)": 1000, "Discount": 0},
            "O2": {"Product ID": "P2", "Order Status": "Completed", "Quantity": 1,
                   "Unit Price (RM)": 100, "Discount": 0},
            "O3": {"Product ID": "P999", "Order Status": "Completed", "Quantity": 1,
                   "Unit Price (RM)": 500, "Discount": 0},
        }
        products = {
            "P1": {"Category": "Home Office", "Unit Cost (RM)": 900},
            "P2": {"Category": "Home Office", "Unit Cost (RM)": 50},
        }
        rows, summary = calculate_financial_metrics(orders, products)
        self.assertAlmostEqual(rows[0]["weighted_margin_pct"], 100 * 150 / 1100)
        self.assertEqual(summary["revenue_rm"], 1100)
        self.assertEqual(summary["gross_profit_rm"], 150)
        self.assertEqual(summary["excluded_unknown_cost_orders"], 1)
        self.assertEqual(summary["excluded_unknown_cost_revenue_rm"], 500)

    def test_service_counts_keep_orders_with_missing_product_information(self):
        orders = {"O1": {"Product ID": "P999"}, "O2": {"Product ID": "P1"}}
        experience = {
            "O1": {"Complaint Category": "late delivery", "Returned Order": "No"},
            "O2": {"Complaint Category": "Product Not Suitable", "Returned Order": "Yes"},
        }
        rows, summary = calculate_service_metrics(orders, experience)
        self.assertEqual(summary["unique_orders"], 2)
        self.assertEqual(summary["total_complaints"], 2)
        self.assertEqual(summary["delivery_suitability_returns"], 1)
        self.assertEqual({row["category"] for row in rows},
                         {"Late Delivery", "Product Not Suitable"})


class ImageExportTests(unittest.TestCase):
    def test_exports_omit_optional_metadata_and_preserve_pixels(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory)
            figure, axes = plt.subplots(figsize=(2, 1))
            axes.barh(["Example"], [1])
            save_chart(figure, output, "example", dpi=100)

            with Image.open(output / "example.png") as image:
                self.assertEqual(image.size, (200, 100))
                self.assertEqual(set(image.info), {"dpi"})
                self.assertFalse(image.getexif())

            svg = ET.parse(output / "example.svg").getroot()
            self.assertIsNone(svg.find("{http://www.w3.org/2000/svg}metadata"))


if __name__ == "__main__":
    unittest.main()
