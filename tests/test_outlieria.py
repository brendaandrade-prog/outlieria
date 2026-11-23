import unittest

import numpy as np
import pandas as pd

import streamlit_app as app


class OutlieriaSpatialTests(unittest.TestCase):
    def test_radius_validation(self):
        self.assertEqual(app.validate_radius(125), 125.0)
        with self.assertRaises(ValueError):
            app.validate_radius(0)
        with self.assertRaises(ValueError):
            app.validate_radius(-10)

    @unittest.skipUnless(app.LOF_AVAILABLE, "scikit-learn não está instalado neste ambiente")
    def test_lof_keeps_depths_separate_and_uses_at_least_five(self):
        rows = []
        # Sete históricos próximos permitem k=5 sem usar 2026 no treino.
        for i in range(7):
            rows.append({
                "year": "2023", "crop": "L01", "depth": "0-10 cm",
                "longitude": -50.0 + i * 0.00003, "latitude": -20.0,
                "PH": 5.0 + i * 0.01,
            })
        rows.append({
            "year": "2026", "crop": "L01", "depth": "0-10 cm",
            "longitude": -50.0, "latitude": -20.0, "PH": 6.0,
        })
        rows.append({
            "year": "2026", "crop": "L01", "depth": "10-20 cm",
            "longitude": -50.0, "latitude": -20.0, "PH": 6.0,
        })
        sample = pd.DataFrame(rows)
        result = app.run_spatial_lof(sample, ["PH"], 300)
        self.assertFalse(result.empty)
        scored = result[result["ml_status"].isin(["normal", "anomalia espacial"])]
        self.assertTrue((scored["ml_neighbor_count"] >= app.MIN_LOF_NEIGHBORS).all())
        self.assertTrue((scored["ml_k_effective"] >= app.MIN_LOF_NEIGHBORS).all())
        self.assertTrue((result["ml_radius_m"] == 300).all())
        self.assertIn("10-20 cm", set(result["depth"]) if "depth" in result else {"10-20 cm"})

    def test_classification_requires_iqr_lof_and_regression(self):
        base = {"status": "outlier", "ml_status": "anomalia espacial", "reg_status": "anomalia temporal"}
        self.assertEqual(app.combine_detection(base), "outlier confirmado")
        base["reg_status"] = "normal"
        self.assertEqual(app.combine_detection(base), "atenção")
        base["status"] = "normal"
        self.assertEqual(app.combine_detection(base), "normal")
        base["ml_status"] = "normal"
        self.assertEqual(app.combine_detection(base), "normal")


if __name__ == "__main__":
    unittest.main()
