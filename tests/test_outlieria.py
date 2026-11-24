import io
import unittest

import numpy as np
import pandas as pd
from openpyxl import Workbook
from shapely.geometry import Polygon

import streamlit_app as app


class OutlieriaSpatialTests(unittest.TestCase):
    def test_radius_validation(self):
        self.assertEqual(app.validate_radius(125), 125.0)
        with self.assertRaises(ValueError):
            app.validate_radius(0)
        with self.assertRaises(ValueError):
            app.validate_radius(-10)

    @unittest.skipUnless(app.LOF_AVAILABLE, "scikit-learn não está instalado neste ambiente")
    def test_lof_keeps_depths_separate_and_uses_at_least_three(self):
        rows = []
        # Sete históricos próximos permitem k>=3 sem usar 2026 no treino.
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
        self.assertTrue((scored["ml_neighbor_count"] >= 3).all())
        self.assertTrue((scored["ml_k_effective"] >= 3).all())
        self.assertTrue((result["ml_radius_m"] == 300).all())
        self.assertIn("10-20 cm", set(result["depth"]) if "depth" in result else {"10-20 cm"})

    def test_classification_uses_lof_and_regression_only(self):
        base = {"status": "normal", "ml_status": "anomalia espacial", "reg_status": "anomalia temporal"}
        self.assertEqual(app.combine_detection(base), "outlier confirmado")
        base["ml_status"] = "normal"
        self.assertEqual(app.combine_detection(base), "anomalia temporal")
        base["reg_status"] = "normal"
        base["ml_status"] = "anomalia espacial"
        self.assertEqual(app.combine_detection(base), "anomalia espacial")
        base["ml_status"] = "normal"
        self.assertEqual(app.combine_detection(base), "normal")

    def test_spatial_regression_does_not_pair_by_point_id(self):
        rows = []
        # Os IDs mudam em cada ciclo; a coordenada permanece próxima do alvo.
        for year, point_id, value in [(2018, "A001", 2.0), (2021, "B044", 3.5), (2023, "C900", 5.0)]:
            rows.append({
                "year": str(year), "crop": "L01", "depth": "0-10 cm",
                "longitude": -50.0 + (year - 2018) * 0.000001, "latitude": -20.0,
                "point_id": point_id, "MO": value,
            })
        rows.append({
            "year": "2026", "crop": "L01", "depth": "0-10 cm",
            "longitude": -50.0, "latitude": -20.0, "point_id": "R012", "MO": 5.2,
        })
        sample = pd.DataFrame(rows)
        result = app.run_spatial_regression(sample, ["MO"], 300)
        self.assertEqual(len(result), 1)
        self.assertEqual(int(result.iloc[0]["reg_cycles"]), 3)
        self.assertNotEqual(result.iloc[0]["reg_status"], "dados insuficientes")

    def test_spatial_temporal_reference_ignores_id(self):
        historical = pd.DataFrame([
            {"year": "2021", "crop": "L01", "depth": "0-10 cm", "longitude": -50.0, "latitude": -20.0, "point_id": "OLD", "MO": 4.0},
            {"year": "2023", "crop": "L01", "depth": "0-10 cm", "longitude": -50.0, "latitude": -20.0, "point_id": "OTHER", "MO": 5.0},
        ])
        target = pd.Series({"crop": "L01", "depth": "0-10 cm", "longitude": -50.0, "latitude": -20.0, "point_id": "NEW"})
        series = app.spatial_temporal_reference(target, historical, "MO", 300)
        self.assertEqual([item[0] for item in series], [2021, 2023])
        self.assertEqual([item[1] for item in series], [4.0, 5.0])

    def test_ph_above_smp_is_flagged_on_same_point(self):
        result = pd.DataFrame([
            {"sample_index": 10, "element": "pH", "value": 6.2, "final_status": "normal"},
            {"sample_index": 10, "element": "SMP", "value": 5.8, "final_status": "normal"},
            {"sample_index": 11, "element": "pH", "value": 5.1, "final_status": "normal"},
            {"sample_index": 11, "element": "SMP", "value": 5.8, "final_status": "normal"},
        ])
        checked = app.apply_ph_smp_validation(result)
        self.assertEqual(checked.loc[0, "ph_smp_status"], "anomalia: pH > SMP")
        self.assertEqual(checked.loc[2, "ph_smp_status"], "")
        self.assertEqual(app.display_classification(checked.loc[0]), "anomalia (pH > SMP)")

    def test_display_table_has_only_core_columns_and_classification(self):
        result = pd.DataFrame([{
            "point_id": "C001", "crop": "L01", "depth": "0-10 cm", "element": "PH",
            "value": 6.2, "display_status": "anomalia (pH > SMP)", "q1": 1.0,
            "ml_score": 2.0, "reg_status": "anomalia temporal",
        }])
        displayed = app.frame_for_display(result)
        self.assertEqual(list(displayed.columns), ["ponto", "lavoura", "profundidade", "elemento", "valor", "classificação"])

    def test_parser_accepts_coordinate_aliases_and_decimal_comma(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "2026"
        sheet.append(["LONG", "LAT", "CAMADA", "PH"])
        sheet.append(["-54,2300", "-28,4980", "0-10 cm", 5.8])
        content = io.BytesIO()
        workbook.save(content)
        polygon = app.PolygonRecord(Polygon([(-55, -30), (-53, -30), (-53, -27), (-55, -27)]), 1.0)
        sample, elements, _ = app.parse_workbook(content.getvalue(), [polygon], "", "")
        self.assertEqual(len(sample), 1)
        self.assertEqual(elements, ["PH"])
        self.assertAlmostEqual(sample.iloc[0]["longitude"], -54.23)
        self.assertAlmostEqual(sample.iloc[0]["latitude"], -28.498)


if __name__ == "__main__":
    unittest.main()
