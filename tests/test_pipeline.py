"""
test_pipeline.py - Unit tests for data pipeline, cleaning, and PDF report generation.
"""

import os
import sys
import tempfile
import unittest
import pandas as pd
import numpy as np

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from src.data_pipeline import clean_dataframe, build_preprocessor, get_feature_names
from src.report_generator import generate_pdf_report


class TestDataPipeline(unittest.TestCase):
    def test_clean_dataframe(self):
        sample_data = {
            "ID": [101, 102],
            "A1_Score": [1, 0],
            "A2_Score": [0, 1],
            "A3_Score": [1, 0],
            "A4_Score": [0, 1],
            "A5_Score": [1, 0],
            "A6_Score": [0, 1],
            "A7_Score": [1, 0],
            "A8_Score": [0, 1],
            "A9_Score": [1, 0],
            "A10_Score": [1, 0],
            "age": [38.17, 300.5],
            "gender": ["m", "f"],
            "ethnicity": ["?", "White-European"],
            "jaundice": ["no", "yes"],
            "austim": ["no", "yes"],
            "contry_of_res": ["Austria", "United States"],
            "used_app_before": ["no", "yes"],
            "result": [6.35, 2.25],
            "age_desc": ["18 and more", "18 and more"],
            "relation": ["Self", "?"],
            "Class/ASD": [0, 1]
        }
        df = pd.DataFrame(sample_data)
        cleaned = clean_dataframe(df)

        self.assertNotIn("ID", cleaned.columns)
        self.assertNotIn("age_desc", cleaned.columns)
        self.assertLessEqual(cleaned.loc[1, "age"], 95)
        self.assertEqual(cleaned.loc[0, "ethnicity"], "Others")
        self.assertEqual(cleaned.loc[1, "relation"], "Others")
        self.assertEqual(cleaned["gender"].tolist(), [1, 0])
        self.assertEqual(cleaned["jaundice"].tolist(), [0, 1])
        self.assertEqual(cleaned["austim"].tolist(), [0, 1])

    def test_build_preprocessor(self):
        sample_data = {
            "A1_Score": [1, 0],
            "A2_Score": [0, 1],
            "A3_Score": [1, 0],
            "A4_Score": [0, 1],
            "A5_Score": [1, 0],
            "A6_Score": [0, 1],
            "A7_Score": [1, 0],
            "A8_Score": [0, 1],
            "A9_Score": [1, 0],
            "A10_Score": [1, 0],
            "age": [35, 42],
            "gender": [1, 0],
            "ethnicity": ["White-European", "Asian"],
            "jaundice": [0, 1],
            "austim": [0, 1],
            "contry_of_res": ["United States", "India"],
            "used_app_before": [0, 0],
            "result": [5.0, 3.0],
            "relation": ["Self", "Parent"]
        }
        df = pd.DataFrame(sample_data)
        preprocessor = build_preprocessor(include_result=True)
        X_trans = preprocessor.fit_transform(df)

        self.assertEqual(X_trans.shape[0], 2)
        self.assertFalse(np.isnan(X_trans).any())

        feature_names = get_feature_names(preprocessor, include_result=True)
        self.assertEqual(len(feature_names), X_trans.shape[1])

    def test_generate_pdf_report(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_path = os.path.join(tmpdir, "test_report.pdf")
            patient_data = {
                "patient_name": "Demo Test Patient",
                "age": 28,
                "gender": "m",
                "ethnicity": "White-European",
                "jaundice": "no",
                "austim": "yes",
                "contry_of_res": "United States",
                "relation": "Self",
                "used_app_before": "no",
                "A1_Score": 1, "A2_Score": 1, "A3_Score": 1, "A4_Score": 1, "A5_Score": 1,
                "A6_Score": 1, "A7_Score": 0, "A8_Score": 1, "A9_Score": 1, "A10_Score": 1
            }
            prediction_result = {
                "probability": 0.88,
                "risk_level": "High Risk (>70%)",
                "decision": "ASD Indicators Identified",
                "top_factors": [
                    ("A1_Score (Sensory)", "+0.22 risk impact"),
                    ("A9_Score (Facial cues)", "+0.18 risk impact")
                ]
            }
            res_path = generate_pdf_report(patient_data, prediction_result, pdf_path)
            self.assertTrue(os.path.exists(res_path))
            self.assertGreater(os.path.getsize(res_path), 1000)


if __name__ == "__main__":
    unittest.main()
