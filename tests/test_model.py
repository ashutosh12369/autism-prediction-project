"""
test_model.py - Unit tests for model inference, serialization, and explainability.
"""

import os
import sys
import unittest
import joblib
import pandas as pd

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from src.data_pipeline import clean_dataframe
from src.explainability import explain_instance


class TestModelInference(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
        model_path = os.path.join(project_root, "models", "best_pipeline.joblib")
        cls.pipeline = joblib.load(model_path)
        cls.feature_names = [f"feat_{i}" for i in range(100)]

    def test_low_risk_neurotypical_prediction(self):
        neurotypical_sample = pd.DataFrame({
            "age": [25],
            "gender": ["f"],
            "ethnicity": ["White-European"],
            "jaundice": ["no"],
            "austim": ["no"],
            "contry_of_res": ["United States"],
            "used_app_before": ["no"],
            "relation": ["Self"],
            "result": [0.0],
            "A1_Score": [0], "A2_Score": [0], "A3_Score": [0], "A4_Score": [0], "A5_Score": [0],
            "A6_Score": [0], "A7_Score": [0], "A8_Score": [0], "A9_Score": [0], "A10_Score": [0]
        })
        cleaned = clean_dataframe(neurotypical_sample)
        prob = self.pipeline.predict_proba(cleaned)[0, 1]
        self.assertGreaterEqual(prob, 0.0)
        self.assertLessEqual(prob, 0.50)  # Should predict low risk

    def test_high_risk_asd_prediction(self):
        asd_sample = pd.DataFrame({
            "age": [22],
            "gender": ["m"],
            "ethnicity": ["Asian"],
            "jaundice": ["yes"],
            "austim": ["yes"],
            "contry_of_res": ["India"],
            "used_app_before": ["no"],
            "relation": ["Self"],
            "result": [9.0],
            "A1_Score": [1], "A2_Score": [1], "A3_Score": [1], "A4_Score": [1], "A5_Score": [1],
            "A6_Score": [1], "A7_Score": [1], "A8_Score": [0], "A9_Score": [1], "A10_Score": [1]
        })
        cleaned = clean_dataframe(asd_sample)
        prob = self.pipeline.predict_proba(cleaned)[0, 1]
        self.assertGreaterEqual(prob, 0.50)  # Should predict high risk

    def test_explain_instance_structure(self):
        sample = pd.DataFrame({
            "age": [30],
            "gender": ["m"],
            "ethnicity": ["Others"],
            "jaundice": ["no"],
            "austim": ["no"],
            "contry_of_res": ["United Kingdom"],
            "used_app_before": ["no"],
            "relation": ["Self"],
            "result": [5.0],
            "A1_Score": [1], "A2_Score": [0], "A3_Score": [1], "A4_Score": [0], "A5_Score": [1],
            "A6_Score": [0], "A7_Score": [0], "A8_Score": [1], "A9_Score": [0], "A10_Score": [0]
        })
        cleaned = clean_dataframe(sample)
        explanation = explain_instance(self.pipeline, cleaned, self.feature_names)

        self.assertIn("base_value", explanation)
        self.assertIn("predicted_probability", explanation)
        self.assertIn("contributions", explanation)
        self.assertIsInstance(explanation["contributions"], list)
        self.assertGreater(len(explanation["contributions"]), 0)


if __name__ == "__main__":
    unittest.main()
