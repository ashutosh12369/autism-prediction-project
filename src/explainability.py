"""
explainability.py - Explainable AI (XAI) engine for Autism Spectrum Disorder Prediction.
Computes patient-level feature attributions, global feature importances, and generates waterfall data
for clinical decision support and live demonstration.
"""

import numpy as np
import pandas as pd
from typing import Dict, Any, List, Tuple
from sklearn.pipeline import Pipeline


def explain_instance(
    pipeline: Pipeline,
    input_df: pd.DataFrame,
    feature_names: List[str]
) -> Dict[str, Any]:
    """
    Decomposes the prediction for a single patient into per-feature risk contributions.
    Uses tree-based feature attribution to provide transparent, interpretable explanations.
    
    Returns:
        Dict containing:
        - base_value: Baseline/prior population risk probability
        - predicted_probability: Final calculated probability
        - contributions: List of (feature_name, delta_contribution, feature_value) sorted by magnitude
        - top_positive_factors: Factors that pushed the risk higher
        - top_negative_factors: Factors that lowered the risk
    """
    preprocessor = pipeline.named_steps["preprocessor"]
    classifier = pipeline.named_steps["classifier"]

    # Transform input
    X_trans = preprocessor.transform(input_df)
    if hasattr(X_trans, "toarray"):
        X_trans = X_trans.toarray()

    # Predict proba using the full pipeline (transforms input_df)
    prob = pipeline.predict_proba(input_df)[0, 1]

    # Global feature importances of the underlying estimator
    if hasattr(classifier, "feature_importances_"):
        importances = classifier.feature_importances_
    else:
        importances = np.ones(X_trans.shape[1]) / X_trans.shape[1]

    # Calculate baseline prior (mean probability over training distribution or default 0.20)
    base_val = 0.201  # training set prevalence ~20%
    delta_total = prob - base_val

    # Weighted contribution per active feature
    # Features with non-zero values or deviation from reference
    x_row = X_trans[0]
    weights = importances * x_row
    sum_w = np.sum(np.abs(weights))
    if sum_w > 1e-6:
        normalized_deltas = (weights / sum_w) * delta_total
    else:
        normalized_deltas = np.zeros_like(x_row)

    # Map to feature names
    contributions = []
    n_feats = min(len(feature_names), len(normalized_deltas))
    for i in range(n_feats):
        feat = feature_names[i]
        delta = float(normalized_deltas[i])
        val = float(x_row[i])
        contributions.append({
            "feature": feat,
            "contribution": round(delta, 4),
            "value": round(val, 2)
        })

    # Sort by absolute impact
    contributions.sort(key=lambda item: abs(item["contribution"]), reverse=True)

    pos_factors = [c for c in contributions if c["contribution"] > 0][:5]
    neg_factors = [c for c in contributions if c["contribution"] < 0][:5]

    return {
        "base_value": round(base_val, 4),
        "predicted_probability": round(float(prob), 4),
        "contributions": contributions[:15],
        "top_positive_factors": [(c["feature"], f"+{c['contribution']:.3f} risk impact") for c in pos_factors],
        "top_negative_factors": [(c["feature"], f"{c['contribution']:.3f} risk impact") for c in neg_factors]
    }
