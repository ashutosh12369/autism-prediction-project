"""
data_pipeline.py - Autism Spectrum Disorder Screening Preprocessing & Transformation Pipeline.
Handles raw data cleaning, imputation, outlier clipping, and scikit-learn compatible ColumnTransformer.
"""

import os
import numpy as np
import pandas as pd
from typing import Tuple, List, Optional
from sklearn.compose import ColumnTransformer
from sklearn.preprocessing import OneHotEncoder, StandardScaler, RobustScaler
from sklearn.impute import SimpleImputer
from sklearn.pipeline import Pipeline


# Constant lists for validation and alignment
BINARY_QUESTIONS = [f"A{i}_Score" for i in range(1, 11)]
BINARY_DEMOGRAPHICS = ["gender", "jaundice", "austim", "used_app_before"]
CATEGORICAL_COLS = ["ethnicity", "contry_of_res", "relation"]
NUMERICAL_COLS_WITH_RESULT = ["age", "result"]
NUMERICAL_COLS_NO_RESULT = ["age"]

TOP_COUNTRIES = [
    "United States", "India", "New Zealand", "United Kingdom", "United Arab Emirates",
    "Australia", "Jordan", "Canada", "Afghanistan", "Sri Lanka", "Austria",
    "Netherlands", "Brazil", "Russia", "South Africa"
]

ETHNICITY_MAP = {
    "?": "Others",
    "others": "Others",
    "Others": "Others",
    "Middle Eastern ": "Middle Eastern",
    "middle eastern": "Middle Eastern",
    "White-European": "White-European",
    "white-european": "White-European",
    "Asian": "Asian",
    "asian": "Asian",
    "South Asian": "South Asian",
    "south asian": "South Asian",
    "Black": "Black",
    "black": "Black",
    "Hispanic": "Hispanic",
    "hispanic": "Hispanic",
    "Latino": "Latino",
    "latino": "Latino",
    "Pasifika": "Pasifika",
    "pasifika": "Pasifika",
    "Turkish": "Turkish",
    "turkish": "Turkish"
}

RELATION_MAP = {
    "?": "Others",
    "Self": "Self",
    "Parent": "Parent",
    "Relative": "Relative",
    "Others": "Others",
    "Health care professional": "Healthcare Professional"
}


def clean_dataframe(df: pd.DataFrame, drop_target_leakage: bool = False) -> pd.DataFrame:
    """
    Cleans raw ASD screening dataframe:
    - Drops ID and age_desc (zero variance)
    - Replaces '?' missing placeholders
    - Standardizes categorical spellings
    - Fixes and clips age values to plausible medical ranges (1-100)
    - Optionally drops 'result' column to eliminate target leakage
    """
    df = df.copy()

    # Rename jundice typo from UCI ARFF if present
    if "jundice" in df.columns and "jaundice" not in df.columns:
        df["jaundice"] = df["jundice"]

    # Drop non-informative columns
    cols_to_drop = [col for col in ["ID", "age_desc", "jundice"] if col in df.columns]
    if cols_to_drop:
        df = df.drop(columns=cols_to_drop)

    # Clean age: convert to numeric, clip outliers, round to nearest integer
    if "age" in df.columns:
        df["age"] = pd.to_numeric(df["age"], errors="coerce")
        median_age = df["age"].median() if not df["age"].isna().all() else 30.0
        df["age"] = df["age"].fillna(median_age)
        # Clip to [1, 95] to eliminate impossible sensor/synthetic anomalies
        df["age"] = df["age"].clip(lower=1.0, upper=95.0).round().astype(int)

    # Clean binary mappings
    binary_yes_no = ["jaundice", "austim", "used_app_before"]
    for col in binary_yes_no:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.lower().map({"yes": 1, "no": 0}).fillna(0).astype(int)

    if "gender" in df.columns:
        df["gender"] = df["gender"].astype(str).str.strip().str.lower().map({"m": 1, "f": 0}).fillna(1).astype(int)

    # Clean categorical text columns
    if "ethnicity" in df.columns:
        df["ethnicity"] = df["ethnicity"].astype(str).str.strip()
        df["ethnicity"] = df["ethnicity"].map(lambda x: ETHNICITY_MAP.get(x, "Others"))

    if "relation" in df.columns:
        df["relation"] = df["relation"].astype(str).str.strip()
        df["relation"] = df["relation"].map(lambda x: RELATION_MAP.get(x, "Others"))

    if "contry_of_res" in df.columns:
        df["contry_of_res"] = df["contry_of_res"].astype(str).str.strip()
        df["contry_of_res"] = df["contry_of_res"].apply(lambda x: x if x in TOP_COUNTRIES else "Other")

    # Clean AQ-10 scores
    for col in BINARY_QUESTIONS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).clip(0, 1).astype(int)

    # Handle result
    if "result" in df.columns:
        df["result"] = pd.to_numeric(df["result"], errors="coerce")
        # Impute missing result with sum of AQ-10 scores
        aq_sum = df[BINARY_QUESTIONS].sum(axis=1)
        df["result"] = df["result"].fillna(aq_sum)
        if drop_target_leakage:
            df = df.drop(columns=["result"])

    # Target column cleaning (handles YES/NO strings from UCI ARFF or 0/1 integers)
    if "Class/ASD" in df.columns:
        if df["Class/ASD"].dtype == object or df["Class/ASD"].astype(str).str.contains("YES|NO|yes|no").any():
            df["Class/ASD"] = df["Class/ASD"].astype(str).str.strip().str.upper().map({"YES": 1, "NO": 0, "1": 1, "0": 0}).fillna(0).astype(int)
        else:
            df["Class/ASD"] = pd.to_numeric(df["Class/ASD"], errors="coerce").fillna(0).astype(int)

    return df


def build_preprocessor(include_result: bool = True) -> ColumnTransformer:
    """
    Constructs a ColumnTransformer that handles:
    - Numerical features (RobustScaler for outlier resilience)
    - Multi-class categorical features (OneHotEncoder with handle_unknown='ignore')
    - Binary features passed through as-is
    """
    num_cols = NUMERICAL_COLS_WITH_RESULT if include_result else NUMERICAL_COLS_NO_RESULT

    num_transformer = Pipeline(steps=[
        ("imputer", SimpleImputer(strategy="median")),
        ("scaler", RobustScaler())
    ])

    cat_transformer = Pipeline(steps=[
        ("imputer", SimpleImputer(strategy="constant", fill_value="Others")),
        ("onehot", OneHotEncoder(handle_unknown="ignore", sparse_output=False))
    ])

    preprocessor = ColumnTransformer(
        transformers=[
            ("num", num_transformer, num_cols),
            ("cat", cat_transformer, CATEGORICAL_COLS),
            ("bin", "passthrough", BINARY_QUESTIONS + BINARY_DEMOGRAPHICS)
        ],
        remainder="drop"
    )

    return preprocessor


def get_feature_names(preprocessor: ColumnTransformer, include_result: bool = True) -> List[str]:
    """
    Extracts post-transformation feature names for SHAP and model explainability.
    """
    feature_names = []

    # Numerical features
    num_cols = NUMERICAL_COLS_WITH_RESULT if include_result else NUMERICAL_COLS_NO_RESULT
    feature_names.extend(num_cols)

    # Categorical one-hot features
    cat_step = preprocessor.named_transformers_["cat"]
    if hasattr(cat_step, "named_steps") and "onehot" in cat_step.named_steps:
        onehot = cat_step.named_steps["onehot"]
        if hasattr(onehot, "get_feature_names_out"):
            cat_features = onehot.get_feature_names_out(CATEGORICAL_COLS).tolist()
            feature_names.extend(cat_features)
        else:
            feature_names.extend(CATEGORICAL_COLS)
    else:
        feature_names.extend(CATEGORICAL_COLS)

    # Binary features
    feature_names.extend(BINARY_QUESTIONS + BINARY_DEMOGRAPHICS)

    return feature_names
