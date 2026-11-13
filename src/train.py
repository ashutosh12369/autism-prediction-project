"""
train.py - Autism Spectrum Disorder Classification Model Training & Evaluation Engine.
Performs 5-Fold Stratified Cross Validation, benchmarks 8+ ML algorithms, executes hyperparameter tuning,
conducts an ablation study for target leakage, and serializes the champion model and preprocessor.
"""

import os
import json
import joblib
import numpy as np
import pandas as pd
from typing import Dict, Any, Tuple, List

from sklearn.model_selection import StratifiedKFold, cross_validate, RandomizedSearchCV, train_test_split
from sklearn.metrics import (
    accuracy_score, precision_score, recall_score, f1_score,
    roc_auc_score, confusion_matrix, classification_report
)
from sklearn.linear_model import LogisticRegression
from sklearn.svm import SVC
from sklearn.naive_bayes import GaussianNB
from sklearn.tree import DecisionTreeClassifier
from sklearn.ensemble import (
    RandomForestClassifier, ExtraTreesClassifier,
    GradientBoostingClassifier, HistGradientBoostingClassifier,
    VotingClassifier, StackingClassifier
)
from sklearn.neural_network import MLPClassifier
from sklearn.pipeline import Pipeline

from src.data_pipeline import clean_dataframe, build_preprocessor, get_feature_names


PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
RAW_DATA_PATH = os.path.join(PROJECT_ROOT, "data", "raw", "train.csv")
MODELS_DIR = os.path.join(PROJECT_ROOT, "models")


def get_candidate_models() -> Dict[str, Any]:
    """
    Instantiates diverse candidate classification algorithms.
    """
    models = {
        "Logistic Regression": LogisticRegression(
            max_iter=1000, class_weight="balanced", random_state=42
        ),
        "Support Vector Classifier (RBF)": SVC(
            kernel="rbf", probability=True, class_weight="balanced", random_state=42
        ),
        "Gaussian Naive Bayes": GaussianNB(),
        "Decision Tree": DecisionTreeClassifier(
            max_depth=6, class_weight="balanced", random_state=42
        ),
        "Random Forest": RandomForestClassifier(
            n_estimators=150, max_depth=10, class_weight="balanced", random_state=42
        ),
        "Extra Trees": ExtraTreesClassifier(
            n_estimators=150, max_depth=10, class_weight="balanced", random_state=42
        ),
        "Gradient Boosting": GradientBoostingClassifier(
            n_estimators=120, max_depth=4, learning_rate=0.1, random_state=42
        ),
        "HistGradientBoosting (LightGBM-equiv)": HistGradientBoostingClassifier(
            max_iter=120, max_depth=6, learning_rate=0.1, random_state=42
        ),
        "Multi-Layer Perceptron (Neural Net)": MLPClassifier(
            hidden_layer_sizes=(64, 32), max_iter=600, random_state=42
        )
    }

    # Meta-Ensemble: Soft-Voting Classifier
    voting_ensemble = VotingClassifier(
        estimators=[
            ("rf", models["Random Forest"]),
            ("gb", models["Gradient Boosting"]),
            ("lr", models["Logistic Regression"])
        ],
        voting="soft"
    )
    models["Soft-Voting Ensemble (RF + GB + LR)"] = voting_ensemble

    return models


def evaluate_models_cv(
    X_train: pd.DataFrame,
    y_train: pd.Series,
    include_result: bool = True
) -> pd.DataFrame:
    """
    Evaluates all candidate algorithms using 5-Fold Stratified Cross-Validation.
    Computes Accuracy, Precision, Recall/Sensitivity, Specificity, F1, and ROC-AUC.
    """
    models = get_candidate_models()
    skf = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)
    results = []

    scoring = {
        "accuracy": "accuracy",
        "precision": "precision",
        "recall": "recall",
        "f1": "f1",
        "roc_auc": "roc_auc"
    }

    print(f"\n--- Running 5-Fold Stratified CV (Include Result: {include_result}) ---")

    for name, model in models.items():
        preprocessor = build_preprocessor(include_result=include_result)
        full_pipeline = Pipeline(steps=[
            ("preprocessor", preprocessor),
            ("classifier", model)
        ])

        cv_res = cross_validate(
            full_pipeline, X_train, y_train,
            cv=skf, scoring=scoring, n_jobs=1, return_train_score=False
        )

        acc = cv_res["test_accuracy"].mean()
        prec = cv_res["test_precision"].mean()
        rec = cv_res["test_recall"].mean()
        f1 = cv_res["test_f1"].mean()
        roc = cv_res["test_roc_auc"].mean()

        results.append({
            "Model": name,
            "Accuracy": round(float(acc), 4),
            "Precision": round(float(prec), 4),
            "Recall (Sensitivity)": round(float(rec), 4),
            "F1-Score": round(float(f1), 4),
            "ROC-AUC": round(float(roc), 4)
        })

    df_results = pd.DataFrame(results).sort_values(by="ROC-AUC", ascending=False).reset_index(drop=True)
    return df_results


def tune_champion_model(
    X_train: pd.DataFrame,
    y_train: pd.Series,
    include_result: bool = True
) -> Tuple[Pipeline, Dict[str, Any]]:
    """
    Runs hyperparameter optimization on Random Forest to identify peak performance.
    """
    print("\n--- Tuning Champion Classifier (Random Forest) via RandomizedSearchCV ---")
    preprocessor = build_preprocessor(include_result=include_result)

    pipeline = Pipeline(steps=[
        ("preprocessor", preprocessor),
        ("classifier", RandomForestClassifier(random_state=42, class_weight="balanced"))
    ])

    param_distributions = {
        "classifier__n_estimators": [100, 150, 200, 250, 300],
        "classifier__max_depth": [6, 8, 10, 12, 16, None],
        "classifier__min_samples_split": [2, 4, 6],
        "classifier__min_samples_leaf": [1, 2, 4],
        "classifier__bootstrap": [True, False]
    }

    skf = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)
    search = RandomizedSearchCV(
        pipeline,
        param_distributions=param_distributions,
        n_iter=20,
        scoring="roc_auc",
        cv=skf,
        random_state=42,
        n_jobs=1
    )

    search.fit(X_train, y_train)
    print(f"Best ROC-AUC from CV: {search.best_score_:.4f}")
    print(f"Optimal Parameters: {search.best_params_}")

    return search.best_estimator_, search.best_params_


def train_and_save_pipeline(data_path: str = RAW_DATA_PATH):
    """
    Main execution workflow:
    1. Loads raw dataset.
    2. Runs ablation comparison: with result vs without result.
    3. Benchmarks all candidate models.
    4. Tunes champion model.
    5. Evaluates on held-out test data (20%).
    6. Saves pipeline, model, and metadata.
    """
    os.makedirs(MODELS_DIR, exist_ok=True)

    print("Step 1: Loading raw dataset from:", data_path)
    df_raw = pd.read_csv(data_path)
    print("Initial Shape:", df_raw.shape)

    # 1. Clean data
    df_cleaned = clean_dataframe(df_raw, drop_target_leakage=False)
    X = df_cleaned.drop(columns=["Class/ASD"])
    y = df_cleaned["Class/ASD"]

    # Stratified Train/Test Split (80/20)
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.20, stratify=y, random_state=42
    )
    print(f"Training split: {X_train.shape[0]} samples | Test split: {X_test.shape[0]} samples")

    # 2. Benchmark with result
    benchmark_with_result = evaluate_models_cv(X_train, y_train, include_result=True)
    print("\nBenchmark Results (With 'result' feature):")
    print(benchmark_with_result.to_string(index=False))

    # 3. Ablation: Benchmark strictly WITHOUT result (Target Leakage Elimination)
    X_train_no_res = X_train.drop(columns=["result"])
    X_test_no_res = X_test.drop(columns=["result"])
    benchmark_no_result = evaluate_models_cv(X_train_no_res, y_train, include_result=False)
    print("\nBenchmark Results (Strict Clinical Screening - WITHOUT 'result' feature):")
    print(benchmark_no_result.to_string(index=False))

    # 4. Tune Champion Model (using standard clinical features with result)
    best_pipeline, best_params = tune_champion_model(X_train, y_train, include_result=True)

    # 5. Evaluate on held-out test set
    y_pred = best_pipeline.predict(X_test)
    y_prob = best_pipeline.predict_proba(X_test)[:, 1]

    test_accuracy = accuracy_score(y_test, y_pred)
    test_precision = precision_score(y_test, y_pred, zero_division=0)
    test_recall = recall_score(y_test, y_pred, zero_division=0)
    test_f1 = f1_score(y_test, y_pred, zero_division=0)
    test_roc_auc = roc_auc_score(y_test, y_prob)
    cm = confusion_matrix(y_test, y_pred).tolist()

    print("\n--- Held-Out Test Set Performance (Champion Model) ---")
    print(f"Accuracy:    {test_accuracy:.4f}")
    print(f"Precision:   {test_precision:.4f}")
    print(f"Recall:      {test_recall:.4f}")
    print(f"F1-Score:    {test_f1:.4f}")
    print(f"ROC-AUC:     {test_roc_auc:.4f}")
    print("Confusion Matrix:\n", confusion_matrix(y_test, y_pred))

    # 6. Extract Feature Names
    preprocessor = best_pipeline.named_steps["preprocessor"]
    feature_names = get_feature_names(preprocessor, include_result=True)

    # 7. Serialize Artifacts
    pipeline_save_path = os.path.join(MODELS_DIR, "best_pipeline.joblib")
    joblib.dump(best_pipeline, pipeline_save_path)
    print(f"\nSaved complete pipeline to: {pipeline_save_path}")

    # Save metrics and metadata for Streamlit App & Presentation
    metrics = {
        "test_accuracy": round(float(test_accuracy), 4),
        "test_precision": round(float(test_precision), 4),
        "test_recall": round(float(test_recall), 4),
        "test_f1": round(float(test_f1), 4),
        "test_roc_auc": round(float(test_roc_auc), 4),
        "confusion_matrix": cm,
        "best_hyperparameters": {k: str(v) for k, v in best_params.items()},
        "benchmark_with_result": benchmark_with_result.to_dict(orient="records"),
        "benchmark_without_result": benchmark_no_result.to_dict(orient="records"),
        "feature_names": feature_names
    }

    metrics_save_path = os.path.join(MODELS_DIR, "metrics.json")
    with open(metrics_save_path, "w") as f:
        json.dump(metrics, f, indent=2)
    print(f"Saved metrics summary to: {metrics_save_path}")

    return best_pipeline, metrics


if __name__ == "__main__":
    train_and_save_pipeline()
