# 🧠 Autism Spectrum Disorder (ASD) Prediction & Clinical Decision Support System

> **Final Year B.Tech Project in Machine Learning & Healthcare Informatics**  
> Based on the Cambridge AQ-10 Clinical Screening Protocol, inspired by Siddhardhan's ML Project series, elevated with multi-model benchmarking, target leakage ablation, explainable AI (XAI), and interactive Streamlit deployment.

---

## 📌 Project Overview
Autism Spectrum Disorder (ASD) is a pervasive neurodevelopmental condition. Early screening is critical because behavioral interventions (speech and occupational therapy) yield maximum developmental gains when initiated during early childhood neuroplasticity.

This project delivers:
1. **Clinical Triaging Engine:** Rapid behavioral risk evaluation using the validated Cambridge AQ-10 instrument.
2. **Multi-Model Machine Learning:** Benchmarked 8+ classification algorithms under 5-Fold Stratified Cross-Validation.
3. **Academic Target Leakage Ablation Study:** Solved the mathematical dependence between `result` and AQ-10 questions.
4. **Explainable AI (XAI):** Decomposes patient risk scores into positive and negative behavioral decision drivers.
5. **Interactive Full-Stack Web App:** Live Streamlit demo with quick-load clinical profiles, real-time probability gauge, and automated PDF clinical screening report generation.

---

## 🏆 Model Performance Benchmark (5-Fold Stratified CV)

| Algorithm Family | Model | Accuracy | Recall (Sensitivity) | F1-Score | ROC-AUC |
| :--- | :--- | :---: | :---: | :---: | :---: |
| **Ensemble (Bagging)** | **Random Forest (Tuned)** | **83.91%** | **77.48%** | **0.6610** | **0.9209** |
| **Margin-Based** | Support Vector Classifier (RBF) | 81.87% | **82.18%** | 0.6525 | 0.9181 |
| **Ensemble (Extreme)** | Extra Trees Classifier | 84.38% | 69.63% | 0.6390 | 0.9178 |
| **Meta-Ensemble** | Soft-Voting (RF + GB + LR) | 84.06% | 73.54% | 0.6476 | 0.9148 |
| **Linear / Regularized**| Logistic Regression | 81.72% | 81.42% | 0.6457 | 0.9111 |
| **Boosting (Histogram)**| HistGradientBoosting (LightGBM-equiv) | 84.69% | 56.49% | 0.5941 | 0.9096 |
| **Boosting (Sequential)**| Gradient Boosting | 83.75% | 55.66% | 0.5719 | 0.9036 |
| **Neural Network** | Multi-Layer Perceptron (MLP) | 84.06% | 52.58% | 0.5577 | 0.8766 |
| **Decision Tree** | Decision Tree Classifier | 81.41% | 76.80% | 0.6279 | 0.8203 |

> **Held-out 20% Test Split Performance (Champion Tuned Random Forest):**
> - **Test ROC-AUC:** `0.8994`
> - **Test Recall (Sensitivity):** `81.25%` (Correctly flagged 26 of 32 ASD cases)
> - **Test Accuracy:** `80.63%`

---

## 🚀 Quick Start & Live Demo

### 1. Launch the Web Application
Simply run the startup batch file or execute in PowerShell/Command Prompt:

```bash
# Option A: Double-click run_demo.bat
run_demo.bat

# Option B: Run via Streamlit command
streamlit run app/app.py
```
Open your browser at: `http://localhost:8501`

### 2. How to Give a 5-Minute Lab / Viva Presentation:
1. **Explain the Clinical Problem:** Open **Tab 1**. Mention that traditional clinical evaluation has 12–18 month waitlists, and machine learning with AQ-10 allows early screening triaging.
2. **Show the Live Demo (Zero Risk):**
   - Click **`🔴 Profile: ASD Indicators (High Risk)`** ➔ Click **`Run Machine Learning Diagnostic Screening`** ➔ Shows **87%+ Risk Probability**, Red Badge, and the SHAP waterfall chart explaining which answers drove the risk.
   - Click **`🟢 Profile: Neurotypical (Low Risk)`** ➔ Shows **12% Risk**, Green Badge, and protective neurotypical factors.
3. **Download PDF Report:** Click **`Download Official Clinical Screening Summary (PDF)`** to show external examiners the formatted clinical document ready for doctors.
4. **Show Academic Rigor (Tab 2):** Show the 8-model comparative table and explain the **Target Leakage Ablation Study** (proving the model still achieves 0.918 ROC-AUC even without the `result` sum feature).
5. **Answer Examiner Questions (Tab 3):** Tab 3 has the answers to the top 5 examiner questions directly visible on screen.

---

## 📂 Repository Structure
```text
autism_prediction_project/
├── data/
│   ├── raw/
│   │   ├── train.csv              # 800-patient dataset from video
│   │   └── Autism-Adult-Data.arff # Benchmark UCI dataset
├── src/
│   ├── data_pipeline.py           # Preprocessing & ColumnTransformer
│   ├── train.py                   # 5-Fold Stratified CV & Model Tuning
│   ├── explainability.py          # Explainable AI (XAI) Attribution
│   └── report_generator.py        # PDF report generator with ReportLab
├── models/
│   ├── best_pipeline.joblib       # Tuned champion model pipeline
│   └── metrics.json               # Benchmark results & feature names
├── app/
│   └── app.py                     # Interactive Streamlit application
├── tests/
│   ├── test_pipeline.py           # Unit tests for preprocessing & PDF
│   └── test_model.py              # Unit tests for predictions & XAI
├── docs/
│   ├── btech_project_report.md    # Complete IEEE-style thesis report
│   └── viva_defense_questions.md  # 15 Viva Voce defense questions & answers
├── run_demo.bat                   # 1-Click Windows Batch launcher
├── run_demo.ps1                   # 1-Click PowerShell launcher
├── requirements.txt
└── README.md
```

---

## 🧪 Automated Testing
Run the automated test suite with pytest:
```bash
pytest -v
```
*All 6 unit tests verify data cleaning, model inference, XAI outputs, and PDF creation.*
