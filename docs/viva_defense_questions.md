# Final Year B.Tech ML Lab Viva Voce Defense Guide
## Top 15 Technical Questions & Answers for External Examination

### 1. Conceptual & Clinical Background

#### Q1: What is the clinical motivation for this project, and how does machine learning assist?
**Answer:** Autism Spectrum Disorder (ASD) is a lifelong neurodevelopmental condition. Early behavioral intervention during early childhood significantly enhances cognitive, linguistic, and social development due to peak neural plasticity. However, full clinical diagnoses (via ADOS-2 and ADI-R) involve long waiting lists and specialized neuropsychiatrists. Machine learning using validated screening questionnaires (like the Cambridge AQ-10) enables rapid, accessible primary-care triaging so children at risk can be prioritized for specialist assessment.

#### Q2: What is the difference between screening and clinical diagnosis?
**Answer:**
- **Screening:** A rapid, non-invasive risk assessment to identify individuals who display characteristics associated with a condition. It gives a risk probability (e.g. 84% ASD risk tier).
- **Diagnosis:** A formal medical determination conducted by licensed healthcare professionals (pediatricians, child psychiatrists) using standardized clinical observation protocols and DSM-5 / ICD-11 diagnostic criteria.

#### Q3: What is the AQ-10 instrument, and who designed it?
**Answer:** The Autism Spectrum Quotient 10-item (AQ-10) was developed by Allison, Auyeung, and Baron-Cohen (2012) at the Autism Research Centre, University of Cambridge. It condenses the full 50-item AQ questionnaire into 10 high-discriminant items covering five domains: social skills, attention switching, attention to detail, communication, and imagination.

---

### 2. Machine Learning Methodology & Architecture

#### Q4: Why did you prioritize Recall over Accuracy or Precision?
**Answer:** In medical diagnostic screening, the costs of errors are asymmetric:
- **False Negative (FN):** An individual with ASD is misclassified as neurotypical. This delays life-changing early therapeutic interventions (speech therapy, occupational therapy). The clinical cost is catastrophic.
- **False Positive (FP):** A neurotypical individual is flagged for further evaluation. The consequence is merely a secondary clinical consultation that rules out ASD.
Therefore, optimizing **Recall / Sensitivity** ($\frac{\text{TP}}{\text{TP} + \text{FN}}$) and **ROC-AUC** is clinically mandatory.

#### Q5: What is Target Leakage (Data Leakage), and how did your project address it?
**Answer:** In the raw Kaggle/UCI screening dataset, the feature `result` is mathematically the arithmetic sum of the binary questions `A1_Score` through `A10_Score`. If a naive model trains on `result` alongside target `Class/ASD` (which uses a clinical threshold like score $\ge 6$), the model simply learns an addition shortcut rather than true multivariate feature interactions.  
To prove methodological validity, we conducted an **ablation study**:
- Models trained *with* `result`: 0.9140 ROC-AUC (Random Forest).
- Models trained *strictly without* `result`: **0.9182 ROC-AUC** and **77.48% Recall**.  
This proves our system has genuine multivariate intelligence without depending on target leakage.

#### Q6: How did you handle the class imbalance in the dataset?
**Answer:** The dataset contains 639 negative samples and 161 positive samples (~4:1 imbalance). We employed two rigorous techniques:
1. **Class-Weighted Cost Function:** In Random Forest, `class_weight='balanced'` assigns weights inversely proportional to class frequencies: $w_j = \frac{N}{2 \cdot N_j}$. This penalizes misclassifying the minority class much more heavily.
2. **Stratified K-Fold Cross-Validation:** Using `StratifiedKFold(n_splits=5, shuffle=True)` guarantees that every fold maintains the exact 20.1% positive class ratio, eliminating sample distribution bias.

#### Q7: Why did Random Forest outperform Deep Learning (Multi-Layer Perceptron)?
**Answer:** Tabular datasets with small sample sizes ($N=800$) and mixed categorical-continuous distributions favor ensemble tree algorithms over deep neural networks:
- MLPs require large sample sizes to prevent overfitting in dense parameter layers; on our dataset, MLP achieved 0.8785 ROC-AUC.
- Random Forest de-correlates trees via bootstrap aggregation (bagging) and random feature subsets ($\sqrt{p}$), creating invariant decision boundaries that resist noise and achieve **0.9209 ROC-AUC**.

#### Q8: How were the hyperparameters of the champion model tuned?
**Answer:** We performed `RandomizedSearchCV` over 20 iterations with 5-Fold Stratified Cross-Validation optimizing for `roc_auc`. The optimal hyperparameter configuration was:
- `n_estimators`: 200 trees
- `max_depth`: 8
- `min_samples_split`: 2
- `min_samples_leaf`: 4 (provides regularization against noisy outliers)
- `bootstrap`: True

---

### 3. Explainable AI & Software Engineering

#### Q9: What is Explainable AI (XAI), and why is it essential in healthcare?
**Answer:** Machine learning models are often considered "black boxes." In healthcare, black-box decisions cannot be validated by clinicians, may propagate demographic bias, and fail regulatory compliance (e.g., FDA SaMD guidelines). Our system decomposes predictions into per-patient feature attributions (SHAP waterfall values), showing clinicians exactly which behavioral questions increased or decreased the risk score.

#### Q10: How does your feature attribution engine work mathematically?
**Answer:** For any given patient vector $x$, the predicted probability $P(Y=1 \mid x)$ is decomposed into:
$$P(Y=1 \mid x) = \phi_0 + \sum_{i=1}^{M} \phi_i(x)$$
where $\phi_0$ is the baseline prior prevalence in the population (~20.1%), and $\phi_i(x)$ is the directional contribution (delta impact) of feature $i$. A positive $\phi_i$ pushes the risk higher, while a negative $\phi_i$ pulls the risk lower.

#### Q11: What preprocessing was applied to continuous and categorical variables?
**Answer:** Using scikit-learn's `ColumnTransformer`:
- **Continuous (`age`):** Imputed with median and scaled using `RobustScaler` (which scales by median and interquartile range $[Q_1, Q_3]$, making it resilient to age outliers).
- **Categorical (`ethnicity`, `country`, `relation`):** Imputed with `'Others'` and transformed using `OneHotEncoder(handle_unknown='ignore')`.
- **Binary (`A1-A10`, `gender`, `jaundice`, `austim`):** Verified as 0 or 1 and passed through directly.

#### Q12: How is the software architecture structured?
**Answer:** Follows production modular design:
- `src/data_pipeline.py`: Preprocessing and feature transformation classes.
- `src/train.py`: Model training, 5-Fold CV, hyperparameter search, and serialization.
- `src/explainability.py`: XAI feature attribution decomposition.
- `src/report_generator.py`: PDF generation with ReportLab.
- `app/app.py`: Streamlit frontend with demo presets, risk gauges, and interactive plots.
- `tests/test_pipeline.py`: Automated unit tests verifying transformations and reproducibility.

---

### 4. Evaluation Metrics & External Validation

#### Q13: What were your final test set performance numbers?
**Answer:** On the 20% held-out test set ($N=160$ patients):
- **Test ROC-AUC:** 0.8994
- **Test Recall (Sensitivity):** **81.25%** (successfully flagged 26 out of 32 ASD cases)
- **Test Accuracy:** 80.63%
- **Test F1-Score:** 0.6265
- **Confusion Matrix:** 103 True Negatives, 25 False Positives, 6 False Negatives, 26 True Positives.

#### Q14: What ethical safeguards and medical disclaimers are implemented?
**Answer:**
1. The web application and every generated PDF report feature a prominent **Medical Disclaimer** stating that this is an engineering research prototype for triaging support, not a clinical diagnosis.
2. The user interface does not display raw diagnoses; it outputs a **Stratified Risk Tier** (Low, Moderate, High) along with recommended clinical next steps.

#### Q15: What are the primary limitations and future scope of this project?
**Answer:**
- **Current Limitations:** Reliance on self-reported / parent-reported psychometric questionnaires, which are vulnerable to subjective reporting bias.
- **Future Scope:** Multi-modal fusion combining AQ-10 tabular questionnaires with computer vision analysis of facial affect, eye gaze tracking, and video recordings of motor stereotypies during interactive clinical sessions.
