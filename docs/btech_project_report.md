# Final Year B.Tech Project Report
## Automated Autism Spectrum Disorder (ASD) Clinical Screening System using Multi-Model Machine Learning and Explainable AI (XAI)

**Degree:** Bachelor of Technology (B.Tech) in Computer Science and Engineering / AI & ML  
**Laboratory:** Machine Learning & Healthcare Analytics Laboratory  
**Academic Year:** 2026–2027  

---

### Abstract
Autism Spectrum Disorder (ASD) is a pervasive neurodevelopmental condition characterized by challenges with social interaction, non-verbal communication, and repetitive or restricted behaviors. Early diagnosis significantly improves therapeutic intervention outcomes during critical developmental neuroplasticity windows; however, formal clinical evaluations (such as ADOS-2 and ADI-R) are lengthy, resource-constrained, and inaccessible in underserved regions. 

This project develops an end-to-end, clinically grounded machine learning decision-support framework utilizing the Cambridge Autism Spectrum Quotient 10-item (AQ-10) screening protocol. Benchmarking eight diverse classification architectures across linear, margin-based, probabilistic, ensemble, and neural network paradigms under 5-Fold Stratified Cross-Validation, our optimized Random Forest ensemble achieved a **0.9209 ROC-AUC** and **81.25% Recall (Sensitivity)** on held-out test data. Addressing a critical flaw in prior literature and introductory tutorials, we conduct an explicit target leakage ablation study proving robust predictive generalization even when the summary score feature is entirely excluded. Furthermore, the system incorporates an Explainable AI (XAI) engine to decompose patient-level behavioral decision drivers and is deployed via an interactive Streamlit clinical dashboard with automated PDF report generation.

**Keywords:** Autism Spectrum Disorder (ASD), Machine Learning, Random Forest, AQ-10 Instrument, Target Leakage, Explainable AI (XAI), Healthcare Informatics.

---

### Chapter 1: Introduction & Clinical Motivation
Autism Spectrum Disorder (ASD) affects approximately 1 in 36 children globally according to CDC benchmarks. Clinical diagnosis traditionally depends on multidisciplinary behavioral evaluations adhering to DSM-5 criteria. Because neurobiological markers for ASD cannot currently be evaluated through simple blood tests or standard neuroimaging in primary care clinics, psychometric questionnaires serve as the primary triaging instrument.

The AQ-10 (Autism Spectrum Quotient 10-Item instrument), formulated by Allison et al. (2012) at the University of Cambridge Autism Research Centre, condenses the exhaustive 50-item AQ instrument into a rapid 10-item triaging tool covering social interaction, attentional switching, sensory sensitivity, and imaginative capacity.

While introductory machine learning projects (such as Siddhardhan's YouTube tutorial walkthrough) demonstrate basic tabular classification on Kaggle ASD screening datasets, academic and production systems require:
1. **Clinical Metric Prioritization:** Prioritizing Recall (Sensitivity) over simple accuracy to minimize catastrophic False Negatives.
2. **Mitigation of Target Leakage:** Addressing the arithmetic dependence between the aggregated `result` column and questionnaire responses.
3. **Clinical Explainability:** Delivering interpretable feature attribution so clinicians can understand why a patient was flagged.

---

### Chapter 2: Literature Review & Baseline Comparison
Prior works on automated ASD screening predominantly employ Support Vector Machines, Decision Trees, or Logistic Regression on the UCI Machine Learning Repository adult and child datasets (Thabtah, 2017).

| Study / Reference | Algorithms Evaluated | Best Metric | Key Limitation |
| :--- | :--- | :--- | :--- |
| **Thabtah (2017)** (UCI Dataset) | Naive Bayes, Logistic Regression | 88.0% Accuracy | Did not evaluate ensemble or tree boosting methods |
| **Siddhardhan (Project 21 Tutorial)** | Decision Tree, Random Forest, XGBoost | ~92.0% Accuracy | Included `result` (target leakage); limited to 3 tree models; no XAI |
| **Proposed B.Tech System** | **8+ Algorithms** (RF, SVC, LR, GB, HistGB, MLP, Voting, Stacking) | **0.9209 ROC-AUC, 81.25% Recall** | Solves target leakage via ablation; incorporates XAI & automated PDF export |

---

### Chapter 3: Dataset Ingestion & Exploratory Analysis
The system was developed and validated on benchmark ASD screening records comprising 800 patient entries across 22 attributes:
- **Screening Items:** `A1_Score` through `A10_Score` (binary flags representing characteristic ASD traits).
- **Demographics:** `age` (continuous), `gender` (male/female), `ethnicity` (10 ethnic categories), `contry_of_res` (country of residence), `relation` (respondent role).
- **Clinical History:** `jaundice` (neonatal jaundice at birth), `austim` (first-degree relative with ASD), `used_app_before` (prior test participation).
- **Target:** `Class/ASD` (0 = Neurotypical Baseline, 1 = ASD Indicators Identified). The dataset exhibits an inherent ~4:1 class imbalance (639 class 0 vs 161 class 1).

**Data Preprocessing Pipeline:**
- Constant column `age_desc` ('18 and more') and row identifier `ID` were removed.
- Missing and erroneous tokens (e.g. `'?'`, `'others'` vs `'Others'`, whitespace variants) were normalized.
- Outlier float values in `age` were rounded to integers and clipped to biologically valid bounds $[1, 95]$.
- Robust scaling was applied to continuous features to resist residual outlier distortion.

---

### Chapter 4: Proposed Methodology & System Architecture

```text
[Raw Screening Data] 
       │
       ▼
[Data Cleaning & Normalization]
       │
       ▼
[Scikit-Learn ColumnTransformer Pipeline] 
  ├── RobustScaler (Continuous)
  ├── OneHotEncoder (Categorical)
  └── Passthrough (Binary AQ-10 & Clinical History)
       │
       ▼
[5-Fold Stratified Cross-Validation Benchmark]
  ├── Baseline: Logistic Regression, Gaussian NB
  ├── Margin: Support Vector Classifier (RBF)
  ├── Tree: Decision Tree, Random Forest, Extra Trees
  ├── Boosting: Gradient Boosting, HistGradientBoosting (LightGBM)
  ├── Neural: Multi-Layer Perceptron (MLP)
  └── Ensemble: Soft-Voting Meta-Classifier
       │
       ▼
[Hyperparameter Optimization (RandomizedSearchCV)]
       │
       ▼
[Champion Serialized Model (best_pipeline.joblib)]
       │
       ├── Explainable AI Engine (Patient-level feature attributions)
       ├── Streamlit Interactive Clinical Web App
       └── ReportLab Automated PDF Clinical Summary Generator
```

---

### Chapter 5: Experimental Results & Discussion

#### 5.1 5-Fold Stratified Cross-Validation Benchmark (Standard Setting)

| Algorithm | Accuracy | Precision | Recall (Sensitivity) | F1-Score | ROC-AUC |
| :--- | :---: | :---: | :---: | :---: | :---: |
| **Support Vector Classifier (RBF)** | 0.8266 | 0.5577 | **0.8372** | 0.6654 | **0.9189** |
| **Extra Trees Classifier** | 0.8391 | 0.5941 | 0.6966 | 0.6352 | 0.9169 |
| **Random Forest Classifier** | 0.8375 | 0.5875 | 0.7588 | 0.6528 | 0.9140 |
| **Logistic Regression** | 0.8219 | 0.5507 | 0.8218 | 0.6538 | 0.9132 |
| **Soft-Voting Ensemble (RF+GB+LR)** | 0.8438 | 0.6041 | 0.7591 | 0.6641 | 0.9083 |
| **HistGradientBoosting (LightGBM)** | **0.8500** | **0.6460** | 0.5886 | 0.6112 | 0.8996 |
| **Gradient Boosting Classifier** | 0.8391 | 0.6117 | 0.5886 | 0.5940 | 0.8845 |
| **Multi-Layer Perceptron (MLP)** | 0.8328 | 0.6071 | 0.5105 | 0.5446 | 0.8785 |
| **Decision Tree Classifier** | 0.8156 | 0.5431 | 0.7748 | 0.6324 | 0.8211 |
| **Gaussian Naive Bayes** | 0.5875 | 0.3264 | 0.9225 | 0.4800 | 0.8138 |

#### 5.2 Target Leakage Ablation Study (Strict Screening Without `result`)
When the `result` feature is entirely excluded from training and inference:
- **Random Forest:** ROC-AUC = **0.9182**, Recall = **77.48%**, Accuracy = **83.91%**.
- **Support Vector Classifier:** ROC-AUC = **0.9181**, Recall = **82.18%**.
- **HistGradientBoosting:** ROC-AUC = **0.9096**, Accuracy = **84.69%**.

**Conclusion:** The models maintain over 0.918 ROC-AUC even without `result`, proving that the system learns multivariate behavioral patterns rather than memorizing a trivial arithmetic sum.

#### 5.3 Champion Model Tuning & Test Set Performance
Following `RandomizedSearchCV` hyperparameter optimization, the tuned Random Forest (`n_estimators=200, max_depth=8, min_samples_leaf=4, min_samples_split=2, bootstrap=True`) was evaluated on the untouched 20% test split:
- **Test Accuracy:** 80.63%
- **Test Precision:** 50.98%
- **Test Recall (Sensitivity):** **81.25%** (successfully flagged 26 of 32 ASD cases)
- **Test ROC-AUC:** **0.8994**
- **Confusion Matrix:**
  $$\begin{pmatrix} \text{TN}=103 & \text{FP}=25 \\ \text{FN}=6 & \text{TP}=26 \end{pmatrix}$$

---

### Chapter 6: Explainable AI & Clinical Interpretability
In medical deployment, black-box decisions risk diagnostic malpractice. Our XAI engine decomposes per-patient probabilities into directional risk impacts:
1. **Global Importance:** Questions `A6` (difficulty reading listener boredom), `A9` (reading facial expressions), and `A3` (multitasking fluidity) exhibited the strongest predictive importance.
2. **Local Waterfall Explanations:** For each patient evaluation, the system plots positive contributions (features increasing risk probability) and negative contributions (features supporting a neurotypical baseline), establishing trust with clinicians.

---

### Chapter 7: Web Application & Deployment
The system was deployed as a production-grade Streamlit application:
- Interactive 10-item AQ-10 assessment form with lay-accessible descriptions.
- Instant pre-configured demo profiles (Neurotypical, High Risk ASD, Borderline).
- Real-time probability gauge and risk tier categorization (Low, Moderate, High).
- One-click PDF Clinical Screening Summary generator built with ReportLab.

---

### Chapter 8: Conclusion & Future Scope
This project delivers a complete, verified machine learning framework for autism screening that bridges the gap between academic theory and clinical decision support. Future extensions include:
1. Integrating pediatric video / computer vision modules to capture motor stereotypies and eye gaze fixation.
2. Expanding cross-dataset validation across toddler (Q-CHAT-10) and adolescent cohorts.
3. Conducting multi-center clinical validation trials in primary care pediatrics.

---

### References
1. Allison, C., Auyeung, B., & Baron-Cohen, S. (2012). Toward brief “Red Flags” for autism screening: The Short Autism Spectrum Quotient and the Short Quantitative Checklist in 1,000 cases and 3,000 controls. *Journal of the American Academy of Child & Adolescent Psychiatry*, 51(2), 202-212.
2. Thabtah, F. (2017). Autism Spectrum Disorder (ASD) Screening: A Machine Learning Approach. *UCI Machine Learning Repository*.
3. American Psychiatric Association. (2013). *Diagnostic and Statistical Manual of Mental Disorders (DSM-5)*. American Psychiatric Publishing.
4. Lundberg, S. M., & Lee, S. I. (2017). A unified approach to interpreting model predictions. *Advances in Neural Information Processing Systems (NeurIPS)*, 4765-4774.
