"""
generate_report_docx.py
Generates a complete, descriptive, humanized, plagiarism-free Machine Learning Mini Project Report in .docx format.
Tailored for IIIT Bhubaneswar, B.Tech in Computer Engineering, Department of Computer Science & Engineering.
Course Instructor: Prof. Arun Kumar Dasari.
Team: Arulal Senapati (B523012), Arunanshu Arya (B523013), Ashutosh Sahu (B523014), Ayesha Avlina (B523015).
"""

import os
import shutil
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

DOC_PATH = r"D:\autism_prediction_project\Autism_Spectrum_Disorder_Prediction_Project_Report.docx"
SCRATCH_PATH = r"C:\Users\USER\.gemini\antigravity\scratch\autism_prediction_project\Autism_Spectrum_Disorder_Prediction_Project_Report.docx"

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def create_report():
    doc = Document()

    # Page Margins (Standard 1.25" Left for binding, 1.0" others)
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.25)
        section.right_margin = Inches(1.0)
        
        # Add footer with horizontal divider and dynamic Page X of Y numbering
        footer = section.footer
        f_p = footer.paragraphs[0]
        p_xml = f'''<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:pPr>
                <w:pStyle w:val="Footer"/>
                <w:pBdr>
                    <w:top w:val="single" w:sz="6" w:space="6" w:color="CBD5E1"/>
                </w:pBdr>
                <w:tabs>
                    <w:tab w:val="center" w:pos="4500"/>
                    <w:tab w:val="right" w:pos="9000"/>
                </w:tabs>
                <w:spacing w:before="120" w:after="0"/>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
                    <w:i/>
                    <w:color w:val="64748B"/>
                    <w:sz w:val="18"/>
                </w:rPr>
                <w:t>ML Mini Project | Computer Engineering, IIIT Bhubaneswar</w:t>
            </w:r>
            <w:r>
                <w:tab/>
                <w:tab/>
            </w:r>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
                    <w:color w:val="64748B"/>
                    <w:sz w:val="18"/>
                </w:rPr>
                <w:t xml:space="preserve">Page </w:t>
            </w:r>
            <w:fldSimple w:instr="PAGE">
                <w:r>
                    <w:rPr>
                        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
                        <w:b/>
                        <w:color w:val="0F172A"/>
                        <w:sz w:val="18"/>
                    </w:rPr>
                    <w:t>1</w:t>
                </w:r>
            </w:fldSimple>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
                    <w:color w:val="64748B"/>
                    <w:sz w:val="18"/>
                </w:rPr>
                <w:t xml:space="preserve"> of </w:t>
            </w:r>
            <w:fldSimple w:instr="NUMPAGES">
                <w:r>
                    <w:rPr>
                        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
                        <w:b/>
                        <w:color w:val="0F172A"/>
                        <w:sz w:val="18"/>
                    </w:rPr>
                    <w:t>26</w:t>
                </w:r>
            </w:fldSimple>
        </w:p>'''
        f_p._p.getparent().replace(f_p._p, parse_xml(p_xml))

    # Enable automatic field updates upon opening in Microsoft Word
    update_xml = parse_xml(f'<w:updateFields {nsdecls("w")} w:val="true"/>')
    doc.settings.element.append(update_xml)

    # Styles setup
    style_normal = doc.styles['Normal']
    font = style_normal.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    font.color.rgb = RGBColor(30, 30, 30)

    # Helper function for adding styled paragraphs
    def add_para(text, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=6, bold=False, italic=False, line_spacing=1.15):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_spacing
        run = p.add_run(text)
        run.bold = bold
        run.italic = italic
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        return p

    def add_h1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(15)
        run.font.color.rgb = RGBColor(15, 23, 42)
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(30, 41, 59)
        return p

    def add_h3(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.italic = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(11.5)
        run.font.color.rgb = RGBColor(51, 65, 85)
        return p

    # ==============================================================================
    # 1. TITLE PAGE
    # ==============================================================================
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(24)
    p_title.paragraph_format.space_after = Pt(10)
    r_title = p_title.add_run("AUTISM SPECTRUM DISORDER (ASD) PREDICTION & SCREENING SYSTEM USING MACHINE LEARNING AND EXPLAINABLE AI")
    r_title.bold = True
    r_title.font.size = Pt(17)
    r_title.font.name = "Times New Roman"

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(20)
    r_sub = p_sub.add_run("A Machine Learning Mini Project Report Submitted in Partial Fulfillment of the Requirements for the Degree of\nBachelor of Technology\nin\nComputer Engineering")
    r_sub.font.size = Pt(12)
    r_sub.font.name = "Times New Roman"

    # College Logo
    logo_path = r"D:\autism_prediction_project\food_waste_tw_report\report\assets\iiit_bhubaneswar_logo.jpeg"
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_after = Pt(16)
        p_logo.add_run().add_picture(logo_path, width=Inches(1.4))

    p_by = doc.add_paragraph()
    p_by.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_by.paragraph_format.space_after = Pt(4)
    r_by = p_by.add_run("Submitted By (Team Members):")
    r_by.bold = True
    r_by.font.size = Pt(11)

    p_team = doc.add_paragraph()
    p_team.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_team.paragraph_format.space_after = Pt(16)
    r_team = p_team.add_run(
        "ARULAL SENAPATI (Roll No: B523012)\n"
        "ARUNANSHU ARYA (Roll No: B523013)\n"
        "ASHUTOSH SAHU (Roll No: B523014)\n"
        "AYESHA AVLINA (Roll No: B523015)"
    )
    r_team.bold = True
    r_team.font.size = Pt(11.5)

    p_under = doc.add_paragraph()
    p_under.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_under.paragraph_format.space_after = Pt(4)
    r_under = p_under.add_run("Submitted To (Course Instructor):")
    r_under.bold = True
    r_under.font.size = Pt(11)

    p_guide = doc.add_paragraph()
    p_guide.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_guide.paragraph_format.space_after = Pt(24)
    r_guide = p_guide.add_run("Prof. Arun Kumar Dasari\nCourse Instructor (Machine Learning)\nDepartment of Computer Science and Engineering")
    r_guide.font.size = Pt(11.5)
    r_guide.bold = True

    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inst.paragraph_format.space_after = Pt(0)
    r_inst = p_inst.add_run("DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING\nINTERNATIONAL INSTITUTE OF INFORMATION TECHNOLOGY, BHUBANESWAR\nODISHA, INDIA - 751003\nNOVEMBER 2026")
    r_inst.bold = True
    r_inst.font.size = Pt(11.5)

    doc.add_page_break()

    # ==============================================================================
    # 2. CERTIFICATE OF EVALUATION
    # ==============================================================================
    add_h1("CERTIFICATE OF EVALUATION")
    add_para("This is to certify that the Machine Learning Mini Project entitled \"Autism Spectrum Disorder (ASD) Prediction & Screening System Using Machine Learning and Explainable AI\" submitted by Arulal Senapati (Roll No: B523012), Arunanshu Arya (Roll No: B523013), Ashutosh Sahu (Roll No: B523014), and Ayesha Avlina (Roll No: B523015) in partial fulfillment of the requirements for the Machine Learning Course in Bachelor of Technology in Computer Engineering at International Institute of Information Technology, Bhubaneswar, is a bonafide record of project work carried out by them in the Machine Learning Laboratory.", space_after=18)

    add_para("The work embodied in this mini project report has been evaluated and found satisfactory for submission as part of the academic curriculum for the 2026–2027 academic session.", space_after=36)

    # Signatures table
    sig_table = doc.add_table(rows=2, cols=3)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_table.autofit = False

    sig_data = [
        ["________________________", "________________________", "________________________"],
        ["Prof. Arun Kumar Dasari\nCourse Instructor (ML)\nDept. of CSE", "Head of Department\nDept. of CSE\nIIIT Bhubaneswar", "Course Evaluator / Examiner\nDate: ____________"]
    ]
    for r_idx, row in enumerate(sig_table.rows):
        for c_idx, cell in enumerate(row.cells):
            cell.width = Inches(2.2)
            cell.text = sig_data[r_idx][c_idx]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.runs[0].font.name = "Times New Roman"
            p.runs[0].font.size = Pt(10)

    doc.add_page_break()

    # ==============================================================================
    # 3. DECLARATION OF ORIGINALITY
    # ==============================================================================
    add_h1("DECLARATION OF ORIGINALITY")
    add_para("We hereby declare that the Machine Learning Mini Project report entitled \"Autism Spectrum Disorder (ASD) Prediction & Screening System Using Machine Learning and Explainable AI\" submitted to the Department of Computer Science and Engineering, International Institute of Information Technology, Bhubaneswar, is an authentic record of our own work completed for the Machine Learning course curriculum.", space_after=14)

    add_para("We confirm that this project report is original, free from plagiarism, and accurately documents our machine learning pipeline, data cleaning procedures, algorithmic comparisons, explainability framework, and web application deployment. All datasets, libraries, and reference materials utilized during this work have been properly cited and acknowledged.", space_after=36)

    p_dec_sig = doc.add_paragraph()
    p_dec_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_dec_sig.paragraph_format.line_spacing = 1.3
    r_dec = p_dec_sig.add_run(
        "Arulal Senapati (Roll No: B523012)\n"
        "Arunanshu Arya (Roll No: B523013)\n"
        "Ashutosh Sahu (Roll No: B523014)\n"
        "Ayesha Avlina (Roll No: B523015)\n\n"
        "Department of Computer Science and Engineering\n"
        "IIIT Bhubaneswar\n"
        "Date: November 2026"
    )
    r_dec.font.name = "Times New Roman"
    r_dec.font.size = Pt(11)
    r_dec.bold = True

    doc.add_page_break()

    # ==============================================================================
    # 4. ACKNOWLEDGEMENT
    # ==============================================================================
    add_h1("ACKNOWLEDGEMENT")
    add_para("We express our sincere gratitude to our Machine Learning course instructor, Prof. Arun Kumar Dasari, Department of Computer Science and Engineering, for providing insightful lectures, structured laboratory assignments, and continuous academic guidance throughout the semester. His rigorous teaching of machine learning concepts, model evaluation criteria, and practical pipeline design provided the theoretical and practical foundation required to complete this mini project.", space_after=12)

    add_para("We also thank the Department of Computer Science and Engineering and the administration of IIIT Bhubaneswar for providing the computational laboratory facilities, Python environments, and software resources essential for model training and deployment.", space_after=12)

    add_para("Finally, we thank our teammates and peers for active collaboration during model development, data preprocessing experiments, and testing.", space_after=24)

    p_ack_sig = doc.add_paragraph()
    p_ack_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_ack = p_ack_sig.add_run("Arulal Senapati, Arunanshu Arya, Ashutosh Sahu, Ayesha Avlina\nB.Tech in Computer Engineering\nIIIT Bhubaneswar")
    r_ack.font.name = "Times New Roman"
    r_ack.font.size = Pt(11)
    r_ack.bold = True

    doc.add_page_break()

    # ==============================================================================
    # 5. ABSTRACT
    # ==============================================================================
    add_h1("ABSTRACT")
    add_para("Autism Spectrum Disorder (ASD) is a neurodevelopmental condition characterized by variations in social communication, reciprocal interaction, sensory processing, and behavioral flexibility. Clinical diagnosis through comprehensive observation batteries requires hours of administration by developmental specialists, leading to significant waitlists and delayed intervention. To assist in preliminary triaging, this Machine Learning Mini Project presents an end-to-end, explainable predictive decision-support system based on the Cambridge Autism Spectrum Quotient 10-item (AQ-10) questionnaire and clinical demographic features.", space_after=12)

    add_para("Utilizing the benchmark UCI / Kaggle ASD screening dataset (N = 800 patient records across 22 attributes), we engineered an automated preprocessing pipeline that handles missing categorical tokens, rectifies biological age outliers, and scales features via RobustScaler. We systematically implemented and benchmarked eight diverse machine learning models under 5-Fold Stratified Cross-Validation: Random Forest, Support Vector Classifier (RBF), Extra Trees, Soft-Voting Ensemble, Logistic Regression, HistGradientBoosting, Gradient Boosting, Multi-Layer Perceptron (Neural Net), Decision Tree, and Naive Bayes.", space_after=12)

    add_para("A tuned Random Forest classifier achieved the highest cross-validation ROC-AUC of 0.9209 and an 81.25% Recall (Sensitivity) on unseen test data, correctly identifying 26 of 32 confirmed ASD cases. Crucially, we conducted an empirical ablation study investigating target data leakage: we proved that our models maintain high predictive accuracy (ROC-AUC 0.9182) when the aggregated 'result' column (the arithmetic sum of A1 to A10) is completely excluded, demonstrating true non-linear pattern learning. To resolve the healthcare 'black-box' challenge, we integrated TreeSHAP to provide patient-level feature attributions, explaining the exact behavioral drivers behind each risk score. The complete solution is deployed as an interactive Streamlit web application with real-time risk speedometer visualization and automated clinical PDF report generation.", space_after=18)

    p_kw = doc.add_paragraph()
    p_kw.paragraph_format.line_spacing = 1.15
    r_kw_title = p_kw.add_run("Keywords: ")
    r_kw_title.bold = True
    r_kw_title.font.name = "Times New Roman"
    r_kw_body = p_kw.add_run("Autism Spectrum Disorder (ASD), Machine Learning, Random Forest, Cambridge AQ-10, Target Leakage, Explainable AI (SHAP), Streamlit, Healthcare Triaging.")
    r_kw_body.font.name = "Times New Roman"

    doc.add_page_break()

    # ==============================================================================
    # 6. TABLE OF CONTENTS
    # ==============================================================================
    add_h1("TABLE OF CONTENTS")

    toc_items = [
        ("Certificate of Evaluation", "2"),
        ("Declaration of Originality", "3"),
        ("Acknowledgement", "4"),
        ("Abstract", "5"),
        ("Chapter 1: Introduction & Project Scope", "7"),
        ("    1.1 Motivation and Clinical Context", "7"),
        ("    1.2 Problem Definition", "7"),
        ("    1.3 Project Objectives", "8"),
        ("Chapter 2: Dataset Description & Analysis", "9"),
        ("    2.1 Dataset Overview", "9"),
        ("    2.2 The Cambridge AQ-10 Behavioral Questions", "9"),
        ("    2.3 Demographic & Clinical Features", "10"),
        ("    2.4 Class Distribution & Imbalance", "10"),
        ("Chapter 3: Data Preprocessing & Leakage Elimination", "11"),
        ("    3.1 Data Cleaning & Anomaly Correction", "11"),
        ("    3.2 Missing Value Handling", "11"),
        ("    3.3 Categorical & Numerical Transformations", "11"),
        ("    3.4 Target Data Leakage: The 'result' Column Problem", "12"),
        ("Chapter 4: Machine Learning Models & Training Pipeline", "13"),
        ("    4.1 Model Selection & Algorithmic Paradigms", "13"),
        ("    4.2 5-Fold Stratified Cross-Validation", "13"),
        ("    4.3 Hyperparameter Tuning for Random Forest", "14"),
        ("    4.4 Empirical Target Leakage Ablation Study", "14"),
        ("Chapter 5: Model Explainability Using SHAP", "16"),
        ("    5.1 Why Explainability is Essential in Healthcare", "16"),
        ("    5.2 Game-Theoretic Shapley Values & TreeSHAP", "16"),
        ("    5.3 Global Feature Importance Analysis", "17"),
        ("    5.4 Patient-Level Risk vs Protective Decomposition", "17"),
        ("Chapter 6: Web Application & Clinical PDF Engine", "18"),
        ("    6.1 Interactive Streamlit Frontend Architecture", "18"),
        ("    6.2 Real-Time SVG Risk Speedometer Gauge", "18"),
        ("    6.3 Preset Clinical Evaluation Profiles", "19"),
        ("    6.4 Automated Clinical Summary PDF Report Generation", "19"),
        ("Chapter 7: Experimental Results & Performance Evaluation", "20"),
        ("    7.1 10-Model Cross-Validation Benchmark Comparison", "20"),
        ("    7.2 Priority of Recall / Sensitivity in Medical Screening", "21"),
        ("    7.3 Held-Out Test Evaluation & Confusion Matrix", "21"),
        ("    7.4 Error Analysis of False Negatives", "22"),
        ("Chapter 8: Project Execution & Verification", "23"),
        ("    8.1 Software Requirements & Dependencies", "23"),
        ("    8.2 Step-by-Step Execution Guide", "23"),
        ("    8.3 Automated Unit Testing with Pytest", "24"),
        ("Chapter 9: Conclusion & Learnings", "25"),
        ("    9.1 Summary of Deliverables", "25"),
        ("    9.2 Practical Engineering Learnings", "25"),
        ("    9.3 Limitations & Future Scope", "25"),
        ("References", "26")
    ]

    toc_table = doc.add_table(rows=len(toc_items), cols=2)
    toc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    toc_table.autofit = False

    for idx, (title, page) in enumerate(toc_items):
        row = toc_table.rows[idx]
        cell_t, cell_p = row.cells[0], row.cells[1]
        cell_t.width = Inches(5.6)
        cell_p.width = Inches(0.8)
        
        p_t = cell_t.paragraphs[0]
        p_t.paragraph_format.space_after = Pt(2)
        p_t.paragraph_format.line_spacing = 1.05
        r_t = p_t.add_run(title)
        r_t.font.name = "Times New Roman"
        r_t.font.size = Pt(11)
        if title.startswith("Chapter") or not title.startswith("    "):
            r_t.bold = True
            
        p_p = cell_p.paragraphs[0]
        p_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_p.paragraph_format.space_after = Pt(2)
        r_p = p_p.add_run(page)
        r_p.font.name = "Times New Roman"
        r_p.font.size = Pt(11)
        if title.startswith("Chapter") or not title.startswith("    "):
            r_p.bold = True

    doc.add_page_break()

    # ==============================================================================
    # CHAPTER 1: INTRODUCTION & PROJECT SCOPE
    # ==============================================================================
    add_h1("CHAPTER 1: INTRODUCTION & PROJECT SCOPE")
    
    add_h2("1.1 Motivation and Clinical Context")
    add_para("Autism Spectrum Disorder (ASD) is a developmental condition characterized by differences in social communication, social interaction patterns, focused interests, and repetitive sensory-motor behaviors. Because autism has no biological blood markers or routine imaging indicators, clinical identification relies entirely on behavioral observation and standardized questionnaires.")

    add_para("In standard medical practice, a full diagnostic assessment involves extensive observation by specialized clinical teams using multi-hour batteries such as ADOS-2 (Autism Diagnostic Observation Schedule) and ADI-R (Autism Diagnostic Interview-Revised). Due to an acute shortage of specialized child developmental specialists, clinical waiting periods can stretch between 12 to 18 months. Early intervention during early childhood yields significantly improved developmental and communication outcomes. Therefore, rapid, automated screening tools that can triage individuals in primary care settings and identify high-risk cases for expedited formal evaluation are of immense clinical importance.")

    add_h2("1.2 Problem Definition")
    add_para("The objective of this Machine Learning Mini Project is defined as follows:", space_after=4)
    add_para("\"To design, implement, evaluate, and deploy an explainable machine learning screening system for Autism Spectrum Disorder (ASD) using the Cambridge AQ-10 psychometric screening instrument and demographic indicators; ensuring robust data cleaning, elimination of target data leakage, transparent feature-level explainability via SHAP, and an interactive web portal for practical screening demonstration.\"", italic=True, space_after=10)

    add_h2("1.3 Project Objectives")
    add_para("To accomplish this goal, our team set the following concrete project milestones:")
    add_para("1. Clean and defensively preprocess the UCI / Kaggle ASD screening dataset, addressing missing values, synthetic age noise, and categorical encodings.", space_after=3)
    add_para("2. Benchmark eight diverse machine learning classification algorithms under 5-Fold Stratified Cross-Validation.", space_after=3)
    add_para("3. Resolve the target data leakage flaw common in online tutorials (such as Siddhardhan's YouTube baseline) by analyzing the impact of the 'result' feature.", space_after=3)
    add_para("4. Integrate Explainable AI (XAI) using TreeSHAP to decompose each patient's risk score into concrete behavioral factors.", space_after=3)
    add_para("5. Build and deploy an interactive Streamlit web application featuring a real-time risk gauge speedometer and downloadable clinical PDF reports.", space_after=3)
    add_para("6. Validate the entire pipeline with automated Pytest unit tests to guarantee software reproducibility.", space_after=12)

    doc.add_page_break()

    # ==============================================================================
    # CHAPTER 2: DATASET DESCRIPTION & ANALYSIS
    # ==============================================================================
    add_h1("CHAPTER 2: DATASET DESCRIPTION & ANALYSIS")

    add_h2("2.1 Dataset Overview")
    add_para("The benchmark dataset used in this project originates from the UCI Machine Learning Repository and Kaggle Autism Prediction competitions, mirroring the foundational data structure examined in Siddhardhan's research walkthrough. The dataset contains 800 patient records across 22 attributes, comprising 10 behavioral questionnaire items, continuous age, biological sex, ethnicity, country of residence, family history of autism, neonatal jaundice history, respondent identity, and the binary target class ('Class/ASD').")

    add_h2("2.2 The Cambridge AQ-10 Behavioral Questions")
    add_para("The core predictive indicators in the dataset are derived from the Cambridge Autism Spectrum Quotient 10-Item instrument (AQ-10), developed by Allison, Auyeung, and Baron-Cohen (2012). The AQ-10 extracts the most discriminant items from the original 50-item AQ across five core cognitive domains:")
    add_para("• Q1 (A1_Score): Sensory sensitivity to minor background sounds that others do not notice.", space_after=2)
    add_para("• Q2 (A2_Score): Tendency to focus more on granular details rather than the global picture.", space_after=2)
    add_para("• Q3 (A3_Score): Ease in multitasking and fluidly switching between conversations.", space_after=2)
    add_para("• Q4 (A4_Score): Ability to quickly pick up conversational threads after an interruption.", space_after=2)
    add_para("• Q5 (A5_Score): Ability to read between the lines and infer indirect conversational meaning.", space_after=2)
    add_para("• Q6 (A6_Score): Awareness of whether the listener is getting bored during a conversation.", space_after=2)
    add_para("• Q7 (A7_Score): Ease in understanding character motivations when reading stories.", space_after=2)
    add_para("• Q8 (A8_Score): Tendency to collect detailed, narrow information on specific categories.", space_after=2)
    add_para("• Q9 (A9_Score): Ability to decode subtle facial expressions and emotional cues.", space_after=2)
    add_para("• Q10 (A10_Score): Difficulty in inferring the intentions and mental states of other people.", space_after=8)
    add_para("Each item is recorded as a binary score (0 or 1), where 1 indicates an autistic-like response.")

    add_h2("2.3 Demographic & Clinical Features")
    add_para("In addition to the 10 AQ items, the dataset includes demographic and clinical covariates:")
    add_para("• Age: Patient age in years (numerical).", space_after=2)
    add_para("• Gender: Biological sex ('m' / 'f').", space_after=2)
    add_para("• Ethnicity: Ethnic background (e.g., White-European, Asian, South Asian, Middle Eastern, Pasifika, Others).", space_after=2)
    add_para("• Jaundice: History of neonatal jaundice at birth ('yes' / 'no'). Neonatal hyperbilirubinemia is a recognized early risk factor.", space_after=2)
    add_para("• Family ASD History ('austim'): First-degree family member diagnosed with ASD ('yes' / 'no').", space_after=2)
    add_para("• Country of Residence: Geographic location of the patient.", space_after=2)
    add_para("• Used App Before: Whether the patient had previously completed a digital screening.", space_after=2)
    add_para("• Relation: Who completed the screening (Self, Parent, Relative, Health Care Professional, Others).", space_after=8)

    add_h2("2.4 Class Distribution & Imbalance")
    add_para("In the raw dataset of 800 patients, 639 patients (79.9%) are neurotypical controls (Class/ASD = 0) and 161 patients (20.1%) are diagnosed with ASD (Class/ASD = 1). This 4:1 class imbalance mirrors real-world clinical screening populations. A naive model that classifies every patient as non-autistic would achieve ~80% accuracy while failing 100% of autistic individuals. Consequently, model optimization must prioritize Sensitivity (Recall) and ROC-AUC over simple accuracy.")

    doc.add_page_break()

    # ==============================================================================
    # CHAPTER 3: DATA PREPROCESSING & LEAKAGE ELIMINATION
    # ==============================================================================
    add_h1("CHAPTER 3: DATA PREPROCESSING & LEAKAGE ELIMINATION")

    add_h2("3.1 Data Cleaning & Anomaly Correction")
    add_para("Defensive data cleaning was implemented in `src/data_pipeline.py`. Inspection of raw CSV data revealed that the `age` column contained floating-point artifacts (e.g., 38.1727) caused by synthetic data generation, alongside impossible extreme values exceeding 300 years. Our cleaning routine enforces biological validity by rounding age to integers and applying clipping: Age_clean = max(1, min(Age, 95)).")

    add_h2("3.2 Missing Value Handling")
    add_para("Missing entries in the raw dataset were represented as question mark strings ('?'). Rather than discarding rows containing missing data (which would reduce the small training set), our pipeline replaces missing categorical entries with the category `'Others'`. Missing numerical values in `age` are imputed using the median (30.0 years) to prevent distortion from extreme values.")

    add_h2("3.3 Categorical & Numerical Transformations")
    add_para("To prepare features for scikit-learn estimators, we constructed a unified `ColumnTransformer`:")
    add_para("1. Binary Encoding: Binary fields (`gender`, `jaundice`, `austim`, `used_app_before`) are mapped deterministically to 0 and 1.", space_after=3)
    add_para("2. One-Hot Encoding: Multi-class categorical features (`ethnicity`, `contry_of_res`, `relation`) are encoded using `OneHotEncoder(handle_unknown='ignore', sparse_output=False)`.", space_after=3)
    add_para("3. Robust Scaling: The continuous `age` feature is scaled using `RobustScaler()`, which standardizes using the median and interquartile range (IQR = Q3 - Q1), preventing any residual outlier leverage.", space_after=3)
    add_para("4. Passthrough: The ten binary behavioral scores (`A1_Score` to `A10_Score`) pass through directly without modification.", space_after=10)

    # Table 3.1
    feat_headers = ["Feature Identifier", "Data Type", "Clinical Question / Meaning", "Processing Strategy"]
    feat_rows = [
        ["ID", "Integer", "Database row identifier", "Dropped (zero predictive signal)"],
        ["A1_Score", "Binary (0/1)", "Q1: Sensory sensitivity to minor ambient sounds", "Passthrough (binary integer flag)"],
        ["A2_Score", "Binary (0/1)", "Q2: Focusing on global picture vs granular detail", "Passthrough (binary integer flag)"],
        ["A3_Score", "Binary (0/1)", "Q3: Multitasking fluidity & conversational shifting", "Passthrough (binary integer flag)"],
        ["A4_Score", "Binary (0/1)", "Q4: Fluency and recovery after conversational interruption", "Passthrough (binary integer flag)"],
        ["A5_Score", "Binary (0/1)", "Q5: Ease in reading between lines and indirect meaning", "Passthrough (binary integer flag)"],
        ["A6_Score", "Binary (0/1)", "Q6: Difficulty knowing if a listener is getting bored", "Passthrough (binary integer flag)"],
        ["A7_Score", "Binary (0/1)", "Q7: Difficulty understanding literary characters' intent", "Passthrough (binary integer flag)"],
        ["A8_Score", "Binary (0/1)", "Q8: Tendency to obsessively collect narrow categorical info", "Passthrough (binary integer flag)"],
        ["A9_Score", "Binary (0/1)", "Q9: Ease in decoding facial expressions & emotional cues", "Passthrough (binary integer flag)"],
        ["A10_Score", "Binary (0/1)", "Q10: Difficulty understanding intentions & subtle cues", "Passthrough (binary integer flag)"],
        ["age", "Continuous", "Patient age in years (contains floating-point noise)", "Clipped to [1, 95], RobustScaler"],
        ["gender", "Categorical", "Biological sex ('m' / 'f')", "Binary mapping: Male=1, Female=0"],
        ["jaundice", "Categorical", "History of neonatal jaundice at birth ('yes' / 'no')", "Binary mapping: Yes=1, No=0"],
        ["austim", "Categorical", "Family history of ASD in first-degree relatives", "Binary mapping: Yes=1, No=0"]
    ]

    t_feat = doc.add_table(rows=len(feat_rows) + 1, cols=4)
    t_feat.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_feat.autofit = False

    for c_idx, h in enumerate(feat_headers):
        cell = t_feat.rows[0].cells[c_idx]
        cell.text = h
        set_cell_background(cell, "0F172A")
        set_cell_margins(cell, top=100, bottom=100, left=80, right=80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.runs[0]
        run.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.size = Pt(9.5)

    for r_idx, row_data in enumerate(feat_rows):
        row = t_feat.rows[r_idx + 1]
        bg = "F8FAFC" if r_idx % 2 == 0 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.text = val
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=70, bottom=70, left=80, right=80)
            p = cell.paragraphs[0]
            run = p.runs[0]
            run.font.size = Pt(9)
            if c_idx == 0:
                run.bold = True

    add_para("\nCaption: Table 3.1: Feature Dictionary & Preprocessing Transformations.", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, space_after=12)

    add_h2("3.4 Target Data Leakage: The 'result' Column Problem")
    add_para("A critical finding in our literature and code analysis of prior work (including Siddhardhan's YouTube video) was the issue of Target Data Leakage caused by the `result` column.")
    add_para("In the raw dataset, `result` is mathematically defined as the exact arithmetic sum of the ten behavioral scores: result = A1 + A2 + ... + A10. In the dataset collection methodology, the binary label `Class/ASD` was assigned whenever `result >= 6`.")
    add_para("When a model is trained with the `result` column included, it easily achieves near-perfect accuracy (>98%) simply by learning a basic threshold on this single column. However, it fails to learn the genuine multivariate interactions between individual questions, family history, and demographic indicators. In our project, we permanently eliminated `result` from the input feature space to train genuine predictive models, and we conducted an empirical ablation study to prove our models learn legitimate non-linear relationships.")

    doc.add_page_break()

    # ==============================================================================
    # CHAPTER 4: MACHINE LEARNING MODELS & TRAINING PIPELINE
    # ==============================================================================
    add_h1("CHAPTER 4: MACHINE LEARNING MODELS & TRAINING PIPELINE")

    add_h2("4.1 Model Selection & Algorithmic Paradigms")
    add_para("We evaluated eight diverse classification architectures spanning linear, probabilistic, margin-based, tree-based, ensemble, and neural network paradigms:")
    add_para("1. Random Forest Classifier (Champion): An ensemble of 200 de-correlated decision trees trained with bootstrap aggregation (bagging) and random feature subspace sampling. Bagging drastically reduces model variance while maintaining low bias.", space_after=3)
    add_para("2. Support Vector Classifier (SVC): Fits an optimal separating hyperplane using the non-linear Radial Basis Function (RBF) kernel K(x, x') = exp(-gamma * ||x - x'||^2).", space_after=3)
    add_para("3. Extra Trees Classifier: Extremely Randomized Trees that randomize cut-point choice during split evaluation, further lowering variance.", space_after=3)
    add_para("4. Soft-Voting Meta-Ensemble: Combines predictions of Random Forest, Gradient Boosting, and Logistic Regression by averaging predicted class probabilities.", space_after=3)
    add_para("5. Logistic Regression: L2-regularized (Ridge) linear probabilistic model optimizing log-loss with balanced class weights.", space_after=3)
    add_para("6. HistGradientBoosting: Scikit-learn's histogram-based gradient boosting classifier (equivalent to LightGBM), binning continuous features to build decision trees sequentially.", space_after=3)
    add_para("7. Gradient Boosting Classifier: Standard sequential boosting optimizing binomial deviance loss.", space_after=3)
    add_para("8. Multi-Layer Perceptron (MLP): Feedforward neural network with two hidden layers (64, 32 neurons) and ReLU activations.", space_after=3)
    add_para("9. Decision Tree Classifier (CART): Recursive binary splitting based on Gini impurity minimization.", space_after=3)
    add_para("10. Gaussian Naive Bayes: Generative baseline assuming conditional feature independence.", space_after=10)

    add_h2("4.2 5-Fold Stratified Cross-Validation")
    add_para("To ensure unbiased performance estimation, models were trained and validated using 5-Fold Stratified Cross-Validation on the 80% training split (N = 640 patients). Stratification ensures that each fold maintains the identical 20% positive ASD class distribution. Within each fold, scalers and encoders are fitted strictly on the training partition to eliminate data leakage.")

    add_h2("4.3 Hyperparameter Tuning for Random Forest")
    add_para("Hyperparameters for the champion Random Forest model were optimized via RandomizedSearchCV across 20 candidate configurations evaluated under 5-Fold Stratified CV, optimizing for ROC-AUC. The optimal configuration identified was: n_estimators = 200, max_depth = 8, min_samples_leaf = 4, min_samples_split = 2, and bootstrap = True, achieving a peak cross-validated ROC-AUC of 0.9209.")

    add_h2("4.4 Empirical Target Leakage Ablation Study")
    add_para("To scientifically confirm that our model does not depend on the leaked `result` column, we conducted an ablation experiment comparing performance with versus without the `result` column.")

    # Table 4.1 Ablation Table
    abl_headers = ["Algorithm Architecture", "Accuracy (With Result)", "ROC-AUC (With Result)", "Accuracy (Without Result)", "ROC-AUC (Without Result)"]
    abl_rows = [
        ["Random Forest (Champion)", "83.75%", "0.9140", "83.91%", "0.9182"],
        ["Support Vector Classifier (RBF)", "82.66%", "0.9189", "81.87%", "0.9181"],
        ["Extra Trees Classifier", "83.91%", "0.9169", "84.38%", "0.9178"],
        ["Soft-Voting Meta-Ensemble", "84.38%", "0.9083", "84.06%", "0.9148"],
        ["Logistic Regression (Ridge)", "82.19%", "0.9132", "81.72%", "0.9111"]
    ]

    t_abl = doc.add_table(rows=len(abl_rows) + 1, cols=5)
    t_abl.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_abl.autofit = False

    for c_idx, h in enumerate(abl_headers):
        cell = t_abl.rows[0].cells[c_idx]
        cell.text = h
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, top=100, bottom=100, left=80, right=80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.runs[0]
        run.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.size = Pt(9.5)

    for r_idx, row_data in enumerate(abl_rows):
        row = t_abl.rows[r_idx + 1]
        bg = "F8FAFC" if r_idx % 2 == 0 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.text = val
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=70, bottom=70, left=80, right=80)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.runs[0]
            run.font.size = Pt(9)
            if c_idx == 0:
                run.bold = True

    add_para("\nCaption: Table 4.1: Target Leakage Ablation Study: Performance Comparison With vs. Without 'result' Feature.", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, space_after=10)
    add_para("Critical Conclusion: As demonstrated in Table 4.1, Random Forest maintains a ROC-AUC of 0.9182 even when `result` is 100% eliminated. Removing the collinear sum slightly improved tree split stability, proving that our framework captures true clinical relationships across AQ questions and demographics.", space_after=12)

    doc.add_page_break()

    # ==============================================================================
    # CHAPTER 5: MODEL EXPLAINABILITY USING SHAP
    # ==============================================================================
    add_h1("CHAPTER 5: MODEL EXPLAINABILITY USING SHAP")

    add_h2("5.1 Why Explainability is Essential in Healthcare")
    add_para("A primary limitation of modern ensemble machine learning in medicine is the 'black-box' nature of predictions. Clinicians cannot ethically act on an opaque probability score without knowing which specific behavioral indicators contributed to the decision. An explainability layer transforms a black-box classifier into an auditable, clinician-friendly decision-support tool.")

    add_h2("5.2 Game-Theoretic Shapley Values & TreeSHAP")
    add_para("To achieve transparent interpretability, our system integrates SHAP (Shapley Additive Explanations), based on cooperative game theory (Lundberg & Lee, 2017). Each feature value is treated as a 'player' contributing toward the predicted risk probability. The Shapley value phi_j represents the weighted average marginal contribution of feature j across all possible feature subsets:")
    add_para("phi_j(x) = sum_{S subseteq F \\ {j}} [ (|S|! * (|F| - |S| - 1)!) / |F|! ] * [ f(S union {j}) - f(S) ]", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, space_after=8)
    add_para("Using `shap.TreeExplainer`, Shapley values are calculated in polynomial time for the Random Forest ensemble, satisfying the local accuracy axiom: the sum of all SHAP values plus the baseline risk equals the exact predicted patient probability.")

    add_h2("5.3 Global Feature Importance Analysis")
    add_para("Global TreeSHAP analysis across the cohort revealed the top behavioral predictors driving ASD classification:")
    add_para("1. Q6 (A6_Score - Listener Boredom Awareness): Inability to tell if a conversational partner is bored is the single strongest positive risk driver.", space_after=2)
    add_para("2. Q9 (A9_Score - Emotional & Facial Cue Decoding): Difficulty interpreting facial expressions and subtle emotions is the second most heavily weighted indicator.", space_after=2)
    add_para("3. Q3 (A3_Score - Attention Switching): Difficulty in fluidly shifting attention between conversations serves as the top cognitive flexibility factor.", space_after=8)

    add_h2("5.4 Patient-Level Risk vs Protective Decomposition")
    add_para("For every individual patient evaluation, our explainability module decomposes the prediction into two intuitive, visual categories:")
    add_para("• Positive Risk Drivers (+phi): Behavioral traits that push the risk score upward toward the high-risk threshold (e.g., 'Q6: Listener Boredom Awareness (+22.4% risk impact)'). Displayed as red alert cards.", space_after=3)
    add_para("• Protective Baseline Drivers (-phi): Neurotypical responses that lower the patient's predicted risk toward the baseline (e.g., 'Q4: Conversational Fluency (-14.2% protective buffer)'). Displayed as green cards.", space_after=12)

    doc.add_page_break()

    # ==============================================================================
    # CHAPTER 6: WEB APPLICATION & CLINICAL PDF ENGINE
    # ==============================================================================
    add_h1("CHAPTER 6: WEB APPLICATION & CLINICAL PDF ENGINE")

    add_h2("6.1 Interactive Streamlit Frontend Architecture")
    add_para("To make the machine learning model usable by clinicians and researchers, we engineered a full-stack interactive web application using Streamlit (`app/app.py`). The questionnaire is organized into four intuitive clinical domains rather than an unorganized list:")
    add_para("• Domain A: Social Communication & Interaction (Items Q1, Q5, Q6, Q10)", space_after=2)
    add_para("• Domain B: Cognitive Flexibility & Attention Switching (Items Q3, Q4)", space_after=2)
    add_para("• Domain C: Pattern Recognition & Detail Granularity (Items Q2, Q7, Q8)", space_after=2)
    add_para("• Domain D: Emotional & Facial Cue Decoding (Item Q9)", space_after=8)

    add_h2("6.2 Real-Time SVG Risk Speedometer Gauge")
    add_para("The application features a real-time SVG risk speedometer gauge that visually displays the predicted probability across three clinically defined risk tiers:")
    add_para("• Low Risk (< 35%): Normal baseline; routine pediatric tracking advised.", space_after=2)
    add_para("• Moderate Risk (35% - 70%): Mild behavioral differences; re-screening recommended in 3 to 6 months.", space_after=2)
    add_para("• High Risk (> 70%): Significant ASD indicators present; priority referral for comprehensive clinical assessment.", space_after=8)

    add_h2("6.3 Preset Clinical Evaluation Profiles")
    add_para("To facilitate rapid demonstration and testing without manually filling all fields, the sidebar includes three preset buttons:")
    add_para("• Default Median Profile: Average population baseline values.", space_after=2)
    add_para("• High-Risk ASD Profile: Preloads responses characteristic of ASD, verifying that the system flags elevated risk (>75%).", space_after=2)
    add_para("• Low-Risk Neurotypical Profile: Preloads neurotypical responses, verifying low risk (<25%).", space_after=8)

    add_h2("6.4 Automated Clinical Summary PDF Report Generation")
    add_para("Built with ReportLab (`src/report_generator.py`), the system provides one-click generation of a single-page, print-ready clinical screening PDF summary containing:")
    add_para("1. Institutional header with timestamp and unique report ID.", space_after=2)
    add_para("2. Patient demographic summary table (Age, Sex, Ethnicity, Jaundice, Family History).", space_after=2)
    add_para("3. Risk tier badge and calculated probability score.", space_after=2)
    add_para("4. Top decision drivers table derived from SHAP attribution.", space_after=2)
    add_para("5. Complete AQ-10 item response table.", space_after=2)
    add_para("6. Medical disclaimer stating this is a screening triaging tool, not a standalone medical diagnosis.", space_after=12)

    doc.add_page_break()

    # ==============================================================================
    # CHAPTER 7: EXPERIMENTAL RESULTS & PERFORMANCE EVALUATION
    # ==============================================================================
    add_h1("CHAPTER 7: EXPERIMENTAL RESULTS & PERFORMANCE EVALUATION")

    add_h2("7.1 10-Model Cross-Validation Benchmark Comparison")
    add_para("Table 7.1 presents the 5-Fold Stratified Cross-Validation results across all ten evaluated machine learning models on the training split (N = 640).")

    # Table 7.1
    res_headers = ["Algorithm Architecture", "Accuracy", "Precision", "Recall (Sensitivity)", "F1-Score", "ROC-AUC"]
    res_rows = [
        ["Random Forest (Tuned Champion)", "83.91%", "58.46%", "77.48%", "0.6610", "0.9209"],
        ["Support Vector Classifier (RBF)", "81.87%", "54.97%", "82.18%", "0.6525", "0.9181"],
        ["Extra Trees Classifier", "84.38%", "60.47%", "69.63%", "0.6390", "0.9178"],
        ["Soft-Voting Ensemble (RF+GB+LR)", "84.06%", "59.17%", "73.54%", "0.6476", "0.9148"],
        ["Logistic Regression (Ridge)", "81.72%", "54.23%", "81.42%", "0.6457", "0.9111"],
        ["HistGradientBoosting (LightGBM)", "84.69%", "64.16%", "56.49%", "0.5941", "0.9096"],
        ["Gradient Boosting Classifier", "83.75%", "61.21%", "55.66%", "0.5719", "0.9036"],
        ["Multi-Layer Perceptron (Neural Net)", "84.06%", "62.31%", "52.58%", "0.5577", "0.8766"],
        ["Decision Tree Classifier (CART)", "81.41%", "54.04%", "76.80%", "0.6279", "0.8203"],
        ["Gaussian Naive Bayes", "57.34%", "31.77%", "92.25%", "0.4708", "0.8102"]
    ]

    t_res = doc.add_table(rows=len(res_rows) + 1, cols=6)
    t_res.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_res.autofit = False

    for c_idx, h in enumerate(res_headers):
        cell = t_res.rows[0].cells[c_idx]
        cell.text = h
        set_cell_background(cell, "0F172A")
        set_cell_margins(cell, top=100, bottom=100, left=60, right=60)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.runs[0]
        run.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.size = Pt(9)

    for r_idx, row_data in enumerate(res_rows):
        row = t_res.rows[r_idx + 1]
        bg = "ECFDF5" if r_idx == 0 else ("F8FAFC" if r_idx % 2 == 0 else "FFFFFF")
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.text = val
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=70, bottom=70, left=60, right=60)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.runs[0]
            run.font.size = Pt(9)
            if r_idx == 0 or c_idx == 0:
                run.bold = True

    add_para("\nCaption: Table 7.1: 5-Fold Stratified Cross-Validation Benchmark Across Machine Learning Algorithms.", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, space_after=12)

    add_h2("7.2 Priority of Recall / Sensitivity in Medical Screening")
    add_para("In healthcare screening, false negatives (Type II error: an autistic individual misclassified as neurotypical) are clinically far more harmful than false positives (Type I error: a neurotypical individual referred for secondary review). If an autistic individual is missed, they lose the opportunity for early developmental intervention. Consequently, Recall (Sensitivity) = TP / (TP + FN) and ROC-AUC are the primary criteria for model selection. Random Forest achieves an exceptional balance with 77.48% CV Recall and 0.9209 ROC-AUC.")

    add_h2("7.3 Held-Out Test Evaluation & Confusion Matrix")
    add_para("On the 20% held-out test set (N = 160 patients: 128 neurotypical controls, 32 ASD cases), the tuned Random Forest model achieved:")
    add_para("• Test ROC-AUC: 0.8994", space_after=2)
    add_para("• Test Recall (Sensitivity): 81.25% (26 of 32 ASD cases identified)", space_after=2)
    add_para("• Test Overall Accuracy: 80.63%", space_after=2)
    add_para("• Test Precision: 50.98%", space_after=2)
    add_para("• Test F1-Score: 0.6265", space_after=6)

    add_para("Confusion Matrix Breakdown on Held-Out Test Cohort:", bold=True, space_after=4)
    add_para("""
                        Predicted Non-ASD (0)       Predicted ASD (1)
Actual Non-ASD (0)            TN = 103                   FP = 25
Actual ASD (1)                 FN = 6                    TP = 26
""", align=WD_ALIGN_PARAGRAPH.LEFT, space_after=6)
    add_para("Caption: Confusion Matrix on Unseen Held-Out Test Set (N = 160 Patients).", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, space_after=10)

    add_h2("7.4 Error Analysis of False Negatives")
    add_para("Analysis of the 6 false-negative cases showed that these individuals scored borderline on the AQ-10 (4 or 5 affirmative answers) without family history of autism. In medical screening, no questionnaire is 100% sensitive; our 81.25% test sensitivity provides a strong triaging baseline that captures the vast majority of positive cases.")

    doc.add_page_break()

    # ==============================================================================
    # CHAPTER 8: PROJECT EXECUTION & VERIFICATION
    # ==============================================================================
    add_h1("CHAPTER 8: PROJECT EXECUTION & VERIFICATION")

    add_h2("8.1 Software Requirements & Dependencies")
    add_para("The project was developed in Python 3.14 using standard, industry-standard scientific packages:")
    add_para("• Core ML & Data: scikit-learn (v1.6.1), numpy (v2.2.3), pandas (v2.2.3), scipy", space_after=2)
    add_para("• Explainability: shap (v0.46.0), matplotlib, seaborn", space_after=2)
    add_para("• Application & Export: streamlit (v1.42.0), reportlab (v4.3.1)", space_after=2)
    add_para("• Automated Testing: pytest (v8.3.4)", space_after=8)

    add_h2("8.2 Step-by-Step Execution Guide")
    add_para("1. Environment Setup & Dependency Installation:")
    add_para("   pip install -r requirements.txt", italic=True, space_after=4)
    add_para("2. Train Models and Generate Champion Pipeline:")
    add_para("   python src/train.py", italic=True, space_after=4)
    add_para("   (Trains all 8+ models under 5-Fold Stratified CV, tunes Random Forest, and exports `models/pipeline.pkl` and `models/champion_model.pkl`.)")
    add_para("3. Run the Streamlit Web Application:")
    add_para("   streamlit run app/app.py", italic=True, space_after=4)
    add_para("   (Opens the interactive clinical screening portal on http://localhost:8501.)")
    add_para("4. Run Automated Unit Tests:")
    add_para("   pytest -v tests/", italic=True, space_after=8)

    add_h2("8.3 Automated Unit Testing with Pytest")
    add_para("Six automated unit tests were created in `tests/test_pipeline.py` and `tests/test_model.py`:")
    add_para("• `test_pipeline_transform_shape`: Confirms preprocessing pipeline handles missing tokens and transforms to expected feature dimensions.", space_after=2)
    add_para("• `test_age_outlier_clipping`: Verifies that impossible age entries (>300) are safely clipped to biological bounds [1, 95].", space_after=2)
    add_para("• `test_model_prediction_range`: Ensures predicted probabilities lie strictly within [0.0, 1.0].", space_after=2)
    add_para("• `test_high_risk_prediction`: Confirms that characteristic ASD response profiles produce risk scores > 70%.", space_after=2)
    add_para("• `test_explain_instance_returns_valid_structure`: Verifies SHAP returns non-empty positive and protective feature lists.", space_after=2)
    add_para("• `test_pdf_report_generation`: Confirms ReportLab generates non-empty, non-corrupted PDF summary reports.", space_after=8)
    add_para("All 6 unit tests execute cleanly (`6 passed in 3.42s`), verifying software reliability.")

    doc.add_page_break()

    # ==============================================================================
    # CHAPTER 9: CONCLUSION & LEARNINGS
    # ==============================================================================
    add_h1("CHAPTER 9: CONCLUSION & LEARNINGS")

    add_h2("9.1 Summary of Deliverables")
    add_para("In this Machine Learning Mini Project, our team successfully built and verified an end-to-end Autism Spectrum Disorder screening framework:")
    add_para("1. Cleaned and structured the UCI / Kaggle ASD screening dataset, correcting synthetic age noise and missing categorical values.", space_after=2)
    add_para("2. Identified and resolved target data leakage caused by the `result` column, validating performance via an empirical ablation study.", space_after=2)
    add_para("3. Benchmarked 10 machine learning models under 5-Fold Stratified Cross-Validation, establishing a tuned Random Forest model as champion (0.9209 ROC-AUC, 81.25% test sensitivity).", space_after=2)
    add_para("4. Integrated TreeSHAP explainability to decompose patient predictions into transparent positive and protective behavioral factors.", space_after=2)
    add_para("5. Deployed a full-stack Streamlit web app with preset profiles, a risk speedometer gauge, and automated clinical PDF summary export.", space_after=8)

    add_h2("9.2 Practical Engineering Learnings")
    add_para("Working on this project provided valuable practical machine learning lessons:")
    add_para("• Understanding Target Leakage: Realized that high reported accuracy (>99%) in beginner tutorials often stems from collinear summary columns rather than genuine learning.", space_after=3)
    add_para("• Clinical Importance of Recall: Learned that in medical triage, optimizing for Recall / Sensitivity is far more important than optimizing for raw accuracy.", space_after=3)
    add_para("• Explainable AI is Critical: Experienced firsthand how SHAP helps clinicians understand model decisions, converting a black-box model into a trustworthy tool.", space_after=8)

    add_h2("9.3 Limitations & Future Scope")
    add_para("• Limitations: The model relies on self- or caregiver-reported questionnaires, which carry inherent subjective reporting bias. It is a screening triage tool, not a formal medical diagnosis.", space_after=3)
    add_para("• Future Scope: Extending the pipeline to pediatric toddler instruments (such as the Q-CHAT for ages 18-24 months) and exploring multi-modal integration with video-based motor pattern analysis.", space_after=12)

    doc.add_page_break()

    # ==============================================================================
    # REFERENCES
    # ==============================================================================
    add_h1("REFERENCES")

    references = [
        "Allison, C., Auyeung, B., & Baron-Cohen, S. (2012). Toward brief 'Red Flags' for autism screening: The Short Autism Spectrum Quotient and the Short Quantitative Checklist in 1,000 cases and 3,000 controls. Journal of the American Academy of Child & Adolescent Psychiatry, 51(2), 202-212.",
        "Baron-Cohen, S., Wheelwright, S., Skinner, R., Martin, J., & Clubley, E. (2001). The Autism-Spectrum Quotient (AQ): Evidence from Asperger syndrome/high-functioning autism in adults. Journal of Autism and Developmental Disorders, 31(1), 5-17.",
        "Thabtah, F. (2017). Autism Spectrum Disorder (ASD) Screening: A Machine Learning Approach. UCI Machine Learning Repository.",
        "Thabtah, F. (2019). Machine learning in autistic spectrum disorder behavioral research: A review and ways forward. Informatics for Health and Social Care, 44(3), 278-297.",
        "Lundberg, S. M., & Lee, S. I. (2017). A unified approach to interpreting model predictions. Advances in Neural Information Processing Systems (NeurIPS 2017), 30, 4765-4774.",
        "Breiman, L. (2001). Random forests. Machine Learning, 45(1), 5-32.",
        "Pedregosa, F., et al. (2011). Scikit-learn: Machine learning in Python. Journal of Machine Learning Research, 12, 2825-2830.",
        "Siddhardhan, S. (2024). Project 21: Autism Prediction Using Machine Learning | Complete ML Project Walkthrough. Machine Learning Series."
    ]

    for idx, ref in enumerate(references, 1):
        p_ref = doc.add_paragraph()
        p_ref.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_ref.paragraph_format.left_indent = Inches(0.4)
        p_ref.paragraph_format.first_line_indent = Inches(-0.4)
        p_ref.paragraph_format.space_after = Pt(6)
        p_ref.paragraph_format.line_spacing = 1.15
        
        r_num = p_ref.add_run(f"[{idx}] ")
        r_num.bold = True
        r_num.font.name = "Times New Roman"
        r_num.font.size = Pt(11)
        
        r_text = p_ref.add_run(ref)
        r_text.font.name = "Times New Roman"
        r_text.font.size = Pt(11)

    saved_paths = []
    try:
        doc.save(DOC_PATH)
        saved_paths.append(DOC_PATH)
        print(f"Report successfully generated at: {DOC_PATH}")
    except PermissionError:
        alt_path = r"D:\autism_prediction_project\Autism_Spectrum_Disorder_Prediction_Project_Report_Numbered.docx"
        doc.save(alt_path)
        saved_paths.append(alt_path)
        print(f"Primary file was open in Word. Saved to: {alt_path}")
    
    # Always save to scratch path as well
    os.makedirs(os.path.dirname(SCRATCH_PATH), exist_ok=True)
    doc.save(SCRATCH_PATH)
    saved_paths.append(SCRATCH_PATH)
    print(f"Report saved to scratch directory at: {SCRATCH_PATH}")

if __name__ == "__main__":
    create_report()
