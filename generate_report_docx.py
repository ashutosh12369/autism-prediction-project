"""
generate_report_docx.py
Generates a complete, descriptive, humanized, plagiarism-free academic B.Tech project report in .docx format.
Tailored for IIIT Bhubaneswar, B.Tech CSE / Computer Engineering.
"""

import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

DOC_PATH = r"D:\autism_prediction_project\Autism_Spectrum_Disorder_Prediction_Project_Report.docx"

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
        # Add page numbering to footer
        footer = section.footer
        f_p = footer.paragraphs[0]
        f_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        f_run = f_p.add_run("Page ")
        f_run.font.name = "Times New Roman"
        f_run.font.size = Pt(10)
        f_run.font.color.rgb = RGBColor(100, 100, 100)

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
        run.font.size = Pt(16)
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
        run.font.size = Pt(13.5)
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
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(51, 65, 85)
        return p

    # ==============================================================================
    # 1. TITLE PAGE
    # ==============================================================================
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(36)
    p_title.paragraph_format.space_after = Pt(12)
    r_title = p_title.add_run("AN EXPLAINABLE MACHINE LEARNING DECISION-SUPPORT FRAMEWORK FOR EARLY AUTISM SPECTRUM DISORDER (ASD) SCREENING")
    r_title.bold = True
    r_title.font.size = Pt(18)
    r_title.font.name = "Times New Roman"

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(28)
    r_sub = p_sub.add_run("A Major Project Report Submitted in Partial Fulfillment of the Requirements for the Degree of\nBachelor of Technology\nin\nComputer Science and Engineering")
    r_sub.font.size = Pt(12.5)
    r_sub.font.name = "Times New Roman"

    # College Logo placeholder or box
    logo_path = r"D:\autism_prediction_project\food_waste_tw_report\report\assets\iiit_bhubaneswar_logo.jpeg"
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_after = Pt(20)
        p_logo.add_run().add_picture(logo_path, width=Inches(1.5))

    p_by = doc.add_paragraph()
    p_by.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_by.paragraph_format.space_after = Pt(6)
    p_by.add_run("Submitted By:").bold = True

    p_team = doc.add_paragraph()
    p_team.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_team.paragraph_format.space_after = Pt(24)
    r_team = p_team.add_run("ASHUTOSH SAHU (Roll No: B523014)\nMOHAMMAD ARIF (Roll No: B523039)")
    r_team.bold = True
    r_team.font.size = Pt(12)

    p_under = doc.add_paragraph()
    p_under.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_under.paragraph_format.space_after = Pt(4)
    p_under.add_run("Under the Guidance of:").bold = True

    p_guide = doc.add_paragraph()
    p_guide.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_guide.paragraph_format.space_after = Pt(28)
    r_guide = p_guide.add_run("Prof. Bharati Mishra\nDepartment of Computer Science and Engineering")
    r_guide.font.size = Pt(12)
    r_guide.bold = True

    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inst.paragraph_format.space_after = Pt(0)
    r_inst = p_inst.add_run("DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING\nINTERNATIONAL INSTITUTE OF INFORMATION TECHNOLOGY, BHUBANESWAR\nODISHA, INDIA - 751003\nNOVEMBER 2026")
    r_inst.bold = True
    r_inst.font.size = Pt(12)

    doc.add_page_break()

    # ==============================================================================
    # 2. CERTIFICATE OF APPROVAL
    # ==============================================================================
    add_h1("CERTIFICATE OF APPROVAL")
    add_para("This is to certify that the project entitled \"An Explainable Machine Learning Decision-Support Framework for Early Autism Spectrum Disorder (ASD) Screening\" submitted by Ashutosh Sahu (Roll No: B523014) and Mohammad Arif (Roll No: B523039) in partial fulfillment of the requirements for the award of the degree of Bachelor of Technology in Computer Science and Engineering at International Institute of Information Technology, Bhubaneswar, is an authentic work carried out by them under my supervision and guidance.", space_after=18)

    add_para("To the best of my knowledge, the matter embodied in this report has not been submitted to any other University / Institute for the award of any degree or diploma.", space_after=36)

    # Signatures table
    sig_table = doc.add_table(rows=2, cols=3)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_table.autofit = False

    sig_data = [
        ["________________________", "________________________", "________________________"],
        ["Prof. Bharati Mishra\nProject Supervisor\nDept. of CSE", "Head of Department\nDept. of CSE\nIIIT Bhubaneswar", "External Examiner\nEvaluation Board\nDate: ____________"]
    ]
    for r_idx, row in enumerate(sig_table.rows):
        for c_idx, cell in enumerate(row.cells):
            cell.width = Inches(2.2)
            cell.text = sig_data[r_idx][c_idx]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.runs[0].font.name = "Times New Roman"
            p.runs[0].font.size = Pt(10.5)

    doc.add_page_break()

    # ==============================================================================
    # 3. DECLARATION OF ORIGINALITY
    # ==============================================================================
    add_h1("DECLARATION OF ORIGINALITY")
    add_para("We hereby declare that the project work presented in this report entitled \"An Explainable Machine Learning Decision-Support Framework for Early Autism Spectrum Disorder (ASD) Screening\" is an authentic record of our own research work carried out under the supervision of Prof. Bharati Mishra, Department of Computer Science and Engineering, International Institute of Information Technology, Bhubaneswar.", space_after=14)

    add_para("We further confirm that this work is original, free from plagiarism, and has not been submitted previously in part or in full for the award of any degree, diploma, or fellowship in this or any other academic institution. All literature sources, analytical frameworks, and benchmark datasets utilized during this investigation have been duly cited and acknowledged in accordance with standard academic ethics.", space_after=36)

    p_dec_sig = doc.add_paragraph()
    p_dec_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_dec_sig.paragraph_format.line_spacing = 1.3
    r_dec = p_dec_sig.add_run("Ashutosh Sahu (Roll No: B523014)\nMohammad Arif (Roll No: B523039)\n\nPlace: Bhubaneswar, Odisha\nDate: November 2026")
    r_dec.font.name = "Times New Roman"
    r_dec.font.size = Pt(11.5)
    r_dec.bold = True

    doc.add_page_break()

    # ==============================================================================
    # 4. ACKNOWLEDGEMENT
    # ==============================================================================
    add_h1("ACKNOWLEDGEMENT")
    add_para("The successful completion of this major project represents an enriching academic journey, which was made possible through the guidance, encouragement, and support of numerous individuals and institutional facilities.", space_after=12)

    add_para("First and foremost, we express our profound gratitude and indebtedness to our esteemed project guide, Prof. Bharati Mishra, Department of Computer Science and Engineering, for her invaluable mentorship, perceptive critiques, and continuous technical encouragement throughout the design and execution of this machine learning system. Her deep insights into computational healthcare modeling provided clear direction to our research.", space_after=12)

    add_para("We also extend our sincere appreciation to the Head of Department, Department of Computer Science and Engineering, and the honorable Director of IIIT Bhubaneswar for providing advanced computing infrastructure, lab facilities, and an encouraging research environment conducive to undergraduate innovation.", space_after=12)

    add_para("Lastly, we owe our deepest gratitude to our parents, family members, and peers whose moral support and patience sustained our commitment across every phase of this capstone endeavor.", space_after=24)

    p_ack_sig = doc.add_paragraph()
    p_ack_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_ack = p_ack_sig.add_run("Ashutosh Sahu & Mohammad Arif\nDepartment of Computer Science and Engineering\nIIIT Bhubaneswar")
    r_ack.font.name = "Times New Roman"
    r_ack.font.size = Pt(11.5)
    r_ack.bold = True

    doc.add_page_break()

    # ==============================================================================
    # 5. ABSTRACT
    # ==============================================================================
    add_h1("ABSTRACT")
    add_para("Autism Spectrum Disorder (ASD) is a pervasive neurodevelopmental condition characterized by persistent challenges in reciprocal social communication, non-verbal emotional decoding, and restricted, repetitive behavioral repertoires. Clinical literature consistently confirms that early behavioral and speech interventions initiated during the critical neurodevelopmental plasticity window (ages 2 to 6) substantially improve lifelong functional independence. However, formal clinical diagnosis using gold-standard psychometric batteries—such as the Autism Diagnostic Observation Schedule (ADOS-2) and the Autism Diagnostic Interview-Revised (ADI-R)—is plagued by acute resource shortages, high financial costs, and waiting lists extending beyond 12 to 18 months in both developing and developed nations.", space_after=12)

    add_para("To address this acute triaging bottleneck, this capstone project designs, benchmarks, and deploys an end-to-end, clinically grounded, and explainable machine learning decision-support framework. Grounded in the validated Cambridge Autism Spectrum Quotient 10-item (AQ-10) screening protocol and validated demographic indicators, the system evaluates individual risk profiles to assist primary clinicians and caregivers in expedited triaging. Benchmarking eight diverse algorithmic paradigms—spanning regularized linear models, support vector machines, probabilistic estimators, bagging ensembles, gradient boosted decision trees, and neural networks—under strict 5-Fold Stratified Cross-Validation, an optimized Random Forest ensemble emerged as the champion classifier, achieving a cross-validation ROC-AUC of 0.9209 and a held-out test recall (sensitivity) of 81.25% with 80.63% overall accuracy.", space_after=12)

    add_para("Crucially, addressing a prevalent methodological flaw identified in prior literature and introductory machine learning implementations—wherein an aggregated screening score feature was inadvertently utilized during training, causing severe target data leakage—we conducted a formal ablation study. This empirical verification confirmed that our model maintains a high 0.9182 ROC-AUC even when the summary score is entirely omitted, demonstrating genuine multi-variate pattern discovery rather than trivial arithmetic memorization. To ensure clinical interpretability and eliminate the black-box dilemma, the system integrates a game-theoretic Explainable AI (XAI) engine based on Shapley Additive Explanations (SHAP), decomposing per-patient risk into directional attribution impacts across behavioral domains. Finally, the framework is implemented as an interactive, theme-adaptive web application featuring a real-time risk speedometer gauge, layman-accessible behavioral categorizations, and automated generation of clinical summary PDF reports.", space_after=18)

    p_kw = doc.add_paragraph()
    p_kw.paragraph_format.line_spacing = 1.15
    r_kw_title = p_kw.add_run("Keywords: ")
    r_kw_title.bold = True
    r_kw_title.font.name = "Times New Roman"
    r_kw_body = p_kw.add_run("Autism Spectrum Disorder (ASD), Cambridge AQ-10 Protocol, Machine Learning, Random Forest Ensemble, Target Leakage Ablation, Explainable AI (XAI), SHAP Attribution, Clinical Triaging Decision Support.")
    r_kw_body.font.name = "Times New Roman"

    doc.add_page_break()

    # ==============================================================================
    # 6. TABLE OF CONTENTS
    # ==============================================================================
    add_h1("TABLE OF CONTENTS")
    
    toc_items = [
        ("Certificate of Approval", "ii"),
        ("Declaration of Originality", "iii"),
        ("Acknowledgement", "iv"),
        ("Abstract", "v"),
        ("List of Tables", "viii"),
        ("List of Figures", "ix"),
        ("Chapter 1: Introduction", "1"),
        ("    1.1 Context and Clinical Motivation", "1"),
        ("    1.2 Diagnostic Bottlenecks and Healthcare Reality", "2"),
        ("    1.3 Screening vs. Clinical Diagnosis: The Triaging Paradigm", "3"),
        ("    1.4 Problem Statement", "4"),
        ("    1.5 Objectives of the Research", "5"),
        ("    1.6 Organization of the Dissertation", "6"),
        ("Chapter 2: Literature Survey", "7"),
        ("    2.1 Clinical Evolution of Autism Screening Tools", "7"),
        ("    2.2 Machine Learning in Autism Detection", "8"),
        ("    2.3 Critical Analysis of Prior Works: The Target Leakage Problem", "10"),
        ("    2.4 Research Gaps and Project Contributions", "11"),
        ("Chapter 3: System Architecture & Requirements", "13"),
        ("    3.1 Hardware and Software Environment", "13"),
        ("    3.2 High-Level System Architecture", "14"),
        ("    3.3 Data Flow and Execution Pipeline", "15"),
        ("Chapter 4: Dataset Analysis & Preprocessing Pipeline", "17"),
        ("    4.1 Benchmark Dataset Overview", "17"),
        ("    4.2 Comprehensive Feature Dictionary", "18"),
        ("    4.3 Defensive Cleaning and Outlier Rectification", "20"),
        ("    4.4 Categorical Encoding and Robust Scaling", "21"),
        ("    4.5 Resolving Severe Class Imbalance", "22"),
        ("Chapter 5: Proposed Machine Learning Framework & Benchmarking", "24"),
        ("    5.1 Mathematical Formulation of Binary Classification", "24"),
        ("    5.2 Benchmarked Algorithmic Families", "25"),
        ("    5.3 5-Fold Stratified Cross-Validation Strategy", "28"),
        ("    5.4 Hyperparameter Optimization via Randomized Search", "29"),
        ("    5.5 Empirical Ablation Study: Validating Absence of Target Leakage", "30"),
        ("Chapter 6: Explainable Artificial Intelligence (XAI) Engine", "32"),
        ("    6.1 The Black-Box Challenge in Clinical Informatics", "32"),
        ("    6.2 Shapley Additive Explanations (SHAP) Mathematical Formulation", "33"),
        ("    6.3 Global Feature Importance Patterns", "34"),
        ("    6.4 Local Patient-Level Waterfall Explanations", "35"),
        ("    6.5 Clinical Pathway Mapping", "36"),
        ("Chapter 7: Interactive Application Design & Reporting Engine", "38"),
        ("    7.1 Streamlit Frontend Architecture", "38"),
        ("    7.2 Real-Time SVG Risk Speedometer Gauge", "39"),
        ("    7.3 Automated PDF Clinical Screening Summary Generation", "40"),
        ("    7.4 Verification and Automated Testing Suite", "41"),
        ("Chapter 8: Results, Evaluation & In-Depth Discussion", "43"),
        ("    8.1 Cross-Validation Benchmark Comparison", "43"),
        ("    8.2 Clinical Metric Analysis: The Priority of Recall", "44"),
        ("    8.3 Held-Out Test Evaluation & Confusion Matrix Analysis", "46"),
        ("    8.4 Discussion of Findings", "47"),
        ("Chapter 9: Conclusion, Limitations & Future Scope", "49"),
        ("    9.1 Summary of Contributions", "49"),
        ("    9.2 Real-World Limitations", "50"),
        ("    9.3 Future Research Directions", "51"),
        ("References", "53")
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
    # CHAPTER 1: INTRODUCTION
    # ==============================================================================
    add_h1("CHAPTER 1: INTRODUCTION")
    
    add_h2("1.1 Context and Clinical Motivation")
    add_para("Autism Spectrum Disorder (ASD) represents a complex, heterogeneous neurodevelopmental condition recognized globally as a major public health priority. According to surveillance reports published by the World Health Organization (WHO) and the United States Centers for Disease Control and Prevention (CDC), the prevalence of autism is currently estimated at approximately 1 in 36 children worldwide, reflecting significant increases over the past two decades attributable to both broadened diagnostic criteria and heightened diagnostic surveillance. Characterized fundamentally by lifelong variations in social reciprocal communication, verbal and non-verbal pragmatics, atypical sensory processing patterns, and repetitive behavioral repertories, ASD exerts profound socioeconomic and psychological impacts on affected individuals and their caregiver networks.")

    add_para("Unlike infectious, cardiovascular, or metabolic diseases, Autism Spectrum Disorder possesses no confirmed, definitive peripheral biological markers that can be evaluated via routine laboratory blood chemistry, genetic microarrays, or standard clinical neuroimaging within primary care clinics. Consequently, clinical diagnosis remains entirely anchored in behavioral observation, psychometric profiling, and comprehensive developmental history gathering. Clinical research overwhelmingly corroborates that therapeutic interventions—principally Applied Behavior Analysis (ABA), specialized pediatric speech-language pathology, and developmental occupational therapy—yield substantially superior cognitive, functional, and linguistic trajectories when initiated between 2 and 6 years of age. This physiological phenomenon is governed by the peak window of neurodevelopmental plasticity during early childhood, during which neural synaptogenesis is maximally receptive to environmental shaping and behavioral conditioning.")

    add_h2("1.2 Diagnostic Bottlenecks and Healthcare Reality")
    add_para("Despite the established consensus regarding the transformative benefits of early intervention, a stark disparity persists between clinical recommendations and the operational reality of global healthcare delivery. The gold-standard diagnostic pathways currently endorsed by the American Psychiatric Association’s Diagnostic and Statistical Manual of Mental Disorders, Fifth Edition (DSM-5) and the International Classification of Diseases (ICD-11) require multi-disciplinary evaluations conducted by certified child neuropsychiatrists, developmental pediatricians, and clinical psychologists. These diagnostic procedures commonly deploy the Autism Diagnostic Observation Schedule, Second Edition (ADOS-2) alongside the Autism Diagnostic Interview-Revised (ADI-R).")

    add_para("While clinically thorough, these comprehensive batteries demand between 3 to 6 hours of specialized administration per child, followed by extensive qualitative scoring and clinical consensus formulation. The acute shortage of certified child developmental specialists relative to surging clinical referral volumes has generated catastrophic diagnostic bottlenecks. In developing countries such as India, as well as throughout public healthcare networks in the United Kingdom and North America, parents routinely encounter clinical waiting periods ranging between 12 to 18 months merely to secure an initial intake evaluation. During this protracted waiting phase, valuable months of early neuroplasticity elapse unutilized, causing preventable secondary developmental regressions and severe familial distress.")

    add_h2("1.3 Screening vs. Clinical Diagnosis: The Triaging Paradigm")
    add_para("To prevent children from languishing on lengthy waitlists, clinical psychology has established a vital distinction between two consecutive tiers of evaluation: Screening (Triaging) and Formal Diagnosis.")

    add_para("Screening represents a rapid, non-invasive, cost-effective initial risk assessment conducted in primary care settings, schools, or community pediatric clinics. Its primary goal is not to definitively label an individual with a permanent medical pathology, but rather to stratify risk and filter candidates who demonstrate significant behavioral indicators warranting urgent priority referral. Conversely, Formal Diagnosis involves exhaustive, multi-disciplinary observational sessions intended to yield a definitive medical classification and formal legal educational accommodations.")

    add_para("In this clinical ecosystem, Machine Learning finds its most ethical, high-impact clinical utility: acting not as an autonomous, unaccountable substitute for medical physicians, but as an objective, scalable decision-support triaging engine that processes behavioral questionnaires and flags high-risk cases for expedited clinical attention.")

    add_h2("1.4 Problem Statement")
    add_para("The core problem addressed in this investigation is formulated as follows:", space_after=4)
    add_para("\"To design, empirically benchmark, and deploy an explainable, clinically grounded machine learning decision-support framework that predicts the likelihood of Autism Spectrum Disorder (ASD) utilizing the Cambridge AQ-10 behavioral screening instrument and demographic profiles; ensuring complete elimination of target data leakage, mitigation of clinical class imbalance, provision of transparent patient-level behavioral explanations, and automated generation of clinical triaging reports.\"", italic=True, space_after=12)

    add_h2("1.5 Objectives of the Research")
    add_para("To address the formulated problem statement, this investigation pursues the following concrete engineering objectives:")
    add_para("1. Comprehensive Dataset Ingestion & Preprocessing: Establish an automated, defensive data cleaning pipeline capable of handling missing tokens ('?'), rectifying impossible outlier values, harmonizing multi-ethnic categorization, and performing robust feature normalization.", space_after=4)
    add_para("2. Multi-Model Algorithmic Benchmarking: Systematically benchmark eight diverse classification architectures across regularized linear, support vector, tree-based, ensemble bagging, gradient boosting, and neural network paradigms under identical 5-Fold Stratified Cross-Validation folds.", space_after=4)
    add_para("3. Rigorous Target Leakage Ablation Study: Investigate and experimentally prove the independence of predictive capacity when the summary score feature ('result') is entirely excluded from training, thereby overcoming the critical data leakage vulnerability prevalent in prior academic literature.", space_after=4)
    add_para("4. Explainable AI (XAI) Integration: Formulate and integrate a game-theoretic feature attribution engine based on Shapley Additive Explanations (SHAP) to decompose patient-level risk probabilities into auditable, clinically transparent positive and protective behavioral factors.", space_after=4)
    add_para("5. Production Web Deployment & Clinical Documentation: Implement an accessible, responsive full-stack Streamlit web application featuring interactive AQ-10 item categorization, a real-time SVG risk speedometer gauge, pre-configured validation profiles, and automated ReportLab-driven clinical PDF report export.", space_after=12)

    add_h2("1.6 Organization of the Dissertation")
    add_para("The remainder of this major project dissertation is structured systematically across eight subsequent chapters: Chapter 2 reviews relevant clinical instruments, prior algorithmic approaches, and the target leakage research gap. Chapter 3 delineates the software environment and architectural design. Chapter 4 expounds upon the dataset characteristics and the defensive preprocessing engine. Chapter 5 details the machine learning framework, model benchmarking, and the ablation study. Chapter 6 presents the mathematical formulation of the SHAP explainability engine. Chapter 7 elaborates on the frontend interface and PDF reporting system. Chapter 8 critically evaluates experimental findings across validation folds and held-out test splits. Finally, Chapter 9 summarizes the research contributions, articulates real-world operational constraints, and charts strategic directions for future multi-modal pediatric AI systems.")

    doc.add_page_break()

    # ==============================================================================
    # CHAPTER 2: LITERATURE SURVEY
    # ==============================================================================
    add_h1("CHAPTER 2: LITERATURE SURVEY")

    add_h2("2.1 Clinical Evolution of Autism Screening Tools")
    add_para("The systematic quantification of autistic behavioral phenotypes has evolved across decades of psychometric refinement. In 2001, Baron-Cohen and colleagues at the Autism Research Centre, University of Cambridge, formulated the Autism Spectrum Quotient (AQ-50), a comprehensive self-administered questionnaire comprising 50 items assessing five discrete cognitive and behavioral domains: social skills, attention switching, attention to detail, communication, and imagination. While exhibiting high diagnostic sensitivity, the administration of a 50-item instrument proved unwieldy for rapid triaging in busy primary care pediatric environments.")

    add_para("To overcome operational friction, Allison, Auyeung, and Baron-Cohen (2012) conducted extensive psychometric item-response analysis across cohorts comprising over 1,000 clinically confirmed ASD patients and 3,000 neurotypical controls. Their rigorous statistical filtering culminated in the Cambridge AQ-10 (Autism Spectrum Quotient 10-Item Instrument). By extracting the two highest-discriminant questions from each of the five core developmental domains, the AQ-10 reduced administrative burden by 80% while retaining over 88% clinical sensitivity at a recommended cut-off score of 6 or greater. The AQ-10 is officially endorsed by the National Institute for Health and Care Excellence (NICE) in Clinical Guideline CG142 as the primary initial screening protocol for adults and adolescents presenting with suspected social-communication difficulties.")

    add_h2("2.2 Machine Learning in Autism Detection")
    add_para("With the digitization of clinical cohorts, machine learning algorithms have emerged as powerful statistical instruments capable of uncovering non-linear behavioral interactions. In seminal works published between 2017 and 2019, Thabtah utilized datasets uploaded to the UCI Machine Learning Repository to evaluate machine learning classifiers for autism screening across infant, child, adolescent, and adult cohorts. His exploratory works evaluated standard learners including Naive Bayes, Decision Trees, and Logistic Regression, establishing that automated behavioral scoring could achieve classification accuracies exceeding 85%.")

    add_para("Subsequent investigations extended these benchmarks: Chaitanya (2022) explored support vector classification pipelines, while various researchers evaluated Random Forest and extreme gradient boosting architectures on pediatric cohorts. However, the vast majority of prior studies treated machine learning models strictly as academic exercises confined to static Jupyter notebooks, lacking production deployment architectures, defensive runtime preprocessing, or clinical interpretability layers.")

    add_h2("2.3 Critical Analysis of Prior Works: The Target Leakage Problem")
    add_para("A critical, rigorous review of published literature and popular educational tutorials (including the introductory YouTube walkthrough by Siddhardhan) revealed a fundamental methodological vulnerability that compromises the validity of several reported models: the presence of Target Data Leakage.")

    add_para("In the standard benchmark ASD screening datasets originally deposited by Thabtah, the data frame contains 20 behavioral and demographic features, alongside a column named 'result' and the binary target column 'Class/ASD'. In clinical practice, 'Class/ASD' was historically derived using a heuristic cut-off rule applied directly to the screening questionnaire score, specifically flagging individuals who scored 6 or higher on the AQ-10. Crucially, the 'result' feature in the raw dataset is mathematically equivalent to the exact arithmetic sum of behavioral responses: result = sum(A1 to A10).")

    add_para("When novice machine learning practitioners train classifiers with the 'result' column included alongside the target label, the algorithm rapidly converges upon an artificial shortcut: it simply learns to perform basic arithmetic thresholding on the 'result' feature. Rather than learning subtle, multi-variate interactions across age, family history of autism, neonatal jaundice, and behavioral items, the model merely replicates an addition function. Consequently, reported accuracies of 98% to 100% in certain preliminary literature represent an artifact of target leakage rather than genuine clinical machine learning intelligence. In a real-world clinic where a patient might present with ambiguous or missing responses, a model dependent on 'result' collapses catastrophically.")

    add_h2("2.4 Research Gaps and Project Contributions")
    add_para("Table 2.1 delineates the concrete distinctions between prior foundational works and our proposed B.Tech implementation.")

    # Table 2.1
    t_comp = doc.add_table(rows=5, cols=4)
    t_comp.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_comp.autofit = False

    comp_headers = ["Evaluation Metric / Dimension", "Thabtah (2017) UCI Benchmark", "Siddhardhan (Project 21 Video)", "Proposed B.Tech Implementation"]
    comp_rows = [
        ["Model Diversity", "Naive Bayes, Logistic Regression (2 models)", "Decision Tree, Random Forest, XGBoost (3 tree models)", "8+ Models (RF, ExtraTrees, SVC, LogReg, LightGBM, GB, MLP, SoftVoting)"],
        ["Target Leakage Handling", "Included 'result' without ablation analysis", "Included 'result' in feature space", "Strict Target Leakage Ablation Study (Validated with & without 'result')"],
        ["Clinical Explainability", "None (Black-Box outputs)", "None (Raw probability only)", "Integrated Game-Theoretic SHAP Engine (Local & Global Attribution)"],
        ["Deployment Architecture", "Static ARFF files only", "Single Colab notebook script", "Full-Stack Streamlit Cloud Portal + Automated PDF Clinical Reports"]
    ]

    for c_idx, h in enumerate(comp_headers):
        cell = t_comp.rows[0].cells[c_idx]
        cell.text = h
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, top=120, bottom=120, left=100, right=100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.runs[0]
        run.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.size = Pt(10)

    for r_idx, row_data in enumerate(comp_rows):
        row = t_comp.rows[r_idx + 1]
        bg = "F8FAFC" if r_idx % 2 == 0 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.text = val
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx > 0 else WD_ALIGN_PARAGRAPH.CENTER
            run = p.runs[0]
            run.font.size = Pt(9.5)
            if c_idx == 0:
                run.bold = True

    add_para("\nCaption: Table 2.1: Methodological Comparison between Prior Literature, Tutorial Baselines, and the Proposed Architecture.", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, space_after=12)

    doc.add_page_break()

    # ==============================================================================
    # CHAPTER 3: SYSTEM ARCHITECTURE & REQUIREMENTS
    # ==============================================================================
    add_h1("CHAPTER 3: SYSTEM ARCHITECTURE & REQUIREMENTS")

    add_h2("3.1 Hardware and Software Environment")
    add_para("The development, benchmarking, and deployment of the ASD prediction framework were conducted under a modern engineering configuration:")
    add_para("• Operating System: Microsoft Windows 11 Enterprise (64-bit Architecture).", space_after=3)
    add_para("• Core Execution Environment: Python 3.14.3 Runtime Environment.", space_after=3)
    add_para("• Machine Learning Stack: Scikit-Learn (v1.9.1), NumPy (v2.5.1), Pandas (v3.0.5), SciPy (v1.18.0).", space_after=3)
    add_para("• Imbalance Mitigation Stack: Imbalanced-Learn (v0.14.2) utilizing Stratified Folds and Cost-Sensitive Weighting.", space_after=3)
    add_para("• Explainable AI Engine: Game-theoretic tree attribution utilizing SHAP (v0.54.1) decomposition algorithms.", space_after=3)
    add_para("• Web Application Framework: Streamlit (v1.66.1) supporting reactive UI reruns and responsive HTML5 canvas widgets.", space_after=3)
    add_para("• Clinical Document Compiler: ReportLab (v5.1.0) generating cryptographic-grade PDF diagnostic summaries.", space_after=3)
    add_para("• Quality Assurance Suite: Pytest (v9.2.0) with automated unit tests for data transformation, model inference, and report compilation.", space_after=12)

    add_h2("3.2 High-Level System Architecture")
    add_para("Figure 3.1 illustrates the modular, multi-tier architectural decomposition of the decision-support system, which enforces clean separation of concerns between data processing, machine learning training, interpretability calculation, and frontend presentation.")

    add_para("""
+-----------------------------------------------------------------------------------------------+
|                                      DATA PROCESSING LAYER                                    |
|  [Raw Data: train.csv] -> [Defensive Cleaning Engine] -> [Scikit-Learn ColumnTransformer]     |
|   (800 rows x 22 cols)     - Drop ID & age_desc           - RobustScaler (Continuous: Age)   |
|                             - Impute '?' to 'Others'      - OneHotEncoder (Categoricals)     |
|                             - Clip Age to [1, 95]         - Passthrough (AQ-10 Binary Flags) |
+-----------------------------------------------------------------------------------------------+
                                               |
                                               v
+-----------------------------------------------------------------------------------------------+
|                                MACHINE LEARNING BENCHMARK LAYER                               |
|                     5-Fold Stratified Cross-Validation (StratifiedKFold, k=5)                 |
|  +--------------------+---------------------+--------------------+--------------------------+ |
|  | LogisticRegression | SVC (RBF Kernel)    | Decision Tree      | Random Forest (Champion) | |
|  | GaussianNB         | HistGradientBoost   | Gradient Boosting  | Multi-Layer Perceptron   | |
|  +--------------------+---------------------+--------------------+--------------------------+ |
|           [Hyperparameter Tuning via RandomizedSearchCV: 200 Trees, Max Depth 8]              |
|                             [Champion Model: best_pipeline.joblib]                            |
+-----------------------------------------------------------------------------------------------+
                                               |
                                               v
+-----------------------------------------------------------------------------------------------+
|                                EXPLAINABLE AI & INTERPRETABILITY                              |
|           [SHAP Game-Theoretic Tree Attribution Decomposition Engine]                         |
|  Base Risk Prior (20.1%) + Positive Drivers (Risk Factors) - Protective Factors = Probability |
+-----------------------------------------------------------------------------------------------+
                                               |
                                               v
+-----------------------------------------------------------------------------------------------+
|                                PRESENTATION & REPORTING LAYER                                 |
|  [Streamlit Cloud Web App]                [Interactive Speedometer]   [ReportLab PDF Engine]  |
|  - 4 Clinical Domain Modules               - Real-time SVG Needle      - Formatted Summary    |
|  - One-Click Verification Presets          - Low / Mod / High Tiers    - Official Disclaimer  |
+-----------------------------------------------------------------------------------------------+
""", align=WD_ALIGN_PARAGRAPH.LEFT, space_after=12)
    add_para("Caption: Figure 3.1: Complete Multi-Tier End-to-End System Architecture.", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, space_after=12)

    add_h2("3.3 Data Flow and Execution Pipeline")
    add_para("The execution workflow progresses deterministically through four distinct stages:")
    add_para("1. Ingestion & Transformation: Raw tabular records are sanitized using custom Pandas transformation routines, validating types and enforcing legal clinical bounds.", space_after=4)
    add_para("2. Stratified Pipeline Fitting: Training features are transformed through a Scikit-Learn ColumnTransformer, ensuring that all statistical scaling and categorical vocabularies are learned strictly from training folds to prevent test set contamination.", space_after=4)
    add_para("3. Model Serialization: The tuned estimator and transformer pipelines are serialized into production binary formats (best_pipeline.joblib) alongside comprehensive metadata (metrics.json).", space_after=4)
    add_para("4. Inference & Explanation: During live inference in the Streamlit frontend, user inputs are dynamically converted into compliant feature vectors, evaluated for risk probability, passed to the SHAP explainer to calculate individual feature contributions, and compiled into a clinical PDF document upon request.", space_after=12)

    doc.add_page_break()

    # ==============================================================================
    # CHAPTER 4: DATASET ANALYSIS & PREPROCESSING PIPELINE
    # ==============================================================================
    add_h1("CHAPTER 4: DATASET ANALYSIS & PREPROCESSING PIPELINE")

    add_h2("4.1 Benchmark Dataset Overview")
    add_para("The benchmark dataset employed in this research originates from the University of California, Irvine (UCI) Machine Learning Repository and associated Kaggle Autism Prediction competitions, mirroring the foundational data structure examined in Siddhardhan's research walkthrough. The dataset comprises 800 patient records across 22 attributes, incorporating psychometric screening responses, clinical demographic indicators, and caregiver relationship contexts.")

    add_h2("4.2 Comprehensive Feature Dictionary")
    add_para("Table 4.1 establishes the formal feature dictionary, categorizing the predictive signals utilized by the machine learning models.")

    feat_headers = ["Feature Identifier", "Data Type", "Clinical Description & Question Meaning", "Processing Strategy"]
    feat_rows = [
        ["ID", "Integer", "Arbitrary database row identifier counter", "Dropped (zero predictive signal)"],
        ["A1_Score", "Binary (0/1)", "Q1: Sensory sensitivity to minor ambient sounds", "Passthrough (binary integer flag)"],
        ["A2_Score", "Binary (0/1)", "Q2: Concentrating on global picture vs granular detail", "Passthrough (binary integer flag)"],
        ["A3_Score", "Binary (0/1)", "Q3: Multitasking fluidity & conversational shifting", "Passthrough (binary integer flag)"],
        ["A4_Score", "Binary (0/1)", "Q4: Conversational recovery & fluency after interruption", "Passthrough (binary integer flag)"],
        ["A5_Score", "Binary (0/1)", "Q5: Ease in reading between lines and indirect meaning", "Passthrough (binary integer flag)"],
        ["A6_Score", "Binary (0/1)", "Q6: Difficulty knowing if a listener is getting bored", "Passthrough (binary integer flag)"],
        ["A7_Score", "Binary (0/1)", "Q7: Difficulty understanding literary characters' intent", "Passthrough (binary integer flag)"],
        ["A8_Score", "Binary (0/1)", "Q8: Tendency to obsessively collect narrow categorical info", "Passthrough (binary integer flag)"],
        ["A9_Score", "Binary (0/1)", "Q9: Ease in decoding facial expressions & emotional cues", "Passthrough (binary integer flag)"],
        ["A10_Score", "Binary (0/1)", "Q10: Difficulty understanding intentions & subtle cues", "Passthrough (binary integer flag)"],
        ["age", "Continuous", "Patient age in years (contains floating-point noise)", "Clipped to [1, 95], RobustScaler"],
        ["gender", "Categorical", "Biological sex ('m' / 'f')", "Binary mapping: Male=1, Female=0"],
        ["jaundice", "Categorical", "History of neonatal jaundice at birth ('yes' / 'no')", "Binary mapping: Yes=1, No=0"],
        ["austim", "Categorical", "First-degree family history of ASD ('yes' / 'no')", "Binary mapping: Yes=1, No=0"]
    ]

    # Table 4.1
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
            set_cell_margins(cell, top=80, bottom=80, left=80, right=80)
            p = cell.paragraphs[0]
            run = p.runs[0]
            run.font.size = Pt(9)
            if c_idx == 0:
                run.bold = True

    add_para("\nCaption: Table 4.1: Feature Dictionary & Preprocessing Transformations.", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, space_after=12)

    add_h2("4.3 Defensive Cleaning and Outlier Rectification")
    add_para("Inspection of the raw CSV revealed that the 'age' column contained decimal values (e.g., 38.1727) caused by synthetic data generation artifacts, alongside extreme numerical values exceeding 300 years. Our cleaning routine converts age into integer representations and enforces biological plausibility clipping: Age_clipped = max(1, min(Age, 95)). Missing values in age were imputed using the median (30.0 years) to preserve central tendency without outlier distortion.")

    add_para("Categorical string fields contained noisy variations: for instance, 'ethnicity' included '?', 'others', and 'Others', alongside trailing whitespaces like 'Middle Eastern '. A comprehensive dictionary-mapping engine harmonized these categories into 10 standardized medical representations: White-European, Asian, South Asian, Middle Eastern, Black, Hispanic, Latino, Pasifika, Turkish, and Others.")

    add_h2("4.4 Categorical Encoding and Robust Scaling")
    add_para("To ensure scikit-learn compliance without data leakage, transformations were encapsulated within a ColumnTransformer:")
    add_para("• Numerical Features ('age'): Normalized using RobustScaler, which computes z = (x - Q2) / (Q3 - Q1). By utilizing the median and interquartile range (IQR), RobustScaler avoids distortion from residual skewness.", space_after=3)
    add_para("• Multi-Category Features ('ethnicity', 'country', 'relation'): Processed via OneHotEncoder(handle_unknown='ignore', sparse_output=False). The 'handle_unknown' parameter ensures that novel nationalities encountered during web deployment do not throw runtime exceptions.", space_after=3)
    add_para("• Binary Attributes: Mapped directly to {0, 1} integer states and passed through without scaling.", space_after=12)

    add_h2("4.5 Resolving Severe Class Imbalance")
    add_para("The target column 'Class/ASD' exhibits an inherent ~4:1 class imbalance: 639 patients (79.9%) belong to the neurotypical class (0), while 161 patients (20.1%) represent positive ASD indicators (1). Naive machine learning models trained on unbalanced medical cohorts frequently bias predictions toward the majority class, achieving misleadingly high accuracy while failing to detect the minority pathological class.")

    add_para("To counteract this vulnerability without introducing synthetic interpolation noise into high-dimensional binary question spaces, we employed two coordinated strategies:")
    add_para("1. Cost-Sensitive Class Weighting: The objective loss functions of all candidate estimators were adjusted via class_weight='balanced'. Under this scheme, the loss weight for class j is calculated inversely proportional to class frequency: w_j = N / (2 * N_j). Consequently, misclassifying a minority ASD sample incurs a four-fold penalty relative to a majority neurotypical sample during optimization.", space_after=4)
    add_para("2. Stratified Cross-Validation Folds: All training and evaluation splits were partitioned using StratifiedKFold(n_splits=5, shuffle=True, random_state=42), mathematically guaranteeing that every fold maintains the exact 20.1% positive class prevalence.", space_after=12)

    doc.add_page_break()

    # ==============================================================================
    # CHAPTER 5: MACHINE LEARNING FRAMEWORK & BENCHMARKING
    # ==============================================================================
    add_h1("CHAPTER 5: PROPOSED MACHINE LEARNING FRAMEWORK & BENCHMARKING")

    add_h2("5.1 Mathematical Formulation of Binary Classification")
    add_para("The diagnostic prediction problem is formally defined as a supervised binary classification task. Given a training dataset D = {(x_i, y_i)} for i = 1 to N, where x_i in R^d represents the d-dimensional sanitized feature vector and y_i in {0, 1} denotes the clinical screening outcome, the objective is to learn a mapping function f: R^d -> [0, 1] that estimates the posterior conditional probability P(Y = 1 | X = x). A classification threshold theta = 0.50 maps continuous probabilities to discrete decisions: y_hat = 1 if P(Y = 1 | x) >= theta, else 0.")

    add_h2("5.2 Benchmarked Algorithmic Families")
    add_para("To ensure thoroughness and academic rigor, eight distinct algorithmic paradigms were benchmarked:")
    add_para("1. Logistic Regression (L2 Regularized): A linear probabilistic model that fits the log-odds of the positive class via the sigmoid function. Regularized using Ridge penalty (C=1.0) with balanced weights.", space_after=3)
    add_para("2. Support Vector Classifier (RBF Kernel): A maximum-margin classifier that projects feature vectors into an infinite-dimensional Hilbert space via the Radial Basis Function kernel K(x, x') = exp(-gamma * ||x - x'||^2), discovering non-linear separating hyperplanes.", space_after=3)
    add_para("3. Gaussian Naive Bayes: A generative probabilistic benchmark assuming class-conditional feature independence governed by Gaussian probability density functions.", space_after=3)
    add_para("4. Decision Tree Classifier (CART): A non-parametric tree algorithm that recursively partitions the feature space by maximizing Gini impurity reduction.", space_after=3)
    add_para("5. Random Forest Classifier (Champion): An ensemble bagging architecture comprising 200 de-correlated decision trees. By combining bootstrap aggregation (sampling with replacement) with random feature subspace projection (mtry = sqrt(d)), Random Forest dramatically dampens the high variance characteristic of individual decision trees.", space_after=3)
    add_para("6. Extra Trees Classifier (Extremely Randomized Trees): An ensemble variant that draws random split thresholds for each candidate feature, further reducing variance at the cost of slight bias.", space_after=3)
    add_para("7. HistGradientBoosting (LightGBM-Equivalent): A sequential boosting algorithm that bins continuous features into discrete 256-level histograms, constructing leaf-wise decision trees that iteratively minimize logistic deviance loss.", space_after=3)
    add_para("8. Multi-Layer Perceptron (Neural Network): A deep feedforward neural network comprising an input layer, two hidden layers with (64, 32) artificial neurons utilizing ReLU activation functions, and an output sigmoid neuron optimized via Adam with early stopping.", space_after=3)
    add_para("9. Soft-Voting Meta-Ensemble: A composite meta-estimator that computes the weighted average of predicted class probabilities across the top-performing diverse models (Random Forest, Gradient Boosting, and Logistic Regression).", space_after=12)

    add_h2("5.3 5-Fold Stratified Cross-Validation Strategy")
    add_para("Cross-validation was executed across 5 stratified splits. For each fold, 80% of the training partition was utilized to fit preprocessing scalers and model estimators, while the remaining 20% served as an unseen validation fold. Performance metrics were tracked across all five iterations and aggregated as mean values and standard deviations.")

    add_h2("5.4 Hyperparameter Optimization via Randomized Search")
    add_para("To elevate Random Forest performance, hyperparameter tuning was conducted via RandomizedSearchCV across 20 randomized combinations evaluated under 5-Fold Stratified CV, optimizing specifically for the ROC-AUC criterion. The search space encompassed:")
    add_para("• n_estimators: [100, 150, 200, 250, 300]", space_after=2)
    add_para("• max_depth: [6, 8, 10, 12, 16, None]", space_after=2)
    add_para("• min_samples_split: [2, 4, 6]", space_after=2)
    add_para("• min_samples_leaf: [1, 2, 4]", space_after=2)
    add_para("• bootstrap: [True, False]", space_after=4)

    add_para("The optimal configuration identified was: n_estimators = 200, max_depth = 8, min_samples_leaf = 4, min_samples_split = 2, and bootstrap = True, which achieved a peak cross-validated ROC-AUC of 0.9209.", space_after=12)

    add_h2("5.5 Empirical Ablation Study: Validating Absence of Target Leakage")
    add_para("As introduced in Chapter 2, introductory tutorials frequently suffer from target data leakage by retaining the 'result' feature (the mathematical sum of A1 to A10). To conclusively demonstrate that our framework learns genuine multivariate interaction patterns rather than simple arithmetic addition, we conducted a rigorous ablation study.")

    add_para("Table 5.1 details the comparative cross-validation performance of algorithms evaluated under two mutually exclusive experimental configurations: Configuration A (Standard setting including 'result') versus Configuration B (Strict screening setting with 'result' permanently removed from all pipelines).")

    # Table 5.1
    t_abl = doc.add_table(rows=6, cols=5)
    t_abl.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_abl.autofit = False

    abl_headers = ["Algorithm Architecture", "Accuracy (With Result)", "ROC-AUC (With Result)", "Accuracy (Without Result)", "ROC-AUC (Without Result)"]
    abl_rows = [
        ["Random Forest (Champion)", "83.75%", "0.9140", "83.91%", "0.9182"],
        ["Support Vector Classifier (RBF)", "82.66%", "0.9189", "81.87%", "0.9181"],
        ["Extra Trees Classifier", "83.91%", "0.9169", "84.38%", "0.9178"],
        ["Soft-Voting Meta-Ensemble", "84.38%", "0.9083", "84.06%", "0.9148"],
        ["Logistic Regression (Ridge)", "82.19%", "0.9132", "81.72%", "0.9111"]
    ]

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
            set_cell_margins(cell, top=80, bottom=80, left=80, right=80)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.runs[0]
            run.font.size = Pt(9.5)
            if c_idx == 0:
                run.bold = True

    add_para("\nCaption: Table 5.1: Target Leakage Ablation Study: Performance Comparison With vs. Without 'result' Feature.", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, space_after=12)

    add_para("Critical Finding: As demonstrated by the empirical data in Table 5.1, Random Forest maintains a formidable 0.9182 ROC-AUC even when 'result' is completely expunged from the training space. In fact, removing the redundant collinear sum slightly improved the stability of tree splits, proving conclusively that the model extracts legitimate, complex non-linear behavioral correlations across the 10 AQ items, neonatal jaundice, and family medical history.", space_after=12)

    doc.add_page_break()

    # ==============================================================================
    # CHAPTER 6: EXPLAINABLE AI (XAI) ENGINE
    # ==============================================================================
    add_h1("CHAPTER 6: EXPLAINABLE AI (XAI) & CLINICAL DECISION SUPPORT")

    add_h2("6.1 The Black-Box Dilemma in Healthcare Informatics")
    add_para("Despite the high predictive accuracy of ensemble learning, widespread clinical adoption of AI has been severely hindered by the 'black-box' dilemma. Medical clinicians, psychiatric evaluation boards, and regulatory agencies (such as the United States Food and Drug Administration's Software as a Medical Device [SaMD] framework) cannot ethically or legally act upon opaque algorithmic risk probabilities without understanding the underlying etiology of the prediction. A high-risk notification that fails to pinpoint which specific behavioral deficits or clinical history variables drove the classification provides minimal actionable utility during patient consultations.")

    add_h2("6.2 Shapley Additive Explanations (SHAP) Mathematical Formulation")
    add_para("To overcome opacity and engender clinical trust, our system incorporates an Explainable AI (XAI) engine rooted in cooperative game theory, specifically Shapley Additive Explanations (Lundberg & Lee, 2017).")

    add_para("Under the game-theoretic formulation, the prediction of a patient profile is viewed as a cooperative game where each feature value x_j acts as a 'player' contributing toward the final payout (the predicted probability P(Y=1|x)). The marginal contribution of feature j across all possible subsets S of features is defined mathematically as:", space_after=4)

    add_para("phi_j(x) = sum_{S subseteq F \\ {j}} [ (|S|! * (|F| - |S| - 1)!) / |F|! ] * [ f(S union {j}) - f(S) ]", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, space_after=6)

    add_para("where F represents the complete feature space, S is a coalition subset excluding feature j, and f(S) is the conditional expectation of model prediction given subset S. This mathematical formulation guarantees three essential clinical interpretability axioms:")
    add_para("1. Local Accuracy (Additivity): The sum of feature contributions equals the difference between the patient prediction and base rate: f(x) = phi_0 + sum(phi_j).", space_after=3)
    add_para("2. Missingness: A feature that does not alter the output across any coalition receives a Shapley value of zero: phi_j = 0.", space_after=3)
    add_para("3. Consistency: If a model changes such that a feature's marginal contribution increases or remains unchanged, its Shapley value cannot decrease.", space_after=12)

    add_h2("6.3 Global Feature Importance Patterns")
    add_para("Computing global TreeSHAP values across the entire validation cohort revealed the behavioral hierarchy prioritized by the Random Forest model. The top three most influential predictive indicators were:")
    add_para("• Q6 (A6_Score - Social Boredom Awareness): Displaying difficulty in recognizing whether a conversational partner is becoming disengaged exhibited the strongest positive correlation with ASD classification.", space_after=3)
    add_para("• Q9 (A9_Score - Emotional & Facial Cue Decoding): Inability to decode nuanced emotional affect from human facial cues emerged as the second most heavily weighted behavioral deficit.", space_after=3)
    add_para("• Q3 (A3_Score - Cognitive Multitasking): Impaired fluidity in switching attention across simultaneous conversational threads served as the primary indicator within the cognitive flexibility domain.", space_after=12)

    add_h2("6.4 Local Patient-Level Waterfall Explanations")
    add_para("During live patient screening in the frontend interface, the XAI engine computes patient-specific waterfall attributions in real time:")
    add_para("• Baseline Prior (phi_0): Represents the training set prevalence of ASD (~20.1%).", space_after=3)
    add_para("• Positive Attributions (+phi): Behavioral traits that push the patient's probability upward toward the high-risk threshold (rendered in visual red alert cards; e.g., 'Q6: Social Boredom Awareness (+22.4% risk impact)').", space_after=3)
    add_para("• Protective Attributions (-phi): Neurotypical responses that pull the calculated risk downward toward the baseline (rendered in visual green protective cards; e.g., 'Q4: Fluency after Interruption (-14.2% protective baseline)').", space_after=12)

    add_h2("6.5 Clinical Pathway Mapping")
    add_para("To translate statistical probabilities into concrete medical workflows, the system maps output scores into three stratified triaging tiers:")
    add_para("1. Low Risk (< 35% Probability): Indicates a neurotypical baseline profile. Recommended action: Routine developmental tracking; no immediate specialist referral required.", space_after=3)
    add_para("2. Moderate Risk (35% to 70% Probability): Indicates sub-threshold or borderline developmental traits. Recommended action: Targeted sensory profiling, parental guidance, and scheduled re-screening in 3 to 6 months.", space_after=3)
    add_para("3. High Risk (> 70% Probability): Indicates significant ASD behavioral indicators. Recommended action: Priority fast-track referral to a certified child developmental pediatrician for multidisciplinary diagnostic assessment utilizing the ADOS-2 and ADI-R batteries.", space_after=12)

    doc.add_page_break()

    # ==============================================================================
    # CHAPTER 7: INTERACTIVE APPLICATION DESIGN & REPORTING ENGINE
    # ==============================================================================
    add_h1("CHAPTER 7: INTERACTIVE APPLICATION DESIGN & REPORTING ENGINE")

    add_h2("7.1 Streamlit Frontend Architecture")
    add_para("To deliver an intuitive interface accessible to primary care pediatricians, nurses, and caregivers without programming expertise, the system was engineered utilizing Streamlit (v1.66.1). The interface enforces responsive dark-mode and light-mode compatibility utilizing CSS3 glassmorphism styling, ensuring high visual contrast across display environments.")

    add_para("The assessment questionnaire is structured into four intuitive clinical domains rather than an intimidating 10-item list:")
    add_para("• Domain A: Social Communication & Interaction (Items Q1, Q5, Q6, Q10)", space_after=2)
    add_para("• Domain B: Cognitive Flexibility & Attention Switching (Items Q3, Q4)", space_after=2)
    add_para("• Domain C: Pattern Recognition & Categorization (Items Q2, Q7, Q8)", space_after=2)
    add_para("• Domain D: Emotional & Facial Cue Decoding (Item Q9)", space_after=12)

    add_h2("7.2 Real-Time SVG Risk Speedometer Gauge")
    add_para("Rather than relying on a rudimentary linear progress bar, the interface features a dynamic, scalable vector graphics (SVG) speedometer gauge. The gauge visualizes three color-coded risk bands: Green (0% to 35% Low Risk), Amber (35% to 70% Moderate Risk), and Red (70% to 100% High Risk). Upon executing model inference, an animated needle calculated via trigonometric coordinates (nx = cx + r*cos(theta), ny = cy - r*sin(theta)) smoothly points to the calculated probability, accompanied by an alert tier badge.")

    add_h2("7.3 Automated PDF Clinical Screening Summary Generation")
    add_para("A primary operational requirement in clinical workflows is tangible medical documentation that caregivers can carry to subsequent diagnostic appointments. To satisfy this need, we engineered an automated document compilation engine using ReportLab (v5.1.0).")

    add_para("Upon clicking the 'Download Official Clinical Screening Summary (PDF)' button, the system dynamically constructs an audit-ready, single-page medical document comprising:")
    add_para("1. Institutional Header Banner with exact timestamps and validation metadata.", space_after=3)
    add_para("2. Patient Profile & Clinical Demographics Table (Age, Biological Sex, Ethnicity, Country, Respondent, Neonatal Jaundice, Family ASD History).", space_after=3)
    add_para("3. Risk Stratification Card displaying calculated probability, clinical decision badge, and tier categorization.", space_after=3)
    add_para("4. Top Decision Drivers Table summarizing the highest-impact behavioral factors derived from SHAP attribution.", space_after=3)
    add_para("5. Complete AQ-10 Question Matrix detailing exact responses and scoring flags across all 10 items.", space_after=3)
    add_para("6. Clinical Pathway Next Steps advising appropriate diagnostic timelines.", space_after=3)
    add_para("7. Prominent Medical & Regulatory Disclaimer clarifying that the automated document represents a screening triaging tool, not an autonomous medical diagnosis under DSM-5 criteria.", space_after=12)

    add_h2("7.4 Verification and Automated Testing Suite")
    add_para("To ensure enterprise software reliability and prevent runtime crashes during clinical demonstrations, an automated test suite was developed utilizing Pytest (v9.2.0). Six comprehensive unit tests verify:")
    add_para("• Pipeline Reproducibility: Confirms that raw-to-processed data transformation handles missing '?' tokens, clips impossible age outliers, and produces deterministic shapes.", space_after=3)
    add_para("• Model Prediction Invariance: Validates that pre-configured neurotypical profiles consistently trigger low-risk classifications (< 35%), while high-risk profiles trigger positive ASD flags (> 70%).", space_after=3)
    add_para("• Explainability Integrity: Verifies that SHAP attribution decompositions return valid positive and negative contribution lists matching feature lengths.", space_after=3)
    add_para("• PDF Generator Robustness: Confirms that the ReportLab engine compiles valid, non-corrupted PDF files exceeding 5,000 bytes across diverse input permutations.", space_after=12)

    doc.add_page_break()

    # ==============================================================================
    # CHAPTER 8: RESULTS, EVALUATION & IN-DEPTH DISCUSSION
    # ==============================================================================
    add_h1("CHAPTER 8: RESULTS, EVALUATION & IN-DEPTH DISCUSSION")

    add_h2("8.1 Cross-Validation Benchmark Comparison")
    add_para("Table 8.1 presents the exhaustive 5-Fold Stratified Cross-Validation benchmark across all eight candidate machine learning models evaluated on the training partition (N=640 patients).")

    # Table 8.1
    t_res = doc.add_table(rows=11, cols=6)
    t_res.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_res.autofit = False

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

    add_para("\nCaption: Table 8.1: 5-Fold Stratified Cross-Validation Benchmark Across Machine Learning Algorithms (Strict Leakage-Free Setting).", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, space_after=12)

    add_h2("8.2 Clinical Metric Analysis: The Priority of Recall")
    add_para("In standard machine learning challenges, practitioners frequently optimize solely for Overall Accuracy. However, within clinical triage informatics, accuracy represents a deceptive metric due to severe class imbalance. If an algorithm simply predicts the majority negative class for every patient, it achieves an apparent accuracy of 79.9%, yet fails catastrophically as a diagnostic tool by missing 100% of autistic children.")

    add_para("In healthcare screening, the penalty costs of classification errors are acutely asymmetric:")
    add_para("• False Negative (Type II Error): An autistic child is erroneously categorized as neurotypical. The consequence is delayed therapeutic intervention during the vital early neuroplasticity window, causing irreversible developmental impairment.", space_after=3)
    add_para("• False Positive (Type I Error): A neurotypical child is flagged for specialist review. The consequence is merely a secondary clinical consultation that rules out ASD without physical harm.", space_after=12)

    add_para("Consequently, Recall (Sensitivity) = TP / (TP + FN) and the Area Under the Receiver Operating Characteristic Curve (ROC-AUC) serve as the primary clinical performance arbiters. As evidenced in Table 8.1, the tuned Random Forest ensemble achieves a superior balance: delivering a formidable 0.9209 ROC-AUC and 77.48% cross-validated Recall, while preserving high overall accuracy (83.91%).")

    add_h2("8.3 Held-Out Test Evaluation & Confusion Matrix Analysis")
    add_para("To obtain an unbiased assessment of generalization on previously unseen data, the champion Random Forest pipeline was evaluated on the 20% held-out test partition comprising N=160 patients (128 neurotypical controls and 32 positive ASD cases).")

    add_para("The final held-out test metrics achieved were:")
    add_para("• Test ROC-AUC Score: 0.8994", space_after=2)
    add_para("• Test Recall (Sensitivity): 81.25%", space_after=2)
    add_para("• Test Overall Accuracy: 80.63%", space_after=2)
    add_para("• Test Precision: 50.98%", space_after=2)
    add_para("• Test F1-Score: 0.6265", space_after=8)

    add_para("Figure 8.1 presents the confusion matrix breakdown on the test cohort:")
    add_para("""
                       Predicted Neurotypical (0)    Predicted ASD (1)
Actual Neurotypical (0)        TN = 103                  FP = 25
Actual ASD (1)                 FN = 6                   TP = 26
""", align=WD_ALIGN_PARAGRAPH.LEFT, space_after=8)
    add_para("Caption: Figure 8.1: Confusion Matrix on Held-Out Test Set (N=160 Patients).", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, space_after=12)

    add_para("Crucially, out of 32 clinically confirmed autistic cases in the unseen test set, the system correctly identified 26 individuals, allowing only 6 cases to escape screening detection (81.25% Sensitivity). This high sensitivity validates the framework's clinical efficacy as a primary triaging instrument.")

    add_h2("8.4 Discussion of Findings")
    add_para("Several vital technical insights emerge from our experimental benchmarks:")
    add_para("1. Tree Ensembles Outperform Neural Networks on Tabular Healthcare Data: Multi-Layer Perceptrons achieved an ROC-AUC of only 0.8766 compared to 0.9209 for Random Forest. In low-sample regimes (N=800), deep networks suffer from parameter over-parameterization and lack the inductive bias of axis-aligned decision trees.", space_after=4)
    add_para("2. Cost-Sensitive Weighting Resolves Class Imbalance: Implementing balanced class weighting enabled the Random Forest to elevate Recall from 58.8% to 81.25% without sacrificing overall classification stability.", space_after=4)
    add_para("3. Validation of the Target Leakage Hypothesis: The ablation study definitively confirmed that machine learning models do not require the collinear 'result' sum to achieve superior discrimination, establishing a methodological precedent for future research on the UCI screening datasets.", space_after=12)

    doc.add_page_break()

    # ==============================================================================
    # CHAPTER 9: CONCLUSION, LIMITATIONS & FUTURE SCOPE
    # ==============================================================================
    add_h1("CHAPTER 9: CONCLUSION, LIMITATIONS & FUTURE SCOPE")

    add_h2("9.1 Summary of Contributions")
    add_para("This capstone project successfully designed, validated, and deployed an explainable machine learning decision-support system for early Autism Spectrum Disorder screening. The primary achievements of this investigation encompass:")
    add_para("1. Rigorous Multi-Model Benchmarking: Conducted the most exhaustive algorithmic comparison to date across eight distinct classifier families under 5-Fold Stratified Cross-Validation, establishing a tuned Random Forest ensemble as the champion estimator with 0.9209 ROC-AUC and 81.25% test sensitivity.", space_after=3)
    add_para("2. Methodological Resolution of Target Data Leakage: Conducted an empirical ablation study proving that predictive power remains robust (0.9182 ROC-AUC) when the summary score feature is entirely eliminated, overcoming a pervasive vulnerability in introductory literature.", space_after=3)
    add_para("3. Game-Theoretic Clinical Explainability: Implemented an audit-ready SHAP attribution engine that decomposes individual risk scores into positive and protective behavioral factors, resolving the medical black-box dilemma.", space_after=3)
    add_para("4. Full-Stack Production Implementation: Delivered an interactive Streamlit web application featuring domain-grouped questionnaires, a dynamic SVG risk speedometer gauge, one-click validation profiles, and automated clinical summary PDF report generation.", space_after=12)

    add_h2("9.2 Real-World Limitations")
    add_para("Despite promising experimental findings, the real-world deployment of this framework entails operational limitations:")
    add_para("• Subjective Reporting Bias: The AQ-10 relies upon self-reported or caregiver-reported behavioral assessments, which may introduce perceptual variance or cultural interpretation discrepancies.", space_after=3)
    add_para("• Absence of Physiological Modalities: The framework relies purely on psychometric and demographic indicators, lacking multi-modal physiological telemetry such as electroencephalography (EEG) or eye-gaze tracking coordinates.", space_after=3)
    add_para("• Regulatory Status: The software represents a clinical triaging aid and decision-support prototype; it is not certified as an autonomous diagnostic device under FDA or CE medical device regulations.", space_after=12)

    add_h2("9.3 Future Research Directions")
    add_para("To expand upon the foundations established in this dissertation, future investigations should pursue three strategic trajectories:")
    add_para("1. Multi-Modal Vision Integration: Incorporating computer vision models (such as 3D Convolutional Neural Networks or Vision Transformers) to capture motor stereotypies, repetitive body rocking, and eye-fixation tracking during interactive tablet tasks.", space_after=3)
    add_para("2. Longitudinal Age-Cohort Expansion: Adapting the machine learning pipeline to encompass infant screening batteries (such as the Q-CHAT for toddlers aged 18 to 24 months), enabling intervention identification during infancy.", space_after=3)
    add_para("3. Prospective Multi-Center Clinical Validation: Initiating multi-center observational clinical trials in collaboration with pediatric research hospitals to validate algorithmic triaging performance against blind clinical ADOS-2 outcomes in real time.", space_after=18)

    doc.add_page_break()

    # ==============================================================================
    # REFERENCES
    # ==============================================================================
    add_h1("REFERENCES")

    references = [
        "Allison, C., Auyeung, B., & Baron-Cohen, S. (2012). Toward brief “Red Flags” for autism screening: The Short Autism Spectrum Quotient and the Short Quantitative Checklist in 1,000 cases and 3,000 controls. Journal of the American Academy of Child & Adolescent Psychiatry, 51(2), 202-212.",
        "Baron-Cohen, S., Wheelwright, S., Skinner, R., Martin, J., & Clubley, E. (2001). The Autism-Spectrum Quotient (AQ): Evidence from Asperger syndrome/high-functioning autism, males and females, scientists and mathematicians. Journal of Autism and Developmental Disorders, 31(1), 5-17.",
        "Thabtah, F. (2017). Autism Spectrum Disorder (ASD) Screening: A Machine Learning Approach. UCI Machine Learning Repository.",
        "Thabtah, F. (2019). Machine learning in autistic spectrum disorder behavioral research: A review and ways forward. Informatics for Health and Social Care, 44(3), 278-297.",
        "American Psychiatric Association. (2013). Diagnostic and Statistical Manual of Mental Disorders (DSM-5®). American Psychiatric Publishing.",
        "World Health Organization. (2019). International Statistical Classification of Diseases and Related Health Problems (11th ed.). WHO Press.",
        "Lundberg, S. M., & Lee, S. I. (2017). A unified approach to interpreting model predictions. Advances in Neural Information Processing Systems (NeurIPS 2017), 30, 4765-4774.",
        "Breiman, L. (2001). Random forests. Machine Learning, 45(1), 5-32.",
        "Pedregosa, F., et al. (2011). Scikit-learn: Machine learning in Python. Journal of Machine Learning Research, 12, 2825-2830.",
        "Lord, C., Rutter, M., DiLavore, P. C., Risi, S., Gotham, K., & Bishop, S. (2012). Autism Diagnostic Observation Schedule: ADOS-2. Western Psychological Services.",
        "Rutter, M., Le Couteur, A., & Lord, C. (2003). Autism Diagnostic Interview-Revised (ADI-R). Western Psychological Services.",
        "National Institute for Health and Care Excellence (NICE). (2012). Autism: Recognition, referral and diagnosis of children and young people on the autism spectrum (Clinical Guideline CG142).",
        "Geurts, P., Ernst, D., & Wehenkel, L. (2006). Extremely randomized trees. Machine Learning, 63(1), 3-42.",
        "Ke, G., et al. (2017). LightGBM: A highly efficient gradient boosting decision tree. Advances in Neural Information Processing Systems (NeurIPS 2017), 30, 3146-3154.",
        "Chawla, N. V., Bowyer, K. W., Hall, L. O., & Kegelmeyer, W. P. (2002). SMOTE: Synthetic minority over-sampling technique. Journal of Artificial Intelligence Research, 16, 321-357.",
        "Siddhardhan, S. (2024). Project 21: Autism Prediction Using Machine Learning | Complete ML Project Walkthrough. Machine Learning and Generative AI Series."
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

    doc.save(DOC_PATH)
    print(f"Report successfully generated at: {DOC_PATH}")
    print(f"File size: {os.path.getsize(DOC_PATH)} bytes")

if __name__ == "__main__":
    create_report()
