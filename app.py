import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os
import re

try:
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

# ------------------------------------------------------------
# CONFIGURATION
# ------------------------------------------------------------
DEPARTMENTS = [
    "Administration and Human Resources (1001)",
    "Finance and Accounts Unit (1002)",
    "Policy and Planning Division (1003)",
    "Internal Audit Unit (1004)",
    "Legal Services Unit (1005)",
    "Government Communication Unit (1006)",
    "Procurement Management Unit (1007)",
    "Monitoring and Evaluation Unit (1008)",
    "Management Information Systems Unit (1009)",
    "Minerals Division (2001)"
]

# Financial Year display (July - June)
YEAR_DISPLAY = [
    "2025/2026", "2026/2027", "2027/2028", "2028/2029",
    "2029/2030", "2030/2031"
]
YEAR_MAP = {disp: disp.split("/")[0] for disp in YEAR_DISPLAY}

TFRS_GROUPS = {
    "Standard Submission": "Narrative, KPIs, Achievements, Challenges, Risks, Audit",
    "Organisational Profile & Strategy": "Mandate, Vision, Structure, Performance Overview",
    "Sector Performance & KPIs": "GDP, Production, Revenue, ASM, ESG",
    "Financial Governance": "Revenue, Budget, Liquidity, Financial Statements",
    "Infrastructure & Procurement": "Projects, Vehicles, Contracts, Tenders",
    "ICT & Digital Transformation": "Systems, Security, Data, Digital Projects",
    "Governance, Risk & Compliance": "Legal, Ethics, Litigation, Internal Controls",
    "HR & Capacity Building": "Staffing, Training, Welfare, Integrity"
}

# ------------------------------------------------------------
# THE FULL QUESTION LIST – ALL DEPARTMENTS COVERED
# Format: (Department, TFRS_Group, Question, Evidence/Guidance)
# ------------------------------------------------------------
QUESTION_LIST = []

# --- 1. STANDARD SUBMISSION PACKAGE (Section 5 of the matrix) ---
# Applies to EVERY department
for dept in DEPARTMENTS:
    for q_text, g_text in [
        ("Provide a short narrative summary of achievements, challenges, and priorities for the year.",
         "Submit: Narrative summary covering key events, successes, major obstacles, and forward-looking priorities."),
        ("Provide the Target vs Actual Performance table with variance and status.",
         "Submit: KPI/Activity table with columns: Target, Actual, Variance, Status."),
        ("List your key quantified achievements linked to supporting evidence.",
         "Submit: A list of achievements with evidence references."),
        ("Identify major challenges, their impact, mitigation actions taken, and recommendations.",
         "Submit: Challenge description, impact, mitigation, recommendations."),
        ("Provide budget/financial implications: budget, expenditure, commitments, pending items.",
         "Submit: Budget allocation, actual expenditure, commitments, pending bills, and variance explanations."),
        ("Report on audit findings, legal/compliance matters, and corrective actions taken.",
         "Submit: Audit issues, legal/compliance breaches, status of corrective actions.")
    ]:
        QUESTION_LIST.append((dept, "Standard Submission", q_text, g_text))

# --- 2. ADMINISTRATION AND HUMAN RESOURCES (1001) ---
hr_qs = [
    ("Organisational Profile & Strategy",
     "Submit the approved organisational profile, mandate, vision, mission, and structure.",
     "Submit: Approved organisation structure, establishment records, circulars, official mandate documents."),
    ("HR & Capacity Building",
     "Submit detailed HR data: staff establishment, filled/vacant posts, recruitment, transfers, retirements, deaths, promotions, gender profile, training, capacity gaps, welfare matters, staff claims, disciplinary matters, HIV/AIDS and NCD initiatives.",
     "Submit: HR registers, approved establishment, training reports, PEPMIS reports, minutes, attendance sheets, welfare records, HoD certification."),
    ("HR & Capacity Building",
     "Report on the risk of 'Shortage of skilled staff' including vacancy analysis, training plan, recruitment needs, and capacity gaps.",
     "Submit: Vacancy analysis, training plan, recruitment needs, capacity gap assessment."),
    ("Governance, Risk & Compliance",
     "Submit management meetings held, key decisions made, implementation progress, and pending decisions.",
     "Submit: Meeting minutes, attendance registers, action trackers."),
    ("Governance, Risk & Compliance",
     "Report on integrity, ethics, and compliance: Integrity initiatives, disciplinary matters, ethics compliance, fraud-prevention controls, and staff awareness activities.",
     "Submit: Integrity committee records, training reports, disciplinary registers, compliance reports.")
]
for group, q, g in hr_qs:
    QUESTION_LIST.append(("Administration and Human Resources (1001)", group, q, g))

# --- 3. FINANCE AND ACCOUNTS UNIT (1002) ---
fin_qs = [
    ("Financial Governance",
     "Submit revenue collection data: targets, actual revenue, collection efficiency, variance explanations, arrears, and reconciliation with official systems.",
     "Submit: MUSE/IFMS reports, revenue collection schedules, bank/treasury confirmations, reconciliation reports."),
    ("Financial Governance",
     "Submit the budget execution and variance analysis: releases, expenditure, commitments, payables, receivables, deposits, imprests, assets, WIP, and liabilities.",
     "Submit: MUSE/IFMS reports, ledgers, bank confirmations, reconciliation statements, commitment register, pending bills."),
    ("Financial Governance",
     "Submit financial statements and notes: Statement inputs, notes, disclosures, comparative figures, commitments, pending bills, assets, WIP, liabilities, and confirmations.",
     "Submit: Financial statements, schedules, ledgers, confirmations, asset register, commitment register."),
    ("Financial Governance",
     "Submit the financial control environment: IPSAS compliance, reconciliations, financial risks, and financial reporting status.",
     "Submit: Financial statements, ledgers, reconciliations, audit files, IPSAS disclosure schedules."),
    ("Financial Governance",
     "Submit CAG Audit Issues and Responses: Audit issues, management responses, corrective actions, implementation status, and evidence for closure.",
     "Submit: CAG report, management letter, audit action plan, supporting closure evidence."),
    ("Financial Governance",
     "Report on the risk of 'Unreliable or untimely financial statements': Reconciliation status, reporting timetable, responsible officers, and audit trail.",
     "Submit: Reconciliation status, reporting timetable, responsible officers, audit trail.")
]
for group, q, g in fin_qs:
    QUESTION_LIST.append(("Finance and Accounts Unit (1002)", group, q, g))

# --- 4. POLICY AND PLANNING DIVISION (1003) ---
plan_qs = [
    ("Organisational Profile & Strategy",
     "Submit the Minister's statement inputs: Sector priorities, strategic direction, major achievements, challenges, reforms, and forward-looking commitments.",
     "Submit: Approved annual performance reports, sector statistics, speech inputs, communication clearance."),
    ("Organisational Profile & Strategy",
     "Submit the Permanent Secretary's statement: Institutional performance, implementation of plans, revenue performance, infrastructure, reforms, HR capacity, controls, and priorities.",
     "Submit: Consolidated performance report, management meeting records, certified inputs from all units."),
    ("Organisational Profile & Strategy",
     "Submit strategy and performance overview: Strategic plan implementation, MTEF performance, annual plan results, KPIs, target vs actual, and variance explanations.",
     "Submit: Strategic plan, annual plan, MTEF, quarterly implementation reports, KPI validation records."),
    ("Organisational Profile & Strategy",
     "Submit strategic plan implementation status: Objectives implemented, activities completed, outputs delivered, delayed activities, and reasons.",
     "Submit: Annual plan implementation report, performance dashboard, M&E validation notes."),
    ("Organisational Profile & Strategy",
     "Submit the governance and oversight report: Governance structures, committees, management decisions, implementation status, and unresolved governance issues.",
     "Submit: Minutes, attendance sheets, decision registers, implementation follow-up reports."),
    ("Governance, Risk & Compliance",
     "Submit the risk management framework: Risk register, risk ratings, controls, mitigation actions, residual risks, and unresolved high-risk items.",
     "Submit: Risk register, action plans, M&E/risk review minutes, supporting evidence."),
    ("Governance, Risk & Compliance",
     "Report on the risk of 'Outdated or incomplete risk register' with updated risk register, ratings, controls, residual risk, and action owners.",
     "Submit: Updated risk register with ratings, controls, residual risk, and action owners."),
    ("Governance, Risk & Compliance",
     "Report on the risk of 'Inaccurate budgeting or weak budget execution': Budget variance analysis, reallocation records, and planning corrective action.",
     "Submit: Budget variance analysis, reallocation records, planning corrective action.")
]
for group, q, g in plan_qs:
    QUESTION_LIST.append(("Policy and Planning Division (1003)", group, q, g))

# --- 5. INTERNAL AUDIT UNIT (1004) ---
audit_qs = [
    ("Governance, Risk & Compliance",
     "Submit the annual audit plan and charter: Scope coverage, number of planned audits vs completed.",
     "Submit: Audit plan, audit reports, audit committee minutes."),
    ("Governance, Risk & Compliance",
     "Submit key audit findings, recommendations, management responses, and implementation status.",
     "Submit: Audit reports, management responses, action tracker, verification evidence."),
    ("Governance, Risk & Compliance",
     "Submit internal controls and audit committee matters: Internal control findings, audit committee meetings, recommendations, and management responses.",
     "Submit: Audit reports, audit committee minutes, management responses, follow-up tracker."),
    ("Governance, Risk & Compliance",
     "Report on the risk of 'Internal audit independence or control weaknesses' including internal audit charter status, audit committee minutes, and management responses.",
     "Submit: Internal audit charter, audit committee minutes, management responses.")
]
for group, q, g in audit_qs:
    QUESTION_LIST.append(("Internal Audit Unit (1004)", group, q, g))

# --- 6. LEGAL SERVICES UNIT (1005) ---
legal_qs = [
    ("Governance, Risk & Compliance",
     "Submit legal and regulatory reforms: Mining law amendments, regulations, directives, legal compliance, legal opinions, and regulatory implementation status.",
     "Submit: Acts, regulations, legal opinions, circulars, gazette notices, implementation reports."),
    ("Governance, Risk & Compliance",
     "Submit litigation, claims, and contingent liabilities: Court cases, claims, probability of loss, provisions, contingent liabilities, legal commitments, and subsequent events.",
     "Submit: Litigation register, case files, legal opinions, claims schedules, court documents."),
    ("Governance, Risk & Compliance",
     "Report on the risk of 'Poor enforcement of mining regulations' including legal reforms, enforcement records, compliance reports, and awareness activities.",
     "Submit: Legal reforms, enforcement records, compliance reports, awareness activities.")
]
for group, q, g in legal_qs:
    QUESTION_LIST.append(("Legal Services Unit (1005)", group, q, g))

# --- 7. GOVERNMENT COMMUNICATION UNIT (1006) ---
comm_qs = [
    ("Organisational Profile & Strategy",
     "Submit transparency, public disclosure, and stakeholder communication: Media engagement, public disclosure, stakeholder communication, publications, social media/website analytics, and feedback.",
     "Submit: Media reports, website analytics, press releases, publications, stakeholder meeting records."),
    ("Organisational Profile & Strategy",
     "Report on the risk of 'Poor communication and stakeholder tension' including communication plan, stakeholder engagement log, and feedback handling.",
     "Submit: Communication plan, stakeholder engagement log, feedback handling records.")
]
for group, q, g in comm_qs:
    QUESTION_LIST.append(("Government Communication Unit (1006)", group, q, g))

# --- 8. PROCUREMENT MANAGEMENT UNIT (1007) ---
proc_qs = [
    ("Infrastructure & Procurement",
     "Submit infrastructure projects: Major construction, office infrastructure, contract value, physical progress, financial progress, delays, and completion status.",
     "Submit: Contracts, IPCs, completion certificates, site reports, procurement files, payment records."),
    ("Infrastructure & Procurement",
     "Submit procurement of vehicles, equipment, contracts, and LPOs: Procurement plan implementation, tenders, contracts, LPOs, framework contracts, variations, guarantees, and delays.",
     "Submit: APP, NeST reports, tender board minutes, contract files, LPOs, guarantees, delivery notes."),
    ("Infrastructure & Procurement",
     "Submit Ministerial Tender Board matters: Tender board meetings, approvals, procurement decisions, contract awards, pending procurements, and procurement risks.",
     "Submit: Tender board minutes, evaluation reports, approvals, contract award notices."),
    ("Governance, Risk & Compliance",
     "Report on the risk of 'Leakage of tender information' including tender controls, confidentiality declarations, board minutes, and investigation reports.",
     "Submit: Tender controls, confidentiality declarations, board minutes, investigation reports."),
    ("Governance, Risk & Compliance",
     "Report on the risk of 'Delays in procurement' including delayed procurements, causes, impact, and mitigation plan.",
     "Submit: Delayed procurements, causes, impact, mitigation plan.")
]
for group, q, g in proc_qs:
    QUESTION_LIST.append(("Procurement Management Unit (1007)", group, q, g))

# --- 9. MONITORING AND EVALUATION UNIT (1008) ---
me_qs = [
    ("Sector Performance & KPIs",
     "Submit KPI validation, project monitoring, data quality, and performance evidence.",
     "Submit: M&E reports, field visit reports, data validation sheets, performance dashboards, evaluation reports, evidence validation notes."),
    ("Sector Performance & KPIs",
     "Report on the risk of 'Poor statistical data integrity' including data validation procedures, quality checks, methodology notes, and sign-offs.",
     "Submit: Data validation procedures, quality checks, methodology notes, sign-offs.")
]
for group, q, g in me_qs:
    QUESTION_LIST.append(("Monitoring and Evaluation Unit (1008)", group, q, g))

# --- 10. MANAGEMENT INFORMATION SYSTEMS UNIT (1009) ---
ict_qs = [
    ("ICT & Digital Transformation",
     "Submit ICT systems, digital transformation, and data security: Systems used, uptime, users supported, incidents, backups, cybersecurity, licences, digital projects, and data governance.",
     "Submit: System inventory, incident logs, uptime reports, backup reports, licence register, user support logs, DR plan, ICT project reports."),
    ("ICT & Digital Transformation",
     "Submit Ministerial ICT Steering Committee matters: ICT governance decisions, ICT projects, system risks, digital transformation, and implementation status.",
     "Submit: Committee minutes, project reports, ICT risk reports, implementation trackers."),
    ("ICT & Digital Transformation",
     "Report on the risk of 'Loss of data, downtime or system failure' including backup status, uptime, incident reports, DR arrangements, and cybersecurity controls.",
     "Submit: Backup status, uptime, incident reports, DR arrangements, cybersecurity controls.")
]
for group, q, g in ict_qs:
    QUESTION_LIST.append(("Management Information Systems Unit (1009)", group, q, g))

# --- 11. MINERALS DIVISION (2001) ---
min_qs = [
    ("Sector Performance & KPIs",
     "Submit sector KPIs: GDP contribution, mineral production, mineral buying centres, trained miners, production statistics, training data, buying centres, licences, market performance, and validation.",
     "Submit: Sector reports, Mining Commission/official records, M&E sign-off, data methodology notes."),
    ("Sector Performance & KPIs",
     "Submit small-scale mining and ASM development: ASM licences, training, loans, formalisation, safety, compliance, markets, and productivity support.",
     "Submit: Training reports, licence schedules, loan schedules, attendance sheets, inspection reports, photos."),
    ("Sector Performance & KPIs",
     "Submit sustainability, ESG, and climate-related matters: Environmental management, rehabilitation, ASM safety, CSR, local content, climate-related risks, and financial implications.",
     "Submit: Inspection reports, ESG schedules, CSR reports, environmental compliance records, risk register, financial schedules."),
    ("Sector Performance & KPIs",
     "Submit detailed mineral sector performance: Mining sector GDP contribution, mineral production (tonnages/value), mineral revenue, licences, mineral markets, ASM support, local content, CSR, value addition, strategic minerals, environmental management, rehabilitation, mining inspections, illegal mining, ESG/climate data.",
     "Submit: Sector reports, inspection reports, licence schedules, mineral production records, CSR reports, local content reports, environmental compliance records, data methodology note."),
    ("Sector Performance & KPIs",
     "Report on the risk of 'Weak inter-agency coordination' including coordination meetings, stakeholder engagement records, and action plans.",
     "Submit: Coordination meetings, stakeholder engagement records, action plans."),
    ("Sector Performance & KPIs",
     "Report on the risk of 'Informal/illegal mining operations' including inspection reports, enforcement actions, stakeholder awareness, and legal updates.",
     "Submit: Inspection reports, enforcement actions, stakeholder awareness, legal updates."),
    ("Sector Performance & KPIs",
     "Report on the risk of 'Incomplete revenue data from mineral rights holders' including reconciled revenue and production data, data gaps, and corrective action.",
     "Submit: Reconciled revenue and production data, data gaps, corrective action."),
    ("Sector Performance & KPIs",
     "Report on the risk of 'Mineral smuggling and revenue leakage' including control measures, enforcement reports, and risk mitigation actions.",
     "Submit: Control measures, enforcement reports, risk mitigation actions."),
    ("Sector Performance & KPIs",
     "Report on the risk of 'Environmental pollution and mine rehabilitation issues' including inspection reports, rehabilitation plans, compliance letters, and ESG schedules.",
     "Submit: Inspection reports, rehabilitation plans, compliance letters, ESG schedules."),
    ("Sector Performance & KPIs",
     "Report on the risk of 'Low ASM productivity' including training, loans, market access, safety support, and productivity indicators.",
     "Submit: Training, loans, market access, safety support, productivity indicators.")
]
for group, q, g in min_qs:
    QUESTION_LIST.append(("Minerals Division (2001)", group, q, g))

# Build the dictionary used by the app
DEPARTMENT_QUESTIONS = {}
for dept, group, q, g in QUESTION_LIST:
    DEPARTMENT_QUESTIONS.setdefault(dept, []).append(
        {"group": group, "question": q, "guidance": g}
    )

# ------------------------------------------------------------
# FILE HELPERS
# ------------------------------------------------------------
DATA_FILE = "tfrs_data"
SYNTHESIS_FILE = "synthesis_data"
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

def count_words(text):
    if not text or not isinstance(text, str):
        return 0
    return len(re.findall(r'\b\w+\b', text))

def load_data(year_key):
    file_key = f"{DATA_FILE}_{year_key}.csv"
    try:
        if os.path.exists(file_key):
            df = pd.read_csv(file_key)
            required = ['Department', 'Group', 'Question', 'Guidance',
                        'Comments', 'Narrative', 'Attachments', 'Year', 'Last_Updated']
            for col in required:
                if col not in df.columns:
                    df[col] = ''
            df['Year'] = df.get('Year', year_key)
            return df
        else:
            return create_new_data(year_key)
    except Exception:
        if os.path.exists(file_key):
            os.remove(file_key)
        return create_new_data(year_key)

def create_new_data(year_key):
    rows = []
    for dept, questions in DEPARTMENT_QUESTIONS.items():
        for q in questions:
            rows.append({
                "Department": dept,
                "Group": q["group"],
                "Question": q["question"],
                "Guidance": q["guidance"],
                "Comments": "",
                "Narrative": "",
                "Attachments": "",
                "Year": year_key,
                "Last_Updated": datetime.now().strftime("%Y-%m-%d %H:%M")
            })
    df = pd.DataFrame(rows)
    df.to_csv(f"{DATA_FILE}_{year_key}.csv", index=False)
    return df

def save_data(df, year_key):
    df.to_csv(f"{DATA_FILE}_{year_key}.csv", index=False)

def load_synthesis(year_key):
    file_key = f"{SYNTHESIS_FILE}_{year_key}.csv"
    try:
        if os.path.exists(file_key):
            df = pd.read_csv(file_key)
            if 'Group' not in df.columns or 'Synthesis' not in df.columns:
                os.remove(file_key)
                raise Exception("Bad format")
            for group in TFRS_GROUPS.keys():
                if group not in df['Group'].values:
                    df = pd.concat([df, pd.DataFrame([{'Group': group, 'Synthesis': ''}])], ignore_index=True)
            return df[['Group', 'Synthesis']]
        else:
            rows = [{'Group': group, 'Synthesis': ''} for group in TFRS_GROUPS.keys()]
            df = pd.DataFrame(rows)
            df.to_csv(file_key, index=False)
            return df
    except Exception:
        if os.path.exists(file_key):
            os.remove(file_key)
        rows = [{'Group': group, 'Synthesis': ''} for group in TFRS_GROUPS.keys()]
        df = pd.DataFrame(rows)
        df.to_csv(file_key, index=False)
        return df

def save_synthesis(df, year_key):
    df = df[['Group', 'Synthesis']]
    df.to_csv(f"{SYNTHESIS_FILE}_{year_key}.csv", index=False)

# ------------------------------------------------------------
# STREAMLIT APP
# ------------------------------------------------------------
st.set_page_config(layout="wide")
st.title("🏛 TCWG & TFRS 1 Report Builder – Ministry of Minerals")
st.caption("Based on the Official TCWG Report Data Responsibility Matrix (FY 2025/2026)")

try:
    # Sidebar
    st.sidebar.title("📋 Reporting Period")
    selected_display = st.sidebar.selectbox("Select Financial Year", YEAR_DISPLAY)
    selected_year = YEAR_MAP[selected_display]

    st.sidebar.markdown("---")
    st.sidebar.title("📌 Division/Unit")
    selected_dept = st.sidebar.selectbox("Select Your Division/Unit",
                                         DEPARTMENTS)

    st.sidebar.markdown("---")
    st.sidebar.caption("📌 Write at least **100 words** per narrative.")
    st.sidebar.caption("📌 Attach supporting evidence (PDF, Excel, Images).")

    # Load data
    if 'data' not in st.session_state or st.session_state.get('current_year') != selected_year:
        st.session_state.data = load_data(selected_year)
        st.session_state.current_year = selected_year
        st.session_state.synthesis = load_synthesis(selected_year)

    if 'uploaded_files' not in st.session_state:
        st.session_state.uploaded_files = {}

    # Admin: Add Question
    with st.sidebar.expander("🔧 Admin - Add New Question (Password: admin123)"):
        admin_pass = st.text_input("Password", type="password", key="admin_pass")
        if admin_pass == "admin123":
            st.success("Admin access granted.")
            new_dept = st.selectbox("Department", DEPARTMENTS)
            new_group = st.selectbox("TFRS Group", list(TFRS_GROUPS.keys()))
            new_q = st.text_area("Question")
            new_guidance = st.text_area("Evidence Required")
            if st.button("➕ Add Question"):
                if new_q and new_guidance:
                    new_row = {
                        "Department": new_dept,
                        "Group": new_group,
                        "Question": new_q,
                        "Guidance": new_guidance,
                        "Comments": "",
                        "Narrative": "",
                        "Attachments": "",
                        "Year": selected_year,
                        "Last_Updated": datetime.now().strftime("%Y-%m-%d %H:%M")
                    }
                    st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame([new_row])], ignore_index=True)
                    save_data(st.session_state.data, selected_year)
                    st.success("✅ Added!")
                    st.rerun()
        elif admin_pass:
            st.error("Wrong password.")

    st.sidebar.markdown("---")
    st.sidebar.caption(f"Financial Year {selected_display}")

    # ---------- DATA ENTRY ----------
    st.subheader(f"📋 Data Request: {selected_dept}")
    st.caption(f"**Reporting Period:** Financial Year {selected_display}")
    st.caption("Minimum 100 words per narrative. Attach evidence for every submission.")

    dept_mask = st.session_state.data["Department"] == selected_dept
    dept_data = st.session_state.data[dept_mask].copy()

    if dept_data.empty:
        st.warning(f"No questions found for {selected_dept}.")
        st.stop()

    with st.form(key="entry_form"):
        updated_narratives = []
        updated_comments = []
        updated_attachments = []

        q_counter = 1
        for idx, row in dept_data.iterrows():
            question = row["Question"]
            guidance = row["Guidance"]
            current_narrative = row["Narrative"]
            current_comment = row["Comments"]
            current_attachments = row["Attachments"]

            with st.expander(f"Q{q_counter}: {question}", expanded=False):
                col1, col2 = st.columns([2, 1])
                with col1:
                    st.caption(f"📎 Evidence Required: {guidance}")
                    narrative = st.text_area(
                        "Narrative / Data Submission",
                        value=current_narrative,
                        key=f"narrative_{idx}",
                        height=150,
                        placeholder="Provide detailed data, numbers, and analysis as requested..."
                    )
                    comment = st.text_input(
                        "Additional Comments / References",
                        value=current_comment,
                        key=f"comment_{idx}",
                        placeholder="e.g., see page 5 of the official report"
                    )
                with col2:
                    wc = count_words(narrative)
                    if wc >= 100:
                        st.success(f"✅ {wc} words - Detailed!")
                    else:
                        st.warning(f"⚠️ {wc} / 100 words")
                        st.progress(wc / 100)

                    uploaded_file = st.file_uploader(
                        "Attach Evidence",
                        type=["pdf", "xlsx", "xls", "docx", "png", "jpg", "jpeg"],
                        key=f"upload_{idx}",
                        label_visibility="collapsed"
                    )
                    if uploaded_file is not None:
                        ext = uploaded_file.name.split('.')[-1]
                        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                        safe_name = f"{selected_dept.replace(' ', '_')}_{ts}.{ext}"
                        file_path = os.path.join(UPLOAD_DIR, safe_name)
                        with open(file_path, "wb") as f:
                            f.write(uploaded_file.getbuffer())
                        st.success(f"✅ Evidence saved: {safe_name}")
                        current_attachments = safe_name if not current_attachments else current_attachments + ", " + safe_name
                    if current_attachments:
                        st.info(f"📎 Attached: {current_attachments}")

                updated_narratives.append(narrative)
                updated_comments.append(comment)
                updated_attachments.append(current_attachments)

            q_counter += 1

        submitted = st.form_submit_button("💾 Submit Division/Unit Data")
        if submitted:
            for i, (idx, row) in enumerate(dept_data.iterrows()):
                st.session_state.data.at[idx, "Narrative"] = updated_narratives[i]
                st.session_state.data.at[idx, "Comments"] = updated_comments[i]
                st.session_state.data.at[idx, "Attachments"] = updated_attachments[i]
                st.session_state.data.at[idx, "Last_Updated"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                st.session_state.data.at[idx, "Year"] = selected_year
            save_data(st.session_state.data, selected_year)
            st.success("✅ Division/Unit data submitted successfully!")

    # ---------- DASHBOARD ----------
    st.markdown("---")
    st.header("📊 Live Submission Dashboard")
    st.caption(f"**Reporting Period:** Financial Year {selected_display}")
    st.info("💡 **Compliance %** = (questions with ≥100 words) / (total questions).")

    all_depts = st.session_state.data["Department"].unique()
    summary = []
    for dept in all_depts:
        dept_filter = st.session_state.data[st.session_state.data["Department"] == dept]
        total = len(dept_filter)
        compliant = sum(1 for _, r in dept_filter.iterrows() if count_words(r["Narrative"]) >= 100)
        compliance = (compliant / total * 100) if total > 0 else 0
        summary.append({
            "Division/Unit": dept,
            "Total Questions": total,
            "✅ Submitted (≥100 words)": compliant,
            "❌ Pending/Incomplete": total - compliant,
            "Compliance %": round(compliance, 1)
        })
    df_summary = pd.DataFrame(summary)
    st.dataframe(df_summary, width='stretch', hide_index=True)
    st.subheader("Overall Submission Progress by Division/Unit (%)")
    st.bar_chart(df_summary.set_index("Division/Unit")["Compliance %"])

    # ---------- ADMIN SYNTHESIS ----------
    st.markdown("---")
    st.header("🔧 Admin Consolidation / Synthesis")

    with st.expander("📝 Admin: Synthesize Division Inputs into TCWG Report (Password: admin123)"):
        synth_pass = st.text_input("Enter Admin Password", type="password", key="synth_pass")
        if synth_pass == "admin123":
            st.success("✅ Admin access granted.")
            st.info("Rewrite raw division inputs into one cohesive paragraph per TCWG Report Area.")

            st.subheader("📄 Raw Division/Unit Inputs")
            for group in TFRS_GROUPS.keys():
                group_data = st.session_state.data[
                    st.session_state.data["Group"] == group
                ]
                if not group_data.empty:
                    with st.expander(f"View Raw Data for {group}"):
                        for dept in DEPARTMENTS:
                            dept_group = group_data[group_data["Department"] == dept]
                            if not dept_group.empty:
                                st.markdown(f"**{dept}:**")
                                for _, row in dept_group.iterrows():
                                    st.caption(f"Q: {row['Question']}")
                                    st.write(row['Narrative'] if row['Narrative'] else "⚠ No data provided.")
                                    st.caption(f"Word count: {count_words(row['Narrative'])}")
                                st.divider()

            st.subheader("✍️ Write Your Polished Synthesis Paragraphs")
            synth_df = st.session_state.synthesis.copy()
            updated = {}
            for group in TFRS_GROUPS.keys():
                current = synth_df[synth_df['Group'] == group]['Synthesis'].values[0] if not synth_df[synth_df['Group'] == group].empty else ""
                new_text = st.text_area(
                    f"Synthesis for {group}",
                    value=current,
                    height=120,
                    key=f"synth_{group}",
                    placeholder="Write the final report narrative for this section..."
                )
                updated[group] = new_text

            if st.button("💾 Save All Synthesis Text"):
                for group, text in updated.items():
                    idx = st.session_state.synthesis[st.session_state.synthesis['Group'] == group].index
                    if not idx.empty:
                        st.session_state.synthesis.at[idx[0], 'Synthesis'] = text
                    else:
                        st.session_state.synthesis = pd.concat([st.session_state.synthesis, pd.DataFrame([{'Group': group, 'Synthesis': text}])], ignore_index=True)
                save_synthesis(st.session_state.synthesis, selected_year)
                st.success(f"✅ Synthesis saved for FY {selected_display}!")
                st.rerun()
        elif synth_pass:
            st.error("Wrong password.")

    # ---------- REPORT GENERATOR ----------
    st.markdown("---")
    st.header("📄 Generate Draft TCWG & TFRS 1 Report")
    st.warning(f"Report for Financial Year {selected_display}.")

    if st.button("📝 Generate Consolidated Draft Report (Word)"):
        output = io.BytesIO()
        if DOCX_AVAILABLE:
            doc = Document()
            doc.add_heading('REPORT BY THOSE CHARGED WITH GOVERNANCE', 0)
            doc.add_heading('Ministry of Minerals', level=1)
            doc.add_paragraph(f'Reporting Period: Financial Year {selected_display}')
            doc.add_paragraph(f'Date: {datetime.now().strftime("%d %B %Y")}')
            doc.add_paragraph('Prepared in accordance with TFRS 1 and TCWG Framework')
            doc.add_paragraph('')

            doc.add_heading('1. Executive Summary', level=2)
            for _, row in df_summary.iterrows():
                doc.add_paragraph(f'• {row["Division/Unit"]}: {row["Compliance %"]}% submission completeness ({row["✅ Submitted (≥100 words)"]} out of {row["Total Questions"]})')
            doc.add_paragraph('')

            doc.add_heading('2. Consolidated Report by TCWG/TFRS Section', level=2)
            synth_df = st.session_state.synthesis
            for group in TFRS_GROUPS.keys():
                doc.add_heading(f'2.{list(TFRS_GROUPS.keys()).index(group)+1} {group}', level=3)
                doc.add_paragraph(TFRS_GROUPS[group])
                synth_row = synth_df[synth_df['Group'] == group]
                if not synth_row.empty and synth_row.iloc[0]['Synthesis'] and synth_row.iloc[0]['Synthesis'].strip():
                    doc.add_paragraph(synth_row.iloc[0]['Synthesis'], style='List Bullet')
                    doc.add_paragraph("*(Consolidated by Management.)*", style='List Bullet')
                else:
                    group_data = st.session_state.data[
                        st.session_state.data["Group"] == group
                    ]
                    if group_data.empty:
                        doc.add_paragraph('⚠ No data provided for this section.')
                    else:
                        for dept in DEPARTMENTS:
                            dept_group = group_data[group_data["Department"] == dept]
                            if dept_group.empty:
                                continue
                            doc.add_heading(f'{dept}:', level=4)
                            for _, row in dept_group.iterrows():
                                doc.add_paragraph(f'• {row["Question"]}')
                                if row["Narrative"]:
                                    doc.add_paragraph(row["Narrative"], style='List Bullet')
                                else:
                                    doc.add_paragraph('⚠ Not provided.', style='List Bullet')
                                if row["Attachments"]:
                                    doc.add_paragraph(f'Evidence: {row["Attachments"]}', style='List Bullet')
                doc.add_paragraph('')

            doc.save(output)
            mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ext = "docx"
        else:
            output.write("="*80 + "\n".encode())
            output.write("TCWG REPORT - MINISTRY OF MINERALS\n".encode())
            output.write(f"Financial Year {selected_display}\n".encode())
            output.write("="*80 + "\n\n".encode())
            output.write("(Install python-docx for a proper Word document.)\n".encode())
            mime = "text/plain"
            ext = "txt"

        st.download_button(
            label=f"⬇ Download Draft Report (.{ext})",
            data=output.getvalue(),
            file_name=f"TCWG_Report_FY{selected_display.replace('/','_')}_{datetime.now().strftime('%Y%m%d')}.{ext}",
            mime=mime
        )
        st.success(f"Report generated for FY {selected_display}!")

except Exception as e:
    st.error(f"An error occurred: {e}")
    st.stop()
