import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os
import base64

# Try to import docx for Word report generation
try:
    from docx import Document
    from docx.shared import Inches
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False
    st.warning("python-docx not installed. Report will be generated as plain text. Run 'pip install python-docx' for Word export.")

# ---------- CONFIG ----------
DEPARTMENTS = [
    "Human Resource",
    "Monitoring & Evaluation",
    "Information & Communication Technology",
    "Government Communication Unit",
    "Policy & Planning",
    "Procurement",
    "Commissioner for Minerals",
    "Finance & Accounts",
    "Internal Auditing"
]

# Define TFRS Groups for consistent grouping (USED ONLY FOR THE FINAL REPORT)
TFRS_GROUPS = {
    "A. Nature of Operation": "Description of the industry, markets, products, services, and regulatory environment.",
    "B. Objectives, Strategies & Operating Model": "Entity's strategic objectives, resource allocation plans, and operating model (inputs/outputs/outcomes).",
    "C. Resources": "Financial, manufactured, intellectual, human, social/relationship, and natural resources.",
    "D. Principal Risks & Uncertainties": "Key risks, likelihood, impact, and mitigation strategies.",
    "E. Stakeholder Relationships": "Significant relationships with employees, customers, suppliers, and communities.",
    "F. Capital Structure, Treasury & Liquidity": "Capital structure, treasury policies, cash flows, and liquidity position.",
    "G. Performance & KPIs": "Key performance indicators (financial and non-financial), service delivery targets, and efficiency.",
    "H. Corporate Governance": "Governance structure, committees, membership, ethics, and auditor appointment.",
    "I. Forward-looking & Future Prospects": "Anticipated trends, challenges, and strategic responses for the next 1-3 years.",
    "J. Compliance & Responsibility": "Statement of responsibility, compliance with TFRS 1, publication, and approvals."
}

# ---------- DEPARTMENT-SPECIFIC QUESTIONS (Properly phrased) ----------
DEPARTMENT_QUESTIONS = {
    "Human Resource": [
        {"group": "A. Nature of Operation", "question": "What is the role of HR in supporting the Ministry's operational mandate, and how is it structured?",
         "guidance": "Outline staffing levels, organizational structure, and key HR functions (recruitment, payroll, training)."},
        {"group": "C. Resources", "question": "What are the current staff recruitment, retention, and development strategies, and what are the key metrics?",
         "guidance": "Provide vacancy rates, turnover %, training days per staff, and reference to strategic plans."},
        {"group": "C. Resources", "question": "What employee welfare initiatives (health, pension, safety) are in place, and what is the coverage?",
         "guidance": "Coverage percentages, budget allocated, specific schemes (PPF/NSSF)."},
        {"group": "C. Resources", "question": "How does the department ensure gender parity and equal opportunity in employment and promotion?",
         "guidance": "Current gender ratio (M/F). Recruitment/promotion by gender."},
        {"group": "C. Resources", "question": "What training and career progression opportunities are provided for persons with disabilities?",
         "guidance": "Numbers of disabled staff recruited, trained, promoted."},
        {"group": "E. Stakeholder Relationships", "question": "What grievance handling mechanisms are in place for staff, and what are the resolution statistics?",
         "guidance": "Number of grievances lodged, resolved, average resolution time."},
        {"group": "G. Performance & KPIs", "question": "What are the key HR performance metrics, and what do the trends show?",
         "guidance": "Turnover rate, cost per hire, staff satisfaction score, vacancy rate."},
        {"group": "I. Forward-looking & Future Prospects", "question": "What major HR trends are anticipated in the next 1-3 years, and what is the strategic response?",
         "guidance": "Digital HR, talent shortages, remote work, skills gaps."}
    ],
    "Monitoring & Evaluation": [
        {"group": "A. Nature of Operation", "question": "How is the M&E function structured, and what is its role in tracking Ministry performance?",
         "guidance": "Outline the M&E framework, reporting cycles, and key stakeholders."},
        {"group": "G. Performance & KPIs", "question": "What are the actual service delivery targets (outputs) achieved, and what performance indicators were used?",
         "guidance": "Actual outputs vs. targets (e.g., inspections, reports produced)."},
        {"group": "G. Performance & KPIs", "question": "How efficiently were resources utilized in terms of inputs versus outputs, and what ratios were achieved?",
         "guidance": "Cost per output, productivity ratios, budget absorption rate."},
        {"group": "E. Stakeholder Relationships", "question": "How has stakeholder feedback and citizen satisfaction been incorporated, and what were the scores?",
         "guidance": "Satisfaction scores, feedback response rates, actions taken."},
        {"group": "G. Performance & KPIs", "question": "To what extent have the intended societal impacts (outcomes) been achieved, and what is the evidence?",
         "guidance": "Outcome indicators (e.g., health improvements, economic growth)."},
        {"group": "H. Corporate Governance", "question": "What is the implementation status of previous M&E recommendations (closed, in progress, not implemented)?",
         "guidance": "Number closed, in progress, not implemented."},
        {"group": "G. Performance & KPIs", "question": "Which Key Performance Indicators (KPIs) are tracked consistently, and what are the trends compared to targets?",
         "guidance": "Targets vs. actuals for top 5 KPIs."},
        {"group": "I. Forward-looking & Future Prospects", "question": "What major M&E trends are anticipated in the next 1-3 years, and what is the strategic response?",
         "guidance": "New data collection tools, impact evaluation, real-time reporting."}
    ],
    "Information & Communication Technology": [
        {"group": "A. Nature of Operation", "question": "How is the ICT function structured, and what critical services does it provide to the Ministry?",
         "guidance": "Outline systems supported, user base, and key services (email, networks, databases)."},
        {"group": "D. Principal Risks & Uncertainties", "question": "What are the key cybersecurity risks, data privacy measures, and IT policies in place?",
         "guidance": "Number of incidents, compliance level, policy version."},
        {"group": "B. Objectives, Strategies & Operating Model", "question": "What is the digital transformation strategy, and what progress has been made in implementation?",
         "guidance": "Milestones achieved (systems launched, automation, adoption rates)."},
        {"group": "D. Principal Risks & Uncertainties", "question": "What business continuity and disaster recovery plans are in place, and when were they last tested?",
         "guidance": "Last testing date, Recovery Time Objective (RTO) achieved."},
        {"group": "F. Capital Structure, Treasury & Liquidity", "question": "What are the costs and benefits of major IT investments, and what is the ROI?",
         "guidance": "Total capex/opex, ROI, efficiency gains (e.g., time reduced by X%)."},
        {"group": "C. Resources", "question": "What is the current IT asset inventory, and how is it managed?",
         "guidance": "Total assets, age profile, replacement value, disposal records."},
        {"group": "G. Performance & KPIs", "question": "What is the system downtime record, and how quickly are incidents resolved?",
         "guidance": "Number of outages, average resolution time, root causes."},
        {"group": "I. Forward-looking & Future Prospects", "question": "What major ICT trends are anticipated in the next 1-3 years, and what is the strategic response?",
         "guidance": "Cloud migration, AI, mobile apps, data analytics, remote work support."}
    ],
    "Government Communication Unit": [
        {"group": "A. Nature of Operation", "question": "What is the role of the Government Communication Unit, and what are its core functions?",
         "guidance": "Outline public engagement, media relations, and information dissemination."},
        {"group": "B. Objectives, Strategies & Operating Model", "question": "What communication strategies and public engagement plans are in place, and what is their reach?",
         "guidance": "Campaigns, events, media engagements, reach (estimated audience)."},
        {"group": "G. Performance & KPIs", "question": "How effectively does the unit disseminate information to citizens, and what are the metrics?",
         "guidance": "Number of press releases, website traffic, social media impressions."},
        {"group": "E. Stakeholder Relationships", "question": "What mechanisms are in place for handling citizen feedback/complaints, and how effective are they?",
         "guidance": "Feedback volumes, response rates, resolution times."},
        {"group": "J. Compliance & Responsibility", "question": "How does the unit ensure transparency and access to information in compliance with the law?",
         "guidance": "Number of information requests handled, proactive disclosures."},
        {"group": "G. Performance & KPIs", "question": "What public awareness campaigns have been conducted, and what was their effectiveness?",
         "guidance": "Campaign reach, behavior change results, cost per person reached."},
        {"group": "D. Principal Risks & Uncertainties", "question": "What risks related to misinformation and reputation have been identified, and how are they managed?",
         "guidance": "Number of reputation incidents, media monitoring outcomes."},
        {"group": "I. Forward-looking & Future Prospects", "question": "What major communication trends are anticipated in the next 1-3 years, and what is the strategic response?",
         "guidance": "Digital engagement, fake news, citizen journalism, crisis communication."}
    ],
    "Policy & Planning": [
        {"group": "A. Nature of Operation", "question": "What is the role of the Policy & Planning function, and how does it coordinate with national priorities?",
         "guidance": "Outline policy formulation, planning, and coordination with national priorities."},
        {"group": "B. Objectives, Strategies & Operating Model", "question": "How are the Ministry's strategic objectives aligned with national development plans (e.g., Five-Year Plan)?",
         "guidance": "Map specific goals to national indicators."},
        {"group": "G. Performance & KPIs", "question": "What is the progress of strategic plan implementation, and what are the major milestones achieved?",
         "guidance": "Status of each goal (achieved, ongoing, delayed). Milestones reached."},
        {"group": "A. Nature of Operation", "question": "What legislative and regulatory changes affecting the minerals sector have occurred, and what is their impact?",
         "guidance": "List new laws/amendments and their expected impact on operations."},
        {"group": "B. Objectives, Strategies & Operating Model", "question": "What is the policy development framework, and how are stakeholders consulted?",
         "guidance": "Number of consultations, feedback incorporated."},
        {"group": "D. Principal Risks & Uncertainties", "question": "What risks are associated with policy implementation, and what mitigation measures are in place?",
         "guidance": "Identify top risks and mitigation measures."},
        {"group": "I. Forward-looking & Future Prospects", "question": "What major policy trends are anticipated in the next 1-3 years, and what is the strategic response?",
         "guidance": "New mining codes, environmental regulations, regional integration."}
    ],
    "Procurement": [
        {"group": "A. Nature of Operation", "question": "What is the role of the Procurement function, and what are its critical responsibilities?",
         "guidance": "Outline procurement of goods/services, contract management, and supplier relations."},
        {"group": "B. Objectives, Strategies & Operating Model", "question": "What are the strategic procurement plans and annual budgets, and what is the actual spend?",
         "guidance": "Total planned vs. actual spend. Number of tenders planned vs issued."},
        {"group": "J. Compliance & Responsibility", "question": "What is the level of adherence to the Public Procurement Act regulations, and are there any breaches?",
         "guidance": "Percentage of tenders fully compliant. Any major breaches/penalties?"},
        {"group": "G. Performance & KPIs", "question": "What are the major contracts awarded, and how have suppliers performed?",
         "guidance": "Top 5 contracts by value. Supplier scores (e.g., 4.5/5)."},
        {"group": "D. Principal Risks & Uncertainties", "question": "What risks related to supply chain, fraud, and delays have been identified, and how are they mitigated?",
         "guidance": "List top risks and controls. Number of fraud incidents."},
        {"group": "G. Performance & KPIs", "question": "What value-for-money and cost savings have been achieved through procurement?",
         "guidance": "Estimated savings from competitive bidding vs. estimated price."},
        {"group": "J. Compliance & Responsibility", "question": "How are sustainable and local content procurement policies being implemented?",
         "guidance": "Percentage local content, environmental criteria applied."},
        {"group": "I. Forward-looking & Future Prospects", "question": "What major procurement trends are anticipated in the next 1-3 years, and what is the strategic response?",
         "guidance": "e-Procurement, strategic sourcing, green procurement."}
    ],
    "Commissioner for Minerals": [
        {"group": "A. Nature of Operation", "question": "What is the mandate of the Commissioner for Minerals, and what are the core regulatory functions?",
         "guidance": "Outline licensing, inspection, revenue collection, and sector development."},
        {"group": "G. Performance & KPIs", "question": "What are the mineral production statistics and royalty collections, and how do they compare to targets?",
         "guidance": "Tonnes produced (by mineral), total value, royalties collected vs. target."},
        {"group": "A. Nature of Operation", "question": "What is the status of licensing processes and monitoring of licensed activities?",
         "guidance": "Number of new/renewed licenses, compliance rate."},
        {"group": "G. Performance & KPIs", "question": "What inspection and compliance activities have been carried out, and what were the findings?",
         "guidance": "Number of inspections, findings, enforcement actions."},
        {"group": "J. Compliance & Responsibility", "question": "What environmental protection and rehabilitation activities have been undertaken?",
         "guidance": "Rehabilitation area (ha), number of EIAs, closure plans."},
        {"group": "G. Performance & KPIs", "question": "How do actual revenue collections compare to targets, and what are the reasons for variances?",
         "guidance": "Variance analysis: reasons for over/under collection."},
        {"group": "D. Principal Risks & Uncertainties", "question": "What risks relate to mining compliance, illegal mining, and revenue leakage, and how are they controlled?",
         "guidance": "Number of illegal mining cases, revenue lost, controls."},
        {"group": "I. Forward-looking & Future Prospects", "question": "What major minerals sector trends are anticipated in the next 1-3 years, and what is the strategic response?",
         "guidance": "Global commodity prices, new discoveries, local beneficiation policy."}
    ],
    "Finance & Accounts": [
        {"group": "F. Capital Structure, Treasury & Liquidity", "question": "What is the current cash flow position and liquidity status of the Ministry?",
         "guidance": "Opening/closing cash balances, cash inflow/outflow, liquidity ratio."},
        {"group": "G. Performance & KPIs", "question": "What is the budget absorption rate, and what are the major variances (actual vs planned)?",
         "guidance": "Overall budget execution percentage (>90%?). Explain variances >10%."},
        {"group": "D. Principal Risks & Uncertainties", "question": "What financial risks (currency, inflation, fraud) have been identified, and how are they mitigated?",
         "guidance": "Quantify exposure, risk mitigation measures."},
        {"group": "F. Capital Structure, Treasury & Liquidity", "question": "What treasury policies are in place, and how are public funds managed?",
         "guidance": "Cash management strategies, bank reconciliation status."},
        {"group": "F. Capital Structure, Treasury & Liquidity", "question": "What significant payments, commitments, and outstanding obligations exist?",
         "guidance": "Large contracts pending, bills, committed funds."},
        {"group": "F. Capital Structure, Treasury & Liquidity", "question": "What is the financial impact of major capital projects (cost, funding, expected returns)?",
         "guidance": "Total cost, funding sources, expected returns/benefits."},
        {"group": "I. Forward-looking & Future Prospects", "question": "What major financial trends are anticipated in the next 1-3 years, and what is the strategic response?",
         "guidance": "Budget constraints, funding reforms, PPPs, revenue diversification."}
    ],
    "Internal Auditing": [
        {"group": "H. Corporate Governance", "question": "What is the annual audit plan and charter, and what is the coverage scope?",
         "guidance": "Scope coverage, number of planned audits vs completed."},
        {"group": "H. Corporate Governance", "question": "What are the key audit findings, recommendations, and management responses?",
         "guidance": "List top 5 high-risk findings and whether management accepted them."},
        {"group": "H. Corporate Governance", "question": "What is the implementation status of previous audit recommendations?",
         "guidance": "Number closed, in progress, not implemented. Provide percentages."},
        {"group": "H. Corporate Governance", "question": "How effective is the internal control environment, and what is the assessment?",
         "guidance": "Qualitative opinion on control design and operation."},
        {"group": "J. Compliance & Responsibility", "question": "What is the compliance status with financial laws and regulations, and were there any breaches?",
         "guidance": "Breaches detected, penalties, remedial actions."},
        {"group": "D. Principal Risks & Uncertainties", "question": "What governance, fraud, and operational risks have been identified, and what is their status?",
         "guidance": "Risk assessment results, top risks, mitigation status."},
        {"group": "I. Forward-looking & Future Prospects", "question": "What major audit trends are anticipated in the next 1-3 years, and what is the strategic response?",
         "guidance": "Data analytics, continuous auditing, forensic services."}
    ]
}

# ---------- GLOBAL/CORPORATE QUESTIONS ----------
GLOBAL_QUESTIONS = [
    {"group": "J. Compliance & Responsibility", "question": "Has the statement of responsibility (TFRS 1, para 47-48) been formally adopted, confirming accountability for true and fair financial statements?",
     "guidance": "Draft a formal statement acknowledging responsibility."},
    {"group": "J. Compliance & Responsibility", "question": "Has full compliance with TFRS 1 and all relevant laws been declared in the report?",
     "guidance": "Declare full compliance with TFRS 1 and all relevant laws."},
    {"group": "H. Corporate Governance", "question": "Who is the external auditor (name, address, registration, TIN, PF number), and how are they appointed?",
     "guidance": "Provide full contact details and registration numbers."},
    {"group": "J. Compliance & Responsibility", "question": "What political and charitable donations were made, and to which political recipients?",
     "guidance": "Disclose total political donations and recipients; charitable donations totals (names not required)."},
    {"group": "J. Compliance & Responsibility", "question": "Has the report been published on the Ministry website within 30 days of approval?",
     "guidance": "Confirm if published on the Ministry website within 30 days of approval."},
    {"group": "J. Compliance & Responsibility", "question": "Who approved the report, and on what date (list signatories and designations)?",
     "guidance": "List signatories (names, designations) and approval date."},
    {"group": "H. Corporate Governance", "question": "How does the Ministry comply with best practice corporate governance codes?",
     "guidance": "Explain how the Ministry complies with best practice governance codes."},
    {"group": "H. Corporate Governance", "question": "Who are the governance members (Permanent Secretary, Commissioners, etc.), and what is their meeting attendance record?",
     "guidance": "List governance members and meeting attendance."},
    {"group": "I. Forward-looking & Future Prospects", "question": "What are the major strategic priorities for the entire Ministry in the next 1-3 years, and what is the outlook?",
     "guidance": "Based on national plans, mineral sector strategy, and budget outlook."}
]

# ---------- FILE & SESSION MANAGEMENT ----------
DATA_FILE = "tfrs_data.csv"
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

def load_data():
    if os.path.exists(DATA_FILE):
        df = pd.read_csv(DATA_FILE)
        # Add missing columns if upgrading from old version
        if 'Group' not in df.columns:
            df['Group'] = ''
        if 'Attachments' not in df.columns:
            df['Attachments'] = ''
        if 'Narrative' not in df.columns:
            df['Narrative'] = ''
        df['Status'] = df['Status'].fillna('NA')
        df['Comments'] = df['Comments'].fillna('')
        df['Narrative'] = df['Narrative'].fillna('')
        df['Attachments'] = df['Attachments'].fillna('')
        df['Group'] = df['Group'].fillna('')
        return df
    else:
        return create_new_data()

def create_new_data():
    rows = []
    # Department-specific questions
    for dept, questions in DEPARTMENT_QUESTIONS.items():
        for q in questions:
            rows.append({
                "Department": dept,
                "Group": q["group"],
                "Question": q["question"],
                "Guidance": q["guidance"],
                "Status": "NA",
                "Comments": "",
                "Narrative": "",
                "Attachments": "",
                "Last_Updated": datetime.now().strftime("%Y-%m-%d %H:%M")
            })
    # Global questions
    for q in GLOBAL_QUESTIONS:
        rows.append({
            "Department": "**CORPORATE / GENERAL**",
            "Group": q["group"],
            "Question": q["question"],
            "Guidance": q["guidance"],
            "Status": "NA",
            "Comments": "",
            "Narrative": "",
            "Attachments": "",
            "Last_Updated": datetime.now().strftime("%Y-%m-%d %H:%M")
        })
    df = pd.DataFrame(rows)
    df.to_csv(DATA_FILE, index=False)
    return df

def save_data(df):
    df.to_csv(DATA_FILE, index=False)

# ---------- STREAMLIT APP ----------
st.set_page_config(layout="wide")
st.title("🏛 TFRS 1 Report Builder – Ministry of Minerals")

if 'data' not in st.session_state:
    st.session_state.data = load_data()
if 'uploaded_files' not in st.session_state:
    st.session_state.uploaded_files = {}

# ---------- SIDEBAR ----------
st.sidebar.title("📋 Navigation")
selected_dept = st.sidebar.selectbox("Select Your Department", DEPARTMENTS + ["**CORPORATE / GENERAL**"])
st.sidebar.markdown("---")
st.sidebar.caption("📌 **Y** = Fully Met, **N** = Not Met, **NA** = Not Applicable")
st.sidebar.caption("📌 Provide **Narrative/Figures** – these will become the contents of your final report.")
st.sidebar.caption("📌 Attach supporting files (PDF, Excel, Images) under each question.")

# ---------- ADMIN: EXTEND QUESTIONS ----------
with st.sidebar.expander("🔧 Admin - Add New Question (Password: admin123)"):
    admin_pass = st.text_input("Password", type="password", key="admin_pass")
    if admin_pass == "admin123":
        st.success("Admin access granted.")
        new_dept = st.selectbox("Department", DEPARTMENTS + ["**CORPORATE / GENERAL**"])
        new_group = st.selectbox("TFRS Group", list(TFRS_GROUPS.keys()))
        new_q = st.text_area("Question", placeholder="e.g., How does your department ensure compliance with the Public Procurement Act?")
        new_guidance = st.text_area("Guidance", placeholder="e.g., State the compliance percentage, number of audits, and any corrective actions taken.")
        if st.button("➕ Add Question"):
            if new_q and new_guidance:
                new_row = {
                    "Department": new_dept,
                    "Group": new_group,
                    "Question": new_q,
                    "Guidance": new_guidance,
                    "Status": "NA",
                    "Comments": "",
                    "Narrative": "",
                    "Attachments": "",
                    "Last_Updated": datetime.now().strftime("%Y-%m-%d %H:%M")
                }
                st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame([new_row])], ignore_index=True)
                save_data(st.session_state.data)
                st.success(f"✅ Added question to {new_dept}!")
                st.rerun()
    elif admin_pass:
        st.error("Wrong password.")

st.sidebar.markdown("---")
st.sidebar.caption(f"Data file: {DATA_FILE}")
st.sidebar.caption(f"Last update: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

# ---------- DATA ENTRY (SIMPLIFIED - NO CONFUSING HEADERS) ----------
st.subheader(f"📋 Checklist: {selected_dept}")
st.caption("Answer each question below. Expand each item to see guidance and upload files.")

dept_mask = st.session_state.data["Department"] == selected_dept
dept_data = st.session_state.data[dept_mask].copy()

if dept_data.empty:
    st.warning(f"No questions found for {selected_dept}.")
    st.stop()

with st.form(key="entry_form"):
    updated_statuses = []
    updated_comments = []
    updated_narratives = []
    updated_attachments = []
    
    # Simple counter for questions
    q_counter = 1
    for idx, row in dept_data.iterrows():
        question = row["Question"]
        guidance = row["Guidance"]
        current_status = row["Status"]
        current_comment = row["Comments"]
        current_narrative = row["Narrative"]
        current_attachments = row["Attachments"]
        
        # Display as a simple expander with a number
        with st.expander(f"Q{q_counter}: {question}", expanded=False):
            col1, col2 = st.columns([1, 2])
            with col1:
                try:
                    default_idx = ["Y", "N", "NA"].index(current_status)
                except ValueError:
                    default_idx = 2
                status = st.selectbox(
                    "Status",
                    options=["Y", "N", "NA"],
                    index=default_idx,
                    key=f"status_{idx}",
                    label_visibility="collapsed"
                )
                st.caption(f"💡 {guidance}")
            with col2:
                narrative = st.text_area(
                    "Narrative / Key Figures (to be included in final report)",
                    value=current_narrative,
                    key=f"narrative_{idx}",
                    height=100,
                    placeholder="Paste actual data, numbers, percentages, and analysis here..."
                )
                comment = st.text_input(
                    "Additional Comments / References",
                    value=current_comment,
                    key=f"comment_{idx}",
                    placeholder="e.g., See page 5 of policy"
                )
                uploaded_file = st.file_uploader(
                    "Attach supporting document (PDF, Excel, Image)",
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
                    st.success(f"✅ File saved: {safe_name}")
                    current_attachments = safe_name if not current_attachments else current_attachments + ", " + safe_name
                if current_attachments:
                    st.info(f"📎 Attached: {current_attachments}")
            
            updated_statuses.append(status)
            updated_narratives.append(narrative)
            updated_comments.append(comment)
            updated_attachments.append(current_attachments)
        
        q_counter += 1
    
    submitted = st.form_submit_button("💾 Save My Data")
    if submitted:
        for i, (idx, row) in enumerate(dept_data.iterrows()):
            st.session_state.data.at[idx, "Status"] = updated_statuses[i]
            st.session_state.data.at[idx, "Narrative"] = updated_narratives[i]
            st.session_state.data.at[idx, "Comments"] = updated_comments[i]
            st.session_state.data.at[idx, "Attachments"] = updated_attachments[i]
            st.session_state.data.at[idx, "Last_Updated"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        save_data(st.session_state.data)
        st.success("✅ Data saved successfully!")

# ---------- DASHBOARD ----------
st.markdown("---")
st.header("📊 Live Compliance Dashboard")
all_depts = st.session_state.data["Department"].unique()
summary = []
for dept in all_depts:
    if dept == "**CORPORATE / GENERAL**":
        continue
    dept_filter = st.session_state.data[st.session_state.data["Department"] == dept]
    total = len(dept_filter)
    y_count = len(dept_filter[dept_filter["Status"] == "Y"])
    n_count = len(dept_filter[dept_filter["Status"] == "N"])
    na_count = len(dept_filter[dept_filter["Status"] == "NA"])
    compliance = (y_count / (y_count + n_count) * 100) if (y_count + n_count) > 0 else 0
    completion = ((y_count + n_count + na_count) / total) * 100
    summary.append({
        "Department": dept,
        "Yes": y_count,
        "No": n_count,
        "NA": na_count,
        "Compliance %": round(compliance, 1),
        "Completion %": round(completion, 1)
    })
df_summary = pd.DataFrame(summary)
st.dataframe(df_summary, width='stretch', hide_index=True)
st.bar_chart(df_summary.set_index("Department")["Compliance %"])

# ---------- REPORT GENERATOR ----------
st.markdown("---")
st.header("📄 Generate Draft TFRS 1 Report")

st.warning("The report compiles all **Narrative** fields, grouped by TFRS sections. Ensure all departments have saved their narratives.")

if st.button("📝 Generate Consolidated Draft Report (Word)"):
    output = io.BytesIO()
    
    if DOCX_AVAILABLE:
        doc = Document()
        doc.add_heading('REPORT BY THOSE CHARGED WITH GOVERNANCE', 0)
        doc.add_heading('Ministry of Minerals', level=1)
        doc.add_paragraph(f'Date: {datetime.now().strftime("%d %B %Y")}')
        doc.add_paragraph('Prepared in accordance with TFRS 1 (Effective 1st January 2021)')
        doc.add_paragraph('')
        
        # 1. Executive Summary
        doc.add_heading('1. Executive Summary', level=2)
        doc.add_paragraph(f'Compliance status across {len(df_summary)} departments:')
        for _, row in df_summary.iterrows():
            doc.add_paragraph(f'• {row["Department"]}: {row["Compliance %"]}% compliance (Yes: {row["Yes"]}, No: {row["No"]}, N/A: {row["NA"]})')
        doc.add_paragraph('')
        
        # 2. Corporate / General Section
        corp_data = st.session_state.data[st.session_state.data["Department"] == "**CORPORATE / GENERAL**"]
        if not corp_data.empty:
            doc.add_heading('2. Corporate Governance and General Disclosures', level=2)
            for _, row in corp_data.iterrows():
                doc.add_heading(row["Question"], level=3)
                doc.add_paragraph(row["Narrative"] if row["Narrative"] else "⚠ No narrative provided yet.")
                if row["Attachments"]:
                    doc.add_paragraph(f'Attachments: {row["Attachments"]}')
        doc.add_paragraph('')
        
        # 3. Grouped by TFRS Sections (across all departments)
        doc.add_heading('3. Operational and Financial Review by TFRS Section', level=2)
        for group in TFRS_GROUPS.keys():
            group_data = st.session_state.data[
                (st.session_state.data["Group"] == group) &
                (st.session_state.data["Department"] != "**CORPORATE / GENERAL**")
            ]
            if group_data.empty:
                continue
            doc.add_heading(f'3.{list(TFRS_GROUPS.keys()).index(group)+1} {group}', level=3)
            doc.add_paragraph(TFRS_GROUPS[group])
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
        mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        file_ext = "docx"
    else:
        # Fallback to plain text
        output.write("="*80 + "\n".encode())
        output.write("REPORT BY THOSE CHARGED WITH GOVERNANCE - MINISTRY OF MINERALS\n".encode())
        output.write(f"Date: {datetime.now().strftime('%d %B %Y')}\n".encode())
        output.write("="*80 + "\n\n".encode())
        
        corp_data = st.session_state.data[st.session_state.data["Department"] == "**CORPORATE / GENERAL**"]
        if not corp_data.empty:
            output.write("CORPORATE GOVERNANCE & GENERAL DISCLOSURES\n".encode())
            output.write("-"*40 + "\n".encode())
            for _, row in corp_data.iterrows():
                output.write(f"{row['Question']}\n".encode())
                output.write(f"{row['Narrative'] if row['Narrative'] else 'N/A'}\n".encode())
                if row['Attachments']:
                    output.write(f"Attachments: {row['Attachments']}\n".encode())
                output.write("\n".encode())
        
        for group in TFRS_GROUPS.keys():
            group_data = st.session_state.data[
                (st.session_state.data["Group"] == group) &
                (st.session_state.data["Department"] != "**CORPORATE / GENERAL**")
            ]
            if group_data.empty:
                continue
            output.write(f"\n\n{group}\n".encode())
            output.write("="*len(group) + "\n".encode())
            for dept in DEPARTMENTS:
                dept_group = group_data[group_data["Department"] == dept]
                if dept_group.empty:
                    continue
                output.write(f"\n{dept}:\n".encode())
                for _, row in dept_group.iterrows():
                    output.write(f"  - {row['Question']}\n".encode())
                    output.write(f"    {row['Narrative'] if row['Narrative'] else 'N/A'}\n".encode())
                    if row['Attachments']:
                        output.write(f"    Evidence: {row['Attachments']}\n".encode())
        
        mime_type = "text/plain"
        file_ext = "txt"
    
    st.download_button(
        label=f"⬇ Download Draft Report (.{file_ext})",
        data=output.getvalue(),
        file_name=f"TFRS1_Draft_Report_{datetime.now().strftime('%Y%m%d')}.{file_ext}",
        mime=mime_type
    )
    st.success("Draft report generated! Review and edit the narrative fields for missing data.")
