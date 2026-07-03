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

# Define TFRS Groups for consistent grouping
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

# ---------- DEPARTMENT-SPECIFIC QUESTIONS WITH GROUPS ----------
DEPARTMENT_QUESTIONS = {
    "Human Resource": [
        {"group": "A. Nature of Operation", "question": "Describe the role of HR in supporting the Ministry's operational mandate.",
         "guidance": "Outline staffing levels, structure, and key HR functions (recruitment, payroll, training)."},
        {"group": "C. Resources", "question": "Staff recruitment, retention, and development strategies",
         "guidance": "Provide vacancy rates, turnover %, training days per staff, and strategic plans."},
        {"group": "C. Resources", "question": "Employee welfare initiatives (health, pension, safety)",
         "guidance": "Coverage percentages, budget allocated, specific schemes (PPF/NSSF)."},
        {"group": "C. Resources", "question": "Gender parity and equal opportunity policies",
         "guidance": "Current gender ratio (M/F). Recruitment/promotion by gender."},
        {"group": "C. Resources", "question": "Training and career progression for persons with disabilities",
         "guidance": "Numbers of disabled staff recruited, trained, promoted."},
        {"group": "E. Stakeholder Relationships", "question": "Grievance handling mechanisms for staff",
         "guidance": "Number of grievances lodged, resolved, average resolution time."},
        {"group": "G. Performance & KPIs", "question": "Key HR performance metrics",
         "guidance": "Turnover rate, cost per hire, staff satisfaction score, vacancy rate."},
        {"group": "I. Forward-looking & Future Prospects", "question": "Major HR trends and strategic response (next 1-3 years)",
         "guidance": "Digital HR, talent shortages, remote work, skills gaps."}
    ],
    "Monitoring & Evaluation": [
        {"group": "A. Nature of Operation", "question": "Describe the M&E function and its role in tracking Ministry performance.",
         "guidance": "Outline the M&E framework, reporting cycles, and key stakeholders."},
        {"group": "G. Performance & KPIs", "question": "Service delivery targets (outputs) and performance indicators",
         "guidance": "Actual outputs vs. targets (e.g., inspections, reports produced)."},
        {"group": "G. Performance & KPIs", "question": "Efficiency of resource utilization (inputs vs outputs)",
         "guidance": "Cost per output, productivity ratios, budget absorption rate."},
        {"group": "E. Stakeholder Relationships", "question": "Stakeholder feedback and citizen satisfaction",
         "guidance": "Satisfaction scores, feedback response rates, actions taken."},
        {"group": "G. Performance & KPIs", "question": "Achievement of intended societal impacts (outcomes)",
         "guidance": "Outcome indicators (e.g., health improvements, economic growth)."},
        {"group": "H. Corporate Governance", "question": "Implementation status of previous M&E recommendations",
         "guidance": "Number closed, in progress, not implemented."},
        {"group": "G. Performance & KPIs", "question": "Key performance indicators (KPIs) tracked consistently",
         "guidance": "Targets vs. actuals for top 5 KPIs."},
        {"group": "I. Forward-looking & Future Prospects", "question": "Major M&E trends and strategic response (next 1-3 years)",
         "guidance": "New data collection tools, impact evaluation, real-time reporting."}
    ],
    "Information & Communication Technology": [
        {"group": "A. Nature of Operation", "question": "Describe the ICT function and its critical role in Ministry operations.",
         "guidance": "Outline systems supported, user base, and key services (email, networks, databases)."},
        {"group": "D. Principal Risks & Uncertainties", "question": "Cybersecurity risks, data privacy measures, and IT policies",
         "guidance": "Number of incidents, compliance level, policy version."},
        {"group": "B. Objectives, Strategies & Operating Model", "question": "Digital transformation strategy and implementation progress",
         "guidance": "Milestones achieved (systems launched, automation, adoption rates)."},
        {"group": "D. Principal Risks & Uncertainties", "question": "Business continuity and disaster recovery plans",
         "guidance": "Last testing date, Recovery Time Objective (RTO) achieved."},
        {"group": "F. Capital Structure, Treasury & Liquidity", "question": "Costs and benefits of major IT investments",
         "guidance": "Total capex/opex, ROI, efficiency gains (e.g., time reduced by X%)."},
        {"group": "C. Resources", "question": "IT asset inventory and management",
         "guidance": "Total assets, age profile, replacement value, disposal records."},
        {"group": "G. Performance & KPIs", "question": "System downtime incidents and resolutions",
         "guidance": "Number of outages, average resolution time, root causes."},
        {"group": "I. Forward-looking & Future Prospects", "question": "Major IT trends and strategic response (next 1-3 years)",
         "guidance": "Cloud migration, AI, mobile apps, data analytics, remote work support."}
    ],
    "Government Communication Unit": [
        {"group": "A. Nature of Operation", "question": "Describe the role of the Government Communication Unit.",
         "guidance": "Outline public engagement, media relations, and information dissemination."},
        {"group": "B. Objectives, Strategies & Operating Model", "question": "Communication strategies and public engagement plans",
         "guidance": "Campaigns, events, media engagements, reach (estimated audience)."},
        {"group": "G. Performance & KPIs", "question": "Dissemination of information to citizens",
         "guidance": "Number of press releases, website traffic, social media impressions."},
        {"group": "E. Stakeholder Relationships", "question": "Mechanisms for handling citizen feedback/complaints",
         "guidance": "Feedback volumes, response rates, resolution times."},
        {"group": "J. Compliance & Responsibility", "question": "Role in transparency and access to information",
         "guidance": "Number of information requests handled, proactive disclosures."},
        {"group": "G. Performance & KPIs", "question": "Public awareness campaigns and their effectiveness",
         "guidance": "Campaign reach, behavior change results, cost per person reached."},
        {"group": "D. Principal Risks & Uncertainties", "question": "Risks related to misinformation and reputation",
         "guidance": "Number of reputation incidents, media monitoring outcomes."},
        {"group": "I. Forward-looking & Future Prospects", "question": "Major communication trends and strategic response (next 1-3 years)",
         "guidance": "Digital engagement, fake news, citizen journalism, crisis communication."}
    ],
    "Policy & Planning": [
        {"group": "A. Nature of Operation", "question": "Describe the Policy & Planning function and its strategic role.",
         "guidance": "Outline policy formulation, planning, and coordination with national priorities."},
        {"group": "B. Objectives, Strategies & Operating Model", "question": "Alignment of strategic objectives with national development plans",
         "guidance": "Map specific goals to national indicators (e.g., NDP, Five-Year Plan)."},
        {"group": "G. Performance & KPIs", "question": "Progress of strategic plan implementation and major milestones",
         "guidance": "Status of each goal (achieved, ongoing, delayed). Milestones reached."},
        {"group": "A. Nature of Operation", "question": "Legislative and regulatory changes affecting the minerals sector",
         "guidance": "List new laws/amendments and their expected impact on operations."},
        {"group": "B. Objectives, Strategies & Operating Model", "question": "Policy development framework and stakeholder consultation",
         "guidance": "Number of consultations, feedback incorporated."},
        {"group": "D. Principal Risks & Uncertainties", "question": "Risks related to policy implementation",
         "guidance": "Identify top risks and mitigation measures."},
        {"group": "I. Forward-looking & Future Prospects", "question": "Major policy trends and strategic response (next 1-3 years)",
         "guidance": "New mining codes, environmental regulations, regional integration."}
    ],
    "Procurement": [
        {"group": "A. Nature of Operation", "question": "Describe the Procurement function and its critical role.",
         "guidance": "Outline procurement of goods/services, contract management, and supplier relations."},
        {"group": "B. Objectives, Strategies & Operating Model", "question": "Strategic procurement plans and annual budgets",
         "guidance": "Total planned vs. actual spend. Number of tenders planned vs issued."},
        {"group": "J. Compliance & Responsibility", "question": "Adherence to the Public Procurement Act regulations",
         "guidance": "Percentage of tenders fully compliant. Any major breaches/penalties?"},
        {"group": "G. Performance & KPIs", "question": "Major contracts and supplier performance evaluations",
         "guidance": "Top 5 contracts by value. Supplier scores (e.g., 4.5/5)."},
        {"group": "D. Principal Risks & Uncertainties", "question": "Risks related to supply chain, fraud, and delays",
         "guidance": "List top risks and controls. Number of fraud incidents."},
        {"group": "G. Performance & KPIs", "question": "Value-for-money achieved and cost savings",
         "guidance": "Estimated savings from competitive bidding vs. estimated price."},
        {"group": "J. Compliance & Responsibility", "question": "Sustainable and local content procurement policies",
         "guidance": "Percentage local content, environmental criteria applied."},
        {"group": "I. Forward-looking & Future Prospects", "question": "Major procurement trends and strategic response (next 1-3 years)",
         "guidance": "e-Procurement, strategic sourcing, green procurement."}
    ],
    "Commissioner for Minerals": [
        {"group": "A. Nature of Operation", "question": "Describe the Commissioner's mandate in regulating the minerals sector.",
         "guidance": "Outline licensing, inspection, revenue collection, and sector development."},
        {"group": "G. Performance & KPIs", "question": "Mineral production statistics and royalty collections",
         "guidance": "Tonnes produced (by mineral), total value, royalties collected vs. target."},
        {"group": "A. Nature of Operation", "question": "Licensing processes and monitoring",
         "guidance": "Number of new/renewed licenses, compliance rate."},
        {"group": "G. Performance & KPIs", "question": "Inspection and compliance activities",
         "guidance": "Number of inspections, findings, enforcement actions."},
        {"group": "J. Compliance & Responsibility", "question": "Environmental protection and rehabilitation activities",
         "guidance": "Rehabilitation area (ha), number of EIAs, closure plans."},
        {"group": "G. Performance & KPIs", "question": "Revenue targets vs actual collections",
         "guidance": "Variance analysis: reasons for over/under collection."},
        {"group": "D. Principal Risks & Uncertainties", "question": "Risks related to mining compliance, illegal mining, and revenue leakage",
         "guidance": "Number of illegal mining cases, revenue lost, controls."},
        {"group": "I. Forward-looking & Future Prospects", "question": "Major minerals sector trends and strategic response (next 1-3 years)",
         "guidance": "Global commodity prices, new discoveries, local beneficiation policy."}
    ],
    "Finance & Accounts": [
        {"group": "F. Capital Structure, Treasury & Liquidity", "question": "Cash flow position and liquidity status",
         "guidance": "Opening/closing cash balances, cash inflow/outflow, liquidity ratio."},
        {"group": "G. Performance & KPIs", "question": "Budget absorption and major variances",
         "guidance": "Overall budget execution percentage (>90%?). Explain variances >10%."},
        {"group": "D. Principal Risks & Uncertainties", "question": "Financial risks (currency, inflation, fraud)",
         "guidance": "Quantify exposure, risk mitigation measures."},
        {"group": "F. Capital Structure, Treasury & Liquidity", "question": "Treasury policies and management of public funds",
         "guidance": "Cash management strategies, bank reconciliation status."},
        {"group": "F. Capital Structure, Treasury & Liquidity", "question": "Significant payments, commitments, and outstanding obligations",
         "guidance": "Large contracts pending, bills, committed funds."},
        {"group": "F. Capital Structure, Treasury & Liquidity", "question": "Financial impact of major capital projects",
         "guidance": "Total cost, funding sources, expected returns/benefits."},
        {"group": "I. Forward-looking & Future Prospects", "question": "Major financial trends and strategic response (next 1-3 years)",
         "guidance": "Budget constraints, funding reforms, PPPs, revenue diversification."}
    ],
    "Internal Auditing": [
        {"group": "H. Corporate Governance", "question": "Annual audit plan and charter",
         "guidance": "Scope coverage, number of planned audits vs completed."},
        {"group": "H. Corporate Governance", "question": "Key audit findings, recommendations, and management responses",
         "guidance": "List top 5 high-risk findings and whether management accepted them."},
        {"group": "H. Corporate Governance", "question": "Implementation status of previous audit recommendations",
         "guidance": "Number closed, in progress, not implemented. Provide percentages."},
        {"group": "H. Corporate Governance", "question": "Assessment of internal control environment",
         "guidance": "Qualitative opinion on control design and operation."},
        {"group": "J. Compliance & Responsibility", "question": "Compliance with financial laws and regulations",
         "guidance": "Breaches detected, penalties, remedial actions."},
        {"group": "D. Principal Risks & Uncertainties", "question": "Governance, fraud, and operational risks identified",
         "guidance": "Risk assessment results, top risks, mitigation status."},
        {"group": "I. Forward-looking & Future Prospects", "question": "Major audit trends and strategic response (next 1-3 years)",
         "guidance": "Data analytics, continuous auditing, forensic services."}
    ]
}

# ---------- GLOBAL/CORPORATE QUESTIONS (no department, goes into general report) ----------
GLOBAL_QUESTIONS = [
    {"group": "J. Compliance & Responsibility", "question": "Statement of Responsibility (TFRS 1, para 47-48)",
     "guidance": "Draft a formal statement acknowledging responsibility for true and fair financial statements."},
    {"group": "J. Compliance & Responsibility", "question": "Compliance Statement (TFRS 1, para 55)",
     "guidance": "Declare full compliance with TFRS 1 and all relevant laws."},
    {"group": "H. Corporate Governance", "question": "Auditor Appointment (TFRS 1, para 45)",
     "guidance": "Provide full name, address, registration, TIN, and PF number of the external auditor."},
    {"group": "J. Compliance & Responsibility", "question": "Political and Charitable Donations (TFRS 1, para 49)",
     "guidance": "Disclose total political donations and recipients; charitable donations totals."},
    {"group": "J. Compliance & Responsibility", "question": "Publication (TFRS 1, para 56)",
     "guidance": "Confirm if published on the Ministry website within 30 days of approval."},
    {"group": "J. Compliance & Responsibility", "question": "Approval and Signing (TFRS 1, para 57)",
     "guidance": "List signatories (names, designations) and approval date."},
    {"group": "H. Corporate Governance", "question": "Corporate Governance (TFRS 1, para 41)",
     "guidance": "Explain how the Ministry complies with best practice governance codes."},
    {"group": "H. Corporate Governance", "question": "Membership (TFRS 1, para 42-43)",
     "guidance": "List governance members (Permanent Secretary, Commissioners, etc.) and meeting attendance."},
    {"group": "I. Forward-looking & Future Prospects", "question": "Major strategic priorities for the entire Ministry (next 1-3 years)",
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
        new_q = st.text_area("Question")
        new_guidance = st.text_area("Guidance")
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
                st.success(f"Added question to {new_dept}!")
                st.rerun()
    elif admin_pass:
        st.error("Wrong password.")

st.sidebar.markdown("---")
st.sidebar.caption(f"Data file: {DATA_FILE}")
st.sidebar.caption(f"Last update: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

# ---------- DATA ENTRY ----------
st.subheader(f"📋 Checklist: {selected_dept}")
dept_mask = st.session_state.data["Department"] == selected_dept
dept_data = st.session_state.data[dept_mask].copy()

if dept_data.empty:
    st.warning(f"No questions found for {selected_dept}.")
    st.stop()

# Group questions by TFRS Group for display
grouped = dept_data.groupby("Group")

with st.form(key="entry_form"):
    st.caption("Expand each section and provide the required data. Attach files where available.")
    
    updated_statuses = []
    updated_comments = []
    updated_narratives = []
    updated_attachments = []
    
    # We need to iterate over the original index to update correctly
    for idx, row in dept_data.iterrows():
        group = row["Group"]
        question = row["Question"]
        guidance = row["Guidance"]
        current_status = row["Status"]
        current_comment = row["Comments"]
        current_narrative = row["Narrative"]
        current_attachments = row["Attachments"]
        
        # Display group header only once per group
        if group and group not in st.session_state.get('displayed_groups', []):
            st.markdown(f"### 📂 {group}")
            st.caption(TFRS_GROUPS.get(group, ""))
            st.markdown("---")
            if 'displayed_groups' not in st.session_state:
                st.session_state.displayed_groups = []
            st.session_state.displayed_groups.append(group)
        
        with st.expander(f"Q: {question}", expanded=False):
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
                # File uploader
                uploaded_file = st.file_uploader(
                    "Attach supporting document (PDF, Excel, Image)",
                    type=["pdf", "xlsx", "xls", "docx", "png", "jpg", "jpeg"],
                    key=f"upload_{idx}",
                    label_visibility="collapsed"
                )
                if uploaded_file is not None:
                    # Save file
                    ext = uploaded_file.name.split('.')[-1]
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    safe_name = f"{selected_dept.replace(' ', '_')}_{ts}.{ext}"
                    file_path = os.path.join(UPLOAD_DIR, safe_name)
                    with open(file_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    st.success(f"✅ File saved: {safe_name}")
                    current_attachments = safe_name if not current_attachments else current_attachments + ", " + safe_name
                # Show existing attachments
                if current_attachments:
                    st.info(f"📎 Attached: {current_attachments}")
            
            updated_statuses.append(status)
            updated_narratives.append(narrative)
            updated_comments.append(comment)
            updated_attachments.append(current_attachments)
    
    # Reset displayed groups for next render
    if 'displayed_groups' in st.session_state:
        del st.session_state.displayed_groups
    
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
            # Get all rows from all departments (except corporate) belonging to this group
            group_data = st.session_state.data[
                (st.session_state.data["Group"] == group) &
                (st.session_state.data["Department"] != "**CORPORATE / GENERAL**")
            ]
            if group_data.empty:
                continue
            doc.add_heading(f'3.{list(TFRS_GROUPS.keys()).index(group)+1} {group}', level=3)
            doc.add_paragraph(TFRS_GROUPS[group])
            # Group by department for readability
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
        
        # Corporate
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
        
        # TFRS Groups
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