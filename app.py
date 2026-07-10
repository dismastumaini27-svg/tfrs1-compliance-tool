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

YEARS = ["2025", "2026", "2027", "2028", "2029", "2030"]

TFRS_GROUPS = {
    "A. Nature of Operation": "Description of the industry, markets, products, services, and regulatory environment.",
    "G. Performance & KPIs": "Key performance indicators, service delivery targets, and efficiency.",
    "I. Forward-looking & Future Prospects": "Anticipated trends, challenges, and strategic responses.",
    "J. Compliance & Responsibility": "Statement of responsibility, compliance with TFRS 1."
}

# ---------- SAMPLE QUESTIONS (only 2 per department) ----------
DEPARTMENT_QUESTIONS = {
    "Human Resource": [
        {"group": "A. Nature of Operation", "question": "What is the role of HR in supporting the Ministry's operational mandate?", "guidance": "Outline staffing levels and key HR functions."},
        {"group": "G. Performance & KPIs", "question": "What are the key HR performance metrics, and what do the trends show?", "guidance": "Provide turnover rate, cost per hire, etc."}
    ],
    "Monitoring & Evaluation": [
        {"group": "A. Nature of Operation", "question": "How is the M&E function structured?", "guidance": "Outline the M&E framework and reporting cycles."},
        {"group": "G. Performance & KPIs", "question": "What are the actual service delivery targets achieved?", "guidance": "Actual outputs vs. targets (e.g., inspections)."}
    ],
    "Information & Communication Technology": [
        {"group": "A. Nature of Operation", "question": "How is the ICT function structured?", "guidance": "Outline systems supported and key services."},
        {"group": "G. Performance & KPIs", "question": "What is the system downtime record?", "guidance": "Number of outages, average resolution time."}
    ],
    "Government Communication Unit": [
        {"group": "A. Nature of Operation", "question": "What is the role of the Government Communication Unit?", "guidance": "Outline public engagement and media relations."},
        {"group": "G. Performance & KPIs", "question": "How effectively does the unit disseminate information?", "guidance": "Number of press releases, website traffic."}
    ],
    "Policy & Planning": [
        {"group": "A. Nature of Operation", "question": "What is the role of the Policy & Planning function?", "guidance": "Outline policy formulation and planning."},
        {"group": "I. Forward-looking & Future Prospects", "question": "What major policy trends are anticipated?", "guidance": "New mining codes, environmental regulations."}
    ],
    "Procurement": [
        {"group": "A. Nature of Operation", "question": "What is the role of the Procurement function?", "guidance": "Outline procurement of goods and contract management."},
        {"group": "J. Compliance & Responsibility", "question": "What is the level of adherence to the Public Procurement Act?", "guidance": "Percentage of tenders fully compliant."}
    ],
    "Commissioner for Minerals": [
        {"group": "A. Nature of Operation", "question": "What is the mandate of the Commissioner for Minerals?", "guidance": "Outline licensing, inspection, and revenue collection."},
        {"group": "G. Performance & KPIs", "question": "What are the mineral production statistics?", "guidance": "Tonnes produced, royalty collections."}
    ],
    "Finance & Accounts": [
        {"group": "A. Nature of Operation", "question": "What is the current cash flow position?", "guidance": "Opening/closing balances, liquidity ratio."},
        {"group": "G. Performance & KPIs", "question": "What is the budget absorption rate?", "guidance": "Overall budget execution percentage."}
    ],
    "Internal Auditing": [
        {"group": "A. Nature of Operation", "question": "What is the annual audit plan?", "guidance": "Scope coverage, planned audits."},
        {"group": "J. Compliance & Responsibility", "question": "What is the compliance status with financial laws?", "guidance": "Breaches detected, remedial actions."}
    ]
}

GLOBAL_QUESTIONS = [
    {"group": "J. Compliance & Responsibility", "question": "Has the statement of responsibility been formally adopted?", "guidance": "Draft a formal statement."},
    {"group": "J. Compliance & Responsibility", "question": "Who is the external auditor?", "guidance": "Provide full contact details."}
]

# ---------- FILE HELPERS ----------
DATA_FILE = "tfrs_data"
SYNTHESIS_FILE = "synthesis_data"
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

def count_words(text):
    if not text or not isinstance(text, str):
        return 0
    return len(re.findall(r'\b\w+\b', text))

def safe_read_csv(file_key, expected_columns=None):
    if os.path.exists(file_key):
        try:
            df = pd.read_csv(file_key)
            if expected_columns:
                if all(col in df.columns for col in expected_columns):
                    return df
                else:
                    os.remove(file_key)
                    return None
            return df
        except Exception:
            os.remove(file_key)
            return None
    return None

def load_data(year):
    file_key = f"{DATA_FILE}_{year}.csv"
    df = safe_read_csv(file_key, ['Department', 'Group', 'Question', 'Guidance', 'Comments', 'Narrative', 'Attachments', 'Year', 'Last_Updated'])
    if df is not None:
        for col in ['Comments', 'Narrative', 'Attachments', 'Group']:
            if col not in df.columns:
                df[col] = ''
        df['Year'] = df.get('Year', year)
        return df
    else:
        return create_new_data(year)

def create_new_data(year):
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
                "Year": year,
                "Last_Updated": datetime.now().strftime("%Y-%m-%d %H:%M")
            })
    for q in GLOBAL_QUESTIONS:
        rows.append({
            "Department": "**CORPORATE / GENERAL**",
            "Group": q["group"],
            "Question": q["question"],
            "Guidance": q["guidance"],
            "Comments": "",
            "Narrative": "",
            "Attachments": "",
            "Year": year,
            "Last_Updated": datetime.now().strftime("%Y-%m-%d %H:%M")
        })
    df = pd.DataFrame(rows)
    file_key = f"{DATA_FILE}_{year}.csv"
    df.to_csv(file_key, index=False)
    return df

def save_data(df, year):
    file_key = f"{DATA_FILE}_{year}.csv"
    df.to_csv(file_key, index=False)

def load_synthesis(year):
    file_key = f"{SYNTHESIS_FILE}_{year}.csv"
    df = safe_read_csv(file_key, ['Group', 'Synthesis'])
    if df is not None:
        for group in TFRS_GROUPS.keys():
            if group not in df['Group'].values:
                df = pd.concat([df, pd.DataFrame([{'Group': group, 'Synthesis': ''}])], ignore_index=True)
        df = df[['Group', 'Synthesis']]
        return df
    else:
        rows = [{'Group': group, 'Synthesis': ''} for group in TFRS_GROUPS.keys()]
        df = pd.DataFrame(rows)
        file_key = f"{SYNTHESIS_FILE}_{year}.csv"
        df.to_csv(file_key, index=False)
        return df

def save_synthesis(df, year):
    file_key = f"{SYNTHESIS_FILE}_{year}.csv"
    df = df[['Group', 'Synthesis']]
    df.to_csv(file_key, index=False)

# ---------- STREAMLIT APP ----------
st.set_page_config(layout="wide")
st.title("🏛 TFRS 1 Report Builder – Ministry of Minerals")

try:
    # Sidebar
    st.sidebar.title("📋 Reporting Period")
    selected_year = st.sidebar.selectbox("Select Financial Year", YEARS)

    st.sidebar.markdown("---")
    st.sidebar.title("📌 Department")
    selected_dept = st.sidebar.selectbox("Select Your Department", DEPARTMENTS + ["**CORPORATE / GENERAL**"])

    st.sidebar.markdown("---")
    st.sidebar.caption("📌 Write at least **100 words** per narrative.")
    st.sidebar.caption("📌 Attach supporting files.")

    # Load data for selected year
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
            new_dept = st.selectbox("Department", DEPARTMENTS + ["**CORPORATE / GENERAL**"])
            new_group = st.selectbox("TFRS Group", list(TFRS_GROUPS.keys()))
            new_q = st.text_area("Question", placeholder="e.g., How does your department ensure compliance?")
            new_guidance = st.text_area("Guidance", placeholder="e.g., Provide compliance percentage.")
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
                    st.success(f"✅ Added question to {new_dept}!")
                    st.rerun()
        elif admin_pass:
            st.error("Wrong password.")

    st.sidebar.markdown("---")
    st.sidebar.caption(f"Financial Year {selected_year}")

    # ---------- DATA ENTRY ----------
    st.subheader(f"📋 Checklist: {selected_dept}")
    st.caption(f"**Reporting Period:** Financial Year {selected_year}")
    st.caption("Minimum 100 words per narrative.")

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
                    st.caption(f"💡 {guidance}")
                    narrative = st.text_area(
                        "Narrative / Detailed Explanation",
                        value=current_narrative,
                        key=f"narrative_{idx}",
                        height=150,
                        placeholder="Write a detailed explanation (minimum 100 words)..."
                    )
                    comment = st.text_input(
                        "Additional Comments / References",
                        value=current_comment,
                        key=f"comment_{idx}",
                        placeholder="e.g., See page 5"
                    )
                with col2:
                    new_word_count = count_words(narrative)
                    if new_word_count >= 100:
                        st.success(f"✅ {new_word_count} words - Compliant!")
                    else:
                        st.warning(f"⚠️ {new_word_count} / 100 words")
                        st.progress(new_word_count / 100)

                    uploaded_file = st.file_uploader(
                        "Attach supporting document",
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

                updated_narratives.append(narrative)
                updated_comments.append(comment)
                updated_attachments.append(current_attachments)

            q_counter += 1

        submitted = st.form_submit_button("💾 Save My Data")
        if submitted:
            for i, (idx, row) in enumerate(dept_data.iterrows()):
                st.session_state.data.at[idx, "Narrative"] = updated_narratives[i]
                st.session_state.data.at[idx, "Comments"] = updated_comments[i]
                st.session_state.data.at[idx, "Attachments"] = updated_attachments[i]
                st.session_state.data.at[idx, "Last_Updated"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                st.session_state.data.at[idx, "Year"] = selected_year
            save_data(st.session_state.data, selected_year)
            st.success("✅ Data saved successfully!")

    # ---------- DASHBOARD ----------
    st.markdown("---")
    st.header("📊 Live Compliance Dashboard")
    st.caption(f"**Reporting Period:** Financial Year {selected_year}")
    st.info("💡 **Compliance %** = (questions with ≥100 words) / (total questions).")

    all_depts = st.session_state.data["Department"].unique()
    summary = []
    for dept in all_depts:
        if dept == "**CORPORATE / GENERAL**":
            continue
        dept_filter = st.session_state.data[st.session_state.data["Department"] == dept]
        total = len(dept_filter)
        compliant = sum(1 for _, r in dept_filter.iterrows() if count_words(r["Narrative"]) >= 100)
        compliance = (compliant / total * 100) if total > 0 else 0
        summary.append({
            "Department": dept,
            "Total Questions": total,
            "✅ Compliant (≥100 words)": compliant,
            "❌ Incomplete": total - compliant,
            "Compliance %": round(compliance, 1)
        })
    df_summary = pd.DataFrame(summary)
    st.dataframe(df_summary, use_container_width=True, hide_index=True)
    st.subheader("Overall Compliance by Department (%)")
    st.bar_chart(df_summary.set_index("Department")["Compliance %"])

    # ---------- ADMIN SYNTHESIS ----------
    st.markdown("---")
    st.header("🔧 Admin Consolidation / Synthesis")

    with st.expander("📝 Admin: Synthesize Departmental Inputs (Password: admin123)"):
        synth_pass = st.text_input("Enter Admin Password", type="password", key="synth_pass")
        if synth_pass == "admin123":
            st.success("✅ Admin access granted.")
            st.info("Rewrite raw inputs into one cohesive paragraph per TFRS group.")

            st.subheader("📄 Raw Departmental Inputs")
            for group in TFRS_GROUPS.keys():
                group_data = st.session_state.data[
                    (st.session_state.data["Group"] == group) &
                    (st.session_state.data["Department"] != "**CORPORATE / GENERAL**")
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
                    placeholder="Write a cohesive paragraph here..."
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
                st.success(f"✅ Synthesis saved for FY {selected_year}!")
                st.rerun()
        elif synth_pass:
            st.error("Wrong password.")

    # ---------- REPORT GENERATOR ----------
    st.markdown("---")
    st.header("📄 Generate Draft TFRS 1 Report")
    st.warning(f"Report for Financial Year {selected_year}.")

    if st.button("📝 Generate Consolidated Draft Report (Word)"):
        output = io.BytesIO()
        if DOCX_AVAILABLE:
            doc = Document()
            doc.add_heading('REPORT BY THOSE CHARGED WITH GOVERNANCE', 0)
            doc.add_heading('Ministry of Minerals', level=1)
            doc.add_paragraph(f'Reporting Period: Financial Year {selected_year}')
            doc.add_paragraph(f'Date: {datetime.now().strftime("%d %B %Y")}')
            doc.add_paragraph('Prepared in accordance with TFRS 1')
            doc.add_paragraph('')

            doc.add_heading('1. Executive Summary', level=2)
            for _, row in df_summary.iterrows():
                doc.add_paragraph(f'• {row["Department"]}: {row["Compliance %"]}% compliance ({row["✅ Compliant (≥100 words)"]} out of {row["Total Questions"]})')
            doc.add_paragraph('')

            corp_data = st.session_state.data[st.session_state.data["Department"] == "**CORPORATE / GENERAL**"]
            if not corp_data.empty:
                doc.add_heading('2. Corporate Governance', level=2)
                for _, row in corp_data.iterrows():
                    doc.add_heading(row["Question"], level=3)
                    doc.add_paragraph(row["Narrative"] if row["Narrative"] else "⚠ Not provided.")
                    if row["Attachments"]:
                        doc.add_paragraph(f'Attachments: {row["Attachments"]}')
                doc.add_paragraph('')

            doc.add_heading('3. Operational and Financial Review by TFRS Section', level=2)
            synth_df = st.session_state.synthesis
            for group in TFRS_GROUPS.keys():
                doc.add_heading(f'3.{list(TFRS_GROUPS.keys()).index(group)+1} {group}', level=3)
                doc.add_paragraph(TFRS_GROUPS[group])
                synth_row = synth_df[synth_df['Group'] == group]
                if not synth_row.empty and synth_row.iloc[0]['Synthesis'] and synth_row.iloc[0]['Synthesis'].strip():
                    doc.add_paragraph(synth_row.iloc[0]['Synthesis'], style='List Bullet')
                    doc.add_paragraph("*(Consolidated by Management.)*", style='List Bullet')
                else:
                    group_data = st.session_state.data[
                        (st.session_state.data["Group"] == group) &
                        (st.session_state.data["Department"] != "**CORPORATE / GENERAL**")
                    ]
                    if group_data.empty:
                        doc.add_paragraph('⚠ No data provided.')
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
            output.write("REPORT BY THOSE CHARGED WITH GOVERNANCE - MINISTRY OF MINERALS\n".encode())
            output.write(f"Financial Year {selected_year}\n".encode())
            output.write("="*80 + "\n\n".encode())
            output.write("(Install python-docx for a proper Word document.)\n".encode())
            mime = "text/plain"
            ext = "txt"

        st.download_button(
            label=f"⬇ Download Draft Report (.{ext})",
            data=output.getvalue(),
            file_name=f"TFRS1_FY{selected_year}_{datetime.now().strftime('%Y%m%d')}.{ext}",
            mime=mime
        )
        st.success(f"Report generated for FY {selected_year}!")

except Exception as e:
    st.error(f"An error occurred: {e}")
    st.stop()
