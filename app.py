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

YEAR_DISPLAY = ["2025/2026", "2026/2027", "2027/2028", "2028/2029", "2029/2030", "2030/2031"]
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
# LOAD QUESTIONS FROM EXTERNAL CSV
# ------------------------------------------------------------
MASTER_FILE = "questions_master.csv"

def load_master_questions():
    """Load questions from the master CSV file."""
    if not os.path.exists(MASTER_FILE):
        st.error(f"CRITICAL: {MASTER_FILE} not found! Please upload it to the repository.")
        return {}
    try:
        df = pd.read_csv(MASTER_FILE)
        # Ensure required columns
        required = ['Department', 'Group', 'Question', 'Guidance']
        for col in required:
            if col not in df.columns:
                st.error(f"Missing column '{col}' in {MASTER_FILE}")
                return {}
        # Build dictionary
        dept_questions = {}
        for _, row in df.iterrows():
            dept = row['Department']
            dept_questions.setdefault(dept, []).append({
                "group": row['Group'],
                "question": row['Question'],
                "guidance": row['Guidance']
            })
        return dept_questions
    except Exception as e:
        st.error(f"Error reading {MASTER_FILE}: {e}")
        return {}

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

def get_question_count():
    """Return total number of questions in master file."""
    df = pd.read_csv(MASTER_FILE)
    return len(df)

def load_data(year_key):
    """Load user data. If it doesn't match master question count, rebuild."""
    file_key = f"{DATA_FILE}_{year_key}.csv"
    master_count = get_question_count()
    
    if os.path.exists(file_key):
        try:
            df = pd.read_csv(file_key)
            # Check if the number of rows matches the master question count
            if len(df) == master_count:
                required = ['Department', 'Group', 'Question', 'Guidance', 'Comments', 'Narrative', 'Attachments', 'Year', 'Last_Updated']
                for col in required:
                    if col not in df.columns:
                        df[col] = ''
                df['Year'] = df.get('Year', year_key)
                return df
            else:
                # Mismatch found – delete and rebuild
                os.remove(file_key)
                return create_new_data(year_key)
        except Exception:
            if os.path.exists(file_key):
                os.remove(file_key)
            return create_new_data(year_key)
    else:
        return create_new_data(year_key)

def create_new_data(year_key):
    """Create a new user data file from the master question list."""
    dept_questions = load_master_questions()
    if not dept_questions:
        # Fallback: create empty df with required columns
        return pd.DataFrame(columns=['Department', 'Group', 'Question', 'Guidance', 'Comments', 'Narrative', 'Attachments', 'Year', 'Last_Updated'])
    
    rows = []
    for dept, questions in dept_questions.items():
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
    file_key = f"{DATA_FILE}_{year_key}.csv"
    df.to_csv(file_key, index=False)
    return df

def save_data(df, year_key):
    file_key = f"{DATA_FILE}_{year_key}.csv"
    df.to_csv(file_key, index=False)

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
    selected_dept = st.sidebar.selectbox("Select Your Division/Unit", DEPARTMENTS)

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
            st.warning("Adding questions via the Admin panel will NOT update the master file. To add questions permanently, edit the 'questions_master.csv' file in the repository.")
            new_dept = st.selectbox("Department", DEPARTMENTS)
            new_group = st.selectbox("TFRS Group", list(TFRS_GROUPS.keys()))
            new_q = st.text_area("Question")
            new_guidance = st.text_area("Evidence Required")
            if st.button("➕ Add Question (Temporary)"):
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
                    st.success("✅ Added temporarily! To make it permanent, add it to 'questions_master.csv'.")
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
        st.warning(f"No questions found for {selected_dept}. Please check the master questions file.")
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
