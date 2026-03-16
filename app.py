import streamlit as st
import pdfplumber
import pandas as pd
import re
import pytesseract
from pdf2image import convert_from_bytes
import io

st.title("ServiceNow RITM PDF Extractor")

st.write("Upload RITM PDFs to extract ticket details and download Excel.")

uploaded_files = st.file_uploader(
    "Upload RITM PDFs",
    type=["pdf"],
    accept_multiple_files=True
)

# -----------------------------
# TEXT EXTRACTION
# -----------------------------

def extract_text_from_pdf(file):

    text = ""

    try:
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except:
        pass

    # If PDF has no selectable text → use OCR
    if text.strip() == "":
        images = convert_from_bytes(file.read())
        for img in images:
            text += pytesseract.image_to_string(img)

    return text


# -----------------------------
# FIELD EXTRACTION
# -----------------------------

def extract_fields(text):

    def search(pattern):
        match = re.search(pattern, text, re.IGNORECASE)
        return match.group(1).strip() if match else ""

    # RITM Number
    ritm = re.search(r'RITM\d+', text)
    ritm_number = ritm.group(0) if ritm else ""

    # Requested For
    requested_for = search(r'Requested\s*for\s*(.+)')

    # Opened
    opened = search(r'Opened\s*(\d{2}/\d{2}/\d{4}\s*\d{2}:\d{2}:\d{2})')

    # Opened By
    opened_by = search(r'Opened\s*by\s*(.+)')

    # State
    state = search(r'State\s*(Closed Complete|Closed|Open|Completed)')

    # Action Required
    action_required = search(
        r'What action do you require on the account\?\s*(.+)'
    )

    # -----------------------------
    # APPROVER EXTRACTION (FIXED)
    # -----------------------------

    approver = ""

    approver_match = re.search(
        r'Approved\s*\n\s*([A-Za-z\s]+)',
        text
    )

    if approver_match:
        approver = approver_match.group(1).strip()

    # -----------------------------
    # CREATED DATE EXTRACTION
    # -----------------------------

    created = ""

    created_match = re.search(
        r'Created\s*\n\s*(\d{2}/\d{2}/\d{4}\s\d{2}:\d{2}:\d{2})',
        text
    )

    if created_match:
        created = created_match.group(1)

    data = {
        "RITM Number": ritm_number,
        "Requested For": requested_for,
        "Opened": opened,
        "Opened By": opened_by,
        "Approvers": approver,
        "Created": created,
        "State": state,
        "What action do you require on the account?": action_required
    }

    return data


# -----------------------------
# PROCESS FILES
# -----------------------------

if uploaded_files:

    results = []

    progress = st.progress(0)

    for i, file in enumerate(uploaded_files):

        text = extract_text_from_pdf(file)

        data = extract_fields(text)

        results.append(data)

        progress.progress((i + 1) / len(uploaded_files))

    df = pd.DataFrame(results)

    st.subheader("Extracted Data")

    st.dataframe(df)

    # -----------------------------
    # CREATE EXCEL FILE
    # -----------------------------

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)

    output.seek(0)

    st.download_button(
        label="Download Excel",
        data=output,
        file_name="RITM_Audit_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )