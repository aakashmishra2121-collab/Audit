import streamlit as st
import pdfplumber
import pandas as pd
import re
import pytesseract
from pdf2image import convert_from_bytes

st.title("ServiceNow RITM Extractor")

st.write("Upload RITM PDFs and download Excel file")

uploaded_files = st.file_uploader(
    "Upload RITM PDFs",
    type=["pdf"],
    accept_multiple_files=True
)

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

    if text.strip() == "":
        images = convert_from_bytes(file.read())
        for img in images:
            text += pytesseract.image_to_string(img)

    return text


def extract_fields(text):

    def search(pattern):
        match = re.search(pattern, text, re.IGNORECASE)
        return match.group(1).strip() if match else ""

    ritm = re.search(r'RITM\d+', text)

    data = {
        "RITM Number": ritm.group(0) if ritm else "",
        "Requested For": search(r'Requested\s*for\s*(.+)'),
        "Opened": search(r'Opened\s*(\d{2}/\d{2}/\d{4}.*)'),
        "Opened By": search(r'Opened\s*by\s*(.+)'),
        "Approvers": search(r'Approver\s*(.+)'),
        "Created": search(r'Created\s*(\d{2}/\d{2}/\d{4}.*)'),
        "State": search(r'State\s*(Closed Complete|Closed|Open|Completed)'),
        "What action do you require on the account?": search(
            r'What action do you require on the account\?\s*(.+)'
        )
    }

    return data


if uploaded_files:

    results = []

    for file in uploaded_files:

        text = extract_text_from_pdf(file)

        data = extract_fields(text)

        results.append(data)

    df = pd.DataFrame(results)

    st.dataframe(df)

    df.to_excel("output.xlsx", index=False)

    with open("output.xlsx", "rb") as f:
        st.download_button(
            "Download Excel",
            f,
            file_name="RITM_Output.xlsx"
        )