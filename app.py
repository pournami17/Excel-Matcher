# Reads the PDF exported from a monday.com form

# Extracts FAQ question–answer pairs

# Reads the Microsoft Excel Template file

# Matches similar questions using AI embeddings

# Writes the best answer into the Excel sheet

# Saves the final Excel file

import streamlit as st
import pdfplumber
import pandas as pd
from sentence_transformers import SentenceTransformer, util
import tempfile

# ----------------------------
# Streamlit App Title
# ----------------------------
st.title("PDF to FAQ Excel Matcher")

# ----------------------------
# File Upload
# ----------------------------
pdf_file = st.file_uploader("Upload the Monday Board form in PDF format", type=["pdf"])
excel_file = st.file_uploader("Upload the Excel Template", type=["xlsx"])

# ----------------------------
# Load model
# ----------------------------
@st.cache_resource
def load_model():
    return SentenceTransformer("all-MiniLM-L6-v2")

model = load_model()

# ----------------------------
# Process files
# ----------------------------
if pdf_file and excel_file:

    # Save uploaded PDF temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        temp_pdf.write(pdf_file.read())
        PDF_FILE = temp_pdf.name

    # Save uploaded Excel temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_excel:
        temp_excel.write(excel_file.read())
        EXCEL_FILE = temp_excel.name

    st.info("Processing files... This may take a minute depending on file size.")

    # ----------------------------
    # Extract text from PDF
    # ----------------------------
    text = ""
    with pdfplumber.open(PDF_FILE) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"

    lines = [line.strip() for line in text.split("\n") if line.strip()]

    # ----------------------------
    # Extract Property Name from PDF
    # ----------------------------
    property_name = ""
    for i, line in enumerate(lines):
        if "Property Name" in line:
            if i + 1 < len(lines):
                property_name = lines[i + 1].strip()
            break

    if not property_name:
        st.warning("Property Name not found in PDF. Defaulting to 'Unknown Property'.")
        property_name = "Unknown Property"

    st.info(f"Property Name detected: {property_name}")

    # ----------------------------
    # Extract FAQ pairs
    # ----------------------------
    faq_pairs = []
    current_question = None
    current_answer = []

    for line in lines:
        if line.endswith("?"):
            if current_question and current_answer:
                faq_pairs.append((current_question, " ".join(current_answer)))
            current_question = line
            current_answer = []
        else:
            if current_question:
                current_answer.append(line)

    if current_question and current_answer:
        faq_pairs.append((current_question, " ".join(current_answer)))

    faq_df = pd.DataFrame(faq_pairs, columns=["Question", "Answer"])

    if faq_df.empty:
        st.error("No FAQs detected. Check PDF format.")
    else:
        st.success(f"Extracted {len(faq_df)} FAQ pairs from PDF")

        # ----------------------------
        # Read Excel
        # ----------------------------
        excel_df = pd.read_excel(EXCEL_FILE)
        if "Question" not in excel_df.columns:
            st.error("Excel file must contain a 'Question' column")
        else:
            excel_questions = excel_df["Question"].tolist()

            # ----------------------------
            # Generate embeddings
            # ----------------------------
            faq_embeddings = model.encode(faq_df["Question"].tolist(), convert_to_tensor=True)
            answers = []

            # ----------------------------
            # Match questions
            # ----------------------------
            for q in excel_questions:
                q_embedding = model.encode(q, convert_to_tensor=True)
                scores = util.cos_sim(q_embedding, faq_embeddings)
                best_idx = scores.argmax()
                best_score = scores.max()
                best_answer = faq_df.iloc[int(best_idx)]["Answer"] if best_score > 0.55 else ""
                answers.append(best_answer)

            excel_df["Answer"] = answers

            # ----------------------------
            # Save & download with extracted property name
            # ----------------------------
            output_file = f"{property_name} – Jonah Import FAQ.xlsx"
            excel_df.to_excel(output_file, index=False)

            st.success(f"Processing complete! Output file: {output_file}")
            st.download_button(
                label="Download Excel with Answers",
                data=open(output_file, "rb"),
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )