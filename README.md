# PDF to FAQ Excel Matcher

This app allows you to automatically extract FAQ question-answer pairs from a **PDF** and match them to an **Excel template**. The output is an Excel file named using the **Property Name** from the PDF.

---

## **Features**

- Reads a **PDF exported from a Monday.com form**.
- Automatically extracts the **Property Name**.
- Extracts **FAQ question-answer pairs**.
- Matches Excel questions with PDF FAQs using **AI embeddings** (`sentence-transformers`).
- Generates a new Excel with matched answers:

```
[Property Name] – Jonah Import FAQ.xlsx
```

- Hosted as a **Streamlit app** for easy access by teammates.

---

## **Requirements**

- Python 3.9+
- Libraries (see `requirements.txt`):

```
streamlit
pdfplumber
pandas
sentence-transformers
openpyxl
torch
```

---

## **Setup**

1. Clone this repository:

```bash
git clone <your-repo-url>
cd <repo-folder>
```

2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Run the app locally:

```bash
streamlit run app.py
```

---

## **Usage**

1. Upload the **Monday PDF** form submission.
2. Upload the **Excel template** (must have a column named `Question`).
3. The app will automatically:  
   - Extract the property name  
   - Extract FAQ question-answer pairs  
   - Match the Excel questions with FAQs
4. Click **Download Excel with Answers** to get the output file:

```
[Property Name] – Jonah Import FAQ.xlsx
```

---