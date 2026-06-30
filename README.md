# Duplicate Question Detector

A web app to detect exact and near-duplicate questions in question bank PDFs.

## Live App

https://duplicate-question-detector-xawkvzisyskp4lecuatrvh.streamlit.app/

## What it does

- Upload any question bank PDF (supports 1000+ questions)
- Detects **exact duplicates** and **near-duplicates** (different wording, different options, fill-in-blank vs MCQ versions of the same question)
- Works for any subject — Physics, Chemistry, Math, Biology, etc.
- Outputs a results table with similarity scores and page numbers
- Download results as **Excel** or **Text** report

## How to run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Output format

Each duplicate group is listed as:

```
N. QX and QY (Pages A & B): [question description].
```

The Excel report has two sheets:
- **Summary** — all duplicate groups with similarity scores
- **Question Details** — full question text for each group

## How it works

1. Extracts text from PDF using `pdfplumber` (handles two-column layouts)
2. Parses questions by number pattern (Q1, Q2, ...)
3. Normalizes text — strips MCQ labels, blanks, watermarks, punctuation
4. Detects exact duplicates via MD5 hash
5. Detects near-duplicates via `rapidfuzz` fuzzy matching
6. Groups overlapping pairs into clusters using union-find
