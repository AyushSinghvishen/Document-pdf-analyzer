import os
import pdfplumber
from docx import Document
import pandas as pd

pdf_folder = "pdf_reports"
docx_folder = "docx_reports"
output_folder = "output"
os.makedirs(output_folder, exist_ok=True)

all_text = ""
all_tables = []

def clean_words(text):
    punctuation = ['.', ',', '!', '?', ':', ';', '(', ')', '[', ']', '{', '}', '"', "'"]
    words = text.lower().split()
    cleaned = []
    for w in words:
        for p in punctuation:
            w = w.strip(p)
        if w:
            cleaned.append(w)
    return cleaned

# Extract from PDFs
for file in os.listdir(pdf_folder):
    if file.endswith(".pdf"):
        with pdfplumber.open(os.path.join(pdf_folder, file)) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                all_text += text + "\n"
                tables = page.extract_tables()
                for table in tables:
                    df = pd.DataFrame(table[1:], columns=table[0])
                    all_tables.append(df)

# Extract from DOCX
for file in os.listdir(docx_folder):
    if file.endswith(".docx"):
        doc = Document(os.path.join(docx_folder, file))
        for para in doc.paragraphs:
            all_text += para.text + "\n"

# Word frequency without .get() or lambda
words = clean_words(all_text)
freq = {}

# Count each word manually
for w in words:
    if w in freq:
        freq[w] += 1
    else:
        freq[w] = 1

# Convert dict to list of (word, count) tuples
word_count_list = []
for word in freq:
    word_count_list.append((word, freq[word]))

# Sort the list by count descending (simple sort)
for i in range(len(word_count_list)):
    for j in range(i + 1, len(word_count_list)):
        if word_count_list[j][1] > word_count_list[i][1]:
            word_count_list[i], word_count_list[j] = word_count_list[j], word_count_list[i]

# Get top 20 words
top_words = word_count_list[:20]

# Save word frequency to CSV
df_freq = pd.DataFrame(top_words, columns=["Word", "Count"])
df_freq.to_csv(os.path.join(output_folder, "word_frequency.csv"), index=False)

# Save extracted tables to Excel
with pd.ExcelWriter(os.path.join(output_folder, "extracted_tables.xlsx"), engine='openpyxl') as writer:
    for i, df in enumerate(all_tables):
        df.columns = [str(c).strip() for c in df.columns]
        df.to_excel(writer, sheet_name=f"Table_{i+1}", index=False)

print("Done! Files saved in 'output' folder.")
