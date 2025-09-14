#PDF & DOCX Analyzer  

This project extracts **text** and **tables** from PDF and DOCX files, cleans the text, and performs **word frequency analysis**. It saves results as:  
- `word_frequency.csv` → Top 20 words  
- `extracted_tables.xlsx` → All tables  

##  Usage  
1. Add PDFs to `pdf_reports/` and DOCX files to `docx_reports/`  
2. Run:  
   ```bash
   python main.py

