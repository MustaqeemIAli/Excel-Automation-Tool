# Excel Automation Tool

A Python-based **Excel/CSV automation and data consolidation tool** with a clean **PyQt5 GUI**.  
It simplifies merging, cleaning, and categorizing financial transactions from multiple monthly files into a single **main Excel file** — with full automation, duplicate detection, and pivot table refresh.  

---
<img width="499" height="678" alt="Excel-Automation" src="https://github.com/user-attachments/assets/8a514db7-9bd8-492e-8854-1452822053e1" />

## Features
- **User-Friendly GUI** – Select a main Excel file, add multiple monthly files, and process them with one click.  
- **Smart File Handling** – Supports `.xlsx`, `.xls`, and `.csv`, detects headers automatically, and validates missing columns.  
- **Data Cleaning & Standardization**  
  - Unifies inconsistent headers across files.  
  - Detects date and amount formats automatically.  
  - Combines multiple description columns safely.  
- **Automatic Classification** – Assigns categories/subcategories using rules, frequency history, and fuzzy matching.  
- **Duplicate Detection** – Flags and skips already existing transactions.  
- **Return & Check Handling** – Identifies returned checks and pairs related transactions.  
- **Backups & Safety** – Automatically creates a backup of the original file before any changes.  
- **Excel Output**  
  - Appends new transactions with correct formatting.  
  - Highlights “Need Review” and “Duplicate” rows with colors.  
  - Auto-generates Month column via formula.  
  - Refreshes all Pivot Tables instantly.  

---

## Why Create This Tool?
**Save Hours of Work** – No more manual copy-paste between monthly and main Excel sheets.  
**Reduce Errors** – Automated classification, duplicate checks, and formatting consistency.  
**Boost Productivity** – Process all files in one go and get updated pivot tables instantly.  
**Ensure Consistency** – Standardized categories and descriptions across multiple reports.  

---

## Tech Stack
- **Python 3**  
- **PyQt5** – GUI framework  
- **Pandas** – Data manipulation 
- **OpenPyXL / xlrd** – Excel reading/writing  
- **xlwings**  
- **Chardet** 
- **Requests** 

---

## Contact

For support, licensing inquiries, or customization requests, please contact:

**Mustaqeem Ali**  
mustaqeemimtiazali@gmail.com

**Luqman Ali**
luqmanmoizali@gmail.com  
