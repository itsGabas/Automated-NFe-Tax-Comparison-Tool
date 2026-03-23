# 📊 Automated NFe Tax Comparison Pipeline (Python)

Internal automation project developed to compare tax data from Brazilian electronic invoices (NFe) across two systems, reducing manual effort and improving accuracy.

---

## 🎯 Objective
Automate the comparison of tax values between two financial systems, enabling faster validation and reducing human error.

---

## 🧠 Problem
The fiscal team needed to manually compare 11+ tax fields (PIS, COFINS, ICMS, etc.) between two Excel reports from different systems.

This process was:
- Time-consuming  
- Repetitive  
- Error-prone  

---

## 🔄 Data Pipeline (ETL Approach)

This solution was structured as a data pipeline:

### 1️⃣ Data Cleaning
- Removed unnecessary columns  
- Standardized column names  
- Filtered only relevant tax fields  

### 2️⃣ Data Preparation
- Structured datasets into a consistent format  
- Ensured compatibility between both systems  
- Aligned invoice records for accurate comparison  

### 3️⃣ Data Comparison & Reporting
- Matched invoices using NFe number  
- Compared 11+ tax fields  
- Generated automated Excel report with:
  - ✅ Matching values (green)
  - ❌ Discrepancies (red)
  - 💬 Comments with both system values  

---

## 📊 Output
- Structured datasets ready for analysis  
- Automated comparison results  
- Excel report highlighting inconsistencies  

---

## 💼 Business Impact
- Reduced manual comparison time  
- Improved accuracy in tax validation  
- Increased productivity of the fiscal team  
- Faster identification of discrepancies  

---

## 🛠️ Tech Stack
Python • Pandas • Excel Processing • Data Cleaning • Automation  

---

## ⚠️ Limitations
- Inconsistent comment generation in some cases  
- Misalignment when reports have different row counts  
- Edge cases with missing or unordered invoices  

---

## 🔧 Future Improvements
- Improve matching logic for robustness  
- Add validation and alignment checks  
- Enhance logging and error handling  

---

## 🤖 Use of AI Tools
AI-assisted tools were used to accelerate development and learning.  
All logic and business rules were reviewed and validated manually.

---

## 📂 Context
- Internal automation project  
- Real-world accounting use case  
- Applied in a production-like environment  

---

## 🔎 Why this project matters
- Real-world data pipeline (ETL)  
- Financial data validation experience  
- Automation of repetitive processes  
- Focus on accuracy and data quality  
