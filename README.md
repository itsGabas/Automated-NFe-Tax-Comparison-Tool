# Automated NFe Tax Comparison Tool

Internal automation project developed to **compare tax data from Brazilian electronic invoices (NFe)** across two different reporting systems, reducing manual effort and minimizing errors in tax validation.

---

## 📌 Project Overview

In the accounting firm where I currently work, the tax (fiscal) team faced significant challenges when comparing **multiple tax values** (such as PIS, COFINS, ICMS, and others) across two separate systems.

The comparison process involved:
- Exporting an Excel report from the **Domínio** system
- Exporting a second Excel report from an external tax validation tool
- Manually comparing **11+ tax fields** for each NFe across both reports

This manual process was time-consuming and prone to human error.

---

## 🎯 Objective

The goal of this project was to **automate the comparison process**, allowing the team to:
- Quickly identify discrepancies
- Focus only on invoices with inconsistencies
- Reduce repetitive manual work
- Improve accuracy and productivity

---

## ⚙️ How It Works

1. Two Excel reports are provided as input:
   - Report A: Exported from Domínio
   - Report B: Exported from the external tax tool

2. Each report contains:
   - One row per NFe
   - One column per tax type (PIS, COFINS, ICMS, etc.)

3. The script:
   - Matches invoices using the **NFe number**
   - Compares **11+ tax values** for each matched invoice
   - Generates a new Excel file with:
     - Rows grouped by NFe number
     - **Green cells** when values match
     - **Red cells** when discrepancies are found
     - Cell comments showing values from both systems for easier validation

4. The fiscal team reviews only the highlighted discrepancies and manually corrects values in the internal system when needed.

---

## 🧠 Impact & Results

- Significant reduction in manual comparison time
- Improved visibility of tax inconsistencies
- Faster identification of invoices requiring correction
- Increased productivity for the fiscal team

---

## 🛠️ Technologies Used

- **Python**
- **Pandas**
- **Excel file manipulation**
- **Conditional formatting**
- **Automated data comparison and validation**

---

## ⚠️ Known Limitations & Edge Cases

During testing and real-world usage, the following limitations were identified:

- In some scenarios, **cell comments were not consistently generated**, even when discrepancies were correctly detected and highlighted.
- When input reports contained **different numbers of rows**, some invoice comparisons could become misaligned, leading to incorrect matches.
- Certain edge cases related to data ordering and missing invoices were observed and logged during execution.

These limitations were documented and partially traced through application logs, highlighting opportunities for improved data alignment, validation, and error handling.

---

## 🔧 Future Improvements

- Improve invoice matching logic to handle missing or extra records more robustly
- Add stricter validation and alignment checks before comparisons
- Enhance error handling and logging for easier debugging
- Refactor comment generation to ensure consistent annotation behavior

---

## 🤖 Use of AI Tools

This project was developed with the support of **AI-assisted tools** to accelerate learning and problem-solving, as it was one of my first automation projects of this nature.

AI was used as a **support tool**, while all comparison logic, validations, adjustments, and business rules were actively reviewed and validated to ensure correctness.

---

## 📂 Project Context

- Internal automation project
- Developed to address a real operational challenge
- Used in a production-like environment to support daily workflows
- Focused on improving efficiency and data accuracy

---

## 🔎 Why this project matters

This project demonstrates:
- Practical application of Python for data comparison and reporting
- Experience working with real-world financial and tax datasets
- Ability to identify system limitations and edge cases
- Transparency, documentation, and continuous improvement mindset
