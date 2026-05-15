# VAT-Issuance-Support

> 📂 **[VIEW FULL PROJECT PRESENTATION (PDF)](presentation/VAT_Issuance_Support.pdf)**

# 📊 Automated VAT Invoicing Data Transformation Tool

### 📌 Project Overview
This project automates the process of merging **Daily Sales Transaction Data** with **Customer VAT Profiles**. The goal is to transform raw POS data into a standardized format ready for direct import into tax declaration software, ensuring 100% accuracy in tax reporting and compliance.

### 🛠️ Key Data Logic
The script processes and reconciles data through several stages:

*   **Data Ingestion:** Imports daily sales reports and system master lists via authenticated Google Sheets CSV export links for centralized processing.
*   **Intelligent Mapping:** Merges internal item names with official VAT-compliant descriptions using string normalization (lowercasing, stripping whitespace) to ensure a perfect join.
*   **Tax Calculation & Categorization:** Dynamically assigns VAT rates (8% or 10%) and maps items to their respective accounting groups.
*   **Output Optimization:** Cleans, deduplicates, and reorders columns to match the specific schema required by tax declaration software.

### 💻 Tech Stack
*   **Language:** Python
*   **Libraries:** Pandas, NumPy, Openpyxl.
*   **Platform:** Google Colab / Jupyter Notebook.

### 📈 Business Value
*   **Speed:** Transforms hours of manual data reconciliation into a sub-minute automated task.
*   **Compliance:** Eliminates manual entry errors, ensuring the exported data matches legal VAT requirements.
*   **Integration:** Designed as a bridge between internal POS systems and external tax filing platforms.

---
*Note: All data displayed in screenshots has been anonymized or modified to protect business confidentiality.*
