# SAP-Duplicate-Invoice-Automation
Automated solution using Excel VBA and SAP GUI Scripting to extract, cross-reference, and bulk block duplicate invoices directly in SAP.

# SAP Duplicate Invoice Automation & Payment Control 🚀

## 📌 Project Overview
This project is an automated solution built with **Excel VBA** and **SAP GUI Scripting** designed to overhaul the Duplicate Invoice checking process. It replaces a heavy, multi-step manual workflow with a centralized, 5-click automated interface.

## 🛠️ Tech Stack
* **Language:** VBA (Visual Basic for Applications)
* **Integration:** SAP GUI Scripting API
* **Platform:** Microsoft Excel

## 🛑 The Problem (Manual DTP)
Previously, the process required:
* Manually executing the `D_FIAPDUPINVREP` job in SAP via SM37.
* Downloading multiple spool files and reports (FBL1N, VIM).
* Performing complex, error-prone manual VLOOKUPs across different data silos.
* Manually re-entering SAP to apply payment blocks item by item.

## ✅ The Automated Solution
The VBA tool connects directly to the active SAP session to:
1. **Retrieve Data:** Automates the extraction of raw text files from SAP.
2. **Clean & Match:** Algorithms process the text files, removing unnecessary data and identifying exact duplicate matches.
3. **Assign Claims:** Automatically generates sequential Claim Numbers for full traceability.
4. **Bulk Block in SAP:** Loops through identified unpaid duplicates and autonomously applies the "B" payment block in SAP.

## 📈 Business Impact
* **Efficiency:** Reduced a process taking hours of daily manual effort to just minutes.
* **Accuracy:** Completely eliminated data entry and formatting errors.
* **Compliance:** Improved the integrity of large-scale payment runs by ensuring duplicates are blocked before payment cycles.

## 📂 Repository Structure
* `/src`: Contains the raw `.bas` VBA modules for code review.
* `/assets`: Visual workflow improvements (Before & After).
* `/docs`: Original manual process documentation (DTP).
