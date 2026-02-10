# Excel Data Consolidation Tool (VBA)

A simple and reusable Excel VBA utility to merge multiple Excel files into a single consolidated workbook.

This tool is designed for real-world data operations where:
- Multiple Excel files exist in a folder  
- Each file contains multiple sheets  
- Sheet names may repeat across files  
- Column headers may appear in different orders  
- Data needs to be appended and aligned correctly  

---

## 🚀 Features

- Select a folder and automatically merge all Excel files
- Combine same sheet names into a single `_Combined` sheet
- Align columns by header name (even if column order is shuffled)
- Append all rows safely without overwriting original files
- Fully dynamic: supports any number of files, sheets, rows, and columns
- Replaces old combined file if already present

---

## 🧠 Use Case

This utility is useful for:
- Data Analysts
- Finance & Operations teams
- Reporting automation
- Month-end consolidation
- Handling ERP / system exports

Instead of manually copying and pasting data from multiple files, this tool automates the entire process in one click.

---

## 🛠 How to Use

1. Open Excel  
2. Press `Alt + F11`  
3. Insert → Module  
4. Paste the VBA code from `Combine_Excel_Files.bas`  
5. Run the macro  
6. Select the folder containing Excel files  
7. The combined file will be created as `Combined_Excel.xlsx`

---

## 📸 Demo

(Include a screenshot of before & after)

---

## 📌 Why this project?

This project demonstrates:
- Practical automation skills
- Understanding of real business data problems
- Clean and reusable VBA coding practices
- Data consolidation and transformation logic

---

## 📬 Author

Created by [Sukesh S]  
Data Analytics | Automation | Excel | VBA  
