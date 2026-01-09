# SQL to Excel – Monthly SFTP vs GTO Reconciliation Automation

This project is a C# console application that automates
monthly reconciliation between **SFTP** and **GTO** values
for multiple outlets and exports the results into a
formatted Excel report.

---

## 🔹 Business Use Case
- Monthly financial reconciliation
- Store-wise SFTP vs GTO comparison
- Automated Excel report generation
- Reduces manual reconciliation effort

---

## 🔹 Key Features
- Executes SQL Server stored procedure per outlet
- Generates a **month-based Excel report**
- Store-wise grouped layout with borders
- Calculates daily difference and totals
- Auto-freezes headers for easy navigation
- Fully automated using C#

---

## 🔹 Technologies Used
- C#
- SQL Server
- Stored Procedures
- ClosedXML (Excel generation)
- ADO.NET

---

## 🔹 Report Highlights
- Month displayed at the top
- Each store shown in a separate block
- Columns: Date | SFTP | GTO | Difference
- Total row per store
- Auto-adjusted column widths
- Excel output ready for finance review

---

## 🔹 Project Structure
