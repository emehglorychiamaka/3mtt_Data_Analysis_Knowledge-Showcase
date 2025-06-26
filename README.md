# 3mtt_Data_Analysis_Knowledge-Showcase
An AI-powered system that predicts maternal health risk levels (Low, Mid, High) using vital signs input from an Excel form, with results stored in a SQL Server database.

 Features
 Excel-based form for easy data entry (no technical skills needed)

 VBA button triggers a Python script invisibly (pythonw.exe)

 AI model classifies risk level in real-time

 Results are saved to MSSQL and written back to Excel

 Form auto-clears after each submission

 Tech Stack
Python (with openpyxl, pyodbc)

Excel (.xlsm with VBA)

MSSQL (local server)

Trained ML model (classification)

📂 Folder Structure

3MTT_Data_Analysis/
│
├── MaternalHealth_Risk.xlsm     # Excel input form
├── submit_form.py               # Python script for reading input, prediction & DB write
├── submit_form.bat              # Batch file to run Python script silently

 How It Works
Fill form in Excel (B4:B12)

Click "Submit"

AI predicts risk level

Result is displayed in Excel and logged in database

 Credits
Developed by Glory Chiamaka Emeh as part of the
#My3MTT project — AI Category
 #3MTTLearningCommunity | @3MTTNigeria
