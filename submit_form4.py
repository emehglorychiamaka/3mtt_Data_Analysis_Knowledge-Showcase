# import libraries

import pyodbc
from datetime import datetime
import ctypes
import sys
import pythoncom
from win32com.client import Dispatch

# STEP 1: Access the already open Excel workbook via COM
try:
    pythoncom.CoInitialize()
    xl = Dispatch("Excel.Application")
    wb = xl.Workbooks("MaternalHealth_Risk.xlsm")  # must be open
    ws = wb.Sheets("Form")

    form_data = [ws.Range(f"B{row}").Value for row in range(4, 13)]  # B4 to B12
except Exception as e:
    ctypes.windll.user32.MessageBoxW(0, f"❌ Error accessing Excel: {e}", "Excel Access Error", 0)
    sys.exit()

# STEP 2: Validate and extract
try:
    surname    = str(form_data[0]).strip()
    names      = str(form_data[1]).strip()
    age        = int(form_data[2])
    systolic   = float(form_data[3])
    diastolic  = float(form_data[4])
    bs         = float(form_data[5])
    bodytemp   = float(form_data[6])
    heartrate  = float(form_data[7])
    visit_raw  = form_data[8]

    if not all([surname, names, age, systolic, diastolic, bs, bodytemp, heartrate, visit_raw]):
        raise ValueError("All form fields (B4:B12) must be filled.")

    # Convert visit date
    if isinstance(visit_raw, datetime):
        visit_date = visit_raw.date()
    else:
        visit_date = datetime.strptime(str(visit_raw), "%Y-%m-%d").date()

except Exception as e:
    ctypes.windll.user32.MessageBoxW(0, f"❌ Form input error: {e}", "Input Error", 0)
    sys.exit()

# STEP 3: Predict
if systolic > 140 and diastolic > 90 and bs > 120:
    risk = "High risk"
elif systolic > 120 or bs > 110:
    risk = "Mid risk"
else:
    risk = "Low risk"

# STEP 4: Database insert
conn_str = (
    r"DRIVER={ODBC Driver 17 for SQL Server};"
    r"SERVER=Emeh_GLory_C\SQLEXPRESS;"
    r"DATABASE=Maternal_Health_Risk;"
    r"Trusted_Connection=yes;"
)

try:
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    cursor.execute("""
    IF NOT EXISTS (
        SELECT * FROM sysobjects WHERE name='MaternalHealth_Extended' AND xtype='U'
    )
    CREATE TABLE MaternalHealth_Extended (
        ID INT IDENTITY(1,1) PRIMARY KEY,
        Surname NVARCHAR(50),
        Names NVARCHAR(50),
        Age INT,
        SystolicBP FLOAT,
        DiastolicBP FLOAT,
        BS FLOAT,
        BodyTemp FLOAT,
        HeartRate FLOAT,
        Visit_date DATE,
        RiskLevel NVARCHAR(20)
    )
    """)
    conn.commit()

    cursor.execute("""
        INSERT INTO MaternalHealth_Extended
        (Surname, Names, Age, SystolicBP, DiastolicBP, BS, BodyTemp, HeartRate, Visit_date, RiskLevel)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        surname, names, age, systolic, diastolic, bs, bodytemp, heartrate, visit_date, risk
    )
    conn.commit()
    conn.close()
except Exception as e:
    ctypes.windll.user32.MessageBoxW(0, f"❌ DB error: {e}", "Database Error", 0)
    sys.exit()

# STEP 5: Write back result and clear form (via COM)
try:
    ws.Range("B13").Value = f"Prediction Risk Level result: {risk}"

    for row in range(4, 13):
        ws.Range(f"B{row}").Value = ""

except Exception as e:
    ctypes.windll.user32.MessageBoxW(0, f"❌ COM write error: {e}", "Excel Write Error", 0)
    sys.exit()

# STEP 6: Confirmation
ctypes.windll.user32.MessageBoxW(
    0,
    f" Record submitted!\n Predicted Risk Level: {risk}",
    "Submission Successful",
    0
)
