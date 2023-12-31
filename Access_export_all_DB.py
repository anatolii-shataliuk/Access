import pandas as pd
import pyodbc
from datetime import datetime
import os, openpyxl
from openpyxl import load_workbook

# Database file path and password
db_file_path = 'S:/Штатка tabl.accdb'
password = 'zknafein1990'
today = datetime.now().strftime("%d.%m.%Y_%H%M")

# Establish connection
conn_str = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    rf"DBQ={db_file_path};"
    f"PWD={password};"
)
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# Get tables from the database
tables = cursor.tables(tableType='TABLE')
tables_to_export = [table.table_name for table in tables]

# Create an Excel writer object
excel_file_path = 'База ' + today + '.xlsx'  # Define your output Excel file path
excel_writer = pd.ExcelWriter(excel_file_path, engine='xlsxwriter')

# Export tables to Excel
for table in tables_to_export:
    query = f"SELECT * FROM [{table}]"
    df = pd.read_sql(query, conn)
    df.to_excel(excel_writer, sheet_name=table, index=False)

# Save and close the Excel writer
excel_writer.save()
conn.close()

#Del first sheet of workbook
wb = load_workbook(excel_file_path)
#if 'log' in wb.get_sheet_names:
wb.remove(wb['log'])
wb.save(excel_file_path)


print("Exported " + excel_file_path)
exit()