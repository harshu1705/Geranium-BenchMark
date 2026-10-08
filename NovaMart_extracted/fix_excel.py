import openpyxl
import os

path = r'c:\Users\user\OneDrive\Desktop\geranium project\NovaMart_extracted\Inputs\Third_Party_Security_Scan.xlsx'
wb = openpyxl.load_workbook(path)
ws = wb.active

for row in ws.iter_rows(min_row=1, max_col=7, max_row=ws.max_row):
    # Just print the first run
    print([cell.value for cell in row])

    # Now modify IP addresses
    if row[1].value and isinstance(row[1].value, str) and row[1].value.startswith('192.168.1.'):
        new_ip = row[1].value.replace('192.168.1.', '203.0.113.')
        row[1].value = new_ip
    
    # Fix CVE-2024-3094
    if row[3].value == 'CVE-2024-3094':
        row[4].value = 'SSH'
        row[5].value = 22

wb.save(path)
print("Updated Third_Party_Security_Scan.xlsx successfully.")
