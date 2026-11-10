import csv
import json
import os
import openpyxl

folder_path = r"C:\Users\Administrator\Downloads\23PT17\Ques(csv)\csv"

all_files = os.listdir(folder_path)
converted_count = 0

for file_name in all_files:
    # Skip temporary Excel lock files
    if file_name.startswith('~$'):
        continue
        
    file_path = os.path.join(folder_path, file_name)
    base_name, extension = os.path.splitext(file_name)
    extension = extension.lower()
    
    json_file_name = f"{base_name}.json"
    json_file_path = os.path.join(folder_path, json_file_name)
    
    if extension == '.csv':
        print(f"Processing CSV: {file_name}")
        with open(file_path, mode='r', encoding='latin-1') as csv_file:
            csv_reader = csv.DictReader(csv_file)
            data = list(csv_reader)
            
    # Handle Excel XLSX files
    elif extension == '.xlsx':
        print(f"Processing XLSX: {file_name}")
        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheet = wb.active  # Reads the first active sheet
        
        headers = [cell.value for cell in sheet[1]]
        
        data = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not any(row):
                continue
            row_dict = {headers[i]: row[i] for i in range(len(headers))}
            data.append(row_dict)
            
    else:
        continue

    # Save data to JSON
    with open(json_file_path, mode='w', encoding='utf-8') as json_file:
        json.dump(data, json_file, indent=4)
    converted_count += 1

print(f"\nSuccess! Total JSON files generated: {converted_count}")