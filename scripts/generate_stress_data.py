import pandas as pd
import numpy as np
import os
import xlwings as xw
import uuid

BASE_DIR = r"D:\DB Merging Automation\DB Merging Automation"
DATA_DIR = os.path.join(BASE_DIR, "data", "base_files")
MAPPER_PATH = os.path.join(BASE_DIR, "Column_Mapper.xlsx")

def generate_large_file(filename, key, date_type, row_count=1000000):
    print(f"Generating massive {filename} ({row_count} rows)...")
    
    cols = {}
    if date_type == 'Split':
        cols = {"Year_Col": "Year", "Month_Col": "Month", "Country": "Region", "Material": "Product", "Type": "Type", "Value": "USDValue"}
    else:
        cols = {"Date_Col": "TransactionDate", "Country": "Area", "Material": "Item", "Type": "Direction", "Value": "TotalAmount"}
    
    headers = list(cols.values())
    # Add many columns to increase size index
    headers += [f"Unique_ID_{i}" for i in range(1, 10)]
    headers += [f"Random_Category_{i}" for i in range(1, 10)]
    headers += ["BusinessUnit", "Category", "Currency", "Quantity"]

    # --- High-Performance Date Range Generation ---
    start_date = pd.to_datetime("2021-01-01")
    end_date = pd.to_datetime("2026-03-31")
    
    # Generate random timestamps within the range
    random_timestamps = np.random.randint(start_date.value, end_date.value, row_count, dtype=np.int64)
    random_dates = pd.to_datetime(random_timestamps)
    
    # Cached month mapping for speed
    month_map = {1:"Jan", 2:"Feb", 3:"Mar", 4:"Apr", 5:"May", 6:"Jun", 7:"Jul", 8:"Aug", 9:"Sep", 10:"Oct", 11:"Nov", 12:"Dec"}

    data = {}
    for h in headers:
        if h == "Year": 
            data[h] = random_dates.year
        elif h == "Month": 
            # Vectorized mapping is much faster than .strftime for 1M rows
            data[h] = random_dates.month.map(month_map)
        elif h in ["TransactionDate", "Doc_Date"]: 
            data[h] = random_dates.strftime('%Y-%m-%d')
        elif h in ["USDValue", "TotalAmount"]: 
            data[h] = np.random.uniform(100, 10000, size=row_count)
        elif "Unique_ID" in h: 
            data[h] = [str(uuid.uuid4())[:8] for _ in range(row_count)]
        elif "Random_Category" in h:
            data[h] = np.random.choice(["Red", "Blue", "Green", "Yellow", "Purple", "Orange"], size=row_count)
        else: 
            data[h] = np.random.choice(["Standard", "Priority", "Express", "Slow", "Bulk"], size=row_count)
            
    df = pd.DataFrame(data)
    file_path = os.path.join(DATA_DIR, filename)
    
    print(f"Writing {filename} to disk...")
    df.to_excel(file_path, index=False, engine='openpyxl') # openpyxl is safer for very large files on some systems
        
    size_mb = os.path.getsize(file_path) / (1024*1024)
    print(f"Done: {filename}. Size: {size_mb:.2f} MB")
    return {"key": key, "path": file_path, "type": date_type, "cols": cols}

def update_mapper(configs):
    app = xw.App(visible=False)
    try:
        wb = app.books.open(MAPPER_PATH)
        sheet = wb.sheets['Export'] 
        
        last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
        if last_row > 1:
            sheet.range(f'A2:Q{last_row}').clear_contents()
            
        curr_row = 2
        for cfg in configs:
            row_data = [
                cfg['key'], "Sheet1", "Data", cfg['type'],
                cfg['cols'].get('Date_Col', ""), cfg['cols'].get('Year_Col', ""), cfg['cols'].get('Month_Col', ""),
                cfg['cols'].get('Date_Col', ""),
                cfg['cols'].get('Country'), "Region_Name", "BusinessUnit_Name", "Category_Name", cfg['cols'].get('Material'),
                cfg['cols'].get('Type'), "Currency", "Quantity", cfg['cols'].get('Value')
            ]
            sheet.range(f"A{curr_row}").value = row_data
            curr_row += 1
            
        wb.save()
        wb.close()
    finally:
        app.quit()

if __name__ == "__main__":
    # Create 1 Massive file at the Excel limit
    c1 = generate_large_file("Stress_Large_Excel_Limit.xlsx", "STRESS_MAX", "Split", 1000000)
    update_mapper([c1])
