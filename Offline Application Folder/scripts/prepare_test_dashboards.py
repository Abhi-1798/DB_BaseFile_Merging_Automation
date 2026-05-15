import pandas as pd
import xlwings as xw
import os
import json

def prepare_dashboards():
    config_path = r'D:\DB Merging Automation\DB Merging Automation\config.json'
    with open(config_path, 'r') as f:
        config = json.load(f)
    
    dashboards = [config['export_dashboard_path'], config['import_dashboard_path']]
    headers = [
        "Date", "Country", "Region", "Business Unit", 
        "Product Category", "SKU Name", "Transaction Type", 
        "Currency", "Quantity", "Value USD"
    ]
    
    app = xw.App(visible=False)
    for path in dashboards:
        print(f"Preparing {path}...")
        dir_name = os.path.dirname(path)
        if not os.path.exists(dir_name):
            os.makedirs(dir_name)
        
        # Remove existing file if it exists to ensure a clean slate
        if os.path.exists(path):
            try:
                os.remove(path)
            except Exception as e:
                print(f"Warning: Could not remove existing file {path}. It might be open. {e}")
            
        # Create new workbook
        wb = app.books.add()
        sheet = wb.sheets[0]
        sheet.name = "Data"
        
        # Add headers
        sheet.range('A1').value = headers
        
        # Professional Dashboard Formatting
        header_range = sheet.range('A1:J1')
        header_range.api.Font.Bold = True
        header_range.color = (31, 73, 125)  # Professional Dark Blue
        header_range.font.color = (255, 255, 255) # White
        header_range.api.HorizontalAlignment = -4108 # Center
        
        # Auto-fit columns
        sheet.range('A1:J1').columns.autofit()
        
        # Freeze panes (Row 1)
        app.api.ActiveWindow.FreezePanes = False
        sheet.range('A2').select()
        app.api.ActiveWindow.FreezePanes = True
        
        wb.save(path)
        wb.close()
    
    app.quit()
    print("Dashboards successfully reset and formatted for final testing.")

if __name__ == "__main__":
    prepare_dashboards()
