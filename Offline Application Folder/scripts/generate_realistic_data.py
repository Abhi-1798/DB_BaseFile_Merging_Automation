import pandas as pd
import os
import xlwings as xw

BASE_DIR = r"D:\DB Merging Automation\DB Merging Automation"
DATA_DIR = os.path.join(BASE_DIR, "data", "base_files")
MAPPER_PATH = os.path.join(BASE_DIR, "Column_Mapper.xlsx")

if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR)

# Logic to generate realistic demo files
def generate_files():
    # 6 files with unique variations
    configs = [
        {
            "filename": "Arg_Sales_Split.xlsx",
            "key": "Test_Argentina",
            "date_type": "Split",
            "cols": {
                "Year_Col": "Year", "Month_Col": "Month", 
                "Country": "Nacion", "Region": "Mercado", "Business_Unit": "Division", 
                "Product_Category": "Grupo", "SKU_Name": "Articulo", "Transaction_Type": "Sentido",
                "Currency": "Moneda", "Quantity": "Kilos", "Value_USD": "Venta_USD"
            }
        },
        {
            "filename": "Bra_Sales_Direct.xlsx",
            "key": "Test_Brazil",
            "date_type": "Direct",
            "cols": {
                "Date_Col": "Data_Relatorio", 
                "Country": "Estado", "Region": "Cluster", "Business_Unit": "Unidade", 
                "Product_Category": "Familia", "SKU_Name": "Descricao", "Transaction_Type": "Tipo",
                "Currency": "Simbolo", "Quantity": "Volume", "Value_USD": "Valor_Dolar"
            }
        },
        {
            "filename": "Chi_Sales_Split.xlsx",
            "key": "Test_Chile",
            "date_type": "Split",
            "cols": {
                "Year_Col": "Anio", "Month_Col": "Mes", 
                "Country": "Zona", "Region": "Canal", "Business_Unit": "Dept", 
                "Product_Category": "Linea", "SKU_Name": "Nombre", "Transaction_Type": "Flujo",
                "Currency": "ISO", "Quantity": "Unidades", "Value_USD": "USD_Total"
            }
        },
        {
            "filename": "Den_Sales_Direct.xlsx",
            "key": "Test_Denmark",
            "date_type": "Direct",
            "cols": {
                "Date_Col": "Doc_Date", 
                "Country": "Country_Code", "Region": "Area", "Business_Unit": "Org", 
                "Product_Category": "Main_Category", "SKU_Name": "Item_Name", "Transaction_Type": "Category",
                "Currency": "Curr", "Quantity": "Qty", "Value_USD": "Amount_USD"
            }
        },
        {
            "filename": "Egy_Sales_Split.xlsx",
            "key": "Test_Egypt",
            "date_type": "Split",
            "cols": {
                "Year_Col": "FY", "Month_Col": "Period", 
                "Country": "Territory", "Region": "Segment", "Business_Unit": "Vertical", 
                "Product_Category": "Class", "SKU_Name": "Product", "Transaction_Type": "Trade",
                "Currency": "Local_Curr", "Quantity": "Count", "Value_USD": "Total_USD"
            }
        },
        {
            "filename": "Fra_Sales_Direct.xlsx",
            "key": "Test_France",
            "date_type": "Direct",
            "cols": {
                "Date_Col": "Facture_Date", 
                "Country": "Pays_Source", "Region": "District", "Business_Unit": "Branche", 
                "Product_Category": "Rayon", "SKU_Name": "Reference", "Transaction_Type": "Sens",
                "Currency": "Code_Monnaie", "Quantity": "Qte", "Value_USD": "Montant_Global"
            }
        }
    ]

    for config in configs:
        data = []
        for i in range(5):
            row = {}
            for logical_name, header in config['cols'].items():
                if logical_name == 'Year_Col': row[header] = 2024
                elif logical_name == 'Month_Col': row[header] = "Mar"
                elif logical_name == 'Date_Col': row[header] = "2024-03-01"
                elif logical_name == 'Value_USD': row[header] = 1000 + (i * 100)
                else: row[header] = f"{logical_name}_{i}"
            
            # Add noise columns to reach 15+
            for n in range(1, 8):
                row[f"Analytics_Tag_{n}"] = f"Tag_{n}_{i}"
            
            data.append(row)
        
        df = pd.DataFrame(data)
        file_path = os.path.join(DATA_DIR, config['filename'])
        df.to_excel(file_path, index=False)
        print(f"Generated: {config['filename']}")

    return configs

def update_mapper(configs):
    app = xw.App(visible=False)
    try:
        wb = app.books.open(MAPPER_PATH)
        for sheet_name in ['Export', 'Import']:
            sheet = wb.sheets[sheet_name]
            # Clear existing mappings (keeping headers)
            last_existing = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
            if last_existing > 1:
                sheet.range(f'A2:K{last_existing}').clear_contents()
            
            # Write new headers first to ensure consistency (matching the new 10-column dashboard)
            headers = ["Key", "Source_Sheet", "Target_Sheet", "Date_Type", "Date_Col", "Year_Col", "Month_Col", 
                       "Date", "Country", "Region", "Business_Unit", "Product_Category", "SKU_Name", "Transaction_Type", "Currency", "Quantity", "Value_USD"]
            sheet.range("A1").value = headers
            
            curr_row = 2
            for config in configs:
                # Key(A), Source(B), Target(C), DateType(D), DateCol(E), Year(F), Month(G), 
                # Meta mappings start from index 7 (Column H): Date, Country, etc.
                row_base = [
                    config['key'],
                    "Sheet1",
                    "Data",
                    config['date_type'],
                    config['cols'].get('Date_Col', ""),
                    config['cols'].get('Year_Col', ""),
                    config['cols'].get('Month_Col', ""),
                    config['cols'].get('Date_Col', ""), # Map Date to Date_Col if direct
                    config['cols'].get('Country'),
                    config['cols'].get('Region'),
                    config['cols'].get('Business_Unit'),
                    config['cols'].get('Product_Category'),
                    config['cols'].get('SKU_Name'),
                    config['cols'].get('Transaction_Type'),
                    config['cols'].get('Currency'),
                    config['cols'].get('Quantity'),
                    config['cols'].get('Value_USD')
                ]
                # Adjusting for Split cases where "Date" logically doesn't exist as a single column in base
                if config['date_type'] == 'Split':
                    # Maybe map Date to Year_Col or leave blank for now? 
                    # If the dashboard has a single 'Date' column, the current engine might struggle if it's split.
                    # Actually, the engine just copies what's mapped.
                    pass

                sheet.range(f"A{curr_row}").value = row_base
                curr_row += 1
            
            sheet.autofit()
        
        wb.save()
        wb.close()
        print("Updated Mapper with realistic mappings.")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        app.quit()

if __name__ == "__main__":
    configs = generate_files()
    update_mapper(configs)
