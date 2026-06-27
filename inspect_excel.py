import pandas as pd

EXCEL_PATH = "Graduating_college.xlsx"

def inspect_excel():
    print("=== EXCEL RAW INSPECTION REPORT ===")
    try:
        # header=None forces pandas to read every single row as raw data
        df = pd.read_excel(EXCEL_PATH, header=None)
        
        print("Displaying the first 15 rows of your Excel file...\n")
        print("-" * 80)
        
        for index, row in df.head(15).iterrows():
            # Convert row to a list, replacing NaN (empty cells) with empty strings for readability
            clean_row = [str(item) if pd.notna(item) else "" for item in row.tolist()]
            print(f"Row {index}: {clean_row}")
            
        print("-" * 80)
        print("\nLook for the Row number that contains 'Student Name' and 'Program'.")
        
    except Exception as e:
        print(f"Error reading Excel file: {e}")

if __name__ == "__main__":
    inspect_excel()