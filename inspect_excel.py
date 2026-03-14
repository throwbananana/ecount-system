
import pandas as pd
import os

file_path = r"12yue\1122.xlsx"

if not os.path.exists(file_path):
    print(f"File not found: {file_path}")
else:
    try:
        df = pd.read_excel(file_path)
        print("Columns:", df.columns.tolist())
        print("-" * 30)
        print(df.head(5).to_string())
    except Exception as e:
        print(f"Error reading Excel: {e}")
