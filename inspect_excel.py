import pandas as pd
import os

file_path = r"c:\Users\USER\Downloads\english learning\2026 new1  (1).xlsx"
if os.path.exists(file_path):
    df = pd.read_excel(file_path)
    print("Columns:", df.columns.tolist())
    for i, row in df.head(10).iterrows():
        print(f"Row {i}: {row.to_dict()}")
else:
    print(f"File not found: {file_path}")
