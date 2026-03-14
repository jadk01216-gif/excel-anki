import pandas as pd
import os

file_path = r"c:\Users\USER\Downloads\english learning\2026 new1  (1).xlsx"
if os.path.exists(file_path):
    # Try reading with header=1 (the status, Translation row)
    df = pd.read_excel(file_path, header=1)
    print("Detected Columns:", df.columns.tolist())
    print("\nFirst 3 rows of data:")
    print(df.head(3).to_string())
    
    # Also check raw positions for row 2
    df_raw = pd.read_excel(file_path, header=None)
    row2 = df_raw.iloc[2].to_list()
    print("\nRaw Row 2 (Internal Index 2):", row2)
    for idx, val in enumerate(row2):
        print(f"  Col {idx}: {val}")
else:
    print(f"File not found: {file_path}")
