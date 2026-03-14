import pandas as pd
import os

file_path = r"c:\Users\USER\Downloads\english learning\2026 new1  (1).xlsx"
if os.path.exists(file_path):
    df_raw = pd.read_excel(file_path, header=None)
    row2 = df_raw.iloc[2].to_list()
    print("RAW INDEX 2 FULL DATA:")
    for idx, val in enumerate(row2):
        print(f"  Column {idx}: {val}")
else:
    print(f"File not found: {file_path}")
