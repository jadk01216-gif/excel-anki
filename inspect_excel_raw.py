import pandas as pd
import os

file_path = r"c:\Users\USER\Downloads\english learning\2026 new1  (1).xlsx"
if os.path.exists(file_path):
    df_raw = pd.read_excel(file_path, header=None)
    print("INDEX 0:", df_raw.iloc[0].to_list())
    print("INDEX 1:", df_raw.iloc[1].to_list())
    print("INDEX 2:", df_raw.iloc[2].to_list())
else:
    print(f"File not found: {file_path}")
