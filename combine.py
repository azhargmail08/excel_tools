import pandas as pd

all_sheets = pd.read_excel(r"C:\Users\StudentQR\Desktop\SENARAI NAMA MURID 2025_2026 (1).xlsx", sheet_name=None)

df_lists = []
for sheet_name, df in all_sheets.items():
    df['KELAS'] = sheet_name
    df_lists.append(df)

combined_df = pd.concat(df_lists, ignore_index=True)
combined_df.to_excel('combined_data.xlsx', index=False)