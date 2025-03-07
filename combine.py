import pandas as pd

file_path = input("Enter the path to your Excel file: ").strip().strip('"').strip("'")

all_sheets = pd.read_excel(file_path, sheet_name=None)

df_lists = []
for sheet_name, df in all_sheets.items():
    df['KELAS'] = sheet_name
    df_lists.append(df)

combined_df = pd.concat(df_lists, ignore_index=True)
combined_df.to_excel('combined_data.xlsx', index=False)