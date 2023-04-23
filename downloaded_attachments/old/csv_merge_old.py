import pandas as pd
import glob
from openpyxl import load_workbook

def get_formid_index(worksheet):
    for i, cell in enumerate(worksheet[1]):
        if cell.value == "formid":
            return i
    return None

def align_formid(df, formid_index):
    if formid_index is not None:
        df['formid'] = pd.to_numeric(df['formid'], errors='coerce')
        df = df.sort_values(by='formid').reset_index(drop=True)

    return df

def merge_excel_files(files):
    dfs = []

    for file in files:
        wb = load_workbook(file)
        ws = wb.active

        formid_index = get_formid_index(ws)
        df = pd.DataFrame(ws.values)
        df.columns = df.iloc[0]
        df = df[1:]

        df = align_formid(df, formid_index)
        dfs.append(df)

    merged_df = dfs[0]

    for df in dfs[1:]:
        merged_df = pd.merge(merged_df, df, on='formid', how='outer')

    merged_df.to_excel('janssons_kranar_merged.xlsx', index=False, engine='openpyxl')

if __name__ == "__main__":
    xlsx_files = glob.glob("janssons_kranar_*.xlsx")

    merge_excel_files(xlsx_files)
