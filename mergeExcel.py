import os
import pandas as pd

from datetime import date

cwd = os.path.abspath('C:\\pipelineAnalysis')
files = os.listdir(cwd)
print(files)

df_total = pd.DataFrame()
for file in files:
    if file.endswith('.xlsx'):
        fpath = os.path.join(cwd,file)
        excel_file = pd.ExcelFile(fpath)
        sheets = excel_file.sheet_names
        for sheet in sheets:
            df = excel_file.parse(sheet_name = sheet)
            df_total = df_total.append(df)
df_total.to_excel('combined_file.xlsx')