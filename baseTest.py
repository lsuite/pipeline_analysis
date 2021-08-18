import pandas as pd
from pandas import ExcelWriter

import openpyxl
from openpyxl import workbook

from datetime import date

#Record date of new data export from ServiceNow
today = date.today()
newDate = today.strftime("%Y%m%d")
print("Today's date:", newDate)



base_location = 'C:\\Users\\Administrator\\pipelineAnalysis\\PL_20210409_Workflow.xlsx'
base_file = pd.read_excel(base_location)
copy_location = 'C:\\Users\\Administrator\\pipelineAnalysis\\PL_20210409.xlsx'

#writer = pd.ExcelWriter(copy_location, engine = 'xlsxwriter')
#base_file.to_excel(writer)

with ExcelWriter(copy_location) as writer:
    base_file.to_excel(writer)

print(base_file)

