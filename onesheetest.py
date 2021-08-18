import openpyxl
import pathlib
import win32com.client as win32
import os

import logging
logging.basicConfig(filename='pipelineanalysis.log', level=logging.INFO)

#Creates variable for the necessary file destinations.
tsp5_updated = 'C:\\Users\\Administrator\\Downloads\\tsp5_updated.xlsx'
pl_template_rows = 'C:\\Users\\Administrator\\Downloads\\PL_rows.xlsx'
test_copy_merge = "C:\\Users\\Administrator\\Downloads\\copy_merge.xlsx"
PL_file = 'C:\\Users\\Administrator\\Downloads\\PL_20210521_Workflow.xlsx'
updated_file = 'C:\\Users\\Administrator\\Downloads\\PL_Updated_Workflow.xlsx'
tsp5_updated = 'C:\\Users\\Administrator\\Downloads\\tsp5_updated.xlsx'
test_save = 'C:\\Users\\Administrator\\Downloads\\test_save.xlsx'
excel = win32.gencache.EnsureDispatch('Excel.Application')

row_file = openpyxl.load_workbook(tsp5_updated)
row_sheet = row_file.worksheets[0]
row_file.close()

maxRow = row_sheet.max_row + 15
strRow = str(maxRow)


print("Test1")
print("Test2")
copy_workbook = excel.Workbooks.Open(test_copy_merge)
print('Test')

# Merges the updated tsp5_demand file and the autofilled template file into the master Excel file as tabs.
excel.Visible = False
PL_workbook = excel.Workbooks.Open(PL_file)
tsp5_workbook = excel.Workbooks.Open(tsp5_updated)
copy_workbook.Worksheets(1).Move(Before=PL_workbook.Worksheets(1))
tsp5_workbook.Worksheets(1).Move(Before=PL_workbook.Worksheets(1))
excel.DisplayAlerts = False
PL_workbook.SaveAs(updated_file)
excel.DisplayAlerts = True
excel.Visible = True
excel.Application.Quit()
print('Files have been successfully merged.')
logging.info('Workbooks Merged into Final Document')

#Removes copies of intermediate files completed during the workflow.
#os.remove(tsp5_updated)
#os.remove(pl_template_rows)
#os.remove(test_copy_merge)

#Removed code, saved in case it is needed later
    # Iterate again, printing new values
    #for cell_tuple in cells_with_formulae:
    #    for cell in cell_tuple:
    #        print(cell.value)

    # Iterate again, printing new values
    #for cell_tuple in cells_with_formulae2:
    #    for cell in cell_tuple:
    #        print(cell.value)

    # Iterate again, printing new values
    #for cell_tuple in cells_with_formulae3:
    #    for cell in cell_tuple:
