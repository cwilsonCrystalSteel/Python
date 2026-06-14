# -*- coding: utf-8 -*-
"""
Created on Sun Jun 14 12:18:27 2026

@author: Netadmin
"""

import win32com.client as win32
from pathlib import Path
import shutil
import os



def run_macro(template_file, output_file, bs_file, tax_report_file):
    print('Generating copy of {template_file.name} to {output_file.name}')
    shutil.copy2(template_file, output_file)

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    
    # macro can only accept str-path, not Path
    if isinstance(bs_file, Path):
        bs_file = str(bs_file)
        
    if isinstance(tax_report_file, Path):
        bs_file = str(tax_report_file)

    try:
        wb = excel.Workbooks.Open(output_file)
        
        print(f'Running ImportCRYSSWSpectrumEmployerTaxExpenseJournalEntry on: {bs_file}')
        excel.Application.Run("ImportCRYSSWSpectrumEmployerTaxExpenseJournalEntry", bs_file, True)
        
        print(f'Running ImportLaborDistributionPayrollSummary on: {tax_report_file}')
        excel.Application.Run("ImportLaborDistributionPayrollSummary", tax_report_file, True)
        
        print('Running ProcessData...')
        excel.Application.Run("ProcessData")
        
        wb.Save()
        wb.Close()
    finally:
        excel.Quit()





# template_file = r"C:\Users\Netadmin\Downloads\test-g-drive\2026\2026.06.12\Macro File v2 - TEMPLATE.xlsm"
# print(os.path.exists(template_file))
# working_file = r"C:\Users\Netadmin\Downloads\test-g-drive\2026\2026.06.12\Macro File v2 - STATE xyz DATE yyyy.mm.dd.xlsm"
# print(os.path.exists(working_file))
# shutil.copy2(template_file, output_file)


# try:
#     wb = excel.Workbooks.Open(working_file)
    
#     f_labor = r"C:\Users\Netadmin\Downloads\test-g-drive\2026\2026.06.12\Data Input Files\DE 06.12.2026 BS 20260614_130056.xlsx.xlsx"
#     print(os.path.exists(f_labor))
#     f_cryss = r"C:\Users\Netadmin\Downloads\test-g-drive\2026\2026.06.12\Data Input Files\DE 06.12.2026 Tax Report 20260614_130110.xls.xls"
#     print(os.path.exists(f_cryss))

#     excel.Application.Run("ImportCRYSSWSpectrumEmployerTaxExpenseJournalEntry", f_cryss, True)
#     excel.Application.Run("ImportLaborDistributionPayrollSummary", f_labor, True)
#     excel.Application.Run("ProcessData")
    
#     wb.Save()
#     # wb.SaveAs(r"C:\Users\Netadmin\Downloads\test-g-drive\2026\2026.06.12\Macro File v2 - test.xlsm")
#     wb.Close(SaveChanges=False)
# finally:
#     excel.Quit()
