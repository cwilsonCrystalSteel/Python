# -*- coding: utf-8 -*-
"""
Created on Sun Jun 14 12:18:27 2026

@author: Netadmin
"""

import win32com.client as win32
import win32process
import psutil
import gc
from pathlib import Path
import shutil
import os


def run_macro(template_file, output_file, bs_filepath, tax_report_filepath):    
    
    # macro can only accept str-path, not Path
    if isinstance(bs_filepath, Path):
        bs_filepath = str(bs_filepath)
        
    if isinstance(tax_report_filepath, Path):
        tax_report_filepath = str(tax_report_filepath)
    
    
    print(f'Generating copy of {template_file.name} to {output_file.name}')
    shutil.copy2(template_file, output_file)

    wb = None
    excel = None
    excel_pid = None

    try:

        excel = win32.DispatchEx("Excel.Application")

        # Capture PID of the Excel instance we created
        hwnd = excel.Hwnd
        _, excel_pid = win32process.GetWindowThreadProcessId(hwnd)

        print(f'Created Excel PID: {excel_pid}')

        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False

        wb = excel.Workbooks.Open(str(output_file))

        print(f'Running ImportCRYSSWSpectrumEmployerTaxExpenseJournalEntry on: {bs_filepath}')
        excel.Application.Run("ImportCRYSSWSpectrumEmployerTaxExpenseJournalEntry", bs_filepath, True)

        print(f'Running ImportLaborDistributionPayrollSummary on: {tax_report_filepath}')
        excel.Application.Run("ImportLaborDistributionPayrollSummary", tax_report_filepath, True)

        print('Running ProcessData...')
        excel.Application.Run("ProcessData")

        wb.Save()
        print('Workbook saved')

    finally:

        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
                print('Workbook closed.')
            except Exception as ex:
                print(f'Error closing workbook: {ex}')

        if excel is not None:
            try:
                excel.Quit()
                print('Excel quit requested.')
            except Exception as ex:
                print(f'Error quitting Excel: {ex}')

        # Release COM references
        try:
            del wb
        except:
            pass

        try:
            del excel
        except:
            pass

        gc.collect()

        # Force kill only the Excel instance we created if it survived
        if excel_pid is not None and psutil.pid_exists(excel_pid):
            try:
                proc = psutil.Process(excel_pid)

                proc.wait(timeout=5)

            except psutil.TimeoutExpired:

                print(f'Excel PID {excel_pid} still running. Killing process.')

                try:
                    proc.kill()
                except Exception as ex:
                    print(f'Could not kill Excel PID {excel_pid}: {ex}')




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
