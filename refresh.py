from os.path import abspath
import win32com.client as win32
# Only available for Windows

def open_excel(headless=True):
  excel = win32.gencache.EnsureDispatch('Excel.Application')
  excel.Visible = !headless # Default to run in headless-mode
  return excel

def open_workbook(excel, path):
  return excel.Workbooks.Open(Filename=path)

# References: https://stackoverflow.com/questions/40893870/refresh-excel-external-data-with-python
# VBA References: https://docs.microsoft.com/en-us/office/vba/api/overview/excel/object-model

def main():
  # Open Excel.Application, make sure you close all running Excel instances before running this script 
  excel = open_excel()
  
  # Open the dashboard Excel file
  path = abspath('dashboard.xlsx')
  wb = open_workbook(excel, path)

  # Trigger 'Refresh All'
  wb.RefreshAll()
  
  # Wait until all async queries done
  # Refer to: https://docs.microsoft.com/en-us/office/vba/api/excel.application.calculateuntilasyncqueriesdone 
  excel.CalculateUntilAsyncQueriesDone() 
  
  # Close dan Save Workbook
  wb.Close(SaveChanges=True)
  
  # Close Excel.Application
  excel.Quit()
    
if __name__ == "__main__":
  main()

  # Need Research
  # Refer To: https://docs.microsoft.com/en-us/office/vba/api/excel.application.calculatefull
  #excel.CalculateFull()
