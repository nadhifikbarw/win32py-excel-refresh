from os.path import abspath
import win32com.client as win32
# So Far library ini hanya ada di Windows, kemungkinan tidak tersedia untuk Mac/Linux


def open_excel():
  excel = win32.gencache.EnsureDispatch('Excel.Application')
  excel.Visible = True
  return excel

def open_workbook(excel, path):
  return excel.Workbooks.Open(Filename=path)


# References: https://stackoverflow.com/questions/40893870/refresh-excel-external-data-with-python
# VBA References: https://docs.microsoft.com/en-us/office/vba/api/overview/excel/object-model

def main():
  # Buka Excel
  path = abspath('dashboard.xlsx')
  excel = open_excel()
  wb = open_workbook(excel, path)

  # Melakukan query ulang dengan menjalankan Refresh All pada seluruh 
  wb.RefreshAll()
  # Menunggu hingga semua Async Queries selesai. Refer to: https://docs.microsoft.com/en-us/office/vba/api/excel.application.calculateuntilasyncqueriesdone 
  excel.CalculateUntilAsyncQueriesDone() 
  # Close dan Save Workbook
  wb.Close(SaveChanges=True)
  # Tutup Excel
  excel.Quit()
    
if __name__ == "__main__":
  main()

  # Unknown Behaviour
  # Berguna untuk mengkalkulasi ulang PivotTable. Refer To: https://docs.microsoft.com/en-us/office/vba/api/excel.application.calculatefull
  #excel.CalculateFull()