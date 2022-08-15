import win32com.client
import os

cd = os.getcwd()
excel = win32com.client.Dispatch("Excel.Application")
wb = excel.Workbooks.Open(cd + '/test.xlsm')
excel.visible = True
excel.Application.Run("Module1.test")

