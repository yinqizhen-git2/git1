# import pandas as pd 
filename=r'C:\Users\hongji\Desktop\123.xlsx'
# df=pd.read_excel(filename)
# print(df)
import win32com.client as com 
# ie=com.Dispatch('InternetExplorer.Application')
# ie.Navigate('https://e.cebbank.com/cebent/prelogin.do?_locale=zh_CN')
# ie.visible=True
xlsx=com.Dispatch('Excel.Application')
xlsx.Visible=False
xlsx.DisplayAlerts=False
wb=xlsx.Workbooks.Open(filename)
sht=wb.Worksheets('Sheet1')
a1=sht.Cells(1,1)
print(a1)
wb.SaveAs(r'C:\Users\hongji\Desktop\123.xls')
wb.Close()
#123