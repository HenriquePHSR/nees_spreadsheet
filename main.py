import sys
import sheet_process as sp
#import win32com.client as win32

# fname = "Alunos_concluintes.xls"
# excel = win32.gencache.EnsureDispatch('Excel.Application')
# wb = excel.Workbooks.Open(fname)

# wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
# wb.Close()                               #FileFormat = 56 is for .xls extension
# excel.Application.Quit()


# python main.py input.xlsx
print("init: args file "+str(sys.argv[1:]))
sp.build_df(sys.argv[1:])

#sp.dispatch_email_from_df(sys.argv[1:])