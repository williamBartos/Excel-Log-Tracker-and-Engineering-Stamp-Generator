import openpyxl
import pythoncom
pythoncom.CoInitialize()
import win32com
import win32com.client


totalSDs = int(input('How many shop drawings are being reviewed? '))
print(totalSDs)
numList = []


for i in range(totalSDs):
        drawingNums = int(input('Enter the Shop Drawing Number: '))
        if drawingNums == '':
            break
        else: 
            numList.append(drawingNums)

        


def stampWriter(numList):

     
    log = openpyxl.load_workbook(r'C:\Users\wbartos\Documents\Pyscripts\Stamp\ShopDrawingLog.xlsx')
    logSheet = log.get_sheet_by_name('Log')
   
    for i in range(11, logSheet.get_highest_row()+1):
        
        if logSheet['A' + str(i)].value in numList:
            
            if logSheet['F' + str(i)].value == 'NET':
                
                wb = openpyxl.load_workbook(r'C:\Users\wbartos\Documents\Pyscripts\Stamp\Stamp2.xlsx')
                sheet = wb.get_sheet_by_name('Sheet1')
                sheet['B1'].value = logSheet['C' + str(i)].value
                sheet['B2'].value = 'X'
                wb.save(r'C:\Users\wbartos\Documents\Pyscripts\Stamp\newstamp' + str(i) + '.xlsx')
                
                xlApp = win32com.client.Dispatch("Excel.Application")
                books = xlApp.Workbooks.Open(r'C:\Users\wbartos\Documents\Pyscripts\Stamp\newstamp' + str(i) + '.xlsx')
                ws = books.Worksheets[0]
                ws.Visible = 1
                ws.ExportAsFixedFormat(0, r'C:\Users\wbartos\Documents\Pyscripts\Stamp\newstamp' + str(i) + '.pdf')
                
            elif logSheet['F' + str(i)].value == 'ET':
                
                wb = openpyxl.load_workbook(r'C:\Users\wbartos\Documents\Pyscripts\Stamp\Stamp2.xlsx')
                sheet = wb.get_sheet_by_name('Sheet1')
                sheet['B1'].value = logSheet['C' + str(i)].value
                sheet['B3'].value = 'X'
                wb.save(r'C:\Users\wbartos\Documents\Pyscripts\Stamp\newstamp' + str(i) + '.xlsx')
                
                xlApp = win32com.client.Dispatch("Excel.Application")
                books = xlApp.Workbooks.Open(r'C:\Users\wbartos\Documents\Pyscripts\Stamp\newstamp' + str(i) + '.xlsx')
                ws = books.Worksheets[0]
                ws.Visible = 1
                ws.ExportAsFixedFormat(0, r'C:\Users\wbartos\Documents\Pyscripts\Stamp\newstamp' + str(i) + '.pdf')
                
            elif logSheet['F' + str(i)].value == 'E&C':
                
                wb = openpyxl.load_workbook(r'C:\Users\wbartos\Documents\Pyscripts\Stamp\Stamp2.xlsx')
                sheet = wb.get_sheet_by_name('Sheet1')
                sheet['B1'].value = logSheet['C' + str(i)].value
                sheet['B4'].value = 'X'
                wb.save(r'C:\Users\wbartos\Documents\Pyscripts\Stamp\newstamp' + str(i) + 'xlsx')
                
                xlApp = win32com.client.Dispatch("Excel.Application")
                books = xlApp.Workbooks.Open(r'C:\Users\wbartos\Documents\Pyscripts\Stamp\newstamp' + str(i) + '.xlsx')
                ws = books.Worksheets[0]
                ws.Visible = 1
                ws.ExportAsFixedFormat(0, r'C:\Users\wbartos\Documents\Pyscripts\Stamp\newstamp' + str(i) + '.pdf')
                
            elif logSheet['F' + str(i)].value == 'R&R':
                
                wb = openpyxl.load_workbook(r'C:\Users\wbartos\Documents\Pyscripts\Stamp\Stamp2.xlsx')
                sheet = wb.get_sheet_by_name('Sheet1')
                sheet['B1'].value = logSheet['C' + str(i)].value
                sheet['B6'].value = 'X'
                wb.save(r'C:\Users\wbartos\Documents\Pyscripts\Stamp\newstamp' + str(i) + '.xlsx')
                
                xlApp = win32com.client.Dispatch("Excel.Application")
                books = xlApp.Workbooks.Open(r'C:\Users\wbartos\Documents\Pyscripts\Stamp\newstamp' + str(i) + '.xlsx')
                ws = books.Worksheets[0]
                ws.Visible = 1
                ws.ExportAsFixedFormat(0, r'C:\Users\wbartos\Documents\Pyscripts\Stamp\newstamp' + str(i) + '.pdf')
                
            elif logSheet['F' + str(i)].value == 'REJ':
                
                wb = openpyxl.load_workbook(r'C:\Users\wbartos\Documents\Pyscripts\Stamp\Stamp2.xlsx')
                sheet = wb.get_sheet_by_name('Sheet1')
                sheet['B1'].value = logSheet['C' + str(i)].value
                sheet['B6'].value = 'X'
                wb.save(r'C:\Users\wbartos\Documents\Pyscripts\Stamp\newstamp' + str(i) + '.xlsx')
                
                xlApp = win32com.client.Dispatch("Excel.Application")
                books = xlApp.Workbooks.Open(r'C:\Users\wbartos\Documents\Pyscripts\Stamp\newstamp' + str(i) + '.xlsx')
                ws = books.Worksheets[0]
                ws.Visible = 1
                ws.ExportAsFixedFormat(0, r'C:\Users\wbartos\Documents\Pyscripts\Stamp\newstamp' + str(i) + '.pdf')
            
            

stampWriter(numList)
    