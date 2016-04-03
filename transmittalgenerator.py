import openpyxl
from openpyxl.styles.borders import Border, Side
import pythoncom
pythoncom.CoInitialize()
import win32com
import win32com.client
import PyPDF2
import os


shopDict = {'Project':'171',
            'Log':(r'ShopDrawingLog.xlsx'),
            'Stamp':(r'Stamp2.xlsx'),
            'Out': (os.path.abspath('OUT')),
            'In': (os.path.abspath('IN')),
            'Transmittal': (r'transmittal')
            }
            
            
stampDict = {'NET':('B13'),
            'ET':('B14'),
            'E&C':('B15'),
            'R&R':('B17'),
            'REJ':('B18')
            }

totalSDs = int((input('How many shop drawings are being reviewed? ')))

numList = []


for i in range((totalSDs)):
     drawingNums = (input('Enter the Shop Drawing Number: '))
     numList.append(drawingNums)

                
                
                
def transmittalWriter(numList):
    
    thinBorder = Border(left=Side(style='thin'), 
                        right=Side(style='thin'), 
                        top=Side(style='thin'), 
                        bottom=Side(style='thin')
                     )
    
    wbpath = (shopDict['Transmittal'])
    wbpathext = wbpath + '.xlsx'
     
    log = openpyxl.load_workbook(shopDict['Log'])
    logSheet = log.get_sheet_by_name('Log')
    
    wb = openpyxl.load_workbook(wbpathext)
    sheet = wb.get_sheet_by_name('Sheet1')
  
    currentRow = 18
      
    for logrow in range(11, logSheet.get_highest_row()+1):
        if str(logSheet['A' + str(logrow)].value) in numList:
            pdfFileObj = open(logSheet['K'+str(logrow)].value, 'rb')
            pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
            sheet['A8'].value = logSheet['A3'].value
            sheet['A9'].value = logSheet['A4'].value
            sheet['A10'].value = logSheet['A5'].value
            sheet['A'+str(currentRow)].value = str(logSheet['A' + str(logrow)].value)
            sheet['A'+str(currentRow)].border = thinBorder
            sheet['B'+str(currentRow)].value = pdfReader.numPages
            sheet['B'+str(currentRow)].border = thinBorder
            sheet['C'+str(currentRow)].value = str(logSheet['C' + str(logrow)].value)
            sheet['C'+str(currentRow)].border = thinBorder
            sheet['D'+str(currentRow)].value = str(logSheet['F' + str(logrow)].value)
            sheet['D'+str(currentRow)].border = thinBorder
            
            currentRow+= 1
         
    newTransmittal = ((shopDict['Out'] + '\\' + shopDict['Transmittal'] + str(logrow)))
    wb.save(newTransmittal + '.xlsx')
    
    xlsxToPdf(newTransmittal)
    addHeader(newTransmittal + '.pdf') 
                

def xlsxToPdf(path):
    xlApp = win32com.client.Dispatch("Excel.Application")
    books = xlApp.Workbooks.Open(path + '.xlsx')
    ws = books.Worksheets[0]
    ws.Visible = 1
    ws.ExportAsFixedFormat(0, path + '.pdf')
    stampPdf = path + '.pdf'
    books.Close(True)
    return stampPdf
     

def pdfMerger(stamp,submittal, path):
    
    stampFile = open(stamp, 'rb')
    submittalFile = open(submittal,'rb')  
    merger = PyPDF2.PdfFileMerger()
    
    try:
        merger.merge(position = 0, fileobj = stampFile) 
        merger.merge(position = 2, fileobj = submittalFile)
        merger.write(open(path, 'wb'))

    finally:
        stampFile.close()
        submittalFile.close()
        
def addHeader(path):
    
    pdfNoHeader = open(path, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfNoHeader)
    firstPage = pdfReader.getPage(0)
    pdfHeaderReader = PyPDF2.PdfFileReader(open('watermark3.pdf', 'rb'))
    firstPage.mergePage(pdfHeaderReader.getPage(0))
    pdfWriter = PyPDF2.PdfFileWriter()
    pdfWriter.addPage(firstPage)
    
    for pageNum in range(1, pdfReader.numPages):
           pageObj = pdfReader.getPage(pageNum)
           pdfWriter.addPage(pageObj)
           

    resultPdfFile = open('finishedTransmittal.pdf', 'wb')
    pdfWriter.write(resultPdfFile)
    pdfNoHeader.close()
        
 
       
input('PRESS ENTER TO EXIT')   

 
transmittalWriter(numList)







    