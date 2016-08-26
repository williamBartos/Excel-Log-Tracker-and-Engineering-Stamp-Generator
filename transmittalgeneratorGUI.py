#WRITTEN BY WILLIAM BARTOS IN PYTHON 3.5#



#THESE PACKAGES ARE NECESSARY FOR PYINSTALLER TO CONVERT TO .EXE
import six 
import packaging
import packaging.version
import packaging.specifiers
import packaging.requirements


import re
import openpyxl
from openpyxl.styles.borders import Border, Side
import pythoncom
pythoncom.CoInitialize()
import win32com
import win32com.client
import PyPDF2
import os
import datetime

now = datetime.datetime.now()

shopDict = {
            'Project':'171',
            'Log':(r'ShopDrawingLog.xlsx'),
            'Stamp':(r'./Templates/Stamp.xlsx'),
            'Out': (os.path.abspath('OUT')),
            'In': (os.path.abspath('IN')),
            'Transmittal': ('.\\Templates\\transmittal1'),
            'Header':(r'./Templates/header.pdf')
            }
            
            
stampDict = {
            'NET':('B12'),
            'ET':('B13'),
            'E&C':('B14'),
            'R&R':('B16'),
            'REJ':('B17')
            }

            
numList=[]

thinBorder = Border(left=Side(style='thin'), 
                    right=Side(style='thin'), 
                    top=Side(style='thin'), 
                    bottom=Side(style='thin')
                 )
                 
def transmittalWriter(numList):
    
    
    wbpath = shopDict['Transmittal'] + '.xlsx'
     
    log = openpyxl.load_workbook(shopDict['Log'])
    logSheet = log.get_sheet_by_name('Log')    
    wb = openpyxl.load_workbook(wbpath)
    sheet = wb.get_sheet_by_name('Sheet1')
        
    currentRow = 29 #STARTING ROW FOR SUBMITTAL TABLE
    headerRow = currentRow - 1

    def copyValues(currentRow):
        
        sheet['E13'].value = logSheet['A3'].value #Project Name
        sheet['E14'].value = logSheet['A4'].value #Client Name
        sheet['G11'].value = logSheet['A5'].value #MC Project No.
        sheet['A'+str(currentRow)].value = str(logSheet['A' + str(logrow)].value) #SD NO. 
        sheet['C'+str(currentRow)].value = str(logSheet['C' + str(logrow)].value) #DESCRIPTION
        sheet['H'+str(currentRow)].value = str(logSheet['F' + str(logrow)].value) #STATUS
        
        sheet['A'+str(currentRow)].border = thinBorder
        sheet['B'+str(currentRow)].border = thinBorder
        sheet['C'+str(currentRow)].border = thinBorder
        sheet['D'+str(currentRow)].border = thinBorder
        sheet['E'+str(currentRow)].border = thinBorder
        sheet['F'+str(currentRow)].border = thinBorder
        sheet['G'+str(currentRow)].border = thinBorder
        sheet['H'+str(currentRow)].border = thinBorder
    
        sheet['A'+str(headerRow)].border = thinBorder
        sheet['B'+str(headerRow)].border = thinBorder
        sheet['C'+str(headerRow)].border = thinBorder
        sheet['D'+str(headerRow)].border = thinBorder
        sheet['E'+str(headerRow)].border = thinBorder
        sheet['F'+str(headerRow)].border = thinBorder
        sheet['G'+str(headerRow)].border = thinBorder
        sheet['H'+str(headerRow)].border = thinBorder

    
      
    for logrow in range(11, logSheet.get_highest_row()+1):
        if str(logSheet['A' + str(logrow)].value) in numList:
            try:
                pdfFileObj = open(logSheet['K'+str(logrow)].value, 'rb')
                pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
                copyValues(currentRow)
                
            except:
                copyValues(currentRow)
                currentRow+= 1
                 
                try:
                    sheet['B'+str(currentRow)].value = pdfReader.numPages #TOTAL PAGES 
                except:
                    continue
                       
        currentRow+= 1
                
    try:
        newTransmittal = ((shopDict['Out'] + '\\Transmittal'))
        wb.save(newTransmittal)
    
        xlsxToPdf(newTransmittal)
        addHeader(newTransmittal + '.pdf') 
        
        os.remove(newTransmittal)
        os.remove(newTransmittal + '.pdf')
        
    except:
        os.remove(newTransmittal)
        os.remove(newTransmittal + '.pdf')
                



def xlsxToPdf(path):
    xlApp = win32com.client.Dispatch("Excel.Application")
    books = xlApp.Workbooks.Open(path)
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
    
    pdfNoHeader = open(path, 'rb') #opens the generated PDF(from xlsxToPdf)
    pdfReader = PyPDF2.PdfFileReader(pdfNoHeader) #creates an object out of the NoHeader PDF
    firstPage = pdfReader.getPage(0) #grabs the first page 
    pdfHeader= open(shopDict['Header'], 'rb')
    pdfHeaderReader = PyPDF2.PdfFileReader(pdfHeader) #opens the header template
    firstPage.mergePage(pdfHeaderReader.getPage(0)) #merges the first page of the NoHeader PDF with the header template
    pdfWriter = PyPDF2.PdfFileWriter() #creates a new PDF
    pdfWriter.addPage(firstPage) #adds the watermarked page to the first page of the new PDF
    
    for pageNum in range(1, pdfReader.numPages):
           pageObj = pdfReader.getPage(pageNum) #gets each page from the no header PDF
           pdfWriter.addPage(pageObj) #adds each page of the no header PDF to the new PDF
           

    resultPdfFile = open(shopDict['Out'] + '\\Transmittal_' + now.strftime("%Y-%m-%d") + '.pdf', 'wb') #FINISHED PDF
    pdfWriter.write(resultPdfFile)
    pdfNoHeader.close()
    pdfHeader.close()
    





    