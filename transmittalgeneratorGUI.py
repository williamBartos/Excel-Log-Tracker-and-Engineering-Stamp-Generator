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
            'Transmittal': ('.\\Templates\\transmittal2'),
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
        
    currentRow = 28 #STARTING ROW FOR SUBMITTAL TABLE


    def copyValues(currentRow):
        
        sheet['G13'].value = logSheet['A3'].value #Project Name
        sheet['G14'].value = logSheet['A4'].value #Client Name
        sheet['J11'].value = logSheet['A5'].value #MC Project No.
        sheet['A'+str(currentRow)].value = str(logSheet['A' + str(logrow)].value) #SD NO. 
        sheet['C'+str(currentRow)].value = str(logSheet['C' + str(logrow)].value) #DESCRIPTION
        sheet['K'+str(currentRow)].value = str(logSheet['F' + str(logrow)].value.upper()) #STATUS
  
      
    for logrow in range(11, logSheet.get_highest_row()+1):
        if str(logSheet['A' + str(logrow)].value) in numList:
            try:
                pdfFileObj = open(logSheet['K'+str(logrow)].value, 'rb')
                pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
                copyValues(currentRow)
                sheet['B'+str(currentRow)].value = pdfReader.numPages #TOTAL PAGES 
                currentRow+= 1
               
         
            except:
                copyValues(currentRow)
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
           
    try:
        n=0
        
        if os.path.isfile((shopDict['Out'] + '\\Transmittal_' + now.strftime("%Y-%m-%d") +'-' + '%s' + '.pdf') % n) == False:  
            resultPdfFile = open((shopDict['Out'] + '\\Transmittal_' + now.strftime("%Y-%m-%d") +'-' + '%s' + '.pdf') % n, 'wb' ) #FINISHED PDF
        
        else:
            while(os.path.isfile((shopDict['Out'] + '\\Transmittal_' + now.strftime("%Y-%m-%d") +'-' + '%s' + '.pdf') % n)):
                n+=1
            resultPdfFile = open((shopDict['Out'] + '\\Transmittal_' + now.strftime("%Y-%m-%d") +'-' + '%s' + '.pdf') % n, 'wb' ) #FINISHED PDF 
            
        pdfWriter.write(resultPdfFile)
        pdfNoHeader.close()
        pdfHeader.close()
            
    except:
        pdfNoHeader.close()
        pdfHeader.close()
        pass





    