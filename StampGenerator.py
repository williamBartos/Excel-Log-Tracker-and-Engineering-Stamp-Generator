import openpyxl
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

                 
def stampWriter(numList):
    
    log = openpyxl.load_workbook(shopDict['Log'])
    logSheet = log.get_sheet_by_name('Log')
    
    for i in range(11, logSheet.get_highest_row()+1):               
        if str(logSheet['A' + str(i)].value) in numList:
            if logSheet['F' + str(i)].value in stampDict:
    
                wb = openpyxl.load_workbook(shopDict['Stamp'])
                sheet = wb.get_sheet_by_name('Sheet1')
                sheet['B10'].value = logSheet['C' + str(i)].value
                sheet[stampDict[logSheet['F' + str(i)].value]].value = '✓'
                sheetPath = (shopDict['Out'] + '\\newstamp' + str(i))
                transmittalPath = (shopDict['Out'] +'\\testout' + str(i) + '.pdf')
                submittalPath = logSheet['K'+str(i)].value
                print('Generated')
        
                wb.save(sheetPath + '.xlsx')
     
                pdfMerger(xlsxToPdf(sheetPath),submittalPath,transmittalPath)
                addHeader(transmittalPath, logSheet['A' + str(i)].value, logSheet['C' + str(i)].value )
                os.remove(transmittalPath)
                os.remove(sheetPath + '.pdf')
                os.remove(sheetPath + '.xlsx')
                

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
        
def addHeader(path, sdNo, sdTitle):
    
    pdfNoHeader = open(path, 'rb') #opens the generated PDF(from xlsxToPdf)
    pdfReader = PyPDF2.PdfFileReader(pdfNoHeader) #creates an object out of the NoHeader PDF
    firstPage = pdfReader.getPage(0) #grabs the first page 
    pdfHeader= open('header.pdf', 'rb')
    pdfHeaderReader = PyPDF2.PdfFileReader(pdfHeader) #opens the header template
    firstPage.mergePage(pdfHeaderReader.getPage(0)) #merges the first page of the NoHeader PDF with the header template
    pdfWriter = PyPDF2.PdfFileWriter() #creates a new PDF
    pdfWriter.addPage(firstPage) #adds the watermarked page to the first page of the new PDF
    
    for pageNum in range(1, pdfReader.numPages):
           pageObj = pdfReader.getPage(pageNum) #gets each page from the no header PDF
           pdfWriter.addPage(pageObj) #adds each page of the no header PDF to the new PDF
           

    resultPdfFile = open(shopDict['Out'] + '\\SD#' + str(sdNo) + '_' + str(sdTitle) + '.pdf', 'wb') #FINISHED PDF
    pdfWriter.write(resultPdfFile)
    print(resultPdfFile)
    pdfNoHeader.close()
    pdfHeader.close()
        
 
       
input('PRESS ENTER TO EXIT')    


            
stampWriter(numList)   








    