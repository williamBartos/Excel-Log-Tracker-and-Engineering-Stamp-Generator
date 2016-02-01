import openpyxl

wbList = [
          r'R:\General\Projects\Msa\MSA-132\Project_Info\Shop Drawings\ShopDrawingLog.xlsx',
          r'R:\General\Projects\Msa\MSA-153\Project_Info\Shop Drawings\ShopDrawingLog.xlsx',
          r'R:\General\Projects\Msa\MSA-156\Construction\Shop Drawings\ShopDrawings(MSA156).xlsx',
          r'R:\General\Projects\Msa\MSA-157\Construction\Shop Drawings\ShopDrawingLog.xlsx',
          r'R:\General\Projects\Msa\MSA-162\Project_Info\Shop Drawings\ShopDrawingLog.xlsx',
          r'R:\General\Projects\Msa\MSA-167\Project Information\Shop Drawings\ShopDrawingLog.xlsx',
          r'R:\General\Projects\Msa\MSA-168\Project Information\Shop Drawings\ShopDrawingLog.xlsx',
          r'R:\General\Projects\Msa\MSA-170\Project Information\Shop Drawings\ShopDrawingLog.xlsx',
          r'R:\General\Projects\Msa\MSA-171\Project Information\Shop Drawings\ShopDrawingLog.xlsx',
          r'R:\General\Projects\Msa\MSA-177\Project Information\Shop Drawings\ShopDrawingLog.xlsx',
          r'R:\General\Projects\Msa\MSA-178\Project Information\Shop Drawings\ShopDrawingLog.xlsx',
          r'R:\General\Projects\Msa\MSA-185\Project Information\Shop Drawings\ShopDrawingsLog.xlsx',
          r'R:\General\Projects\Msa\MSA-188\Project Information\Shop Drawings\ShopDrawingsLog.xlsx',
          r'R:\General\Projects\Msa\MSA-195\Project Information\Shop Drawings\ShopDrawingLog.xlsx',
          ]
listLen = len(wbList)
reportFile = open(r'C:\Users\wbartos\Documents\Pyscripts\logreport7.txt', 'w')


for i in range(listLen):
    wbCurr = openpyxl.load_workbook(wbList[i])
    sheet = wbCurr.get_sheet_by_name('Log')
    reportFile.write('\n' + '********* ' + wbList[i] + ' *********' + '\n')
    for row in range(11, sheet.get_highest_row()+1):
        if sheet['E' + str(row)].value is None and sheet['D' + str(row)].value is not None:
            reportFile.write('-' + str(sheet['C' + str(row)].value) +' submitted on ' + str(sheet['D' + str(row)].value) + ' has not been resubmitted ' + '\n')
        
reportFile.close()
            
    
        

