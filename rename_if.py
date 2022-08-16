import os
import pandas


reptPath = 'C:\\SurosBI\\Panorays_Reports_220810A\\'
excelFile = 'dbData_rename_220815.xlsx'
reptList = pandas.read_excel(excelFile, sheet_name=0, usecols=[1,2,3,4,5,6,7])
reptHead = ['報告檔案夾', 'Date', 'pdfName', 'csvName', 'NewDate','NewpdfName','NewcsvName']


for m in range (0, 11):
    reptFolder = reptList[reptHead[0]].values[m].strip()    #get report dir
    iniPdf = reptList[reptHead[2]].values[m].strip()    #get pdf file name
    iniCsv = reptList[reptHead[3]].values[m].strip()    #get initial csv file name
    revisPdf = reptList[reptHead[5]].values[m].strip()	 #get Newpdf file name
    revisCsv = reptList[reptHead[6]].values[m].strip()	 #get revised csv file name

    if (iniPdf == 'no.pdf') and (iniCsv == 'no.csv'):
        m+=1
        continue
    
    if iniPdf != 'no.pdf' :
        inipdfPath = reptPath + reptFolder + '\\' + iniPdf    #get pdf path
        revispdfPath = reptPath + reptFolder + '\\' + revisPdf  #get revised pdf path
        os.rename(inipdfPath,revispdfPath)
    
    if iniCsv != 'no.csv' :
        inicsvPath = reptPath + reptFolder + '\\' + iniCsv    #get initial csv path
        reviscsvPath = reptPath + reptFolder + '\\' + revisCsv  #get revised csv path
        os.rename(inicsvPath,reviscsvPath)
    
    m+=1

