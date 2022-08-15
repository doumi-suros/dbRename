import os
import pandas


reptPath = 'C:\\SurosBI\\Panorays_Reports_220810A\\'
excelFile = 'dbData_rename_220812.xlsx'
reptList = pandas.read_excel(excelFile, sheet_name=0, usecols=[1,2,3,4,5,6,7])
reptHead = ['報告檔案夾', 'Date', 'pdfName', 'csvName', 'NewDate','NewpdfName','NewcsvName']


for m in range (15, 22):
	reptFolder = reptList[reptHead[0]].values[m].strip()    #get report dir
	iniPdf = reptList[reptHead[2]].values[m].strip()    #get initial pdf file name
	iniCsv = reptList[reptHead[3]].values[m].strip()    #get initial csv file name
	revisPdf = reptList[reptHead[5]].values[m].strip()	 #get revised pdf file name
	revisCsv = reptList[reptHead[6]].values[m].strip()	 #get revised csv file name

	while iniPdf != 'no.pdf':
		inipdfPath = reptPath + reptFolder + '\\' + iniPdf    #get initial pdf path
		revispdfPath = reptPath + reptFolder + '\\' + revisPdf  #get revised pdf path
		print(m)
		os.rename(inipdfPath,revispdfPath)
		break

	while iniCsv != 'no.csv':
		inicsvPath = reptPath + reptFolder + '\\' + iniCsv    #get initial csv path
		reviscsvPath = reptPath + reptFolder + '\\' + revisCsv  #get revised csv path
		os.rename(inicsvPath,reviscsvPath)
		break


	
	m+=1



# ##開Excel取出目前檔案路徑的值A
# old_name = "C:\\Users\\Esther.Chang\\Desktop\\LAB\\NEW_TEST.xlsx"

# ##開Excel取出新的檔案路徑的值B
# new_name = "C:\\Users\\Esther.Chang\\Desktop\\LAB\\NEW_NEW_TEST.xlsx"


# ##用os更改檔名

# if os.path.isfile(old_name):
# 	os.rename(old_name, new_name)
# else:
# 	print("There is no such file.")


###改成while迴圈
	# if iniPdf == 'no.pdf':
	# 	m+=1
	# 	continue
	
	# inipdfPath = reptPath + reptFolder + '\\' + iniPdf    #get pdf path
	# revispdfPath = reptPath + reptFolder + '\\' + revisPdf  #get Newpdf path		

	# os.rename(inipdfPath,revispdfPath)