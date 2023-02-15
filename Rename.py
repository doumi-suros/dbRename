import os    # for rename and logs
import pandas    # for reading excel format records

excelFile = 'C:\\Users\\chesh\\Desktop\\testRun\\服務.xlsx'    # excel record file name
dbDirPath = 'C:\\Users\\chesh\\Desktop\\testRun\\'    # directory path till where db layer 絕對路徑

reptList = pandas.read_excel(excelFile, sheet_name=0, usecols=[1,3,4,5])    # set excel record reading columns
reptHead = ['報告檔案夾', 'summary.pdf', 'detail.pdf','issue.csv' ]    # set select columns' header

for m in range (0, 119):    # rename range per excel record lines, count from "0", excluded end number   for excel
    dirPath = reptList[reptHead[0]].values[m].strip()    # get company folder 從2行開始 strip去前後空格

    if not os.path.exists(dirPath):    # verify if company folder existed
        logFile = open ('missingFolderLog.txt', 'a+')
        logFile.write(str(m+2) + ', ' + dirPath + '\n')    # add-in log file if company folder not existed
        m += 1
        continue

    iniDirList = os.listdir (dirPath)    # get file list of company folder
    #print (str(m) + ': ' + dirPath + '\n', iniDirList)
    sortDirList = ['', '', '']    # set sorted list of company folder
    findKey = ['Scorecard', 'Detailed-Report', 'IssuesReport']    # set mapping key words
    
    for i in range (0,3):    # for sorting file list order
        a=0
        for k in iniDirList:    # for finding file name position of initial getting file list
            if findKey[i] in k:    # find file name by mapping key words
                sortDirList[i] = iniDirList[a]    # put found file name into new list
            a += 1

        i += 1
    print (sortDirList)
       
    for n in range (0, 3):  # for company folder
        oldNamePath = dbDirPath + dirPath + '\\' + sortDirList[n]    # get original file name from file list 抓取舊檔案名
        revNamePath = dbDirPath + dirPath + '\\' + reptList[reptHead[n+1]].values[m].strip()    # get new file name from excel record 位置對好
        #print (oldNamePath+'\n', revNamePath)
        os.rename(oldNamePath, revNamePath) 
        n += 1
    
    #reDirList = os.listdir (dirPath)    # get file list after rename
    #print (reDirList)

    m += 1