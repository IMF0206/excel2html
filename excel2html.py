#coding=utf-8
'''
Created on 2016-08-25
 
@author: PJL

function: excel2html
'''
import os,sys
from win32com.client.gencache import EnsureDispatch
from win32com.client import constants
 
def GetExcelList(dir, fileList):
    newDir = dir
    if os.path.isfile(dir):
        fileList.append(dir.decode('gbk'))
    elif os.path.isdir(dir):  
        for s in os.listdir(dir):
            isExcel = s.find('xls')
            if isExcel == -1:
                continue
            newDir=os.path.join(dir,s)
            GetExcelList(newDir, fileList)  
    return fileList


yourpath = sys.argv[1]
newpath = yourpath+'/excel2html'
if not os.path.exists(newpath):
    os.mkdir(newpath)
list = GetExcelList(yourpath, [])
xl = EnsureDispatch('Excel.Application')
for file in list:
    yourExcelFile = os.path.split(file)[1]
    newFileName = newpath+'/'+yourExcelFile.split('.')[0]+'.htm'
    if os.path.exists(newFileName):
        continue
    try:
        wb = xl.Workbooks.Open(file)
        wb.SaveAs(newFileName, constants.xlHtml)
    except Exception,e:
        print newFileName,e
    xl.Workbooks.Close()
xl.Quit()
del xl