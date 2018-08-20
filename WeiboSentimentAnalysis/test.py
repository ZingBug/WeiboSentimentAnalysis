from sina_login import Login
from spider import CollectData
from analysis import SemanticAnalysis
import openpyxl
import requests
import os

if __name__=="__main__":
    keyword = '华泰'
    startTime = '2018-08-13'
    interval = '40'
    excelPath = 'data/weibo.xlsx'
    excelDir = 'data'
    session=requests.session()
    cd=CollectData(keyword,startTime,excelPath,excelDir,session,interval)
    #sa = SemanticAnalysis(startTime, keyword, excelPath)



    print("完成")