from sina_login import Login
from spider import CollectData
from analysis import SemanticAnalysis
import openpyxl
import requests

if __name__=="__main__":
    keyword = '华泰'
    startTime = '2018-08-13'
    interval = '40'
    excelPath = 'data/weibo.xlsx'
    session=requests.session()
    #cd=CollectData(keyword,startTime,excelPath,session,interval)
    sa = SemanticAnalysis(startTime, keyword, excelPath)

    print("完成")