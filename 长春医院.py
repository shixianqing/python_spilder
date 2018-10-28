#!/usr/bin/python
# -*- coding: UTF-8 -*-

import requests
from lxml import etree
import xlrd
import xlwt
from xlutils.copy import copy
import os
from bs4 import BeautifulSoup
import time
url = "http://222.168.33.108:8899/yyselect?yy="

"""
    params : 请求参数字典
"""
def sendReq(params):
    print("请求第"+str(params["pageNo"])+"页数据")
    reponse = requests.post(url,data=params)
    htmlStr = reponse.text
    docment = etree.HTML(htmlStr)
    result = {"html":docment,"pageNo":params["pageNo"]}
    # parseDoc(docment, params["pageNo"])
    parseDoc(result)


"""
    row : 行数
    list : 数据集合
    currPage 当前页
    wk : 工作簿
    sheetc : 工作表
    fileName : 文件名称
"""
def writeData(row,list, currPage, wk, sheetc,fileName):
    # 计算从哪行开始写
    pageNum = (currPage-1)*15
    # 遍历单元格数
    for j in range(len(list)):
        sheetc.write(pageNum + row,j,list[j])
    wk.save(fileName)


"""
@:param row 行数
@:param list 数据集合
@:param currPage 当前页
@:param fileName 保存文件名称
"""
def writeExcel(row,list,currPage,fileName):
    print("开始写入第" + str(currPage) + "页数据============" + ','.join(list))
    isExsit = os.path.exists(fileName)
    if isExsit:
        xls = xlrd.open_workbook(fileName)
        xlsc = copy(xls)
        sheetc = xlsc.get_sheet(0)
        writeData(row,list,currPage,xlsc,sheetc,fileName)
    else:
        wk = xlwt.Workbook(encoding="utf-8");
        sheet = wk.add_sheet("医院", cell_overwrite_ok=True)
        sheet.write(0,0,"医院名称")
        sheet.write(0,1,"地址")
        sheet.write(0,2,"所属地区")
        writeData(row,list,currPage,wk,sheet,fileName)



def parseDoc(result):
    doc = result["html"]
    currPage = result["pageNo"]
    liTag = doc.xpath("/html/body/form/div/div[2]/div[2]/div[2]/div/ul/li")
    for i in range(len(liTag)):
        li = liTag[i]
        div = li.xpath("/html/body/form/div/div[2]/div[2]/div[2]/div/ul/li[{}]/div".format(i+1))
        list = []
        for j in range(len(div)):
            text = div[j].text.replace(" ","").replace("\r","").replace("\t","").replace("\n","")
            list.append(text)
        # 开始写入一行数据
        writeExcel(i,list, currPage, "d:/长春医院.xls")





# def spider():
#     for i in range(1,12311):
#         params = {"pageNo": i, "totalPageCount": 12310}
#         sendReq(params)
#
#
# spider()

def main():
    startTime = time.time()

    for i in range(1,182):#12311
        params = {"pageNo": i, "totalPageCount": 181}
        sendReq(params)
    endTime = time.time()

    print("耗时======="+str(endTime-startTime))

if __name__ == '__main__':
    main()



"""
    下载多线程
    解析多线程

"""
