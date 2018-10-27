#!/usr/bin/python
# -*- coding: UTF-8 -*-
# http请求包
import urllib3

# 下载pip install beautifulsoup4
from bs4 import BeautifulSoup

import xlwt
import os
import xlrd
from xlutils.copy import copy
import requests
import json

def process(formdata):

    print("------------请求第"+str(formdata["curPage"])+"页数据----------------")
    # 请求url
    url = "http://gdrst.gdhrss.gov.cn//sofpro/otherproject/yaopin/yaopin.jsp"

    # 添加User-Agent，完善请求信息
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko)"
                             "Chrome/51.0.2704.103 Safari/537.36",
               "Content-Type":"application/x-www-form-urlencoded"}
    # 创建一个http
    # http = urllib3.PoolManager()

    # 创建请求
    # r = http.request("POST", url, fields=formdata, headers=headers)
    r = requests.post(url,data = formdata)
    # 获取响应数据
    data = r.text
    doc = BeautifulSoup(data, "lxml")

    trTag = doc.find_all("tr", attrs={"bgcolor": "#F7D8E0"})
    # 解析元素数据
    print("------------------------准备解析数据----------------------")
    writeExcel(trTag,formdata["curPage"])
    return trTag

def parseData( trTag ):
    with open("d:/guangdongMedicine.txt", "a") as file:
        for tr in trTag:
            tds = tr.find_all("td")
           # list = []
            leng = len(tds)
            for index in range(leng):
                text = tds[index].text
                if index == leng - 1:
                    file.write(text + "\r" + "\n")
                else:
                    file.write(text + "\t")
                # list.append(text)
                print(text)
    file.close()
    return


def writeData(wk,sheetc, trTag,currPage):

    pageNum = (int(currPage) - 1) * 20
    for i in range(len(trTag)):
        tds = trTag[i].find_all("td")
        cells = []
        leng = len(tds)
        for index in range(leng):
            text = tds[index].text
            cells.append(text)
        for cell in range(len(cells)):
            sheetc.write(pageNum+i+1, cell, cells[cell])
    wk.save("d:/ghuangdongPaoPin.xls")


def writeExcel(trTag,currPage):

    # 判断文件是否存在，存在，则追加内容，反之创建
    isExsit = os.path.exists("d:/ghuangdongPaoPin.xls")
    if isExsit:
        # 打开文件
        xls = xlrd.open_workbook("d:/ghuangdongPaoPin.xls", formatting_info=True)
        xlsc = copy(xls) # 复制文件
        # 获取sheet
        sheetc = xlsc.get_sheet(0)
        writeData(xlsc,sheetc,trTag,currPage)
    else:
        # 创建工作簿
        wk = xlwt.Workbook(encoding="utf-8")
        # 创建工作表
        sheet = wk.add_sheet("yaopin", cell_overwrite_ok=True)
        sheet.write(0, 0, "中文名称")
        sheet.write(0, 1, "分类")
        sheet.write(0, 2, "剂型")
        sheet.write(0, 3,"备注")
        sheet.write(0, 4, "编号")
        sheet.write(0, 5, "大类")
        sheet.write(0, 6, "中类")
        sheet.write(0, 7, "小类")
        sheet.write(0, 8, "细类")
        sheet.write(0, 9, "英文名称")
        writeData(wk,sheet,trTag,currPage)



for i in range(1,145):

    formdata = {
        "curPage":  i,
        "totalPages": 144
    }
    process(formdata)






