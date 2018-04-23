# -*- coding: utf-8 -*-
"""
Created on Thu Apr 19 14:56:23 2018

@author: Administrator
"""

import requests
import xlwt
import xlrd


class DFCF(object):
    def __init__(self):
        self.f = xlwt.Workbook(encoding='utf-8',style_compression = 0) 
        self.sheet1=self.f.add_sheet('dete',cell_overwrite_ok=True)
        self.rowsTitle = [ u'序号',u'股票号码',u'股票名称',u'最新价',u'涨跌幅',u'主力净额',u'主力净占比',u'超大单净额',u'超大单净占比',u'大单净流入',u'大单净占比',u'中单净流入',u'中单净占比',u'小单净流入',u'小单净占比',u'日期',u'涨跌额']
        for i in range(0,len(self.rowsTitle)):
            self.sheet1.write(0, i, self.rowsTitle[i], self.set_style('Times new Roman', 220, True))
        self.f.save('e:\Elearning\gg.xls')
    
    def set_style(self,name,height,bold=False):
        style = xlwt.XFStyle()  # 初始化样式
        font = xlwt.Font()  # 为样式创建字体
        font.name = name
        font.bold = bold
        font.colour_index = 2
        font.height = height
        style.font = font
        return style
    
    def getURL(self):
        for i in range(3):
            url = "http://nufm.dfcfw.com/EM_Finance2014NumericApplication/JS.aspx?type=ct&st=(BalFlowMain)&sr=-1&p={}".format(i+1) + "&ps=50&js=var%20QvPvHhmY={pages:(pc),date:%222014-10-22%22,data:[(x)]}&token=894050c76af8597a853f5b408b759f5d&cmd=C._AB&sty=DCFFITA&rt=50804877"
            self.spiderPage(url)

    def spiderPage(self,url):
        if url is None:
            return u'不存在网页'
            
        try:
            data = xlrd.open_workbook('gg.xls')
            table = data.sheets()[0]
            rowCount = table.nrows
            headers = {'Accept': '*/*',
               'Accept-Language': 'en-US,en;q=0.8',
               'Cache-Control': 'max-age=0',
               'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.116 Safari/537.36',
               'Connection': 'keep-alive',
               'Referer': 'http://www.baidu.com/'
               } 
            m=0
            req = requests.get(url,headers = headers)
            req.encoding='utf-8'
            reqt=req.text
            reqt=reqt[47:].replace('"','')
            reqt_list=reqt.strip(',').split(',')
            for i in range(0,len(reqt_list),17):
                reqt_list2=[]
                stock_num=reqt_list[i+1]
                stock_name=reqt_list[i+2]
                stock_price=reqt_list[i+3]
                stock_chg=reqt_list[i+4]
                td_mf=reqt_list[i+5]
                td_mfp=reqt_list[i+6]
                td_sp=reqt_list[i+7]
                td_spp=reqt_list[i+8]
                td_big=reqt_list[i+9]
                td_bigp=reqt_list[i+10]
                td_mid=reqt_list[i+11]
                td_midp=reqt_list[i+12]
                td_small=reqt_list[i+13]
                td_sp=reqt_list[i+14]
                td_time=reqt_list[i+15]
                td_chgprice=reqt_list[i+16]
                
                reqt_list2.append(rowCount+m)
                reqt_list2.append(stock_num)
                reqt_list2.append(stock_name)
                reqt_list2.append(stock_price)
                reqt_list2.append(stock_chg)
                reqt_list2.append(td_mf)
                reqt_list2.append(td_mfp)
                reqt_list2.append(td_sp)
                reqt_list2.append(td_spp)
                reqt_list2.append(td_big)
                reqt_list2.append(td_bigp)
                reqt_list2.append(td_mid)
                reqt_list2.append(td_midp)
                reqt_list2.append(td_small)
                reqt_list2.append(td_sp)
                reqt_list2.append(td_time)
                reqt_list2.append(td_chgprice)
                print(reqt_list2)
                
                for i in range(len(reqt_list2)):
                 self.sheet1.write(rowCount+m,i,reqt_list2[i])
                m+=1
                print(m)
        except Exception as e:
             print(e)
        finally:
             self.f.save('e:\Elearning\gg.xls')
             
if '_main_':
    dfcf = DFCF()
    dfcf.getURL()