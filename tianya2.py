# -*- coding: utf-8 -*-
"""
Created on Tue May  8 23:36:21 2018

@author: Administrator
"""

from requests_html import HTMLSession
import re
import xlsxwriter

class ZF():
    def __init__(self):       
        self.f = xlsxwriter.Workbook('G:\elearning\pl.xlsx')
        self.sheet2=self.f.add_worksheet()
        self.sheet3=self.f.add_worksheet()
        self.rowsTitle = [ u'标题',u'点击率',u'回复数',u'时间']
        self.sheet2.write_row('A1',self.rowsTitle)
        self.sheet3.write_row('A1',self.rowsTitle)
        
    def GetDetail(self):    
        session=HTMLSession()
        n=1
        for i in range(4):       
            r=session.get('http://bbs.tianya.cn/list.jsp?item=329&order=0&nextid={}&k=%E9%9D%9E%E6%B4%B2'.format(i))
            contents = r.html.find('div.mt5')[3].find('td')   
            for j in range(0,len(contents)-1,5):
                click=contents[j+2].text
                respon=contents[j+3].text
                time=contents[j+4].text
                title=contents[j].text
                self.sheet2.write('A{}'.format(n+1),title)
                self.sheet2.write('B{}'.format(n+1),click)
                self.sheet2.write('C{}'.format(n+1),respon)
                self.sheet2.write('D{}'.format(n+1),time)
                n+=1
                
    def GetDail2(self):
        session=HTMLSession()
        m=1
        for j in range(3):       
            u=session.get('http://bbs.tianya.cn/list.jsp?item=329&order=1&nextid={}&k=%E9%BB%91%E4%BA%BA'.format(j))
            contents = u.html.find('div.mt5')[3].find('td')   
            for j in range(0,len(contents)-1,5):
                click1=contents[j+2].text
                respon1=contents[j+3].text
                time1=contents[j+4].text
                title1=contents[j].text
                self.sheet3.write('A{}'.format(m+1),title1)
                self.sheet3.write('B{}'.format(m+1),click1)
                self.sheet3.write('C{}'.format(m+1),respon1)
                self.sheet3.write('D{}'.format(m+1),time1)
                m+=1
        self.f.close()
                
    def GetLink():
        session=HTMLSession()
        g=1
        for i in range(3):       
            r=session.get('http://bbs.tianya.cn/list.jsp?item=329&order=0&nextid={}&k=%E9%BB%91%E4%BA%BA'.format(i))
            contents = r.html.find('div.mt5')[3].absolute_links
            hh=re.compile(r'http://bbs.tianya.cn/.*?.shtml')
            rp=re.findall(hh,str(contents))
            for link in rp:          
                r=session.get(link)
                contents = r.html.find('div.bbs-content')
                for content in contents:
                    y=content.text
                    print(y)
    
    
    def GetLink3():
        session=HTMLSession()
        for i in range(3):       
            r=session.get('http://bbs.tianya.cn/list.jsp?item=329&order=0&nextid={}&k=%E9%9D%9E%E6%B4%B2'.format(i))
            contents = r.html.find('div.mt5')[3].absolute_links
            hh=re.compile(r'http://bbs.tianya.cn/.*?.shtml')
            rp=re.findall(hh,str(contents))
            for link in rp:          
                r=session.get(link)
                contents = r.html.find('div.bbs-content')
                for content in contents:
                    y=content.text
                    print(y)

GetLink() 

