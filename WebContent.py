# -*- coding: utf-8 -*-
# @Time    : 11/28/2018 17:11
# @Author  : MARX·CBR
# @File    : __init__.py.py
import lxml
from bs4 import BeautifulSoup

class get_content:
    def __init__(self):
        self.html=""
        with open('text.txt','rb') as c:
            self.html=c.read().decode()
        self.soup=BeautifulSoup(self.html,'lxml')
    def show(self,cx,cy):
        x=int(cx)
        y=int(cy)
        localx=14
        localy=30
        xline=int(x/localx)
        # 确定行
        yline=int(y/localy+1)
        content=self.soup.find('span',href="#{}".format(yline)).get_text()
        return content[xline:xline+1:]

class get_fu_content:
    def __init__(self):
        self.html=""
        with open('fu.txt','rb') as c:
            self.html=c.read().decode()
        self.soup=BeautifulSoup(self.html,'lxml')
    def show(self,cx,cy):
        x=int(cx)
        y=int(cy)
        localx=14
        localy=30

        xline=int(x/localx)
        yline=int(y/localy+1)
        content=self.soup.find('span',href="#{}".format(yline)).get_text()
        return content[xline:xline+1:]

class get_css:
    def __init__(self):

        self.html=""
        # self.content=""
        with open('css.txt','rb') as c:
            self.html=c.read().decode()
        self.soup=BeautifulSoup(self.html,'lxml')
        self.mydict={}
    def get_position(self,name):
        return self.soup.find('span',href='{}'.format(name)).get_text()

    def add_value_in_dict(self):
        v=self.soup.findAll('span')
        for i in v:
            self.mydict['{}'.format(i.get('href'))]=i.get_text()
        print(self.mydict)
        return True

# p=get_css()
# print(p.mydict)
# print(p.add_value_in_dict())
# print(p.mydict)
# print(p.mydict['my-e3r3'])

# print(get_css().get_position('ka-1ARY'))