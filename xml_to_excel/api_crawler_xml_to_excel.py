 # 1. xml 다운로드
# 공시정보-공시보고서원본파일
# https://opendart.fss.or.kr/api/document.xml
# ?crtfc_key=*****&rcept_no=(예:유진증권 20190401002982)
# 를 동작시키면 접수번호별 zip파일 나오고 압축을 풀면 xml이 3개 나옵니다.

# 여기부터 코딩하시면 됩니다.
# 접수번호.xml에서 "4. 재무제표" 를 찾아서(42000Line정도 있음)
# 그 이후 값들을 읽어서 Excel에 적용 하십시오
# (공시원본파일을 수행시켜서 다운로드 하고 압축 푸는 것은 수작업으로 하십시오.
# 코딩할 필요는 없습니다. Excel화 하는 것이 우선입니다. 다 했는데도 시간이 남으면
# 뭐.. 코딩하여 자동화 하는 것 말리지는 않겠습니다.)

# 20190401004107 유진
# https://opendart.fss.or.kr/api/document.xml?crtfc_key=fbd3f31ee413a318c81b0fe2bc0ad8b283dcfe21&rcept_no=20190401004107
# 20190814001889 동양
# 20190401002982 유진증권
import time
from bs4 import BeautifulSoup
import urllib.parse as parser
import selenium
import pandas as pd
from html_table_parser import parser_functions as parser
from urllib.request import urlopen
from bs4 import BeautifulSoup
from selenium import webdriver
import traceback
import os



euxml=open(r'C:\Users\user\Desktop\document\1.xml',"r")
eusoup=BeautifulSoup(euxml,'html.parser')
eusoup.encoding='utf-8'
body = eusoup.find('body')
table = body.findAll('table')


for i in range(len(table)):
    a=pd.DataFrame(parser.make2d(table[i]))
    
    if a.iloc[0,0]=='재무상태표':
        df=pd.DataFrame(parser.make2d(table[i+1]))
        break

# fp = open(r"C:\Users\user\Desktop\document\1.xml", "r")

# soup = BeautifulSoup(fp, 'html.parser')
