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
# https://opendart.fss.or.kr/api/document.xml?crtfc_key=fbd3f31ee413a318c81b0fe2bc0ad8b283dcfe21&rcept_no=20190401003691

# 20190401002982 유진증권
# https://opendart.fss.or.kr/api/document.xml?crtfc_key=fbd3f31ee413a318c81b0fe2bc0ad8b283dcfe21&rcept_no=20190401002982
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
import zipfile

# 00184667 유진기업
# 00131054 유진증권
# 00117337 동양
# 00163266 한일합섬 *
# 00165149 유진저축은행
# rcept_no 찾기

company_code_list = ['00184667', '00117337', '00131054']


# report_no 찾기 : xml -> rcept_no
def report_no_make():
    API_KEY = "fbd3f31ee413a318c81b0fe2bc0ad8b283dcfe21"

    url = "https://opendart.fss.or.kr/api/list.xml?crtfc_key="+API_KEY+"&corp_code="+company_code + \
        "&bgn_de=20160101&end_de=20191231&pblntf_ty=A&pblntf_detail_ty=A002&page_no=1&page_count=10"

    resultXML = urlopen(url)
    result = resultXML.read()
    xmlsoup = BeautifulSoup(result, 'html.parser')
    data = pd.DataFrame()
    te = xmlsoup.findAll("list")

    for t in te:
        temp = pd.DataFrame(([[t.corp_cls.string, t.corp_name.string, t.corp_code.string, t.stock_code.string,
                               t.report_nm.string, t.rcept_no.string, t.flr_nm.string, t.rcept_dt.string, t.rm.string]]),
                            columns=["corp_cls", "corp_name", "corp_code", "stock_code", "report_nm", "rcept_no", "flr_nm", "rcept_dt", "rm"])
        data = pd.concat([data, temp])

    data = data.reset_index(drop=True)

    report_code = data[data['report_nm'] == '사업보고서 (2018.12)']
    rcept_no = report_code['rcept_no']
    print(rcept_no)

    return rcept_no


for company_code in company_code_list:
    report_no_make()


# fp = open(r'C:\Users\user\Desktop\document\1.xml', "r")
# soup = BeautifulSoup(fp, 'html.parser', encoding='utf-8')
# body = soup.find('body')
# table = body.findAll('table')


# for i in range(len(table)):
#     a = pd.DataFrame(parser.make2d(table[i]))

#     if a.iloc[0, 0] == '재무상태표':
#         df = pd.DataFrame(parser.make2d(table[i+1]))
#         break

# # fp = open(r"C:\Users\user\Desktop\document\1.xml", "r")

# # soup = BeautifulSoup(fp, 'html.parser')
