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
import urllib.parse as parser
import selenium
import pandas as pd
import zipfile
import traceback
import traceback
import os
import zipfile
import webbrowser
import shutil

from openpyxl.utils.datetime import to_excel
from html_table_parser import parser_functions as parser
from urllib.request import urlopen
from bs4 import BeautifulSoup
from selenium import webdriver
from bs4 import BeautifulSoup


# 00184667 유진기업
# 00131054 유진증권
# 00117337 동양
# 00163266 한일합섬 *
# 00165149 유진저축은행
# rcept_no 찾기


#변수 , 필요한 리스트 지정
company_code_list = ['00184667', '00117337', '00131054']
API_KEY = "fbd3f31ee413a318c81b0fe2bc0ad8b283dcfe21"
writer = pd.ExcelWriter('output_excel.xlsx', mode='w', engine='openpyxl')
reports = ['20190401004107', '20190401003691','20190401002982']

# report_no 찾기 : xml -> rcept_no

def report_no_find():

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



def download():
    # 회사 별 리포트 번호에 따라 다운로드 -> zip 파일 생성 -> 압축 해제

    if report == '20190401004107':
        url_eu = "https://opendart.fss.or.kr/api/document.xml?crtfc_key=" + API_KEY + "&rcept_no=" + report
        webbrowser.open(url_eu)
        time.sleep(5)


        os.mkdir('C:/Users/user/Downloads/유진기업'),
        os.rename('C:/Users/user/Downloads/document.xml',
                'C:/Users/user/Downloads/유진.zip'),

        shutil.move('C:/Users/user/Downloads/유진.zip','C:/Users/user/Downloads/유진기업/유진.zip'),
        os.chdir('C:/Users/user/Downloads/유진기업'),
        ex_zip = zipfile.ZipFile('C:/Users/user/Downloads/유진기업/유진.zip')
        ex_zip.extractall()
        ex_zip.close()

    elif report == '20190401003691':
        url_dong = "https://opendart.fss.or.kr/api/document.xml?crtfc_key=" + API_KEY + "&rcept_no=" + report
        webbrowser.open(url_dong)
        time.sleep(5)

        os.mkdir('C:/Users/user/Downloads/동양'),
        os.rename('C:/Users/user/Downloads/document.xml',
                'C:/Users/user/Downloads/동양.zip')

        shutil.move('C:/Users/user/Downloads/동양.zip','C:/Users/user/Downloads/동양/동양.zip'),
        os.chdir('C:/Users/user/Downloads/동양')
        ex_zip = zipfile.ZipFile('C:/Users/user/Downloads/동양/동양.zip')
        ex_zip.extractall()
        ex_zip.close()
        
    elif report == "20190401002982":
        url_fi = "https://opendart.fss.or.kr/api/document.xml?crtfc_key=" + API_KEY + "&rcept_no=" + report
        webbrowser.open(url_fi)
        time.sleep(5)

        os.mkdir('C:/Users/user/Downloads/유진증권'),
        os.rename('C:/Users/user/Downloads/document.xml',
                'C:/Users/user/Downloads/유진증권.zip')

        shutil.move('C:/Users/user/Downloads/유진증권.zip','C:/Users/user/Downloads/유진증권/유진증권.zip'),
        os.chdir('C:/Users/user/Downloads/유진증권')
        ex_zip = zipfile.ZipFile('C:/Users/user/Downloads/유진증권/유진증권.zip')
        ex_zip.extractall()
        ex_zip.close()
        
    else:
        print(traceback.format_exc())



def url_to_excel():

    #유진기업과 동양은 재무상태표라는 제목을 찾아 표출할 수 있었습니다.
    if report == "20190401004107":

        fp =   open(r'C:/Users/user/Downloads/유진기업/'+report+'.xml', 'r')
        soup = BeautifulSoup(fp, 'html.parser')
        body = soup.find('body')

        table1 = body.findAll('table')


        for i in range(len(table1)):
            a=pd.DataFrame(parser.make2d(table1[i]))

            if a.iloc[0,0]=='재무상태표':
                df1=pd.DataFrame(parser.make2d(table1[i+1]))
        
        df1.to_excel(writer, sheet_name = '유진기업',startrow = 1, startcol = 1)
        writer.save()

    elif report == "20190401003691":
        
        fp =   open(r'C:/Users/user/Downloads/동양/'+report+'.xml', 'r')
        soup = BeautifulSoup(fp, 'html.parser')
        body = soup.find('body')

        table1 = body.findAll('table')


        for i in range(len(table1)):
            a2=pd.DataFrame(parser.make2d(table1[i]))

            if a2.iloc[0,0]=='재무상태표':
                df2=pd.DataFrame(parser.make2d(table1[i+1]))
                
        df2.to_excel(writer, sheet_name = '동양',startrow = 1, startcol = 1)
        writer.save()


#하지만 유진증권은 형태가 다르고, xpath나 태그 또한 다르게 나와 테이블 순서인 번호로 찾아 데이터프레임을 생성하였습니다. 
    elif report == "20190401002982":
        
        fp =   open(r'C:/Users/user/Downloads/유진증권/'+report+'.xml', 'r')
        soup = BeautifulSoup(fp, 'html.parser')
        body = soup.find('body')

        table2 = body.findAll('table')


        for i in range(len(table2)):
            df3=pd.DataFrame(parser.make2d(table2[306]))
            
        df3.to_excel(writer, sheet_name = '유진증권',startrow = 1, startcol = 1)
        writer.save()
        writer.close()


    else:
        print(traceback.format_exc())

                


#runnable main
if __name__ == "__main__":
    for company_code in company_code_list:
        report_no_find()



    for report in reports:
        download()
        url_to_excel()