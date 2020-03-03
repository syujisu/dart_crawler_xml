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

#2018년 사업보고서를 찾아서 엑셀을 완성하십시오.
# 2. 크롤링
# https://dart.fss.or.kr/dsaf001/main.do?rcpNo=(예:유진증권 20190401002982)
# 를 하면 좌측 메뉴에 III.재무에 관한사항 - 4.재무제표를 선택하면
# 우측에 나오는 "가. 재무상태표"의 값을 읽어서 Excel에 적용 하시면 됩니다.
# - 둘 중 하나를 선택하여 구현하시면 되고. 둘다 동일한 값입니다.
# - 포괄상태표"이하는 무시하십시오.
# - 유진기업, 동양, 유진증권 3개 기업을 대상으로 하십시오

#유진기업 2018 재무제표 url
#http://dart.fss.or.kr/report/viewer.do?rcpNo=20190401004107&dcmNo=6612848&eleId=15&offset=1908112&length=146218&dtd=dart3.xsd

#동양
#http://dart.fss.or.kr/report/viewer.do?rcpNo=20190814001889&dcmNo=6845532&eleId=17&offset=1354185&length=99327&dtd=dart3.xsd

#유진증권
#http://dart.fss.or.kr/report/viewer.do?rcpNo=20190401002982&dcmNo=6607495&eleId=15&offset=1221195&length=219209&dtd=dart3.xsd


import time
from bs4 import BeautifulSoup
import urllib.parse as parser
import selenium
import pandas as pd
import os
from html_table_parser import parser_functions as parser
from openpyxl import load_workbook
from urllib.request import urlopen
from bs4 import BeautifulSoup
from selenium import webdriver
import traceback


def url_to_excel(com_url):
    api_key = "fbd3f31ee413a318c81b0fe2bc0ad8b283dcfe21"

    if com_url == "http://dart.fss.or.kr/report/viewer.do?rcpNo=20190401002982&dcmNo=6607495&eleId=15&offset=1221195&length=219209&dtd=dart3.xsd":
        report= urlopen(com_url)
        r=report.read()

        xmlsoup=BeautifulSoup(r,'html.parser')
        body = xmlsoup.find("body")
        table = body.find_all("table")
        p = parser.make2d(table[1])

        df = pd.DataFrame(p[0:])
        header = df.iloc[0]
        df.rename (columns = header , inplace = True)
        df = df.reset_index(inplace=True)

    else:
        try:
            report=urlopen(com_url)
            r=report.read()

            xmlsoup=BeautifulSoup(r,'html.parser')
            body = xmlsoup.find("body")
            table = body.find_all("table")
            p = parser.make2d(table[1])

            df = pd.DataFrame(p[0:])
            header = df.iloc[0]

            df.rename (columns = header , inplace = True)
            df=df.set_index('')
        
        except Exception as e :
            print(traceback.format_exc())


    if not os.path.exists('output.xlsx'):#파일 초기에 생성하기 위해 유진기업은 mode = "w"로 지정!
        with pd.ExcelWriter('output.xlsx', mode='w', engine='openpyxl') as writer:
            if com_url == "http://dart.fss.or.kr/report/viewer.do?rcpNo=20190401004107&dcmNo=6612848&eleId=15&offset=1908112&length=146218&dtd=dart3.xsd":
                df.to_excel(writer, sheet_name = '유진기업', startrow = 1, startcol = 1) 
                writer.save()
                writer.close()

            else :
                print("url이 잘못되었습니다.")
                
    else:#만약 이미 파일이 존재한다면 그 파일에 시트를 append! 
        with pd.ExcelWriter('output.xlsx', mode='a', engine='openpyxl') as writer:
            if com_url == "http://dart.fss.or.kr/report/viewer.do?rcpNo=20190814001889&dcmNo=6845532&eleId=17&offset=1354185&length=99327&dtd=dart3.xsd":
                df.to_excel(writer, sheet_name = '동양', startrow = 1, startcol = 1)
                writer.save()
                writer.close()
            
            elif com_url == "http://dart.fss.or.kr/report/viewer.do?rcpNo=20190401002982&dcmNo=6607495&eleId=15&offset=1221195&length=219209&dtd=dart3.xsd":
                df.to_excel(writer, sheet_name = '유진증권', startrow = 1, startcol = 1)
                writer.save()
                writer.close()

            else :
                print("잘못입력하셨습니다.")
    return 



#실행 코드 
url_list = ["http://dart.fss.or.kr/report/viewer.do?rcpNo=20190401004107&dcmNo=6612848&eleId=15&offset=1908112&length=146218&dtd=dart3.xsd", #유진기업
            "http://dart.fss.or.kr/report/viewer.do?rcpNo=20190814001889&dcmNo=6845532&eleId=17&offset=1354185&length=99327&dtd=dart3.xsd", #동양
            "http://dart.fss.or.kr/report/viewer.do?rcpNo=20190401002982&dcmNo=6607495&eleId=15&offset=1221195&length=219209&dtd=dart3.xsd"] #유진증권

for com_url in url_list:
    url_to_excel(com_url)