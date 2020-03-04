# 2018년 사업보고서를 찾아서 엑셀을 완성하십시오.
# 2. 크롤링
# https://dart.fss.or.kr/dsaf001/main.do?rcpNo=(예:유진증권 20190401002982)
# 를 하면 좌측 메뉴에 III.재무에 관한사항 - 4.재무제표를 선택하면
# 우측에 나오는 "가. 재무상태표"의 값을 읽어서 Excel에 적용 하시면 됩니다.
# - 둘 중 하나를 선택하여 구현하시면 되고. 둘다 동일한 값입니다.
# - 포괄상태표"이하는 무시하십시오.
# - 유진기업, 동양, 유진증권 3개 기업을 대상으로 하십시오

# 유진기업 2018 재무제표 url
# http://dart.fss.or.kr/report/viewer.do?rcpNo=20190401004107&dcmNo=6612848&eleId=15&offset=1908112&length=146218&dtd=dart3.xsd

# 동양
# http://dart.fss.or.kr/report/viewer.do?rcpNo=20190814001889&dcmNo=6845532&eleId=17&offset=1354185&length=99327&dtd=dart3.xsd

# 유진증권
# http://dart.fss.or.kr/report/viewer.do?rcpNo=20190401002982&dcmNo=6607495&eleId=15&offset=1221195&length=219209&dtd=dart3.xsd


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
        report = urlopen(com_url)
        r = report.read()

        xmlsoup = BeautifulSoup(r, 'html.parser')
        body = xmlsoup.find("body")
        table = body.find_all("table")
        p = parser.make2d(table[1])

        df = pd.DataFrame(p[0:], columns=['과목', '2018.12.31', '2018.12.31',
                                          '2017.12.31', '2017.12.31', '2016.12.31', '2016.12.31'])
        df = df.set_index('과목')

    else:
        try:
            report = urlopen(com_url)
            r = report.read()

            xmlsoup = BeautifulSoup(r, 'html.parser')
            body = xmlsoup.find("body")
            table = body.find_all("table")
            p = parser.make2d(table[1])

            df = pd.DataFrame(p[0:])
            header = df.iloc[0]

            df.rename(columns=header, inplace=True)
            df = df.set_index('')

        except Exception as e:
            print(traceback.format_exc())

    if not os.path.exists('output.xlsx'):  # 파일 초기에 생성하기 위해 유진기업은 mode = "w"로 지정!
        with pd.ExcelWriter('output.xlsx', mode='w', engine='openpyxl') as writer:
            if com_url == "http://dart.fss.or.kr/report/viewer.do?rcpNo=20190401004107&dcmNo=6612848&eleId=15&offset=1908112&length=146218&dtd=dart3.xsd":
                df.to_excel(writer, sheet_name='유진기업', startrow=1, startcol=1)
                writer.save()
                writer.close()

            else:
                print("url이 잘못되었습니다.")

    else:  # 만약 이미 파일이 존재한다면 그 파일에 시트를 append!
        with pd.ExcelWriter('output.xlsx', mode='a', engine='openpyxl') as writer:
            if com_url == "http://dart.fss.or.kr/report/viewer.do?rcpNo=20190814001889&dcmNo=6845532&eleId=17&offset=1354185&length=99327&dtd=dart3.xsd":
                df.to_excel(writer, sheet_name='동양', startrow=1, startcol=1)
                writer.save()
                writer.close()

            elif com_url == "http://dart.fss.or.kr/report/viewer.do?rcpNo=20190401002982&dcmNo=6607495&eleId=15&offset=1221195&length=219209&dtd=dart3.xsd":
                df.to_excel(writer, sheet_name='유진증권', startrow=1, startcol=1)
                writer.save()
                writer.close()

            else:
                pass
    return


# 실행 코드
url_list = ["http://dart.fss.or.kr/report/viewer.do?rcpNo=20190401004107&dcmNo=6612848&eleId=15&offset=1908112&length=146218&dtd=dart3.xsd",  # 유진기업
            "http://dart.fss.or.kr/report/viewer.do?rcpNo=20190814001889&dcmNo=6845532&eleId=17&offset=1354185&length=99327&dtd=dart3.xsd",  # 동양
            "http://dart.fss.or.kr/report/viewer.do?rcpNo=20190401002982&dcmNo=6607495&eleId=15&offset=1221195&length=219209&dtd=dart3.xsd"]  # 유진증권

for com_url in url_list:
    url_to_excel(com_url)
