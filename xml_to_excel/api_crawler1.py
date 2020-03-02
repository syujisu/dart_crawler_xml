##필요한 모듈 설치 
import pandas as pd
import requests
import json
import pandas as pd
from bs4 import BeautifulSoup
import webbrowser
from urllib.request import urlopen
from pandas.io.json import json_normalize
import os
import traceback


# A.py 
# - 입력 : 인증키, 고유번호
# - 출력 : 대량보유 상황보고와 임원주요주주 소유보고

# A.py 
# - 입력 : 인증키, 고유번호
# - 출력 : 대량보유 상황보고와 임원주요주주 소유보고


#필요한 정보 입력받기

api_key = str(input("1.인증키를 입력하세요 : "))
com_no = str(input("2.기업 고유 번호를 입력하세요 : "))
rep_kind = int(input("3.보고서 종류를 숫자로 입력하세요 (1.대량보유 상황보고서 / 2.임원주요주주 소유보고서) : "))
file_path = input("4.엑셀 파일을 저장할 폴더명만 쓰세요 : (예:C:\py_temp) ")

#보고서 종류에 따라 excel로 정리 

if rep_kind == 1:
    try: 
        #대량보유상황보고서
        url1 = "https://opendart.fss.or.kr/api/majorstock.xml?crtfc_key="+api_key+"&corp_code="+com_no
        print(url1)
        resultXML1 = urlopen(url1)
        result1 = resultXML1.read()

        xmlsoup1 = BeautifulSoup(result1, 'html.parser')

        data = pd.DataFrame()
        te1 = xmlsoup1.findAll("list")
        
        #컬럼명 바꿔 정렬
        for t1 in te1:
            temp1 = pd.DataFrame(([[t1.rcept_no.string,t1.rcept_dt.string,t1.corp_name.string,t1.corp_code.string,t1.report_tp.string,
                        t1.repror.string,t1.stkqy.string,t1.stkqy_irds.string,t1.stkrt.string,t1.stkrt_irds.string,t1.ctr_stkqy.string,
                        t1.ctr_stkrt.string,t1.report_resn.string]]),
                        columns = ["접수번호","접수일자","회사명","종목코드","보고구분","대표보고자","보유주식등의 수","보유주식등의 증감",
                                   "보유비율","보유비율 증감","주요체결 주식등의 수","주요체결 보유비율","보고사유"])
            data = pd.concat([data,temp1])
            #둘 이상의 Dataframe이 동일한 컬럼을 갖고 있다는 가정에서 row가 늘어하는 형태로 데이터가 늘어납니다.
            file_nm = "유진기업_대량보유상황보고서.xlsx"
            data.to_excel(os.path.join(file_path, file_nm),encoding="euc-kr",index=False)
            
    except Exception as e :
            print(traceback.format_exc())
        
elif rep_kind ==2 :
    try:
        #임원주요주주 소유 보고서
        url2 = "https://opendart.fss.or.kr/api/elestock.xml?crtfc_key="+api_key+"&corp_code="+com_no
        print(url2)
        
        resultXML2 = urlopen(url2)
        result2 = resultXML2.read()

        xmlsoup2 = BeautifulSoup(result2, 'html.parser')
        status_list = xmlsoup2.findAll("status")[0].get_text()
        
        #조회된 데이터가 없을 때 처리 
        if status_list == "NON_DATA":
            print("조회된 데이터가 없습니니다.")
        else:
            data2 = pd.DataFrame()
            te2 = xmlsoup2.findAll("list")
            
            for t2 in te2:
                temp2 = pd.DataFrame(([[t2.rcept_no.string,t2.rcept_dt.string,t2.corp_name.string,t2.corp_code.string,
                            t2.repror.string,t2.isu_exctv_rgist_at.string,t2.isu_exctv_ofcps.string,t2.isu_main_shrholdr.string,t2.sp_stock_lmp_cnt.string,t2.sp_stock_lmp_irds_cnt.string,
                            t2.sp_stock_lmp_rate.string,t2.sp_stock_lmp_irds_rate.string]]),
                            columns = ["접수번호","접수일자","회사명","회사번호","보고자","발행 회사 관계 임원(등기여부)","발행 회사 관계 임원 직위","발행 회사 관계 주요 주주","특정 증권 등 소유 수",
                                       "특정 증권 등 소유 증감 수 ","특정 증권 등 소유 비율","특정 증권 등 소유 증감 비율"])
                data2 = pd.concat([data2,temp2])
                file_nm = "유진기업_임원주요주주 소유 보고서.xlsx"
                data2.to_excel(os.path.join(file_path, file_nm),encoding="euc-kr",index=False)

    except Exception as e :
            print(traceback.format_exc())
else:
    print("보고서 종류를 다시 선택해주세요")

    
#xml.html -> xml 파일로 저장 함수 
def to_xml(df, filename=None, mode='w'):
    def row_to_xml(row):
        xml = ['<item>']
        for i, col_name in enumerate(row.index):
            xml.append('  <field name="{0}">{1}</field>'.format(col_name, row.iloc[i]))
        xml.append('</item>')
        return '\n'.join(xml)
    res = '\n'.join(df.apply(row_to_xml, axis=1))

    if filename is None:
        return res
    with open(filename, mode) as f:
        f.write(res)
pd.DataFrame.to_xml = to_xml

print (temp1.to_xml())
temp1.to_xml('output.xml')
        