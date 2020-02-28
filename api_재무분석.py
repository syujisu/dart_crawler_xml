import pandas as pd
import requests
import json
import pandas as pd
from bs4 import BeautifulSoup
import webbrowser
from urllib.request import urlopen
from pandas.io.json import json_normalize
import os
from html_table_parser import parser_functions as parser
import traceback
import openpyxl
from openpyxl import Workbook

def company_to_excel(com_no): #유진기업
    
    API_KEY=""
    com_no = str(input('''회사 코드를 입력하세요 
    - 유진기업 : 00184667 
    - 동양 : 00117337 
    - 유진증권 : 00131054
    - 한일합성(비상장) : 00163266
    '''))

    total_dataframe = pd.DataFrame(columns = ['접수번호', '고유번호', '종목 코드', '계정명', '개별/연결구분', '개별/연결명', '재무제표구분', '재무제표명',
       '당기명', '당기일자', '당기금액', '전기명', '전기일자', '전기금액', '계정과목 정렬순서', '당기누적금액',
       '전기누적금액'])


    total_dataframe2 = pd.DataFrame(columns = ['접수번호', '고유번호', '종목 코드', '계정명', '개별/연결구분', '개별/연결명', '재무제표구분', '재무제표명',
       '당기명', '당기일자', '당기금액', '전기명', '전기일자', '전기금액', '계정과목 정렬순서', '당기누적금액',
       '전기누적금액'])
    for i in range(11011, 11015):
        for j in range(2016, 2020):
            com_url = "https://opendart.fss.or.kr/api/fnlttSinglAcnt.json?crtfc_key="+API_KEY+"&corp_code="+com_no+"&bsns_year="+str(j)+"&reprt_code="+str(i)

            try:
                response = requests.get(com_url)
                data = json.loads(response.content)
                df = json_normalize(data['list'])
                df_dataframe = pd.DataFrame(df)
                df_dataframe.rename(columns={'rcept_no':'접수번호',
                              'corp_code':'고유번호',
                             'stock_code':'종목 코드',
                              'account_nm':'계정명',
                              'fs_div':'개별/연결구분',
                              'fs_nm':'개별/연결명',
                              'sj_div':'재무제표구분',
                              'sj_nm':'재무제표명',
                              'thstrm_nm':'당기명',
                              'thstrm_dt':'당기일자',
                              'thstrm_amount':'당기금액',
                              'frmtrm_nm':'전기명',
                              'frmtrm_dt':'전기일자',
                              'frmtrm_amount':'전기금액',
                              'bfefrmtrm_nm': '전전기명',
                              'bfefrmtrm_dt': '전전기일자',
                              'bfefrmtrm_amount': '전전기금액',
                              'ord': '계정과목 정렬순서',
                            'thstrm_add_amount' : '당기누적금액',
                            'frmtrm_add_amount' : '전기누적금액',
                             },inplace=True)
                total_dataframe = pd.concat([total_dataframe, df_dataframe])
                total_dataframe2 = pd.concat([total_dataframe2, df_dataframe],sort=True)
                i += 1
                j += 1



            #유진증권과 한일합성은 자료가 없기 때문에 pass 처리 
            except BaseException:
                pass



    #재무제표 엑셀화 데이터 선택
    total_dataframe=total_dataframe[total_dataframe['재무제표명']=='재무상태표']
    total_dataframe=total_dataframe[total_dataframe['개별/연결명']=='재무제표']
    total_dataframe=total_dataframe[['계정명', '당기일자','당기금액']]

    #손익계산서 엑셀화 데이터 선택
    total_dataframe2=total_dataframe2[total_dataframe2['재무제표명']=='손익계산서']
    total_dataframe2=total_dataframe2[total_dataframe2['개별/연결명']=='연결재무제표']
    total_dataframe2=total_dataframe2[['계정명', '당기일자','당기금액']]

    #피벗테이블 사용
    output2 = total_dataframe2.pivot_table(values = "당기금액", index = "계정명", columns = "당기일자" ,aggfunc='first')
    output = total_dataframe.pivot_table(values = "당기금액", index = "계정명", columns = "당기일자" ,aggfunc='first')

    if not os.path.exists('output.xlsx'):
        with pd.ExcelWriter('output.xlsx', mode='w', engine='openpyxl') as writer:
            if com_no == "00184667":
                output.to_excel(writer, sheet_name = '유진기업', startrow = 1, startcol = 1)
                output2.to_excel(writer, sheet_name = '유진기업', startrow = 14, startcol = 1)
                writer.save()
                writer.close()
                
            elif com_no == "00117337":
                output.to_excel(writer, sheet_name = '동양', startrow = 1, startcol = 1)
                output2.to_excel(writer, sheet_name = '동양', startrow = 14, startcol = 1)
                writer.save()
                writer.close()
            
            elif com_no == "00131054":
                output.to_excel(writer, sheet_name = '유진증권', startrow = 1, startcol = 1)
                output2.to_excel(writer, sheet_name = '유진증권', startrow = 14, startcol = 1)
                writer.save()
                writer.close()
            
            elif com_no == "00163266":
                output.to_excel(writer, sheet_name = '한일합성', startrow = 1, startcol = 1)
                output2.to_excel(writer, sheet_name = '한일합성', startrow = 14, startcol = 1)
                writer.save()
                writer.close()
                
            else :
                print("잘못입력하셨습니다.")
    else:
        with pd.ExcelWriter('output.xlsx', mode='a', engine='openpyxl') as writer:
            if com_no == "00184667":
                output.to_excel(writer, sheet_name = '유진기업', startrow = 1, startcol = 1)
                output2.to_excel(writer, sheet_name = '유진기업', startrow = 14, startcol = 1)
                writer.save()
                writer.close()
                
            elif com_no == "00117337":
                output.to_excel(writer, sheet_name = '동양', startrow = 1, startcol = 1)
                output2.to_excel(writer, sheet_name = '동양', startrow = 14, startcol = 1)
                writer.save()
                writer.close()
            
            elif com_no == "00131054":
                output.to_excel(writer, sheet_name = '유진증권', startrow = 1, startcol = 1)
                output2.to_excel(writer, sheet_name = '유진증권', startrow = 14, startcol = 1)
                writer.save()
                writer.close()
            
            elif com_no == "00163266":
                output.to_excel(writer, sheet_name = '한일합성', startrow = 1, startcol = 1)
                output2.to_excel(writer, sheet_name = '한일합성', startrow = 14, startcol = 1)
                writer.save()
                writer.close()
            
            else :
                print("잘못입력하셨습니다.")
    return 