##필요한 모듈 설치 
import pandas as pd
import requests
import json
import pandas as pd
from bs4 import BeautifulSoup
import webbrowser
from urllib.request import urlopen
from pandas.io.json import json_normalize
import traceback
import os

# B.py
# - 입력 : 인증키, 고유번호, 사업년도, 보고서 코드
# - 출력 : 증자(감자)현황 배당에 관한 사항, 자기취득 및 처분현황, 최대 주주현황, 최대주주 변동현황, 
#            소액 주주현황, 임원현황 직원 현황, 타법인 출자현황, 단일회사 주요계정, 다중회사 주요계정

### 필요한 정보 입력받기

api_key = str(input("1.인증키를 입력하세요 : "))
com_no = str(input("2.기업 고유 번호를 입력하세요 : "))
bsns_year = str(input("3. 사업년도를 입력하세요 : "))
rep_code = str(input("4. 보고서 코드를 숫자로 입력하세요 (1분기보고서 / 11013 | 반기보고서 : 11012 | 3분기보고서 : 11014 |사업보고서 : 11011) : "))
rep_kind = int(input('''5. 보고서 종류를 숫자로 입력해주세요 
        1. 증자(감자)현황 배당에 관한 사항    2. 자기취득 및 처분현황
        3. 최대 주주현황    4. 최대주주 변동현황
        5. 소액 주주현황    6. 임원현황 & 직원현황
        7. 타법인 출자현황  8. 단일회사 주요계정
        9. 다중회사 주요계정
'''))
file_path = input("6.엑셀 파일을 저장할 폴더명만 쓰세요 : (예:C:\py_temp) ")


if rep_kind == 1:
    try:
        #증자(감자)현황 배당에 관한 사항
        url3 = "https://opendart.fss.or.kr/api/irdsSttus.xml?crtfc_key="+api_key+"&corp_code="+com_no+"&bsns_year="+bsns_year+"&reprt_code="+rep_code
        print(url3)
        
        resultXML3 = urlopen(url3)
        result3 = resultXML3.read()
        xmlsoup3 = BeautifulSoup(result3, 'html.parser')
        data3 = pd.DataFrame()
        te3 = xmlsoup3.findAll("list")
        
        for t3 in te3:
            temp3 = pd.DataFrame(([[t3.rcept_no.string,t3.corp_cls.string,t3.corp_code.string,t3.corp_name.string,t3.isu_dcrs_de.string,
                        t3.isu_dcrs_stle.string,t3.isu_dcrs_stock_knd.string,t3.isu_dcrs_qy.string,t3.isu_dcrs_mstvdv_fval_amount.string,
                                    t3.isu_dcrs_mstvdv_amount.string]]),
                        columns = ["접수번호","법인구분","고유번호","법인명","주식발행 감소일자","발행 감소 형태",
                                   "발행 감소 주식 종류","발행 감소 수량",
                                   "발행 감소 주당 액면가액","발행 감소 주당 가액"])
            data3 = pd.concat([data3,temp3])
            file_nm = "유진기업_"+bsns_year+"년도 증자(감자)현황 배당에 관한 사항.xlsx"
            data3.to_excel(os.path.join(file_path, file_nm),encoding="euc-kr",index=False)
        
    except Exception as e :
            print(traceback.format_exc())
            
elif rep_kind == 2:
    try:
        #자기취득 및 처분현황
        url4 = "https://opendart.fss.or.kr/api/tesstkAcqsDspsSttus.xml?crtfc_key="+api_key+"&corp_code="+com_no+"&bsns_year="+bsns_year+"&reprt_code="+rep_code
        print(url4)
        resultXML4 = urlopen(url4)
        result4 = resultXML4.read()
        xmlsoup4 = BeautifulSoup(result4, 'html.parser')
        data4 = pd.DataFrame()
        te4 = xmlsoup4.findAll("list")


        for t4 in te4:
            temp4 = pd.DataFrame(([[t4.rcept_no.string,t4.corp_cls.string,t4.corp_code.string,t4.corp_name.string,t4.stock_knd.string,
                        t4.acqs_mth1.string,t4.acqs_mth2.string,t4.acqs_mth3.string,t4.bsis_qy.string,
                                    t4.change_qy_acqs.string,t4.change_qy_dsps.string,t4.change_qy_incnr.string,
                                   t4.trmend_qy.string,t4.rm.string]]),
                                 columns = ["접수번호","법인구분","고유번호","법인명","주식 종류", "취득 방법 1", "취득 방법 2" , 
                                            "취득 방법 3", "기초 수량","변동 수량 취득","변동 수량 처분", "변동 수량 소각", "기말 수량", "비고"])
            data4 = pd.concat([data4,temp4])
            file_nm = "유진기업_"+bsns_year+"년도 자기취득 및 처분현황.xlsx"
            data4.to_excel(os.path.join(file_path, file_nm),encoding="euc-kr",index=False)

    except Exception as e :
            print(traceback.format_exc())
        
elif rep_kind ==3:
    try:
        #최대 주주현황
        url5 = "https://opendart.fss.or.kr/api/hyslrSttus.xml?crtfc_key="+api_key+"&corp_code="+com_no+"&bsns_year="+bsns_year+"&reprt_code="+rep_code
        print(url5)
        
        resultXML5 = urlopen(url5)
        result5 = resultXML5.read()
        xmlsoup5 = BeautifulSoup(result5, 'html.parser')
        data5 = pd.DataFrame()
        te5 = xmlsoup5.findAll("list")
        
        
        for t5 in te5:
            temp5 = pd.DataFrame(([[t5.rcept_no.string, t5.corp_cls.string, t5.corp_code.string, t5.corp_name.string, t5.stock_knd.string, t5.rm.string, t5.nm.string, t5.relate.string,
                                    t5.bsis_posesn_stock_co.string, t5.bsis_posesn_stock_qota_rt.string,t5.trmend_posesn_stock_co.string, t5.trmend_posesn_stock_qota_rt.string
                                   ]]),
                                 columns = ["접수번호","법인구분","고유번호","법인명","주식 종류", "비고", "성명" , 
                                            "관계", "기초 소유 주식 수","기초 소유 주식 지분율","기말 소유 주식 수", "기말 소유 주식 지분율"])
            data5 = pd.concat([data5,temp5])
            file_nm = "유진기업_"+bsns_year+"년도 최대 주주현황.xlsx"
            data5.to_excel(os.path.join(file_path, file_nm),encoding="euc-kr",index=False)

        #relate 요소가 없는 리스트가 있지만 데이터프레임 생성하는데 오류는 없기에 오류만 찍고 넘어갑니다.
    except Exception as e :
            print(traceback.format_exc())
            
elif rep_kind == 4:
    try:
        #최대주주 변동 현황
        url6 = "https://opendart.fss.or.kr/api/hyslrChgSttus.xml?crtfc_key="+api_key+"&corp_code="+com_no+"&bsns_year="+bsns_year+"&reprt_code="+rep_code
        print(url6)
        
        resultXML6 = urlopen(url6)
        result6 = resultXML6.read()
        xmlsoup6 = BeautifulSoup(result6, 'html.parser')
        data6 = pd.DataFrame()
        te6 = xmlsoup6.findAll("list")
        
        for t6 in te6:
            temp6 = pd.DataFrame(([[t6.rcept_no.string,t6.corp_cls.string,t6.corp_code.string,t6.corp_name.string,t6.rm.string,
                                    t6.change_on.string,t6.mxmm_shrholdr_nm.string,t6.posesn_stock_co.string,t6.qota_rt.string,t6.change_cause.string
                                   ]]),
                                 columns = ["접수번호","법인구분","고유번호","법인명", "비고", "변동 일", 
                                            "최대 주주 명", "소유 주식 수","지분 율","변동 원인"])
            
            data6 = pd.concat([data6,temp6])
            file_nm = "유진기업_"+bsns_year+"년도 최대주주 변동 현황.xlsx"
            data6.to_excel(os.path.join(file_path, file_nm),encoding="euc-kr",index=False)
            
        
    except Exception as e :
            print(traceback.format_exc())
        
elif rep_kind == 5:
    try:
        url7 = "https://opendart.fss.or.kr/api/mrhlSttus.xml?crtfc_key="+api_key+"&corp_code="+com_no+"&bsns_year="+bsns_year+"&reprt_code="+rep_code
        print(url7)
        
        
        resultXML7 = urlopen(url7)
        result7 = resultXML7.read()
        xmlsoup7 = BeautifulSoup(result7, 'html.parser')
        data7 = pd.DataFrame()
        te7 = xmlsoup7.findAll("list")
        
        for t7 in te7:
            temp7 = pd.DataFrame(([[t7.rcept_no.string,t7.corp_cls.string,t7.corp_code.string,t7.corp_name.string,t7.se.string,
                                    t7.shrholdr_co.string,t7.shrholdr_rate.string,t7.hold_stock_co.string,t7.hold_stock_rate.string
                                   ]]),
                                 columns = ["접수번호","법인구분","고유번호","법인명", "구분", "주주 수", 
                                            "주주 비율", "보유 주식 수"," 보유 주식 비율"])
            
            data7 = pd.concat([data7,temp7])
            file_nm = "유진기업_"+bsns_year+"년도 소액주주현황.xlsx"
            data7.to_excel(os.path.join(file_path, file_nm),encoding="euc-kr",index=False)
    
    except Exception as e :
            print(traceback.format_exc())
        
elif rep_kind == 6:
    try:
        #임원현황
        url8 = "https://opendart.fss.or.kr/api/exctvSttus.xml?crtfc_key="+api_key+"&corp_code="+com_no+"&bsns_year="+bsns_year+"&reprt_code="+rep_code 
        print(url8)
        resultXML8 = urlopen(url8)
        result8 = resultXML8.read()
        xmlsoup8 = BeautifulSoup(result8, 'html.parser')
        data8 = pd.DataFrame()
        te8 = xmlsoup8.findAll("list")
        
        for t8 in te8:
            temp8 = pd.DataFrame(([[t8.rcept_no.string,t8.corp_cls.string,t8.corp_code.string,t8.corp_name.string,t8.nm.string,t8.sexdstn.string,t8.birth_ym.string,t8.ofcps.string,
                                    t8.rgist_exctv_at.string,t8.fte_at.string,t8.chrg_job.string,t8.main_career.string,t8.mxmm_shrholdr_relate.string,t8.hffc_pd.string,t8.tenure_end_on.string
                                   ]]),
                                 columns = ["접수번호","법인구분","고유번호","법인명","성명","성별","출생 년월","직위","등기 임원 여부",
                                            "상근 여부","담당 업무","주요 경력","최대 주주 관계","재직 기간","임기 만료 일"])
            data8 = pd.concat([data8, temp8])
            file_nm = "유진기업_"+bsns_year+"년도 임원현황.xlsx"
            data8.to_excel(os.path.join(file_path, file_nm),encoding="euc-kr",index=False)
        
        #직원현황
        url9 = "https://opendart.fss.or.kr/api/empSttus.xml?crtfc_key="+api_key+"&corp_code="+com_no+"&bsns_year="+bsns_year+"&reprt_code="+rep_code
        print(url9)
        resultXML9 = urlopen(url9)
        result9 = resultXML9.read()
        xmlsoup9 = BeautifulSoup(result9, 'html.parser')
        data9 = pd.DataFrame()
        te9 = xmlsoup9.findAll("list")
        
        for t9 in te9:
            temp9 = pd.DataFrame(([[t9.rcept_no.string,t9.corp_cls.string,t9.corp_code.string,t9.corp_name.string,t9.rm.string,t9.sexdstn.string,t9.fo_bbm.string,
                                    t9.reform_bfe_emp_co_rgllbr.string,t9.reform_bfe_emp_co_cnttk.string,t9.reform_bfe_emp_co_etc.string,t9.rgllbr_co.string,t9.rgllbr_abacpt_labrr_co.string,
                                    t9.cnttk_co.string,t9.cnttk_abacpt_labrr_co.string,t9.sm.string,t9.avrg_cnwk_sdytrn.string,
                                    t9.fyer_salary_totamt.string,t9.jan_salary_am.string
                                   ]]),
                                 columns = ["접수번호","법인구분","고유번호","법인명","비고","성별","사업부문","개정 전 직원수 정규직","개정 전 직원수 계약직",
                                            "개정 전 직원 수 기타","정규직 수","정규직 단시간 근로자 수 ","계약직 수","계약직 단시간 근로자 수",
                                            "합계","평균 근속 연수","연간 급여 총액","1인평균 급여액"])
            data9 = pd.concat([data9, temp9])
            file_nm = "유진기업_"+bsns_year+"년도 직원현황.xlsx"
            data9.to_excel(os.path.join(file_path, file_nm),encoding="euc-kr",index=False)
        
        
    except Exception as e :
            print(traceback.format_exc())
        
elif rep_kind == 7:
    try:
        #타법인출자현황
        url10 = "https://opendart.fss.or.kr/api/otrCprInvstmntSttus.xml?crtfc_key="+api_key+"&corp_code="+com_no+"&bsns_year="+bsns_year+"&reprt_code="+rep_code
        print(url10)
        
        resultXML10 = urlopen(url10)
        result10 = resultXML10.read()
        xmlsoup10 = BeautifulSoup(result10, 'html.parser')
        data10 = pd.DataFrame()
        te10 = xmlsoup10.findAll("list")
        
        for t10 in te10:
            temp10 = pd.DataFrame(([[t10.rcept_no.string,t10.corp_cls.string,t10.corp_code.string,t10.corp_name.string,t10.inv_prm.string,t10.frst_acqs_de.string,
                                     t10.invstmnt_purps.string,t10.frst_acqs_amount.string,
                                     t10.bsis_blce_qy.string,t10.bsis_blce_qota_rt.string,t10.bsis_blce_acntbk_amount.string,t10.incrs_dcrs_acqs_dsps_qy.string,t10.incrs_dcrs_acqs_dsps_amount.string,
                                     t10.incrs_dcrs_evl_lstmn.string,t10.trmend_blce_qy.string,t10.trmend_blce_qota_rt.string,t10.trmend_blce_acntbk_amount.string,
                                     t10.recent_bsns_year_fnnr_sttus_tot_assets.string,t10.recent_bsns_year_fnnr_sttus_thstrm_ntpf.string
                                   ]]),
                                 columns = ["접수번호","법인구분","고유번호","회사명","법인명","최초 취득 일자","출자 목적","최초 취득 금액","기초 잔액 수량",
                                            "기초 잔액 지분 율","기초 잔액 장부 가액","증가 감소 취득 처분 수량","증가 감소 취득 처분 금액","증가 감소 평가 손액","기말 잔액 수량",
                                            "기말 잔액 지분 율","기말 잔액 장부 가액","최근 사업 연도 재무 현황 총 자산","최근 사업 연도 재무 현황 당기 순이익"])
            
            data10 = pd.concat([data10,temp10])
            file_nm = "유진기업_"+bsns_year+"년도 타법인출자현황.xlsx"
            data10.to_excel(os.path.join(file_path, file_nm),encoding="euc-kr",index=False)
            
    except Exception as e :
            print(traceback.format_exc())
        

        
elif rep_kind == 8:
    try:
        #단일회사 주요 계정
        url11 = "https://opendart.fss.or.kr/api/fnlttSinglAcnt.xml?crtfc_key="+api_key+"&corp_code="+com_no+"&bsns_year="+bsns_year+"&reprt_code="+rep_code
        print(url11)
        
        resultXML11 = urlopen(url11)
        result11 = resultXML11.read()
        xmlsoup11 = BeautifulSoup(result11, 'html.parser')
        data11 = pd.DataFrame()
        te11 = xmlsoup11.findAll("list")
        
        for t11 in te11:
            temp11 = pd.DataFrame(([[t11.rcept_no.string,t11.corp_code.string,t11.stock_code.string,t11.account_nm.string,t11.fs_div.string,t11.fs_nm.string,t11.sj_div.string,
                                     t11.sj_nm.string,t11.thstrm_nm.string,t11.thstrm_dt.string,t11.thstrm_amount.string,
                                     t11.frmtrm_nm.string,t11.frmtrm_dt.string,t11.frmtrm_amount.string,t11.bfefrmtrm_nm.string,
                                     t11.bfefrmtrm_dt.string,t11.bfefrmtrm_amount.string,t11.ord.string
                                   ]]),
                                 columns = ["접수번호","고유번호","종목 코드","계정명","개별/연결구분","개별/연결명",
                                            "재무제표구분","재무제표명","당기명","당기일자","당기금액","전기명",
                                            "전기일자","전기금액","전전기명","전전기일자","전전기금액","계정과목 정렬순서"])
            
            data11 = pd.concat([data11,temp11])
            file_nm = "유진기업_"+bsns_year+"년도 단일회사 주요 계정.xlsx"
            data11.to_excel(os.path.join(file_path, file_nm),encoding="euc-kr",index=False)

    except Exception as e :
            print(traceback.format_exc())
        
elif rep_kind == 9:
    try:
        #다중회사 주요 계정
        com_a = str(input("7. 추가로 선택할 기업 1개의 고유번호를 입력하시오 : "))
        url12 = "https://opendart.fss.or.kr/api/fnlttMultiAcnt.xml?crtfc_key="+api_key+"&corp_code="+com_a+","+com_no+"&bsns_year="+bsns_year+"&reprt_code="+rep_code
        print(url12)
        
        resultXML12 = urlopen(url12)
        result12 = resultXML12.read()
        xmlsoup12 = BeautifulSoup(result12, 'html.parser')
        data12 = pd.DataFrame()
        te12 = xmlsoup12.findAll("list")
        
        for t12 in te12:
            temp12 = pd.DataFrame(([[t12.rcept_no.string,t12.corp_code.string,t12.stock_code.string,t12.account_nm.string,t12.fs_div.string,t12.fs_nm.string,t12.sj_div.string,
                                     t12.sj_nm.string,t12.thstrm_nm.string,t12.thstrm_dt.string,t12.thstrm_amount.string,
                                     t12.frmtrm_nm.string,t12.frmtrm_dt.string,t12.frmtrm_amount.string,t12.bfefrmtrm_nm.string,
                                     t12.bfefrmtrm_dt.string,t12.bfefrmtrm_amount.string,t12.ord.string
                                   ]]),
                                 columns = ["접수번호","고유번호","종목 코드","계정명","개별/연결구분","개별/연결명",
                                            "재무제표구분","재무제표명","당기명","당기일자","당기금액","전기명",
                                            "전기일자","전기금액","전전기명","전전기일자","전전기금액","계정과목 정렬순서"])
            
            data12 = pd.concat([data12,temp12])
            file_nm = "유진기업+"+bsns_year+"년도 다중회사 주요 계정.xlsx"
            data12.to_excel(os.path.join(file_path, file_nm),encoding="euc-kr",index=False)
    except Exception as e :
            print(traceback.format_exc())
else:
    print("레포트 종류 번호를 선택해주세요")