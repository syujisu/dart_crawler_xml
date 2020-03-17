import pandas as pd


def DART_data_to_excel(df, corpx, yearx, reportx):
    if "11011" == reportx:
        option = 'year'
    else:
        option = 'quarter'
    df = xml_work(read_file_name(corpx, yearx, reportx), df, option)
    return df


def read_file_name(corpx, yearx, reportx):
    xml_file_name = "./data/majorstock_" + corpx + "_" + yearx + "_" + reportx + ".xml"
    return xml_file_name


def excel_work_year(treex, dfx):
    corp_list = treex.findall('./list')
    for i in corp_list:
        if "OFS" == i.find('./fs_div').text:
            account_nm = i.find('./account_nm').text.strip()
            thstrm_dt = i.find('./thstrm_dt').text.strip()
            frmtrm_dt = i.find('./frmtrm_dt').text.strip()
            bfefrmtrm_dt = i.find('./bfefrmtrm_dt').text.strip()

            if ("매출액" == account_nm) or ("영업이익" == account_nm) or ("법인세차감전 순이익" == account_nm) or (
                    "당기순이익" == account_nm):
                thstrm_dt = thstrm_dt[13:23]
                frmtrm_dt = frmtrm_dt[13:23]
                bfefrmtrm_dt = bfefrmtrm_dt[13:23]
            else:
                thstrm_dt = thstrm_dt[:10]
                frmtrm_dt = frmtrm_dt[:10]
                bfefrmtrm_dt = bfefrmtrm_dt[:10]

            thstrm_amount = int(i.find('./thstrm_amount').text.strip().replace(',', ''))
            frmtrm_amount = int(i.find('./frmtrm_amount').text.strip().replace(',', ''))
            bfefrmtrm_amount = int(i.find('./bfefrmtrm_amount').text.strip().replace(',', ''))
            print(account_nm)
            dfx.loc[account_nm, thstrm_dt] = thstrm_amount
            dfx.loc[account_nm, frmtrm_dt] = frmtrm_amount
            dfx.loc[account_nm, bfefrmtrm_dt] = bfefrmtrm_amount
    return dfx


def excel_work_quarter(treex, dfx):
    corp_list = treex.findall('./list')

    for i in corp_list:
        if "OFS" == i.find('./fs_div').text:
            account_nm = i.find('./account_nm').text.strip()
            thstrm_dt = i.find('./thstrm_dt').text.strip()

            if ("매출액" == account_nm) or ("영업이익" == account_nm) or ("법인세차감전 순이익" == account_nm) or (
                    "당기순이익" == account_nm):
                thstrm_dt = thstrm_dt[13:23]
            else:
                thstrm_dt = thstrm_dt[:10]

            thstrm_amount = int(i.find('./thstrm_amount').text.strip().replace(',', ''))
            print(account_nm)
            dfx.loc[account_nm, thstrm_dt] = thstrm_amount
    return dfx


def xml_work(xml_work_filex, dfx, option):
    print(xml_work_filex)
    import xml.etree.ElementTree as elemTree
    tree = elemTree.parse(xml_work_filex)
    response_message = tree.find('./status').text
    error_message = '초기화메시지입니다'
    if "000" == response_message:
        error_message = '정상입니다'
        if "year" == option:
            dfx = excel_work_year(tree, dfx)
        elif "quarter" == option:
            dfx = excel_work_quarter(tree, dfx)
        else:
            error_message = '옵션이 잘못되어 처리되지 않았습니다.'
    elif "010" == response_message:
        error_message = '등록되지 않은 키입니다.'
    elif "011" == response_message:
        error_message = '사용할 수 없는 키입니다.'
    elif "013" == response_message:
        error_message = '조회된 데이타가 없습니다'
    elif "020" == response_message:
        error_message = '요청 제한을 초과하였습니다.'
    elif "100" == response_message:
        error_message = '필드의 부적절한 값입니다.'
    elif "800" == response_message:
        error_message = '원활한 공시서비스를 위하여 오픈API 서비스가 중지 중입니다.'
    elif "900" == response_message:
        error_message = '정의되지 않은 오류가 발생하였습니다.'
    print(xml_work_filex + "파일 처리 내용 : " + error_message)
    return dfx


def init_def(corpx):
    left_index = ['유동자산', '비유동자산', '자산총계', '유동부채', '비유동부채', '부채총계', '자본금', '자본총계',
                  '구분선1',
                  '현금 및 현금성자산', '당좌자산', '외상매출금', '재고자산', '유형자산', '토지', '건설중인자산', '무형자산',
                  '외상매입금', '단기차입부채', '유동성장기차입부채', '장기차입부채', '차입금 계', '매출채권', '매입채무',
                  '구분선2',
                  '매출액', '매출원가', '매출총이익', '판관비', '영업이익', 'EBITDA', '기타수익', '기타비용', '금융수익',
                  '금융비용', '법인세차감전 순이익', '법인세비용', '당기순이익',
                  '구분선3',
                  '급여(상여 포함)', '퇴직급여', '복리후생비', '경상개발비(연구비)', '이자비용', '이자수익',
                  '감가상각비(원가)', '감가상각비(판관)'
                  ]
    top_label = ['2015.12.31',
                 '2016.12.31',
                 '2017.12.31',
                 '2018.03.31',
                 '2018.06.30',
                 '2018.09.30',
                 '2018.12.31',
                 '2019.03.31',
                 '2019.06.30',
                 '2019.09.30',
                 '2019.12.31',
                 ]
    df = pd.DataFrame(0, index=left_index, columns=top_label)
    df.columns.name = '<재무상태표>-' + corpx
    return df


corp_list = ["00184667","00117337","00163266", "00131054", "00165149"]
report_list = ["11011", "11012", "11013", "11014"]
year_list = ["2016", "2018","2019"]
df_list = []

try:
    import os
    if os.path.isdir("./output"):
        pass
    else:
        os.mkdir("output")

    for corpx in corp_list:
        df = init_def(corpx)
        for yearx in year_list:
            for reportx in report_list:
                df = DART_data_to_excel(df, corpx, yearx, reportx)
        df.to_csv("./output/" + corpx + ".csv", encoding='cp949')
        df_list.append(df)

    writer = pd.ExcelWriter('./output/eugeneGroup.xlsx', engine='xlsxwriter')
    indexx = 0
    for dfx in df_list:
        dfx.to_excel(writer, sheet_name=corp_list[indexx])
        indexx += 1
    writer.save()

except Exception as ex:
    print(f'Error Occurred : {ex}')

'''
# 회사 고유번호
00184667 유진기업
00131054 유진증권
00117337 동양
00163266 한일합섬 *
00165149 유진저축은행 *
*표시한 기업은 비상장

# 보고서 항목
11011 사업보고서(년)
11012 반기보고서
11013 1분기보고서
11014 3분기보고서
'''
