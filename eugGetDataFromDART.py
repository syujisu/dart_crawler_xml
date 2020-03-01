def xml_file_save(saveFileName, corp_list_url):
    file = open(saveFileName, "w", encoding='utf8')
    file.write(corp_list_url)
    file.close()

def get_data_from_DART(corp_code, reprt_code, bsns_year):
    from urllib.request import urlopen
    api_key = "??"
    url = "https://opendart.fss.or.kr/api/fnlttSinglAcnt.xml?crtfc_key="+api_key+\
      "&corp_code="+corp_code+"&bsns_year="+bsns_year+"&reprt_code="+reprt_code
    corp_list_url = urlopen(url).read().decode('utf8')
    saveFileName = "./data/majorstock_" + corp_code + "_" + bsns_year + "_" + reprt_code + ".xml"
    xml_file_save(saveFileName, corp_list_url)


corp_list = ["00184667","00117337","00163266", "00131054", "00165149"]
report_list = ["11011","11012","11013","11014"]
year_list = ["2016", "2018","2019"]

try:
    import os
    if os.path.isdir("./data"):
        pass
    else:
        os.mkdir("data")

    for corpx in corp_list:
        for reportx in report_list:
            for yearx in year_list:
                get_data_from_DART(corpx, reportx, yearx)

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

