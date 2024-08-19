from pypdf import PdfWriter #pip install pypdf #pdf병합기능 쓰기 위해.
from datetime import datetime #병합된 pdf이름에 오늘 날짜 쓰기 위해.
import pandas as pd #pip install pandas openpyxl #엑셀의 데이터를 읽어오기 위해.
import re

#문자열에서 중량 정보를 추출
def extract_weight(product_data):
    if not isinstance(product_data, str):
        return None
    product_data = product_data.lower().replace('ml', 'g')
    
    match = re.search(r'(\d+(\.\d+)?)\s*(kg|g)', product_data, re.IGNORECASE)
    
    if match:
        weight = float(match.group(1))
        unit = match.group(3).lower()
        
        # 단위가 g일 경우 kg으로 변환
        if unit == 'g':
            weight = weight / 1000
        
        return weight
    else:
        return None

#조건에 따라 중량 정보를 선택
def get_final_weight(row):
    if pd.notna(row['중량']) == False :#중량열에 데이터 없음
        #->스마트 스토어나 톡스토어 주문건임. 상품명에서 중량 추출해야함.  
        weight_from_name_column = extract_weight(row['상품명(한국어 쇼핑몰)'])
        if weight_from_name_column is not None:
            print(f"상품명에서 중량 사용: {weight_from_name_column}")
        return weight_from_name_column
    
    else: #중량열에 데이터 있음
        #->카페24주문 건임(몇몇 연동된 상품 제외). 상품옵션열의 데이터에 중량 정보기 중량열의 데이터가 다를 때만 상품옵션 데이터 쓰기. 
        weight_from_option_column = extract_weight(row['상품옵션'])
        if weight_from_option_column is None: #상품옵션에 중량 정보가 없을경우
            print(f"중량에서 중량 사용: {weight_from_option_column}")
            return row['중량']
        else: #상품옵션에 중량 정보가 있을 경우
            print(f"상품옵션에서 중량 사용: {weight_from_option_column}")
            return weight_from_option_column



####엑셀 파일 읽어오기
def ready_to_convert(order_list_pd): 
    # 상품코드 열의 데이터를 문자열로 변환하고 NaN 값을 빈 문자열로 대체
    codes = order_list_pd['상품코드'].astype(str).fillna("").tolist()

    #카페24상품코드와 네이버 상품코드를 매핑한 엑셀파일 읽어오기
    product_code_mapping = pd.read_excel(r"resources\product_code_mapping.xlsx", engine='openpyxl')

    product_code_mapping['naver_code'] = product_code_mapping['naver_code'].astype(str).str.strip().str.replace('-', '')
    product_code_mapping['kakao_code'] = product_code_mapping['kakao_code'].astype(str).str.strip().str.replace('-', '')
    product_code_mapping['cafe24_code'] = product_code_mapping['cafe24_code'].astype(str).str.strip()
    
    ####상품코드를 카페24의 코드로 통일
    def convert_to_cafe24(code, column):
        # 'column' 열에 'code'가 있는 행 ex) naver_code열에서 9708250509가 있는 행
        result = product_code_mapping.query(f"{column} == @code")
        return result['cafe24_code'].iat[0]

    converted_codes = []
    for code in codes:
        if code.startswith("P00") : #카페24
            converted_codes.append(code)
        elif code.startswith("9") or code.startswith("1") : #네이버
            #상품코드 맵핑된 엑셀파일에서 네이버 상품코드에 해당하는 카페24상품코드 가져오기
            result = convert_to_cafe24(code, "naver_code")
            converted_codes.append(result)
        elif code.startswith("3") : #카카오
            result = convert_to_cafe24(code, "kakao_code")
            converted_codes.append(result)
            
    return converted_codes

#### PDF 파일 병합
def merge_pdf(converted_codes):
    merge_pdf = PdfWriter()
    not_found_files = {}

    #convert된 코드에 해당하는 설명지를 찾아 append
    for converted_code in converted_codes:
        try:
            merge_pdf.append(f"resources\\product_instruction\\{converted_code}.pdf")
        except FileNotFoundError:
            not_found_files[converted_code] = ''

    # 현재 날짜 가져오기
    now = datetime.now().strftime("%m.%d.%a") #월.일.요일

    merge_pdf.write(f"result\\{now}_product_instruction.pdf")
    merge_pdf.close()
    return not_found_files

####설명지 없는 상품코드와 상품명 알려주기
def report_result(order_list_pd, not_found_files):
    #설명지 없는 상품의 코드를 전채널주문리스트에서 찾고, 해당 상품의 이름을 가져와서 딕셔너리의 값으로 넣기
    #매크로 돌리기 전의 열이름은 '상품명(한국어 쇼핑몰)' 돌린 후는 '상품명'이라서 상품명이 포함된 열을 지정
    product_name_col = [col for col in order_list_pd.columns if "상품명" in col]
    for key in not_found_files:
        key_in_order_list = order_list_pd.query(f"상품코드 == @key")
        not_found_files[key] = key_in_order_list[product_name_col[0]].iat[0]
        
    
    ####pyautogui로 프로그램 실행 결과 알려주기
    if len(not_found_files) == 0:
        alert_msg = "모든 상품의 설명지를 찾았습니다!"
    else:
        #csv파일로 저장
        now = datetime.now().strftime("%m.%d.%a") #월.일.요일
        converted_codes_df = pd.DataFrame(list(not_found_files.items()), columns=['상품코드', '상품명'])
        converted_codes_df.to_csv(f"result\\{now}_not_found_files.csv", index=False, encoding='utf-8-sig')
        alert_msg=f"{len(not_found_files)}개의 설명지를 찾지 못했습니다"
        print(alert_msg)
