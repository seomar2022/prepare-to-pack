import pandas as pd
import os
from datetime import datetime


def match_cafe24_with_hanjin(cafe24_file, hanjin_file, output_file):
    # 카페24 주문 데이터 불러오기
    cafe24_data = pd.read_csv(cafe24_file, encoding='utf-8')
    print("카페24 데이터 불러오기 완료.")
    
    # 한진택배 배송 데이터 불러오기
    hanjin_data = pd.read_excel(hanjin_file, engine='openpyxl')
    print("한진택배 데이터 불러오기 완료.")
    
    # 데이터 매칭 (예: 주문번호를 기준으로)
    matched_data = pd.merge(cafe24_data, hanjin_data[['주문번호', '운송장번호']], how='left', on='주문번호')
    print("데이터 매칭 완료.")
    
    # D1 셀에 "수량" 추가 
    matched_data.insert(3, '수량', '') 

    # 결과 저장
    matched_data.to_csv(output_file, index=False, encoding='utf-8-sig')
    print("결과 저장 완료")

############한진택배 파일 찾기###############
def find_file_by_partial_name(directory, partial_name):
    # 지정된 디렉토리의 파일 목록을 가져옴
    files = os.listdir(directory)
    
    # 부분 문자열이 파일 이름에 포함된 파일 목록 생성
    matching_files = [file for file in files if partial_name in file]
    
    # 매칭된 파일이 없으면 None 반환
    if not matching_files:
        return None
    
    # 가장 최근에 수정된 파일 찾기
    most_recent_file = max(matching_files, key=lambda f: os.path.getmtime(os.path.join(directory, f)))
    
    # 가장 최근에 수정된 파일의 전체 경로 반환
    return os.path.join(directory, most_recent_file)


# 파일 경로 설정
directory = r"C:\Users\User\Downloads"  # 파일이 있는 디렉토리
partial_name = "출력자료등록_원본_" + datetime.today().strftime('%Y%m%d')  # 검색할 파일 이름의 부분 문자열


# 파일 검색
file_path = find_file_by_partial_name(directory, partial_name)

if file_path:
    print(f"파일을 찾았습니다: {file_path}")
    hanjin_file = file_path
else:
    print("파일을 찾을 수 없습니다.")
########################################################################

# 파일 경로 설정
cafe24_file = r'result\excel_sample_old.csv'
output_file = r'result\excel_sample_old.csv'

# 매칭 실행
match_cafe24_with_hanjin(cafe24_file, hanjin_file, output_file)