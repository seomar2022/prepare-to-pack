import pandas as pd
import os
from datetime import datetime
import csv

####설정폴더에서 경로찾기
def search_path(header_name):
    try:
        # CSV 파일 열기
        with open("settings\path.csv", mode='r', encoding='utf-8') as file:
            reader = csv.reader(file)
            
            # 데이터 검색
            for row in reader:
                if row[0].strip() == header_name:
                    return row[1].strip()
            
            print(f"헤더 '{header_name}'을(를) 찾을 수 없습니다.")
            return ""
                
    except FileNotFoundError:
        print(f"설정 파일을 찾을 수 없습니다")
        return ""
    except Exception as e:
        print(f"설정 파일을 읽는 중 오류가 발생했습니다: {e}")
        return ""


####이름 일부를 검색해서 파일찾기
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


####헤더 이름을 입력하면 몇 열인지 return
def find_header_index(file_path, header_name):
    """
    CSV 파일에서 특정 헤더 이름의 인덱스를 찾습니다.

    Args:
        file_path (str): CSV 파일의 경로
        header_name (str): 찾고자 하는 헤더 이름

    Returns:
        int: 헤더의 인덱스 (0부터 시작), 존재하지 않을 경우 -1
    """
    try:
        with open(file_path, mode='r', encoding='utf-8') as file:
            reader = csv.reader(file)
            headers = next(reader)  # 첫 번째 행에서 헤더를 가져옴
            
            if header_name in headers:
                index = headers.index(header_name)
                return index
            else:
                print(f"'{header_name}' 헤더를 찾을 수 없습니다.")
                return -1
    except FileNotFoundError:
        print(f"CSV 파일을 찾을 수 없습니다: {file_path}")
        return -1
    except Exception as e:
        print(f"파일을 처리하는 중 오류가 발생했습니다: {e}")
        return -1

