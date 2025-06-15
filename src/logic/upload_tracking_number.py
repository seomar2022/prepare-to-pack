import pandas as pd
from datetime import datetime
from .module import search_path, find_path_by_partial_name
import sys
import os


def match_cafe24_with_hanjin(cafe24_file, hanjin_file, output_file):
    # 카페24 주문 데이터 불러오기
    cafe24_data = pd.read_csv(cafe24_file, encoding="utf-8")

    # 한진택배 배송 데이터 불러오기
    hanjin_data = pd.read_excel(hanjin_file, engine="openpyxl")

    # 데이터 매칭 (예: 주문번호를 기준으로)
    matched_data = pd.merge(
        cafe24_data, hanjin_data[["주문번호", "운송장번호"]], how="left", on="주문번호"
    )
    print("데이터 매칭 완료.")

    # D1 셀에 "수량" 추가
    matched_data.insert(3, "수량", "")

    # 결과 저장
    matched_data.to_csv(output_file, index=False, encoding="utf-8-sig")
    print("결과 저장 완료")


def upload_tracking_number():
    # 검색할 파일 이름의 부분 문자열
    partial_name = "출력자료등록_원본_" + datetime.today().strftime("%Y%m%d")
    # 파일 검색
    file_path = find_path_by_partial_name(
        search_path("download_from_internet"), partial_name
    )

    if file_path:
        print(f"파일을 찾았습니다: {file_path}")
        hanjin_file = file_path
    else:
        print("파일을 찾을 수 없습니다.")

    # os.path.dirname(os.path.abspath(__file__) ->현재 실행 중인 스크립트 파일이 위치한 디렉토리 경로
    result_directory = find_path_by_partial_name(
        os.path.dirname(sys.executable), "result_"
    )
    print("os.path.dirname(sys.executable):", os.path.dirname(sys.executable))
    # os.path.dirname(os.path.abspath(__file__)): C:\Users\User\AppData\Local\Temp\_MEI95402
    print("result_directory:", result_directory)

    # 파일 경로 설정
    cafe24_file = rf"{result_directory}\excel_sample_old.csv"
    output_file = rf"{result_directory}\excel_sample_old.csv"

    # 매칭 실행
    match_cafe24_with_hanjin(cafe24_file, hanjin_file, output_file)


def on_upload_tracking_number_button_click(log_set_callback):
    try:
        upload_tracking_number()
        log_set_callback("실행 완료\n카페24 엑셀 일괄배송 처리란에 업로드해 주세요.")
    except Exception as e:
        log_set_callback(f"⚠️오류 발생: {e}")
