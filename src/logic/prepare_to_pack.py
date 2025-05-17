from .module import (
    search_path,
    find_path_by_partial_name,
    get_column_from_csv,
    run_macro,
)
from .before_packing import (
    get_final_weight,
    convert_to_cafe24_product_code,
    merge_product_instructions,
    report_missing_instructions,
    assign_gift,
    match_to_cafe24_example,
)
import pandas as pd
import os
import webbrowser
import time  # GUI에서 멀티스레드 사용하기 위해
from datetime import datetime  # 폴더이름에 현재 날짜 넣기 위해

from pathlib import Path
import sys

sys.path.append(str(Path(__file__).resolve().parent.parent.parent))
from settings.column_mapping import KOR_TO_ENG_COLUMN_MAP


def prepare_to_pack(log_set_callback, log_get_callback):
    try:
        sleep_time = 0.1
        log_set_callback("🍰🍰🍰시작! 프로그램 실행🍰🍰🍰")

        ### 완성된 파일들을 넣어둘 폴더 만들기
        output_folder = "result_" + datetime.now().strftime(
            "%a.%H.%M.%S"
        )  # 요일.시.분.초
        os.makedirs(output_folder)

        ########################################## Split file from download folder ##########################################
        # Get the folder path from setting\path.csv where the raw file is located
        download_from_internet_path = search_path("download_from_internet")

        # Find the file downloaded from Cafe24
        # File name format downloaded from Cafe24: lalapetmall_today's_date_serialnumber_serialnumber
        download_from_cafe24_path = find_path_by_partial_name(
            download_from_internet_path,
            "lalapetmall_" + datetime.today().strftime("%Y%m%d") + "_",
        )
        df_raw_data = pd.read_csv(download_from_cafe24_path)

        # log
        log_set_callback(log_get_callback() + "\n다운로드 받은 파일 검색")
        time.sleep(sleep_time)

        ### Split into two files
        # Order list file
        # order_list_path = rf"{output_folder}\order_list.xlsx"
        order_list_header_list = get_column_from_csv(
            r"settings\header.csv", "order_list_header"
        )
        df_order_list = df_raw_data[order_list_header_list].rename(
            columns=KOR_TO_ENG_COLUMN_MAP
        )
        # df_order_list.to_excel(order_list_path, index=False)

        # Hanjin file
        hanjin_path = rf"{output_folder}\hanjin_file.xlsx"
        hanjin_header_list = get_column_from_csv(
            r"settings\header.csv", "hanjin_header"
        )
        df_hanjin_list = df_raw_data[hanjin_header_list]
        df_hanjin_list.to_excel(hanjin_path, index=False)

        # log
        log_set_callback(log_get_callback() + "\n헤더명에 따라 두 개의 파일로 분리")
        time.sleep(sleep_time)

        ########################################## Print out product instruction ##########################################
        converted_cafe24_codes = convert_to_cafe24_product_code(df_order_list)
        not_found_files = merge_product_instructions(
            output_folder, converted_cafe24_codes
        )
        report_missing_instructions(output_folder, df_order_list, not_found_files)

        # log
        log_set_callback(log_get_callback() + "\n상품 설명지 병합")
        time.sleep(sleep_time)

        ########################################## Data Transformation ##########################################
        #### Update weight column
        df_order_list["weight"] = df_order_list.apply(get_final_weight, axis=1)
        # df_order_list.to_excel(order_list_path, index=False)
        # print(df_order_list["weight"])

        # log
        log_set_callback(log_get_callback() + "\n주문리스트의 중량 정보 입력")
        time.sleep(sleep_time)

        ####매크로 실행(포장할 때 참고할 주문리스트 만들기 위해)
        #  run_macro("전채널주문리스트", order_list_path)
        # log
        log_set_callback(
            log_get_callback() + "\n전채널주문리스트 매크로 실행\n주문리스트 파일 작성"
        )
        time.sleep(sleep_time)

        ### Generate serial numbers for unique order_numbers; set blank for subsequent duplicates
        # Step 1: Mark True for the first occurrence of each order_number
        df_order_list["serial_number"] = ~df_order_list["order_number"].duplicated()

        # Step 2: Assign a running number to the first occurrences. cumsum: cumulative sum
        df_order_list["serial_number"] = df_order_list["serial_number"].cumsum()

        # Step 3. Convert column to object (string-friendly) type
        df_order_list["serial_number"] = df_order_list["serial_number"].astype(object)

        # Step 4: Replace non-first rows (False values) with blank
        df_order_list.loc[
            df_order_list["order_number"].duplicated(), "serial_number"
        ] = ""

        ### Determine gift type
        df_order_list["gift"] = df_order_list.apply(assign_gift, axis=1)

        ####카페24 양식에 맞게 수정한 파일 만들기
        match_to_cafe24_example(output_folder, hanjin_path)
        # log
        log_set_callback(log_get_callback() + "\n송장등록을 위한 카페24양식 파일 작성")
        time.sleep(sleep_time)

        ####매크로 실행(기존 파일을 한진택배 복수내품 양식에 맞게 변경하기 위해)
        run_macro("ProcessMultipleItems", hanjin_path)
        os.rename(hanjin_path, rf"{output_folder}\upload_to_hanjin.xlsx")
        # log
        log_set_callback(
            log_get_callback()
            + "\nProcessMultipleItems 매크로 실행\n한진 사이트에 올릴 파일 작성"
        )
        time.sleep(sleep_time)

        ####한진택배 사이트 열기
        webbrowser.open("https://focus.hanjin.com/login")

        ####result 폴더 열기
        os.startfile(f"{output_folder}")
        print("실행 완료.")
        # log
        log_set_callback(log_get_callback() + "\n🍰🍰🍰끝! 실행 완료🍰🍰🍰")
        time.sleep(sleep_time)
    except Exception as e:
        log_set_callback(log_get_callback() + f"\n❗❗❗오류 발생: {e}")
