from .module import (
    search_path,
    find_path_by_partial_name,
    get_column_from_csv,
    )
from .before_packing import (
    get_adjusted_unit_weight,
    convert_to_cafe24_product_code,
    merge_product_instructions,
    report_missing_instructions,
    assign_gift,
    determine_box_size,
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

        ### Output folder
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

        ### Split into three files
        # Order list file
        order_list_path = rf"{output_folder}\order_list.xlsx"
        order_list_header_list = get_column_from_csv(
            r"settings\header.csv", "order_list_header"
        )
        df_order_list = df_raw_data[order_list_header_list].rename(
            columns=KOR_TO_ENG_COLUMN_MAP
        )

        # Hanjin file
        hanjin_path = rf"{output_folder}\upload_to_hanjin.xlsx"
        hanjin_header_list = get_column_from_csv(
            r"settings\header.csv", "hanjin_header"
        )
        df_hanjin_list = df_raw_data[hanjin_header_list].rename(
            columns=KOR_TO_ENG_COLUMN_MAP
        )

        # Cafe24 upload file
        cafe24_upload_path = rf"{output_folder}\excel_sample_old.csv"
        cafe24_upload_header_list = get_column_from_csv(
            r"settings\header.csv", "cafe24_upload"
        )
        df_cafe24_upload = df_raw_data[cafe24_upload_header_list]
        df_cafe24_upload.to_csv(
            cafe24_upload_path,
            index=False,
            encoding="utf-8-sig",
        )

        # log
        log_set_callback(log_get_callback() + "\n헤더명에 따라 세 개의 파일로 분리")
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

        ########################################## Data Transformation(Determine box size) ##########################################
        #### Adjust weight data
        df_order_list["unit_weight"] = df_order_list.apply(
            get_adjusted_unit_weight, axis=1
        )

        df_order_list["item_total_weight"] = (
            df_order_list["unit_weight"] * df_order_list["quantity"]
        )

        total_weight_by_order = df_order_list.groupby("order_number")[
            "item_total_weight"
        ].sum()
        df_order_list["total_weight_by_order"] = df_order_list["order_number"].map(
            total_weight_by_order
        )

        ### Determine box size
        df_order_list["box_size"] = df_order_list.apply(determine_box_size, axis=1)
        print(
            df_order_list[
                [
                    "order_number",
                    "quantity",
                    "unit_weight",
                    "item_total_weight",
                    "total_weight_by_order",
                    "box_size",
                ]
            ]
        )

        # log
        log_set_callback(log_get_callback() + "\n주문리스트의 박스 정보 입력")
        time.sleep(sleep_time)

        ########################################## Data Transformation ##########################################
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

        ########################################## ##########################################
        ####매크로 실행(기존 파일을 한진택배 복수내품 양식에 맞게 변경하기 위해)
        # run_macro("ProcessMultipleItems", hanjin_path)
        # os.rename(hanjin_path, rf"{output_folder}\upload_to_hanjin.xlsx")
        df_hanjin_list.to_excel(hanjin_path, index=False)

        # log
        log_set_callback(log_get_callback() + "\n한진 사이트에 업로드할 파일 작성")
        time.sleep(sleep_time)

        ########################################## ##########################################
        df_order_list.to_excel(order_list_path, index=False)

        ####한진택배 사이트 열기
        webbrowser.open("https://focus.hanjin.com/login")

        ####result 폴더 열기
        os.startfile(f"{output_folder}")
        # log
        log_set_callback(log_get_callback() + "\n🍰🍰🍰끝! 실행 완료🍰🍰🍰")
        time.sleep(sleep_time)
    except Exception as e:
        log_set_callback(log_get_callback() + f"\n❗❗❗오류 발생: {e}")
