from module import *
from before_packing import *
from upload_tracking_number import upload_tracking_number
import os
import webbrowser
import time #GUI에서 멀티스레드 사용하기 위해
from datetime import datetime #폴더이름에 현재 날짜 넣기 위해


def prepare_to_pack(log_set_callback, log_get_callback):
    try:
        print("start")
        sleep_time = 0.1
        log_set_callback("시작! 프로그램 실행")

        ####완성된 파일들을 넣어둘 폴더 만들기
        result_directory = "result_" + datetime.now().strftime("%a.%H.%M.%S") #요일.시.분.초
        os.makedirs(result_directory)

        ##########################################다운로드폴더에서 가져와서 쪼개기##########################################
        #setting\path.csv에서 쪼갤 파일이 있는 폴더 경로 가져오기
        download_from_internet_path = search_path("download_from_internet")

        #카페24에서 다운받은 파일 찾기
        #카페24에서 다운받는 파일명의 형식: lalapetmall_오늘날짜_일련번호_일련번호
        download_from_cafe24_path = find_path_by_partial_name(download_from_internet_path, "lalapetmall_" + datetime.today().strftime('%Y%m%d') + "_")
        #log
        log_set_callback(log_get_callback() + "\n다운로드 받은 파일 검색")
        time.sleep(sleep_time) 

        ####두 가지 파일로 쪼개기
        #주문 리스트 파일
        order_list_path = rf"{result_directory}\order_list.xlsx"
        order_list_header_list = get_column_from_csv(r"settings\header.csv", "order_list_header")
        order_list_header_index = [find_header_index(download_from_cafe24_path, order_list_header) for order_list_header in order_list_header_list]
        split_csv_by_column_index(download_from_cafe24_path, order_list_path, order_list_header_index)

        #한진택배리스트 파일
        hanjin_path = rf"{result_directory}\hanjin_file.xlsx"
        hanjin_header_list = get_column_from_csv(r"settings\header.csv", "hanjin_header")
        hanjin_header_index = [find_header_index(download_from_cafe24_path, hanjin_header) for hanjin_header in hanjin_header_list]
        split_csv_by_column_index(download_from_cafe24_path, hanjin_path , hanjin_header_index)
        
        #log
        log_set_callback(log_get_callback() + "\n헤더명에 따라 두 개의 파일로 분리")
        time.sleep(sleep_time) 

        ##########################################print_out_product_instruction##########################################
        order_list_pd = pd.read_excel(rf"{result_directory}\order_list.xlsx", engine='openpyxl')

        #### '중량' 열 업데이트
        order_list_pd['중량'] = order_list_pd.apply(get_final_weight, axis=1)
        # 수정된 내용을 새로운 엑셀 파일로 저장
        order_list_pd.to_excel(rf"{result_directory}\order_list.xlsx", index=False, engine='openpyxl')
        #log
        log_set_callback(log_get_callback() + "\n주문리스트의 중량 정보 입력")
        time.sleep(sleep_time) 

        ####설명지 찾아서 병합
        converted_codes = ready_to_convert(order_list_pd)
        not_found_files = merge_pdf(result_directory, converted_codes)
        report_result(result_directory, order_list_pd, not_found_files)
        #log
        log_set_callback(log_get_callback() + "\n상품 설명지 병합")
        time.sleep(sleep_time) 

        ####매크로 실행(포장할 때 참고할 주문리스트 만들기 위해)
        run_macro("전채널주문리스트", order_list_path)
        #log
        log_set_callback(log_get_callback() + "\n전채널주문리스트 매크로 실행\n주문리스트 파일 작성")
        time.sleep(sleep_time) 

        ####카페24 양식에 맞게 수정한 파일 만들기
        match_to_cafe24_example(result_directory, hanjin_path)  
        #log
        log_set_callback(log_get_callback() + "\n송장등록을 위한 카페24양식 파일 작성")
        time.sleep(sleep_time)

        ####매크로 실행(기존 파일을 한진택배 복수내품 양식에 맞게 변경하기 위해)
        run_macro("ProcessMultipleItems", hanjin_path) 
        os.rename(hanjin_path, rf"{result_directory}\upload_to_hanjin.xlsx")
        #log
        log_set_callback(log_get_callback() + "\nProcessMultipleItems 매크로 실행\n한진 사이트에 올릴 파일 작성")
        time.sleep(sleep_time)

        ####한진택배 사이트 열기
        webbrowser.open("https://focus.hanjin.com/login")

        ####result 폴더 열기
        os.startfile(f"{result_directory}")
        print("실행 완료.")
        #log
        log_set_callback(log_get_callback() + "\n끝! 실행 완료")
        time.sleep(sleep_time)
    except Exception as e:
        log_set_callback(log_get_callback() + f"\n오류 발생: {e}") 
