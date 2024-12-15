from module import *
from before_packing import *
from upload_tracking_number import upload_tracking_number
import os
import webbrowser
import tkinter as tk #GUI
import threading #GUI 멀티스레드 사용하기 위해
import time #GUI에서 멀티스레드 사용하기 위해
from datetime import datetime #폴더이름에 현재 날짜 넣기 위해

def prepare_to_pack():
    try:
        sleep_time = 0.1
        log_text.set("시작! 프로그램 실행")

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
        log_text.set(log_text.get() + "\n다운로드 받은 파일 검색")
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
        log_text.set(log_text.get() + "\n헤더명에 따라 두 개의 파일로 분리")
        time.sleep(sleep_time) 

        ##########################################print_out_product_instruction##########################################
        order_list_pd = pd.read_excel(rf"{result_directory}\order_list.xlsx", engine='openpyxl')

        #### '중량' 열 업데이트
        order_list_pd['중량'] = order_list_pd.apply(get_final_weight, axis=1)
        # 수정된 내용을 새로운 엑셀 파일로 저장
        order_list_pd.to_excel(rf"{result_directory}\order_list.xlsx", index=False, engine='openpyxl')
        #log
        log_text.set(log_text.get() + "\n주문리스트의 중량 정보 입력")
        time.sleep(sleep_time) 

        ####설명지 찾아서 병합
        converted_codes = ready_to_convert(order_list_pd)
        not_found_files = merge_pdf(result_directory, converted_codes)
        report_result(result_directory, order_list_pd, not_found_files)
        #log
        log_text.set(log_text.get() + "\n상품 설명지 병합")
        time.sleep(sleep_time) 

        ####매크로 실행(포장할 때 참고할 주문리스트 만들기 위해)
        run_macro("전채널주문리스트", order_list_path)
        #log
        log_text.set(log_text.get() + "\n전채널주문리스트 매크로 실행\n주문리스트 파일 작성")
        time.sleep(sleep_time) 

        ####카페24 양식에 맞게 수정한 파일 만들기
        match_to_cafe24_example(result_directory, hanjin_path)  
        #log
        log_text.set(log_text.get() + "\n송장등록을 위한 카페24양식 파일 작성")
        time.sleep(sleep_time)

        ####매크로 실행(기존 파일을 한진택배 복수내품 양식에 맞게 변경하기 위해)
        run_macro("ProcessMultipleItems", hanjin_path) 
        os.rename(hanjin_path, rf"{result_directory}\upload_to_hanjin.xlsx")
        #log
        log_text.set(log_text.get() + "\nProcessMultipleItems 매크로 실행\n한진 사이트에 올릴 파일 작성")
        time.sleep(sleep_time)

        ####한진택배 사이트 열기
        webbrowser.open("https://focus.hanjin.com/login")

        ####result 폴더 열기
        os.startfile(f"{result_directory}")
        print("실행 완료.")
        #log
        log_text.set(log_text.get() + "\n끝! 실행 완료")
        time.sleep(sleep_time)
    except Exception as e:
        log_text.set(log_text.get() + f"\n오류 발생: {e}") 

def on_before_packing_button_click():
    # 별도의 스레드에서 프로그램 로직 실행
    threading.Thread(target=prepare_to_pack).start()

def on_upload_tracking_number_button_click():
    try:
        upload_tracking_number()
        log_text.set("실행 완료\n카페24 엑셀 일괄배송 처리란에 업로드해 주세요.")
    except Exception as e:
        log_text.set(f"오류 발생: {e}") 

##########################################GUI##########################################
#.pack()은 부모위젯 안에 배치

# 메인 윈도우 생성
if __name__ == "__main__":
    root = tk.Tk()
    root.title("prepare_to_pack")
    root.geometry("350x450")  # 너비x높이
    root.configure()
    root.attributes('-topmost', True) # 창이 포커스를 잃어도 항상 다른 창들보다 위에 표시


    font_size = 14

    # 프레임을 사용하여 내부 여백 추가
    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack()

    label = tk.Label(root, text="출고 준비 프로그램", font=("none", font_size, "bold"))
    label.pack(pady=10)  # 여백 설정  # 위젯을 창에 배치

    ####버튼
    # 버튼 프레임 생성
    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)

    #버튼 이미지
    before_packing_image = tk.PhotoImage(file="resources/img/package-box.png")
    upload_image = tk.PhotoImage(file="resources/img/document.png")
    info_image = tk.PhotoImage(file="resources/img/info.png")

    # 포장 준비 버튼
    before_packing_button = tk.Button(button_frame, image=before_packing_image, command=on_before_packing_button_click)
    before_packing_button.pack(side="left", padx=10)
    ToolTip(before_packing_button, "cafe24에서 '출고준비통합'양식으로 파일을 다운로드 받은 후 이 버튼을 클릭해 주세요.")

    # 송장 업로드 버튼
    upload_tracking_number_button = tk.Button(button_frame, image=upload_image, command=on_upload_tracking_number_button_click, font=font_size)
    upload_tracking_number_button.pack(side="right", padx=10)
    ToolTip(upload_tracking_number_button, "한진택배에서 '원본파일'을 다운로드 받은 후 이 버튼을 클릭해 주세요.")

    # info 버튼
    info_button = tk.Button(root, image=info_image)
    info_button.pack(side="bottom", pady=20)
    ToolTip(info_button, "-cafe24 엑셀파일 다운 양식 수정: \settings\header.csv\n-인터넷에서 다운 받은 파일이 있는 폴더 경로 지정: \settings\path.csv\n-설명지 추가: \\resources\product_instruction\n-image: Flaticon.com\n-기타문의:seomar2022@gmail.com")


    #### 로그 텍스트를 표시할 라벨 생성
    log_text = tk.StringVar()
    log_text.set("")

    log_label = tk.Label(root, textvariable=log_text, justify="left", anchor="nw", font=font_size, wraplength=300)
    log_label.pack(pady=10, padx=20)

    # 메인 루프 시작
    root.mainloop()

    #pyinstaller --onefile --noconsole prepare_to_pack.py
