from module import *
from print_out_product_instruction import *
from make_two_files import *
import os
import webbrowser
import subprocess #GUI
import tkinter as tk #GUI
import threading #GUI 멀티스레드 사용하기 위해
import time #GUI에서 멀티스레드 사용하기 위해

def prepare_to_pack():
    sleep_time = 0.1
    log_text.set("시작! 프로그램 실행")

    ##########################################다운로드폴더에서 가져와서 쪼개기##########################################
    #setting\path.csv에서 쪼갤 파일이 있는 폴더 경로 가져오기
    download_from_internet_path = search_path("download_from_internet")


    #카페24에서 다운받은 파일 찾기
    #카페24에서 다운받는 파일명의 형식: lalapetmall_오늘날짜_일련번호_일련번호
    download_from_cafe24_path = find_file_by_partial_name(download_from_internet_path, "lalapetmall_" + datetime.today().strftime('%Y%m%d') + "_")


    ####두 가지 파일로 쪼개기
    #주문 리스트 파일
    order_list_path = r"result\order_list.xlsx"
    order_list_header_list = get_column_from_csv(r"settings\header.csv", "order_list_header")
    order_list_header_index = [find_header_index(download_from_cafe24_path, order_list_header) for order_list_header in order_list_header_list]
    split_csv_by_column_index(download_from_cafe24_path, order_list_path, order_list_header_index)


    #한진택배리스트 파일
    hanjin_path = r"result\hanjin_file.xlsx"
    hanjin_header_list = get_column_from_csv(r"settings\header.csv", "hanjin_header")
    hanjin_header_index = [find_header_index(download_from_cafe24_path, hanjin_header) for hanjin_header in hanjin_header_list]
    split_csv_by_column_index(download_from_cafe24_path, hanjin_path , hanjin_header_index)
    
    #log
    log_text.set(log_text.get() + "\n다운로드 받은 파일을 두 개의 파일로 분리")
    time.sleep(sleep_time) 

    ##########################################print_out_product_instruction##########################################
    order_list_pd = pd.read_excel(r"result\order_list.xlsx", engine='openpyxl')

    # '중량' 열을 업데이트
    order_list_pd['중량'] = order_list_pd.apply(get_final_weight, axis=1)

    # 수정된 내용을 새로운 엑셀 파일로 저장
    order_list_pd.to_excel(r"result\order_list.xlsx", index=False, engine='openpyxl')
    
    #log
    log_text.set(log_text.get() + "\n주문리스트의 중량 정보 입력")
    time.sleep(sleep_time) 


    converted_codes = ready_to_convert(order_list_pd)
    not_found_files = merge_pdf(converted_codes)
    report_result(order_list_pd, not_found_files)

    #매크로 실행
    run_macro("전채널주문리스트", order_list_path)

    #log
    log_text.set(log_text.get() + "\n주문리스트 파일 완성")
    time.sleep(sleep_time) 

    ##########################################make_two_files##########################################
    match_to_cafe24_example(hanjin_path)
    #매크로 실행(기존 파일을 한진택배 복수내품 양식에 맞게 변경하기 위해)
    run_macro("ProcessMultipleItems", hanjin_path) 
    os.rename(hanjin_path, r"result\upload_to_hanjin.xlsx")

    #log
    log_text.set(log_text.get() + "\n한진 사이트에 올릴 파일 완성")
    time.sleep(sleep_time)

    ####한진택배 사이트 열기
    webbrowser.open("https://focus.hanjin.com/login")

    
    os.startfile("result") #폴더 열기
    print("실행 완료.")

    #log
    log_text.set(log_text.get() + "\n끝! 실행 완료")
    time.sleep(sleep_time)


##########################################GUI##########################################
#.pack()은 부모위젯 안에 배치

# 메인 윈도우 생성
root = tk.Tk()
root.title("prepare_to_pack")
root.geometry("350x350")  # 너비x높이
root.configure()
root.attributes('-topmost', True) # 창이 포커스를 잃어도 항상 다른 창들보다 위에 표시

font_size = 14

# 프레임을 사용하여 내부 여백 추가
frame = tk.Frame(root, padx=20, pady=20)
frame.pack()

label = tk.Label(root, text="상품포장준비 프로그램\n문의:seomar2022@gmail.com", font=font_size)
label.pack(pady=20)  # 여백 설정  # 위젯을 창에 배치

def on_button_click():
    # 별도의 스레드에서 프로그램 로직 실행
    threading.Thread(target=prepare_to_pack).start()

def run_python_program(script_path):
    try:
        # Python 스크립트 실행
        subprocess.run(['python', script_path], check=True)
        print("prepare_to_pack.py 실행 완료.")
    except Exception as e:
        print(f"파일 실행 중 오류가 발생했습니다: {e}")

# 버튼 프레임 생성
button_frame = tk.Frame(root)
button_frame.pack(pady=10)
prepare_image = tk.PhotoImage(file="resources/img/box.png")
#https://www.flaticon.com/free-icon/box_679720?term=packing&page=1&position=1&origin=search&related_id=679720

upload_image = tk.PhotoImage(file="resources/img/order-fulfillment.png")
#https://www.flaticon.com/free-icon/order-fulfillment_11482468?term=delivery+label&page=1&position=1&origin=search&related_id=11482468


# 포장 준비 버튼 추가
prepare_button = tk.Button(button_frame, image=prepare_image, command=on_button_click)
prepare_button.pack(side="left", padx=10)

# 송장 업로드 버튼 추가
upload_button = tk.Button(button_frame, image=upload_image, command=lambda: run_python_program("match_cafe24_with_hanjin.py"), font=font_size)
upload_button.pack(side="right", padx=10)

# 로그 텍스트를 표시할 라벨 생성
log_text = tk.StringVar()
log_text.set("")

log_label = tk.Label(root, textvariable=log_text, justify="left", anchor="nw", font=font_size, wraplength=300)
log_label.pack(pady=10, padx=20)

# 메인 루프 시작
root.mainloop()

#pyinstaller --onefile print_out_product_instruction.py
