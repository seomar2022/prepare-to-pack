from module import *
from print_out_product_instruction import *
import pyautogui
from make_two_files import *
import os
import webbrowser


pyautogui.alert(text="상품포장준비 프로그램입니다!\n문의:seomar2022@gmail.com", title="prepare_to_pack", button="실행!")
#####################다운로드폴더에서 가져와서 쪼개기 #####################
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


#####################print_out_product_instruction#####################
order_list_pd = pd.read_excel(r"result\order_list.xlsx", engine='openpyxl')


# '중량' 열을 업데이트
order_list_pd['중량'] = order_list_pd.apply(get_final_weight, axis=1)

# 수정된 내용을 새로운 엑셀 파일로 저장
order_list_pd.to_excel(r"result\order_list.xlsx", index=False, engine='openpyxl')

converted_codes = ready_to_convert(order_list_pd)
not_found_files = merge_pdf(converted_codes)
report_result(order_list_pd, not_found_files)

#매크로 실행
run_macro("전채널주문리스트", order_list_path)

#####################make_two_files#####################
match_to_cafe24_example(hanjin_path)
#매크로 실행(기존 파일을 한진택배 복수내품 양식에 맞게 변경하기 위해)
run_macro("ProcessMultipleItems", hanjin_path) 
os.rename(hanjin_path, r"result\upload_to_hanjin.xlsx")



####한진택배 사이트 열기
webbrowser.open("https://focus.hanjin.com/login")

pyautogui.alert(text="실행 결과를 확인해주세요", title="prepare_to_pack", button='네!')
os.startfile("result") #폴더 열기
print("실행 완료.")

#pyinstaller --onefile print_out_product_instruction.py
