from module import *

#####################다운로드폴더에서 가져와서 쪼개기 #####################
#setting\path.csv에서 쪼갤 파일이 있는 폴더 경로 가져오기
download_from_internet_path = search_path("download_from_internet")


#카페24에서 다운받은 파일 찾기
#카페24에서 다운받는 파일명의 형식: lalapetmall_오늘날짜_일련번호_일련번호
download_from_cafe24_path = find_file_by_partial_name(download_from_internet_path, "lalapetmall_" + datetime.today().strftime('%Y%m%d') + "_")


####두 가지 파일로 쪼개기
#주문 리스트 파일
order_list_header_list = get_column_from_csv(r"settings\header.csv", "order_list_header")
order_list_header_index = [find_header_index(download_from_cafe24_path, order_list_header) for order_list_header in order_list_header_list]
split_csv_by_column_index(download_from_cafe24_path,r"result\order_list_raw_file.xlsx" ,order_list_header_index)


#한진택배리스트 파일
hanjin_header_list = get_column_from_csv(r"settings\header.csv", "hanjin_header")
hanjin_header_index = [find_header_index(download_from_cafe24_path, hanjin_header) for hanjin_header in hanjin_header_list]
split_csv_by_column_index(download_from_cafe24_path,r"result\hanjin_raw_file.xlsx" ,hanjin_header_index)
