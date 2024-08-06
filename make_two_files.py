import os
import webbrowser
import pyautogui
import sys
import pandas as pd
import xlwings as xw #매크로실행위해
import tkinter as tk

def make_two_files():
    ####배송리스트 파일 읽어오기
    delivery_list_path = f"result\\hanjin_original_file.xlsx"
    delivery_list = pd.read_excel(delivery_list_path, engine='openpyxl')


    ####카페24 양식에 맞게 수정한 파일 만들기
    try:
        # B열의 데이터까지만 남겨두기.
        upload_to_cafe24 = delivery_list.iloc[:, :2]

        # 수정된 내용을 새로운 CSV 파일로 저장
        upload_to_cafe24.to_csv(r"result\excel_sample_old.csv", index=False, encoding='utf-8-sig')
    
    except Exception as e:
        print(f"파일 편집 중 오류가 발생했습니다: {e}")


    ####매크로 실행(기존 파일을 한진택배 복수내품 양식에 맞게 변경하기 위해)
    #엑셀 모두 닫은 상태에서 시작해야할 듯.
    try:
        # 엑셀 애플리케이션 시작 및 파일 열기 (빈 통합 문서 생성을 방지)
        app = xw.App(visible=True, add_book=False)
        workbook = app.books.open(delivery_list_path)
        
        # personal.xlsb 파일 열기(매크로가 저장된 파일)
        # 아마 여기는 각 컴퓨터에 맞춰서 별도로 지정해야할듯? ->초기설정으로 넣기
        personal_wb = app.books.open(r'C:\Users\User\AppData\Roaming\Microsoft\Excel\XLSTART\PERSONAL.XLSB')
        
        # test.xlsx 파일을 활성화(매크로가 적용될 파일이므로)
        workbook.activate()
        
        # 매크로 실행 (personal_wb에서 호출)
        macro = personal_wb.macro('ProcessMultipleItems') 
        macro()

        # 엑셀 파일 저장 및 닫기
        new_file_path = os.path.join(r"result\upload_to_hanjin.xlsx")
        workbook.save(new_file_path)
        workbook.close()

                
        
        # personal.xlsb 파일 닫기
        personal_wb.close()

        # 엑셀 애플리케이션 종료
        app.quit()
        
        print(f"매크로가 성공적으로 실행되었습니다.")
        
    except Exception as e:
        print(f"매크로 실행 중 오류가 발생했습니다: {e}")

    ####한진택배 사이트 열기
    webbrowser.open("https://focus.hanjin.com/login")
