import pandas as pd

def match_to_cafe24_example(hanjin_path):
    ####배송리스트 파일 읽어오기
    
    delivery_list = pd.read_excel(hanjin_path, engine='openpyxl')


    ####카페24 양식에 맞게 수정한 파일 만들기
    try:
        # B열의 데이터까지만 남겨두기.
        upload_to_cafe24 = delivery_list.iloc[:, :2]

        # 수정된 내용을 새로운 CSV 파일로 저장
        upload_to_cafe24.to_csv(r"result\excel_sample_old.csv", index=False, encoding='utf-8-sig')
    
    except Exception as e:
        print(f"파일 편집 중 오류가 발생했습니다: {e}")
