from pypdf import PdfWriter  # pip install pypdf #pdf병합기능 쓰기 위해.
import pandas as pd  # pip install pandas openpyxl #엑셀의 데이터를 읽어오기 위해.
import re


# 문자열에서 중량 정보를 추출
def extract_weight(product_data):
    if not isinstance(product_data, str):
        return None
    product_data = product_data.lower().replace("ml", "g")

    match = re.search(r"(\d+(\.\d+)?)\s*(kg|g)", product_data, re.IGNORECASE)

    if match:
        weight = float(match.group(1))
        unit = match.group(3).lower()

        # 단위가 g일 경우 kg으로 변환
        if unit == "g":
            weight = weight / 1000

        return weight
    else:
        return None


# 조건에 따라 중량 정보를 선택
def get_final_weight(row):
    if pd.notna(row["weight"]) == False:  # 중량열에 데이터 없음
        # ->스마트 스토어나 톡스토어 주문건임. 상품명에서 중량 추출해야함.
        weight_from_name_column = extract_weight(row["product_name"])
        if weight_from_name_column is not None:
            return weight_from_name_column

    else:  # 중량열에 데이터 있음
        # ->카페24주문 건임(몇몇 연동된 상품 제외). 상품옵션열의 데이터에 중량 정보기 중량열의 데이터가 다를 때만 상품옵션 데이터 쓰기.
        weight_from_option_column = extract_weight(row["option"])
        if weight_from_option_column is None:  # 상품옵션에 중량 정보가 없을경우
            return row["weight"]
        else:  # 상품옵션에 중량 정보가 있을 경우
            return weight_from_option_column


def convert_to_cafe24_product_code(order_list_pd):
    ### Clean product code data for mapping
    # Convert values in the 상품코드(product code) column to strings and replace NaN with empty strings
    raw_product_codes = order_list_pd["product_code"].astype(str).fillna("").tolist()

    # Load and clean product code mappings for Cafe24, Naver, and Kakao
    df_product_code_mapping = pd.read_excel(
        r"resources\product_code_mapping.xlsx", engine="openpyxl"
    )

    df_product_code_mapping["naver_code"] = (
        df_product_code_mapping["naver_code"]
        .astype(str)
        .str.strip()
        .str.replace("-", "")
    )
    df_product_code_mapping["kakao_code"] = (
        df_product_code_mapping["kakao_code"]
        .astype(str)
        .str.strip()
        .str.replace("-", "")
    )
    df_product_code_mapping["cafe24_code"] = (
        df_product_code_mapping["cafe24_code"].astype(str).str.strip()
    )

    ### Convert
    def convert_to_cafe24(code, column):
        # Row where 'code' exists in the specified 'column' (e.g., 9708250509 in the 'naver_code' column)
        matched_row = df_product_code_mapping.query(f"{column} == @code")
        return matched_row["cafe24_code"].iat[0]

    # prefixes of the codes
    prefix_to_column = {
        "P00": None,  # cafe24. No need to convert.
        "9": "naver_code",
        "1": "naver_code",
        "3": "kakao_code",
    }

    converted_cafe24_codes = []

    for raw_product_code in raw_product_codes:
        for prefix, column in prefix_to_column.items():
            if raw_product_code.startswith(prefix):
                if column is None:
                    converted_cafe24_codes.append(raw_product_code)
                else:
                    mapped = convert_to_cafe24(raw_product_code, column)
                    converted_cafe24_codes.append(mapped)
                break

    return converted_cafe24_codes


def merge_product_instructions(output_folder, converted_codes):
    merge_pdf = PdfWriter()
    not_found_files = {}

    # Append the instruction sheet corresponding to each converted code
    for converted_code in converted_codes:
        try:
            merge_pdf.append(f"resources\\product_instruction\\{converted_code}.pdf")
        except FileNotFoundError:
            not_found_files[converted_code] = ""

    merge_pdf.write(f"{output_folder}\\product_instruction.pdf")
    merge_pdf.close()
    return not_found_files


### Report product codes and names for which no instruction sheet was found
def report_missing_instructions(result_directory, order_list_pd, not_found_files):
    # Find product codes without instruction sheets in order_list_pd, and get the corresponding product names to use as values in the dictionary
    product_name_col = [col for col in order_list_pd.columns if "product_name" in col]
    for key in not_found_files:
        key_in_order_list = order_list_pd.query(f"product_code == @key")
        not_found_files[key] = key_in_order_list[product_name_col[0]].iat[0]

    # Save product codes without instruction sheets to CSV and show summary message
    if len(not_found_files) == 0:
        alert_msg = "모든 상품의 설명지를 찾았습니다!"
    else:
        converted_codes_df = pd.DataFrame(
            list(not_found_files.items()), columns=["상품코드", "상품명"]
        )
        converted_codes_df.to_csv(
            f"{result_directory}\\not_found_files.csv",
            index=False,
            encoding="utf-8-sig",
        )
        alert_msg = f"{len(not_found_files)}개의 설명지를 찾지 못했습니다"
        print(alert_msg)

### Determine gift type based on priority: gift_selection > pet_type > product_name
def assign_gift(row):
    gift_selection = (
        ""
        if pd.isna(row.get("gift_selection"))
        else str(row.get("gift_selection")).strip()
    )
    pet_type = "" if pd.isna(row.get("pet_type")) else str(row.get("pet_type")).strip()
    product_name = (
        "" if pd.isna(row.get("product_name")) else str(row.get("product_name")).strip()
    )

    if "강아지용" in gift_selection:
        return "독"
    elif "고양이용" in gift_selection:
        return "캣"
    elif gift_selection == "" and pet_type:
        return pet_type
    elif gift_selection == "" and pet_type == "":
        if "독" in product_name:
            return "독"
        elif "캣" in product_name:
            return "캣"
    return "?"


def match_to_cafe24_example(result_directory, hanjin_path):
    ####배송리스트 파일 읽어오기
    delivery_list = pd.read_excel(hanjin_path, engine="openpyxl")

    ####카페24 양식에 맞게 수정한 파일 만들기
    try:
        # B열의 데이터까지만 남겨두기.
        upload_to_cafe24 = delivery_list.iloc[:, :2]

        # 수정된 내용을 새로운 CSV 파일로 저장
        upload_to_cafe24.to_csv(
            rf"{result_directory}\excel_sample_old.csv",
            index=False,
            encoding="utf-8-sig",
        )

    except Exception as e:
        print(f"파일 편집 중 오류가 발생했습니다: {e}")
