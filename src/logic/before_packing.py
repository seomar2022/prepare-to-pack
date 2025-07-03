import os
from src.logic.config import get_instruction_folder, get_product_code_mapping
from pypdf import PdfWriter
import pandas as pd
import re


# 문자열에서 중량 정보를 추출
def extract_weight(product_data):
    if not isinstance(product_data, str):
        return None
    product_data = product_data.lower().replace("ml", "g")

    match = re.search(r"(\d+(\.\d+)?)\s*(kg|g)", product_data, re.IGNORECASE)

    if match:
        unit_weight = float(match.group(1))
        unit = match.group(3).lower()

        # 단위가 g일 경우 kg으로 변환
        if unit == "g":
            unit_weight = unit_weight / 1000

        return unit_weight
    else:
        return None


def get_adjusted_unit_weight(row):
    ### Select unit weight based on conditions
    if not pd.notna(row["unit_weight"]):  # No data in the unit_weight column
        # -> Likely Naver or Kakao. Extract weight from product name.
        weight_from_name_column = extract_weight(row["product_name"])
        if weight_from_name_column is not None:
            return weight_from_name_column

    else:  # Data exists in the unit_weight column
        # -> Likely Cafe24 (except for some synced products).
        # If weight in option column differs from unit_weight column, use the option column.
        weight_from_option_column = extract_weight(row["option"])
        if weight_from_option_column is None:  # No weight info in the option column
            return row["unit_weight"]
        else:  # Weight info exists in the option column
            return weight_from_option_column


def determine_box_size(row):
    if "냉장배송" in row["product_name"]:
        return "Ice"

    total_weight_by_order = row["total_weight_by_order"]
    brand = row["brand"]
    quantity = row["quantity"]

    if total_weight_by_order < 1:
        return 73
    elif total_weight_by_order < 2:
        return 194
    elif total_weight_by_order < 3.8:
        return 41
    elif total_weight_by_order <= 4:
        if brand == "Royal Canin" and quantity == 2:
            return 104
        else:
            return 420
    elif total_weight_by_order <= 4.3:
        return 104
    elif total_weight_by_order < 8:
        return 170
    else:
        return 266


def convert_to_cafe24_product_code(order_list_pd):
    ### Clean product code data for mapping
    # Convert values in the 상품코드(product code) column to strings and replace NaN with empty strings
    raw_product_codes = order_list_pd["product_code"].astype(str).fillna("").tolist()

    # Load and clean product code mappings for Cafe24, Naver, and Kakao
    df_product_code_mapping = pd.read_excel(
        get_product_code_mapping(), engine="openpyxl"
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
            merge_pdf.append(
                os.path.join(get_instruction_folder(), f"{converted_code}.pdf")
            )
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
        key_in_order_list = order_list_pd.query("product_code == @key")
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


# Define a function to restructure each group
def flatten_order_items_by_order_number(df):
    """
    Restructure the DataFrame by grouping rows by 'order_number' and
    flattening product information into a single row per order.

    Args:
        df (pd.DataFrame): The input DataFrame with order data.
    Returns:
        pd.DataFrame: A restructured DataFrame with one row per order_number.
    """
    grouped = df.groupby("order_number")
    base_columns = df.columns.tolist()

    def restructure(group):
        # Get the first row's base info
        first_row = group.iloc[0][base_columns].tolist()
        rest = []
        for _, row in group.iloc[1:].iterrows():
            rest.extend([row["product_name_with_option"], row["quantity"]])
        return first_row + rest

    # Apply restructure and convert to DataFrame
    restructured_data = grouped.apply(restructure).apply(pd.Series)

    # Generate column names
    extra_cols = []
    max_extra = restructured_data.shape[1] - len(base_columns)
    for i in range(0, max_extra, 2):
        extra_cols.extend(["product_name_with_option", "quantity"])

    restructured_data.columns = base_columns + extra_cols
    return restructured_data
