import os
from src.logic.config import get_instruction_folder, get_product_code_mapping
from pypdf import PdfWriter
import pandas as pd
import re
import logging

logger = logging.getLogger(__name__)


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

    if total_weight_by_order < 1:
        return 1
    elif total_weight_by_order < 2:
        return 2
    elif total_weight_by_order < 3.8:
        return 3
    elif total_weight_by_order <= 4:
        return 420
    elif total_weight_by_order < 8:
        return 287
    else:
        return 0


# Vectorized - A way of processing data all at once, applying operations to the entire array/series without using loops.
def convert_to_cafe24_product_code(order_list_df):
    ### Clean product code data for mapping
    logger.info("Cleaning product code data...")

    raw_product_codes = order_list_df["product_code"].astype(str).fillna("").tolist()
    logger.info(f"Number of product codes to convert: {len(raw_product_codes)}")

    # Load and clean product code mappings for Cafe24, Naver, and Kakao
    logger.info("Loading product code mapping file...")
    product_code_mapping_df = pd.read_excel(
        get_product_code_mapping(), engine="openpyxl", dtype=str
    )

    logger.info("Cleaning mapping columns (naver_code, kakao_code, cafe24_code)...")
    # Clean mapping columns
    for col in ["naver_code", "kakao_code", "cafe24_code"]:
        product_code_mapping_df[col] = (
            product_code_mapping_df[col]
            .astype(str)
            .str.strip()
            .str.replace("-", "", regex=False)
        )
    logger.info("Mapping columns cleaned successfully.")

    # prefixes of the codes
    prefix_to_column = {
        "P00": None,  # cafe24. No need to convert.
        "9": "naver_code",
        "1": "naver_code",
        "3": "kakao_code",
    }

    converted_codes = order_list_df["product_code"].astype(str).copy()
    missing_matches = []

    # df.loc[]: Selects rows by index label and columns by name (label-based indexing, not position-based).
    # df.loc[{Boolean mask (Series)}]: Pandas matches the mask’s index labels with the DataFrame’s index, then selects only rows where the mask is True, regardless of order.
    for prefix, column in prefix_to_column.items():
        mask = converted_codes.str.startswith(prefix)
        if not mask.any():
            continue

        if column is None:
            continue  # Already cafe24, no change

        # Remove duplicate keys from the mapping table and select only the necessary columns
        # df[[]]→ select multiple columns
        # subset=[column] → When determining duplicates, only consider values in the specified column (ignore other columns)
        # keep="last" → If duplicates exist, keep only the last occurrence and drop the rest
        right_df = product_code_mapping_df[[column, "cafe24_code"]].drop_duplicates(
            subset=[column], keep="last"
        )

        ### Merge for vectorized mapping
        # left_df.merge(
        #     right_df,
        #     left_on="column",
        #     right_on="column2"
        # )
        temp_df = order_list_df.loc[mask, ["product_code"]].merge(
            right_df,
            left_on="product_code",
            right_on=column,
            how="left",
        )
        # logger.info(f"right_df: {right_df}")
        # logger.info(f"temp_df: {temp_df}")
        # logger.info(f'temp_df["cafe24_code"]: {temp_df["cafe24_code"]}')

        # For unmapped entries, fill with the original code and assign based on position rather than index alignment
        # Using .to_numpy() bypasses Pandas index alignment, allowing assignment strictly by position, and is also lighter in terms of performance.
        fallback = order_list_df.loc[mask, "product_code"].to_numpy()
        filled = temp_df["cafe24_code"].fillna(pd.Series(fallback, index=temp_df.index))
        converted_codes.loc[mask] = filled.to_numpy()

        # logger.info(f"mask: {mask}")
        # logger.info(f"converted_codes: {converted_codes}")
        # logger.info(f"converted_codes[mask]: {converted_codes[mask]}")

        # Track missing matches
        missing_rows = temp_df[temp_df["cafe24_code"].isna()]
        if not missing_rows.empty:
            missing_matches.extend(missing_rows["product_code"].tolist())

    logger.info(f"Conversion completed. Total unmatched codes: {len(missing_matches)}")

    return converted_codes.tolist(), missing_matches


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
def report_missing_instructions(result_directory, order_list_df, not_found_files):
    if len(not_found_files) == 0:
        logger.info("Found instruction sheets for all products")
        return

    # .query("col_name == @variable_from_python")
    # This returns a new DataFrame with only the rows where the cafe24_code column equals the value in the converted_code variable.
    for product_code in not_found_files:
        key_in_order_list = order_list_df.query("product_code == @product_code")
        logger.info(key_in_order_list)

        not_found_files[product_code] = key_in_order_list["product_name"].iat[0]

    # Save product codes without instruction sheets to CSV and show summary message
    no_instruction_sheets_df = pd.DataFrame(
        list(not_found_files.items()), columns=["상품코드", "상품명"]
    )
    no_instruction_sheets_df.to_csv(
        f"{result_directory}\\not_found_files.csv",
        index=False,
        encoding="utf-8-sig",
    )
    logger.info(f"Could not find {len(not_found_files)} instruction sheets.")


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
