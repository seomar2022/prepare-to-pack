# column_mapping.py

KOR_TO_ENG_COLUMN_MAP = {
    "매출경로": "sales_channel",
    "주문번호": "order_number",
    "주문자명": "orderer_name",
    "브랜드": "brand",
    "상품명(한국어 쇼핑몰)": "product_name",
    "상품옵션": "option",
    "수량": "quantity",
    "수령인": "recipient_name",
    "주문서추가항목01_사은품 선택 (공통입력사항)": "gift_selection",
    "옵션+판매가": "price",
    "주문 상태": "order_status",
    "수령인 주소(전체)": "recipient_address",
    "배송메시지": "delivery_message",
    "중량": "unit_weight",
    "정기배송 회차": "subscription_cycle",
    "상품코드": "product_code",
    "회원추가항목_반려견/반려묘의 종류": "pet_type",
    "주문 시 회원등급": "membership_level",
    "품목별 주문번호": "line_item_number",
    "운송장번호": "tracking_number",
    "수령인 전화번호": "recipient_phone",
    "수령인 휴대전화": "recipient_mobile",
    "수령인 우편번호": "recipient_zip_code",
    "주문상품명(옵션포함)": "product_name_with_option",
}

ENG_TO_KOR_COLUMN_MAP = {
    "sales_channel": "매출경로",
    "order_number": "주문번호",
    "orderer_name": "주문자",  # 주문자명 → 주문자
    "brand": "브랜드",
    "product_name": "상품명",  # 상품명(한국어 쇼핑몰) → 상품명
    "option": "옵션",  # 상품옵션 → 옵션
    "quantity": "수량",
    "recipient_name": "수령인",
    "gift_selection": "주문서추가항목",  # 주문서추가항목01_사은품 선택 → 주문서추가항목
    "price": "가격",  # 옵션+판매가 → 가격
    "order_status": "주문 상태",
    "recipient_address": "주소",  # 수령인 주소(전체) → 주소
    "delivery_message": "배송메시지",
    "unit_weight": "중량",
    "subscription_cycle": "정기배송 회차",
    "product_code": "상품코드",
    "pet_type": "견묘종",  # 회원추가항목 → 견묘종
    "membership_level": "주문 시 회원등급",
    "line_item_number": "품목별 주문번호",
    "tracking_number": "운송장번호",
    "recipient_phone": "수령인 전화번호",
    "recipient_mobile": "수령인 휴대전화",
    "recipient_zip_code": "우편번호",  # 수령인 우편번호 → 우편번호
    "product_name_with_option": "상품명(옵션)",
    "box_size": "박스",
}
