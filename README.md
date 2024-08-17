- 전 채널 주문리스트(manage-online-store), 송장번호 관리(manage-tracking-number), 설명지 인쇄 프로그램(print-out-product-instruction)을 통합한 프로그램
- 카페24와 한진택배 사이트
- 프로그램 도입 전후 작업 과정


# 작동 순서
- 카페24에서 다운로드 받은 파일을 찾아서, settings/header.csv에 따라 두 개의 파일로 나눈다.

### 주문리스트 파일
- 주문리스트 파일에서 상품명과 옵션을 읽고 중량 정보를 기입한다.
- 상품 포장할 때 보기 편하도록 전채널주문리스트 매크로를 실행한다.

### 한진 파일
- 복수내품 형식에 맞추도록 ProcessMultipleItems매크로를 실행한다. 
- 카페24에 송장 정보를 일괄 등록할 수 있도록 조정한 파일을 만들고 저장한다. 
