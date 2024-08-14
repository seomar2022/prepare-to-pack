import tkinter as tk
#.pack()은 부모위젯 안에 배치


# 메인 윈도우 생성
root = tk.Tk()
root.title("prepare_to_pack")
root.geometry("350x400")  # 너비x높이
root.configure()

font_size = 14

# 프레임을 사용하여 내부 여백 추가
frame = tk.Frame(root, padx=20, pady=20)
frame.pack()

label = tk.Label(root, text="상품포장준비 프로그램입니다!\n문의:seomar2022@gmail.com", font=font_size)
label.pack(pady=20)  # 여백 설정  # 위젯을 창에 배치

# 버튼 클릭 시 실행될 함수 정의
def prepare_to_pack():
    log_text.set("시작! 프로그램을 실행합니다.\n헤더를 찾았습니다.\n파일을 분리했습니다.\n중량을 지정했습니다.\n매크로를 실행했습니다.\n상품코드를 찾았습니다.\nPdf파일을 병합했습니다.\n프린트를 지정했습니다.\n프린트했습니다.\n완료! 송장을 등록해주세요.")

def match_cafe24_with_hanjin():
    log_text.set("송장을 업로드하는 기능입니다.\n아직 구현되지 않았습니다.")

# 버튼 프레임 생성
button_frame = tk.Frame(root)
button_frame.pack(pady=10)
prepare_image = tk.PhotoImage(file="resources/img/box.png")
#https://www.flaticon.com/free-icon/box_679720?term=packing&page=1&position=1&origin=search&related_id=679720

upload_image = tk.PhotoImage(file="resources/img/order-fulfillment.png")
#https://www.flaticon.com/free-icon/order-fulfillment_11482468?term=delivery+label&page=1&position=1&origin=search&related_id=11482468
# 포장 준비 버튼 추가
prepare_button = tk.Button(button_frame, image=prepare_image, command=prepare_to_pack)
prepare_button.pack(side="left", padx=10)

# 송장 업로드 버튼 추가
upload_button = tk.Button(button_frame, image=upload_image, command=match_cafe24_with_hanjin, font=font_size)
upload_button.pack(side="right", padx=10)

# 로그 텍스트를 표시할 라벨 생성
log_text = tk.StringVar()
log_text.set("")

log_label = tk.Label(root, textvariable=log_text, justify="left", anchor="nw")
log_label.pack(pady=10, padx=20)

# 메인 루프 시작
root.mainloop()