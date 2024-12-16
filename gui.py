# gui.py
import tkinter as tk
from tkinter import PhotoImage, StringVar
import threading
from module import *

class GUI:
    def __init__(self, root, on_before_packing, on_upload_tracking):
        self.root = root
        self.on_before_packing = on_before_packing  # 로직 함수 연결
        self.on_upload_tracking = on_upload_tracking  # 로직 함수 연결

        self.font_size = 14
        self.log_text = StringVar()  # 로그 텍스트를 관리하는 변수
        self.log_text.set("")  # 초기값 설정

        self.setup_ui()  # UI 생성

    def setup_ui(self):
        """GUI를 초기화하는 메서드"""
        self.root.title("Prepare to Pack")
        self.root.geometry("350x450")
        self.root.attributes('-topmost', True)

        #Add margin
        frame = tk.Frame(self.root, padx=20, pady=20)
        frame.pack()

        # 타이틀
        label = tk.Label(self.root, text="출고 준비 프로그램", font=("none", self.font_size, "bold"))
        label.pack(pady=10) # set title on root

        # 버튼 프레임 생성
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)

        # 버튼 이미지
        before_packing_image = PhotoImage(file="resources/img/package-box.png")
        upload_image = PhotoImage(file="resources/img/document.png")
        info_image = PhotoImage(file="resources/img/info.png")


        # 포장 준비 버튼
      # before_packing_button = tk.Button(button_frame, image=before_packing_image, command=on_before_packing_button_click)
        before_packing_button = tk.Button(button_frame, image=before_packing_image, command=lambda: threading.Thread(target=self.on_before_packing, args=(self.update_log, self.get_log)).start())
        before_packing_button.image = before_packing_image
        before_packing_button.pack(side="left", padx=10)
        ToolTip(before_packing_button, "cafe24에서 '출고준비통합'양식으로 파일을 다운로드 받은 후 이 버튼을 클릭해 주세요.")


        # 송장 업로드 버튼
        upload_tracking_number_button = tk.Button(button_frame, image=upload_image, command=lambda: threading.Thread(target=self.on_upload_tracking, args=(self.update_log)).start())
        upload_tracking_number_button.image = upload_image
        upload_tracking_number_button.pack(side="right", padx=10)
        ToolTip(upload_tracking_number_button, "한진택배에서 '원본파일'을 다운로드 받은 후 이 버튼을 클릭해 주세요.")
        # info 버튼
        info_button = tk.Button(self.root, image=info_image)
        info_button.image = info_image
        info_button.pack(side="bottom", pady=20)
        ToolTip(info_button, "-cafe24 엑셀파일 다운 양식 수정: settings\\header.csv\n-인터넷에서 다운 받은 파일이 있는 폴더 경로 지정: settings\\path.csv\n-설명지 추가: \\resources\\product_instruction\n-image: Flaticon.com\n-기타문의:seomar2022@gmail.com")



        # 로그 레이블
        log_label = tk.Label(self.root, textvariable=self.log_text, justify="left", anchor="nw", font=self.font_size, wraplength=300)
        log_label.pack(pady=10, padx=20)

    def update_log(self, message):
        """로그 텍스트를 업데이트하는 메서드"""
        self.log_text.set(message)
    def get_log(self):
        return self.log_text.get()
    

    # def on_before_packing_button_click():
    # # 별도의 스레드에서 프로그램 로직 실행
    #     threading.Thread(target=prepare_to_pack, daemon=True).start()
