import subprocess
import tkinter as tk
from tkinter import Text, Scrollbar, END, filedialog, messagebox
import threading
from src.logic.config import save_download_path, get_instruction_folder
from src.logic.module import resource_path


class GUI:
    def __init__(self, root, on_before_packing, on_upload_tracking):
        self.root = root
        self.on_before_packing = on_before_packing
        self.on_upload_tracking = on_upload_tracking

        self.font_size = 14
        self.setup_ui()

    def setup_ui(self):
        """Initialize the GUI layout"""
        self.root.title("LALA Pet Mall - 출고 준비 프로그램")
        self.root.geometry("400x600")
        self.root.configure(bg="#fdfaf4")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        self.root.iconbitmap(resource_path("resources/img/favicon.ico"))

        # Main Centering Frame
        main_frame = tk.Frame(self.root, bg="#fdfaf4")
        main_frame.pack(fill="both", expand=True)

        ########################################## Title Frame ##########################################
        title_frame = tk.Frame(main_frame, bg="#fdfaf4")
        title_frame.pack(pady=(20, 10), anchor="center")

        ### title
        title = tk.Label(
            title_frame,
            text="LALA Pet Mall 출고 준비 프로그램",
            font=("Noto Sans KR", self.font_size + 2, "bold"),
            bg="#fdfaf4",
            fg="#2d4831",
        )
        title.pack(side="left")

        ### Settings
        self.gear_icon = tk.PhotoImage(file=resource_path("resources/img/gear.png"))
        settings_button = tk.Button(
            title_frame,
            image=self.gear_icon,
            command=self.set_download_folder,
            bg="#fdfaf4",
            bd=0,
            activebackground="#fdfaf4",
            cursor="hand2",
        )
        settings_button.pack(side="left", padx=5)

        ### product_instruction
        instruction_folder = get_instruction_folder()

        self.document_icon = tk.PhotoImage(
            file=resource_path("resources/img/document.png")
        )
        instruction_button = tk.Button(
            title_frame,
            image=self.document_icon,
            command=lambda: subprocess.Popen(f'explorer "{instruction_folder}"'),
            bg="#fdfaf4",
            bd=0,
            activebackground="#fdfaf4",
            cursor="hand2",
        )
        instruction_button.pack(side="left", padx=5)

        ########################################## Step 1 Frame ##########################################
        step1_frame = tk.Frame(main_frame, bg="#fdfaf4")
        step1_frame.pack(pady=(10, 5))

        tk.Label(
            step1_frame,
            text="[Step 1] 주문서 정리하기",
            font=("Noto Sans KR", self.font_size, "bold"),
            bg="#fdfaf4",
            fg="#2d4831",
        ).pack()

        tk.Label(
            step1_frame,
            text="cafe24에서 '출고준비통합'양식으로 파일을 다운로드 받은 후 아래 버튼을 클릭해 주세요.",
            font=("Noto Sans KR", self.font_size - 3),
            bg="#fdfaf4",
            wraplength=350,
        ).pack(pady=(0, 10))

        tk.Button(
            step1_frame,
            text="주문서 정리",
            command=lambda: threading.Thread(
                target=self.on_before_packing, args=(self.append_log, self.get_log)
            ).start(),
            font=("Noto Sans KR", self.font_size - 1, "bold"),
            bg="#4f785c",
            fg="white",
            width=25,
            height=2,
            bd=0,
            activebackground="#3a5c46",
            cursor="hand2",
        ).pack()

        ########################################## Step 2 Frame ##########################################
        step2_frame = tk.Frame(main_frame, bg="#fdfaf4")
        step2_frame.pack(pady=(20, 5))

        tk.Label(
            step2_frame,
            text="[Step 2] 송장번호 입력 및 업로드",
            font=("Noto Sans KR", self.font_size, "bold"),
            bg="#fdfaf4",
            fg="#2d4831",
        ).pack()

        tk.Label(
            step2_frame,
            text="한진택배에서 '원본파일'을 다운로드 받은 후 아래 버튼을 클릭해 주세요.",
            font=("Noto Sans KR", self.font_size - 3),
            bg="#fdfaf4",
            wraplength=350,
        ).pack(pady=(0, 10))

        tk.Button(
            step2_frame,
            text="송장 업로드하기",
            command=lambda: threading.Thread(
                target=self.on_upload_tracking, args=(self.append_log,)
            ).start(),
            font=("Noto Sans KR", self.font_size - 1, "bold"),
            bg="#4f785c",
            fg="white",
            width=25,
            height=2,
            bd=0,
            activebackground="#3a5c46",
            cursor="hand2",
        ).pack()

        ########################################## Log Frame ##########################################
        log_frame = tk.Frame(main_frame, bg="#fdfaf4")
        log_frame.pack(pady=(20, 10), padx=20, fill="both", expand=True)

        tk.Label(
            log_frame,
            text="실시간 작업 로그",
            font=("Noto Sans KR", self.font_size - 1, "bold"),
            bg="#fdfaf4",
            fg="#2d4831",
        ).pack(anchor="w")

        text_frame = tk.Frame(log_frame)
        text_frame.pack(fill="both", expand=True)

        self.log_widget = Text(
            text_frame,
            font=("Noto Sans KR", self.font_size - 2),
            height=8,
            wrap="word",
            bg="#ffffff",
            bd=1,
            relief="solid",
        )
        self.log_widget.pack(side="left", fill="both", expand=True)

        scrollbar = Scrollbar(text_frame, command=self.log_widget.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_widget.config(yscrollcommand=scrollbar.set)

    def append_log(self, message):
        self.log_widget.insert(END, message + "\n")
        self.log_widget.see(END)

    def get_log(self):
        return self.log_widget.get("1.0", END)

    def set_download_folder(self):
        folder = filedialog.askdirectory(title="다운로드 폴더 선택")
        if folder:
            save_download_path(folder)
            messagebox.showinfo(
                "저장됨", f"다운로드 폴더:\n{folder} 로 설정되었습니다."
            )
