from src.gui import GUI
from src.logic.prepare_to_pack import prepare_to_pack
from src.logic.upload_tracking_number import on_upload_tracking_number_button_click

import logging
logging.basicConfig(level=logging.INFO)

logging.info("main.py started")

# 메인 윈도우 생성
if __name__ == "__main__":
    import tkinter as tk

    root = tk.Tk()
    gui = GUI(root, prepare_to_pack, on_upload_tracking_number_button_click)

    # 메인 루프 시작
    root.mainloop()
