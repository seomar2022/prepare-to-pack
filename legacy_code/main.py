from gui import GUI
from prepare_to_pack import *
from upload_tracking_number import *

# 메인 윈도우 생성
if __name__ == "__main__":
    import tkinter as tk  

    root = tk.Tk()
    
    gui = GUI(root, prepare_to_pack, on_upload_tracking_number_button_click)

    # 메인 루프 시작
    root.mainloop()

#pyinstaller --onefile --noconsole prepare_to_pack.py