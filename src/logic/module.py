import pandas as pd
import os
import csv
import tkinter as tk


####설정폴더에서 경로찾기
def search_path(header_name):
    try:
        # CSV 파일 열기
        with open("settings\\path.csv", mode="r", encoding="utf-8-sig") as file:
            reader = csv.reader(file)

            # 데이터 검색
            for row in reader:
                if row[0].strip() == header_name:
                    return os.path.expanduser(row[1].strip())

            print(f"헤더 '{header_name}'을(를) 찾을 수 없습니다.")
            return ""

    except FileNotFoundError:
        print("설정 파일을 찾을 수 없습니다")
        return ""
    except Exception as e:
        print(f"설정 파일을 읽는 중 오류가 발생했습니다: {e}")
        return ""


####이름 일부를 검색해서 파일찾기
def find_path_by_partial_name(directory, partial_name):
    # 지정된 디렉토리의 파일 및 디렉토리 목록을 가져옴
    items = os.listdir(directory)

    # 부분 문자열이 파일 또는 디렉토리 이름에 포함된 항목 목록 생성
    matching_items = [item for item in items if partial_name in item]

    # 매칭된 항목이 없으면 None 반환
    if not matching_items:
        return None

    # 가장 최근에 수정된 파일 또는 디렉토리 찾기
    most_recent_item = max(
        matching_items, key=lambda f: os.path.getmtime(os.path.join(directory, f))
    )

    # 가장 최근에 수정된 파일 또는 디렉토리의 전체 경로 반환
    return os.path.join(directory, most_recent_item)


####setting\header.csv에서 데이터가져오기위해 만듦.
def get_column_from_csv(file_path, column_name):
    """
    CSV 파일에서 특정 열의 데이터를 가져옵니다.

    Args:
        file_path (str): CSV 파일 경로
        column_name (str): 가져올 열 이름

    Returns:
        pd.Series: 해당 열의 데이터 시리즈
    """
    try:
        # CSV 파일 읽기
        df = pd.read_csv(file_path, encoding="utf-8")

        # 해당 열 가져오기
        if column_name in df.columns:
            return df[column_name].dropna()
        else:
            print(f"'{column_name}' 열을 찾을 수 없습니다.")
            return None

    except FileNotFoundError:
        print(f"CSV 파일을 찾을 수 없습니다: {file_path}")
        return None
    except Exception as e:
        print(f"파일을 읽는 중 오류가 발생했습니다: {e}")
        return None


def split_csv_by_column_index(csv_file_path, excel_file_path, column_indices):
    # column_indices를 list로 넣어도 됨.
    try:
        # CSV 파일 읽기
        df = pd.read_csv(csv_file_path, encoding="utf-8")

        # 특정 인덱스의 열만 선택
        selected_columns = df.iloc[:, column_indices]

        # 선택한 열을 새로운 Excel 파일로 저장
        selected_columns.to_excel(excel_file_path, index=False)
        print(f"선택한 열이 성공적으로 '{excel_file_path}'에 저장되었습니다.")

    except FileNotFoundError:
        print(f"CSV 파일을 찾을 수 없습니다: {csv_file_path}")
    except Exception as e:
        print(f"파일을 처리하는 중 오류가 발생했습니다: {e}")


###GUI 툴팁
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event=None):
        if self.tooltip_window or not self.text:
            return
        x, y, _cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + cy + 25
        self.tooltip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)  # 창 프레임 제거
        tw.wm_geometry(f"+{x}+{y}")
        tw.attributes("-topmost", True)

        label = tk.Label(
            tw,
            text=self.text,
            justify="left",
            relief="solid",
            borderwidth=1,
            font=("Arial", 10, "normal"),
        )
        label.pack(ipadx=1)

    def hide_tooltip(self, event=None):
        tw = self.tooltip_window
        self.tooltip_window = None
        if tw:
            tw.destroy()
