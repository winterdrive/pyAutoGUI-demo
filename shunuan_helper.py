from typing import Tuple
import csv
import pandas as pd
import math

from PyAutoGUIHelper import PyAutoGUIHelper
from config.shunuan_config_07_12 import EXCEL_FILE_PATH, CALENDER_POSITIONS, CALENDER_ROW_AND_COLUMN, LOCAL_PHOTO_DIRECTORY, MY_INDESIGN_BUTTON_POSITIONS
from indesign_get_upsplash_image import get_unsplash_image


def read_csv_file(file_path: str) -> None:
    """Reads a CSV file and prints each row."""
    with open(file_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            print(row)


def read_excel_file(file_path: str) -> pd.DataFrame:
    """Reads an Excel file into a DataFrame."""
    df = pd.read_excel(file_path)
    return df


def get_x_y_cell(df: pd.DataFrame, row_number: int, column_number: int):
    """Gets the value at a specific row and column in a DataFrame."""
    return df.iloc[row_number, column_number]



def replace_cell_text(helper: PyAutoGUIHelper, value: str, position: Tuple[int, int]) -> None:
    """Performs a series of actions in a GUI."""
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["左側工具列_文字工具"]).single_click_left()
    x, y = position
    helper.move_mouse_to(x, y, 0.2)
    helper.double_click_left()
    helper.select_all()
    helper.paste_to_clipboard(value)
    helper.paste()


def format_with_zeros(num: int) -> str:
    """Formats a number with leading zeros."""
    return f"{num:02}"


def create_new_page(helper: PyAutoGUIHelper, is_multiple_page:bool=False) -> None:
    """Creates a new page in InDesign."""
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["右側工具列_頁面"]).single_click_left()
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["右側工具列_頁面_頁面清單第一項"]).single_click_right()
    if is_multiple_page:
        helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["右側工具列_頁面_頁面清單第一項_複製跨頁_多頁情境下"]).single_click_left()
    else:
        helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["右側工具列_頁面_頁面清單第一項_複製跨頁"]).single_click_left()
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["右側工具列_頁面_收起頁面"]).single_click_left()


def auto_scaling_workspace(helper: PyAutoGUIHelper) -> None:
    """Automatically scales the workspace in InDesign."""
    # # # 運行前不得選取任何文字框或物件 # # #
    # 自動標準化工作區域位置
    # 1. CTRL + ALT + SHIFT + 0：整個作業範圍
    # 2. CTRL + 1 ：實際大小
    # 3. CTRL + = ：放大
    # 4. SCROLL DOWN：滑鼠滾輪向下
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["左側工具列_手形工具"]).double_click_left()
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["展示區任意點"]).single_click_left()
    helper.use_any_hotkey(['ctrl', 'alt', 'shift', '0']) # 整個作業範圍
    helper.use_any_hotkey(['ctrl', '1']) # 實際大小
    helper.use_any_hotkey(['ctrl', '=']) # 放大3RIhxkIOBcE.jpg
    # helper.scroll_down() # 如果是兩個以上的跨頁，需要向下滾動
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["展示區任意點"])




def single_run(helper: PyAutoGUIHelper, excel_data: pd.DataFrame, data_row_number: int = 0, is_multiple_page:bool = False):

    try:
        # # 取得圖片下載網址
        from urllib.parse import urlparse, unquote
        photo_url = get_x_y_cell(excel_data, data_row_number, 10)
        print(photo_url)
        parsed_url = urlparse(photo_url)
        path = parsed_url.path 
        photo_id = path.split('/')[-1].split('-')[-1]
        print(photo_id)
        get_unsplash_image(photo_id, LOCAL_PHOTO_DIRECTORY)

        # 塞入底圖
        helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["左側工具列_選取工具"]).single_click_left()
        helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["展示區任意點"]).single_click_left()
        helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["上方工具列_檔案"]).single_click_left()
        helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["檔案_置入"]).double_click_left()
        helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["檔案_置入_檔案路徑"]).single_click_left().paste_to_clipboard(LOCAL_PHOTO_DIRECTORY).paste().press_key('enter')
        helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["檔案_置入_檔名"]).single_click_left().paste_to_clipboard(f"{photo_id}.jpg").paste().press_key('enter')
        helper.use_any_hotkey(['ctrl', 'alt', 'shift', 'c']) # 將圖片等比例填滿框架
    except Exception as e:
        print(e)
        print("No photo to download")

    # return # phote test only

    # 填入各欄位
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["左側工具列_直接選取工具"])
    helper.single_click_left()
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["展示區任意點"]).single_click_left()
    my_dict_keys = CALENDER_ROW_AND_COLUMN.keys()
    for key in my_dict_keys:
        # if key != "月份":
        #     return
        current_x = CALENDER_POSITIONS[key][0]
        current_y = CALENDER_POSITIONS[key][1]
        target_row = data_row_number
        target_column = CALENDER_ROW_AND_COLUMN[key][1]
        my_input = get_x_y_cell(excel_data, target_row, target_column)
        if type(my_input) is float and math.isnan(my_input):
            my_input = " "
        if key == "日期":
            replace_cell_text(helper, format_with_zeros(int(my_input)), (current_x, current_y))
        else:
            replace_cell_text(helper, str(my_input), (current_x, current_y))

    # 保留截圖以供下次檢查定位點
    helper.create_a_screen_shot()

    # 建立新圖給下一張操作
    create_new_page(helper=helper, is_multiple_page=is_multiple_page)


def main():
    # # 圖片下載
    # # https://unsplash.com/photos/3RIhxkIOBcE
    # get_unsplash_image("3RIhxkIOBcE")

    # # 運行任務
    # # # 建立自動化任務工具
    helper = PyAutoGUIHelper()

    # # # # 自動標準化工作區域位置
    # auto_scaling_workspace(helper)

    # # # 讀取Excel資料
    excel_data = read_excel_file(EXCEL_FILE_PATH)

    # # # 迴圈運行任務
    todo_rows = range(0, 366)
    for i in todo_rows:

        # # # # 自動標準化工作區域位置
        helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["左側工具列_文字工具"]).single_click_left()
        helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["頂部任意點"]).single_click_left()
        helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["左側工具列_手形工具"]).double_click_left()
        for j in range(0, 10):
            helper.scroll_down()
        helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["左側工具列_手形工具"]).double_click_left()
        helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["底部任意點"]).double_click_left()
        helper.use_any_hotkey(['ctrl', 'alt', 'shift', '0']) # 整個作業範圍
        helper.use_any_hotkey(['ctrl', '1']) # 實際大小
        helper.use_any_hotkey(['ctrl', 'alt', 'shift', '0']) # 整個作業範圍
        helper.use_any_hotkey(['ctrl', '1']) # 實際大小

        if i == 0:
            single_run(helper=helper, excel_data=excel_data, data_row_number=i, is_multiple_page=False)
        else:
            single_run(helper=helper, excel_data=excel_data, data_row_number=i, is_multiple_page=True)
        print(f"Excel內的第{i}筆資料已完成")


if __name__ == "__main__":
    main() # 84% 放大倍率，向下6滾輪
