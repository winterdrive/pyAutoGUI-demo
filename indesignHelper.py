from typing import Tuple
import csv
import pandas as pd
import math
import requests
import json
from PyAutoGUIHelper import PyAutoGUIHelper
from config import UNSPLASH_ACCESS_KEY, EXCEL_FILE_PATH, CALENDER_POSITIONS, CALENDER_ROW_AND_COLUMN, LOCAL_PHOTO_DIRECTORY, MY_INDESIGN_BUTTON_POSITIONS

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
    x, y = position
    helper.move_mouse_to(x, y, 0.2)
    helper.double_click_left()
    helper.select_all()
    helper.paste_to_clipboard(value)
    helper.paste()


def format_with_zeros(num: int) -> str:
    """Formats a number with leading zeros."""
    return f"{num:02}"


def get_unsplash_image(photo_id: str) -> None:
    """Downloads an image from Unsplash."""
    url = f"https://api.unsplash.com/photos/{photo_id}?client_id={UNSPLASH_ACCESS_KEY}"
    # print(url)
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.82 Safari/537'}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        # print(response.text)
        data = json.loads(response.text)
        # print(data)
        download_url = data['urls']['raw']
        with open(f"./photo/{photo_id}.jpg", "wb") as file:
            res = requests.get(download_url, stream=True, headers=headers)
            file.write(res.content)
    else:
        print('Failed to download the image.')


def create_new_page(helper: PyAutoGUIHelper) -> None:
    """Creates a new page in InDesign."""
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["右側工具列_頁面"]).single_click_left()
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["右側工具列_頁面_頁面清單第一項"]).single_click_right()
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["右側工具列_頁面_頁面清單第一項_複製跨頁"]).single_click_left()
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["右側工具列_頁面_收起頁面"]).single_click_left()


def single_run(helper: PyAutoGUIHelper, excel_data: pd.DataFrame, data_row_number: int = 0):

    # # 取得圖片下載網址
    from urllib.parse import urlparse, unquote
    photo_url = get_x_y_cell(excel_data, data_row_number, 11)
    parsed_url = urlparse(photo_url)
    path = parsed_url.path 
    photo_id = path.split('/')[-1]
    # get_unsplash_image(photo_id)

    # 塞入底圖
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["左側工具列_選取工具"]).double_click_left()
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["展示區任意點"]).single_click_left()
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["上方工具列_檔案"]).single_click_left()
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["檔案_置入"]).double_click_left()
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["檔案_置入_檔案路徑"]).single_click_left().paste_to_clipboard(LOCAL_PHOTO_DIRECTORY).paste().press_key('enter')
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["檔案_置入_檔名"]).single_click_left().paste_to_clipboard(f"{photo_id}.jpg").paste().press_key('enter')
   
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
    create_new_page(helper=helper)


def main():
    # # 圖片下載
    # # https://unsplash.com/photos/3RIhxkIOBcE
    # get_unsplash_image("3RIhxkIOBcE")

    # # 查定位點
    # helper = PyAutoGUIHelper()
    # helper.show_mouse_position()
    
    # # 運行任務
    # 自動化任務
    helper = PyAutoGUIHelper()

    # 讀取Excel資料
    excel_data = read_excel_file(EXCEL_FILE_PATH)

    todo_rows = range(0, 1)

    for i in todo_rows:
        single_run(helper=helper, excel_data=excel_data, data_row_number=i)


if __name__ == "__main__":
    main()
