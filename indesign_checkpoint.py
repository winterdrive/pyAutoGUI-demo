from typing import Tuple
import csv
import pandas as pd
import math
import requests
import json
from PyAutoGUIHelper import PyAutoGUIHelper

EXCEL_FILE_PATH = "D:\\Users\\Downloads\\2024電影曆_編輯內容.xlsx"
UNSPLASH_ACCESS_KEY = "bGW_6q4TAp_kKs1X3ZKD3F4wlOqY5pgHEPu92B0m4Uc"


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


my_calender_positions = {
    "月份": [1184, 271],
    "星期": [1365, 266],
    "農曆": [1367, 284],
    "倒數日": [1572, 269],
    "日期": [1317, 471],
    "宜忌": [1348, 604],
    "節日": [1331, 650],
    "歷史": [1325, 672],
    "語錄": [1324, 802],
    "作者": [1362, 965],
    "影視作品": [1356, 936]
}

my_calender_row_and_column = {
    "月份": [0, 0],
    "星期": [0, 1],
    "農曆": [0, 2],
    "倒數日": [0, 3],
    "日期": [0, 4],
    "宜忌": [0, 5],
    "節日": [0, 6],
    "歷史": [0, 7],
    "語錄": [0, 8],
    "作者": [0, 9],
    "影視作品": [0, 10],

}


def single_run(helper: PyAutoGUIHelper, value: str, position: Tuple[int, int]) -> None:
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


def main():
    # # 圖片下載
    # # https://unsplash.com/photos/3RIhxkIOBcE
    # get_unsplash_image("3RIhxkIOBcE")

    # # 查定位點
    # helper = PyAutoGUIHelper()
    # helper.show_mouse_position()



    # 自動化任務
    helper = PyAutoGUIHelper()

    # 讀取Excel資料
    excel_data = read_excel_file(EXCEL_FILE_PATH)

    # # 取得圖片下載網址
    from urllib.parse import urlparse, unquote
    photo_url = get_x_y_cell(excel_data, 0, 11)
    parsed_url = urlparse(photo_url)
    path = parsed_url.path 
    photo_id = path.split('/')[-1]
    # get_unsplash_image(photo_id)

    LOCAL_PHOTO_DIRECTORY = "E:\驅動程式\pyAutoGUI-demo-master\photo"

    # 塞入底圖
    helper.move_mouse_to(997,231).double_click_left()
    helper.move_mouse_to(1190,449).single_click_left()
    helper.move_mouse_to(997,65).single_click_left()
    helper.move_mouse_to(1108,399).double_click_left()
    helper.move_mouse_to(1069,260).single_click_left().paste_to_clipboard(LOCAL_PHOTO_DIRECTORY).paste().press_key('enter')
    helper.move_mouse_to(840,772).single_click_left().paste_to_clipboard(f"{photo_id}.jpg").paste().press_key('enter')

    # 填入各欄位
    helper.move_mouse_to(994, 271)
    helper.single_click_left()
    helper.move_mouse_to(1190,449).single_click_left()
    my_dict_keys = my_calender_row_and_column.keys()
    for key in my_dict_keys:
        # if key != "月份":
        #     return
        current_x = my_calender_positions[key][0]
        current_y = my_calender_positions[key][1]
        target_row = my_calender_row_and_column[key][0]
        target_column = my_calender_row_and_column[key][1]
        my_input = get_x_y_cell(excel_data, target_row, target_column)
        if type(my_input) is float and math.isnan(my_input):
            my_input = " "
        if key == "日期":
            single_run(helper, format_with_zeros(int(my_input)), (current_x, current_y))
        else:
            single_run(helper, str(my_input), (current_x, current_y))


if __name__ == "__main__":
    main()
