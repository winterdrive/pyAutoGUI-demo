# 文字資料
import os
from dotenv import load_dotenv


EXCEL_FILE_PATH = "D:\\Users\\Downloads\\2024電影曆_編輯內容.xlsx"

# Unsplash API
load_dotenv()
UNSPLASH_ACCESS_KEY = os.getenv('UNSPLASH_ACCESS_KEY')

# 存放背景圖片的資料夾
LOCAL_PHOTO_DIRECTORY = "E:\pyAutoGUI-demo\photo"

# my_calender_positions
CALENDER_POSITIONS= {
    "月份": [1147, 274], 
    "星期": [1340, 268],
    "農曆": [1340, 292], 
    "倒數日": [1539, 280],
    "日期": [1340, 478],
    "宜忌": [1340, 605],
    "節日": [1340, 656],
    "歷史": [1340, 681],
    "語錄": [1340, 802],
    "作者": [1340, 972],
    "影視作品": [1340, 949]
}

# my_calender_row_and_column
CALENDER_ROW_AND_COLUMN = {
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

# my_indesign_button_positions
MY_INDESIGN_BUTTON_POSITIONS = {
    "左側工具列_選取工具": [997, 231],
    "左側工具列_手形工具": [994, 928],
    "左側工具列_文字工具": [992, 418],
    "展示區任意點": [1190, 449],
    "上方工具列_檔案": [997, 65],
    "檔案_置入": [1108, 399],
    "檔案_置入_檔案路徑": [826,239], # 視窗位置會變動
    "檔案_置入_檔名": [731, 833], # 視窗位置會變動
    "左側工具列_直接選取工具": [994, 271],
    "右側工具列_頁面": [1746, 232],
    "右側工具列_頁面_頁面清單第一項": [1300, 374],
    "右側工具列_頁面_頁面清單第一項_複製跨頁": [1388, 417], 
    "右側工具列_頁面_頁面清單第一項_複製跨頁_多頁情境下": [1388, 443],
    "右側工具列_頁面_收起頁面": [1594, 216],
    "底部任意點": [1329, 911],
    "頂部任意點": [1321, 280],
    "頂部長寬工具的寬度" : [1310, 114],
    "頂部長寬工具的九宮格左中" : [1014,132],
    "頂部長寬工具的任意位置" : [1521,111],
    "置中區域" : [1341,616],
    }