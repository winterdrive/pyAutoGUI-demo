# # # # 讀取Excel資料
import pandas as pd
from config.lulumi_config import EXCEL_FILE_PATH

def read_excel_file(file_path: str) -> pd.DataFrame:
    """Reads an Excel file into a DataFrame."""
    df = pd.read_excel(file_path)
    return df


def main():
    excel_data = read_excel_file(EXCEL_FILE_PATH)
    # print(excel_data.head())
    # get the first three rows
    print(excel_data.head(9))
    # get the first three rows of the "月份" column
    print(excel_data["月份"].head(9),
            # excel_data["日期"].head(9),
            # excel_data["星期"].head(9),
            # excel_data["%"].head(9),
            excel_data["bar長度(公釐)"].head(9),
            # excel_data["已過%"].head(9),
            excel_data["語錄（中文）"].head(9),
            # excel_data["著作"].head(9),
            # excel_data["佚名"].head(9),
          )


if __name__ == "__main__":
    main()
