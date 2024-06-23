
from PyAutoGUIHelper import PyAutoGUIHelper
from config.base_config import MY_INDESIGN_BUTTON_POSITIONS
from config.lulumi_config import LOCAL_PHOTO_DIRECTORY


def scroll_down_in_indesign_by_pixel(helper: PyAutoGUIHelper,pixels=100) -> None:
    """Automatically scales the workspace in InDesign."""
    helper.move_mouse_to(1614, 275).double_click_left()
    helper.scroll_down(pixels)



def replace_image(helper: PyAutoGUIHelper, photo_id: str) -> None:
    # 塞入底圖
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["左側工具列_選取工具"]).double_click_left()
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["展示區任意點"]).single_click_left()
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["上方工具列_檔案"]).single_click_left()
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["檔案_置入"]).double_click_left()
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["檔案_置入_檔案路徑"]).single_click_left().paste_to_clipboard(LOCAL_PHOTO_DIRECTORY).paste().press_key('enter')
    helper.move_mouse_to(*MY_INDESIGN_BUTTON_POSITIONS["檔案_置入_檔名"]).single_click_left().paste_to_clipboard(f"{photo_id}.jpg").paste().press_key('enter')
   

def check_file_exists(image_path: str) -> bool:
    """Check if the image file exists."""
    import os
    return os.path.exists(image_path)


def main():
    # # 運行任務
    # # # 建立自動化任務工具
    helper = PyAutoGUIHelper()

    # # # 迴圈運行任務
    for i in range(1, 366):
        # 格式化為三位數字
        photo_id = f"{i:03d}"  

        ### 檢查圖片是否存在
        # 如果圖片不存在，則跳過
        image_path = f"{LOCAL_PHOTO_DIRECTORY}/{photo_id}.jpg"
        if check_file_exists(image_path):
            replace_image(helper=helper, photo_id=photo_id)
        else:
            print(f"圖片不存在: {image_path}")

        # 自動拉至下一page (放大比例75%)
        scroll_down_in_indesign_by_pixel(helper=helper, pixels=-2400)

        if i % 2 == 0:
            # 每10張圖片，校正回一次 (放大比例75%)
            scroll_down_in_indesign_by_pixel(helper=helper, pixels=200)



if __name__ == "__main__":
    main()
