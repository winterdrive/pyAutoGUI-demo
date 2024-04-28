import os
import time
import datetime
import pyautogui
import pyperclip

"""
請確保已安裝 pyautogui：!pip install pyautogui，
並在調用 run_chatgpt 方法前，先將瀏覽器視窗移到右側，並將滑鼠移動到 ChatGPT 文字輸入位置
"""


class PyAutoGUIHelper:
    def __init__(self):
        # Add a delay to ensure the system is ready
        time.sleep(3)

    def run_chatgpt(self, start_position=(0, 0)):
        """
        執行 ChatGPT，請在開始位置設定為右側 ChatGPT 文字輸入位置
        """
        self.move_mouse_to(start_position[0], start_position[1])
        self.single_click_left()
        self.type_string("Tell me what ChatGPT can do for me")
        self.press_key('enter')
        return self

    def show_mouse_position(self):
        """
        顯示目前滑鼠位置
        """
        print(pyautogui.position())
        return self

    def drag_to(self, x, y, duration=1.5, button='left'):
        """
        將滑鼠拖曳至指定位置
        """
        pyautogui.dragTo(x, y, duration=duration, button=button)
        return self

    def move_mouse_to(self, x, y, duration=0.5):
        """
        將滑鼠移動至指定位置
        """
        pyautogui.moveTo(x, y, duration=duration)
        return self

    def get_mouse_position(self):
        """
        獲取目前滑鼠位置
        """
        print(pyautogui.position())
        return self

    def get_screen_size(self):
        """
        獲取螢幕尺寸
        """
        print(pyautogui.size())
        return self

    def get_pyautogui_version(self):
        """
        獲取 PyAutoGUI 版本
        """
        print(pyautogui.__version__)
        return self

    def type_string(self, string):
        """
        輸入指定的字串
        """
        pyautogui.typewrite(string)
        return self

    def press_key(self, key_name):
        """
        按下指定的按鍵
        """
        pyautogui.press(key_name)
        return self

    def single_click_right(self):
        """
        單次點擊右鍵
        """
        pyautogui.click(button='right')
        return self

    def single_click_left(self):
        """
        單次點擊左鍵
        """
        pyautogui.click()
        return self

    def double_click_right(self):
        """
        雙次點擊右鍵
        """
        pyautogui.click(clicks=2, interval=0.5, button='right')
        return self

    def double_click_left(self):
        """
        雙次點擊左鍵
        """
        pyautogui.click(clicks=2, interval=0.5)
        return self

    def open_task_manager(self):
        """
        開啟工作管理員
        """
        pyautogui.hotkey('ctrl', 'shift', 'esc')
        return self

    def create_a_screen_shot(self):
        """
        截圖並儲存
        """
        os.makedirs('screenshot', exist_ok=True)
        now = datetime.datetime.now().strftime('%Y_%m_%d_%H_%M_%S')  # Format the datetime string
        filename = f'screenshot/{now}.png'
        pyautogui.screenshot(filename)
        return self

    def copy(self):
        pyautogui.hotkey('ctrl', 'c')
        return self

    def paste(self):
        pyautogui.hotkey('ctrl', 'v')
        return self

    def select_all(self):
        pyautogui.hotkey('ctrl', 'a')
        return self

    def paste_to_clipboard(self, string: str):
        string_to_paste = string
        pyperclip.copy(string_to_paste)
        return self

    def __str__(self):
        """Returns a string representation of the class."""
        return "PyAutoGUIHelper: A helper class for automating tasks using PyAutoGUI."


if __name__ == '__main__':
    """
    請先打開並登入 ChatGPT，並將瀏覽器視窗移到右側，
    如果您的 ChatGPT 文字輸入位置不在 (1202, 965)，請修改 start_position 參數，
    或使用 show_mouse_position 方法獲取目前滑鼠位置
    """
    helper = PyAutoGUIHelper()
    # 使用 ChatGPT 文字輸入位置 (1202, 965)
    helper.show_mouse_position()
    # .run_chatgpt(start_position=(1202, 965)).create_a_screen_shot()
