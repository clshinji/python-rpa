import pyautogui as pag
import pyperclip
from time import sleep
import sys
import os


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def find_image_coordinate(image_path, time_out, confidence):
    """
    プライマリスクリーンから任意の画像の座標を検索する
    :param image_path: 画面から見つけたい画像のパス。
    :param time_out: 画像検索待ち時間(秒)。
    :param confidence: 信頼感。数値が高いほど信頼できるが、見つからない確率も上がる。(0~1?)
    :return: 画像の座標。
    """
    for _ in range(time_out):
        locate = pag.locateOnScreen(image_path, confidence=confidence)
        if locate is not None:
            return locate
        else:
            sleep(1)


def move_click(button):
# if __name__ == '__main__':
    """ マウスを移動してクリックする
    Args:
        button (str): マウス移動先のターゲット画像パス
    """
    try:
        button_locate = find_image_coordinate(image_path=button, time_out=10, confidence=0.9)
        x, y = pag.center(button_locate)
        pag.moveTo(x, y)
        pag.click()
    except TypeError:
        print('Image not found')


def move_click_try2(button1, button2):
# if __name__ == '__main__':
    """ マウスを移動してクリックする（うまくいかないとき用の2つ目の画像を指定する場合）
    Args:
        button1 (str): マウス移動先のターゲット画像パス1
        button2 (str): マウス移動先のターゲット画像パス2
    """
    try:
        button_locate = find_image_coordinate(image_path=button1, time_out=10, confidence=0.9)
        x, y = pag.center(button_locate)
        pag.moveTo(x, y)
        pag.click()
    except Exception as e:
        print(f"Image not found -> Try 2nd Image [Image path]{button1}")
        try:
            button_locate = find_image_coordinate(image_path=button2, time_out=10, confidence=0.9)
            x, y = pag.center(button_locate)
            pag.moveTo(x, y)
            pag.click()
        except TypeError:
            print('Image not found')


def analyze_check_overwrite(button):
    OL_CHECK = False
    try:
        pag.locateOnScreen(image_path=button, time_out=10, confidence=0.9)
        OL_CHECK = True
    except Exception as e:
        print("Over write window -> Not Found")
    return OL_CHECK


def move_click_cal_textarea(button, x_init, y_init):
# if __name__ == '__main__':
    """ マウスを移動してクリックする（PD-Cal-f0用）
    Args:
        button (str): マウス移動先のターゲット画像パス
    """
    try:
        button_locate = find_image_coordinate(image_path=button, time_out=10, confidence=0.9)
        x, y = pag.center(button_locate)
        print(f"x:{x}, y:{y}")
        pag.moveTo(x + x_init, y + y_init)
        pag.click()
    except TypeError:
        print('Image not found')


def move_click_try2_cal_textarea(button1, button2, x_init1, y_init1, x_init2, y_init2):
# if __name__ == '__main__':
    """ マウスを移動してクリックする（PD-Cal-f0用）
    Args:
        button (str): マウス移動先のターゲット画像パス
    """
    try:
        button_locate = find_image_coordinate(image_path=button1, time_out=10, confidence=0.9)
        x, y = pag.center(button_locate)
        print(f"x:{x}, y:{y}")
        pag.moveTo(x + x_init1, y + y_init1)
        pag.click()
    except Exception as e:
        print(f"Image not found -> Try 2nd Image [Image path]{button1}")
        try:
            button_locate = find_image_coordinate(image_path=button2, time_out=10, confidence=0.9)
            x, y = pag.center(button_locate)
            print(f"x:{x}, y:{y}")
            pag.moveTo(x + x_init2, y + y_init2)
            pag.click()
        except TypeError:
            print('Image not found')


def input_text(value):
    """ テキストフォームに変数の値を入力する
    Args:
        value (any): 入力したい値
    """
    pyperclip.copy(value)
    pag.hotkey("ctrl", "v")


def key_bkspace(num):
    """backspaceを num 回押す
    Args:
        num (int): キーを押す回数
    """
    pag.press('backspace', presses=num)


def key_del(num):
    """delを num 回押す
    Args:
        num (int): キーを押す回数
    """
    pag.press('del', presses=num)


def key_enter(num):
    """Enterを num 回押す
    Args:
        num (int): キーを押す回数
    """
    pag.press('enter', presses=num)


def key_CtrlA2del():
    """Ctrl + A を押した後に del を押す
    """
    with pag.hold('ctrl'):
        pag.press('A')
    pag.press('del')


def change_f0(f0_SET):
    """CAL->f0を変更する
    Args:
        f0_SET (int): 変更後のf0(kHz)
    """
    move_click_cal_textarea(resource_path('assets\\pd_cal_f0_textarea.png'), 70, 340)
    # move_click_cal_textarea('assets\\pd_cal_f0_textarea.png', 70, 340)
    sleep(3)
    key_bkspace(9)
    input_text(f0_SET)    # BW変更のために200kHzに一時的に設定
    key_enter(1)
    pag.click()
    sleep(2)