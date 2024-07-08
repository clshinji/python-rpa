import pyautogui as pag
from time import sleep


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
