from time import sleep
import subprocess
import pyautogui as pag
import src.pag_tools as pag_tools


def main():
    # 実行前の確認
    pag_confirm = pag.confirm(text='自動処理を開始しますか？', title='自動処理実行前の確認', buttons=['OK', 'Cancel'])
    if pag_confirm == "Cancel":
        print("Cancelボタンが押されました。処理を中止します。")
        return
    
    # 電卓を起動して足し算するだけのサンプル
    # クリック対象の画像をキャプチャして assets ディレクトリに入れておく
    app = r"calc.exe"
    subprocess.Popen(app)
    sleep(5)
    pag_tools.move_click_try2('assets\calc_1_dark.png', 'assets\calc_1_dark_mo.png')
    sleep(1)
    pag_tools.move_click_try2('assets\calc_plus_dark.png', 'assets\calc_plus_dark_mo.png')
    sleep(1)
    pag_tools.move_click_try2('assets\calc_1_dark.png', 'assets\calc_1_dark_mo.png')
    sleep(1)
    pag_tools.move_click_try2('assets\calc_eq_dark.png', 'assets\calc_eq_dark_mo.png')
    sleep(5)
    pag_tools.move_click('assets\calc_close.png')
    
    print("自動処理が完了しました")


if __name__ == "__main__":
    main()
