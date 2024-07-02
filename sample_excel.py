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
    
    # エクセルを起動するだけのサンプル
    # クリック対象の画像をキャプチャして assets ディレクトリに入れておく
    app = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
    subprocess.Popen(app)
    sleep(7)
    pag_tools.move_click_try2('assets\excel_btn_window_max_active.png', 'assets\excel_btn_window_max_deactive.png')
    sleep(1)
    pag_tools.move_click('assets\excel_btn_window_min.png')
    sleep(1)
    pag_tools.move_click('assets\excel_btn_window_close.png')
    
    print("自動処理が完了しました")


if __name__ == "__main__":
    main()
