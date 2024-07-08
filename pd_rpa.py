from time import sleep
import subprocess
import pyautogui as pag
import src.pag_tools as pag_tools
import src.helpers as helpers


def main():
    # 実行前の確認
    pag_confirm = pag.confirm(text='自動処理を開始しますか？\n👇の条件になっていることを確認してください\n(初期条件)f0=50kHz, BW=50kHz', title='自動処理実行前の確認', buttons=['OK', 'Cancel'])
    if pag_confirm == "Cancel":
        print("Cancelボタンが押されました。処理を中止します。")
        return

    # 共通設定
    MEAS_TIME = 5    # 一時的に5に変更、後で60に戻す
    # CSVから測定条件を読み込む
    MEAS_LIST = helpers.getMeasList("部分放電測定を行うフィルタの組み合わせ.csv")

    # 部分放電測定器用のRPA
    # クリック対象の画像をキャプチャして assets ディレクトリに入れておく

    # アプリを起動する
    app = r"C:\SOKEN\PD9\SampleSoftware260\PD9.exe"
    subprocess.Popen(app)
    sleep(7)

    # Calを初期状態に設定する
    # 測定時間を設定
    pag_tools.move_click('assets\\pd_cal_MeasTime.png')
    pag_tools.key_del(3)
    pag_tools.input_text(MEAS_TIME)
    pag_tools.key_enter(1)
    sleep(2)
    # 測定インターバルを設定
    pag_tools.move_click('assets\\pd_cal_IntervalTime.png')
    pag_tools.key_del(3)
    pag_tools.input_text(0)
    pag_tools.key_enter(1)
    sleep(2)


    for idx, MEAS in enumerate(MEAS_LIST):
        print(f"{idx}回目の測定    条件> {MEAS}")
        if idx != 0:
            # 2回目だけ（1回目はスキップ）
            # Calモードに入る
            # pag_tools.move_click('assets\\pd_cal_toggle.png')
            pag_tools.move_click_try2('assets\\pd_cal_toggle.png', 'assets\\pd_cal_toggle2.png')
            sleep(2)
            # 設定ウィンドウを表示させる
            pag_tools.move_click('assets\\pd_cal_f0.png')
            sleep(2)
            # BWが変わったときだけBWの設定を変更する操作を実行する
            if MEAS['BW_CHANGE']:
                pag_tools.move_click('assets\\pd_cal_bw.png')
                sleep(4)
                pag_tools.move_click_cal_textarea('assets\\pd_cal_bw_textarea.png', 80, -33)
            # f0の設定
            pag_tools.move_click_cal_textarea('assets\\pd_cal_f0_textarea.png', 110, -29)
            sleep(2)
            pag_tools.key_bkspace(9)
            pag_tools.input_text(MEAS['f0'])
            pag_tools.key_enter(1)
            pag.click()
            sleep(2)
            # Calモードを抜ける
            pag_tools.move_click_cal_textarea('assets\\pd_cal_toggle_off.png', 0, 300)
            sleep(2)

        # 測定開始
        pag_tools.move_click('assets\\pd_start.png')
        sleep(2)
        # シリアルNoに測定条件をメモする
        pag_tools.move_click_cal_textarea('assets\\pd_meas_serialno.png', 80, 30)
        pag_tools.key_bkspace(20)
        pag_tools.input_text(MEAS['SN'])
        sleep(2)
        pag_tools.move_click('assets\\pd_meas_enter.png')
        sleep(MEAS_TIME + 5)    # 測定が終わるまで待機

        # デバッグ用 ループを抜ける
        if idx >= 5:
            return

    pag.alert(text='自動操作が完了しました', title='自動処理完了メッセージ', button='OK')

    return


if __name__ == "__main__":
    main()
    print("自動処理が完了しました")
