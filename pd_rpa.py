from time import sleep
from datetime import datetime
import subprocess
import sys
import os
import shutil
import pyautogui as pag
import src.pag_tools as pag_tools
import src.helpers as helpers


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def main():
    # 共通設定
    ROOT_DIR = 'C:/SOKEN/PD9/SampleSoftware260/Data/'
    GRAPH_DIR_NAME = '2D_Graph'
    TEMP_DIR_NAME = 'pd_temp'
    TEMP_DIR = f'{ROOT_DIR}{TEMP_DIR_NAME}'    # 結果を一時保存するディレクトリ
    # MEAS_TIME = 60    # 一時的に5に変更、後で60に戻す

    # デバッグ用
    # STOP_IDX = 0

    # 実行前の確認
    pag_confirm = pag.confirm(text='自動処理を開始しますか？\n⚠注意 自動処理中は絶対にマウスに触らないでください\nただし、処理を中断したいときは、マウスを画面の左上端に移動させてください', title='自動処理実行前の確認', buttons=['OK', 'Cancel'])
    if pag_confirm == "Cancel":
        print("Cancelボタンが押されました。処理を中止します。")
        return

    MODE = pag.confirm(text='実行する操作を選択してください', title='操作モードの選択', buttons=['測定', '解析', 'キャンセル'])
    # MODE = '解析'    # デバッグ用

    # 部分放電測定器用のRPA
    # クリック対象の画像をキャプチャして assets ディレクトリに入れておく
    if MODE == '測定':
        print(">>> 測定モード")

        MEAS_TIME = int(pag.prompt(text="測定時間(秒)を指定してください\n基本：60秒", title="測定条件を設定", default=60))
        # print(f"MEAS_TIME:{MEAS_TIME}, type:{type(MEAS_TIME)}")    # デバッグ用

        now = datetime.now()
        default_name = now.strftime('%Y%m%d%H%M%S')
        OUT_DIR_NAME = pag.prompt(text=f"出力フォルダ名を指定してください\n\n基準フォルダ:{ROOT_DIR}", title="結果出力先を指定", default=default_name)
        if not OUT_DIR_NAME:
            print("キャンセルされました")
            return
        OUT_DIR = ROOT_DIR + OUT_DIR_NAME

        # CSVから測定条件を読み込む
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        MEAS_LIST = helpers.getMeasList(f"{desktop_path}\\部分放電測定を行うフィルタの組み合わせ.csv")

        # アプリを起動する
        pag.FAILSAFE = True
        app = r"C:/SOKEN/PD9/SampleSoftware260/PD9.exe"
        subprocess.Popen(app)
        sleep(7)

        # Calを初期状態に設定する
        # 測定時間を設定
        pag_tools.move_click(resource_path('assets\\pd_cal_MeasTime.png'))
        pag_tools.key_del(3)
        pag_tools.input_text(MEAS_TIME)
        pag_tools.key_enter(1)
        sleep(2)
        # 測定インターバルを設定
        pag_tools.move_click(resource_path('assets\\pd_cal_IntervalTime.png'))
        pag_tools.key_del(3)
        pag_tools.input_text(0)
        pag_tools.key_enter(1)
        sleep(2)

        for idx, MEAS in enumerate(MEAS_LIST):
            print(f"{idx}回目の測定    条件> {MEAS}")
            # 2回目だけ（1回目はスキップ）←f0=200kHzならBWの変更が容易のため１回目から設定するように変更
            # Calモードに入る
            pag_tools.move_click_try2(resource_path('assets\\pd_cal_toggle.png'), resource_path('assets\\pd_cal_toggle2.png'))
            sleep(2)
            # 設定ウィンドウを表示させる
            pag_tools.move_click(resource_path('assets\\pd_cal_f0.png'))
            sleep(3)
            # BWが変わったときだけBWの設定を変更する操作を実行する
            if MEAS['BW_CHANGE']:
                pag_tools.move_click(resource_path('assets\\pd_cal_bw.png'))
                sleep(4)
                # BWをうまく変更できるように、一時的にf0を200kHzに変更する
                pag_tools.move_click_cal_textarea(resource_path('assets\\pd_cal_f0_textarea.png'), 70, 340)
                sleep(3)
                pag_tools.key_bkspace(9)
                pag_tools.input_text(200)    # BW変更のために200kHzに一時的に設定
                pag_tools.key_enter(1)
                pag.click()
                sleep(2)
                # BWを変更する
                pag_tools.move_click_cal_textarea(resource_path('assets\\pd_cal_bw_textarea.png'), 70, 340)
                sleep(3)
                if MEAS['BW_kHz'] == 300:
                    # 50 -> 300 に変更する
                    pag_tools.move_click(resource_path('assets\\pd_cal_bw_50to300.png'))
                elif MEAS['BW_kHz'] == 50:
                    # 300 -> 50 に変更する
                    pag_tools.move_click(resource_path('assets\\pd_cal_bw_300to50.png'))
            # f0の設定
            pag_tools.move_click_cal_textarea(resource_path('assets\\pd_cal_f0_textarea.png'), 70, 340)
            sleep(3)
            pag_tools.key_bkspace(9)
            pag_tools.input_text(MEAS['f0_kHz'])
            pag_tools.key_enter(1)
            pag.click()
            sleep(2)
            # Calモードを抜ける
            pag_tools.move_click_cal_textarea(resource_path('assets\\pd_cal_toggle_off.png'), 0, 300)
            sleep(2)

            # 測定開始
            pag_tools.move_click(resource_path('assets\\pd_start.png'))
            sleep(2)

            if idx == 0:
                # 初回のみ保存フォルダを設定する
                pag_tools.move_click(resource_path('assets\\pd_meas_dir_open.png'))
                sleep(2)
                # フォルダパスを入力する
                # pag_tools.move_click_cal_textarea(resource_path('assets\\pd_meas_dir_path_textarea.png'), 900, 40)
                pag_tools.move_click_try2_cal_textarea(resource_path('assets\\pd_meas_dir_path_textarea.png'), resource_path('assets\\pd_analyze_dir_path_textarea.png'), 900, 40, 900, 40)
                sleep(2)
                pag_tools.key_CtrlA2del()    # 既に入力されている内容を全て消去
                pag_tools.input_text(TEMP_DIR)
                pag_tools.key_enter(1)
                sleep(2)
                # フォルダを指定する
                pag_tools.move_click(resource_path('assets\\pd_meas_dir_path_set.png'))
                sleep(2)

            # シリアルNoに測定条件をメモする
            pag_tools.move_click_cal_textarea(resource_path('assets\\pd_meas_serialno.png'), 80, 30)
            pag_tools.key_bkspace(20)
            pag_tools.input_text(MEAS['Serial_No'])
            sleep(2)
            pag_tools.move_click(resource_path('assets\\pd_meas_enter.png'))
            sleep(MEAS_TIME + 5)    # 測定が終わるまで待機

            # デバッグ用 ループを抜ける
            # if idx >= STOP_IDX:
            #     break

        # 解析結果の後処理
        MOVE_DIR_CONFIRM = pag.confirm(text='一時フォルダから指定した保存フォルダにデータを移動しますか？', title='自動処理実行後の確認', buttons=['OK', 'Cancel'])
        if MOVE_DIR_CONFIRM == 'OK':
            # 一時ディレクトリ内のデータを移動する
            if not os.path.exists(OUT_DIR):
                os.makedirs(OUT_DIR)
            # 一時ディレクトリ内のすべてのファイルとディレクトリをOUT_DIRに移動する
            for item in os.listdir(TEMP_DIR):
                source = os.path.join(TEMP_DIR, item)
                destination = os.path.join(OUT_DIR, item)
                # ディレクトリまたはファイルを移動
                shutil.move(source, destination)
            pag.alert(text='自動操作が完了しました', title='自動処理完了メッセージ', button='OK')
            print(OUT_DIR)
            cmd = 'explorer {}'.format(OUT_DIR.replace('/', '\\'))
            subprocess.Popen(cmd)
        else:
            pag.alert(text='自動操作が完了しました', title='自動処理完了メッセージ', button='OK')
            cmd = 'explorer {}'.format(TEMP_DIR.replace('/', '\\'))
            subprocess.Popen(cmd)

    elif MODE == '解析':
        print(">>> 解析モード")
        CSV_DIR_NAME = pag.prompt(text=f"結果CSVファイルが保存されたパスを指定してください\n\n基本フォルダ:{TEMP_DIR_NAME}\n\n⚠注意 古い2D_Graphフォルダは削除されます", title="結果CSV保存フォルダを指定", default=TEMP_DIR_NAME)
        if not CSV_DIR_NAME:
            print("キャンセルされました")
            return
        CSV_DIR = ROOT_DIR + CSV_DIR_NAME
        print(CSV_DIR)
        # CSVファイルのリストを取得
        all_files = os.listdir(CSV_DIR)
        all_files.sort()
        csv_files = [f for f in all_files if f.endswith('.csv')]
        # 古い 2D_Graph ディレクトリを削除する
        if os.path.exists(CSV_DIR + '/' + GRAPH_DIR_NAME):
            print("既存の2D_Graphディレクトリを削除しました")
            shutil.rmtree(CSV_DIR + '/' + GRAPH_DIR_NAME)

        # アプリを起動する
        pag.FAILSAFE = True
        app = r"C:/SOKEN/PD9/SampleSoftware260/PD9.exe"
        subprocess.Popen(app)
        sleep(7)

        # アナライズモードに切り替え
        pag_tools.move_click(resource_path('assets\\pd_tab_analyze.png'))
        sleep(5)
        for idx, csv_fname in enumerate(csv_files):
            # CSVファイルを指定する
            pag_tools.move_click(resource_path('assets\\pd_analyze_dir_open.png'))
            sleep(3)
            if idx == 0:
                # 最初だけフォルダパスを入力する
                pag_tools.move_click_try2_cal_textarea(resource_path('assets\\pd_analyze_dir_path_textarea.png'), resource_path('assets\\pd_meas_dir_path_textarea.png'), 900, 40, 900, 40)
                sleep(2)
                pag_tools.key_CtrlA2del()    # 既に入力されている内容を全て消去
                pag_tools.input_text(CSV_DIR)
                pag_tools.key_enter(1)
                sleep(2)
            # ファイルを指定する
            pag_tools.move_click_cal_textarea(resource_path('assets\\pd_analyze_file_name_textarea.png'), 800, 0)
            pag_tools.key_CtrlA2del()    # 既に入力されている内容を全て消去
            pag_tools.input_text(csv_fname)
            pag_tools.move_click(resource_path('assets\\pd_analyze_dir_path_set.png'))
            sleep(8)
            # 2D Graphタブに切り替え
            pag_tools.move_click(resource_path('assets\\pd_analyze_2dgraph_tab.png'))
            sleep(3)
            # グラフ表示の設定
            pag_tools.move_click_cal_textarea(resource_path('assets\\pd_analyze_2dgraph_setY.png'), 50, 0)
            sleep(3)
            pag_tools.move_click_try2(resource_path('assets\\pd_analyze_2dgraph_setQ.png'), resource_path('assets\\pd_analyze_2dgraph_setQ_inv.png'))
            sleep(4)
            pag_tools.move_click_cal_textarea(resource_path('assets\\pd_analyze_2dgraph_setX.png'), 50, 0)
            sleep(3)
            pag_tools.move_click_try2(resource_path('assets\\pd_analyze_2dgraph_setPhi.png'), resource_path('assets\\pd_analyze_2dgraph_setPhi_inv.png'))
            sleep(4)
            pag_tools.move_click_cal_textarea(resource_path('assets\\pd_analyze_2dgraph_SineWave.png'), 90, 0)
            sleep(9)
            pag_tools.move_click_cal_textarea(resource_path('assets\\pd_analyze_2dgraph_DistributionMap.png'), 120, 0)
            sleep(9)
            pag_tools.move_click(resource_path('assets\\pd_analyze_2dgraph_SaveImage.png'))
            sleep(5)
            # 次のCSVファイルのためにタブを切り替える
            pag_tools.move_click(resource_path('assets\\pd_analyze_ChooseDataArea_tab.png'))
            sleep(5)

        # 解析結果のメッセージを表示する
        pag.alert(text='自動操作が完了しました\n2D_Graphフォルダに画像が保存されているのでチェックしてください', title='自動処理完了メッセージ', button='OK')
        cmd = 'explorer {}'.format(CSV_DIR.replace('/', '\\'))
        subprocess.Popen(cmd)

    else:
        print("Cancelボタンが押されました。処理を中止します。")
        return

    return



if __name__ == "__main__":
    main()
    print("自動処理が完了しました")
