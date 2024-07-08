import pandas as pd


def getMeasList(fpath):
    """測定条件を記録したCSVから実行用のリストを作成する

    Args:
        fpath (str): csvファイルのパス

    Returns:
        result_list (list): 解析条件を記録したリスト
    """
    df = pd.read_csv(fpath, encoding="cp932")
    # メモ用の Serial_No 列を作る
    df['Serial_No'] = df.apply(lambda row: f"{row['f0_kHz']}BW{row['BW_kHz']}", axis=1)
    # BW_kHzが変化している行はTrue
    df['BW_CHANGE'] = df['BW_kHz'].diff().ne(0)

    # 1行目は False に設定
    df.loc[0, 'BW_CHANGE'] = False

    # データフレームをリストに変換する
    result_list = df.to_dict(orient='records')

    return result_list
