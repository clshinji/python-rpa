# PythonでRPAしてみる！
仕事の都合で手作業を自動化するためにRPAを作る必要があったのでやってみました。<br />手っ取り早いのは `Power Automate Desktop` 様ですが、対象の環境はMicrosoft365のライセンスが導入されておらず、勝手にインストールするのも気が引けたので、前から気になっていた `Pythonで自動化` を試してみることに。<br />ウィジェットの要素を引っ張ってくる方法もありますが、汎用性を求めずに、画面上のアイコンのキャプチャ画像をマッチングしてマウスを移動してポチポチするRPAとして構築しました。<br />とりあえず動くのでとりあえずはOK、ということで…<br />参考にさせていただいた記事（感謝！）：[「画像認識するRPA」を作って、「毎月1,000件の手入力」を自動化された事例](https://forum.pc5bai.com/article/rpa-by-python/)

# 環境構築～動作テスト
## 1. ローカル環境にクローンする
最初にこのリポジトリをクローンします。<br />ターミナルを起動して、任意のディレクトリ（ドキュメント等）で次のコマンドを実行します。
``` bash
git clone https://github.com/clshinji/python-rpa.git
```

## 2. 必要なライブラリをインストールする
まずは必要なライブラリをインストールします。<br />（Pythonの実行環境が無い場合は公式サイトからダウンロードをお願いします…）

可能であれば、`pyenv`等での仮想環境を利用することをおススメします。

念のため、インストールされているPythonのバージョンをチェックします。<br />(後ろについている `-V`は大文字のブイです)
``` bash
python -V
```

もしこれでエラーが出る場合は、パスが通っていない可能性が高いので、次のコマンドも試してみてください。<br />これでうまくいく場合は、以降のコマンドで`python`という箇所は、全て`python3`に置き換えればOKです。
``` bash
python3 -V
```
出力結果の例（インストールされているPythonのバージョンが表示されます）
``` bash
Python 3.8.7
```

``` bash
pip install -r requrements.txt
```

## 3. サンプルコードを実行してみる
うまく環境が構築できているか試すために、サンプルコードを実行します。<br />[参考にさせていただいたサイト](https://forum.pc5bai.com/article/rpa-by-python/)のコードをベースに作成した電卓を自動でポチポチするコードを動かしてみます。<br />※想定している動作環境：Windows11 ダークモード（個人用設定のアクセントカラー：シーフォーム `#00B7C3` ）<br />　電卓のイコールがうまく認識されない場合は、 `assets\calc_eq_dark.png` をご自身の端末でキャプチャした画像に変更してください。また、マウスオーバーされた場合も想定して `assets\calc_eq_dark_mo.png` に ＝ボタン をマウスオーバーした状態のキャプチャも保存しておくと冗長性が少し高くなります。

``` bash
python sample_calc.py
```

ちゃんと動作すると、電卓が起動して足し算をしてくれます。<br />もしうまく動かない場合は、pythonの実行環境がちゃんとインストールできているか、必要なライブラリがインストールできているか、`assets`ディレクトリ内の画像とご自身の環境のアプリの画面が合うか、等を確かめてください。<br />[電卓RPAデモ動画](https://github.com/clshinji/python-rpa/blob/d9a89800858db2495be0b07835715ff7d3af0442/demo/demo_calc.mp4)<br />[動画RPAデモ動画(YouTube)](https://youtu.be/VNADz47d1bg)

続いて、エクセルを起動するだけのサンプルコードを試してみます。<br />これも動けばバッチリです！

``` bash
python sample_excel.py
```

# 4. 自分で自由にRPAしてみる
`画像を認識 -> クリック` する機能は `src\pag_tools.py` にまとめています。<br />ご自身のアプリ用のRPAを作る場合は、サンプルコードを参考に作成してみてください。

