# Python で RPA してみる！

> [!WARNING]
> このブランチは SOKEN 部分放電測定システム 向けの RPA コードを追加しています。<br />main ブランチとの整合は確保していませんので、ご注意ください。

仕事の都合で手作業を自動化するために RPA を作る必要があったのでやってみました。<br />手っ取り早いのは `Power Automate Desktop` 様ですが、対象の環境は Microsoft365 のライセンスが導入されておらず、勝手にインストールするのも気が引けたので、前から気になっていた `Pythonで自動化` を試してみることに。<br />ウィジェットの要素を引っ張ってくる方法もありますが、汎用性を求めずに、画面上のアイコンのキャプチャ画像をマッチングしてマウスを移動してポチポチする RPA として構築しました。<br />とりあえず動くのでとりあえずは OK、ということで…<br />参考にさせていただいた記事（感謝！）：[「画像認識する RPA」を作って、「毎月 1,000 件の手入力」を自動化された事例](https://forum.pc5bai.com/article/rpa-by-python/)

# 環境構築～動作テスト

## 1. 環境構築＆ローカル環境へのクローン

最初に Python の仮想環境を構築しておきます。

```bash
python -m venv env
```

続いて、仮想環境ディレクトリ内に、このリポジトリをクローンします。<br />ターミナルを起動して、任意のディレクトリ（ドキュメント等）で次のコマンドを実行します。

```bash
cd env
git clone https://github.com/clshinji/python-rpa.git
```

部分放電測定用のブランチに切り替えます。

```bash
cd python-rpa
git checkout pd-meas
```

## 2. 必要なライブラリをインストールする

続いて、必要なライブラリをインストールします。<br />これ以降の操作は仮想環境化で実行していきます。

まずは仮想環境を有効化します。

```bash
script\Activate.ps1
```

次に必要なライブラリをインストール。

```bash
pip install -r requrements.txt
```

## 3. Python ファイルを実行する場合

測定条件を`部分放電測定を行うフィルタの組み合わせ.csv`に記載します。<br />`中心周波数f0(kHz)`と`バンド幅BW(kHz)`を 1 行ずつ記載します。<br />うまく自動処理させるコツは、`BW`をなるべく変えずに、`f0`を低い周波数から順番に上げていくことです。<br />うまく動作させられないときは、作成者に問い合わせてください。

それでは、さっそく Python ファイルを実行していきます。<br />自動処理中はマウスから手を離して、コーヒーを飲みながら終わるのを待ちましょう。<br />間違ってもマウスを触ってしまうと、うまく操作できなくなってしまいます。

```bash
python pd_rpa.py
```

操作が終わったら仮想環境を無効化します。

```bash
deactivate
```

## 4. 実行用ファイルで 1 発スタート

`PD_MEAS_RUN.ps1`をダブルクリックして実行します。<br />以上！！<br />(上記の操作が自動で実行されるようになっています…)
