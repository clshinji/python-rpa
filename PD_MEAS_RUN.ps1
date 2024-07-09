# PowerShell スクリプト: run_python_script.ps1

# 仮想環境のパスを設定
$venvPath = "C:\Users\TC\python\pd-analysis"

# 仮想環境を有効化
& "$venvPath\Scripts\Activate.ps1"

# Python スクリプトのパスを設定
$pythonScriptPath = "$venvPath\python-rpa\pd_rpa.py"

# Python スクリプトを実行
python $pythonScriptPath

# 必要に応じて仮想環境を無効化
deactivate
