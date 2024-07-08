# PowerShell スクリプト: run_python_script.ps1

# 仮想環境のパスを設定
$venvPath = "C:\path\to\your\venv\Scripts\"

# 仮想環境を有効化
& "$venvPath\Activate.ps1"

# Python スクリプトのパスを設定
$pythonScriptPath = "C:\path\to\your\script.py"

# Python スクリプトを実行
python $pythonScriptPath

# 必要に応じて仮想環境を無効化
& "$venvPath\deactivate.ps1"