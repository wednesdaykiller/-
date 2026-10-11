#!/bin/bash
set -euo pipefail
cd "$(dirname "$0")"
if [[ -x .venv/bin/python ]]; then
  APP_PYTHON="$PWD/.venv/bin/python"
elif [[ -x /Library/Frameworks/Python.framework/Versions/3.12/bin/python3 ]]; then
  APP_PYTHON='/Library/Frameworks/Python.framework/Versions/3.12/bin/python3'
else
  echo '请安装 Python 3.12，并先运行“安装运行环境.command”。'; read -r _stop; exit 1
fi
if [[ -z "${OCR_DEPS_PATH:-}" && -d '/Users/luzihan/Documents/金融研报智能纠错系统_v1.0/.ocr_deps' ]]; then
 export OCR_DEPS_PATH='/Users/luzihan/Documents/金融研报智能纠错系统_v1.0/.ocr_deps'
fi
echo '启动后浏览器会打开。请上传财报PDF和研报PDF，然后点击“开始核查”。'
echo '关闭此终端或按Ctrl+C即可停止应用。'
"$APP_PYTHON" app.py || { echo '启动未完成，请检查依赖或端口；先运行安装脚本可建立独立环境。'; read -r _stop; }
