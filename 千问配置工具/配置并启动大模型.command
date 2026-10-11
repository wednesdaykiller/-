#!/bin/bash
set -euo pipefail
APP_DIR='/Users/luzihan/Documents/金融研报智能纠错系统_v1.0/两份PDF研报核查助手'
if [[ ! -f "$APP_DIR/app.py" ]]; then echo '未找到应用目录，请保留应用原位置。';read -r _done;exit 1;fi
if [[ -x "$APP_DIR/.venv/bin/python" ]];then APP_PYTHON="$APP_DIR/.venv/bin/python";else APP_PYTHON='/Library/Frameworks/Python.framework/Versions/3.12/bin/python3';fi
"$APP_PYTHON" "$(dirname "$0")/configure_and_start.py" "$APP_DIR" || { echo '配置未完成；请根据上面的提示处理。';read -r _done; }
