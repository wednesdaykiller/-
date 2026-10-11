#!/bin/bash
set -euo pipefail
cd "$(dirname "$0")"
APP_PYTHON='/Library/Frameworks/Python.framework/Versions/3.12/bin/python3'
if [[ ! -x "$APP_PYTHON" ]]; then echo '请先安装Python 3.12。';read -r _stop;exit 1;fi
"$APP_PYTHON" -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
printf '\n安装完成，现在可以双击“启动核查助手.command”。\n'
read -r _done
