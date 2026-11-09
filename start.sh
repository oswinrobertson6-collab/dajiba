#!/usr/bin/env bash
cd "$(dirname "$0")"
source .venv/bin/activate
echo "Starting 无花果项目供应商智能录入系统 on http://127.0.0.1:8000 ..."
python -m uvicorn app:app --host 0.0.0.0 --port 8000
