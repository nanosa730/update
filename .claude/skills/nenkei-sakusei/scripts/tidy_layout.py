#!/usr/bin/env python3
"""年計シートの罫線・配置を整え、会議記号の列にある長い文字を行事の欄へ移す。
使い方: python tidy_layout.py 年計.xlsx
組み替え・手直しのあと、PDFに書き出す前に実行する。"""
import subprocess, sys, os
here = os.path.dirname(os.path.abspath(__file__))
p = sys.argv[1]
subprocess.run([sys.executable, os.path.join(here, '_codefix.py'), p], check=True)
subprocess.run([sys.executable, os.path.join(here, '_border.py'), p], check=True)
