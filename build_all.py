#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""data/ 폴더의 모든 *.json 파일을 output/ 폴더의 pptx로 변환합니다.
GitHub Actions에서 data/**.json 이 push 될 때마다 자동 실행됩니다.
"""
import glob
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))

def main():
    data_files = sorted(glob.glob(os.path.join(HERE, "data", "*.json")))
    if not data_files:
        print("data/ 폴더에 json 파일이 없습니다.")
        return
    os.makedirs(os.path.join(HERE, "output"), exist_ok=True)
    for data_path in data_files:
        stem = os.path.splitext(os.path.basename(data_path))[0]
        out_path = os.path.join(HERE, "output", f"{stem}.pptx")
        print(f"생성 중: {data_path} -> {out_path}")
        result = subprocess.run(
            [sys.executable, os.path.join(HERE, "generate_report.py"), data_path, out_path],
            cwd=HERE,
        )
        if result.returncode != 0:
            print(f"실패: {data_path}", file=sys.stderr)
            sys.exit(1)

if __name__ == "__main__":
    main()
