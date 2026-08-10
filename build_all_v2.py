#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""data/ 폴더뿐 아니라 저장소 루트에 실수로 올라온 *.json 파일까지 찾아서
output/ 폴더의 pptx로 변환합니다. 루트에 있던 파일은 처리 후 자동으로
data/ 폴더 안으로 옮겨서 정리합니다 (업로드 위치를 매번 신경 안 써도 되게).
"""
import glob
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))

# 이 스크립트 자체가 관여하지 않는 설정성 json 파일은 건드리지 않기 위한 제외 목록
IGNORE_NAMES = {"package.json", "package-lock.json"}


def main():
    data_dir = os.path.join(HERE, "data")
    os.makedirs(data_dir, exist_ok=True)
    os.makedirs(os.path.join(HERE, "output"), exist_ok=True)

    # 루트에 잘못 올라온 json을 먼저 data/ 로 이동
    root_files = sorted(glob.glob(os.path.join(HERE, "*.json")))
    for f in root_files:
        name = os.path.basename(f)
        if name in IGNORE_NAMES:
            continue
        dest = os.path.join(data_dir, name)
        if os.path.abspath(f) == os.path.abspath(dest):
            continue
        print(f"루트에서 발견: {name} -> data/ 로 이동")
        shutil.move(f, dest)

    data_files = sorted(glob.glob(os.path.join(data_dir, "*.json")))
    if not data_files:
        print("data/ 폴더에 json 파일이 없습니다.")
        return

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
