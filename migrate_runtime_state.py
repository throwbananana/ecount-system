# -*- coding: utf-8 -*-
import shutil
from pathlib import Path

from runtime_paths import APP_HOME, ensure_app_home


FILES_TO_MOVE = [
    "config.json",
    "base_data.db",
    "reconciliation.db",
    "shipping.bd",
    "reconciliation_header_mapping.json",
]


def main():
    repo_dir = Path(__file__).resolve().parent
    ensure_app_home()
    print(f"运行时目录: {APP_HOME}")

    for name in FILES_TO_MOVE:
        src = repo_dir / name
        dst = APP_HOME / name

        if not src.exists():
            continue
        if dst.exists():
            print(f"[跳过] 目标已存在: {dst}")
            continue
        shutil.move(str(src), str(dst))
        print(f"[迁移] {src} -> {dst}")


if __name__ == "__main__":
    main()
