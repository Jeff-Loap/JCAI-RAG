# -*- coding: utf-8 -*-
import os
import subprocess
import sys

PYTHON_EXE = r"D:\vng\python.exe"
WORKDIR = r"D:\PythonFile\JCAI\RAG"

# 按你的文件名改这里（脚本不存在就会报错）
STEPS = [
    ("OCR", "ocr.py"),
    ("CHUNK", "chunk.py"),
    ("INDEX", "index.py"),      # 如果你用的是 Qdrant，就改成 index_agent_guide.py
    ("QA", "rag_qa.py"),        # 如果你想跑你原来的 rag_qa.py，就改成 rag_qa.py
]

def run_step(name: str, script: str):
    script_path = os.path.join(WORKDIR, script)
    if not os.path.exists(script_path):
        raise FileNotFoundError(f"[{name}] 找不到脚本：{script_path}")

    print(f"\n==================== {name} ====================")
    print(f"运行：{PYTHON_EXE} {script_path}\n")

    # 用同一个解释器跑，且在 WORKDIR 下运行
    p = subprocess.run(
        [PYTHON_EXE, script_path],
        cwd=WORKDIR,
        text=True
    )

    if p.returncode != 0:
        raise RuntimeError(f"[{name}] 执行失败，退出码：{p.returncode}")

def main():
    print("开始依次执行任务：")
    print("WORKDIR:", WORKDIR)
    print("PYTHON :", PYTHON_EXE)

    for name, script in STEPS:
        run_step(name, script)

    print("\n全部执行 完成。")

if __name__ == "__main__":
    main()