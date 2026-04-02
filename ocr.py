# -*- coding: utf-8 -*-
import json
import os
import re
from pathlib import Path

import pytesseract
from pdf2image import convert_from_path
from PIL import Image
import numpy as np
import cv2

PDF_PATH = r"D:\PythonFile\JCAI\RAG\agent_guide.pdf"
OUT_TXT = r"D:\PythonFile\JCAI\RAG\agent_guide_ocr.txt"
OUT_JSONL = r"D:\PythonFile\JCAI\RAG\agent_guide_ocr_pages.jsonl"

# 如果你没把 tesseract 加到环境变量，这里手动指定：
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# 如果 pdf2image 找不到 poppler，就指定 poppler_path：
# POPPLER_PATH = r"C:\poppler\Library\bin"
POPPLER_PATH = None

def preprocess_for_ocr(pil_img: Image.Image) -> Image.Image:
    """扫描件增强：灰度 -> 去噪 -> 自适应阈值 -> 形态学清理"""
    img = np.array(pil_img)
    if img.ndim == 3:
        img = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)

    # 去噪（比 median 更稳一点）
    img = cv2.fastNlMeansDenoising(img, None, h=18, templateWindowSize=7, searchWindowSize=21)

    # 自适应阈值（对光照不均更友好）
    img = cv2.adaptiveThreshold(
        img, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        35, 11
    )

    # 形态学：去小噪点
    kernel = np.ones((2, 2), np.uint8)
    img = cv2.morphologyEx(img, cv2.MORPH_OPEN, kernel, iterations=1)

    return Image.fromarray(img)

def clean_text(text: str) -> str:
    # 去多余空格/重复空行
    text = text.replace("\x0c", "")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()

def main():
    pdf_path = Path(PDF_PATH)
    assert pdf_path.exists(), f"PDF 不存在：{pdf_path}"

    pages = convert_from_path(
        PDF_PATH,
        dpi=300,
        poppler_path=POPPLER_PATH
    )

    os.makedirs(Path(OUT_TXT).parent, exist_ok=True)

    with open(OUT_TXT, "w", encoding="utf-8") as f_txt, open(OUT_JSONL, "w", encoding="utf-8") as f_jsonl:
        for i, pil_page in enumerate(pages, start=1):
            img = preprocess_for_ocr(pil_page)

            # 英文扫描：lang="eng"
            # 中文扫描：lang="chi_sim"
            text = pytesseract.image_to_string(
                img,
                lang="eng",
                config="--oem 3 --psm 6"
            )
            text = clean_text(text)

            f_txt.write(f"\n\n===== PAGE {i} =====\n\n")
            f_txt.write(text + "\n")

            # 每页一条 jsonl，后续切分/入库更方便（用 json.dumps 自动转义，避免控制字符导致 JSON 无效）
            f_jsonl.write(json.dumps({"page": i, "text": text}, ensure_ascii=False) + "\n")

            print(f"OCR 完成：第 {i}/{len(pages)} 页，字符数={len(text)}")

    print("输出完成：")
    print("TXT:", OUT_TXT)
    print("JSONL:", OUT_JSONL)

if __name__ == "__main__":
    main()