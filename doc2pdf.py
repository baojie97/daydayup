#!/usr/bin/env python
# -*- coding: utf-8 -*-
# author : BaoJIE

from comtypes.client import CreateObject
import os
import logging
from pathlib import Path


logging.basicConfig(level=logging.INFO)
# 防止文件重复
def make_unique_filename(target_dir, filename):
    """生成唯一的文件名"""
    base = Path(filename).stem
    ext = Path(filename).suffix
    counter = 0
    while True:
        candidate = f"{base}{f'_{counter}' if counter else ''}{ext}"
        if not (target_dir / candidate).exists():
            return candidate
        counter += 1
class PDFConverter:
    def __init__(self):
        self.wdFormatPDF = 17
        try:
            self.word_app = CreateObject("Word.Application")
            self.word_app.Visible = False
            logging.info("Word COM对象创建成功")
        except Exception as e:
            logging.error(f"Word初始化失败: {str(e)}")
            raise

    def convert_folder(self, folder):
        if not os.path.isdir(folder):
            raise ValueError("无效的目录路径")

        try:
            for root, _, files in os.walk(folder):
                for file in files:
                    if file.startswith('~') or not file.lower().endswith(('.doc', '.docx')):
                        continue

                    src_path = os.path.abspath(os.path.join(root, file))
                    pdf_path = os.path.splitext(src_path)[0] + '.pdf'

                    if os.path.exists(pdf_path):
                        logging.warning(f"跳过已存在的文件: {pdf_path}")
                        continue

                    try:
                        doc = self.word_app.Documents.Open(src_path)
                        doc.SaveAs(pdf_path, FileFormat=self.wdFormatPDF)
                        doc.Close()
                        logging.info(f"转换成功: {file} -> {os.path.basename(pdf_path)}")
                    except Exception as e:
                        logging.error(f"文件转换失败[{file}]: {str(e)}")

        finally:
            self.word_app.Quit()


if __name__ == "__main__":
    try:
        source = input("请输入文档目录路径：").strip('"')
        # 开始进行 文本转pdf
        converter = PDFConverter()
        converter.convert_folder(source)
    except Exception as e:
        logging.error(f"运行时错误: {str(e)}")
    finally:
        input("按回车键退出...")
