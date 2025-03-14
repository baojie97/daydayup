#!/usr/bin/env python
# -*- coding: utf-8 -*-
# author : BaoJIE

from comtypes.client import CreateObject
import os
import logging
import shutil
from pathlib import Path


logging.basicConfig(level=logging.INFO)

# 集中文件
def gather_files(source_dir, target_dir="gathered_files"):
    """
    集中文件到目标文件夹
    参数:
    source_dir - 源目录路径（必须）
    target_dir - 目标目录路径（可选，默认为项目当前目录下的gathered_files）
    """
    # 确保目标路径有效
    target_path = Path(target_dir).resolve()
    target_path.mkdir(parents=True, exist_ok=True)
    counter = {'success': 0, 'duplicate': 0, 'error': 0}
    for root, _, files in os.walk(source_dir):
        for filename in files:
            source_file = Path(root) / filename
            try:
                # 生成唯一的目标文件名
                dest_name = make_unique_filename(target_path, filename)
                # 复制文件并保留元数据
                shutil.copy2(source_file, target_path / dest_name)
                # 更新计数器
                key = 'success' if filename == dest_name else 'duplicate'
                counter[key] += 1
                # 实时打印进度（单行更新）
                print(
                    f"\r处理中：成功 {counter['success']} | 重复 {counter['duplicate']} | 错误 {counter['error']} | 当前文件: {filename[:30]:<30}",
                    end="")
            except Exception as e:
                print(f"\n错误处理文件 [{filename}]: {str(e)}")
                counter['error'] += 1
    return counter
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
        # 创建一个新的文件夹：gathered_files，复制路径作为目标路径
        target = input("请输入目标目录路径（留空使用默认路径）：").strip()
        # 验证源路径
        if not Path(source).exists():
            print("错误：源目录不存在")
            exit(1)
        # 执行复制操作（修复了None传递问题）
        if target:
            result = gather_files(source, target)
        else:
            result = gather_files(source)  # 无目标参数，使用默认值
        # 输出最终集合结果
        print("\n\n=== 处理结果 ===")
        print(f"目标目录位置: {Path(target or 'gathered_files').resolve()}")
        print(f"成功复制文件数: {result['success']}")
        print(f"重命名重复文件数: {result['duplicate']}")
        print(f"处理失败文件数: {result['error']}")
        print(f"开始进行文本转pdf,请等候: ")
        # 开始进行 文本转pdf
        converter = PDFConverter()
        converter.convert_folder(target)
    except Exception as e:
        logging.error(f"运行时错误: {str(e)}")
    finally:
        input("按回车键退出...")
