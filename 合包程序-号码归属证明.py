# @OS:Windows 11
# @Python:3.12.5
# @Coding: UTF-8
# @功能：合并四个脚本的功能：
# 1. 批量生成号码归属证明
# 2. 批量转换Word文档为PDF
# 3. 批量转换PDF为JPG
# 4. 批量添加印章到JPG
# @时间：2025/5/27

import pandas as pd
from docx import Document
from datetime import datetime, timedelta
import random
import os
from win32com import client
import time
import sys
from pdf2image import convert_from_path
from PIL import Image
import logging
from logging.handlers import RotatingFileHandler


def setup_logger():
    """设置日志配置"""
    # 创建logs目录（如果不存在）
    if not os.path.exists('logs'):
        os.makedirs('logs')

    # 获取当前时间作为日志文件名的一部分
    current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_filename = f'logs/process_{current_time}.log'

    # 创建logger对象
    logger = logging.getLogger('DocumentProcessor')
    logger.setLevel(logging.DEBUG)

    # 创建文件处理器（rotating file handler，限制文件大小为10MB，保留5个备份）
    file_handler = RotatingFileHandler(
        log_filename,
        maxBytes=10 * 1024 * 1024,  # 10MB
        backupCount=5,
        encoding='utf-8'
    )
    file_handler.setLevel(logging.DEBUG)

    # 创建控制台处理器
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)

    # 创建格式化器
    file_formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s'
    )
    console_formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s'
    )

    # 将格式化器添加到处理器
    file_handler.setFormatter(file_formatter)
    console_handler.setFormatter(console_formatter)

    # 将处理器添加到logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger


# 创建全局logger对象
logger = setup_logger()


class PhoneCertificateGenerator:
    def __init__(self):
        self.start = datetime(2021, 1, 1)
        self.end = datetime(2025, 3, 31)
        self.logger = logging.getLogger('DocumentProcessor.PhoneCertificateGenerator')

    def random_date_between(self, start_date, end_date):
        """生成两个日期之间的随机日期"""
        time_between = end_date - start_date
        days_between = time_between.days
        random_days = random.randrange(days_between)
        return start_date + timedelta(days=random_days)

    def clean_filename(self, filename):
        """清理文件名，移除不允许的字符"""
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            filename = filename.replace(char, '')
        return filename

    def get_unique_filename(self, company_value, current_date, repeat_count=None):
        """生成唯一的文件名"""
        company_name = self.clean_filename(company_value)
        if repeat_count is not None and repeat_count > 0:
            filename = f'{company_name}_号码归属证明_{current_date}_{repeat_count}.docx'
        else:
            filename = f'{company_name}_号码归属证明_{current_date}.docx'
        return filename

    def replace_text_in_paragraph(self, paragraph, replacements):
        """在段落中替换文本，保持格式"""
        paragraph_text = paragraph.text
        runs = paragraph.runs

        has_replacements = any(key in paragraph_text for key in replacements.keys())
        if not has_replacements:
            return

        full_text = ''
        for run in runs:
            full_text += run.text

        new_text = full_text
        for key, value in replacements.items():
            if key in new_text:
                new_text = new_text.replace(key, str(value))

        if runs:
            runs[0].text = new_text
            for run in runs[1:]:
                run.text = ''

    def generate_certificates(self):
        """生成号码归属证明"""
        self.logger.info("开始生成号码归属证明...")
        try:
            excel_file = 'data.xlsx'
            self.logger.debug(f"正在读取Excel文件: {excel_file}")
            df = pd.read_excel(excel_file, header=0)

            numbers = []
            for num in df['NUMBERS'].tolist()[1:]:
                if pd.isna(num):
                    numbers.append('')
                else:
                    numbers.append(str(num))

            companies = []
            for comp in df['COMPANY'].tolist():
                if pd.isna(comp):
                    continue
                else:
                    companies.append(str(comp))

            if not companies:
                self.logger.error("没有找到有效的公司名称，请检查Excel文件中的COMPANY列数据")
                raise ValueError("没有找到有效的公司名称")

            batch_size = 10
            total_batches = (len(numbers) + batch_size - 1) // batch_size
            self.logger.info(f"共需处理 {total_batches} 批次数据")

            now = datetime.now()
            year, month, day = now.year, now.month, now.day

            available_companies = companies.copy()
            company_usage = {company: 0 for company in companies}

            for batch_idx in range(total_batches):
                self.logger.debug(f"正在处理第 {batch_idx + 1}/{total_batches} 批次")
                start_idx = batch_idx * batch_size
                end_idx = min(start_idx + batch_size, len(numbers))
                current_numbers = numbers[start_idx:end_idx]
                merged_numbers = '、'.join(current_numbers)

                try:
                    doc = Document('numbers.docx')
                except Exception as e:
                    self.logger.error(f"无法打开模板文件 numbers.docx: {str(e)}")
                    raise

                random_date = self.random_date_between(self.start, self.end)
                r_year, r_month, r_day = random_date.year, random_date.month, random_date.day

                if not available_companies:
                    company_value = random.choice(companies)
                else:
                    company_value = random.choice(available_companies)
                    if company_usage[company_value] == 0:
                        available_companies.remove(company_value)

                company_usage[company_value] += 1

                replacements = {
                    '{NUMBERS}': merged_numbers,
                    '{COMPANY}': company_value,
                    '{YEAR}': str(year),
                    '{MONTH}': str(month),
                    '{DAY}': str(day),
                    '{rYEAR}': str(r_year),
                    '{rMONTH}': str(r_month),
                    '{rDAY}': str(r_day)
                }

                for paragraph in doc.paragraphs:
                    self.replace_text_in_paragraph(paragraph, replacements)

                current_date = datetime.now().strftime('%Y%m%d')

                if company_usage[company_value] > 1:
                    output_file = self.get_unique_filename(company_value, current_date,
                                                           company_usage[company_value] - 1)
                else:
                    output_file = self.get_unique_filename(company_value, current_date)

                try:
                    doc.save(output_file)
                    self.logger.info(f'已生成文件：{output_file}')
                except Exception as e:
                    self.logger.error(f"保存文件 {output_file} 时出错: {str(e)}")
                    raise

            self.logger.info("号码归属证明生成完成")

        except Exception as e:
            self.logger.error(f"生成号码归属证明时发生错误: {str(e)}")
            raise


class DocConverter:
    def __init__(self):
        self.current_dir = os.path.dirname(os.path.abspath(__file__))
        self.logger = logging.getLogger('DocumentProcessor.DocConverter')

    def word_to_pdf(self, word_app, word_path, pdf_path, max_retries=3):
        """将Word文档转换为PDF"""
        doc = None
        for attempt in range(max_retries):
            try:
                self.logger.debug(f"尝试转换文件 {word_path} (第 {attempt + 1} 次)")
                doc = word_app.Documents.Open(word_path)
                doc.SaveAs(pdf_path, FileFormat=17)
                doc.Close()
                return True
            except Exception as e:
                if doc:
                    try:
                        doc.Close()
                    except:
                        pass
                self.logger.warning(f"第 {attempt + 1} 次尝试转换失败：{str(e)}")
                if attempt < max_retries - 1:
                    self.logger.info("等待后重试...")
                    time.sleep(2)
                else:
                    self.logger.error(f"转换失败: {os.path.basename(word_path)}")
                    return False
        return False

    def convert_to_pdf(self, template_name="numbers.docx"):
        """批量转换Word文档为PDF"""
        self.logger.info("开始转换Word文档为PDF...")
        try:
            today = datetime.now().strftime("%Y%m%d")
            pdf_folder = os.path.join(self.current_dir, f"{today}PDF")

            if not os.path.exists(pdf_folder):
                os.makedirs(pdf_folder)
                self.logger.info(f"创建文件夹: {pdf_folder}")

            files_to_convert = [
                f for f in os.listdir(self.current_dir)
                if (f.endswith(".doc") or f.endswith(".docx"))
                   and f != template_name
                   and not f.startswith("~$")
            ]

            if not files_to_convert:
                self.logger.warning("没有找到需要转换的Word文档")
                return

            word_app = None
            try:
                word_app = client.Dispatch("Word.Application")
                word_app.Visible = False
                word_app.DisplayAlerts = False

                total_files = len(files_to_convert)
                converted_count = 0

                for idx, file in enumerate(files_to_convert, 1):
                    word_path = os.path.join(self.current_dir, file)
                    pdf_name = os.path.splitext(file)[0] + ".pdf"
                    pdf_path = os.path.join(pdf_folder, pdf_name)

                    self.logger.info(f"[{idx}/{total_files}] 正在转换: {file}")
                    if self.word_to_pdf(word_app, word_path, pdf_path):
                        converted_count += 1
                        self.logger.info(f"成功转换: {file} -> {pdf_name}")

                self.logger.info(f"PDF转换完成！共成功转换 {converted_count}/{total_files} 个文件")
                return pdf_folder

            except Exception as e:
                self.logger.error(f"Word转PDF过程中发生错误：{str(e)}")
                raise

            finally:
                if word_app:
                    try:
                        word_app.Quit()
                    except:
                        pass
                    os.system("taskkill /F /IM WINWORD.EXE /T >nul 2>&1")
                    self.logger.debug("已关闭Word应用程序")

        except Exception as e:
            self.logger.error(f"PDF转换过程中发生错误：{str(e)}")
            raise


class ImageProcessor:
    def __init__(self):
        self.current_dir = os.path.dirname(os.path.abspath(__file__))
        self.logger = logging.getLogger('DocumentProcessor.ImageProcessor')

    def pdf_to_jpg(self, pdf_path, output_folder, dpi=300):
        """将PDF转换为JPG图片"""
        try:
            self.logger.debug(f"开始转换PDF文件: {pdf_path}")
            images = convert_from_path(pdf_path, dpi=dpi)
            pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]

            for i, image in enumerate(images):
                if len(images) == 1:
                    image_path = os.path.join(output_folder, f"{pdf_name}.jpg")
                else:
                    image_path = os.path.join(output_folder, f"{pdf_name}_第{i + 1}页.jpg")
                image.save(image_path, "JPEG")
                self.logger.debug(f"已保存图片: {image_path}")
            return True
        except Exception as e:
            self.logger.error(f"转换PDF文件 {pdf_path} 时失败：{str(e)}")
            return False

    def convert_pdfs_to_jpg(self, pdf_folder):
        """批量转换PDF文件为JPG图片"""
        self.logger.info("开始转换PDF为JPG...")
        try:
            today = datetime.now().strftime("%Y%m%d")
            jpg_folder = os.path.join(self.current_dir, f"{today}JPG")

            if not os.path.exists(jpg_folder):
                os.makedirs(jpg_folder)
                self.logger.info(f"创建文件夹: {jpg_folder}")

            converted_count = 0
            pdf_files = [f for f in os.listdir(pdf_folder) if f.endswith(".pdf")]
            total_files = len(pdf_files)

            for idx, file in enumerate(pdf_files, 1):
                pdf_path = os.path.join(pdf_folder, file)

                self.logger.info(f"[{idx}/{total_files}] 正在转换: {file}")
                if self.pdf_to_jpg(pdf_path, jpg_folder):
                    converted_count += 1
                    self.logger.info(f"成功转换: {file}")
                else:
                    self.logger.error(f"转换失败: {file}")

            self.logger.info(f"JPG转换完成！共成功转换 {converted_count}/{total_files} 个PDF文件")
            return jpg_folder

        except Exception as e:
            self.logger.error(f"PDF到JPG转换过程中发生错误：{str(e)}")
            raise

    def add_stamp_to_image(self, original_path, stamp_path, position, output_dir, opacity=0.5):
        """添加印章到图片"""
        try:
            self.logger.debug(f"开始处理图片: {original_path}")
            base_image = Image.open(original_path)
            stamp = Image.open(stamp_path)

            if stamp.mode != 'RGBA':
                stamp = stamp.convert('RGBA')

            alpha = stamp.getchannel('A')
            alpha = alpha.point(lambda x: int(x * opacity))
            stamp.putalpha(alpha)

            transparent = Image.new('RGBA', base_image.size, (0, 0, 0, 0))
            transparent.paste(stamp, position, stamp)

            output_image = Image.alpha_composite(base_image.convert('RGBA'), transparent)
            output_image = output_image.convert('RGB')

            name_without_ext = os.path.splitext(os.path.basename(original_path))[0]
            output_path = os.path.join(output_dir, f"{name_without_ext}_印章.jpg")
            output_image.save(output_path, 'JPEG', quality=95)

            base_image.close()
            stamp.close()
            output_image.close()

            self.logger.debug(f"成功添加印章并保存: {output_path}")
            return True

        except Exception as e:
            self.logger.error(f"处理图片 {original_path} 时发生错误: {str(e)}")
            return False

    def add_stamps_to_images(self, jpg_folder):
        """批量添加印章到JPG文件"""
        self.logger.info("开始添加印章到JPG文件...")
        try:
            output_dir = os.path.join(self.current_dir, "JPGOK")
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
                self.logger.info(f"已创建输出目录: {output_dir}")

            stamp_path = os.path.join(self.current_dir, "yinzhang.png")
            if not os.path.exists(stamp_path):
                self.logger.error(f"找不到印章文件: {stamp_path}")
                raise FileNotFoundError(f"印章文件不存在: {stamp_path}")

            position = (1773, 1678)

            processed_count = 0
            jpg_files = [f for f in os.listdir(jpg_folder) if f.lower().endswith(('.jpg', '.jpeg'))]
            total_files = len(jpg_files)

            for idx, filename in enumerate(jpg_files, 1):
                input_path = os.path.join(jpg_folder, filename)
                self.logger.info(f"[{idx}/{total_files}] 正在处理: {filename}")

                if self.add_stamp_to_image(input_path, stamp_path, position, output_dir, opacity=0.7):
                    processed_count += 1
                    self.logger.info(f"已成功处理: {filename}")
                else:
                    self.logger.error(f"处理失败: {filename}")

            self.logger.info(f"印章添加完成！共成功处理 {processed_count}/{total_files} 个文件")
            self.logger.info(f"处理后的文件保存在: {output_dir}")

        except Exception as e:
            self.logger.error(f"添加印章过程中发生错误：{str(e)}")
            raise


def main():
    logger.info("=== 开始执行文档处理程序 ===")
    logger.info(f"程序启动时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    try:
        # 1. 生成号码归属证明
        logger.info("第1步: 开始生成号码归属证明")
        generator = PhoneCertificateGenerator()
        generator.generate_certificates()

        # 2. 转换Word文档为PDF
        logger.info("第2步: 开始转换Word文档为PDF")
        converter = DocConverter()
        pdf_folder = converter.convert_to_pdf()

        # 3. 转换PDF为JPG
        logger.info("第3步: 开始转换PDF为JPG")
        image_processor = ImageProcessor()
        jpg_folder = image_processor.convert_pdfs_to_jpg(pdf_folder)

        # 4. 添加印章到JPG
        logger.info("第4步: 开始添加印章到JPG")
        image_processor.add_stamps_to_images(jpg_folder)

        logger.info("=== 所有处理已完成！===")

    except Exception as e:
        logger.error(f"处理过程中发生错误：{str(e)}", exc_info=True)
        raise
    finally:
        logger.info(f"程序结束时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")


if __name__ == "__main__":
    main()