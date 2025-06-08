# @Author: blazewind
# @OS: Windows 11
# @Python: 3.12.5
# @Coding: UTF-8
# @功能：批量文档处理工具（文档生成、格式转换、印章添加）
# @创建时间：2025-05-27 15:34:53
# @最后修改：2025-06-01 15:42:00

import os
import sys
import time
import random
import pandas as pd
import numpy as np
from docx import Document
from datetime import datetime, timedelta
from win32com import client
from pdf2image import convert_from_path
from PIL import Image
import logging

# 配置日志
logging.basicConfig(
	level=logging.INFO,
	format='%(asctime)s - %(levelname)s - %(message)s',
	handlers=[
		logging.FileHandler('document_processor.log', encoding='utf-8'),
		logging.StreamHandler()
	]
)
logger = logging.getLogger(__name__)


class DocumentProcessor:
	def __init__(self):
		"""初始化文档处理器"""
		self.current_dir = os.path.dirname(os.path.abspath(__file__))   #//此部分是为了处理腾讯的公司号码证明函的字段
		self.today = datetime.now().strftime("%Y%m%d")
		self.field_mappings = {
			'FAREN': '{FAREN}',
			'ZHIWU': '{ZHIWU}',
			'ID': '{ID}',
			'YEWU': '{YEWU}',
			'PHONE': '{PHONE}',
			'EMAIL': '{EMAIL}'
		}
		self._init_directories()

	def _init_directories(self):
		"""初始化所需的目录"""
		directories = [
			f"{self.today}PDF",
			f"{self.today}JPG",
			"JPGOK",
			"stamps"
		]
		for directory in directories:
			path = os.path.join(self.current_dir, directory)
			if not os.path.exists(path):
				os.makedirs(path)
				logger.info(f"创建目录: {path}")

	def random_date_between(self, start_date, end_date):
		"""生成两个日期之间的随机日期"""
		time_between = end_date - start_date
		days_between = time_between.days
		random_days = random.randrange(days_between)
		return start_date + timedelta(days=random_days)

	def clean_filename(self, filename):
		"""清理文件名中的非法字符"""
		invalid_chars = '<>:"/\\|?*'
		return ''.join(char for char in filename if char not in invalid_chars)

	def get_unique_filename(self, company_value, current_date, repeat_count=None):
		"""生成唯一的文件名"""
		company_name = self.clean_filename(company_value)
		if repeat_count is not None and repeat_count > 0:
			return f'{company_name}_腾讯公司号码证明函_{current_date}_{repeat_count}.docx'
		return f'{company_name}_腾讯公司号码证明函_{current_date}.docx'

	def replace_text_in_paragraph(self, paragraph, replacements):
		"""在段落中替换文本并保持格式"""
		paragraph_text = paragraph.text
		runs = paragraph.runs

		if not any(key in paragraph_text for key in replacements.keys()):
			return

		full_text = ''.join(run.text for run in runs)
		new_text = full_text
		for key, value in replacements.items():
			if key in new_text and value is not None:
				new_text = new_text.replace(key, str(value))

		if runs:
			runs[0].text = new_text
			for run in runs[1:]:
				run.text = ''

	def process_row_data(self, df, row_idx):
		"""处理单行数据，返回所有字段的值"""
		data = {}
		for field, template_field in self.field_mappings.items():
			if field in df.columns:
				value = df.iloc[row_idx].get(field)
				if pd.notna(value):
					data[template_field] = str(value)
				else:
					data[template_field] = ''
		return data

	def validate_excel_data(self, df):
		"""验证Excel数据的有效性"""
		if 'NUMBERS' not in df.columns:
			raise ValueError("Excel文件缺少必需的NUMBERS列")
		if 'COMPANY' not in df.columns:
			raise ValueError("Excel文件缺少必需的COMPANY列")
		if df.empty:
			raise ValueError("Excel文件没有数据")
		return True

	def generate_documents(self):
		"""生成Word文档"""
		logger.info("=== 开始生成Word文档 ===")

		try:
			# 读取Excel文件，指定PHONE列为字符串类型
			excel_file = 'data.xlsx'
			df = pd.read_excel(
				excel_file,
				dtype={'PHONE': str}  # 将PHONE列指定为字符串类型
			)
			self.validate_excel_data(df)

			# 处理NUMBERS列数据
			numbers_series = df['NUMBERS'].dropna()
			numbers_list = numbers_series.astype(str).tolist()

			# 设置日期范围
			start = datetime(2021, 1, 1)
			end = datetime(2025, 3, 31)
			now = datetime.now()

			# 按每10个号码分组处理
			batch_size = 10
			num_batches = (len(numbers_list) + batch_size - 1) // batch_size

			# 获取所有公司名称
			companies = df['COMPANY'].dropna().unique().tolist()
			company_index = 0

			for batch_idx in range(num_batches):
				# 获取当前批次的号码
				start_idx = batch_idx * batch_size
				end_idx = min(start_idx + batch_size, len(numbers_list))
				current_batch = numbers_list[start_idx:end_idx]

				# 将号码用、连接
				merged_numbers = '、'.join(current_batch)

				# 选择公司（循环使用公司列表）
				company_value = companies[company_index]
				company_index = (company_index + 1) % len(companies)

				# 获取公司对应的行索引
				company_row_idx = df[df['COMPANY'] == company_value].index[0]

				# 创建新文档
				doc = Document('cert.doc')  #// 腾讯公司号码证明函的模板文件名
				random_date = self.random_date_between(start, end)

				# 准备替换数据
				replacements = {
					'{NUMBERS}': merged_numbers,
					'{COMPANY}': company_value,
					'{YEAR}': str(now.year),
					'{MONTH}': str(now.month),
					'{DAY}': str(now.day),
					'{rYEAR}': str(random_date.year),
					'{rMONTH}': str(random_date.month),
					'{rDAY}': str(random_date.day)
				}

				# 添加其他字段的值（与COMPANY同行的数据）
				additional_data = self.process_row_data(df, company_row_idx)
				replacements.update(additional_data)

				# 执行文本替换
				for paragraph in doc.paragraphs:
					self.replace_text_in_paragraph(paragraph, replacements)

				# 保存文档
				output_file = self.get_unique_filename(
					company_value,
					self.today,
					batch_idx if batch_idx > 0 else None
				)

				doc.save(output_file)
				logger.info(f'已生成文件：{output_file}')
				logger.info(f'本批次包含 {len(current_batch)} 个号码')
				logger.debug(f'替换的字段：{", ".join(f"{k}={v}" for k, v in replacements.items() if v)}')

		except Exception as e:
			logger.error(f"生成文档时发生错误：{str(e)}")
			raise

	def force_kill_word_processes(self):
		"""Force kill all WINWORD.EXE processes"""
		try:
			# First try graceful termination
			os.system("taskkill /IM WINWORD.EXE /F /T >nul 2>&1")
			time.sleep(1)  # Give processes time to terminate

			# Then check if any processes remain and force kill
			os.system("taskkill /F /IM WINWORD.EXE /T >nul 2>&1")
		except Exception as e:
			logger.warning(f"清理Word进程时发生错误：{str(e)}")

	def word_to_pdf(self, word_app, word_path, pdf_path, max_retries=3):
		"""将Word文档转换为PDF"""
		doc = None
		for attempt in range(max_retries):
			try:
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
				logger.warning(f"第 {attempt + 1} 次尝试转换失败：{str(e)}")
				if attempt < max_retries - 1:
					logger.info("等待后重试...")
					time.sleep(2)
				else:
					logger.error(f"转换失败: {os.path.basename(word_path)}")
					return False
		return False

	def convert_to_pdf(self):
		"""批量转换Word文档为PDF"""
		logger.info("=== 开始转换Word文档为PDF ===")

		pdf_folder = os.path.join(self.current_dir, f"{self.today}PDF")
		files_to_convert = [
			f for f in os.listdir(self.current_dir)
			if (f.endswith(".doc") or f.endswith(".docx"))
			   and f != "cert.doc"
			   and not f.startswith("~$")
		]

		if not files_to_convert:
			logger.warning("没有找到需要转换的Word文档")
			return

		word_app = None
		try:
			# First ensure no Word processes are running
			self.force_kill_word_processes()

			word_app = client.Dispatch("Word.Application")
			word_app.Visible = False
			word_app.DisplayAlerts = False

			total_files = len(files_to_convert)
			converted_count = 0

			for idx, file in enumerate(files_to_convert, 1):
				word_path = os.path.join(self.current_dir, file)
				pdf_name = os.path.splitext(file)[0] + ".pdf"
				pdf_path = os.path.join(pdf_folder, pdf_name)

				logger.info(f"[{idx}/{total_files}] 正在转换: {file}")
				if self.word_to_pdf(word_app, word_path, pdf_path):
					converted_count += 1
					logger.info(f"成功转换: {file} -> {pdf_name}")

			logger.info(f"转换完成！共成功转换 {converted_count}/{total_files} 个文件")

		except Exception as e:
			logger.error(f"PDF转换过程中发生错误：{str(e)}")
			raise
		finally:
			if word_app:
				try:
					# Properly close Word application
					word_app.Quit()
					del word_app
				except:
					pass
				# Force kill any remaining Word processes
				self.force_kill_word_processes()
				time.sleep(1)  # Additional wait to ensure processes are terminated
				logger.info("已清理所有Word进程")

	def pdf_to_jpg(self, pdf_path, output_folder, dpi=300):
		"""将PDF转换为JPG图片"""
		try:
			images = convert_from_path(pdf_path, dpi=dpi)
			pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]

			for i, image in enumerate(images):
				image_path = os.path.join(
					output_folder,
					f"{pdf_name}{'_第' + str(i + 1) + '页' if len(images) > 1 else ''}.jpg"
				)
				image.save(image_path, "JPEG")
			return True
		except Exception as e:
			logger.error(f"转换PDF到JPG失败：{str(e)}")
			return False

	def convert_to_jpg(self):
		"""批量转换PDF为JPG"""
		logger.info("=== 开始转换PDF为JPG ===")

		pdf_dir = os.path.join(self.current_dir, f"{self.today}PDF")
		jpg_folder = os.path.join(self.current_dir, f"{self.today}JPG")

		if not os.path.exists(pdf_dir):
			logger.error(f"PDF目录不存在：{pdf_dir}")
			return

		converted_count = 0
		pdf_files = [f for f in os.listdir(pdf_dir) if f.endswith('.pdf')]
		total_files = len(pdf_files)

		for idx, file in enumerate(pdf_files, 1):
			pdf_path = os.path.join(pdf_dir, file)
			logger.info(f"[{idx}/{total_files}] 正在转换: {file}")

			if self.pdf_to_jpg(pdf_path, jpg_folder):
				converted_count += 1
				logger.info(f"成功转换: {file}")

		logger.info(f"转换完成！共成功转换 {converted_count}/{total_files} 个PDF文件")

	def get_company_from_filename(self, filename):
		"""从文件名中提取公司名称"""
		# 文件名格式：公司名_号码归属证明_日期.jpg 或 公司名_号码归属证明_日期_序号.jpg
		name_without_ext = os.path.splitext(filename)[0]
		parts = name_without_ext.split('_')
		return parts[0]

	def add_stamp_to_image(self, original_path, stamp_path, position, output_dir, opacity=0.7):
		"""为图片添加印章"""
		try:
			# 验证印章文件
			if not os.path.exists(stamp_path):
				raise FileNotFoundError(f"找不到印章文件：{stamp_path}")

			if not stamp_path.lower().endswith('.png'):
				raise ValueError(f"印章文件必须是PNG格式：{stamp_path}")

			# 打开并处理图片
			base_image = Image.open(original_path)
			stamp = Image.open(stamp_path)

			# 确保印章是RGBA模式
			if stamp.mode != 'RGBA':
				stamp = stamp.convert('RGBA')

			# 设置透明度
			alpha = stamp.getchannel('A')
			alpha = alpha.point(lambda x: int(x * opacity))
			stamp.putalpha(alpha)

			# 创建透明图层并粘贴印章
			transparent = Image.new('RGBA', base_image.size, (0, 0, 0, 0))
			transparent.paste(stamp, position, stamp)

			# 合并图层
			output_image = Image.alpha_composite(base_image.convert('RGBA'), transparent)
			output_image = output_image.convert('RGB')

			# 保存处理后的图片
			name_without_ext = os.path.splitext(os.path.basename(original_path))[0]
			output_path = os.path.join(output_dir, f"{name_without_ext}_印章.jpg")
			output_image.save(output_path, 'JPEG', quality=95)

			# 清理资源
			base_image.close()
			stamp.close()
			output_image.close()

			return True

		except (FileNotFoundError, ValueError) as e:
			logger.error(str(e))
			return False
		except Exception as e:
			logger.error(f"添加印章过程中发生错误：{str(e)}")
			return False

	def add_stamps(self):
		"""为所有JPG图片添加印章"""
		logger.info("=== 开始添加印章 ===")

		jpg_dir = os.path.join(self.current_dir, f"{self.today}JPG")
		output_dir = os.path.join(self.current_dir, "JPGOK")
		stamps_dir = os.path.join(self.current_dir, "stamps")

		# 检查必要的目录
		if not os.path.exists(jpg_dir):
			logger.error(f"JPG目录不存在：{jpg_dir}")
			return

		if not os.path.exists(stamps_dir):
			logger.error(f"印章目录不存在：{stamps_dir}")
			return

		position = (1773, 1678)  # 印章位置坐标
		processed_count = 0
		total_count = 0
		missing_stamps = set()

		# 处理所有JPG文件
		jpg_files = [f for f in os.listdir(jpg_dir) if f.lower().endswith(('.jpg', '.jpeg'))]
		total_count = len(jpg_files)

		for filename in jpg_files:
			# 获取公司名称
			company_name = self.get_company_from_filename(filename)
			stamp_path = os.path.join(stamps_dir, f"{company_name}.png")

			# 检查印章文件
			if not os.path.exists(stamp_path):
				missing_stamps.add(company_name)
				logger.error(f"缺少公司的印章文件：{company_name}.png")
				continue

			# 处理图片
			input_path = os.path.join(jpg_dir, filename)
			if self.add_stamp_to_image(input_path, stamp_path, position, output_dir):
				processed_count += 1
				logger.info(f"已处理: {filename}")

		# 报告处理结果
		logger.info(f"处理完成！成功处理 {processed_count}/{total_count} 个文件")
		if missing_stamps:
			logger.warning("\n以下公司缺少印章文件：")
			for company in sorted(missing_stamps):
				logger.warning(f"- {company}.png")
			logger.warning(f"\n请将缺失的印章文件放入目录：{stamps_dir}")


def main():
	"""主函数"""
	logger.info("=== 批量生成资料程序开始运行 ===")
	logger.info("=== Version：Beta01 ===")
	logger.info(f"当前时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
	logger.info(f"请确保准备好资料清单、模板文件、印章文件")

	processor = DocumentProcessor()

	try:
		processor.generate_documents()
		processor.convert_to_pdf()
		processor.convert_to_jpg()
		processor.add_stamps()
		logger.info("=== 所有处理步骤已完成 ===")
	except Exception as e:
		logger.error(f"处理过程中发生错误：{str(e)}")
		sys.exit(1)


if __name__ == "__main__":
	main()