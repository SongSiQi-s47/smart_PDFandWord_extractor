# -*- coding: utf-8 -*-
"""
PDF和Word表格提取工具 - Web版本
保持本地版的所有功能，只去掉input()函数
"""

import os
import re
import logging
import pdfplumber
import pandas as pd
from docx import Document
from typing import List, Dict, Optional, Tuple
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

# 设置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def parse_sample_to_template(sample):
    """将样例转换为模板"""
    if not sample:
        return ""
    
    # 替换数字为占位符
    template = re.sub(r'\d+', 'N', sample)
    # 替换字母为占位符
    template = re.sub(r'[a-zA-Z]+', 'L', template)
    return template

def template_to_regex(template):
    """将模板转换为正则表达式"""
    if not template:
        return ""
    
    # 替换占位符为对应的正则表达式
    regex = template.replace('N', r'\d+')
    regex = regex.replace('L', r'[a-zA-Z]+')
    # 转义特殊字符
    regex = re.escape(regex)
    # 恢复占位符的正则表达式
    regex = regex.replace(r'\\N', r'\d+')
    regex = regex.replace(r'\\L', r'[a-zA-Z]+')
    
    return f"^{regex}$"

def get_regex_from_sample(sample):
    """从样例直接获取正则表达式"""
    template = parse_sample_to_template(sample)
    return template_to_regex(template)

def get_fuzzy_regex_from_sample(sample):
    """获取模糊匹配的正则表达式"""
    if not sample:
        return ""
    
    # 创建更宽松的匹配模式
    pattern = re.escape(sample)
    # 允许数字变化
    pattern = re.sub(r'\\d\+', r'\\d+', pattern)
    # 允许字母变化
    pattern = re.sub(r'[a-zA-Z]', r'[a-zA-Z]', pattern)
    
    return f"^{pattern}"

def smart_start_match(sample, text, regex):
    """智能开始匹配"""
    if not sample or not text:
        return False
    
    # 直接字符串匹配
    if sample in text:
        return True
    
    # 正则表达式匹配
    if regex and re.search(regex, text):
        return True
    
    # 模糊匹配
    fuzzy_regex = get_fuzzy_regex_from_sample(sample)
    if fuzzy_regex and re.search(fuzzy_regex, text):
        return True
    
    return False

class PDFWordTableExtractor:
    def __init__(self):
        # 默认的目标列映射
        self.target_columns = {
            '功能模块': '一级模块名称',
            '功能子项': '二级模块名称', 
            '三级模块': '三级模块名称',
            '功能描述': '功能描述'
        }
        
        # 自定义表头映射
        self.custom_headers = {}
        
        # 状态变量
        self.current_lvl1 = ""
        self.current_lvl2 = ""
        self.current_lvl3 = ""
        self.current_description = []
        self.collected_data = []
        
    def extract_tables(self, file_path: str, file_type: str) -> List[Dict]:
        """提取表格数据的主入口"""
        if file_type == "pdf":
            return self.extract_tables_from_pdf_bid(file_path)
        elif file_type == "docx":
            if "合同" in file_path:
                return self.extract_tables_from_word_contract(file_path)
            else:
                return self.extract_tables_from_word_bid(file_path)
        else:
            return []
    
    def setup_custom_headers(self):
        """设置自定义表头映射"""
        self.custom_headers = {
            '功能模块': '一级模块名称',
            '功能子项': '二级模块名称',
            '三级模块': '三级模块名称',
            '功能描述': '合同描述'
        }
    
    def extract_tables_from_pdf_bid(self, pdf_path: str) -> List[Dict]:
        """提取PDF标书文件中的表格（完整版）"""
        # 清空之前的状态变量
        def clear_previous_state():
            self.current_lvl1 = ""
            self.current_lvl2 = ""
            self.current_lvl3 = ""
            self.current_description = []
            self.collected_data = []
        
        clear_previous_state()
        
        def reclassify_module(text, current_level):
            """重新分类模块"""
            if not text:
                return current_level
            
            # 检查是否是一级模块
            if re.match(r'^[一二三四五六七八九十]+、', text) or re.match(r'^\d+\.\d+\.\d+\.\d+', text):
                return 1
            # 检查是否是二级模块
            elif re.match(r'^[（(][一二三四五六七八九十]+[）)]', text) or re.match(r'^\d+\.\d+\.\d+\.\d+\.\d+', text):
                return 2
            # 检查是否是三级模块
            elif re.match(r'^\d+[、）)]', text) or re.match(r'^\d+\.\d+\.\d+\.\d+\.\d+\.\d+', text):
                return 3
            else:
                return current_level
        
        def should_enable_verification():
            """是否启用验证"""
            return True
        
        def is_page_number(text):
            """检查是否为页码"""
            if not text:
                return False
            page_indicators = ['第', '页', 'Page', 'page']
            return any(indicator in str(text) for indicator in page_indicators)
        
        data = []
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    text = page.extract_text()
                    if not text:
                        continue
                    
                    lines = text.split('\n')
                    current_level = 0
                    
                    for line in lines:
                        line = line.strip()
                        if not line or is_page_number(line):
                            continue
                        
                        # 重新分类模块
                        new_level = reclassify_module(line, current_level)
                        if new_level != current_level:
                            # 保存之前的数据
                            if self.current_description and (self.current_lvl1 or self.current_lvl2 or self.current_lvl3):
                                self.collected_data.append({
                                    '一级模块名称': self.current_lvl1,
                                    '二级模块名称': self.current_lvl2,
                                    '三级模块名称': self.current_lvl3,
                                    '标书描述': ' '.join(self.current_description),
                                    '合同描述': '',
                                    '来源文件': os.path.basename(pdf_path)
                                })
                            
                            # 更新当前级别
                            current_level = new_level
                            self.current_description = []
                            
                            if current_level == 1:
                                self.current_lvl1 = line
                                self.current_lvl2 = ""
                                self.current_lvl3 = ""
                            elif current_level == 2:
                                self.current_lvl2 = line
                                self.current_lvl3 = ""
                            elif current_level == 3:
                                self.current_lvl3 = line
                        else:
                            # 收集描述
                            if current_level > 0:
                                self.current_description.append(line)
                    
                    # 处理页面末尾的数据
                    if self.current_description and (self.current_lvl1 or self.current_lvl2 or self.current_lvl3):
                        self.collected_data.append({
                            '一级模块名称': self.current_lvl1,
                            '二级模块名称': self.current_lvl2,
                            '三级模块名称': self.current_lvl3,
                            '标书描述': ' '.join(self.current_description),
                            '合同描述': '',
                            '来源文件': os.path.basename(pdf_path)
                        })
                        self.current_description = []
        
        except Exception as e:
            logger.error(f"PDF处理错误: {e}")
        
        return self.collected_data
    
    def extract_tables_from_pdf_bid_with_samples(self, pdf_path: str, lvl1_sample: str, 
                                               lvl2_sample: str = "", lvl3_sample: str = "", 
                                               end_sample: str = "") -> List[Dict]:
        """使用编号样例提取PDF标书"""
        def clear_previous_state():
            self.current_lvl1 = ""
            self.current_lvl2 = ""
            self.current_lvl3 = ""
            self.current_description = []
            self.collected_data = []
        
        clear_previous_state()
        
        def reclassify_module(text, current_level):
            """重新分类模块"""
            if not text:
                return current_level
            
            # 使用用户提供的样例进行匹配
            if lvl1_sample and smart_start_match(lvl1_sample, text, get_regex_from_sample(lvl1_sample)):
                return 1
            elif lvl2_sample and smart_start_match(lvl2_sample, text, get_regex_from_sample(lvl2_sample)):
                return 2
            elif lvl3_sample and smart_start_match(lvl3_sample, text, get_regex_from_sample(lvl3_sample)):
                return 3
            else:
                return current_level
        
        def should_enable_verification():
            """是否启用验证"""
            return True
        
        def is_page_number(text):
            """检查是否为页码"""
            if not text:
                return False
            page_indicators = ['第', '页', 'Page', 'page']
            return any(indicator in str(text) for indicator in page_indicators)
        
        def should_collect_description():
            """是否应该收集描述"""
            return self.current_lvl1 or self.current_lvl2 or self.current_lvl3
        
        data = []
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    text = page.extract_text()
                    if not text:
                        continue
                    
                    lines = text.split('\n')
                    current_level = 0
                    
                    for line in lines:
                        line = line.strip()
                        if not line or is_page_number(line):
                            continue
                        
                        # 检查是否遇到结束标记
                        if end_sample and end_sample in line:
                            break
                        
                        # 重新分类模块
                        new_level = reclassify_module(line, current_level)
                        if new_level != current_level:
                            # 保存之前的数据
                            if self.current_description and should_collect_description():
                                self.collected_data.append({
                                    '一级模块名称': self.current_lvl1,
                                    '二级模块名称': self.current_lvl2,
                                    '三级模块名称': self.current_lvl3,
                                    '标书描述': ' '.join(self.current_description),
                                    '合同描述': '',
                                    '来源文件': os.path.basename(pdf_path)
                                })
                            
                            # 更新当前级别
                            current_level = new_level
                            self.current_description = []
                            
                            if current_level == 1:
                                self.current_lvl1 = line
                                self.current_lvl2 = ""
                                self.current_lvl3 = ""
                            elif current_level == 2:
                                self.current_lvl2 = line
                                self.current_lvl3 = ""
                            elif current_level == 3:
                                self.current_lvl3 = line
                        else:
                            # 收集描述
                            if should_collect_description():
                                self.current_description.append(line)
                    
                    # 处理页面末尾的数据
                    if self.current_description and should_collect_description():
                        self.collected_data.append({
                            '一级模块名称': self.current_lvl1,
                            '二级模块名称': self.current_lvl2,
                            '三级模块名称': self.current_lvl3,
                            '标书描述': ' '.join(self.current_description),
                            '合同描述': '',
                            '来源文件': os.path.basename(pdf_path)
                        })
                        self.current_description = []
        
        except Exception as e:
            logger.error(f"PDF处理错误: {e}")
        
        return self.collected_data
    
    def extract_tables_from_pdf_contract_with_samples(self, pdf_path: str, lvl1_sample: str, 
                                                    lvl2_sample: str = "", lvl3_sample: str = "", 
                                                    end_sample: str = "") -> List[Dict]:
        """使用编号样例提取PDF合同"""
        data = self.extract_tables_from_pdf_bid_with_samples(pdf_path, lvl1_sample, lvl2_sample, lvl3_sample, end_sample)
        
        # 调整描述字段映射
        for item in data:
            if item['标书描述']:
                item['合同描述'] = item['标书描述']
                item['标书描述'] = ''
        
        return data
    
    def extract_tables_from_word_contract(self, docx_path: str, original_filename: str = None) -> List[Dict]:
        """提取Word合同文件中的表格（完整版）"""
        data = []
        found_quotation_section = False
        current_headers = None
        doc = Document(docx_path)
        
        # 如果没有提供原始文件名，使用docx_path
        if not original_filename:
            original_filename = os.path.basename(docx_path)
        
        # 先检查整个文档是否包含分项报价表
        has_quotation_table = False
        for para in doc.paragraphs:
            if "分项报价表" in para.text:
                has_quotation_table = True
                break
        
        if not has_quotation_table:
            return data
        
        # 处理所有表格
        for table_idx, table in enumerate(doc.tables):
            rows = list(table.rows)
            if not rows:
                continue
            
            # 检查表头
            headers = [cell.text.strip().replace('\n', '') for cell in rows[0].cells]
            
            # 检查是否是目标表格
            if self._is_target_table_custom(headers):
                current_headers = headers
                found_quotation_section = True
                start_row = 1
                print(f"✅ 找到匹配的表格，表头：{headers}")
            elif current_headers and found_quotation_section:
                start_row = 0
            else:
                continue
                
            # 处理数据行
            for row_idx, row in enumerate(rows[start_row:], start=start_row):
                row_data = {}
                cells = row.cells
                
                # 处理合并单元格的情况
                for idx, header in enumerate(current_headers):
                    if idx < len(cells):
                        cell_text = cells[idx].text.strip()
                        row_data[header] = cell_text
                    else:
                        row_data[header] = ''
                
                # 检查是否有有效数据
                has_data = False
                
                # 检查序号字段
                for header in current_headers:
                    if '序号' in header and row_data.get(header, '').strip():
                        try:
                            int(row_data[header])
                            has_data = True
                            break
                        except ValueError:
                            pass
                
                # 如果序号不是数字，检查其他关键字段
                if not has_data:
                    key_fields = self._get_key_fields_for_check()
                    for field in key_fields:
                        for header in current_headers:
                            if field in header and row_data.get(header, '').strip():
                                has_data = True
                                break
                        if has_data:
                            break
                
                # 如果关键字段都没有，再检查其他字段
                if not has_data:
                    has_data = any(row_data.get(header, '').strip() for header in current_headers)
                
                # 添加调试信息
                if has_data:
                    mapped = self._map_word_row_custom(row_data, docx_path, original_filename)
                    
                    # 检查重复并处理
                    if len(data) > 0:
                        # 查找上一个非空的一级模块名称
                        last_lvl1 = ""
                        last_lvl2 = ""
                        for i in range(len(data) - 1, -1, -1):
                            if data[i]['一级模块名称'].strip():
                                last_lvl1 = data[i]['一级模块名称']
                                break
                        for i in range(len(data) - 1, -1, -1):
                            if data[i]['二级模块名称'].strip():
                                last_lvl2 = data[i]['二级模块名称']
                                break
                        
                        # 检查一级模块名称是否重复
                        if mapped['一级模块名称'].strip() == last_lvl1.strip() and mapped['一级模块名称'].strip():
                            # 如果一级模块名称重复，清空当前的一级模块名称
                            mapped['一级模块名称'] = ""
                        
                        # 检查二级模块名称是否重复
                        if mapped['二级模块名称'].strip() == last_lvl2.strip() and mapped['二级模块名称'].strip():
                            # 如果二级模块名称重复，清空当前的二级模块名称
                            mapped['二级模块名称'] = ""
                    
                    data.append(mapped)
        
        return data
    
    def extract_tables_from_word_bid(self, docx_path: str) -> List[Dict]:
        """提取Word标书文件中的表格"""
        data = self.extract_tables_from_word_contract(docx_path)
        
        # 调整描述字段映射
        for item in data:
            if item['合同描述']:
                item['标书描述'] = item['合同描述']
                item['合同描述'] = ''
        
        return data
    
    def _is_target_table_custom(self, headers: List[str]) -> bool:
        """检查是否是目标表格（自定义表头）"""
        if self.custom_headers:
            # 使用自定义表头
            found_headers = 0
            for custom_header in self.custom_headers.keys():
                for header in headers:
                    if custom_header in header:
                        found_headers += 1
                        break
            return found_headers >= 2
        else:
            # 使用默认表头
            return self._is_target_table(headers)
    
    def _get_key_fields_for_check(self) -> List[str]:
        """获取用于检查的关键字段列表"""
        if self.custom_headers:
            # 使用自定义表头
            return list(self.custom_headers.keys())
        else:
            # 使用默认表头
            return ['功能描述', '三级模块', '功能模块', '功能子项']
    
    def _map_word_row_custom(self, row_data: Dict, source_file: str, original_filename: str = None) -> Dict:
        """使用自定义表头映射Word行数据"""
        if self.custom_headers:
            # 添加调试信息
            print(f"🔍 自定义表头映射调试:")
            print(f"   自定义表头: {self.custom_headers}")
            print(f"   原始行数据: {row_data}")
            
            # 使用自定义表头映射
            mapped_data = {}
            for header, value in row_data.items():
                for custom_header, standard_field in self.custom_headers.items():
                    if custom_header in header:
                        mapped_data[standard_field] = value
                        print(f"   ✅ 匹配: {header} -> {standard_field} = {value}")
                        break
            
            print(f"   映射结果: {mapped_data}")
            
            # 确保Word合同内容映射到"合同描述"
            desc_value = mapped_data.get('合同描述', '')
            
            # 如果合同描述为空，尝试从原始数据中查找
            if not desc_value:
                print(f"   ⚠️ 合同描述为空，尝试查找其他描述字段...")
                for header, value in row_data.items():
                    if value.strip() and ('描述' in header or '备注' in header or '内容' in header or '功能' in header):
                        desc_value = value
                        print(f"   ✅ 找到描述字段: {header} = {value}")
                        break
            
            mapped = {
                '一级模块名称': mapped_data.get('一级模块名称', ''),
                '二级模块名称': mapped_data.get('二级模块名称', ''),
                '三级模块名称': mapped_data.get('三级模块名称', ''),
                '标书描述': '',  # Word合同文件，标书描述为空
                '合同描述': desc_value,  # Word合同内容放这里
                '来源文件': original_filename if original_filename else (os.path.basename(source_file) if not source_file.endswith('tmp') else '合同.docx'),
            }
            
            print(f"   最终映射: {mapped}")
        else:
            # 使用默认映射，但增强合同描述的处理
            mapped = self._map_word_row(row_data, source_file)
            
            # 如果合同描述为空，尝试从原始数据中查找
            if not mapped['合同描述']:
                for header, value in row_data.items():
                    if value.strip() and ('描述' in header or '备注' in header or '内容' in header or '功能' in header):
                        mapped['合同描述'] = value
                        break
        
        return mapped
    
    def _is_target_table(self, headers: List[str]) -> bool:
        """检查是否是目标表格（默认表头）"""
        target_headers = list(self.target_columns.keys())
        found_headers = []
        
        # 预处理目标字段名，去掉换行符
        normalized_target_headers = {}
        for target in target_headers:
            normalized_target = target.replace('\n', '').replace('\\n', '')
            normalized_target_headers[normalized_target] = target
        
        for normalized_target, original_target in normalized_target_headers.items():
            for header_cell in headers:
                # 对实际字段名也进行预处理
                normalized_header = header_cell.replace('\n', '').replace('\\n', '')
                if normalized_target in normalized_header:
                    found_headers.append(original_target)
                    break
        
        return len(found_headers) >= 2
    
    def _map_word_row(self, row_data: Dict, source_file: str) -> Dict:
        """映射Word行数据到标准格式"""
        # 创建标准化字段映射
        field_mapping = {}
        for original_field in ['功能模块', '功能子项', '三级模块', '功能描述']:
            normalized_field = original_field.replace('\n', '').replace('\\n', '')
            field_mapping[normalized_field] = original_field
        
        # 查找匹配的字段
        mapped_data = {}
        for header, value in row_data.items():
            normalized_header = header.replace('\n', '').replace('\\n', '')
            for normalized_field, original_field in field_mapping.items():
                if normalized_field in normalized_header:
                    mapped_data[original_field] = value
                    break
        
        mapped = {
            '一级模块名称': mapped_data.get('功能模块', ''),
            '二级模块名称': mapped_data.get('功能子项', ''),
            '三级模块名称': mapped_data.get('三级模块', ''),
            '标书描述': '',
            '合同描述': '',
            '来源文件': os.path.basename(source_file) if not source_file.endswith('tmp') else '合同.docx',
        }
        
        # 根据文件类型决定内容放哪个字段
        if "标书" in source_file:
            mapped['标书描述'] = mapped_data.get('功能描述', '')
        elif "合同" in source_file:
            mapped['合同描述'] = mapped_data.get('功能描述', '')
        
        return mapped
    
    def _merge_paragraphs(self, desc_lines):
        """合并自然段"""
        desc_paragraphs = []
        current_para = []
        for line in desc_lines:
            if not line.strip():
                if current_para:
                    desc_paragraphs.append(''.join(current_para))
                    current_para = []
            else:
                current_para.append(line.strip())
        if current_para:
            desc_paragraphs.append(''.join(current_para))
        return desc_paragraphs
    
    def _clean_extracted_data(self, data: List[Dict]) -> List[Dict]:
        """清理提取的数据，过滤无效内容"""
        cleaned_data = []
        
        for item in data:
            # 检查是否为页码信息
            def is_page_content(text):
                if not text:
                    return False
                page_indicators = ['第', '页', 'Page', 'page']
                return any(indicator in str(text) for indicator in page_indicators)
            
            # 清理各字段
            cleaned_item = {}
            for key, value in item.items():
                if is_page_content(value):
                    print(f"DEBUG: 过滤页码内容 '{value}' 从字段 '{key}'")
                    cleaned_item[key] = ""
                else:
                    cleaned_item[key] = value
            
            # 只保留有效数据
            if any([cleaned_item.get("一级模块名称"), cleaned_item.get("二级模块名称"), 
                   cleaned_item.get("三级模块名称"), cleaned_item.get("标书描述"), cleaned_item.get("合同描述")]):
                cleaned_data.append(cleaned_item)
        
        return cleaned_data
    
    def create_excel_output(self, data: List[Dict], output_path: str, append_mode=False):
        """创建Excel输出文件"""
        if not data:
            logger.warning("没有数据需要输出")
            return None
        
        # 新增：智能分段处理函数
        def split_long_description(description, max_length=500):
            """智能分段处理长描述"""
            if not description or len(description) <= max_length:
                return [description]
            
            # 预定义编号格式正则表达式
            number_patterns = [
                # 中文数字格式
                r'[一二三四五六七八九十]+、',  # 一、二、三、
                r'[（(][一二三四五六七八九十]+[）)]',  # （一）（二）
                r'[（(]\d+[）)]',  # （1）（2）
                r'\d+、',  # 1、2、
                r'\d+\.',  # 1.2.3.
                r'\d+）',  # 1）2）
                r'\d+\)',  # 1)2)
                
                # 英文数字格式
                r'\d+\.',  # 1.2.3.
                r'[a-z]+\)',  # a)b)c)
                r'[A-Z]+\.',  # A.B.C.
                r'[i]+\)',  # i)ii)iii)
            ]
            
            # 尝试按编号分割
            for pattern in number_patterns:
                parts = re.split(f'({pattern})', description)
                if len(parts) > 1:
                    # 重新组合分割的部分
                    result = []
                    current_part = ""
                    for i, part in enumerate(parts):
                        if re.match(pattern, part):
                            if current_part:
                                result.append(current_part.strip())
                            current_part = part
                        else:
                            current_part += part
                    
                    if current_part:
                        result.append(current_part.strip())
                    
                    # 如果分割后的部分仍然太长，继续分割
                    final_result = []
                    for part in result:
                        if len(part) > max_length:
                            # 按句号分割
                            sentences = re.split(r'[。！？]', part)
                            current_sentence = ""
                            for sentence in sentences:
                                if len(current_sentence + sentence) <= max_length:
                                    current_sentence += sentence + "。"
                                else:
                                    if current_sentence:
                                        final_result.append(current_sentence.strip())
                                    current_sentence = sentence + "。"
                            if current_sentence:
                                final_result.append(current_sentence.strip())
                        else:
                            final_result.append(part)
                    
                    return final_result
            
            # 如果没有找到编号格式，按句号分割
            sentences = re.split(r'[。！？]', description)
            result = []
            current_sentence = ""
            for sentence in sentences:
                if len(current_sentence + sentence) <= max_length:
                    current_sentence += sentence + "。"
                else:
                    if current_sentence:
                        result.append(current_sentence.strip())
                    current_sentence = sentence + "。"
            if current_sentence:
                result.append(current_sentence.strip())
            
            return result if result else [description]
        
        def get_unique_filename(filepath):
            """获取唯一文件名"""
            if not os.path.exists(filepath):
                return filepath
            
            base, ext = os.path.splitext(filepath)
            counter = 1
            while os.path.exists(f"{base}_{counter}{ext}"):
                counter += 1
            return f"{base}_{counter}{ext}"
        
        def clean_cell_value(value):
            """清理单元格值"""
            if not value:
                return ""
            
            # 移除多余的空白字符
            cleaned = re.sub(r'\s+', ' ', str(value).strip())
            # 移除特殊字符
            cleaned = re.sub(r'[^\w\s\u4e00-\u9fff.,，。！？：:()（）\-]', '', cleaned)
            return cleaned
        
        # 清理数据
        cleaned_data = self._clean_extracted_data(data)
        
        if not cleaned_data:
            logger.warning("清理后没有数据需要输出")
            return None
        
        # 处理长描述
        processed_data = []
        for item in cleaned_data:
            processed_item = item.copy()
            
            # 处理标书描述
            if processed_item.get('标书描述'):
                desc_parts = split_long_description(processed_item['标书描述'])
                for i, part in enumerate(desc_parts):
                    if i == 0:
                        processed_item['标书描述'] = part
                    else:
                        # 创建新行
                        new_item = processed_item.copy()
                        new_item['标书描述'] = part
                        processed_data.append(new_item)
                        processed_item = None
                        break
                if processed_item:
                    processed_data.append(processed_item)
            else:
                processed_data.append(processed_item)
        
        # 创建DataFrame
        df = pd.DataFrame(processed_data)
        
        # 确保列顺序正确
        column_order = ['一级模块名称', '二级模块名称', '三级模块名称', '标书描述', '合同描述', '来源文件']
        for col in column_order:
            if col not in df.columns:
                df[col] = ''
        df = df[column_order]
        
        # 清理单元格值
        for col in df.columns:
            df[col] = df[col].apply(clean_cell_value)
        
        # 获取唯一文件名
        output_path = get_unique_filename(output_path)
        
        # 写入Excel
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='提取结果', index=False)
            worksheet = writer.sheets['提取结果']
            
            # 设置列宽
            column_widths = {'A': 15, 'B': 15, 'C': 15, 'D': 45, 'E': 45, 'F': 10}
            for col, width in column_widths.items():
                worksheet.column_dimensions[col].width = width
            
            # 设置表头样式
            header_font = Font(bold=True, size=12)
            header_alignment = Alignment(horizontal='center', vertical='center')
            for col in range(1, len(column_order) + 1):
                cell = worksheet.cell(row=1, column=col)
                cell.font = header_font
                cell.alignment = header_alignment
            
            # 设置数据行样式和行高
            data_alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
            for row in range(2, len(df) + 2):
                # 计算每行最大字符数
                max_chars = 0
                for col in range(1, len(column_order) + 1):
                    cell_value = str(worksheet.cell(row=row, column=col).value or '')
                    lines = cell_value.split('\n')
                    for line in lines:
                        line_chars = len(line)
                        if line_chars > 30:
                            needed_lines = (line_chars // 30) + 1
                            max_chars = max(max_chars, needed_lines * 30)
                        else:
                            max_chars = max(max_chars, line_chars)
                
                # 计算行高
                estimated_lines = max(1, (max_chars // 30) + 1)
                row_height = max(20, estimated_lines * 18 + 10)
                worksheet.row_dimensions[row].height = row_height
                
                # 设置单元格对齐方式
                for col in range(1, len(column_order) + 1):
                    cell = worksheet.cell(row=row, column=col)
                    cell.alignment = data_alignment
            
            # 设置边框
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            for row in worksheet.iter_rows(min_row=1, max_row=len(df) + 1, min_col=1, max_col=len(column_order)):
                for cell in row:
                    cell.border = thin_border
        
        return output_path

def main():
    """主函数（保留用于测试）"""
    extractor = PDFWordTableExtractor()
    print("PDF和Word表格提取工具 - Web版本")
    print("请通过Streamlit应用使用此工具")

if __name__ == "__main__":
    main()
    