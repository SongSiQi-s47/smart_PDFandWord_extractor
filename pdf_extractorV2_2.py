import os
import re
import logging
import pdfplumber
import pandas as pd
from typing import List, Dict
from openpyxl.styles import Font, Alignment, Border, Side
from docx import Document
from collections import OrderedDict

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 汉字数字
CN_NUM = '零一二三四五六七八九十百千万亿〇壹贰叁肆伍陆柒捌玖拾'

def parse_sample_to_template(sample):
    template = []
    i = 0
    while i < len(sample):
        c = sample[i]
        if c.isdigit():
            num = c
            while i+1 < len(sample) and sample[i+1].isdigit():
                i += 1
                num += sample[i]
            template.append(('digit', num))
        elif c in CN_NUM:
            cn = c
            while i+1 < len(sample) and sample[i+1] in CN_NUM:
                i += 1
                cn += sample[i]
            template.append(('cndigit', cn))
        elif c in '()（）':
            template.append(('paren', c))
        elif c in '【】':
            template.append(('bracket', c))
        elif c in '.、.':
            template.append(('sep', c))
        elif '\u4e00' <= c <= '\u9fa5':
            word = c
            while i+1 < len(sample) and '\u4e00' <= sample[i+1] <= '\u9fa5':
                i += 1
                word += sample[i]
            template.append(('ch', word))
        else:
            template.append(('other', c))
        i += 1
    return template

def template_to_regex(template):
    regex = ''
    for t, val in template:
        if t == 'digit':
            regex += r'\d+'
        elif t == 'cndigit':
            regex += f'[{CN_NUM}]+'
        elif t == 'paren':
            regex += re.escape(val)
        elif t == 'bracket':
            regex += re.escape(val)
        elif t == 'sep':
            regex += re.escape(val)
        elif t == 'ch':
            regex += f'.{{1,{len(val)+3}}}'
        else:
            regex += re.escape(val)
    return regex

def get_regex_from_sample(sample):
    template = parse_sample_to_template(sample)
    regex = template_to_regex(template)
    return re.compile(regex)

def get_fuzzy_regex_from_sample(sample):
    """
    生成更灵活的正则表达式，支持各种编号格式
    """
    # 匹配"数字.数字.数字."结构（如 9.1.4.3.1.），更灵活
    if re.match(r'^\d+(\.\d+)+\.?$', sample):
        # 返回带分组的正则，编号部分更灵活，支持末尾没有点的情况
        # 新增：记录原始样例的数字长度，用于后续长度检查
        sample_digits = re.sub(r'[^\d]', '', sample)
        regex = re.compile(r'^(\d+(?:\.\d+)+[\.\s\u3000．、]*)(.*)$')
        # 返回字典，包含正则和长度信息
        return {
            'regex': regex,
            'expected_digit_length': len(sample_digits)
        }
    
    # 匹配"数字）"或"数字)"结构（如 1） 或 1) ），支持中英文括号
    if re.match(r'^\d+[）)]$', sample):
        return {
            'regex': re.compile(r'^(\d+[\)\）])(.*)$'),
            'expected_digit_length': None
        }
    
    # 匹配"（数字）"或"(数字)"结构（如 （1） 或 (1) ），支持中英文括号
    if re.match(r'^[（(]\d+[）)]$', sample):
        return {
            'regex': re.compile(r'^([（(]\d+[\)\）])(.*)$'),
            'expected_digit_length': None
        }
    
    # 匹配"（汉字数字）"或"(汉字数字)"结构（如 （十一） 或 (十一) ），支持更多汉字数字
    if re.match(r'^[（(][零一二三四五六七八九十百千万亿]+[）)]$', sample):
        return {
            'regex': re.compile(r'^([（(][零一二三四五六七八九十百千万亿]+[\)\）])(.*)$'),
            'expected_digit_length': None
        }
    
    # 新增：匹配"数字."结构（如 1.），生成更严格的正则
    if re.match(r'^\d+\.$', sample):
        return {
            'regex': re.compile(r'^(\d+\.)(.*)$'),
            'expected_digit_length': None
        }
    
    # 其他类型：先去除所有空白字符（包括全角空格），再用模板解析生成正则
    sample = re.sub(r'[\s\u3000]', '', sample)
    template = parse_sample_to_template(sample)
    regex = template_to_regex(template)
    regex = regex.rstrip(r'\.')
    regex += r'[\.\s\u3000．、]*'  # 允许编号后有点、空格、顿号等
    return {
        'regex': re.compile(f'^({regex})(.*)$'),
        'expected_digit_length': None
    }

def smart_start_match(sample, text, regex):
    """
    智能起始编号匹配，支持多种匹配策略
    """
    match = regex.match(text)
    if not match:
        return False, None, None
    
    # 提取数字序列
    sample_digits = re.sub(r'[^\d]', '', sample)
    text_number_part = match.group(1)
    text_digits = re.sub(r'[^\d]', '', text_number_part)
    
    # 多种匹配策略
    if sample_digits == text_digits:
        return True, "完全匹配", text_digits
    elif text_digits.startswith(sample_digits):
        return True, "前缀匹配", text_digits
    elif sample_digits.startswith(text_digits):
        # 前缀匹配时检查长度是否一致
        if len(text_digits) == len(sample_digits):
            return True, "前缀匹配", text_digits
        else:
            return False, "前缀长度不匹配", text_digits
    
    return False, "不匹配", text_digits

class PDFWordTableExtractor:
    def __init__(self):
        # 默认的目标列映射
        self.target_columns = {
            '功能模块': '一级模块名称',
            '功能子项': '二级模块名称',
            '三级模块': '三级模块名称',
            '功能描述': '功能描述'
        }
        
        # 用户自定义的表头映射（用于合同文件）
        self.custom_headers = None
        
        # 初始化提取状态变量
        self.previous_lvl1_sample = None
        self.previous_lvl2_sample = None
        self.previous_lvl3_sample = None
        self.previous_end_sample = None
        self.previous_lvl1_regex = None
        self.previous_lvl2_regex = None
        self.previous_lvl3_regex = None
        self.previous_end_regex = None

    def extract_tables(self, file_path: str, file_type: str) -> List[Dict]:
        if file_type == "pdf":
            if "合同" in file_path:
                return self.extract_tables_from_pdf_contract(file_path)
            else:
                return self.extract_tables_from_pdf_bid(file_path)
        elif file_type == "docx":
            if "合同" in file_path:
                # 对于合同文件，先设置自定义表头
                if not self.custom_headers:
                    self.setup_custom_headers()
                return self.extract_tables_from_word_contract(file_path)
            else:
                return self.extract_tables_from_word_bid(file_path)
        else:
            print(f"暂不支持的文件类型: {file_path}")
            return []
            
    def setup_custom_headers(self):
        """设置用户自定义的表头映射"""
        print("\n" + "="*50)
        print("📋 请根据Word文档中的实际表头输入对应的字段名")
        print("="*50)
        print("提示：如果某个字段在Word文档中没有，请直接回车跳过")
        print()
        
        custom_headers = {}
        
        # 一级模块名称
        lvl1_header = input("请输入Word文档中对应'一级模块名称'的表头（如：功能模块、模块名称等）：").strip()
        if lvl1_header:
            custom_headers[lvl1_header] = '一级模块名称'
        
        # 二级模块名称
        lvl2_header = input("请输入Word文档中对应'二级模块名称'的表头（如：功能子项、子模块等）：").strip()
        if lvl2_header:
            custom_headers[lvl2_header] = '二级模块名称'
        
        # 三级模块名称
        lvl3_header = input("请输入Word文档中对应'三级模块名称'的表头（如：三级模块、子项等）：").strip()
        if lvl3_header:
            custom_headers[lvl3_header] = '三级模块名称'
        
        # 合同描述
        desc_header = input("请输入Word文档中对应'合同描述'的表头（如：功能描述、描述、备注等）：").strip()
        if desc_header:
            custom_headers[desc_header] = '合同描述'
        
        if custom_headers:
            self.custom_headers = custom_headers
            print(f"\n✅ 已设置自定义表头映射：")
            for word_header, standard_field in custom_headers.items():
                print(f"   {word_header} → {standard_field}")
        else:
            print("\n⚠️ 未设置任何自定义表头，将使用默认映射")
            self.custom_headers = None
        
        return custom_headers

    # ----------- PDF 标书（新版分层提取） -----------
    def extract_tables_from_pdf_bid(self, pdf_path: str) -> List[Dict]:
        # 清空之前的状态变量
        def clear_previous_state():
            """清空之前的提取状态"""
            self.previous_lvl1_sample = None
            self.previous_lvl2_sample = None
            self.previous_lvl3_sample = None
            self.previous_end_sample = None
            self.previous_lvl1_regex = None
            self.previous_lvl2_regex = None
            self.previous_lvl3_regex = None
            self.previous_end_regex = None
            # print("DEBUG: 已清空之前的提取状态")
        
        # 调用状态清空
        clear_previous_state()
        
        # 新增：重新分类模块层级（移到函数开头）
        def reclassify_module(text, current_level):
            """重新分类模块层级"""
            if current_level == 3 and has_lvl3_sample:
                # 检查是否应该归为二级
                if lvl2_regex and lvl2_regex.match(text):
                    # if debug_enabled:
                    #     print(f"DEBUG: 三级模块重新分类为二级: '{text}'")
                    return 2
                # 检查是否应该归为一级
                elif lvl1_regex and lvl1_regex.match(text):
                    # if debug_enabled:
                    #     print(f"DEBUG: 三级模块重新分类为一级: '{text}'")
                    return 1
            
            elif current_level == 2 and not has_lvl3_sample:
                # 检查是否应该归为一级
                if lvl1_regex and lvl1_regex.match(text):
                    # if debug_enabled:
                    #     print(f"DEBUG: 二级模块重新分类为一级: '{text}'")
                    return 1
            
            return current_level
        
        # 新增：判断是否启用重新核验（移到函数开头）
        def should_enable_verification():
            """判断是否启用重新核验"""
            if not lvl1_sample or not lvl2_sample:
                return False
            
            lvl1_len = len(lvl1_sample)
            lvl2_len = len(lvl2_sample)
            
            # 特殊情况：如果二级样例包含更多点号，认为二级更长
            lvl1_dots = lvl1_sample.count('.')
            lvl2_dots = lvl2_sample.count('.')
            
            if lvl2_dots > lvl1_dots:
                # if debug_enabled:
                #     print(f"DEBUG: 启用重新核验 - 二级点号({lvl2_dots}) > 一级点号({lvl1_dots})")
                return True
            
            # 一般情况：比较字符串长度
            if lvl2_len < lvl1_len:
                # if debug_enabled:
                #     print(f"DEBUG: 启用重新核验 - 二级长度({lvl2_len}) < 一级长度({lvl1_len})")
                return True
            
            return False
        
        # 添加页码过滤函数
        def is_page_number(text):
            """判断是否为页码信息"""
            page_patterns = [
                r'^第\d+页$',
                r'^Page\s*\d+$',
                r'^-\s*\d+\s*-$',
            ]
            # 更严格的纯数字页码判断
            if re.match(r'^\d+$', text.strip()):
                # 如果数字小于等于3位数，且前后没有其他内容，可能是页码
                if len(text.strip()) <= 3:
                    return True
            return any(re.match(pattern, text.strip()) for pattern in page_patterns)
        
        print("请输入各级编号样例（可回车跳过二级/三级）：")
        lvl1_sample = input("一级模块编号样例（如9.1.3.4或（一））：").strip()
        lvl2_sample = input("二级模块编号样例（如9.1.3.4.1或（二），可跳过）：").strip()
        lvl3_sample = input("三级模块编号样例（如1），可跳过）：").strip()
        end_sample = input("终止编号样例（遇到该编号停止提取）：").strip()

        lvl1_regex_info = get_fuzzy_regex_from_sample(lvl1_sample) if lvl1_sample else None
        lvl2_regex_info = get_fuzzy_regex_from_sample(lvl2_sample) if lvl2_sample else None
        lvl3_regex_info = get_fuzzy_regex_from_sample(lvl3_sample) if lvl3_sample else None
        end_regex_info = get_fuzzy_regex_from_sample(end_sample) if end_sample else None

        lvl1_regex = lvl1_regex_info['regex'] if lvl1_regex_info else None
        lvl2_regex = lvl2_regex_info['regex'] if lvl2_regex_info else None
        lvl3_regex = lvl3_regex_info['regex'] if lvl3_regex_info else None
        end_regex = end_regex_info['regex'] if end_regex_info else None

        results = []
        current_lvl1 = current_lvl2 = current_lvl3 = None
        last_lvl1 = last_lvl2 = last_lvl3 = None
        desc_lines = []
        extracting = False
        start_found = False
        in_lvl3 = False
        in_lvl2 = False

        lvl1_filled = False
        lvl2_filled = False
        lvl1_to_fill = ""
        lvl2_to_fill = ""

        # 关键：判断是否有三级模块样例
        has_lvl3_sample = bool(lvl3_sample)

        # 补充回来：页面和行数统计
        page_count = 0
        line_count = 0
        
        # 新增：DEBUG控制变量
        # debug_enabled = False

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_num = page.page_number if hasattr(page, 'page_number') else pdf.pages.index(page)
                
                # 补充回来：页面计数
                page_count += 1
                
                lines = (page.extract_text() or '').split('\n')
                for i, raw_text in enumerate(lines):
                    # 补充回来：行数计数
                    line_count += 1
                    
                    # 处理原始文本
                    text = re.sub(r'[\s\u3000]', '', raw_text)

                    # 终止编号判断
                    if end_regex and end_regex.match(text):
                        is_match, match_type, actual_digits = smart_start_match(end_sample, text, end_regex)
                        if is_match:
                            # if debug_enabled:
                            #     print(f"DEBUG: 识别到终止编号: '{raw_text.strip()}'")
                            extracting = False
                            in_lvl3 = False
                            in_lvl2 = False
                            continue

                    # 智能起始编号判断
                    if not extracting and lvl1_sample and lvl1_regex and lvl1_regex.match(text):
                        is_match, match_type, actual_digits = smart_start_match(lvl1_sample, text, lvl1_regex)
                        if is_match:
                            extracting = True
                            start_found = True
                            # debug_enabled = True  # 启用DEBUG
                            current_lvl1 = raw_text.strip()
                            lvl1_filled = False
                            lvl1_to_fill = current_lvl1
                            lvl2_to_fill = current_lvl2 if current_lvl2 else ""
                            in_lvl2 = False
                            # print(f"DEBUG: 识别到起始编号，开始提取 - 一级模块: '{current_lvl1}'")

                    if not extracting:
                        continue

                    # 先判断三级
                    m3 = lvl3_regex.match(text) if lvl3_regex else None
                    if m3:
                        # if debug_enabled:
                        #     print(f"DEBUG: 三级模块匹配 - 原始文本: '{raw_text.strip()}', 处理后: '{text}'")
                        
                        # 新增：验证三级模块匹配的有效性
                        def is_valid_lvl3_match(match_obj, original_text):
                            """验证三级模块匹配是否有效"""
                            if not match_obj:
                                return False
                            
                            # 提取匹配的编号部分
                            number_part = match_obj.group(1)
                            title_part = match_obj.group(2).strip()
                            
                            # if debug_enabled:
                            #     print(f"DEBUG: 三级模块验证 - 编号部分: '{number_part}', 标题部分: '{title_part}'")
                            
                            # 检查标题部分是否有内容
                            if not title_part:
                                # if debug_enabled:
                                #     print(f"DEBUG: 三级模块标题为空: '{original_text}'")
                                return False
                            
                            return True
                        
                        # 验证匹配有效性
                        if not is_valid_lvl3_match(m3, raw_text.strip()):
                            # if debug_enabled:
                            #     print(f"DEBUG: 三级模块验证失败，设为None")
                            m3 = None
                        
                        if lvl3_regex_info and lvl3_regex_info['expected_digit_length']:
                            text_digits = re.sub(r'[^\d]', '', m3.group(1))
                            if len(text_digits) != lvl3_regex_info['expected_digit_length']:
                                                            # if debug_enabled:
                            #     print(f"DEBUG: 三级模块数字长度不匹配: 期望{lvl3_regex_info['expected_digit_length']}, 实际{len(text_digits)}")
                                m3 = None

                        if m3:
                            # if debug_enabled:
                            #     print(f"DEBUG: 确认三级模块匹配成功")
                            
                            # 新增：重新核验逻辑
                            if should_enable_verification():
                                new_level = reclassify_module(raw_text.strip(), 3)
                                if new_level != 3:
                                    # 重新分类，跳过三级模块处理
                                    m3 = None
                                    # if debug_enabled:
                                    #     print(f"DEBUG: 三级模块重新分类为{new_level}级，跳过三级处理")
                            
                            if m3:  # 如果没有重新分类，继续正常处理
                                # 遇到新三级编号时，先输出上一组（如果有描述）
                                if in_lvl3 and desc_lines:
                                    results.append({
                                        "一级模块名称": lvl1_to_fill,
                                        "二级模块名称": lvl2_to_fill,
                                        "三级模块名称": last_lvl3,
                                        "标书描述": '\n\n'.join(self._merge_paragraphs(desc_lines)).strip(),
                                        "合同描述": "",
                                        "来源文件": os.path.basename(pdf_path),
                                        # "页码": page_count
                                    })
                                    desc_lines = []
                                    lvl1_filled = True
                                    lvl2_filled = True

                                number = m3.group(1)
                                title = m3.group(2).strip()
                                current_lvl3 = f"{number} {title}".strip()
                                last_lvl3 = current_lvl3
                                in_lvl3 = True
                                in_lvl2 = False
                                lvl1_to_fill = current_lvl1 if not lvl1_filled else ""
                                lvl2_to_fill = current_lvl2 if not lvl2_filled else ""
                                # if debug_enabled:
                                #     print(f"DEBUG: 设置三级模块: '{current_lvl3}'")
                                continue
                    elif lvl3_regex and lvl3_regex.match(text):
                        pass
                        # if debug_enabled:
                        #     print(f"DEBUG: 正则匹配但被过滤 - 原始文本: '{raw_text.strip()}', 处理后: '{text}'")

                    # 再判断二级
                    m2 = lvl2_regex.match(text) if lvl2_regex else None
                    if m2:
                        # if debug_enabled:
                        #     print(f"DEBUG: 二级模块匹配 - 原始文本: '{raw_text.strip()}', 处理后: '{text}'")
                        
                        # 新增：二级模块长度验证
                        def is_valid_lvl2_match(match_obj, original_text):
                            """验证二级模块匹配是否有效"""
                            if not match_obj:
                                return False
                            
                            # 提取匹配的编号部分
                            number_part = match_obj.group(1)
                            title_part = match_obj.group(2).strip()
                            
                            # 长度比较验证
                            def check_lvl2_length_compatibility():
                                """检查二级模块样例长度与匹配内容的兼容性"""
                                if not lvl2_sample:
                                    return True
                                
                                # 只在二级模块样例是简单数字格式时才启用长度判断
                                def is_simple_number_format(sample):
                                    """判断是否为简单数字格式"""
                                    # 匹配 "1."、"1）"、"1)"、"1、" 等简单格式
                                    simple_patterns = [
                                        r'^\d+\.$',      # 1.
                                        r'^\d+[）)]$',   # 1） 或 1)
                                        r'^\d+、$',      # 1、
                                        r'^\d+】$',      # 1】
                                        r'^\d+]$',       # 1]
                                    ]
                                    return any(re.match(pattern, sample) for pattern in simple_patterns)
                                
                                # 如果不是简单数字格式，跳过长度比较
                                if not is_simple_number_format(lvl2_sample):
                                    # if debug_enabled:
                                    #     print(f"DEBUG: 二级样例不是简单数字格式，跳过长度比较: '{lvl2_sample}'")
                                    return True
                                
                                # 提取样例中的数字长度
                                sample_digits = re.sub(r'[^\d]', '', lvl2_sample)
                                sample_digit_length = len(sample_digits)
                                
                                # 提取匹配内容中的数字长度
                                match_digits = re.sub(r'[^\d]', '', number_part)
                                match_digit_length = len(match_digits)
                                
                                # if debug_enabled:
                                #     print(f"DEBUG: 二级长度比较 - 样例数字长度: {sample_digit_length}, 匹配数字长度: {match_digit_length}")
                                
                                # 如果匹配的数字长度明显大于样例，可能是误匹配
                                if match_digit_length > sample_digit_length + 2:  # 允许2位数字的误差
                                    # if debug_enabled:
                                    #     print(f"DEBUG: 二级长度不匹配，可能是误匹配 - 样例: '{lvl2_sample}', 匹配: '{number_part}'")
                                    return False
                                
                                return True
                            
                            # 检查标题部分是否有内容
                            if not title_part:
                                # if debug_enabled:
                                #     print(f"DEBUG: 二级模块标题为空: '{original_text}'")
                                return False
                            
                            # 执行长度比较验证
                            if not check_lvl2_length_compatibility():
                                return False
                            
                            return True
                        
                        # 验证二级模块匹配有效性
                        if not is_valid_lvl2_match(m2, raw_text.strip()):
                            # if debug_enabled:
                            #     print(f"DEBUG: 二级模块验证失败，设为None")
                            m2 = None
                        
                        if lvl2_regex_info and lvl2_regex_info['expected_digit_length']:
                            text_digits = re.sub(r'[^\d]', '', m2.group(1))
                            if len(text_digits) != lvl2_regex_info['expected_digit_length']:
                                # if debug_enabled:
                                #     print(f"DEBUG: 二级模块数字长度不匹配: 期望{lvl2_regex_info['expected_digit_length']}, 实际{len(text_digits)}")
                                m2 = None
                        if m2:
                            # if debug_enabled:
                            #     print(f"DEBUG: 确认二级模块匹配成功")
                            
                            # 新增：重新核验逻辑（没有三级样例时）
                            if not has_lvl3_sample and should_enable_verification():
                                new_level = reclassify_module(raw_text.strip(), 2)
                                if new_level != 2:
                                    # 重新分类，跳过二级模块处理
                                    m2 = None
                                    # if debug_enabled:
                                    #     print(f"DEBUG: 二级模块重新分类为{new_level}级，跳过二级处理")
                            
                            if m2:  # 如果没有重新分类，继续正常处理
                                # 根据是否有三级模块使用不同逻辑
                                if has_lvl3_sample:
                                    # 有三级模块样例：使用老代码逻辑
                                    if in_lvl3 and desc_lines:
                                        results.append({
                                            "一级模块名称": lvl1_to_fill,
                                            "二级模块名称": lvl2_to_fill,
                                            "三级模块名称": last_lvl3,
                                            "标书描述": '\n\n'.join(self._merge_paragraphs(desc_lines)).strip()
                                        })
                                        desc_lines = []
                                        lvl1_filled = True
                                        lvl2_filled = True
                                        in_lvl3 = False
                                    number = m2.group(1)
                                    title = m2.group(2).strip()
                                    current_lvl2 = f"{number} {title}".strip()
                                    current_lvl3 = None
                                    last_lvl2 = current_lvl2
                                    lvl2_filled = False
                                    in_lvl2 = True
                                    in_lvl3 = False
                                    
                                    # 更新填充值
                                    lvl1_to_fill = current_lvl1 if not lvl1_filled else ""
                                    lvl2_to_fill = current_lvl2
                                    # if debug_enabled:
                                    #     print(f"DEBUG: 设置二级模块: '{current_lvl2}'")
                                    continue

                                else:
                                    # 没有三级模块样例：使用新代码逻辑
                                    # 如果之前有描述内容，先输出上一组
                                    if desc_lines and in_lvl2 and last_lvl2:
                                        results.append({
                                            "一级模块名称": lvl1_to_fill,
                                            "二级模块名称": last_lvl2,  # 使用上一个二级模块名称
                                            "三级模块名称": "",
                                            "标书描述": '\n\n'.join(self._merge_paragraphs(desc_lines)).strip()
                                        })
                                        desc_lines = []
                                        lvl1_filled = True
                                        lvl2_filled = True
                                    
                                    # 更新当前二级模块
                                    number = m2.group(1)
                                    title = m2.group(2).strip()
                                    current_lvl2 = f"{number} {title}".strip()
                                    current_lvl3 = None
                                    last_lvl2 = current_lvl2
                                    lvl2_filled = False
                                    in_lvl2 = True
                                    in_lvl3 = False
                                    
                                    # 更新填充值
                                    lvl1_to_fill = current_lvl1 if not lvl1_filled else ""
                                    lvl2_to_fill = current_lvl2

                                    if in_lvl3 and desc_lines:
                                        results.append({
                                            "一级模块名称": lvl1_to_fill,
                                            "二级模块名称": lvl2_to_fill,
                                            "三级模块名称": last_lvl3,
                                            "标书描述": '\n\n'.join(self._merge_paragraphs(desc_lines)).strip()
                                        })
                                        desc_lines = []
                                        lvl1_filled = True
                                        lvl2_filled = True
                                        in_lvl3 = False
                                    # if debug_enabled:
                                    #     print(f"DEBUG: 设置二级模块: '{current_lvl2}'")
                                    continue

                    # 最后判断一级
                    m1 = lvl1_regex.match(text) if lvl1_regex else None
                    if m1:
                        # if debug_enabled:
                        #     print(f"DEBUG: 一级模块匹配 - 原始文本: '{raw_text.strip()}', 处理后: '{text}'")
                        
                        # 新增：一级模块长度验证
                        def is_valid_lvl1_match(match_obj, original_text):
                            """验证一级模块匹配是否有效"""
                            if not match_obj:
                                return False
                            
                            # 提取匹配的编号部分
                            number_part = match_obj.group(1)
                            title_part = match_obj.group(2).strip()
                            
                            # 长度比较验证
                            def check_lvl1_length_compatibility():
                                """检查一级模块样例长度与匹配内容的兼容性"""
                                if not lvl1_sample:
                                    return True
                                
                                # 只在一级模块样例是简单数字格式时才启用长度判断
                                def is_simple_number_format(sample):
                                    """判断是否为简单数字格式"""
                                    # 匹配 "1."、"1）"、"1)"、"1、" 等简单格式
                                    simple_patterns = [
                                        r'^\d+\.$',      # 1.
                                        r'^\d+[）)]$',   # 1） 或 1)
                                        r'^\d+、$',      # 1、
                                        r'^\d+】$',      # 1】
                                        r'^\d+]$',       # 1]
                                    ]
                                    return any(re.match(pattern, sample) for pattern in simple_patterns)
                                
                                # 如果不是简单数字格式，跳过长度比较
                                if not is_simple_number_format(lvl1_sample):
                                    # if debug_enabled:
                                    #     print(f"DEBUG: 一级样例不是简单数字格式，跳过长度比较: '{lvl1_sample}'")
                                    return True
                                
                                # 提取样例中的数字长度
                                sample_digits = re.sub(r'[^\d]', '', lvl1_sample)
                                sample_digit_length = len(sample_digits)
                                
                                # 提取匹配内容中的数字长度
                                match_digits = re.sub(r'[^\d]', '', number_part)
                                match_digit_length = len(match_digits)
                                
                                # if debug_enabled:
                                #     print(f"DEBUG: 一级长度比较 - 样例数字长度: {sample_digit_length}, 匹配数字长度: {match_digit_length}")
                                
                                # 如果匹配的数字长度明显大于样例，可能是误匹配
                                if match_digit_length > sample_digit_length + 2:  # 允许2位数字的误差
                                    # if debug_enabled:
                                    #     print(f"DEBUG: 一级长度不匹配，可能是误匹配 - 样例: '{lvl1_sample}', 匹配: '{number_part}'")
                                    return False
                                
                                return True
                            
                            # 检查标题部分是否有内容
                            if not title_part:
                                # if debug_enabled:
                                #     print(f"DEBUG: 一级模块标题为空: '{original_text}'")
                                return False
                            
                            # 执行长度比较验证
                            if not check_lvl1_length_compatibility():
                                return False
                            
                            return True
                        
                        # 验证一级模块匹配有效性
                        if not is_valid_lvl1_match(m1, raw_text.strip()):
                            # if debug_enabled:
                            #     print(f"DEBUG: 一级模块验证失败，设为None")
                            m1 = None
                        
                        if lvl1_regex_info and lvl1_regex_info['expected_digit_length']:
                            text_digits = re.sub(r'[^\d]', '', m1.group(1))
                            if len(text_digits) != lvl1_regex_info['expected_digit_length']:
                                # if debug_enabled:
                                #     print(f"DEBUG: 一级模块数字长度不匹配: 期望{lvl1_regex_info['expected_digit_length']}, 实际{len(text_digits)}")
                                m1 = None
                        if m1:
                            # if debug_enabled:
                            #     print(f"DEBUG: 确认一级模块匹配成功")
                            # 根据是否有三级模块使用不同逻辑
                            if has_lvl3_sample:
                                # 有三级模块样例：使用老代码逻辑
                                if in_lvl3 and desc_lines:
                                    results.append({
                                        "一级模块名称": lvl1_to_fill,
                                        "二级模块名称": lvl2_to_fill,
                                        "三级模块名称": last_lvl3,
                                        "标书描述": '\n\n'.join(self._merge_paragraphs(desc_lines)).strip(),
                                        "合同描述": "",
                                        "来源文件": os.path.basename(pdf_path),
                                        # "页码": page_count
                                    })
                                    desc_lines = []
                                    lvl1_filled = True
                                    lvl2_filled = True
                                    in_lvl3 = False
                            else:
                                # 没有三级模块样例：使用新代码逻辑
                                if desc_lines and not in_lvl3:
                                    results.append({
                                        "一级模块名称": lvl1_to_fill,
                                        "二级模块名称": lvl2_to_fill,
                                        "三级模块名称": "",
                                        "标书描述": '\n\n'.join(self._merge_paragraphs(desc_lines)).strip(),
                                        "合同描述": "",
                                        "来源文件": os.path.basename(pdf_path),
                                        # "页码": page_count
                                    })
                                    desc_lines = []
                                    lvl1_filled = True
                                    lvl2_filled = True

                                if in_lvl3 and desc_lines:
                                    results.append({
                                        "一级模块名称": lvl1_to_fill,
                                        "二级模块名称": lvl2_to_fill,
                                        "三级模块名称": last_lvl3,
                                        "标书描述": '\n\n'.join(self._merge_paragraphs(desc_lines)).strip(),
                                        "合同描述": "",
                                        "来源文件": os.path.basename(pdf_path),
                                        # "页码": page_count
                                    })
                                    desc_lines = []
                                    lvl1_filled = True
                                    lvl2_filled = True
                                    in_lvl3 = False

                            number = m1.group(1)
                            title = m1.group(2).strip()
                            current_lvl1 = f"{number} {title}".strip()
                            current_lvl2 = current_lvl3 = None
                            last_lvl1 = current_lvl1
                            lvl1_filled = False
                            lvl2_filled = False
                            in_lvl2 = False
                            in_lvl3 = False
                            lvl1_to_fill = current_lvl1
                            lvl2_to_fill = current_lvl2 if current_lvl2 else ""
                            # if debug_enabled:
                            #     print(f"DEBUG: 设置一级模块: '{current_lvl1}'")
                            continue

                    # 修改后的收集逻辑 - 关键修复
                    if extracting:
                        # 新增：更严格的描述收集验证
                        def should_collect_description():
                            """判断是否应该收集描述内容"""
                            # 如果有三级模块样例
                            if has_lvl3_sample:
                                # 只有在三级模块下才收集
                                return in_lvl3
                            else:
                                # 没有三级模块样例时
                                if in_lvl2:
                                    return True
                                # elif current_lvl1 and not current_lvl2:
                                #    return True
                                return False
                        
                        # 只有在应该收集的情况下才添加描述
                        if should_collect_description():
                            # 额外验证：确保不是模块标题行
                            if not (lvl1_regex and lvl1_regex.match(text)) and \
                               not (lvl2_regex and lvl2_regex.match(text)) and \
                               not (lvl3_regex and lvl3_regex.match(text)):
                                desc_lines.append(raw_text.strip())
                        else:
                            pass

        # 补充最后一组
        if in_lvl3 and desc_lines:
            results.append({
                "一级模块名称": lvl1_to_fill,
                "二级模块名称": lvl2_to_fill,
                "三级模块名称": last_lvl3,
                "标书描述": '\n\n'.join(self._merge_paragraphs(desc_lines)).strip(),
                "合同描述": "",
                "来源文件": os.path.basename(pdf_path),
                # "页码": page_count
            })
        elif desc_lines:  # 如果没有三级模块但有描述内容
            # 确保有正确的二级模块名称
            final_lvl2 = last_lvl2 if last_lvl2 else current_lvl2 if current_lvl2 else lvl2_to_fill
            results.append({
                "一级模块名称": current_lvl1 if current_lvl1 else lvl1_to_fill,
                "二级模块名称": final_lvl2,
                "三级模块名称": "",
                "标书描述": '\n\n'.join(self._merge_paragraphs(desc_lines)).strip(),
                "合同描述": "",
                "来源文件": os.path.basename(pdf_path),
                # "页码": page_count
            })

        # 清理数据 - 暂时注释掉
        # results = self._clean_extracted_data(results)
        results = [r for r in results if any([r["一级模块名称"], r["二级模块名称"], r["三级模块名称"], r["标书描述"]])]
        print(f"最终提取到 {len(results)} 条内容")
        return results

    # ----------- Word 合同（修复版，处理表格分割问题） -----------
    def extract_tables_from_word_contract(self, docx_path: str) -> List[Dict]:
        data = []
        found_quotation_section = False
        current_headers = None
        doc = Document(docx_path)
        
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
            
            # 检查是否是目标表格（使用自定义表头或默认表头）
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
                        # 如果序号是数字，认为有效
                        try:
                            int(row_data[header])
                            has_data = True
                            break
                        except ValueError:
                            # 如果不是数字，检查其他字段
                            pass
                
                # 如果序号不是数字，检查其他关键字段
                if not has_data:
                    # 使用自定义表头或默认表头进行检查
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
                    mapped = self._map_word_row_custom(row_data, docx_path)
                    
                    # 新增：检查重复并处理
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
                        if mapped['一级模块名称'] == last_lvl1 and mapped['一级模块名称']:
                            mapped['一级模块名称'] = ''
                        # 检查二级模块名称是否重复
                        if mapped['二级模块名称'] == last_lvl2 and mapped['二级模块名称']:
                            mapped['二级模块名称'] = ''
                    
                    data.append(mapped)
                else:
                    pass
        
        return data

    def extract_tables_from_word_bid(self, docx_path: str) -> List[Dict]:
        """提取Word标书文件中的表格"""
        # 标书文件使用默认的映射逻辑
        data = []
        found_quotation_section = False
        current_headers = None
        doc = Document(docx_path)
        
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
            if self._is_target_table(headers):
                current_headers = headers
                found_quotation_section = True
                start_row = 1
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
                        # 如果序号是数字，认为有效
                        try:
                            int(row_data[header])
                            has_data = True
                            break
                        except ValueError:
                            # 如果不是数字，检查其他字段
                            pass
                
                # 如果序号不是数字，检查其他关键字段
                if not has_data:
                    key_fields = ['功能描述', '三级模块', '功能模块', '功能子项']
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
                    mapped = self._map_word_row(row_data, docx_path)
                    data.append(mapped)
                else:
                    pass
        
        return data

    def extract_tables_from_pdf_contract(self, pdf_path: str) -> List[Dict]:
        """提取PDF合同文件中的表格"""
        # PDF合同文件暂时使用标书的提取逻辑
        # 可以根据合同文件的特点进行后续优化
        return self.extract_tables_from_pdf_bid(pdf_path)

    def _is_target_table_custom(self, headers: List[str]) -> bool:
        """使用自定义表头检查是否是目标表格"""
        if self.custom_headers:
            # 使用自定义表头
            target_headers = list(self.custom_headers.keys())
        else:
            # 使用默认表头
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

    def _get_key_fields_for_check(self) -> List[str]:
        """获取用于检查的关键字段列表"""
        if self.custom_headers:
            # 使用自定义表头
            return list(self.custom_headers.keys())
        else:
            # 使用默认表头
            return ['功能描述', '三级模块', '功能模块', '功能子项']

    def _map_word_row_custom(self, row_data: Dict, source_file: str) -> Dict:
        """使用自定义表头映射Word行数据"""
        if self.custom_headers:
            # 使用自定义表头映射
            mapped_data = {}
            for header, value in row_data.items():
                for custom_header, standard_field in self.custom_headers.items():
                    if custom_header in header:
                        mapped_data[standard_field] = value
                        break
            
            # 新增：检查是否有实际内容，避免重复
            lvl1_value = mapped_data.get('一级模块名称', '')
            lvl2_value = mapped_data.get('二级模块名称', '')
            lvl3_value = mapped_data.get('三级模块名称', '')
            desc_value = mapped_data.get('功能描述', mapped_data.get('合同描述', ''))
            
            # 新增：数据清洗函数
            def clean_module_name(text):
                """清洗模块名称，去除多余空格和换行符"""
                if not text:
                    return ''
                # 去除换行符、制表符等
                cleaned = text.replace('\n', '').replace('\r', '').replace('\t', '')
                # 去除多余空格（包括全角空格）
                cleaned = re.sub(r'\s+', ' ', cleaned)
                # 去除首尾空格
                cleaned = cleaned.strip()
                return cleaned
            
            # 清洗各级模块名称
            lvl1_cleaned = clean_module_name(lvl1_value)
            lvl2_cleaned = clean_module_name(lvl2_value)
            lvl3_cleaned = clean_module_name(lvl3_value)
            desc_cleaned = clean_module_name(desc_value)
            
            # 只有当有实际内容时才填充，否则留空
            mapped = {
                '一级模块名称': lvl1_cleaned if lvl1_cleaned else '',
                '二级模块名称': lvl2_cleaned if lvl2_cleaned else '',
                '三级模块名称': lvl3_cleaned if lvl3_cleaned else '',
                '功能描述': desc_cleaned if desc_cleaned else '',  # 支持功能描述和合同描述
                '标书描述': '',
                '合同描述': clean_module_name(mapped_data.get('合同描述', '')),
                '来源文件': os.path.basename(source_file),
                # '页码': 'Word文档'
            }
        else:
            # 使用默认映射
            mapped = self._map_word_row(row_data, source_file)
        
        return mapped

    # ----------- 公共表格处理 -----------
    def _is_target_table(self, headers: List[str]) -> bool:
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

    def _process_table(self, table: List[List], headers: List[str], source_file: str, page_num: int) -> List[Dict]:
        processed_data = []
        column_indices = {}
        for target_col, output_col in self.target_columns.items():
            for idx, header in enumerate(headers):
                if target_col in header:
                    column_indices[output_col] = idx
                    break
        start_row = 1 if any(str(cell).strip() for cell in table[0]) else 0
        for row in table[start_row:]:
            if not row or all(not cell for cell in row):
                continue
            row_data = {}
            for output_col, col_idx in column_indices.items():
                if col_idx < len(row):
                    cell_value = str(row[col_idx]).strip() if row[col_idx] else ''
                    row_data[output_col] = cell_value
                else:
                    row_data[output_col] = ''
            if "标书" in source_file:
                row_data['标书描述'] = row_data.get('功能描述', '')
                row_data['合同描述'] = ''
            elif "合同" in source_file:
                row_data['合同描述'] = row_data.get('功能描述', '')
                row_data['标书描述'] = ''
            else:
                row_data['标书描述'] = ''
                row_data['合同描述'] = ''
            row_data['来源文件'] = os.path.basename(source_file)
            row_data['页码'] = page_num + 1
            processed_data.append(row_data)
        return processed_data

    def _map_word_row(self, row_data: Dict, source_file: str) -> Dict:
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
            '功能描述': mapped_data.get('功能描述', ''),
            '标书描述': '',
            '合同描述': '',
            '来源文件': os.path.basename(source_file),
            # '页码': 'Word文档'
        }
        
        if "标书" in source_file:
            mapped['标书描述'] = mapped['功能描述']
        elif "合同" in source_file:
            mapped['合同描述'] = mapped['功能描述']
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
                   cleaned_item.get("三级模块名称"), cleaned_item.get("标书描述")]):
                cleaned_data.append(cleaned_item)
        
        return cleaned_data

    # ----------- Excel输出（修复行高计算） -----------
    def create_excel_output(self, data: List[Dict], output_path: str, append_mode=False):
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
                r'[I]+\.',  # I.II.III.
                
                # 特殊符号
                r'[•●○◇□]',  # 项目符号
                r'第一、',  # 第一、第二、
                r'第一条',  # 第一条第二条
            ]
            
            # 合并所有模式
            combined_pattern = '|'.join(number_patterns)
            
            # 查找所有编号位置
            matches = list(re.finditer(combined_pattern, description))
            
            if not matches:
                # 如果没有找到编号，按句号分割
                sentences = re.split(r'[。！？]', description)
                segments = []
                current_segment = ""
                
                for sentence in sentences:
                    if sentence.strip():
                        if len(current_segment + sentence) <= max_length:
                            current_segment += sentence + "。"
                        else:
                            if current_segment:
                                segments.append(current_segment.strip())
                            current_segment = sentence + "。"
                
                if current_segment:
                    segments.append(current_segment.strip())
                
                return segments if segments else [description]
            
            # 根据编号位置分割
            segments = []
            start_pos = 0
            
            for i, match in enumerate(matches):
                match_start = match.start()
                
                # 如果当前段已经超过最大长度，强制分割
                if match_start - start_pos > max_length:
                    # 在最大长度处寻找合适的分割点
                    split_pos = start_pos + max_length
                    # 向前寻找句号或逗号
                    for j in range(split_pos, start_pos, -1):
                        if description[j] in '。，！？':
                            split_pos = j + 1
                            break
                    
                    segments.append(description[start_pos:split_pos].strip())
                    start_pos = split_pos
                
                # 如果这是第一个编号，保留前面的内容
                if i == 0 and match_start > 0:
                    segments.append(description[start_pos:match_start].strip())
                
                # 确定当前段的结束位置
                if i < len(matches) - 1:
                    end_pos = matches[i + 1].start()
                else:
                    end_pos = len(description)
                
                # 添加当前编号段
                current_segment = description[match_start:end_pos].strip()
                if current_segment:
                    segments.append(current_segment)
                
                start_pos = end_pos
            
            # 处理最后一段
            if start_pos < len(description):
                last_segment = description[start_pos:].strip()
                if last_segment:
                    segments.append(last_segment)
            
            return segments if segments else [description]
        
        # 处理同名文件
        def get_unique_filename(filepath):
            if not os.path.exists(filepath):
                return filepath
            directory = os.path.dirname(filepath)
            filename = os.path.basename(filepath)
            name, ext = os.path.splitext(filename)
            counter = 1
            while True:
                new_filename = f"{name}_{counter}{ext}"
                new_filepath = os.path.join(directory, new_filename)
                if not os.path.exists(new_filepath):
                    return new_filepath
                counter += 1
        
        # 清理数据中的非法字符
        def clean_cell_value(value):
            if value is None:
                return ""
            if isinstance(value, str):
                cleaned = ""
                for char in value:
                    if ord(char) < 32 and char not in '\t\n\r':
                        continue
                    cleaned += char
                return cleaned.strip()
            return str(value)
        
        # 清理所有数据并进行智能分段
        cleaned_data = []
        for row in data:
            # 1. 先清理所有字段
            cleaned_row = {}
            for key, value in row.items():
                cleaned_row[key] = clean_cell_value(value)
            
            # 2. 检查是否有需要分段的描述字段
            has_segmentation = False
            segments_data = {}
            
            for key in ['标书描述', '合同描述', '功能描述']:
                if key in cleaned_row and cleaned_row[key]:
                    segments = split_long_description(cleaned_row[key])
                    if len(segments) > 1:
                        has_segmentation = True
                        segments_data[key] = segments
            
            # 3. 根据是否有分段需求处理
            if has_segmentation:
                # 有分段需求，创建多行数据
                max_segments = max(len(segments) for segments in segments_data.values())
                
                for i in range(max_segments):
                    new_row = cleaned_row.copy()
                    
                    # 处理每个描述字段
                    for key in ['标书描述', '合同描述', '功能描述']:
                        if key in segments_data:
                            if i < len(segments_data[key]):
                                new_row[key] = segments_data[key][i]
                            else:
                                new_row[key] = ''  # 如果这个字段的段数不够，设为空
                    
                    # 新增：分段后只在第一段保留模块名称，后续段清空
                    if i > 0:
                        # 后续段清空模块名称字段
                        new_row['一级模块名称'] = ''
                        new_row['二级模块名称'] = ''
                        new_row['三级模块名称'] = ''
                    
                    cleaned_data.append(new_row)
            else:
                # 没有分段需求，直接添加原行
                cleaned_data.append(cleaned_row)
        
        df = pd.DataFrame(cleaned_data)
        
        # 确保所有必需的列都存在
        required_columns = ['一级模块名称', '二级模块名称', '三级模块名称', '标书描述', '合同描述', '来源文件', '页码']
        for col in required_columns:
            if col not in df.columns:
                df[col] = ''  # 添加缺失的列，填充空字符串
        
        column_order = ['一级模块名称', '二级模块名称', '三级模块名称', '标书描述', '合同描述', '来源文件', '页码']
        df = df[column_order]
        
        # 追加模式处理
        if append_mode and os.path.exists(output_path):
            try:
                # 读取现有Excel文件
                existing_df = pd.read_excel(output_path, sheet_name='分项报价表提取结果')
                # 合并数据
                combined_df = pd.concat([existing_df, df], ignore_index=True)
                df = combined_df
                logger.info(f"已追加 {len(data)} 条新记录到现有文件")
            except Exception as e:
                logger.warning(f"读取现有文件失败，将创建新文件: {e}")
                append_mode = False        
        
        # 获取唯一文件名（仅在非追加模式时）
        if not append_mode:
            unique_output_path = get_unique_filename(output_path)
        else:
            unique_output_path = output_path
        
        with pd.ExcelWriter(unique_output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='分项报价表提取结果', index=False)
            worksheet = writer.sheets['分项报价表提取结果']
            
            # 设置列宽
            column_widths = {
                'A': 15, 'B': 15, 'C': 15, 'D': 45, 'E': 45, 'F': 10, 'G': 10
            }
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
                    
        logger.info(f"Excel文件已保存到: {unique_output_path}")
        return unique_output_path

def main():
    extractor = PDFWordTableExtractor()
    files = [
        "标书.PDF",
        "合同.pdf",
        "标书.docx",
        "合同.docx"
    ]
    all_data = []
    
    # 获取唯一的输出文件名（在开始时就确定）
    base_output_file = "分项报价表提取结果.xlsx"
    
    # 获取唯一的输出文件名
    def get_unique_filename(filepath):
        if not os.path.exists(filepath):
            return filepath
        directory = os.path.dirname(filepath)
        filename = os.path.basename(filepath)
        name, ext = os.path.splitext(filename)
        counter = 1
        while True:
            new_filename = f"{name}_{counter}{ext}"
            new_filepath = os.path.join(directory, new_filename)
            if not os.path.exists(new_filepath):
                return new_filepath
            counter += 1
    
    output_file = get_unique_filename(base_output_file)
    
    # 处理标书文件
    bid_files = [f for f in files if "标书" in f and os.path.exists(f)]
    contract_files = [f for f in files if "合同" in f and os.path.exists(f)]
    
    # 标书提取循环
    for bid_file in bid_files:
        ext = os.path.splitext(bid_file)[1].lower()
        file_type = "pdf" if ext == ".pdf" else "docx" if ext == ".docx" else None
        if not file_type:
            continue
            
        logger.info(f"正在处理标书文件: {bid_file}")
        
        while True:
            # 提取标书数据
            bid_data = extractor.extract_tables(bid_file, file_type)
            
            if bid_data:
                print(f"\n📊 当前提取到 {len(bid_data)} 条标书记录")
                
                all_data.extend(bid_data)  # 添加到总数据
                print(f"📊 有效数据: {len(bid_data)} 条")
                
                # 询问是否继续提取
                while True:
                    choice = input("\n是否继续提取标书？(y/n): ").strip().lower()
                    if choice in ['y', 'n']:
                        break
                    print("请输入 y 或 n")
                
                if choice == 'n':
                    break
                else:
                    print("\n请重新输入编号样例进行下一轮提取...")
                    continue
            else:
                print(f"\n❌ 在文件 {bid_file} 中未找到分项报价表")
                break
    
    # 处理合同文件
    for contract_file in contract_files:
        ext = os.path.splitext(contract_file)[1].lower()
        file_type = "pdf" if ext == ".pdf" else "docx" if ext == ".docx" else None
        if not file_type:
            continue
            
        logger.info(f"正在处理合同文件: {contract_file}")
        
        # 对于Word合同文件，提示用户设置表头
        if file_type == "docx" and "合同" in contract_file:
            print(f"\n📋 即将处理Word合同文件: {contract_file}")
            print("请根据Word文档中的实际表头设置字段映射...")
        
        contract_data = extractor.extract_tables(contract_file, file_type)
        if contract_data:
            all_data.extend(contract_data)
            print(f"\n📊 合同文件提取到 {len(contract_data)} 条记录")
    
    # 保存所有数据到一个Excel文件
    if all_data:
        actual_output_file = extractor.create_excel_output(all_data, output_file, append_mode=False)
        if actual_output_file:
            print(f"\n✅ 提取完成！共提取 {len(all_data)} 条记录")
            print(f"📁 结果文件：{actual_output_file}")
        else:
            print("\n❌ 保存文件失败")
    else:
        print("\n❌ 未找到分项报价表，请检查：")
        print("1. 文件是否包含'分项报价表'字样")
        print("2. 表格是否包含'功能模块'、'功能子项'、'三级模块'等列")
        print("3. 对于Word合同文件，请检查表头设置是否正确")

if __name__ == "__main__":
    main()