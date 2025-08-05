# -*- coding: utf-8 -*-
"""
PDF和Word表格提取工具 - Web版本
简化版本，专注于基本功能
"""

import os
import re
import logging
import pdfplumber
import pandas as pd
from docx import Document
from typing import List, Dict, Optional
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

# 设置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class PDFWordTableExtractor:
    def __init__(self):
        # 默认的目标列映射
        self.target_columns = {
            '功能模块': '一级模块名称',
            '功能子项': '二级模块名称', 
            '三级模块': '三级模块名称',
            '功能描述': '功能描述'
        }
        
        # 自定义表头映射（简化版）
        self.custom_headers = {}
        
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
    
    def extract_tables_from_pdf_bid(self, pdf_path: str) -> List[Dict]:
        """提取PDF标书文件中的表格（简化版）"""
        data = []
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    tables = page.extract_tables()
                    for table in tables:
                        if table and len(table) > 1:
                            # 处理表格数据
                            processed_data = self._process_pdf_table(table, pdf_path, page_num)
                            data.extend(processed_data)
        except Exception as e:
            logger.error(f"PDF处理错误: {e}")
        return data
    
    def _process_pdf_table(self, table: List[List], source_file: str, page_num: int) -> List[Dict]:
        """处理PDF表格数据"""
        processed_data = []
        if not table or len(table) < 2:
            return processed_data
            
        # 获取表头
        headers = [str(cell).strip() if cell else '' for cell in table[0]]
        
        # 处理数据行
        for row in table[1:]:
            if not row or all(not cell for cell in row):
                continue
                
            row_data = {}
            for idx, cell in enumerate(row):
                if idx < len(headers):
                    row_data[headers[idx]] = str(cell).strip() if cell else ''
            
            # 映射到标准格式
            mapped = self._map_pdf_row(row_data, source_file)
            if mapped:
                processed_data.append(mapped)
        
        return processed_data
    
    def _map_pdf_row(self, row_data: Dict, source_file: str) -> Optional[Dict]:
        """映射PDF行数据到标准格式"""
        mapped = {
            '一级模块名称': '',
            '二级模块名称': '',
            '三级模块名称': '',
            '标书描述': '',
            '合同描述': '',
            '来源文件': os.path.basename(source_file),
        }
        
        # 简单的字段映射
        for header, value in row_data.items():
            if '功能模块' in header or '模块' in header:
                mapped['一级模块名称'] = value
            elif '功能子项' in header or '子项' in header:
                mapped['二级模块名称'] = value
            elif '三级模块' in header:
                mapped['三级模块名称'] = value
            elif '功能描述' in header or '描述' in header:
                if "标书" in source_file:
                    mapped['标书描述'] = value
                elif "合同" in source_file:
                    mapped['合同描述'] = value
        
        # 只返回有数据的行
        if any([mapped['一级模块名称'], mapped['二级模块名称'], 
               mapped['三级模块名称'], mapped['标书描述'], mapped['合同描述']]):
            return mapped
        return None
    
    def extract_tables_from_word_contract(self, docx_path: str) -> List[Dict]:
        """提取Word合同文件中的表格（简化版）"""
        data = []
        try:
            doc = Document(docx_path)
            
            # 检查是否包含分项报价表
            has_quotation_table = False
            for para in doc.paragraphs:
                if "分项报价表" in para.text:
                    has_quotation_table = True
                    break
            
            if not has_quotation_table:
                return data
            
            # 处理所有表格
            for table in doc.tables:
                if not table.rows:
                    continue
                
                # 获取表头
                headers = [cell.text.strip() for cell in table.rows[0].cells]
                
                # 处理数据行
                for row in table.rows[1:]:
                    row_data = {}
                    cells = row.cells
                    
                    # 构建行数据
                    for idx, header in enumerate(headers):
                        if idx < len(cells):
                            row_data[header] = cells[idx].text.strip()
                        else:
                            row_data[header] = ''
                    
                    # 映射数据
                    mapped = self._map_word_row_simple(row_data, docx_path)
                    if mapped:
                        data.append(mapped)
                        
        except Exception as e:
            logger.error(f"Word合同处理错误: {e}")
        
        return data
    
    def _map_word_row_simple(self, row_data: Dict, source_file: str) -> Optional[Dict]:
        """简化的Word行数据映射"""
        mapped = {
            '一级模块名称': '',
            '二级模块名称': '',
            '三级模块名称': '',
            '标书描述': '',
            '合同描述': '',
            '来源文件': os.path.basename(source_file),
        }
        
        # 查找各个字段
        for header, value in row_data.items():
            if not value.strip():
                continue
                
            # 一级模块名称
            if '功能模块' in header or '模块名称' in header or '一级' in header:
                mapped['一级模块名称'] = value
            # 二级模块名称
            elif '功能子项' in header or '子项' in header or '二级' in header:
                mapped['二级模块名称'] = value
            # 三级模块名称
            elif '三级模块' in header or '三级' in header:
                mapped['三级模块名称'] = value
            # 描述字段
            elif '功能描述' in header or '描述' in header or '备注' in header or '内容' in header:
                mapped['合同描述'] = value
        
        # 只返回有数据的行
        if any([mapped['一级模块名称'], mapped['二级模块名称'], 
               mapped['三级模块名称'], mapped['合同描述']]):
            return mapped
        return None
    
    def extract_tables_from_word_bid(self, docx_path: str) -> List[Dict]:
        """提取Word标书文件中的表格"""
        # 标书文件使用相同的逻辑，但描述字段映射到标书描述
        data = self.extract_tables_from_word_contract(docx_path)
        
        # 调整描述字段映射
        for item in data:
            if item['合同描述']:
                item['标书描述'] = item['合同描述']
                item['合同描述'] = ''
        
        return data
    
    def extract_tables_from_pdf_bid_with_samples(self, pdf_path: str, lvl1_sample: str, 
                                               lvl2_sample: str = "", lvl3_sample: str = "", 
                                               end_sample: str = "") -> List[Dict]:
        """使用编号样例提取PDF标书（简化版）"""
        # 这里可以添加更复杂的PDF提取逻辑
        # 暂时使用基本的提取方法
        return self.extract_tables_from_pdf_bid(pdf_path)
    
    def create_excel_output(self, data: List[Dict], output_path: str, append_mode=False):
        """创建Excel输出文件"""
        if not data:
            logger.warning("没有数据需要输出")
            return None
        
        # 创建DataFrame
        df = pd.DataFrame(data)
        
        # 确保列顺序正确
        column_order = ['一级模块名称', '二级模块名称', '三级模块名称', '标书描述', '合同描述', '来源文件']
        for col in column_order:
            if col not in df.columns:
                df[col] = ''
        df = df[column_order]
        
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
            
            # 设置数据行样式
            data_alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
            for row in range(2, len(df) + 2):
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
    