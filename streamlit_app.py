# streamlit_app.py
import streamlit as st
import os
import tempfile
import pandas as pd
from pdf_extractorV2_2 import PDFWordTableExtractor

def main():
    st.set_page_config(
        page_title="智能表格提取工具",
        page_icon="��",
        layout="wide",
        initial_sidebar_state="collapsed"
    )
    
    # 自定义CSS样式
    st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .upload-area {
        border: 3px dashed #1f77b4;
        border-radius: 15px;
        padding: 3rem;
        text-align: center;
        background: linear-gradient(135deg, #f0f8ff 0%, #e6f3ff 100%);
        margin: 2rem 0;
    }
    .success-box {
        background: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 10px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .error-box {
        background: #f8d7da;
        border: 1px solid #f5c6cb;
        border-radius: 10px;
        padding: 1rem;
        margin: 1rem 0;
    }
    </style>
    """, unsafe_allow_html=True)
    
    st.markdown('<h1 class="main-header">📋 智能表格提取工具</h1>', unsafe_allow_html=True)
    
    # 功能介绍
    with st.expander("ℹ️ 功能介绍", expanded=False):
        st.markdown("""
        **功能特点：**
        - �� 智能识别PDF和Word文档中的表格
        - 📊 自动提取分项报价表数据
        - 🏗️ 支持多级模块结构识别
        - 📥 一键导出Excel文件
        - �� 零安装，即开即用
        
        **支持格式：**
        - PDF文件 (.pdf)
        - Word文档 (.docx)
        
        **使用说明：**
        1. 上传包含"分项报价表"的PDF或Word文件
        2. 系统自动识别表格结构
        3. 点击"开始提取"进行处理
        4. 下载Excel格式的结果文件
        """)
    
    st.markdown("---")
    
    # 文件上传区域
    st.markdown('<div class="upload-area">', unsafe_allow_html=True)
    st.markdown("### �� 文件上传")
    st.markdown("**支持拖拽上传多个文件**")
    
    uploaded_files = st.file_uploader(
        "选择PDF或Word文件",
        type=['pdf', 'docx'],
        accept_multiple_files=True,
        help="可以同时选择多个文件，支持拖拽上传"
    )
    st.markdown('</div>', unsafe_allow_html=True)
    
    if uploaded_files:
        # 显示上传的文件
        st.markdown("### 📋 已上传文件")
        for i, file in enumerate(uploaded_files):
            file_size = len(file.getvalue()) / 1024  # KB
            st.write(f"**{i+1}.** {file.name} ({file_size:.1f} KB)")
        
        # 处理按钮
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("🚀 开始提取", type="primary", use_container_width=True):
                process_files(uploaded_files)
    
    else:
        st.info("👆 请上传PDF或Word文件开始提取")
        
        # 添加示例说明
        st.markdown("### 💡 使用提示")
        st.markdown("""
        - 确保文件包含"分项报价表"相关内容
        - 支持标书和合同文件的自动识别
        - 系统会自动识别表格结构和模块层级
        - 提取结果将保存为Excel格式
        """)

def process_files(uploaded_files):
    """处理上传的文件"""
    extractor = PDFWordTableExtractor()
    all_data = []
    
    # 创建进度条
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for i, uploaded_file in enumerate(uploaded_files):
        status_text.text(f"正在处理: {uploaded_file.name}")
        
        # 修改这里：使用正确的文件扩展名
        file_extension = os.path.splitext(uploaded_file.name)[1]  # 获取 .pdf 或 .docx
        with tempfile.NamedTemporaryFile(delete=False, suffix=file_extension) as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        
        try:
            # 确定文件类型
            file_type = "pdf" if uploaded_file.type == "application/pdf" else "docx"
            
            # 提取数据
            data = extractor.extract_tables(tmp_file_path, file_type)
            
            if data:
                all_data.extend(data)
                st.markdown(f'<div class="success-box">✅ {uploaded_file.name}: 提取到 {len(data)} 条记录</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="error-box">⚠️ {uploaded_file.name}: 未提取到数据</div>', unsafe_allow_html=True)
        
        except Exception as e:
            st.markdown(f'<div class="error-box">❌ {uploaded_file.name}: 处理失败 - {str(e)}</div>', unsafe_allow_html=True)
        
        finally:
            # 清理临时文件
            if os.path.exists(tmp_file_path):
                os.unlink(tmp_file_path)
        
        # 更新进度
        progress_bar.progress((i + 1) / len(uploaded_files))
    
    # 显示结果
    if all_data:
        st.success(f"🎉 提取完成！共 {len(all_data)} 条记录")
        
        # 创建DataFrame
        df = pd.DataFrame(all_data)
        
        # 显示数据预览
        st.subheader("📊 数据预览")
        st.dataframe(df.head(10), use_container_width=True)
        
        # 下载按钮
        st.subheader("📥 下载结果")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # CSV下载
            csv_data = df.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📄 下载CSV文件",
                data=csv_data,
                file_name="提取结果.csv",
                mime="text/csv",
                use_container_width=True
            )
        
        with col2:
            # Excel下载
            excel_buffer = pd.ExcelWriter('temp.xlsx', engine='openpyxl')
            df.to_excel(excel_buffer, index=False, sheet_name='提取结果')
            excel_buffer.close()
            
            with open('temp.xlsx', 'rb') as f:
                excel_data = f.read()
            
            st.download_button(
                label="📊 下载Excel文件",
                data=excel_data,
                file_name="提取结果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            # 清理临时文件
            if os.path.exists('temp.xlsx'):
                os.unlink('temp.xlsx')
    else:
        st.error("❌ 未提取到任何数据")

if __name__ == "__main__":
    main()
