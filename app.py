import streamlit as st
import pandas as pd
import pdfplumber
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
import os
import tempfile
import traceback
from io import BytesIO
from streamlit_extras.app_logo import add_logo
from streamlit_extras.metric_cards import style_metric_cards

st.set_page_config(
    page_title="PDF转Excel神器",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)
add_logo("https://img.icons8.com/ios-filled/100/000000/pdf.png", height=60)

st.markdown(
    """
    <style>
    .main {
        background: linear-gradient(135deg, #f8fafc 0%, #e0e7ef 100%);
    }
    .stButton>button {
        background: linear-gradient(90deg, #4f8cff 0%, #38b6ff 100%);
        color: white;
        font-weight: bold;
        border-radius: 8px;
        padding: 0.5em 2em;
        margin-top: 1em;
    }
    .stDownloadButton>button {
        background: #38b6ff;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        padding: 0.5em 2em;
    }
    .stProgress>div>div>div>div {
        background-image: linear-gradient(90deg, #4f8cff, #38b6ff);
    }
    .stDataFrame {
        background: #fff;
        border-radius: 8px;
        box-shadow: 0 2px 8px #e0e7ef;
    }
    .stSidebar {
        background: #f0f4fa;
    }
    </style>
    """,
    unsafe_allow_html=True
)

def extract_pdf_to_dataframe(pdf_file):
    """
    Extract data from PDF file and convert to DataFrame
    """
    try:
        tables = []
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                # Try to extract table from the page
                page_tables = page.extract_tables()
                if page_tables:
                    for table in page_tables:
                        if table:  # Check if table is not empty
                            tables.extend(table)
                else:
                    # If no tables found, try to extract text and split by lines
                    text = page.extract_text()
                    if text:
                        lines = text.split('\n')
                        for line in lines:
                            if line.strip():
                                # Split by common delimiters
                                row = line.split('\t') if '\t' in line else line.split()
                                if len(row) > 1:
                                    tables.append(row)
        if not tables:
            st.error("无法从PDF中提取任何数据。请确保PDF包含表格或结构化数据。")
            return None
        # Convert to DataFrame
        df = pd.DataFrame(tables)
        # Clean up the DataFrame - remove empty rows and columns
        df = df.dropna(how='all').dropna(axis=1, how='all')
        # Reset index
        df = df.reset_index(drop=True)
        return df
    except Exception as e:
        st.error(f"PDF提取失败: {str(e)}")
        return None

def process_excel_data(df):
    """
    Process the DataFrame according to the specified requirements
    """
    try:
        # Make a copy to avoid modifying original data
        processed_df = df.copy()
        
        # Ensure we have enough columns (at least I column which is index 8)
        if len(processed_df.columns) < 9:
            # Add empty columns if needed
            for i in range(len(processed_df.columns), 9):
                processed_df[i] = ''
        
        # Delete column A (交易单号), H, I (indices 0, 7, 8)
        # After deletion, 交易时间 will become the new A column
        columns_to_drop = []
        if 0 < len(processed_df.columns):  # Column A (index 0) - 交易单号
            columns_to_drop.append(0)
        if 7 < len(processed_df.columns):  # Column H (index 7)
            columns_to_drop.append(7)
        if 8 < len(processed_df.columns):  # Column I (index 8)
            columns_to_drop.append(8)
        
        # Drop columns in reverse order to maintain indices
        for col_idx in sorted(columns_to_drop, reverse=True):
            processed_df = processed_df.drop(processed_df.columns[col_idx], axis=1)
        
        # Reset column names
        processed_df.columns = range(len(processed_df.columns))
        
        # Process columns: G (交易对方) and B (交易时间)
        # After deleting A (交易单号), B (交易时间) becomes index 0, G (交易对方) becomes index 5
        b_column_index = 0  # 交易时间 is now the first column (new A column)
        g_column_index = 5 if len(processed_df.columns) > 5 else len(processed_df.columns) - 1  # 交易对方
        
        if g_column_index < len(processed_df.columns) and b_column_index < len(processed_df.columns):
            # Sort by column G (交易对方) first, then by column B (交易时间) from newest to oldest
            # For time sorting, we'll try to convert to datetime if possible
            try:
                # Convert to datetime and sort first
                temp_datetime = pd.to_datetime(processed_df[b_column_index], errors='coerce')
                
                # Sort by G column (交易对方) first, then by datetime
                processed_df = processed_df.sort_values(
                    by=[g_column_index, b_column_index], 
                    ascending=[True, False],  # G列升序(相同交易对方聚集), B列降序(时间从近到远)
                    na_position='last'
                )
                
                # Format datetime to remove seconds - apply to all cells in the column
                def format_time(time_str):
                    try:
                        if pd.isna(time_str) or time_str == '':
                            return time_str
                        dt = pd.to_datetime(time_str, errors='coerce')
                        if pd.isna(dt):
                            return time_str
                        return dt.strftime('%Y-%m-%d %H:%M')
                    except:
                        return time_str
                
                processed_df[b_column_index] = processed_df[b_column_index].apply(format_time)
                
            except:
                # If datetime conversion fails, sort as strings
                processed_df = processed_df.sort_values(
                    by=[g_column_index, b_column_index], 
                    ascending=[True, False],  # G列升序(相同交易对方聚集), B列降序(时间从近到远)
                    na_position='last'
                )
            
            processed_df = processed_df.reset_index(drop=True)
        
        return processed_df
    
    except Exception as e:
        st.error(f"数据处理失败: {str(e)}")
        return None

def save_to_excel(df, save_path):
    """
    Save DataFrame to Excel file with merged cells and adjusted column widths
    """
    try:
        # Create Excel file using pandas ExcelWriter
        with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, header=False)
            
            # Get the workbook and worksheet for formatting
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            
            # Apply formatting and merge cells strategically
            if worksheet.max_row > 0 and worksheet.max_column > 0:
                # Merge cells in the first row if it exists (as header)
                if worksheet.max_row >= 1 and worksheet.max_column > 1:
                    try:
                        worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=min(worksheet.max_column, 5))
                    except Exception:
                        pass
                
                # Adjust column widths and formatting
                from openpyxl.utils import get_column_letter
                from openpyxl.styles import Alignment
                
                # Get default column width (Excel default is approximately 8.43)
                default_width = 8.43
                a_column_width = default_width * 2    # A列2倍宽度
                f_column_width = default_width * 3    # F列保持3倍宽度
                
                # Set column A (交易时间) width to 2x and add left alignment
                if worksheet.max_column >= 1:
                    worksheet.column_dimensions[get_column_letter(1)].width = a_column_width
                    
                    # Apply left alignment to column A
                    for row in range(1, worksheet.max_row + 1):
                        cell = worksheet.cell(row=row, column=1)
                        cell.alignment = Alignment(horizontal='left')
                
                # Set column F (交易对方) width to 3x
                if worksheet.max_column >= 6:
                    worksheet.column_dimensions[get_column_letter(6)].width = f_column_width
        
        return True
    
    except Exception as e:
        st.error(f"Excel文件保存失败: {str(e)}")
        return False

def main():
    st.title("📄 PDF到Excel转换神器")
    st.markdown("---")
    style_metric_cards(background_color="#e0e7ef", border_left_color="#38b6ff")
    st.header("1. 上传PDF文件")
    uploaded_file = st.file_uploader(
        "选择PDF文件", 
        type="pdf",
        help="支持微信支付明细PDF文件，最大200MB"
    )
    if uploaded_file is not None:
        file_data = uploaded_file.getvalue()
        file_size_mb = len(file_data) / (1024 * 1024)
        st.info(f"文件名: {uploaded_file.name}")
        st.info(f"文件大小: {file_size_mb:.1f}MB")
        if file_size_mb > 200:
            st.error("文件过大，请选择小于200MB的文件")
        elif file_size_mb == 0:
            st.error("文件为空，请选择有效的PDF文件")
        elif not file_data.startswith(b'%PDF'):
            st.error("文件格式不正确，请选择有效的PDF文件")
        else:
            st.success("文件验证通过，自动开始分析...")
            import tempfile as temp_module
            with temp_module.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_file_path = tmp_file.name
            try:
                st.subheader("2. 分析进度")
                progress_bar = st.progress(0)
                status_text = st.empty()
                # 步骤1：提取PDF数据
                status_text.text("正在提取PDF数据...")
                progress_bar.progress(20)
                df = extract_pdf_to_dataframe(tmp_file_path)
                if df is not None:
                    progress_bar.progress(40)
                    status_text.text("PDF数据提取完成")
                    st.subheader("3. 数据预览")
                    st.dataframe(df.head(10))
                    st.info(f"提取到 {len(df)} 行数据，{len(df.columns)} 列")
                    # 步骤2：处理数据
                    progress_bar.progress(60)
                    status_text.text("正在处理Excel数据...")
                    processed_df = process_excel_data(df)
                    if processed_df is not None:
                        progress_bar.progress(80)
                        status_text.text("数据处理完成")
                        st.subheader("4. 处理后的数据预览")
                        st.dataframe(processed_df.head(10))
                        st.info(f"处理后有 {len(processed_df)} 行数据，{len(processed_df.columns)} 列")
                        # 步骤3：保存Excel
                        progress_bar.progress(90)
                        status_text.text("正在保存Excel文件...")
                        save_dir = os.path.expanduser("~/Downloads")
                        if not os.path.exists(save_dir):
                            save_dir = temp_module.gettempdir()
                        base_name = uploaded_file.name.rsplit('.', 1)[0]
                        excel_filename = f"{base_name}_processed.xlsx"
                        save_path = os.path.join(save_dir, excel_filename)
                        counter = 1
                        while os.path.exists(save_path):
                            excel_filename = f"{base_name}_processed_{counter}.xlsx"
                            save_path = os.path.join(save_dir, excel_filename)
                            counter += 1
                        if save_to_excel(processed_df, save_path):
                            progress_bar.progress(100)
                            status_text.text(f"处理完成！已自动保存到 {save_path}")
                            st.success(f"✅ 文件处理完成！已自动保存到 {save_path}")
                            with open(save_path, 'rb') as f:
                                excel_data = f.read()
                            st.download_button(
                                label="下载Excel文件",
                                data=excel_data,
                                file_name=excel_filename,
                                mime="application/vnd.ms-excel"
                            )
                else:
                    progress_bar.progress(0)
                    status_text.text("PDF数据提取失败")
            except Exception as e:
                st.error(f"处理过程中发生错误: {str(e)}")
                st.error("详细错误信息:")
                st.code(traceback.format_exc())
            finally:
                try:
                    os.unlink(tmp_file_path)
                except:
                    pass
    else:
        st.info("👆 请上传PDF文件开始处理")
        st.subheader("使用说明")
        st.markdown("""
        1. **上传PDF文件**: 点击上方的文件上传按钮选择PDF文件
        2. **极速转换**: 应用会自动将PDF转换为Excel格式
        3. **数据处理**: 执行自动化操作
        4. **保存文件**: 处理后的文件将自动保存到下载目录
        5. **下载选项**: 提供文件下载功能
        """)
        st.subheader("注意事项")
        st.warning("""
        - 请确保PDF文件包含表格或结构化数据
        - 处理大文件可能需要较长时间
        - 确保有足够的磁盘空间保存处理后的文件
        - 如果指定保存路径不存在，文件将保存到应用程序目录
        """)

if __name__ == "__main__":
    main()
