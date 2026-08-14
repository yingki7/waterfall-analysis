import pandas as pd
import numpy as np
from pathlib import Path
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
import re
import tempfile
import gradio as gr

def process_csv(file_obj):
    """处理上传的CSV文件，返回Excel文件路径"""
    
    # ==========================================
    # 1. 读取上传的CSV文件
    # ==========================================
    try:
        # file_obj 是 Gradio 上传的文件对象，有 .name 属性
        df = pd.read_csv(file_obj.name, sep='\t', encoding='utf-16')
        print(f"✅ 成功读取文件，共 {len(df)} 行")
    except Exception as e:
        return f"❌ 读取文件失败: {str(e)}", None
    
    # 获取App名称
    if 'App' in df.columns:
        main_app_name = df['App'].mode()[0] if not df['App'].mode().empty else "Unknown"
        main_app_name = re.sub(r'[\\/*?:"<>|]', '', main_app_name)
        if len(main_app_name) > 50:
            main_app_name = main_app_name[:50]
    else:
        main_app_name = "ad_analysis"
    
    # 使用临时文件保存Excel
    temp_dir = tempfile.mkdtemp()
    output_filename = os.path.join(temp_dir, f"{main_app_name}.xlsx")
    
    # 如果文件存在，删除
    if os.path.exists(output_filename):
        os.remove(output_filename)
    
    # ==========================================
    # 2. 数据处理（你的原有逻辑）
    # ==========================================
    numeric_columns = [
        'Estimated earnings (USD)', 'Observed eCPM (USD)', 'Requests',
        'Matched requests', 'Show rate', 'Impressions', 'CTR', 'Clicks',
        'Ads ARPV (USD)', 'Ads ARPU (USD)', 'Ad viewers (AV)', 'Active users (AU)',
        'Ad viewer rate', 'Imps / AV', 'Imps / AU', 'Ads ARPDAV (USD)', 'Ads ARPDAU (USD)',
        'DAV', 'DAU', 'Daily ad viewer rate', 'IMPDAV', 'IMPDAU', 'Ad load latency',
        'Bid requests', 'Bids in auction (%)', 'Bids in auction', 'Win rate', 'Winning bids'
    ]
    
    for col in numeric_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    
    # 处理Match rate
    df['Match rate'] = df['Match rate'].astype(str).str.replace('%', '', regex=False)
    df['Match rate'] = pd.to_numeric(df['Match rate'], errors='coerce')
    
    # 计算Top 5国家
    country_earnings = df.groupby('Country')['Estimated earnings (USD)'].sum().sort_values(ascending=False)
    top_5_countries = country_earnings.head(5).index.tolist()
    
    # 筛选Ad Unit
    interstitial_native_df = df[
        df['Ad unit'].str.contains('Interstitial|Native|Full|Inter', case=False, na=False) &
        ~df['Ad unit'].str.contains('banner', case=False, na=False)
    ].copy()
    
    # ==========================================
    # 3. 辅助函数（完全保留你的样式逻辑）
    # ==========================================
    def apply_excel_styling(worksheet, data_df, start_row):
        """应用Excel样式"""
        content_font = Font(name='等线', size=11)
        header_font = Font(name='等线', bold=True, size=12, color="FFFFFF")
        
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        even_fill = PatternFill(start_color="F5F9FF", end_color="F5F9FF", fill_type="solid")
        odd_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        
        thin_border = Border(
            left=Side(style='thin', color='D0D0D0'),
            right=Side(style='thin', color='D0D0D0'),
            top=Side(style='thin', color='D0D0D0'),
            bottom=Side(style='thin', color='D0D0D0')
        )
        
        columns = data_df.columns.tolist()
        
        # 表头
        worksheet.row_dimensions[start_row].height = 25
        for col_idx, col_name in enumerate(columns, 1):
            cell = worksheet.cell(row=start_row, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        
        # 数据行
        for row_idx in range(start_row + 1, start_row + len(data_df) + 1):
            worksheet.row_dimensions[row_idx].height = 20
            
            if (row_idx - start_row - 1) % 2 == 1:
                row_fill = even_fill
            else:
                row_fill = odd_fill
            
            for col_idx, col_name in enumerate(columns, 1):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                cell.font = content_font
                cell.fill = row_fill
                cell.border = thin_border
                
                if col_name == 'Rank':
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_name == 'Ad Unit':
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                
                if col_name == 'Earnings (USD)':
                    cell.number_format = '"$"#,##0.00'
                elif col_name == 'eCPM (USD)':
                    cell.number_format = '"$"#,##0.00'
                elif col_name in ['Earnings %', 'Match Rate (%)', 'Impressions %', 'Requests %']:
                    cell.number_format = '0.00"%"'
                elif col_name in ['Impressions', 'Requests']:
                    cell.number_format = '#,##0'
        
        # 列宽
        for col_idx in range(1, len(columns) + 1):
            column_letter = get_column_letter(col_idx)
            max_length = len(columns[col_idx - 1])
            
            for row in range(start_row, min(start_row + len(data_df) + 1, start_row + 100)):
                cell_value = worksheet.cell(row=row, column=col_idx).value
                if cell_value:
                    length = len(str(cell_value))
                    if length > max_length:
                        max_length = min(length, 50)
            
            if columns[col_idx - 1] == 'Ad Unit':
                adjusted_width = min(max_length + 3, 45)
            else:
                adjusted_width = min(max_length + 2, 20)
            
            worksheet.column_dimensions[column_letter].width = adjusted_width

    def process_country_data(country_df, country_name):
        """处理单个国家数据"""
        total_earnings = country_df['Estimated earnings (USD)'].sum()
        total_impressions = country_df['Impressions'].sum()
        total_requests = country_df['Requests'].sum()
        
        result_df = pd.DataFrame()
        result_df['Ad Unit'] = country_df['Ad unit']
        result_df['Earnings (USD)'] = country_df['Estimated earnings (USD)'].round(2)
        result_df['Earnings %'] = (country_df['Estimated earnings (USD)'] / total_earnings * 100).round(2)
        result_df['Impressions'] = country_df['Impressions'].fillna(0).astype(int)
        result_df['Impressions %'] = (country_df['Impressions'] / total_impressions * 100).round(2)
        result_df['Requests'] = country_df['Requests'].fillna(0).astype(int)
        result_df['Requests %'] = (country_df['Requests'] / total_requests * 100).round(2)
        result_df['eCPM (USD)'] = country_df['Observed eCPM (USD)'].round(2)
        result_df['Match Rate (%)'] = country_df['Match rate'].round(2)
        
        result_df = result_df.fillna(0)
        result_df = result_df.sort_values('eCPM (USD)', ascending=False).reset_index(drop=True)
        result_df.insert(0, 'Rank', range(1, len(result_df) + 1))
        
        if 'App' in country_df.columns:
            app_name = country_df['App'].mode()[0] if not country_df['App'].mode().empty else "Unknown"
        else:
            app_name = "Not Specified"
        
        return result_df, total_earnings, total_impressions, total_requests, app_name

    # ==========================================
    # 4. 写入Excel
    # ==========================================
    summary_data = []
    
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        for country in top_5_countries:
            country_data = interstitial_native_df[interstitial_native_df['Country'] == country].copy()
            
            if len(country_data) > 0:
                processed_data, total_earnings, total_impressions, total_requests, app_name = process_country_data(country_data, country)
                
                sheet_name = country.replace('/', '_').replace('\\', '_').replace(':', '_')[:31]
                start_row = 6
                processed_data.to_excel(writer, sheet_name=sheet_name, startrow=start_row - 1, index=False)
                
                worksheet = writer.sheets[sheet_name]
                
                for row in worksheet.iter_rows():
                    for cell in row:
                        cell.font = Font(name='等线', size=10)
                
                # App名称
                app_cell = worksheet.cell(row=1, column=1, value=f" App: {app_name}")
                app_cell.font = Font(name='等线', bold=True, size=14)
                app_cell.alignment = Alignment(horizontal="left", vertical="center")
                worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)
                worksheet.row_dimensions[1].height = 28
                
                # 标题
                title_cell = worksheet.cell(row=2, column=1, value=f"📊 {country} - 广告单元分析报告")
                title_cell.font = Font(name='等线', bold=True, size=16)
                title_cell.alignment = Alignment(horizontal="left", vertical="center")
                worksheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=9)
                worksheet.row_dimensions[2].height = 30
                
                # 统计信息
                stats = [
                    (f"总收益: ${total_earnings:,.2f}", 3, 1),
                    (f"总展示次数: {total_impressions:,.0f}", 3, 3),
                    (f"总请求数: {total_requests:,.0f}", 3, 5),
                    (f"Ad Units数量: {len(country_data)}", 3, 7),
                    (f"平均eCPM: ${processed_data['eCPM (USD)'].mean():.2f}", 4, 1),
                    (f"平均Match Rate: {processed_data['Match Rate (%)'].mean():.2f}%", 4, 3),
                    (f"最高eCPM: ${processed_data['eCPM (USD)'].max():.2f}", 4, 5),
                    (f"最低eCPM: ${processed_data['eCPM (USD)'].min():.2f}", 4, 7),
                ]
                
                stat_font = Font(name='等线', size=11, bold=True)
                for text, row, col in stats:
                    cell = worksheet.cell(row=row, column=col, value=text)
                    cell.font = stat_font
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                
                worksheet.row_dimensions[3].height = 22
                worksheet.row_dimensions[4].height = 22
                
                # 分隔线
                for col in range(1, 10):
                    cell = worksheet.cell(row=5, column=col, value="")
                    cell.border = Border(bottom=Side(style='medium', color='4472C4'))
                
                apply_excel_styling(worksheet, processed_data, start_row)
                worksheet.freeze_panes = f'A{start_row + 1}'
                
                summary_data.append({
                    'App Name': app_name,
                    'Country': country,
                    'Total Earnings (USD)': round(total_earnings, 2),
                    'Total Impressions': int(total_impressions),
                    'Total Requests': int(total_requests),
                    'Number of Ad Units': len(country_data),
                    'Avg eCPM (USD)': round(processed_data['eCPM (USD)'].mean(), 2),
                    'Avg Match Rate (%)': round(processed_data['Match Rate (%)'].mean(), 2),
                    'Max eCPM (USD)': round(processed_data['eCPM (USD)'].max(), 2),
                    'Min eCPM (USD)': round(processed_data['eCPM (USD)'].min(), 2)
                })
    
    # ==========================================
    # 5. 创建Summary sheet
    # ==========================================
    summary_df = pd.DataFrame(summary_data)
    
    with pd.ExcelWriter(output_filename, engine='openpyxl', mode='a') as writer:
        summary_df.to_excel(writer, sheet_name='Summary', index=False, startrow=2)
        
        worksheet = writer.sheets['Summary']
        
        for row in worksheet.iter_rows():
            for cell in row:
                cell.font = Font(name='等线', size=10)
        
        title_cell = worksheet.cell(row=1, column=1, value="📈 Top 5国家广告单元汇总分析")
        title_cell.font = Font(name='等线', bold=True, size=16)
        worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)
        worksheet.row_dimensions[1].height = 30
        
        header_row = 3
        worksheet.row_dimensions[header_row].height = 25
        
        header_font = Font(name='等线', bold=True, size=12, color="FFFFFF")
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        
        for col_idx, col_name in enumerate(summary_df.columns, 1):
            cell = worksheet.cell(row=header_row, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
        
        content_font = Font(name='等线', size=11)
        even_fill = PatternFill(start_color="F5F9FF", end_color="F5F9FF", fill_type="solid")
        thin_border = Border(
            left=Side(style='thin', color='D0D0D0'),
            right=Side(style='thin', color='D0D0D0'),
            top=Side(style='thin', color='D0D0D0'),
            bottom=Side(style='thin', color='D0D0D0')
        )
        
        for row_idx in range(header_row + 1, len(summary_df) + header_row + 1):
            worksheet.row_dimensions[row_idx].height = 20
            
            if (row_idx - header_row - 1) % 2 == 1:
                row_fill = even_fill
            else:
                row_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
            
            for col_idx, col_name in enumerate(summary_df.columns, 1):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                cell.font = content_font
                cell.fill = row_fill
                cell.border = thin_border
                
                if col_name in ['Total Earnings (USD)', 'Avg eCPM (USD)', 'Max eCPM (USD)', 'Min eCPM (USD)']:
                    cell.number_format = '"$"#,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif col_name == 'Avg Match Rate (%)':
                    cell.number_format = '0.00"%"'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif col_name in ['Total Impressions', 'Total Requests']:
                    cell.number_format = '#,##0'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
        
        worksheet.freeze_panes = f'A{header_row + 1}'
        
        for col_idx in range(1, len(summary_df.columns) + 1):
            column_letter = get_column_letter(col_idx)
            max_length = len(summary_df.columns[col_idx - 1])
            for row in range(header_row, len(summary_df) + header_row + 1):
                cell_value = worksheet.cell(row=row, column=col_idx).value
                if cell_value:
                    length = len(str(cell_value))
                    if length > max_length:
                        max_length = min(length, 25)
            adjusted_width = max_length + 2
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    return "✅ 分析完成！点击下方下载按钮获取Excel报告", output_filename

# ==========================================
# 6. Gradio界面
# ==========================================
def create_interface():
    with gr.Blocks(title="广告单元分析工具", theme=gr.themes.Soft()) as demo:
        gr.Markdown("""
        # 📊 广告单元瀑布流分析工具
        
        上传你的CSV数据文件，系统将自动生成包含Top 5国家的广告单元分析报告。
        
        **数据格式要求：**
        - UTF-16编码，Tab分隔的CSV文件
        - 必须包含以下列：`App`, `Country`, `Ad unit`, `Estimated earnings (USD)`, `Impressions`, `Requests` 等
        """)
        
        with gr.Row():
            with gr.Column(scale=1):
                file_input = gr.File(
                    label="📁 上传CSV文件",
                    file_types=[".csv", ".txt"],
                    type="filepath"
                )
                submit_btn = gr.Button("🚀 开始分析", variant="primary")
            
            with gr.Column(scale=1):
                status_output = gr.Textbox(label="📋 处理状态", lines=3)
                file_output = gr.File(label="📥 下载Excel报告")
        
        submit_btn.click(
            fn=process_csv,
            inputs=file_input,
            outputs=[status_output, file_output]
        )
        
        # 添加示例说明
        gr.Markdown("""
        ---
        ### 📌 使用说明
        1. 点击"上传CSV文件"选择你的数据文件
        2. 点击"开始分析"按钮
        3. 等待处理完成后，点击下载按钮保存Excel报告
        
        **生成报告包含：**
        - Top 5国家的独立分析工作表
        - Summary汇总表
        - 自动格式化和样式美化
        """)
    
    return demo

# ==========================================
# 7. 启动应用
# ==========================================
if __name__ == "__main__":
    demo = create_interface()
    demo.launch()