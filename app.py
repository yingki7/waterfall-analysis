import streamlit as st
import pandas as pd
import numpy as np
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
import io

st.set_page_config(page_title="广告瀑布流分析工具", layout="wide")

st.title("📊 广告单元瀑布流分析工具")
st.markdown("上传CSV文件，自动生成Top 5国家的广告单元分析报告")

def apply_excel_styling(worksheet, data_df, start_row):
    content_font = Font(name='等线', size=11)
    header_font = Font(name='等线', bold=True, size=12, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    even_fill = PatternFill(start_color="F5F9FF", end_color="F5F9FF", fill_type="solid")
    odd_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    thin_border = Border(left=Side(style='thin', color='D0D0D0'), right=Side(style='thin', color='D0D0D0'),
                        top=Side(style='thin', color='D0D0D0'), bottom=Side(style='thin', color='D0D0D0'))
    columns = data_df.columns.tolist()
    
    worksheet.row_dimensions[start_row].height = 25
    for col_idx, col_name in enumerate(columns, 1):
        cell = worksheet.cell(row=start_row, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    
    for row_idx in range(start_row + 1, start_row + len(data_df) + 1):
        worksheet.row_dimensions[row_idx].height = 20
        row_fill = even_fill if (row_idx - start_row - 1) % 2 == 1 else odd_fill
        for col_idx, col_name in enumerate(columns, 1):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell.font = content_font
            cell.fill = row_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center" if col_name == 'Rank' else "left" if col_name == 'Ad Unit' else "right", vertical="center")
            if col_name in ['Earnings (USD)', 'eCPM (USD)']:
                cell.number_format = '"$"#,##0.00'
            elif col_name in ['Earnings %', 'Match Rate (%)', 'Impressions %', 'Requests %']:
                cell.number_format = '0.00"%"'
            elif col_name in ['Impressions', 'Requests']:
                cell.number_format = '#,##0'
    
    for col_idx in range(1, len(columns) + 1):
        column_letter = get_column_letter(col_idx)
        max_length = len(columns[col_idx - 1])
        for row in range(start_row, min(start_row + len(data_df) + 1, start_row + 100)):
            cell_value = worksheet.cell(row=row, column=col_idx).value
            if cell_value:
                max_length = max(max_length, min(len(str(cell_value)), 50))
        adjusted_width = min(max_length + 3, 45) if columns[col_idx - 1] == 'Ad Unit' else min(max_length + 2, 20)
        worksheet.column_dimensions[column_letter].width = adjusted_width

def process_csv(file_content):
    df = pd.read_csv(io.BytesIO(file_content), sep='\t', encoding='utf-16')
    
    numeric_columns = ['Estimated earnings (USD)', 'Observed eCPM (USD)', 'Requests',
                      'Matched requests', 'Show rate', 'Impressions', 'CTR', 'Clicks',
                      'Ads ARPV (USD)', 'Ads ARPU (USD)', 'Ad viewers (AV)', 'Active users (AU)',
                      'Ad viewer rate', 'Imps / AV', 'Imps / AU', 'Ads ARPDAV (USD)', 'Ads ARPDAU (USD)',
                      'DAV', 'DAU', 'Daily ad viewer rate', 'IMPDAV', 'IMPDAU', 'Ad load latency',
                      'Bid requests', 'Bids in auction (%)', 'Bids in auction', 'Win rate', 'Winning bids']
    for col in numeric_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    
    df['Match rate'] = pd.to_numeric(df['Match rate'].astype(str).str.replace('%', '', regex=False), errors='coerce')
    
    country_earnings = df.groupby('Country')['Estimated earnings (USD)'].sum().sort_values(ascending=False)
    top_5_countries = country_earnings.head(5).index.tolist()
    
    interstitial_native_df = df[
        df['Ad unit'].str.contains('Interstitial|Native|Full|Inter', case=False, na=False) &
        ~df['Ad unit'].str.contains('banner', case=False, na=False)
    ].copy()
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        summary_data = []
        for country in top_5_countries:
            country_data = interstitial_native_df[interstitial_native_df['Country'] == country].copy()
            if len(country_data) > 0:
                total_earnings = country_data['Estimated earnings (USD)'].sum()
                total_impressions = country_data['Impressions'].sum()
                total_requests = country_data['Requests'].sum()
                
                result_df = pd.DataFrame({
                    'Ad Unit': country_data['Ad unit'],
                    'Earnings (USD)': country_data['Estimated earnings (USD)'].round(2),
                    'Earnings %': (country_data['Estimated earnings (USD)'] / total_earnings * 100).round(2),
                    'Impressions': country_data['Impressions'].fillna(0).astype(int),
                    'Impressions %': (country_data['Impressions'] / total_impressions * 100).round(2),
                    'Requests': country_data['Requests'].fillna(0).astype(int),
                    'Requests %': (country_data['Requests'] / total_requests * 100).round(2),
                    'eCPM (USD)': country_data['Observed eCPM (USD)'].round(2),
                    'Match Rate (%)': country_data['Match rate'].round(2)
                }).fillna(0).sort_values('eCPM (USD)', ascending=False).reset_index(drop=True)
                result_df.insert(0, 'Rank', range(1, len(result_df) + 1))
                
                app_name = country_data['App'].mode()[0] if 'App' in country_data.columns and not country_data['App'].mode().empty else "Unknown"
                sheet_name = re.sub(r'[\\/*?:"<>|]', '', country)[:31]
                start_row = 6
                result_df.to_excel(writer, sheet_name=sheet_name, startrow=start_row - 1, index=False)
                worksheet = writer.sheets[sheet_name]
                
                # 设置默认字体
                for row in worksheet.iter_rows():
                    for cell in row:
                        cell.font = Font(name='等线', size=10)
                
                # App名称
                worksheet.cell(row=1, column=1, value=f" App: {app_name}").font = Font(name='等线', bold=True, size=14)
                worksheet.cell(row=1, column=1).alignment = Alignment(horizontal="left", vertical="center")
                worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)
                worksheet.row_dimensions[1].height = 28
                
                # 标题
                worksheet.cell(row=2, column=1, value=f"📊 {country} - 广告单元分析报告").font = Font(name='等线', bold=True, size=16)
                worksheet.cell(row=2, column=1).alignment = Alignment(horizontal="left", vertical="center")
                worksheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=9)
                worksheet.row_dimensions[2].height = 30
                
                # 统计信息
                stats = [
                    (f"总收益: ${total_earnings:,.2f}", 3, 1),
                    (f"总展示次数: {total_impressions:,.0f}", 3, 3),
                    (f"总请求数: {total_requests:,.0f}", 3, 5),
                    (f"Ad Units数量: {len(country_data)}", 3, 7),
                    (f"平均eCPM: ${result_df['eCPM (USD)'].mean():.2f}", 4, 1),
                    (f"平均Match Rate: {result_df['Match Rate (%)'].mean():.2f}%", 4, 3),
                    (f"最高eCPM: ${result_df['eCPM (USD)'].max():.2f}", 4, 5),
                    (f"最低eCPM: ${result_df['eCPM (USD)'].min():.2f}", 4, 7),
                ]
                for text, row, col in stats:
                    cell = worksheet.cell(row=row, column=col, value=text)
                    cell.font = Font(name='等线', size=11, bold=True)
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                
                worksheet.row_dimensions[3].height = 22
                worksheet.row_dimensions[4].height = 22
                
                # 分隔线
                for col in range(1, 10):
                    cell = worksheet.cell(row=5, column=col, value="")
                    cell.border = Border(bottom=Side(style='medium', color='4472C4'))
                
                apply_excel_styling(worksheet, result_df, start_row)
                worksheet.freeze_panes = f'A{start_row + 1}'
                
                summary_data.append({
                    'App Name': app_name,
                    'Country': country,
                    'Total Earnings (USD)': round(total_earnings, 2),
                    'Total Impressions': int(total_impressions),
                    'Total Requests': int(total_requests),
                    'Number of Ad Units': len(country_data),
                    'Avg eCPM (USD)': round(result_df['eCPM (USD)'].mean(), 2),
                    'Avg Match Rate (%)': round(result_df['Match Rate (%)'].mean(), 2),
                    'Max eCPM (USD)': round(result_df['eCPM (USD)'].max(), 2),
                    'Min eCPM (USD)': round(result_df['eCPM (USD)'].min(), 2)
                })
        
        # Summary sheet
        if summary_data:
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False, startrow=2)
            worksheet = writer.sheets['Summary']
            
            # 设置默认字体
            for row in worksheet.iter_rows():
                for cell in row:
                    cell.font = Font(name='等线', size=10)
            
            worksheet.cell(row=1, column=1, value="📈 Top 5国家广告单元汇总分析").font = Font(name='等线', bold=True, size=16)
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
            thin_border = Border(left=Side(style='thin', color='D0D0D0'), right=Side(style='thin', color='D0D0D0'),
                                top=Side(style='thin', color='D0D0D0'), bottom=Side(style='thin', color='D0D0D0'))
            for row_idx in range(header_row + 1, len(summary_df) + header_row + 1):
                worksheet.row_dimensions[row_idx].height = 20
                row_fill = even_fill if (row_idx - header_row - 1) % 2 == 1 else PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
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
                        max_length = max(max_length, min(len(str(cell_value)), 25))
                worksheet.column_dimensions[column_letter].width = max_length + 2
    
    output.seek(0)
    return output

# ========== Streamlit 界面 ==========
uploaded_file = st.file_uploader("📁 上传CSV文件", type=['csv', 'txt'])

if uploaded_file is not None:
    with st.spinner('🔄 正在分析数据，请稍候...'):
        try:
            # 显示数据预览
            df_preview = pd.read_csv(io.BytesIO(uploaded_file.read()), sep='\t', encoding='utf-16')
            with st.expander("📋 数据预览（前10行）"):
                st.dataframe(df_preview.head(10))
            
            # 重新读取文件进行处理（因为上面读取后指针已到末尾）
            uploaded_file.seek(0)
            excel_data = process_csv(uploaded_file.read())
            st.success("✅ 分析完成！")
            
            # 获取App名称用于文件名
            app_name = "ad_analysis"
            try:
                uploaded_file.seek(0)
                df = pd.read_csv(io.BytesIO(uploaded_file.read()), sep='\t', encoding='utf-16')
                if 'App' in df.columns and not df['App'].mode().empty:
                    app_name = re.sub(r'[\\/*?:"<>|]', '', str(df['App'].mode()[0]))[:50]
            except:
                pass
            
            st.download_button(
                label="📥 下载Excel报告",
                data=excel_data,
                file_name=f"{app_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.info("💡 报告已生成，点击上方按钮下载")
            
        except Exception as e:
            st.error(f"❌ 处理失败: {str(e)}")
            st.exception(e)

st.markdown("---")
st.caption("💡 数据格式要求：UTF-16编码的Tab分隔CSV文件，必须包含 App, Country, Ad unit 等列")
