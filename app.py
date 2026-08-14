def apply_excel_styling(worksheet, data_df, start_row):
    """应用Excel样式（包含条件格式 + 优化列宽）"""
    
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
    
    # ========== 找出需要条件格式的列 ==========
    earnings_pct_idx = None
    impressions_pct_idx = None
    for idx, col_name in enumerate(columns, 1):
        if col_name == 'Earnings %':
            earnings_pct_idx = idx
        elif col_name == 'Impressions %':
            impressions_pct_idx = idx
    
    # ========== Excel条件格式风格的颜色 ==========
    # 绿色：浅绿色 (Excel条件格式默认绿色 #C6EFCE)
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    # 红色：浅红色 (Excel条件格式默认红色 #FFC7CE)
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    
    # ========== 设置表头 ==========
    worksheet.row_dimensions[start_row].height = 30
    for col_idx, col_name in enumerate(columns, 1):
        cell = worksheet.cell(row=start_row, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    
    # ========== 设置数据行 ==========
    for row_idx in range(start_row + 1, start_row + len(data_df) + 1):
        worksheet.row_dimensions[row_idx].height = 22
        
        # 基础交替颜色
        base_fill = even_fill if (row_idx - start_row - 1) % 2 == 1 else odd_fill
        
        for col_idx, col_name in enumerate(columns, 1):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell.font = content_font
            cell.fill = base_fill
            cell.border = thin_border
            
            # 对齐方式
            if col_name == 'Rank':
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name == 'Ad Unit':
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center")
            
            # 数字格式
            if col_name == 'Earnings (USD)':
                cell.number_format = '"$"#,##0.00'
            elif col_name == 'eCPM (USD)':
                cell.number_format = '"$"#,##0.00'
            elif col_name in ['Earnings %', 'Match Rate (%)', 'Impressions %', 'Requests %']:
                cell.number_format = '0.00"%"'
            elif col_name in ['Impressions', 'Requests']:
                cell.number_format = '#,##0'
        
        # ========== 条件格式：Earnings %（使用柔和的Excel风格颜色） ==========
        if earnings_pct_idx:
            cell = worksheet.cell(row=row_idx, column=earnings_pct_idx)
            value = cell.value
            if value is not None:
                if value > 10:
                    cell.fill = green_fill  # 柔和的绿色
                elif value < 2:
                    cell.fill = red_fill    # 柔和的红色
        
        # ========== 条件格式：Impressions %（使用柔和的Excel风格颜色） ==========
        if impressions_pct_idx:
            cell = worksheet.cell(row=row_idx, column=impressions_pct_idx)
            value = cell.value
            if value is not None:
                if value > 10:
                    cell.fill = green_fill  # 柔和的绿色
                elif value < 2:
                    cell.fill = red_fill    # 柔和的红色
    
    # ========== 优化列宽（Ad Unit 列宽调窄） ==========
    column_widths = {
        'Rank': 6,                  # 更窄
        'Ad Unit': 30,              # 从45缩减到30
        'Earnings (USD)': 16,
        'Earnings %': 14,
        'Impressions': 14,
        'Impressions %': 14,
        'Requests': 14,
        'Requests %': 14,
        'eCPM (USD)': 14,
        'Match Rate (%)': 16
    }
    
    for col_idx, col_name in enumerate(columns, 1):
        column_letter = get_column_letter(col_idx)
        
        # 从预定义宽度获取基础宽度
        if col_name in column_widths:
            base_width = column_widths[col_name]
        else:
            base_width = max(len(col_name) + 3, 12)
        
        # 检查数据内容长度（限制检查范围，提高性能）
        max_data_length = 0
        for row in range(start_row + 1, min(start_row + len(data_df) + 1, start_row + 200)):
            cell_value = worksheet.cell(row=row, column=col_idx).value
            if cell_value is not None:
                # 处理数字格式化的长度
                if col_name in ['Earnings (USD)', 'eCPM (USD)']:
                    length = len(f"${cell_value:,.2f}") if isinstance(cell_value, (int, float)) else len(str(cell_value))
                elif col_name in ['Earnings %', 'Impressions %', 'Requests %', 'Match Rate (%)']:
                    length = len(f"{cell_value:.2f}%") if isinstance(cell_value, (int, float)) else len(str(cell_value))
                elif col_name in ['Impressions', 'Requests']:
                    length = len(f"{cell_value:,.0f}") if isinstance(cell_value, (int, float)) else len(str(cell_value))
                else:
                    length = len(str(cell_value))
                
                if length > max_data_length:
                    max_data_length = min(length, 50)
        
        # 计算最终宽度
        header_length = len(col_name) + 2
        final_width = max(header_length, max_data_length + 2, base_width)
        
        # Ad Unit 列限制最大宽度为30
        if col_name == 'Ad Unit':
            final_width = min(final_width, 30)
        else:
            final_width = min(final_width, 22)
        
        worksheet.column_dimensions[column_letter].width = final_width
