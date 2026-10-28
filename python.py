import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def regenerate_excel_from_json(json_file_path, output_xlsx_path):
    with open(json_file_path, 'r', encoding='utf-16') as f:
        data = json.load(f)
        
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = data['sheet_name']
    ws.views.sheetView[0].showGridLines = True
    
    # Восстановление ширины колонок
    for col_letter, width in data['geometry']['columns'].items():
        ws.column_dimensions[col_letter].width = width
        
    # Восстановление высоты строк
    for row_num, height in data['geometry']['rows'].items():
        ws.row_dimensions[int(row_num)].height = height
        
    # Сетка тонких границ ячеек
    thin_side = Side(border_style="thin", color="000000")
    
    # Обработка каждой считанной ячейки
    for c_data in data['cells']:
        cell = ws[c_data['addr']]
        
        # Запись контента
        if c_data['f']:
            cell.value = c_data['f'] # Запись формулы
        else:
            cell.value = c_data['v'] # Запись текста
            
        # Конвертация цвета из Long VBA в HEX openpyxl
        bg_color = f"{c_data['bg']:06X}"[2:] if c_data['bg'] != 16777215 else "FFFFFF"
        fg_color = f"{c_data['fc']:06X}"[2:] if c_data['fc'] != 0 else "000000"
        
        cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
        cell.font = Font(name="Times New Roman", size=c_data['sz'], bold=c_data['bold'], italic=c_data['italic'], color=fg_color)
        
        # Выравнивание
        ha_align = "left"
        if c_data['ha'] == -4108: ha_align = "center"
        if c_data['ha'] == -4152: ha_align = "right"
        cell.alignment = Alignment(horizontal=ha_align, vertical="top", wrap_text=c_data['wrap'])
        
        # Границы
        b_list = c_data['b']
        cell.border = Border(
            left=thin_side if b_list[0] == 1 else None,
            right=thin_side if b_list[1] == 1 else None,
            top=thin_side if b_list[2] == 1 else None,
            bottom=thin_side if b_list[3] == 1 else None
        )
        
        # Объединение диапазонов ячеек
        if c_data['merged'] and c_data['ma']:
            try:
                ws.merge_cells(c_data['ma'])
            except:
                pass # Пропуск дублирующихся адресов объединенной области
                
    wb.save(output_xlsx_path)


