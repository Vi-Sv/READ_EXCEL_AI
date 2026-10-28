Sub ExportFullSheetSpecsForAI()
    Dim ws As Worksheet
    Dim usedRng As Range
    Dim cell As Range
    Dim fileNo As Integer
    Dim desktopPath As String
    Dim filePath As String
    Dim borderStr As String
    Dim alignmentStr As String
    
    Set ws = ActiveSheet
    Set usedRng = ws.UsedRange
    
    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    filePath = desktopPath & "\" & ws.Name & "_FULL_SPECS_FOR_AI.txt"
    
    fileNo = FreeFile
    Open filePath For Output As #fileNo
    
    ' Метаданные листа
    Print #fileNo, "=== АРХИТЕКТУРНЫЙ И ТЕХНИЧЕСКИЙ ПАСПОРТ ЛИСТА EXCEL ДЛЯ ИИ ==="
    Print #fileNo, "Лист: " & ws.Name
    Print #fileNo, "Диапазон: " & usedRng.Address
    Print #fileNo, "Дата экспорта: " & Date & " " & Time
    Print #fileNo, "=============================================================="
    Print #fileNo, ""
    
    ' Считывание глобальной геометрии столбцов и строк
    Print #fileNo, "### ГЕОМЕТРИЯ СЕТКИ ЛИСТА ###"
    Dim col As Range
    For Each col In usedRng.Columns
        Print #fileNo, "Колонка " & Split(col.Address, "$")(1) & " | Ширина: " & Round(col.ColumnWidth, 1)
    Next col
    Dim rw As Range
    For Each row In usedRng.Rows
        Print #fileNo, "Строка " & row.row & " | Высота: " & Round(row.RowHeight, 1)
    Next row
    Print #fileNo, ""
    Print #fileNo, "=============================================================="
    Print #fileNo, ""
    
    ' Поячеистый детальный разбор контента и оформления
    Print #fileNo, "### ПОДРОБНАЯ СПЕЦИФИКАЦИЯ КАЖДОЙ ЯЧЕЙКИ ###"
    
    For Each cell In usedRng
        ' Пропускаем абсолютно пустые нераскрашенные ячейки без границ для экономии объема файла
        If cell.Text = "" And cell.Interior.ColorIndex = xlNone And cell.Borders(xlEdgeLeft).LineStyle = xlNone And cell.Borders(xlEdgeRight).LineStyle = xlNone And cell.Borders(xlEdgeTop).LineStyle = xlNone And cell.Borders(xlEdgeBottom).LineStyle = xlNone Then
            GoTo NextCell
        End If
        
        Print #fileNo, "[" & cell.Address(False, False) & "]"
        Print #fileNo, "  Текст: " & Replace(cell.Text, vbCrLf, " [ПЕРЕНОС СТРОКИ] ")
        
        If cell.HasFormula Then
            Print #fileNo, "  Формула: " & cell.FormulaLocal
        End If
        
        ' Шрифты
        Print #fileNo, "  Шрифт: " & cell.Font.Name & " | Размер: " & cell.Font.Size & " | Цвет RGB: " & cell.Font.Color
        If cell.Font.Bold Then Print #fileNo, "  Стиль: ЖИРНЫЙ"
        If cell.Font.Italic Then Print #fileNo, "  Стиль: КУРСИВ"
        
        ' Заливка ячейки
        If cell.Interior.ColorIndex <> xlNone Then
            Print #fileNo, "  Заливка RGB: " & cell.Interior.Color
        End If
        
        ' Выравнивание
        alignmentStr = ""
        Select Case cell.HorizontalAlignment
            Case xlHAlignLeft: alignmentStr = "По левому краю"
            Case xlHAlignCenter: alignmentStr = "По центру"
            Case xlHAlignRight: alignmentStr = "По правому краю"
        End Select
        If cell.WrapText Then alignmentStr = alignmentStr & " (С переносом текста)"
        If alignmentStr <> "" Then Print #fileNo, "  Выравнивание: " & alignmentStr
        
        ' Границы
        borderStr = ""
        If cell.Borders(xlEdgeLeft).LineStyle <> xlNone Then borderStr = borderStr & "Левая "
        If cell.Borders(xlEdgeRight).LineStyle <> xlNone Then borderStr = borderStr & "Правая "
        If cell.Borders(xlEdgeTop).LineStyle <> xlNone Then borderStr = borderStr & "Верхняя "
        If cell.Borders(xlEdgeBottom).LineStyle <> xlNone Then borderStr = borderStr & "Нижняя "
        If borderStr <> "" Then
            Print #fileNo, "  Границы: " & Trim(borderStr)
        End If
        
        ' Объединение ячеек
        If cell.MergeCells Then
            Print #fileNo, "  Объединение: В составе диапазона " & cell.MergeArea.Address(False, False)
        End If
        
        Print #fileNo, "  ---"
NextCell:
    Next cell
    
    Close #fileNo
    
    MsgBox "Полный технический и визуальный паспорт листа успешно сохранен на Рабочий стол под именем: " & _
           ws.Name & "_FULL_SPECS_FOR_AI.txt", vbInformation, "Экспорт завершен"
End Sub
