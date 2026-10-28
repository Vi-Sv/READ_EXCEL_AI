Sub ExportSheetForAI()
    Dim ws As Worksheet
    Dim usedRng As Range
    Dim row As Range
    Dim cell As Range
    Dim fileNo As Integer
    Dim desktopPath As String
    Dim filePath As String
    Dim lineStr As String
    Dim cellInfo As String
    Dim hasFormulas As Boolean
    
    Set ws = ActiveSheet
    Set usedRng = ws.UsedRange
    
    ' Определение пути к рабочему столу
    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    filePath = desktopPath & "\" & ws.Name & "_for_AI.txt"
    
    fileNo = FreeFile
    Open filePath For Output As #fileNo
    
    ' Запись системной метаинформации для ИИ
    Print #fileNo, "### МЕТАДАННЫЕ ЛИСТА EXCEL ###"
    Print #fileNo, "Название листа: " & ws.Name
    Print #fileNo, "Диапазон данных: " & usedRng.Address
    Print #fileNo, "Количество строк: " & usedRng.Rows.Count
    Print #fileNo, "Количество столбцов: " & usedRng.Columns.Count
    Print #fileNo, "--------------------------------------------------"
    Print #fileNo, ""
    
    ' Блок 1: Экспорт структуры таблицы в формате Markdown
    Print #fileNo, "### СТРУКТУРА ТАБЛИЦЫ (MARKDOWN) ###"
    For Each row In usedRng.Rows
        lineStr = "|"
        For Each cell In row.Cells
            ' Замена внутренних переносов строк на пробелы, чтобы не ломать структуру Markdown
            lineStr = lineStr & Replace(cell.Text, vbCrLf, " ") & "|"
        Next cell
        Print #fileNo, lineStr
        
        ' Добавление разделителя после первой строки (шапки)
        If row.row = usedRng.Rows(1).row Then
            lineStr = "|"
            For Each cell In row.Cells
                lineStr = lineStr & "---|"
            Next cell
            Print #fileNo, lineStr
        End If
    Next row
    Print #fileNo, ""
    Print #fileNo, "--------------------------------------------------"
    Print #fileNo, ""
    
    ' Блок 2: Поячеистый технический аудит (формулы и скрытые свойства)
    Print #fileNo, "### ТЕХНИЧЕСКИЙ АУДИТ ЯЧЕЕК И ФОРМУЛ ###"
    hasFormulas = False
    
    For Each cell In usedRng
        If cell.HasFormula Then
            hasFormulas = True
            cellInfo = "Адрес: " & cell.Address(False, False) & _
                       " | Значение: " & cell.Text & _
                       " | Формула: " & cell.FormulaLocal
            Print #fileNo, cellInfo
        End If
    Next cell
    
    If Not hasFormulas Then
        Print #fileNo, "В указанном диапазоне формулы не обнаружены."
    End If
    
    Close #fileNo
    
    MsgBox "Данные листа успешно оцифрованы и сохранены на Рабочий стол под именем: " & _
           ws.Name & "_for_AI.txt", vbInformation, "Успешный экспорт"
End Sub
