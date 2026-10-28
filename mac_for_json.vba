Sub Export100PercentSheetToJSON()
    Dim ws As Worksheet
    Dim cell As Range
    Dim fso As Object, ts As Object
    Dim json As String, cellJson As String
    Dim desktopPath As String, filePath As String
    Dim i As Long, j As Long
    
    Set ws = ActiveSheet
    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    filePath = desktopPath & "\" & ws.Name & "_DATA_FOR_AI.json"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(filePath, True, True) ' Экспорт в UTF-16
    
    json = "{" & vbCrLf
    json = json & "  ""sheet_name"": """ & ws.Name & """," & vbCrLf
    json = json & "  ""geometry"": {" & vbCrLf
    json = json & "    ""columns"": {" & vbCrLf
    
    ' Считывание ширины колонок
    For i = 1 To ws.UsedRange.Columns.Count
        Dim colLetter As String
        colLetter = Split(ws.Cells(1, i).Address, "$")(1)
        json = json & "      """ & colLetter & """: " & Replace(Format(ws.Columns(i).ColumnWidth, "0.00"), ",", ".")
        If i < ws.UsedRange.Columns.Count Then json = json & ","
        json = json & vbCrLf
    Next i
    json = json & "    }," & vbCrLf
    json = json & "    ""rows"": {" & vbCrLf
    
    ' Считывание высоты строк
    For i = 1 To ws.UsedRange.Rows.Count
        json = json & "      """ & i & """: " & Replace(Format(ws.Rows(i).RowHeight, "0.00"), ",", ".")
        If i < ws.UsedRange.Rows.Count Then json = json & ","
        json = json & vbCrLf
    Next i
    json = json & "    }" & vbCrLf
    json = json & "  }," & vbCrLf
    json = json & "  ""cells"": [" & vbCrLf
    
    ts.Write json
    json = ""
    
    ' Считывание 100% свойств ячеек
    Dim isFirstCell As Boolean
    isFirstCell = True
    
    For Each cell In ws.UsedRange
        ' Пропускаем пустые неформатированные ячейки для оптимизации объема данных
        If cell.Text = "" And cell.Interior.ColorIndex = xlNone And _
           cell.Borders(xlEdgeLeft).LineStyle = xlNone And cell.Borders(xlEdgeRight).LineStyle = xlNone And _
           cell.Borders(xlEdgeTop).LineStyle = xlNone And cell.Borders(xlEdgeBottom).LineStyle = xlNone Then
            GoTo NextCell
        End If
        
        If Not isFirstCell Then ts.Write "," & vbCrLf
        isFirstCell = False
        
        ' Экранирование кавычек и спецсимволов в тексте и формулах
        Dim cleanText As String, cleanFormula As String
        cleanText = Replace(cell.Text, """", "\""")
        cleanText = Replace(cleanText, vbCrLf, "\n")
        cleanFormula = Replace(cell.Formula, """", "\""")
        
        cellJson = "    {" & _
          """addr"":""" & cell.Address(False, False) & """," & _
          """v"":""" & cleanText & """," & _
          """f"":""" & IIf(cell.HasFormula, cleanFormula, "") & """," & _
          """bg"":" & cell.Interior.Color & "," & _
          """fc"":" & cell.Font.Color & "," & _
          """sz"":" & Replace(CStr(cell.Font.Size), ",", ".") & "," & _
          """bold"":" & IIf(cell.Font.Bold, "true", "false") & "," & _
          """italic"":" & IIf(cell.Font.Italic, "true", "false") & "," & _
          """ha"":" & cell.HorizontalAlignment & "," & _
          """wrap"":" & IIf(cell.WrapText, "true", "false") & "," & _
          """merged"":" & IIf(cell.MergeCells, "true", "false") & "," & _
          """ma"":""" & IIf(cell.MergeCells, cell.MergeArea.Address(False, False), "") & """," & _
          """b"":[" & _
            IIf(cell.Borders(xlEdgeLeft).LineStyle <> xlNone, "1", "0") & "," & _
            IIf(cell.Borders(xlEdgeRight).LineStyle <> xlNone, "1", "0") & "," & _
            IIf(cell.Borders(xlEdgeTop).LineStyle <> xlNone, "1", "0") & "," & _
            IIf(cell.Borders(xlEdgeBottom).LineStyle <> xlNone, "1", "0") & "]}"
            
        ts.Write cellJson
NextCell:
    Next cell
    
    ts.Write vbCrLf & "  ]" & vbCrLf & "}"
    ts.Close
    
    MsgBox "Данные листа со всеми параметрами оформления успешно сохранены в JSON на Рабочий стол!", vbInformation, "Экспорт завершен"
End Sub
