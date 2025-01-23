Attribute VB_Name = "Module3"
Sub Сброс_фильтров()
    On Error Resume Next
        ThisWorkbook.Sheets("Джобы").ShowAllData
        ThisWorkbook.Sheets("Проходы в системы").ShowAllData
    On Error GoTo 0
End Sub
Sub ФильтрПоБуферу()
    Dim lastRow As Long
    Dim clipboard As MSForms.DataObject
    Dim clipboardText As String
    ' Создаем экземпляр объекта
    Set clipboard = New MSForms.DataObject
    ' Получаем данные из буфера
    clipboard.GetFromClipboard
    ' Получение текста из буфера
    clipboardText = clipboard.GetText
    ' Определяем лист "Джобы ASK"
    Set ws = ThisWorkbook.Worksheets("Джобы")
    ' Определяем диапазон для фильтрации
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set filterRange = ws.Range("A4:A" & lastRow)
    ' Фильтруем лист по массиву из данных о проходах
    filterRange.AutoFilter Field:=1, Criteria1:=clipboardText, Operator:=xlFilterValues
'необходимо включить модуль Tools-References - Microsoft Forms 2.0
' если модуля в списке нет, то найти вручную C:\Windows\System32\FM20.dll
End Sub

