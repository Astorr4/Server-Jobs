Attribute VB_Name = "Module2"
Sub StatusSelectJob()
    Dim xmlDoc As MSXML2.DOMDocument60
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim httpReq As MSXML2.XMLHTTP60
    Dim responseXML As String
    'Объявляем переменную, которая сохраняет номер активной строки
    Dim cellContent As String
    cellContent = ActiveCell.Row
    ' если джоб из AlwaysUp
    If Cells(cellContent, "F") = "AlwaysUp" Then
        ' Создаем объект XML-документа
        Set xmlDoc = New MSXML2.DOMDocument60
        ' Создаем объект HTTP-запроса
        Set httpReq = New MSXML2.XMLHTTP60
        ' Выполняем HTTP-запрос
        httpReq.Open "GET", "http://" & Cells(cellContent, "C") & ":8585/api/get-status?password=21232f2&application=" & Cells(cellContent, "E"), True
        httpReq.send
        ' Ожидание ответа от запроса
        Do While httpReq.readyState <> 4
            DoEvents
        Loop
        ' Получаем ответ в виде XML-строки
        responseXML = httpReq.responseText
        ' Загружаем XML-ответ в объект xmlDoc
        xmlDoc.LoadXML (responseXML)
        ' Находим узел с тегом "state"
        Set xmlNode = xmlDoc.SelectSingleNode("//state")
        ' Получаем значение тега "state"
        Status = xmlNode.Text
        ' Если статус джоба Waiting, то меняется цвет заливки и выводится сообщение
        If Status = "Waiting" Then
            Range("A" & cellContent).Interior.Color = RGB(255, 255, 102)
            MsgBox "Статус джоба " & Cells(cellContent, "A") & " - Waiting. " & vbNewLine & "Ячейка окрашена в жёлтый цвет."
        ' Если статус джоба Stopped, то меняется цвет заливки и выводится сообщение
        ElseIf Status = "Stopped" Then
            Range("A" & cellContent).Interior.Color = RGB(255, 80, 80)
            MsgBox "Статус джоба " & Cells(cellContent, "A") & " - Stopped. " & vbNewLine & "Ячейка окрашена в красный цвет."
        ' Если статус джоба Running, значит джоб запущен
        Else
            Range("A" & cellContent).Interior.Color = RGB(0, 176, 80)
            MsgBox "Статус джоба " & Cells(cellContent, "A") & " - Running. " & vbNewLine & "Джоб работает."
        End If
    ' Если выбран джоб из Scheduler
    Else
        MsgBox "Вместо AlwaysUp выбран джоб из Scheduler, выберите джоб из AlwaysUp"
    End If
End Sub
Sub StatusAllJobs()
    Dim xmlDoc As MSXML2.DOMDocument60
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim httpReq As MSXML2.XMLHTTP60
    Dim responseXML As String
    Dim ws As Worksheet
    Dim lastRow As Long
    'Выбираем активный лист
    Set ws = ActiveSheet
    'Определяем последнюю заполненную строку
    lastRow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
    MsgBox "Состояние джобов обновляется"
    For i = 5 To lastRow
        ' если джоб из AlwaysUp
    If Cells(i, "F") = "AlwaysUp" Then
        ' Создаем объект XML-документа
        Set xmlDoc = New MSXML2.DOMDocument60
        ' Создаем объект HTTP-запроса
        Set httpReq = New MSXML2.XMLHTTP60
        ' Выполняем HTTP-запрос
        httpReq.Open "GET", "http://" & Cells(i, "C") & ":8585/api/get-status?password=21232f2&application=" & Cells(i, "E"), True
        httpReq.send
        ' Ожидание ответа от запроса
        Do While httpReq.readyState <> 4
            DoEvents
        Loop
        ' Получаем ответ в виде XML-строки
        responseXML = httpReq.responseText
        ' Загружаем XML-ответ в объект xmlDoc
        xmlDoc.LoadXML (responseXML)
        ' Находим узел с тегом "state"
        Set xmlNode = xmlDoc.SelectSingleNode("//state")
        ' Получаем значение тега "state"
        Status = xmlNode.Text
        ' Если статус джоба Waiting, то меняется цвет заливки и выводится сообщение
        If Status = "Waiting" Then
            Range("A" & i).Interior.Color = RGB(255, 255, 102)
        ' Если статус джоба Stopped, то меняется цвет заливки и выводится сообщение
        ElseIf Status = "Stopped" Then
            Range("A" & i).Interior.Color = RGB(255, 80, 80)
        Else
            Range("A" & i).Interior.Color = RGB(0, 176, 80)
        End If
    End If
    Next i
    MsgBox "Состояние джобов обновлено"
End Sub

