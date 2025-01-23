Attribute VB_Name = "Module2"
Sub StatusSelectJob()
    Dim xmlDoc As MSXML2.DOMDocument60
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim httpReq As MSXML2.XMLHTTP60
    Dim responseXML As String
    '��������� ����������, ������� ��������� ����� �������� ������
    Dim cellContent As String
    cellContent = ActiveCell.Row
    ' ���� ���� �� AlwaysUp
    If Cells(cellContent, "F") = "AlwaysUp" Then
        ' ������� ������ XML-���������
        Set xmlDoc = New MSXML2.DOMDocument60
        ' ������� ������ HTTP-�������
        Set httpReq = New MSXML2.XMLHTTP60
        ' ��������� HTTP-������
        httpReq.Open "GET", "http://" & Cells(cellContent, "C") & ":8585/api/get-status?password=21232f2&application=" & Cells(cellContent, "E"), True
        httpReq.send
        ' �������� ������ �� �������
        Do While httpReq.readyState <> 4
            DoEvents
        Loop
        ' �������� ����� � ���� XML-������
        responseXML = httpReq.responseText
        ' ��������� XML-����� � ������ xmlDoc
        xmlDoc.LoadXML (responseXML)
        ' ������� ���� � ����� "state"
        Set xmlNode = xmlDoc.SelectSingleNode("//state")
        ' �������� �������� ���� "state"
        Status = xmlNode.Text
        ' ���� ������ ����� Waiting, �� �������� ���� ������� � ��������� ���������
        If Status = "Waiting" Then
            Range("A" & cellContent).Interior.Color = RGB(255, 255, 102)
            MsgBox "������ ����� " & Cells(cellContent, "A") & " - Waiting. " & vbNewLine & "������ �������� � ����� ����."
        ' ���� ������ ����� Stopped, �� �������� ���� ������� � ��������� ���������
        ElseIf Status = "Stopped" Then
            Range("A" & cellContent).Interior.Color = RGB(255, 80, 80)
            MsgBox "������ ����� " & Cells(cellContent, "A") & " - Stopped. " & vbNewLine & "������ �������� � ������� ����."
        ' ���� ������ ����� Running, ������ ���� �������
        Else
            Range("A" & cellContent).Interior.Color = RGB(0, 176, 80)
            MsgBox "������ ����� " & Cells(cellContent, "A") & " - Running. " & vbNewLine & "���� ��������."
        End If
    ' ���� ������ ���� �� Scheduler
    Else
        MsgBox "������ AlwaysUp ������ ���� �� Scheduler, �������� ���� �� AlwaysUp"
    End If
End Sub
Sub StatusAllJobs()
    Dim xmlDoc As MSXML2.DOMDocument60
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim httpReq As MSXML2.XMLHTTP60
    Dim responseXML As String
    Dim ws As Worksheet
    Dim lastRow As Long
    '�������� �������� ����
    Set ws = ActiveSheet
    '���������� ��������� ����������� ������
    lastRow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
    MsgBox "��������� ������ �����������"
    For i = 5 To lastRow
        ' ���� ���� �� AlwaysUp
    If Cells(i, "F") = "AlwaysUp" Then
        ' ������� ������ XML-���������
        Set xmlDoc = New MSXML2.DOMDocument60
        ' ������� ������ HTTP-�������
        Set httpReq = New MSXML2.XMLHTTP60
        ' ��������� HTTP-������
        httpReq.Open "GET", "http://" & Cells(i, "C") & ":8585/api/get-status?password=21232f2&application=" & Cells(i, "E"), True
        httpReq.send
        ' �������� ������ �� �������
        Do While httpReq.readyState <> 4
            DoEvents
        Loop
        ' �������� ����� � ���� XML-������
        responseXML = httpReq.responseText
        ' ��������� XML-����� � ������ xmlDoc
        xmlDoc.LoadXML (responseXML)
        ' ������� ���� � ����� "state"
        Set xmlNode = xmlDoc.SelectSingleNode("//state")
        ' �������� �������� ���� "state"
        Status = xmlNode.Text
        ' ���� ������ ����� Waiting, �� �������� ���� ������� � ��������� ���������
        If Status = "Waiting" Then
            Range("A" & i).Interior.Color = RGB(255, 255, 102)
        ' ���� ������ ����� Stopped, �� �������� ���� ������� � ��������� ���������
        ElseIf Status = "Stopped" Then
            Range("A" & i).Interior.Color = RGB(255, 80, 80)
        Else
            Range("A" & i).Interior.Color = RGB(0, 176, 80)
        End If
    End If
    Next i
    MsgBox "��������� ������ ���������"
End Sub

