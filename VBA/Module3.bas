Attribute VB_Name = "Module3"
Sub �����_��������()
    On Error Resume Next
        ThisWorkbook.Sheets("�����").ShowAllData
        ThisWorkbook.Sheets("������� � �������").ShowAllData
    On Error GoTo 0
End Sub
Sub ��������������()
    Dim lastRow As Long
    Dim clipboard As MSForms.DataObject
    Dim clipboardText As String
    ' ������� ��������� �������
    Set clipboard = New MSForms.DataObject
    ' �������� ������ �� ������
    clipboard.GetFromClipboard
    ' ��������� ������ �� ������
    clipboardText = clipboard.GetText
    ' ���������� ���� "����� ASK"
    Set ws = ThisWorkbook.Worksheets("�����")
    ' ���������� �������� ��� ����������
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set filterRange = ws.Range("A4:A" & lastRow)
    ' ��������� ���� �� ������� �� ������ � ��������
    filterRange.AutoFilter Field:=1, Criteria1:=clipboardText, Operator:=xlFilterValues
'���������� �������� ������ Tools-References - Microsoft Forms 2.0
' ���� ������ � ������ ���, �� ����� ������� C:\Windows\System32\FM20.dll
End Sub

