Attribute VB_Name = "Module1"
Sub Start_jobs()
    '��������� ����������, ������� ��������� ����� �������� ������
    Dim cellContent As String
    Dim selectedRange As Range
    Dim cell As Range
    ' �������� ���������� �������
    Set selectedRange = Selection
    ' ��������� ���������� ���������� �����. ���� �������� ������ ���� ������, ������������ 1 ����
    If selectedRange.Cells.Count = 1 Then
        ' ������������� ���������� ������ � �������� ��������
        cellContent = ActiveCell.Row
        '���� ���� F = AlwaysUp
        If Cells(cellContent, "F") = "AlwaysUp" Then
            '��������� ������ � �����
            link = "http://" & Cells(cellContent, "C") & ":8585/api/start?password=21232f29&application=" & Cells(cellContent, "E")
            ' �������������� ������ ����
            If MsgBox("����� ������� ���� AlwaysUp " & Cells(cellContent, "E") & " �� ������� " & Cells(cellContent, "C"), vbYesNo) = vbYes Then
                '��������� �� ������, ����� ���������
                ThisWorkbook.FollowHyperlink Address:=link
                MsgBox "���� " & Cells(cellContent, "E") & " �����������"
            Else
                '��������� �� ������ ������� ������
                MsgBox "������ - " & Cells(cellContent, "E") & " �������"
            End If
        Else
            ' �������������� ������ ����
            If MsgBox("����� ������� ���� Scheduler " & Cells(cellContent, "E") & " �� ������� " & Cells(cellContent, "C"), vbYesNo) = vbYes Then
                '����������� � ����������� ������� cmd
                Path = "C:\PsTools\PsExec.exe \\" & Cells(cellContent, "C") & " -u gpb\askloader -p Ufpghjv,fyr!234 schtasks /run /tn ""\START\" & Cells(cellContent, "E") & """"
                Shell (Path)
                MsgBox "���� " & Cells(cellContent, "E") & " �����������"
            Else
                '��������� �� ������ ������� ������
                MsgBox "������ - " & Cells(cellContent, "E") & " �������"
            End If
        End If
    ' ���� ������� ��������� �����
    Else
        ' �������������� � ������� ������
        If MsgBox("����� �������� ���������� �����. ���������? ", vbYesNo) = vbYes Then
            ' ���� �������� ������ ����� ������, ��������� ���� �� ���������
            For Each cell In selectedRange
                ' ������������� ���������� ������ � �������� ��������
                cellContent = cell.Row
                '���� ���� F = AlwaysUp
                If Cells(cellContent, "F") = "AlwaysUp" Then
                    '��������� ������ � �����
                    link = "http://" & Cells(cellContent, "C") & ":8585/api/start?password=21232f&application=" & Cells(cellContent, "E")
                    '��������� �� ������, ����� ���������
                    ThisWorkbook.FollowHyperlink Address:=link
                ' ���� ��� ������� Scheduler
                Else
                    '����������� � ����������� ������� cmd
                    Path = "C:\PsTools\PsExec.exe \\" & Cells(cellContent, "C") & " -u domain\user -p password schtasks /run /tn ""\START\" & Cells(cellContent, "E") & """"
                    Shell (Path)
                End If
            Next cell
        Else
            '��������� �� ������ ������� ������
            MsgBox "������ ���������� ������ �������"
        End If
    End If
End Sub
Sub Stop_jobs()
    '��������� ����������, ������� ��������� ����� �������� ������
    Dim cellContent As String
    Dim selectedRange As Range
    Dim cell As Range
    ' �������� ���������� �������
    Set selectedRange = Selection
    ' ��������� ���������� ���������� �����. ���� �������� ������ ���� ������, ������������ 1 ����
    If selectedRange.Cells.Count = 1 Then
        ' ������������� ���������� ������ � �������� ��������
        cellContent = ActiveCell.Row
        '���� ���� F = AlwaysUp
        If Cells(cellContent, "F") = "AlwaysUp" Then
            '��������� ������ � �����
            link = "http://" & Cells(cellContent, "C") & ":8585/api/stop?password=21232f2&application=" & Cells(cellContent, "E")
            '�������������� ������ ����
            If MsgBox("����� ���������� ���� AlwaysUp " & Cells(cellContent, "E") & " �� ������� " & Cells(cellContent, "C"), vbYesNo) = vbYes Then
                '��������� �� ������, ����� ���������
                ThisWorkbook.FollowHyperlink Address:=link
                MsgBox "���� " & Cells(cellContent, "E") & " ���������������"
            Else
                '��������� �� ������ ������� ������
                MsgBox "��������� - " & Cells(cellContent, "E") & " ��������"
            End If
        Else
            ' �������������� ������ ����
            If MsgBox("����� ���������� ���� Scheduler " & Cells(cellContent, "E") & " �� ������� " & Cells(cellContent, "C"), vbYesNo) = vbYes Then
                '����������� � ����������� ������� cmd
                Path = "C:\PsTools\PsExec.exe \\" & Cells(cellContent, "C") & " -u domain\user -p password schtasks /end /tn ""\START\" & Cells(cellContent, "E") & """"
                Shell (Path)
                MsgBox "���� " & Cells(cellContent, "E") & " ���������������"
            Else
                '��������� �� ������ ������� ������
                MsgBox "��������� - " & Cells(cellContent, "E") & " ��������"
            End If
        End If
    ' ���� ������� ��������� �����
    Else
        ' �������������� � ������� ������
        If MsgBox("����� ����������� ���������� �����. ���������? ", vbYesNo) = vbYes Then
            ' ���� �������� ������ ����� ������, ��������� ���� �� ���������
            For Each cell In selectedRange
                ' ������������� ���������� ������ � �������� ��������
                cellContent = cell.Row
                '���� ���� F = AlwaysUp
                If Cells(cellContent, "F") = "AlwaysUp" Then
                    '��������� ������ � �����
                    link = "http://" & Cells(cellContent, "C") & ":8585/api/stop?password=21232f23&application=" & Cells(cellContent, "E")
                    '��������� �� ������, ����� ���������
                    ThisWorkbook.FollowHyperlink Address:=link
                ' ���� ��� ������� Scheduler
                Else
                    '����������� � ����������� ������� cmd
                    Path = "C:\PsTools\PsExec.exe \\" & Cells(cellContent, "C") & " -u domain\user -p password schtasks /end /tn ""\START\" & Cells(cellContent, "E") & """"
                    Shell (Path)
                End If
            Next cell
        Else
            '��������� �� ������ ������� ������
            MsgBox "��������� ���������� ������ ��������"
        End If
    End If
End Sub
Sub Restart_job()
    '��������� ����������, ������� ��������� ����� �������� ������
    Dim cellContent As String
    cellContent = ActiveCell.Row
    '���� ���� F = AlwaysUp
    If Cells(cellContent, "F") = "AlwaysUp" Then
        ' �������������� ������ ����
        If MsgBox("����� ����������� ���� AlwaysUp " & Cells(cellContent, "E") & " �� ������� " & Cells(cellContent, "C"), vbYesNo) = vbYes Then
            '��������� ������ � �����
            link = "http://" & Cells(cellContent, "C") & ":8585/api/restart?password=21232f29&application=" & Cells(cellContent, "E")
            '��������� �� ������, ����� ���������
            ThisWorkbook.FollowHyperlink Address:=link
            MsgBox "���� " & Cells(cellContent, "E") & " ���������������"
        Else
            '��������� �� ������ ������� ������
            MsgBox "���������� - " & Cells(cellContent, "E") & " �������"
        End If
    Else
        ' �������������� ���������� ����
        If MsgBox("����� ����������� ���� Scheduler " & Cells(cellContent, "E") & " �� ������� " & Cells(cellContent, "C"), vbYesNo) = vbYes Then
            '����������� � ����������� ������� ��������� ����� � cmd
            Path_end = "C:\PsTools\PsExec.exe \\" & Cells(cellContent, "C") & " -u domain\user -p password schtasks /end /tn ""\START\" & Cells(cellContent, "E") & """"
            Shell (Path_end)
            MsgBox "���� " & Cells(cellContent, "E") & " ���������������"
            Application.Wait Now + #12:00:35 AM#
            '����������� � ����������� ������� ������ ����� � cmd
            Path_start = "C:\PsTools\PsExec.exe \\" & Cells(cellContent, "C") & " -u domain\user -p password schtasks /run /tn ""\START\" & Cells(cellContent, "E") & """"
            Shell (Path_start)
            MsgBox "���� " & Cells(cellContent, "E") & " �����������"
        Else
            '��������� �� ������ ������� ������
            MsgBox "���������� - " & Cells(cellContent, "E") & " �������"
        End If
    End If
End Sub



