Attribute VB_Name = "Module1"
Sub Start_jobs()
    'Объявляем переменную, которая сохраняет номер активной строки
    Dim cellContent As String
    Dim selectedRange As Range
    Dim cell As Range
    ' Получаем выделенную область
    Set selectedRange = Selection
    ' Проверяем количество выделенных ячеек. Если выделена только одна ячейка, обрабатываем 1 джоб
    If selectedRange.Cells.Count = 1 Then
        ' присваевается переменной строка с активной ячейчкой
        cellContent = ActiveCell.Row
        'Если поле F = AlwaysUp
        If Cells(cellContent, "F") = "AlwaysUp" Then
            'Формируем ссылку к джобу
            link = "http://" & Cells(cellContent, "C") & ":8585/api/start?password=21232f29&application=" & Cells(cellContent, "E")
            ' Подтверждающее запуск окно
            If MsgBox("Будет запущен джоб AlwaysUp " & Cells(cellContent, "E") & " на сервере " & Cells(cellContent, "C"), vbYesNo) = vbYes Then
                'Переходим по ссылке, вывод сообщения
                ThisWorkbook.FollowHyperlink Address:=link
                MsgBox "Джоб " & Cells(cellContent, "E") & " запускается"
            Else
                'Сообщение об отмене нажатия кнопки
                MsgBox "Запуск - " & Cells(cellContent, "E") & " отменен"
            End If
        Else
            ' Подтверждающее запуск окно
            If MsgBox("Будет запущен джоб Scheduler " & Cells(cellContent, "E") & " на сервере " & Cells(cellContent, "C"), vbYesNo) = vbYes Then
                'Формируется и запускается команда cmd
                Path = "C:\PsTools\PsExec.exe \\" & Cells(cellContent, "C") & " -u gpb\askloader -p Ufpghjv,fyr!234 schtasks /run /tn ""\START\" & Cells(cellContent, "E") & """"
                Shell (Path)
                MsgBox "Джоб " & Cells(cellContent, "E") & " запускается"
            Else
                'Сообщение об отмене нажатия кнопки
                MsgBox "Запуск - " & Cells(cellContent, "E") & " отменен"
            End If
        End If
    ' Если выбрано несколько ячеек
    Else
        ' Предупреждение о запуске джобов
        If MsgBox("Будут запущены выделенные джобы. Запустить? ", vbYesNo) = vbYes Then
            ' Если выделено больше одной ячейки, запускаем цикл по значениям
            For Each cell In selectedRange
                ' присваевается переменной строка с активной ячейчкой
                cellContent = cell.Row
                'Если поле F = AlwaysUp
                If Cells(cellContent, "F") = "AlwaysUp" Then
                    'Формируем ссылку к джобу
                    link = "http://" & Cells(cellContent, "C") & ":8585/api/start?password=21232f&application=" & Cells(cellContent, "E")
                    'Переходим по ссылке, вывод сообщения
                    ThisWorkbook.FollowHyperlink Address:=link
                ' Если тип запуска Scheduler
                Else
                    'Формируется и запускается команда cmd
                    Path = "C:\PsTools\PsExec.exe \\" & Cells(cellContent, "C") & " -u domain\user -p password schtasks /run /tn ""\START\" & Cells(cellContent, "E") & """"
                    Shell (Path)
                End If
            Next cell
        Else
            'Сообщение об отмене нажатия кнопки
            MsgBox "Запуск выделенных джобов отменен"
        End If
    End If
End Sub
Sub Stop_jobs()
    'Объявляем переменную, которая сохраняет номер активной строки
    Dim cellContent As String
    Dim selectedRange As Range
    Dim cell As Range
    ' Получаем выделенную область
    Set selectedRange = Selection
    ' Проверяем количество выделенных ячеек. Если выделена только одна ячейка, обрабатываем 1 джоб
    If selectedRange.Cells.Count = 1 Then
        ' присваевается переменной строка с активной ячейчкой
        cellContent = ActiveCell.Row
        'Если поле F = AlwaysUp
        If Cells(cellContent, "F") = "AlwaysUp" Then
            'Формируем ссылку к джобу
            link = "http://" & Cells(cellContent, "C") & ":8585/api/stop?password=21232f2&application=" & Cells(cellContent, "E")
            'Подтверждающее запуск окно
            If MsgBox("Будет остановлен джоб AlwaysUp " & Cells(cellContent, "E") & " на сервере " & Cells(cellContent, "C"), vbYesNo) = vbYes Then
                'Переходим по ссылке, вывод сообщения
                ThisWorkbook.FollowHyperlink Address:=link
                MsgBox "Джоб " & Cells(cellContent, "E") & " останавливается"
            Else
                'Сообщение об отмене нажатия кнопки
                MsgBox "Остановка - " & Cells(cellContent, "E") & " отменена"
            End If
        Else
            ' Подтверждающее запуск окно
            If MsgBox("Будет остановлен джоб Scheduler " & Cells(cellContent, "E") & " на сервере " & Cells(cellContent, "C"), vbYesNo) = vbYes Then
                'Формируется и запускается команда cmd
                Path = "C:\PsTools\PsExec.exe \\" & Cells(cellContent, "C") & " -u domain\user -p password schtasks /end /tn ""\START\" & Cells(cellContent, "E") & """"
                Shell (Path)
                MsgBox "Джоб " & Cells(cellContent, "E") & " останавливается"
            Else
                'Сообщение об отмене нажатия кнопки
                MsgBox "Остановка - " & Cells(cellContent, "E") & " отменена"
            End If
        End If
    ' Если выбрано несколько ячеек
    Else
        ' Предупреждение о запуске джобов
        If MsgBox("Будут остановлены выделенные джобы. Запустить? ", vbYesNo) = vbYes Then
            ' Если выделено больше одной ячейки, запускаем цикл по значениям
            For Each cell In selectedRange
                ' присваевается переменной строка с активной ячейчкой
                cellContent = cell.Row
                'Если поле F = AlwaysUp
                If Cells(cellContent, "F") = "AlwaysUp" Then
                    'Формируем ссылку к джобу
                    link = "http://" & Cells(cellContent, "C") & ":8585/api/stop?password=21232f23&application=" & Cells(cellContent, "E")
                    'Переходим по ссылке, вывод сообщения
                    ThisWorkbook.FollowHyperlink Address:=link
                ' Если тип запуска Scheduler
                Else
                    'Формируется и запускается команда cmd
                    Path = "C:\PsTools\PsExec.exe \\" & Cells(cellContent, "C") & " -u domain\user -p password schtasks /end /tn ""\START\" & Cells(cellContent, "E") & """"
                    Shell (Path)
                End If
            Next cell
        Else
            'Сообщение об отмене нажатия кнопки
            MsgBox "Остановка выделенных джобов отменена"
        End If
    End If
End Sub
Sub Restart_job()
    'Объявляем переменную, которая сохраняет номер активной строки
    Dim cellContent As String
    cellContent = ActiveCell.Row
    'Если поле F = AlwaysUp
    If Cells(cellContent, "F") = "AlwaysUp" Then
        ' Подтверждающее запуск окно
        If MsgBox("Будет перезапущен джоб AlwaysUp " & Cells(cellContent, "E") & " на сервере " & Cells(cellContent, "C"), vbYesNo) = vbYes Then
            'Формируем ссылку к джобу
            link = "http://" & Cells(cellContent, "C") & ":8585/api/restart?password=21232f29&application=" & Cells(cellContent, "E")
            'Переходим по ссылке, вывод сообщения
            ThisWorkbook.FollowHyperlink Address:=link
            MsgBox "Джоб " & Cells(cellContent, "E") & " Перезапускается"
        Else
            'Сообщение об отмене нажатия кнопки
            MsgBox "Перезапуск - " & Cells(cellContent, "E") & " отменен"
        End If
    Else
        ' Подтверждающее перезапуск окно
        If MsgBox("Будет перезапущен джоб Scheduler " & Cells(cellContent, "E") & " на сервере " & Cells(cellContent, "C"), vbYesNo) = vbYes Then
            'Формируется и запускается команда остановки джоба в cmd
            Path_end = "C:\PsTools\PsExec.exe \\" & Cells(cellContent, "C") & " -u domain\user -p password schtasks /end /tn ""\START\" & Cells(cellContent, "E") & """"
            Shell (Path_end)
            MsgBox "Джоб " & Cells(cellContent, "E") & " останавливается"
            Application.Wait Now + #12:00:35 AM#
            'Формируется и запускается команда старта джоба в cmd
            Path_start = "C:\PsTools\PsExec.exe \\" & Cells(cellContent, "C") & " -u domain\user -p password schtasks /run /tn ""\START\" & Cells(cellContent, "E") & """"
            Shell (Path_start)
            MsgBox "Джоб " & Cells(cellContent, "E") & " запускается"
        Else
            'Сообщение об отмене нажатия кнопки
            MsgBox "Перезапуск - " & Cells(cellContent, "E") & " отменен"
        End If
    End If
End Sub



