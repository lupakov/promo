Attribute VB_Name = "mdl_upload"
Sub checkActionData()
    Dim lastRow As Long
    Dim wsheet As Worksheet
    Dim j As clsJournal
    If getRowNumbers() < 1 Then
    
        MsgBox "Нет данных на листе редактор", vbInformation, "Информация"
        Exit Sub
    End If
    If getLastRow() > 10000 Then
        MsgBox "Количество акций на листе превышает 10000", vbCritical, "Информация"
        Exit Sub
    End If
    
    inputResut = MsgBox("ВНИМАНИЕ: Ошибки будут выделены цветом на вкладке ""Редактор"".", vbOKCancel, "Выполнение первичной проверки")
    If inputResut = vbCancel Then
        Exit Sub
    End If
    Set wsheet = getWorkSheet()
    cleanErrorMessage wsheet
    lastRow = getLastRow()
    wsheet.Cells.Validation.Delete
    updateValidation getColumns(), wsheet, 6, lastRow
    If Not firstCheckValues() Then
        MsgBox "Критическая ошибка", vbCritical, "Проверка"
        Exit Sub
    End If

    Set j = New clsJournal
    j.loadJournalFromSheet
    If j.checkDataset() Then
        MsgBox "Проверка пройдена успешно", vbInformation, "Модуль валидации"
    End If
End Sub


Function getRowNumbers() As Long
    Dim wsh As Worksheet
    Dim lastRow As Long
    Dim firstRowData As Long
    
    firstRowData = 6
    Set wsh = getWorkSheet
    lastRow = getLastRow()
    getRowNumbers = lastRow - firstRowData
End Function

Sub checkAndRunActionData()
    Dim wsheet As Worksheet
    Dim lastRow As Long
    Set wsheet = getWorkSheet()
    lastRow = getLastRow()
    If getRowNumbers() < 1 Then
    
        MsgBox "Нет данных на листе редактор", vbInformation, "Информация"
        Exit Sub
    End If
    
    If getLastRow() > 10000 Then
        MsgBox "Количество акций на листе превышает 10000", vbCritical, "Информация"
        Exit Sub
    End If
    
    inputResut = MsgBox("ВНИМАНИЕ: Предварительно будут запущены проверки в последовательности: первичная проверка, проверка крит.ЦО.", vbOKCancel, "Выполнение расчета цен")
    If inputResut = vbCancel Then
        Exit Sub
    End If
    cleanErrorMessage wsheet
    wsheet.Cells.Validation.Delete
    updateValidation getColumns(), wsheet, 6, lastRow
    If Not firstCheckValues() Then
        MsgBox "Критическая ошибка", vbCritical, "Проверка"
        Exit Sub
    End If

    Set j = New clsJournal
    j.loadJournalFromSheet
    If Not j.checkDataset() Then
        Exit Sub
    End If
    j.saveToPersistJournal
    cleaning
End Sub

Sub acceptActionData()
    If getRowNumbers() < 1 Then
    
        MsgBox "Нет данных на листе редактор", vbInformation, "Информация"
        Exit Sub
    End If
    
    inputResut = MsgBox("ВНИМАНИЕ: Согласованы будут только связки со статусом ""Расчитана"".", vbOKCancel, "Выполнение согласования КМ")
    If inputResut = vbCancel Then
        Exit Sub
    End If
    
    Dim j As clsJournal
    Set j = New clsJournal
    j.loadAcceptFromSheet
    If j.acceptActions Then
        cleaning
    End If
End Sub

Sub cancelActionData()
    If getRowNumbers() < 1 Then
    
        MsgBox "Нет данных на листе редактор", vbInformation, "Информация"
        Exit Sub
    End If
    
    inputResut = MsgBox("ВНИМАНИЕ: будет произведено удаление из базы всех связок, находящихся на вкладке ""Редактор"".", vbOKCancel, "Выполнение удаления данных")
    If inputResut = vbCancel Then
        Exit Sub
    End If
    
        Dim j As clsJournal
    Set j = New clsJournal
    j.loadCancelFromSheet
    j.cancelActions
   cleaning
End Sub

