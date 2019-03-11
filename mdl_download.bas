Attribute VB_Name = "mdl_download"
Function getResultRs(con As ADODB.Connection, userId As Long, Optional weekId As Long = 0) As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim query As String
If weekId = 0 Then
    query = "select * from pricing_sale.PROMO.GET_RESULT (" & CStr(userId) & ")"
Else
    query = "select * from pricing_sale.PROMO.GET_RESULT_week (" & CStr(userId) & ", " & CStr(weekId) & ")"
End If
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.CursorLocation = adUseClientBatch
Set rs.ActiveConnection = con
rs.Open (query)
Set getResultRs = rs
End Function

Sub writeResult(con As ADODB.Connection, userId As Long, Optional weekId As Long = 0, Optional sheetName As String = "Редактор", Optional address As String = "A7")
    Dim wsheet As Worksheet
    Dim cpRange As Range
    Dim rs As ADODB.Recordset
    Dim qt As QueryTable
Dim query As String
If weekId = 0 Then
    query = "select * from pricing_sale.PROMO.GET_RESULT (" & CStr(userId) & ")"
Else
    query = "select * from pricing_sale.PROMO.GET_RESULT_week (" & CStr(userId) & ", " & CStr(weekId) & ")"
End If
    Set wsheet = ThisWorkbook.Worksheets(sheetName)
    clearWorkSheet wsheet
    Set cpRange = wsheet.Range(address)
   If weekId = 0 Then
    query = "select * from pricing_sale.PROMO.GET_RESULT (" & CStr(userId) & ")"
Else
    query = "select * from pricing_sale.PROMO.GET_RESULT_week (" & CStr(userId) & ", " & CStr(weekId) & ")"
End If

    
    connstring = "OLEDB;Provider=SQLOLEDB; Server=pricecraft;Database=PRICING_SALE; Trusted_Connection=yes"
    Set qt = wsheet.QueryTables.Add(Connection:=connstring, Destination:=cpRange, Sql:=query)
    qt.BackgroundQuery = False
    qt.FieldNames = False
    qt.AdjustColumnWidth = False

        qt.Refresh

    
    


    qt.Delete
    writeHeaderLine getColumns(), wsheet, 6, getLastRow()
   '  createPivotTable wsheet
 

End Sub

Sub loadWeek()

    Dim week As Long
    Dim userId As Long
    Dim con As ADODB.Connection
    Dim inputResut As Variant
    Set con = getAdoCon
    userId = getUserId(con)
    inputResut = InputBox("ВНИМАНИЕ: Будет очищена вкладка ""Редактор"". Убедитесь, что все необходимые изменения отправлены на расчет. Неделя в формате ""ГГГГНН"": ", "Выполнение недельной выгрузки")
    If CStr(inputResut) Like "" Then
        Exit Sub
    End If
    week = CLng(inputResut)
    logPressButton con, 1, week
    writeResult con, userId, week
    closeAdoCon con
    
End Sub
Sub loadAll()
    Dim week As Long
    Dim userId As Long
    Dim con As ADODB.Connection
    Dim inputResut As Variant
    Set con = getAdoCon
    userId = getUserId(con)
    inputResut = MsgBox("ВНИМАНИЕ: Будет очищена вкладка ""Редактор"". Убедитесь, что все необходимые изменения отправлены на расчет.", vbOKCancel, "Выполнение общей выгрузки")
    If inputResut = vbCancel Then
        Exit Sub
    End If
    logPressButton con, 2
    writeResult con, userId
    
End Sub

Function getUserId(con As ADODB.Connection) As Long
    Dim rs As ADODB.Recordset
    Dim rslt As Long
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    rs.Open ("select ID from PRICING_SALE.PROMO.USERS WHERE USER_LOGIN = " & quotedString(Environ("username")))
    rslt = rs.Fields("ID").value
    getUserId = rslt
End Function


Sub downloadActionToSvema()

    If getRowNumbers() < 1 Then
    
        MsgBox "Нет данных на листе редактор", vbInformation, "Информация"
        Exit Sub
    End If

    inputResut = InputBox("ВНИМАНИЕ: Шаблон будет формироваться только на основании данных на вкладке ""Редактор"" и только по статусам ""Согласовано РК"". Убедитесь, что все необходимые изменения отправлены на расчет. Неделя: ", "Выполнение недельной выгрузки")
    If CStr(inputResut) Like "" Then
        Exit Sub
    End If
End Sub
