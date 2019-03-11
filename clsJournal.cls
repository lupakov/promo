VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private tmpTable As String
Private persTable As String
Private persView As String
Private cancelTable As String
Private acceptTable As String
Private con As ADODB.Connection
Private firstRow As Integer
Private lastRow As Integer
Private actionList As ArrayList
Private wsheetName As String
Private wsheet As Worksheet
Private userId As Long
Private bossId As Long
Private superList As ArrayList
Private statusDict As Dictionary
Private geoDict As Dictionary
Private actTypeDict As Dictionary
Private frmtDict As Dictionary
Private kisTypeDict As Dictionary
Private uniqueList As ArrayList
Private geoUniqueDict As Dictionary
Private lineUniqueDict As Dictionary
Private analogUniqueDict As Dictionary

Private weekNameDict As Dictionary


Private importTable As String





Private Sub Class_Initialize()
    tmpTable = "PRICING_SALE.PROMO.ACTION_LIST_" & Environ("username")
    persTable = "PRICING_SALE.PROMO.SUMMARY"
    persView = "PRICING_SALE.PROMO.V_LAST_SUMMARY"
    cancelTable = "PRICING_SALE.PROMO.CANCEL_LIST_" & Environ("username")
    acceptTable = "PRICING_SALE.PROMO.ACCEPT_LIST_" & Environ("username")
    importTable = "PRICING_SALE.PROMO.T_IMPORT"
    wsheetName = "Редактор"
    Set wsheet = ThisWorkbook.Worksheets(wsheetName)
    firstRow = 7
    lastRow = getLastRow()
    Set con = getAdoCon
    Set actionList = New ArrayList
    Set superList = New ArrayList
    fillUserInfo
    fillGeoDict
    fillFrmtDict
    fillStatusDict
    fillActionTypeDict
    fillKisTypeDict
    fillSuperList
    Set uniqueList = New ArrayList
    Set geoUniqueDict = New Dictionary
    Set lineUniqueDict = New Dictionary
    Set analogUniqueDict = New Dictionary
    
    Set weekNameDict = nameToDict("nmWeeks", 0, 1)

End Sub

Public Sub loadJournalFromSheet()
Dim i As Long
Dim act As clsAction
Dim a As clsAction
Dim rs As ADODB.Recordset
logPressButton con, 4
con.Execute ("EXECUTE PRICING_SALE.PROMO.CLEAR_IMPORT " & CStr(userId))
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockBatchOptimistic
rs.CursorLocation = adUseClientBatch
Set rs.ActiveConnection = con
rs.Open ("select * from " & importTable & " where load_user_id = " & CStr(userId))

For i = firstRow To lastRow

    Set act = New clsAction
    act.setWorkSheet wsheetName
    act.setUserInfo userId, bossId
    act.setActTypeDict actTypeDict
    act.setFrmtDict frmtDict
    act.setGeoDict geoDict
    act.setKisTypeDict kisTypeDict
    act.setStatusDict statusDict
    act.setWeekDict weekNameDict
    act.setSuperList superList
    act.loadAction wsheet.Rows(i), i
    
   ' Application.StatusBar = "Check line " & CStr(i - firstRow + 1) & " of " & CStr(lastRow - firstRow + 1)
    DoEvents
    act.calFields
    
    'act.checkAction con, uniqueList, geoUniqueDict, lineUniqueDict, analogUniqueDict, "pricing_sale.promo.V_LIVE_SUMMARY"
    If act.readyForCalc() Then
        actionList.Add act
        act.upload rs
    End If
Next
rs.UpdateBatch


'Application.StatusBar = False
End Sub

Function checkDataset() As Boolean
Dim rs As ADODB.Recordset
Set rs = con.Execute("select * from pricing_sale.promo.second_check(" & CStr(userId) & ")")
If Not rs.EOF Then
    drawSecondCheck rs
    MsgBox "Критическая ошибка", vbCritical
    checkDataset = False
    Exit Function
End If
con.Execute ("execute pricing_sale.promo.set_art_id " & CStr(userId))
Set rs = con.Execute("select * from pricing_sale.promo.third_check(" & CStr(userId) & ")")
If Not rs.EOF Then
    drawThirdCheck rs
    MsgBox "Критическая ошибка", vbCritical
    checkDataset = False
    Exit Function
End If

Set rs = con.Execute("select * from pricing_sale.promo.fourth_check(" & CStr(userId) & ")")
If Not rs.EOF Then
    drawFourthCheck rs
    MsgBox "Критическая ошибка", vbCritical
    checkDataset = False
    Exit Function
End If
checkDataset = True
End Function

Sub drawSecondCheck(rs As ADODB.Recordset)
    While Not rs.EOF
        wsheet.Cells(rs.Fields("EXCEL_LINE").value, 4).Interior.Color = IIf(rs.Fields("ART").value = 1, vbRed, xlNone)
    wsheet.Cells(rs.Fields("EXCEL_LINE").value, 11).Interior.Color = IIf(rs.Fields("KIS").value = 1, vbRed, xlNone)
    wsheet.Cells(rs.Fields("EXCEL_LINE").value, 20).Interior.Color = IIf(rs.Fields("LINE").value = 1, vbRed, xlNone)
    wsheet.Cells(rs.Fields("EXCEL_LINE").value, 44).Interior.Color = IIf(rs.Fields("CO").value = 1, vbRed, xlNone)
    wsheet.Cells(rs.Fields("EXCEL_LINE").value, 45).Interior.Color = IIf(rs.Fields("FORECAST").value = 1, vbRed, xlNone)
    wsheet.Cells(rs.Fields("EXCEL_LINE").value, 33).Interior.Color = IIf(rs.Fields("DISC").value = 1, vbRed, xlNone)
    
    addErrorMessage wsheet, rs.Fields("EXCEL_LINE").value, IIf(rs.Fields("ART").value = 1, "Не найден код товара; ", "") & _
        IIf(rs.Fields("KIS").value = 1, "Не найден код Поставщика; ", "") & _
        IIf(rs.Fields("LINE").value = 1, "Не найден код линейки; ", "") & _
        IIf(rs.Fields("CO").value = 1, "Не найден код аналога смотки; ", "") & _
        IIf(rs.Fields("FORECAST").value = 1, "Не найден код аналога прогноза; ", "") & _
        IIf(rs.Fields("DISC").value = 1, "Неверная итоговая скидка; ", "")
    rs.MoveNext
    Wend
End Sub
Sub drawThirdCheck(rs As ADODB.Recordset)
    While Not rs.EOF
        wsheet.Cells(rs.Fields("EXCEL_LINE").value, 4).Interior.Color = vbRed
    wsheet.Cells(rs.Fields("EXCEL_LINE").value, 6).Interior.Color = vbRed
    wsheet.Cells(rs.Fields("EXCEL_LINE").value, 7).Interior.Color = vbRed
    wsheet.Cells(rs.Fields("EXCEL_LINE").value, 8).Interior.Color = vbRed
 
   addErrorMessage wsheet, rs.Fields("EXCEL_LINE").value, "Ошибка ключевой связки; "
 
    rs.MoveNext
    
  
    Wend
End Sub

Sub drawFourthCheck(rs As ADODB.Recordset)
    While Not rs.EOF

    wsheet.Cells(rs.Fields("EXCEL_LINE").value, 20).Interior.Color = IIf(rs.Fields("LINE").value = 1, vbRed, xlNone)
    wsheet.Cells(rs.Fields("EXCEL_LINE").value, 44).Interior.Color = IIf(rs.Fields("CO").value = 1, vbRed, xlNone)
    wsheet.Cells(rs.Fields("EXCEL_LINE").value, 45).Interior.Color = IIf(rs.Fields("FORECAST").value = 1, vbRed, xlNone)
    wsheet.Cells(rs.Fields("EXCEL_LINE").value, 7).Interior.Color = IIf(rs.Fields("GEO").value = 1, vbRed, xlNone)
    addErrorMessage wsheet, rs.Fields("EXCEL_LINE").value, _
        IIf(rs.Fields("LINE").value = 1, "Ошибка линейки; ", "") & _
        IIf(rs.Fields("CO").value = 1, "Ошибка аналога смотки; ", "") & _
        IIf(rs.Fields("FORECAST").value = 1, "Ошибка аналога прогноза; ", "") & _
        IIf(rs.Fields("GEO").value = 1, "Ошибка Географии; ", "")
    
    rs.MoveNext
    Wend
End Sub

Sub drawAcceptCheck(rs As ADODB.Recordset)
    While Not rs.EOF

    wsheet.Cells(rs.Fields("EXCEL_LINE").value, 50).Interior.Color = vbRed
    
      addErrorMessage
    
    rs.MoveNext
    Wend
End Sub
Public Sub loadCancelFromSheet()
Dim i As Long
Dim act As clsAction


Dim rs As ADODB.Recordset

con.Execute ("EXECUTE PRICING_SALE.PROMO.CLEAR_IMPORT " & CStr(userId))
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockBatchOptimistic
rs.CursorLocation = adUseClientBatch
Set rs.ActiveConnection = con
rs.Open ("select * from " & importTable & " where load_user_id = " & CStr(userId))

For i = firstRow To lastRow
    If wsheet.Rows(i).RowHeight Then
    Set act = New clsAction
    act.setWorkSheet wsheetName
    act.setUserInfo userId, bossId
    act.setActTypeDict actTypeDict
    act.setStatusDict statusDict
    act.setSuperList superList
    act.loadActionForCancel wsheet.Rows(i), i
        If act.readyForCancel() Then
            actionList.Add act
            act.uploadCancel rs
        End If
    End If
Next
  rs.UpdateBatch
End Sub

Public Sub loadAcceptFromSheet()
Dim i As Long
Dim act As clsAction

Dim rs As ADODB.Recordset

con.Execute ("EXECUTE PRICING_SALE.PROMO.CLEAR_IMPORT " & CStr(userId))
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockBatchOptimistic
rs.CursorLocation = adUseClientBatch
Set rs.ActiveConnection = con
rs.Open ("select * from " & importTable & " where load_user_id = " & CStr(userId))



For i = firstRow To lastRow
    Set act = New clsAction
    act.setWorkSheet wsheetName
    act.setUserInfo userId, bossId
    act.setActTypeDict actTypeDict
    act.setStatusDict statusDict
    act.setSuperList superList
    act.loadActionForAccept wsheet.Rows(i), i
    If act.readyForCmAccept() Then
        actionList.Add act
        act.uploadAccept rs
    End If
Next
  rs.UpdateBatch
  
End Sub

Public Sub saveToPersistJournal()
    Dim rs As ADODB.Recordset
    Dim seriosCount As Long
    logPressButton con, 5
    Set rs = con.Execute("exec pricing_sale.[PROMO].[SAVE_IMPORT] " & CStr(userId))
    seriosCount = CLng(rs.Fields(0).value)
    If seriosCount > 0 Then
        MsgBox "Произведены значительные изменения. Акции будут пересчитаны", vbInformation
    End If
End Sub

Public Sub cancelActions()
logPressButton con, 7
    con.Execute "exec pricing_sale.PROMO.CANCEL " & CStr(userId)
End Sub
Public Function acceptActions() As Boolean
    Dim rs As ADODB.Recordset
    Dim res As Variant
    Dim cmd As ADODB.Command
    Dim inParam As ADODB.Parameter
    Dim outParam As ADODB.Parameter

  '  cmd.Parameters.Append retParam
    Set rs = con.Execute("select * from  pricing_sale.PROMO.CHECK_ACCEPT (" & CStr(userId) & ")")
    If Not rs.EOF Then
        res = rs.Fields("MAX_STATUS").value
    
        drawAcceptCheck rs
        If res > -2 Then
       
            If res = 4 Then
                
                MsgBox "Ошибка целостности информации: внесено более одной уникальной цены", vbCritical, "Ошибка"
            ElseIf res = 5 Then
              
                MsgBox "Ошибка целостности информации: внесенная цена отличается от ранее согласованной КМ", vbCritical, "Ошибка"
            ElseIf res = 6 Then
             
                MsgBox "Ошибка целостности информации: внесенная цена отличается от ранее согласованной РК", vbCritical, "Ошибка"
    
            End If
            acceptActions = False
            Exit Function
        Else
            ans = MsgBox("Найдены акции с ценой выше регулярной. Продолжить?", vbYesNo)
            If ans = vbNo Then
               acceptActions = False
               Exit Function
            Else
                logPressButton con, 6
                con.Execute ("exec pricing_sale.PROMO.accept_cm " & CStr(userId))
               acceptActions = True
        
            End If
        End If
    Else
        logPressButton con, 6
        con.Execute ("exec pricing_sale.PROMO.accept_cm " & CStr(userId))
        acceptActions = True
    End If
    
    
End Function

Private Sub fillUserInfo()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    rs.Open ("select ID, USER_BOSS_ID from pricing_sale.promo.users where user_login = " & quotedString(Environ("username")))
    userId = rs.Fields("ID").value
    bossId = rs.Fields("USER_BOSS_ID").value
End Sub

Private Sub fillGeoDict()
    fillDictFromTable geoDict, "pricing_sale.promo.v_rc", "NAME", "CNTR_ID"
End Sub

Private Sub fillFrmtDict()
    fillDictFromTable frmtDict, "PRICING_SALE.PROMO.FRMT", "NAME", "FRMT_ID"
End Sub

Private Sub fillStatusDict()
fillDictFromTable statusDict, "PRICING_SALE.PROMO.ACTION_STATUS", "STATUS_NAME", "STATUS_ID"
End Sub

Private Sub fillActionTypeDict()
fillDictFromTable actTypeDict, "PRICING_SALE.PROMO.ACTION_TYPE", "ACTION_TYPE_NAME", "ACTION_TYPE_ID"
End Sub
Private Sub fillKisTypeDict()
fillDictFromTable kisTypeDict, "PRICING_SALE.PROMO.KIS_TYPE", "NAME", "KIS_TYPE_ID"
End Sub


Private Sub fillDictFromTable(dict As Dictionary, tableName As String, keyField As String, valueField As String)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    rs.Open ("select " & keyField & ", " & valueField & " from " & tableName)
    Set dict = New Dictionary
    While Not rs.EOF
        dict.Add rs.Fields(keyField).value, rs.Fields(valueField).value
        rs.MoveNext
    Wend
End Sub

Private Sub fillSuperList()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    rs.Open ("select SUPER_TYPE from pricing_sale.promo.user_super where GETDATE() BETWEEN BEGIN_DT AND END_DT AND id_user = " & CStr(userId))
    While Not rs.EOF
        superList.Add (rs.Fields("SUPER_TYPE").value)
    rs.MoveNext
    Wend
End Sub


