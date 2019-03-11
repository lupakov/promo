Attribute VB_Name = "Library_load"
Function prepareTextFileForJDBC(wrange As Range, Optional withMacro As Boolean = False, Optional fileName As String = "tmp_import.csv", Optional visual As Boolean = True, Optional ft As String = "") As String
    On Error GoTo ERRLabel
    Dim tmpwb As Workbook
    Dim filepath As String
    Dim tmpDir As String
    tmpDir = Environ("TEMP")
    filepath = tmpDir & Application.PathSeparator & fileName
    If Not withMacro Then
    Set tmpwb = Workbooks.Add
    wrange.Copy Destination:=tmpwb.Worksheets(1).Range("a1")
    tmpwb.SaveAs fileName:=filepath, _
        FileFormat:=xlText, CreateBackup:=False
    tmpwb.Close (False)
    Else
    save_macro wrange, filepath, ft
    End If
    prepareTextFileForJDBC = filepath
    On Error GoTo 0
    Exit Function
ERRLabel:
MsgBox "Error", vbCritical
End Function
Function save_macro(wrange As Range, failPath As String, Optional ft As String = "") As Integer
Dim data As String
Dim tRow As Range
Dim tCell As Range
Dim beginRow As String
Dim curTf As String

     Open failPath For Output As #1
     For Each tRow In wrange.Rows
               data = ""
               beginRow = ""
        For Each tCell In tRow.Cells
            If ft <> "" Then
            curTf = Mid(ft, tCell.Column, 1)
            Else
            curTf = ""
            End If
             data = data & beginRow & getTextValue(tCell.value, curTf)
             beginRow = Chr(9)
         Next
         Print #1, data
     Next
     Close #1
save_macro = 0
End Function
Function getTextValue(val As Variant, Optional ft As String = "") As String
    If ft <> "" Then
    getTextValue = getValueByType(val, ft)
    Exit Function
    End If
    If VarType(val) = vbDate Then
        getTextValue = Format(val, "m\/d\/yyyy")
    ElseIf VarType(val) = vbString Then
        getTextValue = Replace(Replace(CStr(val), Chr(10), ""), Chr(13), "")
    ElseIf VarType(val) = vbDouble Then
        getTextValue = Replace(CStr(val), ",", ".")
    ElseIf VarType(val) = vbDecimal Then
        getTextValue = Replace(CStr(val), ",", ".")
    ElseIf VarType(val) = vbCurrency Then
        getTextValue = Replace(CStr(val), ",", ".")
    End If
    
End Function
Function getValueByType(val As Variant, tp As String) As String
Select Case tp
Case "s"
    getValueByType = Replace(Replace(CStr(val), Chr(10), ""), Chr(13), "")
Case "d"
    getValueByType = Format(val, "m\/d\/yyyy")
Case "i"
    getValueByType = Replace(CStr(val), ",", ".")
Case "f"
    getValueByType = Replace(CStr(val), ",", ".")
Case "t"
    getValueByType = Format(val, "yyyy-mm-dd Hh:Nn:Ss")
Case Default
getValueByType = CStr(val)
End Select

End Function
Sub uploadFromRngByJavaEngine(wrkRange As Range, _
            fileName As String, _
            tbl As String, fieldTypes As String, _
            Optional withMacro As Boolean = False, _
            Optional enginePath As String = "M:\___Служба уценки\Отдел автоматизации распродаж и оперативной отчетности\Инструменты\third party\bin", _
            Optional visual As Boolean = True)
    On Error GoTo ERRLabel
    Dim textFilePath As String
    Dim superShell As Object
    Dim workdir As String
    Set superShell = CreateObject("WScript.Shell")
    workdir = enginePath
    textFilePath = prepareTextFileForJDBC(wrkRange, withMacro, fileName, visual, fieldTypes)
    fieldTypes = Replace(fieldTypes, "t", "s")
    uploadEnginCmd = "java -jar " & Chr(34) & workdir & "\UploadEngine.jar" & Chr(34) & " " & Chr(34) & textFilePath & Chr(34) & " " & tbl & " " & fieldTypes
    superShell.Run uploadEnginCmd, 0, True
    On Error GoTo 0
    Exit Sub
ERRLabel:
MsgBox "Error", vbCritical
End Sub

Sub createTmpTable(con As ADODB.Connection, tblName As String, fieldTypes As String)
Dim queryDrop As String
Dim queryCreate As String
Dim fieldCounter As Integer
Dim fieldType As String
Dim cmd As ADODB.Command
Dim beginChar As String

Set cmd = New ADODB.Command
Set cmd.ActiveConnection = con

cmd.CommandTimeout = 0
queryDrop = "call pricing_sale.drop_voltable('" & tblName & "')"
Debug.Print queryDrop

cmd.CommandText = queryDrop
cmd.Execute
indexString = ""
queryCreate = "create multiset table " & tblName & " ( "
For i = 1 To Len(fieldTypes)
    Select Case Mid(fieldTypes, i, 1)
    Case "i"
        fieldType = "INTEGER"
    Case "s"
        fieldType = "VARCHAR(256)"
    Case "f"
        fieldType = "FLOAT"
    Case "d"
        fieldType = "DATE"
    End Select
    If i = 1 Then
        beginChar = ""
    Else
        beginChar = " ,"
    End If
    
    queryCreate = queryCreate & beginChar & "a" & CStr(i) & " " & fieldType
    If i < 10 Then
        indexString = indexString & beginChar & "a" & CStr(i)
    End If
Next
    queryCreate = queryCreate & " ) primary index(" & indexString & ")"
    
Debug.Print queryCreate
cmd.CommandText = queryCreate
cmd.Execute
End Sub

Function copyFromRemoteSheetToMainSheet(srcSheet As Worksheet, _
        wsheet As Worksheet, _
        titleSearch As Variant, Optional rowInsert As Integer = 2, _
        Optional visual As Boolean = True)
    On Error GoTo ERRLabel



    Dim srcTotal As Long
    Dim titleRange As Range
    Dim titleCorr() As Integer
    Dim dimen As Integer
    Dim funErrVall As Integer
    dimen = UBound(titleSearch)
    ReDim titleCorr(dimen)
    
    

    funErrVall = clearFilter(srcSheet)
    If funErrVall <> 0 Then
        GoTo ERRLabel
    End If
    Set titleRange = srcSheet.Range("A1").Resize(, 120)
    For i = 0 To dimen
        titleCorr(i) = searchColumnPosition(titleRange, titleSearch(i), visual)
    Next
    srcTotal = srcSheet.Cells(srcSheet.Rows.Count, 1).End(xlUp).Row
    For i = 0 To dimen
        If titleCorr(i) > 0 Then

        srcSheet.Range(srcSheet.Cells(2, titleCorr(i)), _
            srcSheet.Cells(srcTotal, titleCorr(i))).Copy 'Destination:=
            wsheet.Cells(rowInsert, i + 1).PasteSpecial xlPasteValues
            wsheet.Cells(rowInsert, i + 1).PasteSpecial xlPasteFormats

        End If
    Next
    'srcWb.Close (False)
    On Error GoTo 0
    Exit Function
ERRLabel:
    Call errorHandler(visual)
End Function
Function searchColumnPosition(titleRng As Range, listSearch As Variant, _
        Optional visual As Boolean = True) As Integer
    On Error GoTo ERRLabel
    Dim tmpRng As Range
    Dim i As Integer
    i = 0
    Do While i <= UBound(listSearch)
        Set tmpRng = titleRng.Find(listSearch(i), LookAt:=xlWhole)
        If Not tmpRng Is Nothing Then
            Exit Do
        End If
        i = i + 1
    Loop
    
    If tmpRng Is Nothing Then
        searchColumnPosition = 0
    Else
        searchColumnPosition = tmpRng.Column
    End If
        On Error GoTo 0
    Exit Function
ERRLabel:
    Call errorHandler(visual)
End Function


Sub test()
Dim cn As ADODB.Connection
Set cn = getAdoCon
createTmpTable cn, "pricing_sale.lupakov_test_imp", "sifd"
End Sub
