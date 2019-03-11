Attribute VB_Name = "Library_common"
Function clearFilter(wsh As Worksheet, _
        Optional range_str As String = "A1") As Integer
    'Очистка Авто фильтра
    On Error GoTo ERRLabel
        wsh.Range(range_str).AutoFilter
    On Error GoTo 0
    clearFilter = 0
    Exit Function
ERRLabel:
    clearFilter = 1
End Function

Function replaceKeyValue(rng As Range, _
        dict As Dictionary, _
        Optional rev As Boolean = True) As Integer
    'Замена в диапозоне ключей на значение
    'Или наоборот в зависимости от параметра rev
    Dim k As Variant
    On Error GoTo ERRLabel
    If rev Then
        If Not dict Is Nothing Then
            For Each k In dict.Keys()
                If replaceInRange(rng, k, dict(k)) > 0 Then
                    GoTo ERRLabel
            Next
        End If
    Else
        If Not dict Is Nothing Then
            For Each k In dict.Keys()
                If replaceInRange(rng, k, dict(k)) > 0 Then
                    GoTo ERRLabel
            Next
        End If
    End If
    replaceKeyValue = 0
    Exit Function
ERRLabel:
    replaceKeyValue = 1
End Function

Function replaceInRange(rng As Range, _
        k As String, _
        v As String) As Integer
    'Замена в диапозоне
    On Error GoTo ERRLabel
    rng.Replace what:=k, Replacement:=v, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False _
        , ReplaceFormat:=False
    replaceInRange = 0
    Exit Function
ERRLabel:
    replaceInRange = 1
End Function
Function getAdoCon() As ADODB.Connection
        Dim con As ADODB.Connection
        Set con = New ADODB.Connection
        'con.ConnectionString = "Driver={SQL Server};Server=pricecraft;"
        con.ConnectionString = "Provider=SQLOLEDB; Server=pricecraft;Database=PRICING_SALE; Trusted_Connection=yes"
        con.CommandTimeout = 0
        con.ConnectionTimeout = 0

        con.Open
        Set getAdoCon = con

    
End Function

Function closeAdoCon(con As ADODB.Connection) As Integer
If Not con Is Nothing Then
If con.State = 1 Then
    con.Close
End If
    Set con = Nothing
End If

closeAdoCon = 0
End Function

Function getCurrentRegion(initCell As Range, _
        Optional clearHeader As Boolean = True, _
        Optional filtred As Boolean = True, _
        Optional widthConstr As Integer = 0, _
        Optional heightConstr As Integer = 0) As Range
    Dim tmp As Range
    Dim wrkReg As Range
    Dim headerOffset As Integer
    Dim wid_col As Long
    Dim heig_col As Long

    If Not filtred Then
        clearFilter initCell.Worksheet, initCell.address(0, 0)
    End If
    
    If clearHeader Then
        headerOffset = 1
    Else
        headerOffset = 0
    End If
    Set tmp = initCell.CurrentRegion.SpecialCells(xlCellTypeVisible)
    
    If tmp.Rows.Count < 2 Then
        Set getCurrentRegion = Nothing
        Exit Function
    End If
    
    If widthConstr > 0 Then
        wid_col = widthConstr
    Else
        wid_col = tmp.columns.Count
    End If
    
    If heightConstr > 0 Then
        heig_col = heightConstr
    Else
        heig_col = tmp.Rows.Count
    End If

    Set wrkReg = tmp.Resize(heig_col - headerOffset, wid_col).Offset(headerOffset)
    Set getCurrentRegion = wrkReg
        
End Function
Function getFiltredRowCount(initCell As Range)
    getFiltredRowCount = initCell.CurrentRegion.SpecialCells(xlCellTypeVisible).Rows.Count - 1
End Function

Sub test()
getCurrentRegion Range("a1")
End Sub
Function getWorkRegion(curCell As Range, _
        Optional visual As Boolean = True) As Range
    On Error GoTo labelErr
    Dim allReg As Range
    Dim wrkReg As Range
    Set allReg = curCell.CurrentRegion
    If allReg.Rows.Count < 2 Then
        Set getWorkRegion = Nothing
        Exit Function
    End If
    Set wrkReg = allReg.Resize(allReg.Rows.Count - 1, allReg.columns.Count).Offset(1)
    Set getWorkRegion = wrkReg
        On Error GoTo 0
    Exit Function
labelErr:
    errorHandler visual
End Function
Function prepareDate(rng As Range, Optional tmplate As String = "yyyy-mm-dd") As Integer
    Dim item As Range
    On Error GoTo ERRLabel
    rng.NumberFormat = "yyyy-mm-dd"
    For Each item In rng.Cells
        item.value = Format(item.value, tmplate)
    Next
    prepareDate = 0
    Exit Function
ERRLabel:
    prepareDate = 1
End Function
Sub prepareFloat(rng As Range, Optional visual As Boolean = True)
    On Error GoTo labelErr
        rng.NumberFormat = "#0.0#"
rng.value = rng.value

'Dim cell As Range
'For Each cell In rng.Cells
'    cell.NumberFormat = "#0.0#"
'
'Next
        On Error GoTo 0
    Exit Sub
labelErr:
    errorHandler visual
End Sub
Sub preparePercent(rng As Range, Optional visual As Boolean = True)
    On Error GoTo labelErr

        

 

Dim cell As Range
For Each cell In rng.Cells
'cell.Value = Replace(cell.Value, "%", "")
cell.value = Replace(cell.value, ",", ".")
Next
rng.NumberFormat = "General"
rng.value = rng.value
        On Error GoTo 0
    Exit Sub
labelErr:
    errorHandler visual
End Sub



Function prepareTimestamp(rng As Range, Optional tmplate As String = "yyyy-mm-dd Hh:Nn:Ss") As Integer
    Dim item As Range
    On Error GoTo ERRLabel
    rng.NumberFormat = "yyyy-mm-dd hh:mm:ss"
    For Each item In rng.Cells

        item.value = Format(CDate(item.value), tmplate)
    Next
    prepareTimestamp = 0
    Exit Function
ERRLabel:
    prepareTimestamp = 1
End Function

Sub InputWriteTitle(sht As Worksheet, titles As Variant, _
        titleRow As Long, Optional visual As Boolean = True)
    On Error GoTo ERRLabel
    Dim columnsCount As Long
    Dim tit As Variant

    columnsCount = UBound(titles) + 1
    sht.Rows(titleRow).Clear

    With sht
        For i = 1 To columnsCount
            .Cells(titleRow, i).value = titles(i - 1)
        Next
    End With
    
    With sht.Range(sht.Cells(titleRow, 1), sht.Cells(titleRow, columnsCount)).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent3
            .TintAndShade = 0.799981688894314
           .PatternTintAndShade = 0
    End With
    With sht.Range(sht.Cells(titleRow, 1), sht.Cells(titleRow, columnsCount))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
            .columns.ColumnWidth = 15
            .Rows.AutoFit
            .Font.Bold = True
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
    End With

    On Error GoTo 0
    Exit Sub
ERRLabel:
  '  Call errorHandler(visual)
End Sub

Sub trimRng(rng As Range, Optional visual As Boolean = True)
    On Error GoTo labelErr
Dim cell As Range
For Each cell In rng.Cells
    cell.value = Application.WorksheetFunction.Trim(Replace(cell.value, Chr(160), Chr(32)))
    'cell.Value = Replace(cell.Value, Chr(34), "")
    'cell.Value = Replace(cell.Value, Chr(39), "")
    cell.value = Replace(cell.value, Chr(10), "")
    cell.value = Replace(cell.value, Chr(13), "")
Next
        On Error GoTo 0
    Exit Sub
labelErr:
    errorHandler visual
End Sub

Sub trimSpace(rng As Range)
    Dim cell As Range
    For Each cell In rng.Cells
        cell.value = Application.WorksheetFunction.Trim(Replace(cell.value, Chr(32), ""))
    Next
End Sub

Function quotedString(str As String) As String
    quotedString = "'" & str & "'"
End Function
