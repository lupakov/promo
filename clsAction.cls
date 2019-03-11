VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private record_line_id As Long
Private wsheet As Worksheet
Private attrs As Dictionary
Private statusDict As Dictionary
Private geoDict As Dictionary
Private actTypeDict As Dictionary
Private frmtDict As Dictionary
Private kisTypeDict As Dictionary
Private weekDict As Dictionary
Private superList As ArrayList


Private Sub Class_Initialize()
    Set attrs = getColumns()

End Sub


Public Sub loadAction(record As Range, line As Long)

    record_line_id = line
    For Each k In attrs
    'Debug.Print k
    If Not IsNull(attrs(k)(2)) Then
        changeAttrsValue CStr(k), 3, record.columns(CLng(attrs(k)(2)))
    End If
    Next

    
End Sub

Private Function checkArtCode(con As ADODB.Connection) As Long
    Dim rs As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    rs.CursorType = adOpenKeyset
   ' Debug.Print "select art_id from src.dbo.v_art_ext where code = " & quotedString(CStr(attrs("art_code")(3)))
    rs.Open ("select art_id from src.dbo.v_art_ext where code = " & quotedString(CStr(attrs("art_code")(3))))
    If rs.RecordCount = 0 Then
        markErrorCell record_line_id, CLng(attrs("art_code")(2))
        Err.Raise 24002, , "Код товара не найден"
    End If
    checkArtCode = rs.Fields("ART_ID").value
End Function


Private Function checkLineArt(con As ADODB.Connection) As Long
    Dim rs As ADODB.Recordset
    If IsEmpty(attrs("art_id_line")(3)) Then
        checkLineArt = 0
        Exit Function
    End If
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    rs.CursorType = adOpenKeyset
   ' Debug.Print "select art_id from src.dbo.v_art_ext where code = " & quotedString(CStr(attrs("art_code")(3)))
    rs.Open ("select art_id from src.dbo.v_art_ext where code = " & quotedString(CStr(attrs("art_id_line")(3))))
    If rs.RecordCount = 0 Then
        markErrorCell record_line_id, CLng(attrs("art_id_line")(2))
        Err.Raise 24002, , "Код не найден"
    End If
    checkLineArt = rs.Fields("ART_ID").value
End Function


Private Function checkWindAnalogArt(con As ADODB.Connection) As Long
    Dim rs As ADODB.Recordset
    If IsEmpty(attrs("analog_co_id")(3)) Then
        checkWindAnalogArt = 0
        Exit Function
    End If
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    rs.CursorType = adOpenKeyset
   ' Debug.Print "select art_id from src.dbo.v_art_ext where code = " & quotedString(CStr(attrs("art_code")(3)))
    rs.Open ("select art_id from src.dbo.v_art_ext where code = " & quotedString(CStr(attrs("analog_co_id")(3))))
    If rs.RecordCount = 0 Then
        markErrorCell record_line_id, CLng(attrs("analog_co_id")(2))
        Err.Raise 24002, , "Код товара не найден"
    End If
    checkWindAnalogArt = rs.Fields("ART_ID").value
End Function

Private Function checkAnalogInOut(con As ADODB.Connection) As Long
    Dim rs As ADODB.Recordset
    If IsEmpty(attrs("analog_forecast_id")(3)) Then
        checkAnalogInOut = 0
        Exit Function
    End If
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    rs.CursorType = adOpenKeyset
   ' Debug.Print "select art_id from src.dbo.v_art_ext where code = " & quotedString(CStr(attrs("art_code")(3)))
    rs.Open ("select art_id from src.dbo.v_art_ext where code = " & quotedString(CStr(attrs("analog_forecast_id")(3))))
    If rs.RecordCount = 0 Then
        markErrorCell record_line_id, CLng(attrs("analog_forecast_id")(2))
        Err.Raise 24002, , "Код товара не найден"
    End If
    checkAnalogInOut = rs.Fields("ART_ID").value
End Function


Private Function checkFrmtCode() As Long
   If Not frmtDict.Exists(CStr(attrs("frmt")(3))) Then
           markErrorCell record_line_id, CLng(attrs("frmt")(2))
        Err.Raise 24021, , "Формат не найден"
   End If


    checkFrmtCode = frmtDict(CStr(attrs("frmt")(3)))
End Function

Private Function checkGeoCode() As Long
   If Not geoDict.Exists(CStr(attrs("geo_name")(3))) Then
           markErrorCell record_line_id, CLng(attrs("geo_name")(2))
        Err.Raise 24021, , "География не найдена"
   End If


    checkGeoCode = geoDict(CStr(attrs("geo_name")(3)))
End Function

Private Function checkActTypeCode() As Long
   If Not actTypeDict.Exists(CStr(attrs("act_sign_name")(3))) Then
           markErrorCell record_line_id, CLng(attrs("act_sign_name")(2))
        Err.Raise 24021, , "Тип акции не найден"
   End If


    checkActTypeCode = actTypeDict(CStr(attrs("act_sign_name")(3)))
End Function

Private Function checkKisType() As Long
If IsEmpty(attrs("ka_sign_name")(3)) Then
checkKisType = -1
Exit Function

End If
   If Not kisTypeDict.Exists(CStr(attrs("ka_sign_name")(3))) Then
           markErrorCell record_line_id, CLng(attrs("ka_sign_name")(2))
        Err.Raise 24021, , "Тип поставщика не найден"
   End If


    checkKisType = kisTypeDict(CStr(attrs("ka_sign_name")(3)))
End Function
Private Function checkWeekP() As Date
   If Not weekDict.Exists(CStr(attrs("week_id_p")(3))) Then
           markErrorCell record_line_id, CLng(attrs("week_id_p")(2))
        Err.Raise 24021, , "Неделя не найдена"
   End If


    checkWeekP = DateValue(Replace(CStr(weekDict(CStr(attrs("week_id_p")(3)))), """", ""))
End Function


Private Function chekKisCode(con As ADODB.Connection) As Long
    Dim rs As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    rs.CursorType = adOpenKeyset
    rs.Open ("select cntr_kis_id from src.dbo.v_cntr_kis where code = " & quotedString(CStr(attrs("kis_code")(3))))
    If rs.RecordCount = 0 Then
        markErrorCell record_line_id, CLng(attrs("kis_code")(2))
        Err.Raise 24002, , "Код КИС не найден"
    End If
    
    chekKisCode = rs.Fields("CNTR_KIS_ID").value
End Function

Public Sub upload(rs As ADODB.Recordset)
    
    
    rs.AddNew
    For Each k In attrs
        If CLng(attrs(k)(1)) Mod 10 = 1 Then
           ' Debug.Print attrs(k)(0)
            If Not IsEmpty(attrs(k)(3)) Then
                rs.Fields(attrs(k)(0)).value = attrs(k)(3)
            End If
        End If
    Next
    rs.Fields("EXCEL_LINE").value = record_line_id
    

End Sub


Public Sub uploadCancel(rs As ADODB.Recordset)
    
    
    rs.AddNew

    rs.Fields("ROW_ID").value = attrs("row_id")(3)
    rs.Fields("EXCEL_LINE").value = record_line_id
    rs.Fields("LOAD_USER_ID").value = attrs("load_user_id")(3)
    

End Sub

Public Sub uploadAccept(rs As ADODB.Recordset)
    
    
    rs.AddNew
    rs.Fields("EXCEL_LINE").value = record_line_id
    rs.Fields("ROW_ID").value = attrs("row_id")(3)
    rs.Fields("APP_ACT_PRICE").value = attrs("app_act_price")(3)
    rs.Fields("LOAD_USER_ID").value = attrs("load_user_id")(3)
    
    rs.Fields("REG_PRICE_NOTI").value = attrs("reg_price_noti")(3)
    rs.Fields("REGULAR_SALE_PRICE").value = attrs("regular_sale_price")(3)
    rs.Fields("FINAL_DISCOUNT").value = attrs("final_discount")(3)
    rs.Fields("BONUS").value = attrs("bonus")(3)

End Sub

Private Function checkUnique(con As ADODB.Connection, Optional tableName As String = "PRICING_SALE.PROMO.SUMMARY") As Long
    Dim rs As ADODB.Recordset
    If Not IsEmpty(attrs("row_id")(3)) Then
        checkUnique = 0
        Exit Function
    End If
    
    
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenKeyset
    Set rs.ActiveConnection = con
    rs.Open ("select row_id from " & tableName & " where art_id = " & _
    CStr(attrs("art_id")(3)) & " and frmt_id = " & CStr(attrs("frmt_id")(3)) & " and geo_id = " & CStr(attrs("geo_id")(3)) & _
    " and week_id_p = " & CStr(attrs("week_id_p")(3)))
    If rs.RecordCount = 0 Then
    checkUnique = 0
    Else
    markErrorCell record_line_id
    checkUnique = record_line_id
    Err.Raise 24070, , "Акция уже существует"
    End If
    
End Function

Private Function checkLine(con As ADODB.Connection, Optional tableName As String = "PRICING_SALE.PROMO.SUMMARY") As Long
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenKeyset
    Set rs.ActiveConnection = con
    rs.Open ("select distinct ART_ID_LINE from " & tableName & " where art_id = " & _
    CStr(attrs("art_id")(3)) & " and frmt_id = " & CStr(attrs("frmt_id")(3)) & " and week_id_p = " & _
    CStr(attrs("week_id_p")(3)) & " and geo_id <> " & CStr(attrs("geo_id")(3)))
    If rs.RecordCount = 0 Then
    checkLine = 0
    ElseIf IsNull(rs.Fields(0).value) And IsEmpty(attrs("art_id_line")(3)) Then
    checkLine = 0
    ElseIf rs.Fields(0).value Like CStr(attrs("art_id_line")(3)) Then
    checkLine = 0
    Else
    markErrorCell record_line_id, CLng(attrs("art_id_line")(2))
    checkLine = record_line_id
    Err.Raise 24060, , "Не уникальная линейка"
    End If
End Function

Private Function checkAnalogCo(con As ADODB.Connection, Optional tableName As String = "PRICING_SALE.PROMO.SUMMARY") As Long
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenKeyset
    Set rs.ActiveConnection = con
    rs.Open ("select distinct analog_co_id from " & tableName & " where art_id = " & _
    CStr(attrs("art_id")(3)) & " and frmt_id = " & CStr(attrs("frmt_id")(3)) & " and week_id_p = " & _
    CStr(attrs("week_id_p")(3)))
    If rs.RecordCount = 0 Then
    checkAnalogCo = 0
    ElseIf IsNull(rs.Fields(0).value) And IsEmpty(attrs("analog_co_id")(3)) Then
    checkAnalogCo = 0
    ElseIf rs.Fields(0).value Like CStr(attrs("analog_co_id")(3)) Then
    checkAnalogCo = 0
    Else
    markErrorCell record_line_id, CLng(attrs("analog_co_id")(2))
    checkAnalogCo = record_line_id
    Err.Raise 24060, , "Не уникальный аналог"
    End If
End Function

Private Function checkGeo(con As ADODB.Connection, Optional tableName As String = "PRICING_SALE.PROMO.SUMMARY") As Long
  Dim rs As ADODB.Recordset
  
  
  
  If attrs("geo_id")(3) <> 1 Then
    checkGeo = 0
    Exit Function
  Else
  
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenKeyset
    Set rs.ActiveConnection = con
    rs.Open ("select geo_id from " & tableName & " where art_id = " & CStr(attrs("art_id")(3)) & " and frmt_id = " & CStr(attrs("frmt_id")(3)) & " and week_id_p = " & CStr(attrs("week_id_p")(3)))
    If rs.RecordCount = 0 Then
    checkGeo = 0
    ElseIf rs.RecordCount = 1 And rs.Fields(0).value = 1 Then
    checkGeo = 0
    Else
    markErrorCell record_line_id, CLng(attrs("geo_name")(2))
    checkGeo = record_line_id
    Err.Raise 24050, , "Геогафия не уникальная"
    End If
  End If
End Function

Public Sub calFields()


If IsEmpty(attrs("row_id")(3)) Then
    changeAttrsValue "co_status_id", 3, 1
Else
    changeAttrsValue "co_status_id", 3, statusDict(attrs("co_status_name")(3))
End If

changeAttrsValue "frmt_id", 3, checkFrmtCode()
changeAttrsValue "geo_id", 3, checkGeoCode()
'changeAttrsValue "art_id", 3, checkArtCode(con)
changeAttrsValue "act_sign_id", 3, checkActTypeCode()
changeAttrsValue "ka_sign", 3, IIf(checkKisType() = -1, Null, checkKisType())
changeAttrsValue "expansion", 3, checkIntFromYesNo("expansion_name")
changeAttrsValue "not_ord_id", 3, checkIntFromYesNo("not_ord")
'changeAttrsValue "kas_end_dt", 4, IIf(attrs("act_sign_id")(3) = 1, 1, 0)
changeAttrsValue "act_begin_dt", 3, checkWeekP()
End Sub

Public Function checkAction(con As ADODB.Connection, _
                    uniqList As ArrayList, _
                    geoUniDict As Dictionary, _
                    lineUniqDict As Dictionary, _
                    analogUniqDict As Dictionary, _
                    tableName As String) As Long
                    
If attrs("co_status_id")(3) = 0 Or attrs("co_status_id")(3) = 3 Or attrs("co_status_id")(3) = 5 Or attrs("co_status_id")(3) = 6 Then
    Exit Function
End If
checkDataType
checkEmpty
chekKisCode con
checkLineArt con
checkWindAnalogArt con
checkAnalogInOut con
changeUniqIndex con, tableName
checkUnique con, tableName
checkLine con, tableName
checkAnalogCo con, tableName
checkGeo con, tableName
checkUniqueLocal uniqList
checkGeoLocal geoUniDict
checkLineUniqLocal lineUniqDict
checkAnalogUniqLocal analogUniqDict
checkDiscount

End Function

Private Sub changeAttrsValue(attr As String, index As Integer, value As Variant)
    Dim tmpArray As Variant
    tmpArray = attrs(attr)
    tmpArray(index) = value
    attrs(attr) = tmpArray
End Sub

Function getValueByType(value As Variant, tp As VarEnum) As Variant

End Function
Public Sub setUserInfo(userId As Long, bossId As Long)
    changeAttrsValue "create_user_id", 3, userId
    changeAttrsValue "user_id_boss", 3, bossId
    changeAttrsValue "load_user_id", 3, userId
End Sub

Public Sub setStatusDict(val As Dictionary)
Set statusDict = val
End Sub

Public Sub setFrmtDict(val As Dictionary)
Set frmtDict = val
End Sub

Public Sub setActTypeDict(val As Dictionary)
Set actTypeDict = val
End Sub

Public Sub setKisTypeDict(val As Dictionary)
Set kisTypeDict = val
End Sub


Public Sub setGeoDict(val As Dictionary)
Set geoDict = val
End Sub

Public Sub setWeekDict(val As Dictionary)
    Set weekDict = val
End Sub

Private Function checkIntFromYesNo(attrName As String) As Integer
    If IsEmpty(attrs(attrName)(3)) Then
        checkIntFromYesNo = 0
    ElseIf (Not UCase(CStr(attrs(attrName)(3))) Like "ДА") And (Not UCase(CStr(attrs(attrName)(3))) Like "НЕТ") Then
        markErrorCell record_line_id, CLng(attrs(attrName)(2))
        Err.Raise 24010, , "Значение не НЕТ/ДА"

    ElseIf UCase(CStr(attrs(attrName)(3))) Like "ДА" Then
        checkIntFromYesNo = 1
    Else
        checkIntFromYesNo = 0
    End If
End Function

Public Function checkUniqueLocal(uniqList As ArrayList) As Long

    Dim strIndex As String
    If Not IsEmpty(attrs("row_id")(3)) Then
        checkUniqueLocal = 0
        Exit Function
    End If
    
    strIndex = CStr(attrs("art_code")(3)) & "_" & CStr(attrs("frmt")(3)) & "_" & CStr(attrs("geo_id")(3)) & "_" & CStr(attrs("week_id_p")(3))
    If uniqList.Contains(strIndex) Then
        checkUniqueLocal = record_line_id
        markErrorCell record_line_id
        Err.Raise 24010, , "Дублирование ключевых полей"
    End If
    uniqList.Add strIndex
    checkUniqueLocal = 0
End Function

Public Function checkGeoLocal(geoUniDict As Dictionary) As Long
    Dim strIndex As String
    Dim geoList As ArrayList
    If Not IsEmpty(attrs("row_id")(3)) Then
        checkGeoLocal = 0
        Exit Function
    End If
    strIndex = CStr(attrs("art_code")(3)) & "_" & CStr(attrs("frmt")(3)) & "_" & CStr(attrs("week_id_p")(3))
    
    If attrs("geo_id")(3) <> 1 Then
        If geoUniDict.Exists(strIndex) Then
            If geoUniDict(strIndex).Contains("1") Then
               markErrorCell record_line_id, CLng(attrs("geo_name")(2))
               checkGeoLocal = record_line_id
               Err.Raise 24020, , "Геогафия не уникальная"
               Exit Function
            Else
               checkGeoLocal = 0
               Exit Function
            End If
        Else
            
            Set geoList = New ArrayList
            geoList.Add CStr(attrs("geo_id")(3))
            geoUniDict.Add strIndex, geoList
            checkGeoLocal = 0
            Exit Function
        End If
    Else
        If geoUniDict.Exists(strIndex) Then
            markErrorCell record_line_id, CLng(attrs("geo_name")(2))
            checkGeoLocal = record_line_id
            Err.Raise 24020, , "Геогафия не уникальная"
            Exit Function
        Else
            Set geoList = New ArrayList
            geoList.Add CStr(attrs("geo_id")(3))
            geoUniDict.Add strIndex, geoList
            checkGeoLocal = 0
            Exit Function
        End If

    End If
    
End Function

Public Function checkLineUniqLocal(lineUniqDict As Dictionary) As Long
 Dim strIndex As String
 strIndex = CStr(attrs("art_code")(3)) & "_" & CStr(attrs("frmt")(3)) & "_" & CStr(attrs("week_id_p")(3))
 If lineUniqDict.Exists(strIndex) Then
    If IIf(IsNull(attrs("art_id_line")(3)), "", CStr(attrs("art_id_line")(3))) Like lineUniqDict(strIndex) Then
        checkLineUniqLocal = 0
        Exit Function
    Else
        markErrorCell record_line_id, CLng(attrs("art_id_line")(2))
        checkLineUniqLocal = record_line_id
        Err.Raise 24030, , "Не уникальная линейка"
        Exit Function
    End If
 Else
    lineUniqDict.Add strIndex, IIf(IsNull(attrs("art_id_line")(3)), "", CStr(attrs("art_id_line")(3)))
    checkLineUniqLocal = 0
    Exit Function
 End If
End Function

Public Function checkAnalogUniqLocal(analogUniqDict As Dictionary) As Long
 Dim strIndex As String
 strIndex = CStr(attrs("art_code")(3)) & "_" & CStr(attrs("frmt")(3)) & "_" & CStr(attrs("week_id_p")(3))
 If analogUniqDict.Exists(strIndex) Then
    If IIf(IsNull(attrs("analog_co_id")(3)), "", CStr(attrs("analog_co_id")(3))) Like analogUniqDict(strIndex) Then
        checkAnalogUniqLocal = 0
        Exit Function
    Else
        markErrorCell record_line_id, CLng(attrs("analog_co_id")(2))
        checkAnalogUniqLocal = record_line_id
        Err.Raise 24030, , "Не уникальный аналог"
        Exit Function
    End If
 Else
    analogUniqDict.Add strIndex, IIf(IsNull(attrs("analog_co_id")(3)), "", CStr(attrs("analog_co_id")(3)))
    checkAnalogUniqLocal = 0
    Exit Function
 End If
End Function

Public Function changeUniqIndex(con As ADODB.Connection, Optional tableName As String = "PRICING_SALE.PROMO.SUMMARY") As Long
Dim rs As ADODB.Recordset

If IsEmpty(attrs("row_id")(3)) Then
    changeUniqIndex = 0
    Exit Function
Else
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    rs.CursorType = adOpenKeyset
    rs.Open ("select row_id from " & tableName & " where row_id = " & CStr(attrs("row_id")(3)) & " and  art_id = " & _
    CStr(attrs("art_id")(3)) & " and frmt_id = " & CStr(attrs("frmt_id")(3)) & " and geo_id = " & CStr(attrs("geo_id")(3)) & _
    " and week_id_p = " & CStr(attrs("week_id_p")(3)))
    If rs.RecordCount = 1 Then
        changeUniqIndex = 0
        Exit Function
    Else
        markErrorCell record_line_id
        changeUniqIndex = record_line_id
        Err.Raise 24080, , "Несоответствие акции и идентификатора ROW_ID"
    End If
End If
End Function

Public Function checkEmpty() As Long
For Each k In attrs.Keys
    If attrs(k)(4) = 1 Then
        If IsEmpty(attrs(k)(3)) Then
            markErrorCell record_line_id, CLng(attrs(k)(2))
            Err.Raise 24100, , "Обязательное поле не заполено"
        End If
    End If

Next
End Function

Private Sub markErrorCell(rowNum As Long, Optional colNum As Long = 0)
    If colNum = 0 Then
       wsheet.Cells(rowNum, 3).Interior.ColorIndex = 3
       wsheet.Cells(rowNum, 5).Interior.ColorIndex = 3
       wsheet.Cells(rowNum, 6).Interior.ColorIndex = 3
       wsheet.Cells(rowNum, 7).Interior.ColorIndex = 3
    Else
        wsheet.Cells(rowNum, colNum).Interior.Color = vbRed
    End If
End Sub

Public Sub setWorkSheet(sheetName As String)
    Set wsheet = ThisWorkbook.Worksheets(sheetName)
End Sub

Public Function checkDiscount() As Long
If attrs("final_discount")(3) <> attrs("invoice_discount")(3) + attrs("purchase_discount")(3) + attrs("advance_discount")(3) + attrs("promotion_discount")(3) Then
    markErrorCell record_line_id, CLng(attrs("final_discount")(2))
    checkDiscount = record_line_id
    Err.Raise 24500, , "Неверная итогавая скидка"
    Exit Function
    
End If
    checkDiscount = 0
End Function

Public Function checkDataType() As Long
Dim isErr As Boolean

For Each a In attrs
   ' Debug.Print a
    isErr = False
    If (Not IsNull(attrs(a)(5))) And (Not IsEmpty(attrs(a)(3))) Then
        Select Case attrs(a)(5)
        Case "f"
            If Not IsNumeric(attrs(a)(3)) Then
                isErr = True
            End If
        Case "i"
            If Not IsNumeric(attrs(a)(3)) Then
                isErr = True
            
            ElseIf Not Fix(attrs(a)(3)) = attrs(a)(3) Then
                isErr = True
            End If
        Case "d"
            If Not IsDate(attrs(a)(3)) Then
                    isErr = True
            End If
        End Select
    End If
If isErr Then
    markErrorCell record_line_id, CLng(attrs(a)(2))
    checkDataType = record_line_id
    Err.Raise 24525, , "Неверный тип данных"
End If
    
Next
End Function


Public Sub loadActionForCancel(record As Range, line As Long)
    record_line_id = line
    changeAttrsValue "row_id", 3, record.columns(CLng(attrs("row_id")(2)))
    changeAttrsValue "co_status_name", 3, record.columns(CLng(attrs("co_status_name")(2)))
    changeAttrsValue "co_status_id", 3, IIf(IsEmpty(attrs("co_status_name")(3)), 1, statusDict(attrs("co_status_name")(3)))
End Sub
Public Sub loadActionForAccept(record As Range, line As Long)
    record_line_id = line
    changeAttrsValue "row_id", 3, record.columns(CLng(attrs("row_id")(2)))
    changeAttrsValue "co_status_name", 3, record.columns(CLng(attrs("co_status_name")(2)))
    changeAttrsValue "co_status_id", 3, IIf(IsEmpty(attrs("co_status_name")(3)), 1, statusDict(attrs("co_status_name")(3)))
    changeAttrsValue "app_act_price", 3, record.columns(attrs("app_act_price")(2))
    
    changeAttrsValue "reg_price_noti", 3, record.columns(attrs("reg_price_noti")(2))
    changeAttrsValue "regular_sale_price", 3, record.columns(attrs("regular_sale_price")(2))
    changeAttrsValue "final_discount", 3, record.columns(attrs("final_discount")(2))
    changeAttrsValue "bonus", 3, record.columns(attrs("bonus")(2))
    
End Sub

Public Function readyForCancel() As Boolean
    If Not IsEmpty(attrs("row_id")(3)) And attrs("co_status_id")(3) <> 0 Then
        readyForCancel = True
    Else
        readyForCancel = False
    End If
    
End Function

Public Function readyForCmAccept() As Boolean
    If IsEmpty(attrs("app_act_price")(3)) Then
        readyForCmAccept = False
    ElseIf attrs("co_status_id")(3) = 4 Then
        readyForCmAccept = True
    ElseIf attrs("co_status_id")(3) = 5 And superList.Contains(10) Then
        readyForCmAccept = True
    Else
        readyForCmAccept = False
    End If
End Function

Public Sub setSuperList(val As ArrayList)
    Set superList = val
End Sub

Public Function readyForCalc() As Boolean
If attrs("act_begin_dt")(3) - Now() > 70 Or superList.Contains(1) Then

    If attrs("co_status_id")(3) = 1 Then
            readyForCalc = True
        ElseIf attrs("co_status_id")(3) = 2 Then
            readyForCalc = True
        ElseIf attrs("co_status_id")(3) = 4 Then
            readyForCalc = True
        Else
            readyForCalc = False
        End If
Else
    readyForCalc = False
End If
End Function
