Attribute VB_Name = "mdl_main"



Sub clearWorkSheet(sht As Worksheet)
    headerLine = 6

 '  tmpltIn.Unprotect
 If Not IsEmpty(sht.Cells(headerLine, 1)) Then
    sht.Cells(headerLine, 1).AutoFilter
 End If
    sht.Cells.Delete
 '     tmpltIn.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
 '       False, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
 '       AllowFormattingRows:=True, AllowDeletingRows:=True, AllowSorting:=True, _
 '       AllowFiltering:=True, AllowUsingPivotTables:=True
End Sub



Function getArrayFromQuery(con As ADODB.Connection, query As String) As Variant
Dim rs As ADODB.Recordset
Dim arr As Variant
Dim i As Long
Dim index As Long
Dim name As String
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
Set rs.ActiveConnection = con

rs.Open (query)
ReDim arr(rs.RecordCount - 1)
i = 0
While Not rs.EOF
index = rs.Fields(0).value
name = rs.Fields(1).value
arr(i) = Array(index, name)
rs.MoveNext
i = i + 1
Wend
getArrayFromQuery = arr
End Function

Sub filNamesDict()
Dim arr As Variant
Dim con As ADODB.Connection
Set con = getAdoCon
On Error Resume Next
ThisWorkbook.Names("nmFrmt").Delete
ThisWorkbook.Names("nmActStatus").Delete
ThisWorkbook.Names("nmActType").Delete
ThisWorkbook.Names("nmGeo").Delete
ThisWorkbook.Names("nmKisType").Delete
ThisWorkbook.Names("nmWeeks").Delete
On Error GoTo 0
arr = getArrayFromQuery(con, "select frmt_id , name from pricing_sale.promo.frmt")
ThisWorkbook.Names.Add name:="nmFrmt", RefersTo:=arr

arr = getArrayFromQuery(con, "select status_id , status_name from pricing_sale.promo.action_status")
ThisWorkbook.Names.Add "nmActStatus", arr

arr = getArrayFromQuery(con, "select action_type_id , action_type_name from pricing_sale.promo.action_type")
ThisWorkbook.Names.Add "nmActType", arr

arr = getArrayFromQuery(con, "select cntr_id , name from pricing_sale.promo.v_rc")
ThisWorkbook.Names.Add "nmGeo", arr

arr = getArrayFromQuery(con, "select kis_type_id , name from pricing_sale.promo.kis_type")
ThisWorkbook.Names.Add "nmKisType", arr

arr = getArrayFromQuery(con, "select week_id , WEEK_BEGIN_DT from pricing_sale.promo.weeks where week_begin_dt between cast(getdate() as date) and dateadd(d, 380,cast(getdate() as date))")
ThisWorkbook.Names.Add "nmWeeks", arr


End Sub

Function nameToDict(nm As String, keyNum As Long, valNum As Long, Optional recSep = ";", Optional fieldSep = ",") As Dictionary
    Dim dict As Dictionary
    Dim clearString As String
    Dim recordArr As Variant
    Dim fieldArr As Variant
    Set dict = New Dictionary
    clearString = Replace(Replace(ThisWorkbook.Names(nm).value, "={", ""), "}", "")
    recordArr = VBA.split(clearString, recSep)
    For i = 0 To UBound(recordArr)
        fieldArr = VBA.split(recordArr(i), fieldSep)
        dict.Add fieldArr(keyNum), fieldArr(valNum)
    Next
    Set nameToDict = dict
End Function

Function getDictKeysString(dict As Dictionary, Optional sep As String = ",", Optional clearQuote As Boolean = True) As String
    Dim str As String
    str = Join(dict.Keys, sep)
    If clearQuote Then
    getDictKeysString = Replace(str, """", "")
    Else
    getDictKeysString = str
    End If
End Function


Sub unProtectSheet()
    getWorkSheet().Unprotect
End Sub


Sub protectSheet()

    getWorkSheet().Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
        False, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowDeletingRows:=True, AllowSorting:=True, _
        AllowFiltering:=True, AllowUsingPivotTables:=True
End Sub

Function clearQuotes(val As String) As String
clearQuotes = Replace(val, """", "")
End Function
Function clearEqualsSign(val As String) As String
clearEqualsSign = Replace(val, "=", "")
End Function

Function getWorkSheet() As Worksheet
    Dim shName As String
    shName = clearEqualsSign(clearQuotes(ThisWorkbook.Names("sheetName").value))
    Set getWorkSheet = ThisWorkbook.Worksheets(shName)
End Function

Sub setFreezeRange(wsh As Worksheet, Optional freezeRow As Long = 6, Optional freezeColumn As Long = 8)
    ThisWorkbook.Windows(1).FreezePanes = False
     With ThisWorkbook.Windows(1)
      .SplitColumn = freezeColumn
      .SplitRow = freezeRow
      .FreezePanes = True
    End With
End Sub



Function getLastRow() As Long
    Dim lastR As Long
    Dim wsheet As Worksheet
    
    Set wsheet = getWorkSheet()
    lastR = wsheet.Cells(wsheet.Rows.Count, 4).End(xlUp).Row
    getLastRow = lastR
End Function


Function checkInvalid(wsheet As Worksheet) As Boolean
    Dim flagInvalid As Boolean
    Dim validationRange As Range
    Dim c As Range
    flagInvalid = True
    Set wsheet = getWorkSheet
    Set validationRange = wsheet.Cells.SpecialCells(xlCellTypeAllValidation)
    For Each c In validationRange
        If Not c.Validation.value Then
            flagInvalid = False
            c.Interior.Color = vbRed
            addErrorMessage wsheet, c.Row, "Неверный тип данных; "
        End If
    Next
    checkInvalid = flagInvalid
End Function


Function firstCheckValues() As Boolean
    Dim wsheet As Worksheet
    Set wsheet = getWorkSheet()
    
    If Not checkInvalid(wsheet) Then
       ' wsheet.CircleInvalid
        firstCheckValues = False
    Else
        firstCheckValues = True
    End If
End Function






Function getColumns() As Dictionary
Dim attrs As Dictionary
Set attrs = New Dictionary
attrs.Add "row_id", Array("ROW_ID", 1, 1, Null, 0, Null, "Row_id", 5, 12105912, Null, "General", 0, Null, Null, Null, Null, Null)
attrs.Add "co_status_name", Array("CO_STATUS_NAME", 0, 2, Null, 0, Null, "Статус", 8, 12105912, Null, "General", 0, Null, Null, Null, Null, Null)
attrs.Add "template_error", Array(Null, 0, 3, Null, 0, Null, "Ошибки", 10, 12105912, Null, "General", 0, Null, Null, Null, Null, Null)
attrs.Add "art_code", Array("ART_CODE", 1, 4, Null, 1, Null, "Код позиции", 15, 3577398, 264364, "@", 1, xlValidateTextLength, xlEqual, "10", True, Null)
attrs.Add "art_name", Array("ART_NAME", 0, 5, Null, 0, Null, "Наименование ТП", 25, 12105912, Null, "@", 1, Null, Null, Null, Null, Null)
attrs.Add "frmt", Array("FRMT", 0, 6, Null, 1, Null, "Формат", 15, 3577398, 264364, "@", 1, xlValidateList, xlBetween, getDictKeysString(nameToDict("nmFrmt", 1, 0)), True, Null)
attrs.Add "geo_name", Array("GEO_NAME", 0, 7, Null, 1, Null, "География", 15, 3577398, 264364, "@", 1, xlValidateList, xlBetween, getDictKeysString(nameToDict("nmGeo", 1, 0)), True, Null)
attrs.Add "week_id_p", Array("WEEK_ID_P", 1, 8, Null, 1, Null, "№ промо недели (ГГГГНН)", 15, 3577398, 264364, "General", 1, xlValidateList, xlBetween, getDictKeysString(nameToDict("nmWeeks", 0, 1)), True, Null)
attrs.Add "boss_full_name", Array("BOSS_FULL_NAME", 0, 9, Null, 0, Null, "Руководитель категории", 15, 12105912, Null, "General", 1, Null, Null, Null, Null, Null)
attrs.Add "user_full_name", Array("USER_FULL_NAME", 0, 10, Null, 0, Null, "КМ", 15, 12105912, Null, "General", 1, Null, Null, Null, Null, Null)
attrs.Add "kis_code", Array("KIS_CODE", 1, 11, Null, 1, Null, "Код КИС КА", 15, 3577398, 264364, "General", 1, xlValidateTextLength, xlGreaterEqual, "10", True, Null)
attrs.Add "kis_name", Array("KIS_NAME", 0, 12, Null, 0, Null, "Наименование КИС КА", 15, 12105912, Null, "General", 1, Null, xlEqual, Null, Null, Null)
attrs.Add "art_grp_lvl_0_name", Array(Null, 0, 13, Null, 0, Null, "Группа 20", 15, 12105912, Null, "General", 1, Null, Null, Null, Null, Null)
attrs.Add "art_grp_lvl_1_name", Array(Null, 0, 14, Null, 0, Null, "Группа 21", 15, 12105912, Null, "General", 1, Null, Null, Null, Null, Null)
attrs.Add "art_grp_lvl_2_name", Array(Null, 0, 15, Null, 0, Null, "Группа 22", 15, 12105912, Null, "General", 1, Null, Null, Null, Null, Null)
attrs.Add "art_grp_lvl_3_name", Array(Null, 0, 16, Null, 0, Null, "Группа 23", 15, 12105912, Null, "General", 1, Null, Null, Null, Null, Null)
attrs.Add "import", Array(Null, 0, 17, Null, 0, Null, "СИ/МП", 15, 12105912, Null, "General", 1, Null, Null, Null, Null, Null)
attrs.Add "private_label", Array(Null, 0, 18, Null, 0, Null, "СТМ/СП", 15, 12105912, Null, "General", 1, Null, Null, Null, Null, Null)
attrs.Add "brand", Array(Null, 0, 19, Null, 0, Null, "Бренд", 15, 12105912, Null, "General", 1, Null, Null, Null, Null, Null)
attrs.Add "art_id_line", Array("ART_ID_LINE", 1, 20, Null, 0, Null, "Линейка", 15, 12105912, Null, "General", 1, xlValidateTextLength, xlEqual, "10", False, Null)
attrs.Add "act_sign_name", Array(Null, 0, 21, Null, 1, Null, "Признак акции", 15, 3577398, 264364, "General", 1, xlValidateList, xlBetween, getDictKeysString(nameToDict("nmActType", 1, 0)), True, Null)
attrs.Add "act_subtype", Array("ACT_SUBTYPE", 1, 22, Null, 1, Null, "Подтип акции", 15, 3577398, 264364, "General", 1, xlValidateTextLength, xlGreaterEqual, "1", True, Null)
attrs.Add "art_sign", Array("ART_SIGN", 1, 23, Null, 1, Null, "Признак позиции", 15, 3577398, 264364, "General", 1, xlValidateTextLength, xlGreaterEqual, "1", True, Null)
attrs.Add "act_duration", Array("ACT_DURATION", 1, 24, Null, 1, "f", "Продолжительность Акции, дней", 15, 3577398, 264364, "General", 1, xlValidateWholeNumber, xlGreater, "0", True, Null)
attrs.Add "ship_begin_dt", Array("SHIP_BEGIN_DT", 1, 25, Null, 1, "d", "Отгрузка С", 15, 3577398, 264364, "m/d/yyyy", 1, xlValidateDate, xlBetween, Format(Now - 365, "mm\/dd\/yyyy") & "|" & Format(Now + 365, "mm\/dd\/yyyy"), True, Null)
attrs.Add "ship_end_dt", Array("SHIP_END_DT", 1, 26, Null, 1, "d", "Отгрузка ПО", 15, 3577398, 264364, "m/d/yyyy", 1, xlValidateDate, xlBetween, Format(Now - 365, "mm\/dd\/yyyy") & "|" & Format(Now + 365, "mm\/dd\/yyyy"), True, Null)
attrs.Add "act_begin_dt", Array("ACT_BEGIN_DT", 1, 27, Null, 0, Null, "Дата акции С", 15, 12105912, Null, "m/d/yyyy", 1, Null, Null, Null, Null, Null)
attrs.Add "act_end_dt", Array("ACT_END_DT", 0, 28, Null, 0, Null, "Дата акции ПО", 15, 12105912, Null, "m/d/yyyy", 1, Null, Null, Null, Null, Null)
attrs.Add "kas_end_dt", Array("KAS_END_DT", 1, 29, Null, 1, "d", "Дата ПО для акции 'кассовый модуль' при использовании ин-аута", 15, 8574088, 2067705, "m/d/yyyy", 1, xlValidateDate, xlBetween, Format(Now - 365, "mm\/dd\/yyyy") & "|" & Format(Now + 365, "mm\/dd\/yyyy"), False, Null)
attrs.Add "ka_sign_name", Array(Null, 0, 30, Null, 0, Null, "Признак поставщика РЦ/МП", 15, 8574088, 2067705, "General", 1, xlValidateList, xlBetween, getDictKeysString(nameToDict("nmKisType", 1, 0)), False, Null)
attrs.Add "expansion_name", Array(Null, 0, 31, Null, 0, Null, "Расширение географии (допустимы значения да/нет)", 15, 8574088, 2067705, "General", 1, xlValidateList, xlBetween, "Да,Нет", False, Null)
attrs.Add "num_whs_promo", Array("NUM_WHS_PROMO", 0, 32, Null, 0, Null, "Кол-во ТТ с промо", 15, 12105912, Null, "General", 1, Null, Null, Null, Null, "=IF(RC[-1]=""Да"",RC[79],RC[78])")
attrs.Add "final_discount", Array("FINAL_DISCOUNT", 1, 33, Null, 1, "i", "Итоговая скидка Поставщика, %", 15, 3577398, 264364, "General", 1, xlValidateWholeNumber, xlGreaterEqual, "0", True, "=IF(AND(ISBLANK(RC[1]),ISBLANK(RC[2]),ISBLANK(RC[3]),ISBLANK(RC[4])),"""",SUM(RC[1]:RC[4]))")
attrs.Add "invoice_discount", Array("INVOICE_DISCOUNT", 1, 34, Null, 1, "i", "Скидка в накладной  %", 15, 3577398, 264364, "General", 1, xlValidateWholeNumber, xlGreaterEqual, "0", True, Null)
attrs.Add "purchase_discount", Array("PURCHASE_DISCOUNT", 1, 35, Null, 1, "i", "Скидка по итогам закупок, %", 15, 3577398, 264364, "General", 1, xlValidateWholeNumber, xlGreaterEqual, "0", True, Null)
attrs.Add "advance_discount", Array("ADVANCE_DISCOUNT", 1, 36, Null, 1, "i", "Скидка аванс, %", 15, 3577398, 264364, "General", 1, xlValidateWholeNumber, xlGreaterEqual, "0", True, Null)
attrs.Add "promotion_discount", Array("PROMOTION_DISCOUNT", 1, 37, Null, 1, "i", "Скидка по итогам акции, %", 15, 3577398, 264364, "General", 1, xlValidateWholeNumber, xlGreaterEqual, "0", True, Null)
attrs.Add "yield_discount", Array("YIELD_DISCOUNT", 1, 38, Null, 1, "i", "Скидка в доходность, %", 15, 3577398, 264364, "General", 1, xlValidateWholeNumber, xlGreaterEqual, "0", True, Null)
attrs.Add "bonus", Array("BONUS", 1, 39, Null, 0, "i", "Бонус, %", 15, 8574088, 2067705, "General", 1, xlValidateWholeNumber, xlGreaterEqual, "0", False, Null)
attrs.Add "inp_price_plan_p_period", Array("INP_PRICE_PLAN_P_PERIOD", 1, 40, Null, 1, "f", "Акционная цена уведомления", 15, 3577398, 264364, "#,##0.00", 1, xlValidateDecimal, xlGreaterEqual, "0", True, Null)
attrs.Add "reg_price_noti", Array("REG_PRICE_NOTI", 1, 41, Null, 1, "f", "Регулярная цена уведомления", 15, 3577398, 264364, "#,##0.00", 1, xlValidateDecimal, xlGreaterEqual, "0", True, Null)
attrs.Add "cat_man_price", Array("CAT_MAN_PRICE", 1, 42, Null, 0, "f", "Цель проведения промо", 15, 8574088, 2067705, "General", 1, xlValidateList, xlBetween, "Прибыль,РТО", False, Null)
attrs.Add "winding_discount", Array("WINDING_DISCOUNT", 1, 43, Null, 0, "i", "Скидка для Смотки", 15, 8574088, 2067705, "General", 1, xlValidateWholeNumber, xlGreaterEqual, "0", False, Null)
attrs.Add "analog_co_id", Array("ANALOG_CO_ID", 1, 44, Null, 0, Null, "Аналог для Смотки", 15, 8574088, 2067705, "General", 1, xlValidateTextLength, xlEqual, "10", False, Null)
attrs.Add "analog_forecast_id", Array("ANALOG_FORECAST_ID", 1, 45, Null, 0, Null, "Аналог для Прогноза ин-аутов полных новинок", 15, 8574088, 2067705, "General", 1, xlValidateTextLength, xlEqual, "10", False, Null)
attrs.Add "coef_analog_forecast", Array("COEF_ANALOG_FORECAST", 1, 46, Null, 0, "f", "коэф к аналогу прогноза", 15, 8574088, 2067705, "General", 1, xlValidateDecimal, xlGreaterEqual, "0", False, Null)
attrs.Add "presence_on_dmp", Array("PRESENCE_ON_DMP", 1, 47, Null, 0, "i", "Присутствие на ДМП", 15, 8574088, 2067705, "General", 1, xlValidateWholeNumber, xlGreaterEqual, "0", False, Null)

attrs.Add "competitor_p_price", Array("COMPETITOR_P_PRICE", 1, 48, Null, 0, "f", "Промо цена конкурента, руб", 15, 8574088, 2067705, "#,##0.00", 1, xlValidateDecimal, xlGreaterEqual, "0", False, Null)
attrs.Add "app_act_price", Array("APP_ACT_PRICE", 1, 49, Null, 0, "f", "Акционная Цена утвержденная", 15, 8574088, 2067705, "#,##0.00", 1, xlValidateDecimal, xlGreaterEqual, "0", False, Null)
attrs.Add "price_target_profit", Array("PRICE_TARGET_PROFIT", 0, 50, Null, 0, Null, "Цена (Цель: Прибыль)", 15, 12105912, Null, "#,##0.00", 1, Null, Null, Null, Null, Null)
attrs.Add "price_target_turnover", Array("PRICE_TARGET_TURNOVER", 0, 51, Null, 0, Null, "Цена (Цель: Товарооборот)", 15, 12105912, Null, "#,##0.00", 1, Null, Null, Null, Null, Null)

attrs.Add "regular_sale_price", Array("REGULAR_SALE_PRICE", 0, 52, Null, 0, "f", "Регулярная розничная цена", 15, 12105912, Null, "#,##0.00", 1, Null, Null, Null, False, Null)

attrs.Add "profit_by_cm", Array("PROFIT_BY_CM", 0, 53, Null, 0, Null, "Комм маржа Руб (Утвержденная)", 15, 12105912, Null, "#,##0", 1, Null, Null, Null, Null, "=IF(ISBLANK(RC[-4]),"""",IFERROR((RC[-4]-(RC[-12]*(1-RC[-20]*0.01)*(1-RC[-14]*0.01)))*RC[7],"""" ))")
attrs.Add "profit_by_profit", Array("PROFIT_BY_PROFIT", 0, 54, Null, 0, Null, "Комм маржа Руб (Цель: Прибыль)", 15, 12105912, Null, "#,##0", 1, Null, Null, Null, Null, "=IF(ISBLANK(RC[-4]),"""",IFERROR((RC[-4]-(RC[-13]*(1-RC[-21]*0.01)*(1-RC[-15]*0.01)))*RC[7],"""" ))")
attrs.Add "profit_by_turnover", Array("PROFIT_BY_TURNOVER", 0, 55, Null, 0, Null, "Комм маржа Руб (Цель: Товарооборот)", 15, 12105912, Null, "#,##0", 1, Null, Null, Null, False, "=IF(ISBLANK(RC[-4]),"""",IFERROR((RC[-4]-(RC[-14]*(1-RC[-16]*0.01)*(1-RC[-22]*0.01)))*RC[7],"""" ))")
attrs.Add "com_margin_by_cm", Array("COM_MARGIN_BY_CM", 0, 56, Null, 0, Null, "Ком.Маржа % (Утвержденная)", 15, 12105912, Null, "0.00%", 1, Null, Null, Null, False, "=IF(ISBLANK(RC[-7]),"""",IFERROR((RC[-7]-RC[-15]*(1-RC[-23]*0.01)*(1-RC[-17]*0.01))/RC[-7],"""" ))")
attrs.Add "com_margin_by_profit", Array("COM_MARGIN_BY_PROFIT", 0, 57, Null, 0, Null, "Ком.Маржа % (Цель: Прибыль)", 15, 12105912, Null, "0.00%", 1, Null, xlEqual, Null, Null, "=IF(ISBLANK(RC[-7]),"""",IFERROR((RC[-7]-RC[-16]*(1-RC[-24]*0.01)*(1-RC[-18]*0.01))/RC[-7],"""" ))")
attrs.Add "com_margin_by_turnover", Array("COM_MARGIN_BY_TURNOVER", 0, 58, Null, 0, Null, "Ком.Маржа %(Цель: Товарооборот)", 15, 12105912, Null, "0.00%", 1, Null, Null, Null, Null, "=IF(ISBLANK(RC[-7]),"""",IFERROR((RC[-7]-RC[-17]*(1-RC[-25]*0.01)*(1-RC[-19]*0.01))/RC[-7],"""" ))")
attrs.Add "avg_margin_category", Array("AVG_MARGIN_CATEGORY", 0, 59, Null, 0, Null, "СФ ком.маржа категори", 15, 12105912, Null, "General", 1, Null, Null, Null, Null, Null)
attrs.Add "forecast_by_cm", Array("FORECAST_BY_CM", 0, 60, Null, 0, Null, "Прогноз в ШТ (Цена КМ)", 15, 12105912, Null, "#,##0", 1, Null, Null, Null, Null, "=IF(ISBLANK(RC[-11]),"""",IFERROR(RC[9]*RC[-28]*RC[-36]*((EXP(RC[53]*(1-RC[-11]/RC[-8])+RC[52])-EXP(RC[52])+1)),"""" ))")
attrs.Add "forecast_by_profit", Array("FORECAST_BY_PROFIT", 0, 61, Null, 0, Null, "Прогноз в ШТ (Цель: Прибыль)", 15, 12105912, Null, "#,##0", 1, Null, Null, Null, Null, "=IF(ISBLANK(RC[-11]),"""",IFERROR(RC[8]*RC[-29]*RC[-37]*((EXP(RC[52]*(1-RC[-11]/RC[-9])+RC[51])-EXP(RC[51])+1)),"""" ))")
attrs.Add "forecast_by_turnover", Array("FORECAST_BY_TURNOVER", 0, 62, Null, 0, Null, "Прогноз в ШТ (Цель: Товарооборот)", 15, 12105912, Null, "#,##0", 1, Null, Null, Null, Null, "=IF(ISBLANK(RC[-11]),"""",IFERROR(RC[7]*RC[-30]*RC[-38]*((EXP(RC[51]*(1-RC[-11]/RC[-10])+RC[50])-EXP(RC[50])+1)),"""" ))")
attrs.Add "rest_whs_dc", Array("REST_WHS_DC", 0, 63, Null, 0, "f", "Остаток ТТ+РЦ, шт", 15, 12105912, Null, "#,##0", 1, Null, Null, Null, Null, Null)
attrs.Add "need_ord", Array("NEED_ORD", 1, 64, Null, 0, Null, "Требуется заказать", 15, 8574088, 2067705, "#,##0.00", 1, xlValidateDecimal, xlGreaterEqual, "0", False, Null)
attrs.Add "not_ord", Array(Null, 0, 65, Null, 0, Null, "Не заказывать у поставщика", 15, 8574088, 2067705, "General", 1, xlValidateList, xlBetween, "Да,Нет", False, Null)
attrs.Add "inv_before_ord", Array("INV_BEFORE_ORD", 0, 66, Null, 0, Null, "ТЗ до заказа, дней", 15, 12105912, Null, "#,##0", 1, Null, Null, Null, Null, "=IFERROR(RC[-3]/(RC[42]*RC[44]),"""" )")
attrs.Add "inv_after_ord", Array("INV_AFTER_ORD", 0, 67, Null, 0, Null, "ТЗ после заказа, дней", 15, 12105912, Null, "#,##0", 1, Null, Null, Null, Null, Null)
attrs.Add "inv_after_action", Array("INV_AFTER_ACTION", 0, 68, Null, 0, Null, "ТЗ после акции, дней", 15, 12105912, Null, "#,##0", 1, Null, Null, Null, Null, "=IFERROR(IF(RC[-3]=""Да"",(RC[-5]-RC[-8])/(RC[40]*RC[42]),RC[-2]),"""" )")
attrs.Add "avg_sales_regular", Array("AVG_SALES_REGULAR", 0, 69, Null, 0, Null, "Среднеденвые продажи по рег.розн. Цене, шт", 15, 12105912, Null, "General", 1, Null, Null, Null, Null, "=IF(RC[-38]=""Да"",RC[40],RC[39])")
attrs.Add "turnover_min_part_1_whs_day", Array("TURNOVER_MIN_PART_1_WHS_DAY", 1, 70, Null, 0, "i", "Оборачиваемость мин.партии на 1 ТТ, дн", 15, 8574088, 2067705, "General", 1, xlValidateWholeNumber, xlGreaterEqual, "0", False, Null)
attrs.Add "week_id_last_p", Array("WEEK_ID_LAST_P", 1, 71, Null, 0, "i", "Неделя последнего промо", 15, 8574088, 2067705, "General", 1, xlValidateWholeNumber, xlGreaterEqual, "0", False, Null)
attrs.Add "subtype_last_p", Array("SUBTYPE_LAST_P", 1, 72, Null, 0, Null, "Подтип последнего промо", 15, 8574088, 2067705, "General", 1, Null, Null, Null, Null, Null)
attrs.Add "inp_price_p_period", Array("INP_PRICE_P_PERIOD", 1, 73, Null, 0, "f", "Входная цена в период промо", 15, 8574088, 2067705, "#,##0.00", 1, xlValidateDecimal, xlGreaterEqual, "0", False, Null)
attrs.Add "sale_price_last_p", Array("SALE_PRICE_LAST_P", 1, 74, Null, 0, "f", "Прод. цена последнего промо", 15, 8574088, 2067705, "#,##0.00", 1, xlValidateDecimal, xlGreaterEqual, "0", False, Null)
attrs.Add "shelf_discount_last_p", Array("SHELF_DISCOUNT_LAST_P", 1, 75, Null, 0, "i", "Скидка на полку в последнее промо", 15, 8574088, 2067705, "General", 1, xlValidateWholeNumber, xlGreaterEqual, "0", False, Null)
attrs.Add "fact_sales_1_whs_last_p_it", Array("FACT_SALES_1_WHS_LAST_P_IT", 1, 76, Null, 0, "f", "Факт. Продажи на 1 ТТ в последнем промо, шт.", 15, 8574088, 2067705, "#,##0.00", 1, xlValidateDecimal, xlGreaterEqual, "0", False, Null)
attrs.Add "whs_vol_sales_rest", Array("WHS_VOL_SALES_REST", 1, 77, Null, 0, "i", "Количество ТТ с продажами/остатками", 15, 8574088, 2067705, "#,##0", 1, xlValidateWholeNumber, xlGreaterEqual, "0", False, Null)
attrs.Add "perc_implement_p_in_last_p", Array("PERC_IMPLEMENT_P_IN_LAST_P", 1, 78, Null, 0, "i", "% реализации промо продукции в предыдущем промо", 15, 8574088, 2067705, "General", 1, xlValidateWholeNumber, xlGreaterEqual, "0", False, Null)
attrs.Add "compensation_method", Array("COMPENSATION_METHOD", 1, 79, Null, 0, Null, "Способ компенсации", 15, 8574088, 2067705, "General", 1, Null, Null, Null, False, Null)
attrs.Add "avg_reg_inp_price_np_3_m", Array("AVG_REG_INP_PRICE_NP_3_M", 1, 80, Null, 0, "f", "Ср. регулярная входная цена, без промо руб за 3 месяца", 15, 8574088, 2067705, "General", 1, xlValidateDecimal, xlGreaterEqual, "0", False, Null)
attrs.Add "avg_reg_sales_price_np_3_m", Array("AVG_REG_SALES_PRICE_NP_3_M", 1, 81, Null, 0, "f", "Ср. регулярная продажная цена, без промо руб за 3 месяца", 15, 8574088, 2067705, "#,##0.00", 1, xlValidateDecimal, xlGreaterEqual, "0", False, Null)
attrs.Add "cur_reg_inp_price", Array("CUR_REG_INP_PRICE", 1, 82, Null, 0, "f", "Текущая регулярная вх цена, руб", 15, 8574088, 2067705, "General", 1, xlValidateDecimal, xlGreaterEqual, "0", False, Null)
attrs.Add "cur_reg_shelf_price", Array("CUR_REG_SHELF_PRICE", 1, 83, Null, 0, "f", "Текущая регулярная полочная цена, руб", 15, 8574088, 2067705, "#,##0.00", 1, xlValidateDecimal, xlGreaterEqual, "0", False, Null)
attrs.Add "reg_week_turnover_np_3_m_it", Array("REG_WEEK_TURNOVER_NP_3_M_IT", 1, 84, Null, 0, "f", "Регулярный недельный оборот в шт без промо за 3 месяца", 15, 8574088, 2067705, "#,##0.00", 1, xlValidateDecimal, xlGreaterEqual, "0", False, Null)
attrs.Add "reg_week_turnover_np_3_m_rub", Array("REG_WEEK_TURNOVER_NP_3_M_RUB", 1, 85, Null, 0, "f", "Регулярный недельный оборот в руб без промо за 3 месяца", 15, 8574088, 2067705, "#,##0.00", 1, xlValidateDecimal, xlGreaterEqual, "0", False, Null)
attrs.Add "avg_reg_week_turnover", Array("AVG_REG_WEEK_TURNOVER", 1, 86, Null, 0, "f", "Ср. регулярный недельный оборот в шт без промо за 3 месяца на 1 ТТ", 15, 8574088, 2067705, "#,##0.00", 1, xlValidateDecimal, xlGreaterEqual, "0", False, Null)
attrs.Add "avg_week_whs_vol_np_3_m", Array("AVG_WEEK_WHS_VOL_NP_3_M", 1, 87, Null, 0, "f", "Ср. недельное количество ТТ без промо за 3 месяца", 15, 8574088, 2067705, "#,##0.00", 1, xlValidateDecimal, xlGreaterEqual, "0", False, Null)
attrs.Add "dc_vol_con", Array("DC_VOL_CON", 1, 88, Null, 0, "i", "Кол-во РЦ с подключением", 15, 8574088, 2067705, "#,##0", 1, xlValidateWholeNumber, xlGreaterEqual, "0", False, Null)
attrs.Add "max_whs_vol_on_rc_con", Array("MAX_WHS_VOL_ON_RC_CON", 1, 89, Null, 0, "i", "Максимальное кол-во ТТ на РЦ с подключением", 15, 8574088, 2067705, "General", 1, xlValidateWholeNumber, xlGreaterEqual, "0", False, Null)
attrs.Add "cur_whs_con_in_chain", Array("CUR_WHS_CON_IN_CHAIN", 1, 90, Null, 0, "i", "Текущее подключение ТТ в сети", 15, 8574088, 2067705, "#,##0", 1, xlValidateWholeNumber, xlGreaterEqual, "0", False, Null)
attrs.Add "cs", Array("CS", 1, 91, Null, 0, "i", "ЦС", 15, 12105912, Null, "#,##0", 1, xlValidateWholeNumber, xlGreaterEqual, "0", False, Null)
attrs.Add "min_part_pick", Array("MIN_PART_PICK", 1, 92, Null, 0, "i", "Мин партия отборки", 15, 8574088, 2067705, "General", 1, xlValidateWholeNumber, xlGreaterEqual, "0", False, Null)
attrs.Add "cur_inv_whs_it", Array("CUR_INV_WHS_IT", 1, 93, Null, 0, "f", "ТЗ на текущую дату в шт.(ТТ)", 15, 8574088, 2067705, "#,##0", 1, xlValidateDecimal, xlGreaterEqual, "0", False, Null)
attrs.Add "cur_inv_whs_rub", Array("CUR_INV_WHS_RUB", 1, 94, Null, 0, "f", "ТЗ на текущую дату в руб.(ТТ)", 15, 8574088, 2067705, "#,##0.00", 1, xlValidateDecimal, xlGreaterEqual, "0", False, Null)
attrs.Add "cur_inv_whs_day", Array("CUR_INV_WHS_DAY", 1, 95, Null, 0, "f", "ТЗ на текущую дату в дн.(ТТ)", 15, 8574088, 2067705, "#,##0.00", 1, xlValidateDecimal, xlGreaterEqual, "0", False, Null)
attrs.Add "cur_inv_dc_it", Array("CUR_INV_DC_IT", 1, 96, Null, 0, "f", "ТЗ на текущую дату в шт.(РЦ)", 15, 8574088, 2067705, "#,##0.00", 1, xlValidateDecimal, xlGreaterEqual, "0", False, Null)
attrs.Add "cur_inv_dc_rub", Array("CUR_INV_DC_RUB", 1, 97, Null, 0, "f", "ТЗ на текущую дату в руб.(РЦ)", 15, 8574088, 2067705, "#,##0.00", 1, xlValidateDecimal, xlGreaterEqual, "0", False, Null)
attrs.Add "cur_inv_dc_day", Array("CUR_INV_DC_DAY", 1, 98, Null, 0, "f", "ТЗ на текущую дату в дн.(РЦ)", 15, 8574088, 2067705, "#,##0.00", 1, xlValidateDecimal, xlGreaterEqual, "0", False, Null)
attrs.Add "avg_loss_last_3_m", Array("AVG_LOSS_LAST_3_M", 1, 99, Null, 0, "i", "Потери, % (среднее за 3 месяца)", 15, 8574088, 2067705, "#,##0.00", 1, xlValidateWholeNumber, xlGreaterEqual, "0", False, Null)
attrs.Add "matrix_level", Array("MATRIX_LEVEL", 1, 100, Null, 0, "i", "Уровень матрицы", 15, 8574088, 2067705, "#,##0", 1, xlValidateWholeNumber, xlGreaterEqual, "0", False, Null)
attrs.Add "art_equipment", Array("ART_EQUIPMENT", 1, 101, Null, 0, Null, "Оборудование в котором размещается позиция", 15, 8574088, 2067705, "General", 1, Null, Null, Null, Null, Null)
attrs.Add "sold_layout_catalog", Array("SOLD_LAYOUT_CATALOG", 1, 102, Null, 0, Null, "Проданный макет в каталоге", 15, 8574088, 2067705, "General", 1, Null, Null, Null, Null, Null)
attrs.Add "cost_layout", Array("COST_LAYOUT", 1, 103, Null, 0, "f", "Стоимость макета", 15, 8574088, 2067705, "General", 1, xlValidateDecimal, xlGreaterEqual, "0", False, Null)
attrs.Add "dcm_decision", Array("DCM_DECISION", 1, 104, Null, 0, Null, "Решение ДЗ", 15, 8574088, 2067705, "General", 1, Null, Null, Null, True, Null)
attrs.Add "dcm_comment", Array("DCM_COMMENT", 1, 105, Null, 0, Null, "Комментарий ДЗ", 15, 8574088, 2067705, "General", 1, Null, Null, Null, Null, Null)
attrs.Add "protection_decision", Array("PROTECTION_DECISION", 1, 106, Null, 0, Null, "Решение по итогам защиты", 15, 8574088, 2067705, "General", 1, Null, Null, Null, Null, Null)
attrs.Add "protection_comment", Array("PROTECTION_COMMENT", 1, 107, Null, 0, Null, "Комментарий по итогам защиты", 15, 8574088, 2067705, "General", 1, Null, Null, Null, Null, Null)
attrs.Add "avg_wo_expansion", Array("AVG_WO_EXPANSION", 0, 108, Null, 0, Null, "Среднедневные Без расширния", 15, 12105912, Null, "#,##0.00", 1, Null, Null, Null, Null, Null)
attrs.Add "avg_with_expansion", Array("AVG_WITH_EXPANSION", 0, 109, Null, 0, Null, "Среднедневные С расширнием", 15, 12105912, Null, "#,##0.00", 1, Null, Null, Null, Null, Null)
attrs.Add "whs_num_wo_expansion", Array("WHS_NUM_WO_EXPANSION", 0, 110, Null, 0, Null, "Количество ТТ Без расширения", 15, 12105912, Null, "#,##0", 1, Null, Null, Null, Null, Null)
attrs.Add "whs_num_with_expansion", Array("WHS_NUM_WITH_EXPANSION", 0, 111, Null, 0, Null, "Количество ТТ С расширением", 15, 12105912, Null, "#,##0", 1, Null, Null, Null, Null, Null)
attrs.Add "elast_a", Array("ELAST_A", 1, 112, Null, 0, Null, "КОЭФ А", 15, 12105912, Null, "General", 1, Null, Null, Null, Null, Null)
attrs.Add "elast_b", Array("ELAST_B", 1, 113, Null, 0, Null, "КОЭФ В", 15, 12105912, Null, "General", 1, Null, Null, Null, Null, Null)
attrs.Add "turnover", Array(Null, 0, 114, Null, 0, Null, "Оборот за период акции по Утвержденной цене", 15, 12105912, Null, "#,##0.00", 1, Null, Null, Null, Null, "=IFERROR(RC[-54]*RC[-65],"""" )")
attrs.Add "load_dtm", Array("LOAD_DTM", 1, Null, Null, 0, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null)
attrs.Add "art_id", Array("ART_ID", 0, Null, Null, 0, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null)
attrs.Add "frmt_id", Array("FRMT_ID", 1, Null, Null, 0, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null)
attrs.Add "geo_id", Array("GEO_ID", 1, Null, Null, 0, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null)
attrs.Add "user_id_boss", Array("USER_ID_BOSS", 1, Null, Null, 0, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null)
attrs.Add "create_user_id", Array("CREATE_USER_ID", 1, Null, Null, 0, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null)
attrs.Add "act_sign_id", Array("ACT_SIGN_ID", 1, Null, 1, 0, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null)
attrs.Add "ka_sign", Array("KA_SIGN", 1, Null, Null, 0, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null)
attrs.Add "expansion", Array("EXPANSION", 1, Null, Null, 0, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null)
attrs.Add "not_ord_id", Array("NOT_ORD", 1, Null, Null, 0, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null)
attrs.Add "load_user_id", Array("LOAD_USER_ID", 1, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null)
attrs.Add "co_status_id", Array("CO_STATUS", 1, Null, Null, 0, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null)
Set getColumns = attrs
End Function

Sub writeHeaderLine(dict As Dictionary, wsheet As Worksheet, headerLine As Long, formulaLastRow As Long, Optional lastRow As Long = 10000)
    For Each a In dict
        If Not IsNull(dict(a)(2)) Then
            With wsheet.Cells(headerLine, dict(a)(2))
                .value = dict(a)(6)
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = True
                
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).ColorIndex = 0
                .Borders(xlEdgeLeft).TintAndShade = 0
                .Borders(xlEdgeLeft).Weight = xlThin
        
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).ColorIndex = 0
                .Borders(xlEdgeTop).TintAndShade = 0
                .Borders(xlEdgeTop).Weight = xlThin
        
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).ColorIndex = 0
                .Borders(xlEdgeBottom).TintAndShade = 0
                .Borders(xlEdgeBottom).Weight = xlThin
        
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).ColorIndex = 0
                .Borders(xlEdgeRight).TintAndShade = 0
                .Borders(xlEdgeRight).Weight = xlThin
                
            End With

            wsheet.columns(dict(a)(2)).ColumnWidth = dict(a)(7)
            wsheet.Cells(headerLine, dict(a)(2)).Interior.Color = dict(a)(8)
            If Not IsNull(dict(a)(9)) Then
                wsheet.Cells(headerLine - 1, dict(a)(2)).Interior.Color = dict(a)(9)
            End If
            If CStr(dict(a)(10)) Like "m/d/yyyy" Then
                rangeValueToDate wsheet.Range(wsheet.Cells(headerLine + 1, dict(a)(2)), wsheet.Cells(lastRow, dict(a)(2)))
            End If
          
            wsheet.Range(wsheet.Cells(headerLine + 1, dict(a)(2)), wsheet.Cells(lastRow, dict(a)(2))).NumberFormat = CStr(dict(a)(10))
            If dict(a)(11) = 1 Then
                wsheet.Range(wsheet.Cells(headerLine, dict(a)(2)), wsheet.Cells(wsheet.Rows.Count, dict(a)(2))).Locked = False
            End If
            If Not IsNull(dict(a)(12)) Then
                setValidation wsheet.Range(wsheet.Cells(headerLine + 1, dict(a)(2)), _
                    wsheet.Cells(lastRow, dict(a)(2))), dict(a)(12), dict(a)(13), CStr(dict(a)(14))
            End If
            
            If (Not IsNull(dict(a)(16))) And (formulaLastRow > 6 Or CLng(dict(a)(2)) = 33) Then
            
                wsheet.Range(wsheet.Cells(headerLine + 1, dict(a)(2)), wsheet.Cells(IIf(CLng(dict(a)(2)) = 33, 10000, formulaLastRow), dict(a)(2))).FormulaR1C1 = dict(a)(16)
            End If
            
        End If
    
    Next
        wsheet.Rows("5:5").RowHeight = 5
        wsheet.Rows("6:6").RowHeight = 90
        wsheet.Cells(headerLine, 1).AutoFilter
        wsheet.Range("BR21:DC21").EntireColumn.Hidden = True
        
        
    wsheet.Range("D3").FormulaR1C1 = _
        "=IFERROR(SUBTOTAL(9,R[4]C[49]:R[1048573]C[49])/SUBTOTAL(9,R[4]C[110]:R[1048573]C[110]),"""" )"
    wsheet.Range("D3").NumberFormat = "0.00%"
    wsheet.Range("E3").FormulaR1C1 = "=SUBTOTAL(9,R[4]C[48]:R[1048573]C[48])"
    wsheet.Range("E3").NumberFormat = "#,##0"
    wsheet.Range("F3").FormulaR1C1 = "=SUBTOTAL(9,R[4]C[108]:R[1048573]C[108])"
    wsheet.Range("F3").NumberFormat = "#,##0"
    
    wsheet.Range("D2").FormulaR1C1 = "Комм маржа, %"
    
    wsheet.Range("E2").FormulaR1C1 = "Комм маржа, руб"
    wsheet.Range("F2").FormulaR1C1 = "Прогноз РТО"
    
    
    wsheet.Range("D2:F3").Borders(xlEdgeLeft).Weight = xlMedium
    wsheet.Range("D2:F3").Borders(xlEdgeTop).Weight = xlMedium
    wsheet.Range("D2:F3").Borders(xlEdgeBottom).Weight = xlMedium
    wsheet.Range("D2:F3").Borders(xlEdgeRight).Weight = xlMedium
    wsheet.Range("D2:F3").Borders(xlInsideVertical).Weight = xlThin
    wsheet.Range("D2:F3").Borders(xlInsideHorizontal).Weight = xlThin
    
    

    wsheet.Range("D2:F3").HorizontalAlignment = xlCenter
    wsheet.Range("D2:F3").VerticalAlignment = xlCenter

    
    wsheet.Range("D2:F2").Interior.Color = 12105912


    
End Sub

Sub setValidation(rng As Range, valType As Variant, operType As Variant, formula As String, Optional ignor As Boolean = True)
    rng.Validation.Delete
    
    
    a = VBA.split(formula, "|")
    
    
    
    If UBound(a) = 0 Then
        rng.Validation.Add Type:=valType, AlertStyle:=xlValidAlertStop, _
        Operator:=operType, Formula1:=a(0)
    ElseIf UBound(a) = 1 Then
        rng.Validation.Add Type:=valType, AlertStyle:=xlValidAlertStop, _
        Operator:=operType, Formula1:=a(0), Formula2:=a(1)
    End If
    rng.Validation.IgnoreBlank = ignor
    rng.Validation.InCellDropdown = True
    rng.Validation.InputTitle = ""
    rng.Validation.ErrorTitle = ""
    rng.Validation.InputMessage = ""
    rng.Validation.ErrorMessage = ""
    rng.Validation.ShowInput = True
    rng.Validation.ShowError = True
    

End Sub

Sub updateValidation(dict As Dictionary, wsheet As Worksheet, headerLine As Long, Optional lastRow As Long = 10000)
    For Each a In dict
        If Not IsNull(dict(a)(2)) Then
            If Not IsNull(dict(a)(12)) Then
                setValidation wsheet.Range(wsheet.Cells(headerLine + 1, dict(a)(2)), _
                    wsheet.Cells(lastRow, dict(a)(2))), dict(a)(12), dict(a)(13), CStr(dict(a)(14)), Not CBool(dict(a)(15))
            End If
            Debug.Print a
        End If
    Next
End Sub


Sub addErrorMessage(wsheet As Worksheet, rowNum As Long, msg As String)
    If Not Trim(msg) Like Trim(wsheet.Cells(rowNum, 3).value) Then
        wsheet.Cells(rowNum, 3).value = wsheet.Cells(rowNum, 3).value & msg
    End If
End Sub

Sub cleanErrorMessage(wsheet As Worksheet)
    wsheet.Range(wsheet.Cells(7, 3), wsheet.Cells(wsheet.Rows.Count, 3)).Clear
End Sub
Function getClientVersion() As Integer
getClientVersion = 0
End Function
Function getDataBaseVersion(con As ADODB.Connection) As Integer
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim version As Integer
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = con
    cmd.CommandText = "select value from PRICING_SALE.PROMO.INFO where upper(attribute) = 'VERSION'"
    Set rs = cmd.Execute
    If Not rs.EOF Then
        version = CInt(rs.Fields(0).value)
    Else
        version = -1
    End If
    getDataBaseVersion = version
End Function

Function checkVersion(con As ADODB.Connection) As Boolean
    Dim xlVersion As Integer
    Dim dbVersion As Integer
    
    xlVersion = getClientVersion()
    dbVersion = getDataBaseVersion(con)
    checkVersion = (xlVersion = dbVersion)
End Function

Function checkPerson(con As ADODB.Connection) As Boolean
    Dim rs  As ADODB.Recordset
    Set rs = con.Execute("select id from pricing_sale.promo.users where user_login = " & quotedString(Environ("username")))
    
    If rs.EOF Then
        checkPerson = False
    Else
        checkPerson = True
    End If
End Function

Function checkDictionary(con As ADODB.Connection) As Boolean
    On Error Resume Next
        ThisWorkbook.Names("nmFrmt").Delete
        ThisWorkbook.Names("nmActStatus").Delete
        ThisWorkbook.Names("nmActType").Delete
        ThisWorkbook.Names("nmGeo").Delete
        ThisWorkbook.Names("nmKisType").Delete
        ThisWorkbook.Names("nmWeeks").Delete
    On Error GoTo LError
        arr = getArrayFromQuery(con, "select frmt_id , name from pricing_sale.promo.frmt")
        ThisWorkbook.Names.Add name:="nmFrmt", RefersTo:=arr
        
        arr = getArrayFromQuery(con, "select status_id , status_name from pricing_sale.promo.action_status")
        ThisWorkbook.Names.Add "nmActStatus", arr
        
        arr = getArrayFromQuery(con, "select action_type_id , action_type_name from pricing_sale.promo.action_type")
        ThisWorkbook.Names.Add "nmActType", arr
        
        arr = getArrayFromQuery(con, "select cntr_id , name from pricing_sale.promo.v_rc")
        ThisWorkbook.Names.Add "nmGeo", arr
        
        arr = getArrayFromQuery(con, "select kis_type_id , name from pricing_sale.promo.kis_type")
        ThisWorkbook.Names.Add "nmKisType", arr
        
        arr = getArrayFromQuery(con, "select week_id , WEEK_BEGIN_DT from pricing_sale.promo.weeks where week_begin_dt between cast(getdate() as date) and dateadd(d, 380,cast(getdate() as date))")
        ThisWorkbook.Names.Add "nmWeeks", arr
        checkDictionary = True
    Exit Function
LError:
   checkDictionary = False
End Function

Function checkInitial(con As ADODB.Connection) As Boolean

    If Not checkVersion(con) Then
        MsgBox "Версия Excel файла не соответствует версии БД", vbCritical
        checkInitial = False
        Exit Function
    End If
    
    If Not checkPerson(con) Then
        MsgBox "Пользователь не найден", vbCritical
        checkInitial = False
        Exit Function
    End If
    
    If Not checkDictionary(con) Then
        MsgBox "Не обновились необходимые словари", vbCritical
        checkInitial = False
        Exit Function
    End If
    
    checkInitial = True
    
End Function
    
    
    Sub rangeValueToDate(rng As Range)
    Dim c As Range
    For Each c In rng.Cells
        If Not IsEmpty(c.value) Then
            c.value = CDate(c.value)
        End If
    Next
    End Sub
    
    Sub logPressButton(con As ADODB.Connection, buttonId As Integer, Optional buttonValue As Long = -1)
    
        Dim cmd As ADODB.Command
        
        Dim btnPar As ADODB.Parameter
        Dim valPar As ADODB.Parameter
        Dim userPar As ADODB.Parameter
        
        Set cmd = New ADODB.Command
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "pricing_sale.promo.log_press_button"
        
        Set btnPar = cmd.CreateParameter("btn", adInteger, adParamInput, 8, buttonId)
        Set valPar = cmd.CreateParameter("val", adInteger, adParamInput, 8, buttonValue)
        Set userPar = cmd.CreateParameter("user", adVarChar, adParamInput, 100, Environ("username"))
        
        cmd.Parameters.Append btnPar
        cmd.Parameters.Append valPar
        cmd.Parameters.Append userPar
        
        cmd.Execute
    

End Sub
