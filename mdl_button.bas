Attribute VB_Name = "mdl_button"
Sub downloadWeek(ByVal control As IRibbonControl) 'Выгрузить неделю
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False
downloadWeekLnchr
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub downloadAll(ByVal control As IRibbonControl) 'Выгрузить всё
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False
downloadAllLnchr
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub clearSheet(ByVal control As IRibbonControl) 'Очистить редактор
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False
clearSheetLnchr
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub checkData(ByVal control As IRibbonControl) 'Проверить
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False
checkDataLnchr
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub runData(ByVal control As IRibbonControl) 'Рассчитать
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False
runDataLnchr
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub acceptData(ByVal control As IRibbonControl) 'Согласовать
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False
acceptDataLnchr
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub cancelData(ByVal control As IRibbonControl) 'Удалить
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False
cancelDataLnchr
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub downloadRepSvema(ByVal control As IRibbonControl) 'Сформировать шаблон
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False
downloadRepSvemaLnchr
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
