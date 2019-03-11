Attribute VB_Name = "Main"
Sub downloadWeekLnchr()
On Error GoTo LErr
unProtectSheet

loadWeek

setFreezeRange getWorkSheet()
protectSheet
Exit Sub
LErr:
setFreezeRange getWorkSheet()
protectSheet
MsgBox Err.Description, vbCritical
End Sub

Sub downloadAllLnchr()
On Error GoTo LErr
unProtectSheet

loadAll

setFreezeRange getWorkSheet()
protectSheet
Exit Sub
LErr:
setFreezeRange getWorkSheet()
protectSheet
MsgBox Err.Description, vbCritical
End Sub

Sub clearSheetLnchr(Optional silent As Boolean = False)
On Error GoTo LErr
unProtectSheet
doClean silent
setFreezeRange getWorkSheet()
protectSheet
Exit Sub
LErr:
setFreezeRange getWorkSheet()
protectSheet
MsgBox Err.Description, vbCritical
End Sub

Sub checkDataLnchr()
On Error GoTo LErr
unProtectSheet
checkActionData

setFreezeRange getWorkSheet()
protectSheet
Exit Sub
LErr:
setFreezeRange getWorkSheet()
protectSheet
MsgBox Err.Description, vbCritical
End Sub
Sub runDataLnchr()
On Error GoTo LErr
unProtectSheet
checkAndRunActionData

setFreezeRange getWorkSheet()
protectSheet
Exit Sub
LErr:
MsgBox Err.Description, vbCritical
End Sub
Sub acceptDataLnchr()
On Error GoTo LErr
unProtectSheet

acceptActionData
setFreezeRange getWorkSheet()
protectSheet
Exit Sub
LErr:
setFreezeRange getWorkSheet()
protectSheet
MsgBox Err.Description, vbCritical
End Sub
Sub cancelDataLnchr()
On Error GoTo LErr
unProtectSheet
cancelActionData
    setFreezeRange getWorkSheet()
protectSheet
Exit Sub
LErr:
    setFreezeRange getWorkSheet()
protectSheet
MsgBox Err.Description, vbCritical
End Sub


Sub downloadRepSvemaLnchr()
On Error GoTo LErr
unProtectSheet
downloadActionToSvema
setFreezeRange getWorkSheet()
protectSheet
Exit Sub
LErr:
setFreezeRange getWorkSheet()
protectSheet
MsgBox Err.Description, vbCritical
End Sub


