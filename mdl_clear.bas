Attribute VB_Name = "mdl_clear"
Sub doClean(silent As Boolean)
  If Not silent Then
   inputResut = MsgBox("ВНИМАНИЕ: Будет очищена вкладка ""Редактор"". Убедитесь, что все необходимые изменения отправлены на расчет.", vbOKCancel, "Выполнение очистки")
    If inputResut = vbCancel Then
        Exit Sub
    End If
  End If
cleaning
 

End Sub

Sub cleaning()
   clearWorkSheet getWorkSheet()
    writeHeaderLine getColumns(), getWorkSheet(), 6, getLastRow()
 '   createPivotTable getWorkSheet()
End Sub
