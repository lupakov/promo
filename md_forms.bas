Attribute VB_Name = "md_forms"

Function getArtInfoByCode(con As ADODB.Connection, _
                            code As String) As Dictionary
    Dim rs As ADODB.Recordset
    Dim ans As Dictionary
    Dim f As ADODB.Field
    If cacheInfo("ArtInfo").Exists(code) Then
        Set getArtInfoByCode = cacheInfo("ArtInfo")(code)
        Exit Function
    End If
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    rs.Open "select art_id , name , art_grp_lvl_0_id " & _
        " , art_grp_lvl_1_id, art_grp_lvl_2_id, art_grp_lvl_3_id " & _
        " , art_grp_lvl_0_name, art_grp_lvl_1_name, art_grp_lvl_2_name " & _
        " , art_grp_lvl_3_name, private_label, import, grp_vend_art_id, grp_vend_art_name " & _
        " from v_art_ext where code = '" & code & "'"
        
    Set ans = New Dictionary
    For Each f In rs.Fields
        ans.Add f.name, f.value
    Next
    cacheInfo("ArtInfo").Add code, ans
    Set getArtInfoByCode = ans
End Function

Sub fillArtInfo(code As String, _
                txtArtName As MSForms.TextBox, _
                txtGrp0Name As MSForms.TextBox, _
                txtGrp1Name As MSForms.TextBox, _
                txtGrp2Name As MSForms.TextBox, _
                txtGrp3Name As MSForms.TextBox, _
                txtVendor As MSForms.TextBox, _
                txtImport As MSForms.TextBox, _
                txtPrivateLabel As MSForms.TextBox, _
                txtArtId As MSForms.TextBox, _
                txtGrp0Id As MSForms.TextBox, _
                txtGrp1Id As MSForms.TextBox, _
                txtGrp2Id As MSForms.TextBox, _
                txtGrp3Id As MSForms.TextBox, _
                txtVendorId As MSForms.TextBox, _
                txtImportId As MSForms.TextBox, _
                txtPrivateLabelId As MSForms.TextBox)
    Dim con As ADODB.Connection
    Dim artDict As Dictionary
    Set con = getAdoCon
    Set artDict = getArtInfoByCode(con, code)
    con.Close
    txtArtName.Text = artDict("NAME")
    txtGrp0Name.Text = artDict("ART_GRP_LVL_0_NAME")
    txtGrp1Name.Text = artDict("ART_GRP_LVL_1_NAME")
    txtGrp2Name.Text = artDict("ART_GRP_LVL_2_NAME")
    txtGrp3Name.Text = artDict("ART_GRP_LVL_3_NAME")
    txtVendor.Text = artDict("GRP_VEND_ART_NAME")
    txtImport.Text = IIf(artDict("IMPORT") = 1, "СИ", "МП")
    txtPrivateLabel = IIf(artDict("PRIVATE_LABEL") = 1, "СТМ", "СП")
    txtArtId.Text = artDict("ART_ID")
    txtGrp0Id.Text = artDict("ART_GRP_LVL_0_ID")
    txtGrp1Id.Text = artDict("ART_GRP_LVL_1_ID")
    txtGrp2Id.Text = artDict("ART_GRP_LVL_2_ID")
    txtGrp3Id.Text = artDict("ART_GRP_LVL_3_ID")
    txtVendorId.Text = artDict("GRP_VEND_ART_ID")
    txtImportId.Text = artDict("IMPORT")
    txtPrivateLabelId = artDict("PRIVATE_LABEL")
    
End Sub


Function getKisInfoByCode(con As ADODB.Connection, _
                            code As String) As Dictionary
    Dim rs As ADODB.Recordset
    Dim ans As Dictionary
    Dim f As ADODB.Field
    
    If cacheInfo("KisInfo").Exists(code) Then
        Set getKisInfoByCode = cacheInfo("KisInfo")(code)
        Exit Function
    End If
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    rs.Open "select cntr_kis_id , cntr_fullname as NAME " & _
        " from v_cntr_kis where code = '" & code & "'"
        
    Set ans = New Dictionary
    For Each f In rs.Fields
        ans.Add f.name, f.value
    Next
    cacheInfo("KisInfo").Add code, ans
    Set getKisInfoByCode = ans
End Function

Sub fillKisInfo(code As String, _
                txtSupplierName As MSForms.TextBox)
    Dim con As ADODB.Connection
    Dim kisDict As Dictionary
    Set con = getAdoCon
    Set kisDict = getKisInfoByCode(con, code)
    con.Close
    txtSupplierName.Text = kisDict("NAME")
 
End Sub


Function fillListBoxbyQuery(con As ADODB.Connection, targetList As MSForms.ListBox, sqlText As String)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    targetList.Clear
    rs.Open sqlText
    
    While Not rs.EOF
    targetList.AddItem rs.Fields(1).value
    targetList.list(targetList.ListCount - 1, 1) = rs.Fields(0).value
    rs.MoveNext
    Wend

End Function

