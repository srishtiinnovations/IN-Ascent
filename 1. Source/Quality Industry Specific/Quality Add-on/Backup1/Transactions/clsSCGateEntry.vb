
Public Class clsSCGateEntry
    Private oForm As SAPbouiCOM.Form
    Private oMatrix As SAPbouiCOM.Matrix
    Dim objRS As SAPbobsCOM.Recordset
    Dim strSQL As String
    Dim oDT As SAPbouiCOM.DataTable
    Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
    Dim oCFL As SAPbouiCOM.ChooseFromList
    Dim oCons As SAPbouiCOM.Conditions
    Dim oCon As SAPbouiCOM.Condition
    Dim oCmb As SAPbouiCOM.ComboBox
    Dim i As Integer
    Dim oEdit1, oEdit2, oEdit3 As SAPbouiCOM.EditText
    Public Const formtype As String = "Frm_SCGE"
    Public Sub LoadScreen()
        Try
            oForm = objAddOn.objUIXml.LoadScreenXML("SubConGateEntry.xml", SST.enuResourceType.Embeded, formtype)
            oForm.Items.Item("6").Specific.value = objAddOn.objGenFunc.GetDocNum("SST_GATSUB")
            oForm.Items.Item("10").Specific.string = objAddOn.objGenFunc.GetDateTimeValue(objAddOn.SBO_Application.Company.ServerDate).ToString("dd/MM/yy")
            oForm.DataBrowser.BrowseBy = "6"
            AddCflCon()

            LoadCombo()
        Catch ex As Exception
            objAddOn.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub

    Private Sub LoadCombo()
        oCmb = oForm.Items.Item("1000002").Specific

        For i = oCmb.ValidValues.Count - 1 To 0 Step -1
            oCmb.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
        Next
        objRS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strSQL = "select code,name from [@SST_MODE]"
        objRS.DoQuery(strSQL)
        While Not objRS.EoF
            oCmb.ValidValues.Add(objRS.Fields.Item("code").Value, objRS.Fields.Item("name").Value)
            objRS.MoveNext()
        End While
        oCmb.ValidValues.Add("Define", "Define New")
    End Sub

    Private Sub LoadCombo1()
        oCmb = oForm.Items.Item("33").Specific

        For i = oCmb.ValidValues.Count - 1 To 0 Step -1
            oCmb.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
        Next
        objRS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strSQL = "select docnum,series from OWOR where U_ScCode = '" & oForm.Items.Item("4").Specific.string & "' and status = 'R'"
        objRS.DoQuery(strSQL)
        While Not objRS.EoF
            oCmb.ValidValues.Add(objRS.Fields.Item("docnum").Value, objRS.Fields.Item("docnum").Value)
            objRS.MoveNext()
        End While
    End Sub

    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.Before_Action = True Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        oForm = objAddOn.SBO_Application.Forms.Item(FormUID)
                        oMatrix = oForm.Items.Item("19").Specific
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If pVal.ItemUID = "1000002" Then
                            LoadCombo()
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        'If pVal.ItemUID = "1000001" Then
                        '    oMatrix = oForm.Items.Item("19").Specific
                        '    If oMatrix.RowCount = 0 Then
                        '        oMatrix.AddRow()
                        '        oMatrix.GetLineData(oMatrix.RowCount)
                        '        oForm.DataSources.DBDataSources.Item("@SST_SUBGAT1").SetValue("LineId", 0, oMatrix.RowCount)
                        '        oMatrix.SetLineData(oMatrix.RowCount)
                        '    Else
                        '        oForm.DataSources.DBDataSources.Item("@SST_SUBGAT1").Clear()
                        '        oMatrix.AddRow()
                        '        oMatrix.GetLineData(oMatrix.RowCount)
                        '        oForm.DataSources.DBDataSources.Item("@SST_SUBGAT1").SetValue("LineId", 0, oMatrix.RowCount)
                        '        oMatrix.SetLineData(oMatrix.RowCount)
                        '    End If

                        'End If

                        If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            If Validate(pVal) = False Then
                                BubbleEvent = False
                            End If
                        ElseIf pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            If oForm.Items.Item("31").Specific.selected.value = "C" Then
                                objAddOn.SBO_Application.SetStatusBarMessage("Document Closed Cannot be updated")
                                BubbleEvent = False
                            Else
                                objRS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                strSQL = "select * from [@SST_NQCHDR] where U_GENo ='" & oForm.Items.Item("6").Specific.string & "' "
                                objRS.DoQuery(strSQL)
                                If objRS.RecordCount > 0 Then
                                    objAddOn.SBO_Application.SetStatusBarMessage("Cannot be updated")
                                    BubbleEvent = False
                                Else
                                    If Validate(pVal) = False Then
                                        BubbleEvent = False
                                    End If
                                End If
                            End If
                        End If
                End Select
            Else
                Select Case pVal.EventType

                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        If pVal.ItemUID = "4" Or pVal.ItemUID = "12" Or (pVal.ItemUID = "19" Or pVal.ColUID = "0") Then
                            Choose(FormUID, pVal)
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                        Try
                            If pVal.ItemUID = "33" Then
                                Dim code As String
                                oCmb = oForm.Items.Item("33").Specific
                                code = oCmb.Selected.Value
                                objRS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                strSQL = "select a.ItemCode,b.ItemName, a.PlannedQty from owor a inner join OITM b on a.itemcode = b.itemcode where a.docnum ='" & code & "'"
                                objRS.DoQuery(strSQL)

                                If objRS.RecordCount > 0 Then
                                    If oMatrix.RowCount > 0 Then
                                        oMatrix.Clear()
                                    End If

                                    Dim rowno As Integer
                                    rowno = 0
                                    objRS.MoveFirst()

                                    While Not objRS.EoF
                                        oMatrix.AddRow()
                                        rowno += 1
                                        oMatrix.GetLineData(oMatrix.RowCount)
                                        oForm.DataSources.DBDataSources.Item("@SST_SUBGAT1").SetValue("LineId", 0, oMatrix.RowCount)
                                        oForm.DataSources.DBDataSources.Item("@SST_SUBGAT1").SetValue("U_ItmCode", 0, objRS.Fields.Item("ItemCode").Value)
                                        oForm.DataSources.DBDataSources.Item("@SST_SUBGAT1").SetValue("U_ItmDesc", 0, objRS.Fields.Item("ItemName").Value)
                                        oForm.DataSources.DBDataSources.Item("@SST_SUBGAT1").SetValue("U_Qty", 0, objRS.Fields.Item("PlannedQty").Value)
                                        'oForm.DataSources.DBDataSources.Item("@SST_SUBGAT1").SetValue("U_smpsize", 0, objRS.Fields.Item("U_smpsize").Value)
                                        oMatrix.SetLineData(rowno)
                                        objRS.MoveNext()
                                    End While

                                Else
                                    objAddOn.SBO_Application.SetStatusBarMessage("No matching records....", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                End If
                            End If
                        Catch ex As Exception
                        End Try

            If pVal.ItemUID = "1000002" Then
                oCmb = oForm.Items.Item("1000002").Specific
                If oCmb.Selected.Value = "Define" Then
                    Dim j As Integer
                    Dim omenus As SAPbouiCOM.MenuItem
                    omenus = objAddOn.SBO_Application.Menus.Item("47616")
                    For j = 0 To omenus.SubMenus.Count - 1
                        strSQL = omenus.SubMenus.Item(j).String
                        If strSQL.StartsWith("SST_MOD") = True Then
                            objAddOn.SBO_Application.ActivateMenuItem(omenus.SubMenus.Item(j).UID.ToString)
                            Exit For
                        End If
                    Next
                End If
            End If

                End Select
            End If
            If pVal.ItemUID = "1" And pVal.Action_Success = True And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Items.Item("6").Specific.value = objAddOn.objGenFunc.GetDocNum("SST_GATSUB")
                    oForm.Items.Item("10").Specific.string = objAddOn.objGenFunc.GetDateTimeValue(objAddOn.SBO_Application.Company.ServerDate).ToString("dd/MM/yy")
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Function Validate(ByVal pVal) As Boolean
        oForm = objAddOn.SBO_Application.Forms.Item(pVal.FormUID)
        oMatrix = oForm.Items.Item("19").Specific
        oEdit1 = oForm.Items.Item("4").Specific
        oEdit2 = oForm.Items.Item("12").Specific
        oEdit3 = oForm.Items.Item("30").Specific

        '**************Mandatory For Code*****************

        If oEdit1.Value = "" Or oEdit1.Value = Nothing Then
            oEdit1.Active = True
            objAddOn.SBO_Application.SetStatusBarMessage("Choose Vendor....")
            Return False
        End If

        '**************Mandatory For Description*****************

        'If oEdit2.Value = "" Or oEdit2.Value = Nothing Then
        '    oEdit2.Active = True
        '    objAddOn.SBO_Application.SetStatusBarMessage("Choose PO......")
        '    Return False
        'End If

        'objRS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'strSQL = "select * from opor where docnum = " & Trim(oForm.DataSources.DBDataSources.Item("@SST_OGATSUB").GetValue("U_PONum", 0)) & " and cardcode = '" & Trim(oForm.DataSources.DBDataSources.Item("@SST_OGATSUB").GetValue("U_VCode", 0)) & "'"
        'objRS.DoQuery(strSQL)
        'If objRS.RecordCount = 0 Then
        '    objAddOn.SBO_Application.SetStatusBarMessage("Invalid combination of vendor and Purchase order")
        '    Return False
        'End If
        '****************** Mandatory Category Code ***************

        'If oCombo.Selected Is Nothing Then
        '    oCombo.Active = True
        '    objAddOn.SBO_Application.SetStatusBarMessage("Category Code should not be left blank")
        '    Return False
        'Else
        '    If oCombo.Selected.Value Is "" Then
        '        oCombo.Active = True
        '        objAddOn.SBO_Application.SetStatusBarMessage("Category Code should not be left blank")
        '        Return False
        '    End If
        'End If

        '****************** Mandatory Category Description ***************


        'If oEdit3.Value = "" Or oEdit3.Value = Nothing Then
        '    oEdit3.Active = True
        '    objAddOn.SBO_Application.SetStatusBarMessage("Enter Received By.....")
        '    Return False
        'End If
        'If oMatrix.RowCount = 0 Then
        '    objAddOn.SBO_Application.SetStatusBarMessage("Line detail is missing....")
        '    Return False
        'End If
        'For i = 1 To oMatrix.RowCount
        '    If oMatrix.RowCount > 0 And oMatrix.Columns.Item("0").Cells.Item(i).Specific.string = "" Then
        '        objAddOn.SBO_Application.SetStatusBarMessage("Select Item Code....")
        '        Return False
        '    End If
        'Next
        Return True
    End Function

    Private Sub AddPOCon()
        oCFLs = oForm.ChooseFromLists
        oCFL = oCFLs.Item("CFL_4")
        oCons = oCFL.GetConditions()
        oCon = oCons.Add()
        oCon.Alias = "U_ScCode"
        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCon.CondVal = Trim(oForm.DataSources.DBDataSources.Item("@SST_OGATSUB").GetValue("U_VCode", 0))
        oCFL.SetConditions(oCons)
    End Sub

    'Private Sub AddPOCon()
    '    oCFLs = oForm.ChooseFromLists
    '    oCFL = oCFLs.Item("CFL_4")
    '    oCons = oCFL.GetConditions()
    '    oCon = oCons.Add()
    '    oCon.Alias = "DocStatus"
    '    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '    oCon.CondVal = "O"
    '    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
    '    oCon = oCons.Add()
    '    oCon.Alias = "CardCode"
    '    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '    oCon.CondVal = Trim(oForm.DataSources.DBDataSources.Item("@SST_OGATSUB").GetValue("U_VCode", 0))
    '    oCFL.SetConditions(oCons)
    'End Sub

    Private Sub AddCflCon()
        oCFLs = oForm.ChooseFromLists
        oCFL = oCFLs.Item("CFL_2")
        oCons = oCFL.GetConditions()
        oCon = oCons.Add()
        oCon.Alias = "GroupCode"
        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCon.CondVal = "104"
        oCFL.SetConditions(oCons)
    End Sub

    Private Sub Choose(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent)
        Dim rs As SAPbobsCOM.Recordset
        Dim objCFLEvent As SAPbouiCOM.ChooseFromListEvent
        Dim strCFL As String
        objRS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objCFLEvent = pval
        strCFL = objCFLEvent.ChooseFromListUID
        oForm = objAddOn.SBO_Application.Forms.Item(FormUID)
        oMatrix = oForm.Items.Item("19").Specific
        oDT = objCFLEvent.SelectedObjects
        If objCFLEvent.BeforeAction = False Then
            Try

                If strCFL = "CFL_2" Then
                    oForm.DataSources.DBDataSources.Item("@SST_OGATSUB").SetValue("U_VCode", 0, oDT.GetValue("CardCode", 0))
                    oForm.DataSources.DBDataSources.Item("@SST_OGATSUB").SetValue("U_VDesc", 0, oDT.GetValue("CardName", 0))
                    'AddPOCon()
                    LoadCombo1()

                    'ElseIf strCFL = "CFL_4" Then
                    '    Try
                    '        oForm.DataSources.DBDataSources.Item("@SST_OGATSUB").SetValue("U_Series", 0, oDT.GetValue("Series", 0))
                    '        oForm.DataSources.DBDataSources.Item("@SST_OGATSUB").SetValue("U_PONum", 0, oDT.GetValue("DocNum", 0))
                    '        'oForm.DataSources.DBDataSources.Item("@SST_OGATSUB").SetValue("U_PODate", 0, objAddOn.objGenFunc.GetDateTimeValue(oDT.GetValue("DocDate", 0)).ToString("MM/dd/yy"))
                    '        oForm.DataSources.DBDataSources.Item("@SST_OGATSUB").SetValue("U_Docentry", 0, oDT.GetValue("DocEntry", 0))
                    '        AddItemCon()

                    '    Catch ex As Exception
                    '    End Try

                    'ElseIf strCFL = "CFL_3" Then
                    '    oMatrix.GetLineData(pval.Row)
                    '    'strSQL = "select dscription,quantity,unitmsr,linenum from wOR1 where docentry = (select docentry from owor where docnum =" & oForm.DataSources.DBDataSources.Item("@SST_OGATSUB").GetValue("U_PONum", 0) & " and series =" & oForm.DataSources.DBDataSources.Item("@SST_OGATSUB").GetValue("U_Series", 0) & " ) and itemcode = '" & oDT.GetValue("ItemCode", 0) & "'"
                    '    strSQL = "select a.docentry,a.ItemCode,a.PlannedQty,b.ItemName from owor a inner join OITM b on a.itemcode = b.itemcode where a.docnum =" & oForm.DataSources.DBDataSources.Item("@SST_OGATSUB").GetValue("U_PONum", 0) & " and  a.series =" & oForm.DataSources.DBDataSources.Item("@SST_OGATSUB").GetValue("U_Series", 0) & " "
                    '    objRS.DoQuery(strSQL)
                    '    oForm.DataSources.DBDataSources.Item("@SST_SUBGAT1").SetValue("U_ItmCode", 0, oDT.GetValue("ItemCode", 0))
                    '    oForm.DataSources.DBDataSources.Item("@SST_SUBGAT1").SetValue("U_ItmDesc", 0, objRS.Fields.Item(3).Value)

                    '    rs = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '    strSQL = "select sum(a.U_Qty) Qty from [@SST_SUBGAT1] a"
                    '    strSQL = strSQL + " inner join [@SST_OGATSUB] b on a.docentry = b.docentry"
                    '    strSQL = strSQL + " where b.U_VCode = '" & Trim(oForm.DataSources.DBDataSources.Item("@SST_OGAT").GetValue("U_VCode", 0)) & "'  "
                    '    strSQL = strSQL + " and a.U_ItmCode = '" & Trim(oForm.DataSources.DBDataSources.Item("@SST_SUBGAT1").GetValue("U_ItmCode", 0)) & "' and b.U_PONum = " & Trim(oForm.DataSources.DBDataSources.Item("@SST_OGATSUB").GetValue("U_PONum", 0))
                    '    rs.DoQuery(strSQL)
                    '    If rs.RecordCount > 0 Then
                    '        oForm.DataSources.DBDataSources.Item("@SST_SUBGAT1").SetValue("U_Qty", 0, CDbl(objRS.Fields.Item(1).Value) - CDbl(rs.Fields.Item(0).Value))
                    '    Else
                    '        oForm.DataSources.DBDataSources.Item("@SST_SUBGAT1").SetValue("U_Qty", 0, objRS.Fields.Item(1).Value)
                    '    End If

                    '    oForm.DataSources.DBDataSources.Item("@SST_SUBGAT1").SetValue("U_UOM", 0, objRS.Fields.Item(2).Value)
                    '    oForm.DataSources.DBDataSources.Item("@SST_SUBGAT1").SetValue("U_BrefNo", 0, objRS.Fields.Item(3).Value)

                    '    oMatrix.SetLineData(pval.Row)
                End If

            Catch ex As Exception
                objAddOn.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End If
        objRS = Nothing
    End Sub

    Private Sub AddItemCon()
        objRS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strSQL = "select * from por1 where docentry = (select docentry from opor where docnum = " & oForm.DataSources.DBDataSources.Item("@SST_OGATSUB").GetValue("U_PONum", 0) & " and series = " & oForm.DataSources.DBDataSources.Item("@SST_OGATSUB").GetValue("U_Series", 0) & ")"
        objRS.DoQuery(strSQL)
        oCFL.SetConditions(Nothing)
        oCFLs = oForm.ChooseFromLists
        oCFL = oCFLs.Item("CFL_3")
        oCons = oCFL.GetConditions()
        oCon = oCons.Add()
        oCon.Alias = "ItemCode"
        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        For i As Integer = 1 To objRS.RecordCount
            If objRS.EoF = False Then
                oCon.CondVal = objRS.Fields.Item("ItemCode").Value
                If Not i = objRS.RecordCount Then
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCon = oCons.Add()
                    oCon.Alias = "ItemCode"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                End If
            End If
            objRS.MoveNext()
        Next
        oCFL.SetConditions(oCons)
    End Sub

    Public Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)

        If pVal.MenuUID = "1290" Or pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1291" Then
            If oForm.Items.Item("31").Specific.selected.value <> "O" Then
                Disable()
            End If
        End If
        If pVal.MenuUID = "1282" Then
            Enable()
            oForm.Items.Item("6").Specific.value = objAddOn.objGenFunc.GetDocNum("SST_GATSUB")
            oForm.Items.Item("10").Specific.string = objAddOn.objGenFunc.GetDateTimeValue(objAddOn.SBO_Application.Company.ServerDate).ToString("dd/MM/yy")
        End If
    End Sub

    Private Sub Disable()

        oForm.Items.Item("8").Enabled = False

        oForm.Items.Item("16").Enabled = False
        oForm.Items.Item("1000002").Enabled = False
        oForm.Items.Item("27").Enabled = False
        oForm.Items.Item("6").Enabled = False
        oForm.Items.Item("10").Enabled = False
        oForm.Items.Item("14").Enabled = False
        oForm.Items.Item("18").Enabled = False
        oForm.Items.Item("30").Enabled = False


    End Sub

    Private Sub Enable()
        oForm.Items.Item("4").Enabled = True
        oForm.Items.Item("12").Enabled = True
        oForm.Items.Item("16").Enabled = True
        oForm.Items.Item("1000002").Enabled = True
        oForm.Items.Item("27").Enabled = True
        oForm.Items.Item("18").Enabled = True
        oForm.Items.Item("30").Enabled = True


    End Sub

End Class
