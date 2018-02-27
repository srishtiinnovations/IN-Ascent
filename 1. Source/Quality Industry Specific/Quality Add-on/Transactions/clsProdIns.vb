Public Class clsProdIns
    Private oForm As SAPbouiCOM.Form
    Public Const formtype As String = "Frm_PrdInsp"
    Dim oDT As SAPbouiCOM.DataTable
    Private oMatrix As SAPbouiCOM.Matrix
    Dim RS As SAPbobsCOM.Recordset
    Dim strSQL As String
    Dim i As Integer
    Dim k As Integer = 0
    Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
    Dim oCFL As SAPbouiCOM.ChooseFromList
    Dim oCons As SAPbouiCOM.Conditions
    Dim oCon As SAPbouiCOM.Condition
    Dim oCombo, oCombo1 As SAPbouiCOM.ComboBox
    Dim oCol1, oCol2, oCol3 As SAPbouiCOM.Column
    Dim oColumns As SAPbouiCOM.Columns
    Public Sub LoadScreen()
        Try

            oForm = objAddOn.objUIXml.LoadScreenXML("Frm_PrdInspEntry.xml", SST.enuResourceType.Embeded, formtype)
            'oForm.Items.Item("txtsft").Enabled = False
            AddCflCon()
            LoadCombo()
            Dim ofld As SAPbouiCOM.Folder
            'ofld = oForm.Items.Item("1000001").Specific
            'ofld.Select()
            oMatrix = oForm.Items.Item("mat2").Specific
            oColumns = oMatrix.Columns
            'oCol1 = oColumns.Item("tollp")
            'oCol2 = oColumns.Item("tollm")
            'oCol3 = oColumns.Item("actual")
            'oCol3.Editable = True
            oForm.DataBrowser.BrowseBy = "txtinsno"
            oForm.Items.Item("txtinsno").Specific.value = objAddOn.objGenFunc.GetDocNum("SST_PRDINSP")
            oForm.Items.Item("53").Specific.string = objAddOn.objGenFunc.GetDateTimeValue(objAddOn.SBO_Application.Company.ServerDate).ToString("dd/MM/yy")
        Catch ex As Exception
        End Try
    End Sub
    Private Sub AddCflCon()
        oCFLs = oForm.ChooseFromLists
        oCFL = oCFLs.Item("CFL_2")
        oCons = oCFL.GetConditions()
        oCon = oCons.Add()
        oCon.Alias = "Status"
        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCon.CondVal = "R"
        oCFL.SetConditions(oCons)
    End Sub
    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)

        Try
            If pVal.Before_Action = True Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        oForm = objAddOn.SBO_Application.Forms.Item(pVal.FormUID)

                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        'If pVal.ItemUID = "1000001" Then
                        '    oForm.PaneLevel = 1
                        'End If
                        If pVal.ItemUID = "1000002" Then
                            oForm.PaneLevel = 2
                        End If

                        If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            If Validate(pVal) = False Then
                                BubbleEvent = False
                            End If
                        End If
                        If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            'objAddOn.SBO_Application.SetStatusBarMessage("Cannot be updated")
                            'BubbleEvent = False
                            If Validate(pVal) = False Then
                                BubbleEvent = False
                            End If
                        End If

                End Select

            Else

                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        If pVal.ItemUID = "txtgrnno" Then
                            Choose(FormUID, pVal)
                        End If
                        'modified ------------S
                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        If pVal.ItemUID = "1000008" Then
                            If CheckExists(oForm.Items.Item("txtgrnno").Specific.string, oForm.Items.Item("1000008").Specific.selected.description) = False Then

                                RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                strSQL = "select itemname from oitm where itemcode = '" & oForm.Items.Item("1000008").Specific.selected.description & "'"
                                RS.DoQuery(strSQL)
                                If RS.RecordCount > 0 Then
                                    oForm.Items.Item("1000012").Specific.string = RS.Fields.Item(0).Value
                                End If

                                LoadStage(oForm.Items.Item("1000008").Specific.selected.description)
                                RS = Nothing
                                RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                strSQL = ""
                                strSQL = "select sum(plannedqty) plannedqty from wor1 where itemcode = '" & oForm.Items.Item("1000008").Specific.selected.description & "' and  docentry = (select docentry from owor where docnum = " & oForm.Items.Item("txtgrnno").Specific.string & ")"
                                strSQL = strSQL + " group by itemcode"
                                RS.DoQuery(strSQL)
                                If RS.RecordCount > 0 Then
                                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCHDR").SetValue("U_prdqty", 0, RS.Fields.Item(0).Value)
                                End If
                                oForm.Freeze(True)
                                LoadMatrix()
                                oForm.Freeze(False)
                                oForm.Items.Item("60").Specific.string = oForm.Items.Item("1000008").Specific.selected.description
                            Else
                                objAddOn.SBO_Application.SetStatusBarMessage("Production Order and Sub Level Item combination Already Exists........", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                        End If
                        'modified------------E
                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        If pVal.ItemUID = "mat2" And pVal.ColUID = "actual" Then
                            If CDbl(oCol1.Cells.Item(pVal.Row).Specific.string) > CDbl(oCol2.Cells.Item(pVal.Row).Specific.string) Then
                                If CDbl(oCol1.Cells.Item(pVal.Row).Specific.string) >= CDbl(oCol3.Cells.Item(pVal.Row).Specific.string) And CDbl(oCol2.Cells.Item(pVal.Row).Specific.string) <= CDbl(oCol3.Cells.Item(pVal.Row).Specific.string) Then
                                    'oMatrix.Columns.Item("status").Cells.Item(pVal.Row).Specific.string = "A"

                                    'modified ------------S
                                    oMatrix.Columns.Item("status").Cells.Item(pVal.Row).Specific.string = "Accepted"
                                    'modified------------E
                                Else
                                    'oMatrix.Columns.Item("status").Cells.Item(pVal.Row).Specific.string = "N"
                                    oMatrix.Columns.Item("status").Cells.Item(pVal.Row).Specific.string = "Rejection"
                                    'modified------------E
                                End If
                            End If
                            If CDbl(oCol2.Cells.Item(pVal.Row).Specific.string) > CDbl(oCol1.Cells.Item(pVal.Row).Specific.string) Then
                                If CDbl(oCol1.Cells.Item(pVal.Row).Specific.string) <= CDbl(oCol3.Cells.Item(pVal.Row).Specific.string) And CDbl(oCol2.Cells.Item(pVal.Row).Specific.string) >= CDbl(oCol3.Cells.Item(pVal.Row).Specific.string) Then
                                    'oMatrix.Columns.Item("status").Cells.Item(pVal.Row).Specific.string = "A"

                                    'modified ------------S
                                    oMatrix.Columns.Item("status").Cells.Item(pVal.Row).Specific.string = "Accepted"
                                    'modified------------E
                                Else
                                    'oMatrix.Columns.Item("status").Cells.Item(pVal.Row).Specific.string = "N"

                                    'modified ------------S
                                    oMatrix.Columns.Item("status").Cells.Item(pVal.Row).Specific.string = "Rejection"
                                    'modified------------E
                                End If
                            End If

                        End If
                    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        If pVal.ItemUID = "58" Then
                            If CDbl(oForm.Items.Item("58").Specific.string) <= CDbl(oForm.Items.Item("txtsupcd").Specific.string) Then
                                LoadMatrix_Ins(oForm.Items.Item("58").Specific.string)
                            Else
                                objAddOn.SBO_Application.SetStatusBarMessage("Inspection Qty should not be greater than the production qty....", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                        End If
                End Select
            End If

            If pVal.ItemUID = "1" And pVal.Action_Success = True And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Items.Item("txtinsno").Specific.value = objAddOn.objGenFunc.GetDocNum("SST_PRDINSP")
                    oForm.Items.Item("53").Specific.string = objAddOn.objGenFunc.GetDateTimeValue(objAddOn.SBO_Application.Company.ServerDate).ToString("dd/MM/yy")
                End If

            End If
        Catch ex As Exception

        End Try

    End Sub
    Private Sub LoadMatrix_Ins(ByVal Ins As String)
        RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strSQL = ""
        strSQL = "select distinct t1.*,'" & Ins & "' as u_smpsize from [@SST_PRDSTDDTL] t1 "
        strSQL = strSQL + " inner join [@SST_PRDSTDHDR] t2 on t2.code = t1.code"
        strSQL = strSQL + " inner join [@SST_PLANHDR] t3 on  t3.U_itemcode = t2.u_itemcode and t3.U_stage = t2.U_Stage"
        strSQL = strSQL + " inner join [@SST_PLANDTL] t4 on  t4.code = t3.code"
        strSQL = strSQL + " where t2.u_itemcode = '" & oForm.Items.Item("1000008").Specific.selected.description & "' "
        strSQL = strSQL + " and t2.U_Active = 'Y' and t3.U_Active = 'Y'"
        RS.DoQuery(strSQL)
        If RS.RecordCount > 0 Then
            If oMatrix.RowCount > 0 Then
                oMatrix.Clear()
            End If
            k = 0
            For i = 1 To RS.RecordCount
                oMatrix.AddRow()
                oMatrix.GetLineData(oMatrix.RowCount)
                oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("LineId", 0, oMatrix.RowCount)
                oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_pcode", 0, RS.Fields.Item("U_paracode").Value)
                oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_paradesc", 0, RS.Fields.Item("U_paradesc").Value)
                oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_uomcode", 0, RS.Fields.Item("U_uomcode").Value)
                oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_uom", 0, RS.Fields.Item("U_uomdesc").Value)
                oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_value", 0, RS.Fields.Item("U_value").Value)
                oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_tolladd", 0, RS.Fields.Item("U_tollplus").Value)
                oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_tollsub", 0, RS.Fields.Item("U_tollmins").Value)
                oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_samno", 0, 1)
                oMatrix.SetLineData(oMatrix.RowCount)
                For k = 1 To RS.Fields.Item("u_smpsize").Value - 1
                    oMatrix.AddRow()
                    oMatrix.GetLineData(oMatrix.RowCount)
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("LineId", 0, oMatrix.RowCount)
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_pcode", 0, RS.Fields.Item("U_paracode").Value)
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_paradesc", 0, RS.Fields.Item("U_paradesc").Value)
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_uomcode", 0, RS.Fields.Item("U_uomcode").Value)
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_uom", 0, RS.Fields.Item("U_uomdesc").Value)
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_value", 0, RS.Fields.Item("U_value").Value)
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_tolladd", 0, RS.Fields.Item("U_tollplus").Value)
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_tollsub", 0, RS.Fields.Item("U_tollmins").Value)
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_samno", 0, k + 1)
                    oMatrix.SetLineData(oMatrix.RowCount)
                Next
                RS.MoveNext()
            Next
        Else
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strSQL = ""
            strSQL = "select distinct t1.* from [@SST_PRDSTDDTL] t1  "
            strSQL = strSQL + " inner join [@SST_PRDSTDHDR] t2 on t2.code = t1.code "
            strSQL = strSQL + " where t2.u_itemcode = '" & oForm.Items.Item("1000008").Specific.selected.description & "'"
            strSQL = strSQL + " and t2.U_Active = 'Y' "
            RS.DoQuery(strSQL)
            If RS.RecordCount > 0 Then
                If oMatrix.RowCount > 0 Then
                    oMatrix.Clear()
                End If

                k = 0
                For i = 1 To RS.RecordCount
                    oMatrix.AddRow()
                    oMatrix.GetLineData(oMatrix.RowCount)
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("LineId", 0, oMatrix.RowCount)
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_pcode", 0, RS.Fields.Item("U_paracode").Value)
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_paradesc", 0, RS.Fields.Item("U_paradesc").Value)
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_uomcode", 0, RS.Fields.Item("U_uomcode").Value)
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_uom", 0, RS.Fields.Item("U_uomdesc").Value)
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_value", 0, RS.Fields.Item("U_value").Value)
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_tolladd", 0, RS.Fields.Item("U_tollplus").Value)
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_tollsub", 0, RS.Fields.Item("U_tollmins").Value)
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_samno", 0, 1)
                    oMatrix.SetLineData(oMatrix.RowCount)
                    For k = 1 To Trim(oForm.DataSources.DBDataSources.Item("@SST_PRDQCHDR").GetValue("U_Insqty", 0)) - 1
                        oMatrix.AddRow()
                        oMatrix.GetLineData(oMatrix.RowCount)
                        oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("LineId", 0, oMatrix.RowCount)
                        oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_pcode", 0, RS.Fields.Item("U_paracode").Value)
                        oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_paradesc", 0, RS.Fields.Item("U_paradesc").Value)
                        oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_uomcode", 0, RS.Fields.Item("U_uomcode").Value)
                        oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_uom", 0, RS.Fields.Item("U_uomdesc").Value)
                        oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_value", 0, RS.Fields.Item("U_value").Value)
                        oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_tolladd", 0, RS.Fields.Item("U_tollplus").Value)
                        oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_tollsub", 0, RS.Fields.Item("U_tollmins").Value)
                        oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_samno", 0, k + 1)
                        oMatrix.SetLineData(oMatrix.RowCount)
                    Next
                    RS.MoveNext()
                Next
            End If
        End If
    End Sub
    Private Sub LoadCombo()
        Try
            oCombo = oForm.Items.Item("1000005").Specific
            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
            oCombo.ValidValues.Add("Normal", "Normal")
            oCombo.ValidValues.Add("Rework", "Rework")
            oCombo.Select("Normal", SAPbouiCOM.BoSearchKey.psk_ByValue)
        Catch ex As Exception

        End Try

    End Sub
    Private Sub LoadStage(ByVal ItmCode As String)
        Try

            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strSQL = "select * from [@SST_PRDSTDHDR] where U_itemcode = '" & ItmCode & "'"
            RS.DoQuery(strSQL)
            If RS.RecordCount > 0 Then
                oForm.Items.Item("1000006").Specific.string = RS.Fields.Item("U_stagedes").Value.ToString
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Choose(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent)
        Dim objCFLEvent As SAPbouiCOM.ChooseFromListEvent
        Dim strCFL As String
        Dim k As Integer = 0
        RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objCFLEvent = pval
        strCFL = objCFLEvent.ChooseFromListUID
        oForm = objAddOn.SBO_Application.Forms.Item(FormUID)
        oDT = objCFLEvent.SelectedObjects

        If objCFLEvent.BeforeAction = False Then
            Try
                If strCFL = "CFL_2" Then
                    'If CheckExists(oDT.GetValue("DocNum", 0)) = False Then

                    '******* for checking purpose i comment the if condition = true'''''

                    ' If IsIssued(oDT.GetValue("DocNum", 0)) = True Then
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCHDR").SetValue("U_prodno", 0, oDT.GetValue("DocNum", 0))

                    'oForm.Items.Item("54").Specific.string = objAddOn.objGenFunc.GetDateTimeValue(oDT.GetValue("CreateDate", 0))
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCHDR").SetValue("U_itemcode", 0, oDT.GetValue("ItemCode", 0))
                    strSQL = "select * from oitm where itemcode = '" & oDT.GetValue("ItemCode", 0) & "'"
                    RS.DoQuery(strSQL)
                    If RS.RecordCount > 0 Then
                        oForm.DataSources.DBDataSources.Item("@SST_PRDQCHDR").SetValue("U_itemname", 0, RS.Fields.Item("ItemName").Value)
                    End If
                    'oForm.DataSources.DBDataSources.Item("@SST_PRDQCHDR").SetValue("U_prdqty", 0, oDT.GetValue("PlannedQty", 0))
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCHDR").SetValue("U_uom", 0, oDT.GetValue("Uom", 0))
                    Try
                        oCombo = oForm.Items.Item("1000008").Specific
                        For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                            oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                        Next
                        RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'strSQL = "select itemcode from wor1 where docentry = " & oForm.Items.Item("txtgrnno").Specific.value
                        strSQL = "select itemcode from WOR1 where docentry =(select DocEntry from OWOR where DocNum = " & oForm.Items.Item("txtgrnno").Specific.value & ")"
                        strSQL = strSQL + " group by itemcode"
                        RS.DoQuery(strSQL)
                        While Not RS.EoF
                            k = k + 1
                            oCombo.ValidValues.Add(k, RS.Fields.Item("itemcode").Value)
                            RS.MoveNext()
                        End While

                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try
                Else
                    objAddOn.SBO_Application.SetStatusBarMessage("Item Not Issued........", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    'End If
                    'Else
                    '    objAddOn.SBO_Application.SetStatusBarMessage("Production Order Already Exists........", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    'End If
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End If

        RS = Nothing
    End Sub
    Private Function CheckExists(ByVal Dcn As Integer, ByVal ItmCd As String) As Boolean
        Try
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strSQL = "select * from [@SST_PRDQCHDR] where U_sitemcd = '" & ItmCd & "' and U_prodno = " & Dcn
            RS.DoQuery(strSQL)
            If RS.RecordCount > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception

        End Try
    End Function
    Private Function IsIssued(ByVal DocN As Integer) As Boolean
        Try
            Dim rec As SAPbobsCOM.Recordset
            RS = Nothing
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strSQL = "select issuetype,itemcode from wor1 where docentry = (select docentry from owor where docnum = " & DocN & ")"
            RS.DoQuery(strSQL)
            While Not RS.EoF
                If RS.Fields.Item(0).Value = "M" Then
                    rec = Nothing
                    rec = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    strSQL = "select * from ige1 where baseref = " & DocN & " and itemcode = '" & RS.Fields.Item(1).Value & "'"
                    ' strSQL = "select * from ige1 where itemcode = '" & RS.Fields.Item(1).Value & "' "
                    rec.DoQuery(strSQL)
                    If rec.RecordCount = 0 Then
                        Return False
                    End If
                End If
                RS.MoveNext()
            End While

            Return True
        Catch ex As Exception

        End Try
    End Function
    Private Function Validate(ByVal pVal) As Boolean
        oForm = objAddOn.SBO_Application.Forms.Item(pVal.FormUID)
        oMatrix = oForm.Items.Item("mat2").Specific

        If oForm.Items.Item("txtgrnno").Specific.value = "" Or oForm.Items.Item("txtgrnno").Specific.value = Nothing Then

            objAddOn.SBO_Application.SetStatusBarMessage("Choose Production Order....")
            Return False
        End If

        If oMatrix.RowCount = 0 Then
            objAddOn.SBO_Application.SetStatusBarMessage("No line Items....")
            Return False
        End If
        Try
            For i = 1 To oMatrix.RowCount

                If oMatrix.Columns.Item("actual").Cells.Item(i).Specific.string <> "" Then
                    If CDbl(oMatrix.Columns.Item("actual").Cells.Item(i).Specific.string) <> 0 Then
                        'If oMatrix.Columns.Item("status").Cells.Item(i).Specific.string = "N" Then
                        If oMatrix.Columns.Item("status").Cells.Item(i).Specific.string = "Rejection" Then
                            oForm.Items.Item("txtsft").Specific.string = "N"
                            Exit For
                        Else
                            oForm.Items.Item("txtsft").Specific.string = ""
                            oForm.Items.Item("txtsft").Specific.string = "A"
                        End If
                    Else
                        objAddOn.SBO_Application.SetStatusBarMessage("Enter Actual Values....")
                        Return False
                    End If
                Else
                    objAddOn.SBO_Application.SetStatusBarMessage("Enter Actual Values....")
                    Return False
                End If
            Next


        Catch ex As Exception

        End Try

        Return True
    End Function
    Private Sub LoadMatrix()
        Try
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strSQL = ""
            'strSQL = "select distinct t1.*,t4.u_smpsize from [@SST_PRDSTDDTL] t1 "
            'strSQL = strSQL + " inner join [@SST_PRDSTDHDR] t2 on t2.code = t1.code"
            'strSQL = strSQL + " inner join [@SST_PLANHDR] t3 on  t3.U_itemcode = t2.u_itemcode and t3.U_stage = t2.U_Stage"
            'strSQL = strSQL + " inner join [@SST_PLANDTL] t4 on  t4.code = t3.code"
            ''strSQL = strSQL + " where t2.u_itemcode = '" & Trim(oForm.DataSources.DBDataSources.Item("@SST_PRDQCHDR").GetValue("U_itemcode", 0)) & "' "
            'strSQL = strSQL + " where t2.u_itemcode = '" & oForm.Items.Item("1000008").Specific.selected.description & "' "
            'strSQL = strSQL + " and t4.u_fromqty <= " & Trim(oForm.DataSources.DBDataSources.Item("@SST_PRDQCHDR").GetValue("U_prdqty", 0)) & " and t4.u_toqty >= " & Trim(oForm.DataSources.DBDataSources.Item("@SST_PRDQCHDR").GetValue("U_prdqty", 0)) & ""
            'strSQL = strSQL + " and t2.U_Active = 'Y' and t3.U_Active = 'Y'"
            'RS.DoQuery(strSQL)

            strSQL = "select a.U_itemcode,b.U_paradesc,b.U_smpsize,b.U_accpqty,U_rejqty"
            strSQL = strSQL + " from [@SST_planhdr] a inner join [@SST_plandtl] b on a.Code = b.Code"
            strSQL = strSQL + " where a.U_itemcode = '" & oForm.Items.Item("txtitem").Specific.String & "' "
            'strSQL = strSQL + " and a.U_Active = 'Y' "
            RS.DoQuery(strSQL)
            If RS.RecordCount > 0 Then
                If oMatrix.RowCount > 0 Then
                    oMatrix.Clear()
                End If
                k = 0
                For i = 1 To RS.RecordCount
                    oMatrix.AddRow()
                    oMatrix.GetLineData(oMatrix.RowCount)
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("LineId", 0, oMatrix.RowCount)
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_paradesc", 0, RS.Fields.Item("U_paradesc").Value)
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_smpsize", 0, RS.Fields.Item("U_smpsize").Value)
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_acclvl", 0, RS.Fields.Item("U_accpqty").Value)
                    oForm.DataSources.DBDataSources.Item("@SST_PRDQCDTL").SetValue("U_rejlvl", 0, RS.Fields.Item("U_rejqty").Value)
                    oMatrix.SetLineData(oMatrix.RowCount)
                Next
                RS.MoveNext()

            Else
                objAddOn.SBO_Application.SetStatusBarMessage("No matching records....", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If

        Catch ex As Exception
            oForm.Freeze(False)
        End Try
    End Sub
    Public Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        If pVal.MenuUID = "1282" Then
            oForm.Items.Item("txtinsno").Specific.value = objAddOn.objGenFunc.GetDocNum("SST_PRDINSP")
            oForm.Items.Item("53").Specific.string = objAddOn.objGenFunc.GetDateTimeValue(objAddOn.SBO_Application.Company.ServerDate).ToString("dd/MM/yy")
            oForm.Items.Item("59").Enabled = True
            oForm.Items.Item("txtgrnno").Enabled = True
            oForm.Items.Item("1000008").Enabled = True
            oForm.Items.Item("1000005").Enabled = True
            oForm.Items.Item("46").Enabled = True
            oForm.Items.Item("txtbom").Enabled = True
            oForm.Items.Item("txtrwk").Enabled = True
            oForm.Items.Item("1000015").Enabled = True
            oForm.Items.Item("57").Enabled = True
            oForm.Items.Item("txtsno").Enabled = True
            oForm.Items.Item("txtrfd").Enabled = True
            oForm.Items.Item("1000010").Enabled = True
            oForm.Items.Item("txtrsno").Enabled = True
            oForm.Items.Item("txtecn").Enabled = True
            oForm.Items.Item("txtasstim").Enabled = True
            oForm.Items.Item("txtrmrks").Enabled = True
            oForm.Items.Item("58").Enabled = True
            oMatrix = oForm.Items.Item("mat2").Specific
            oMatrix.Columns.Item("actual").Editable = True
        ElseIf pVal.MenuUID = "1289" Or pVal.MenuUID = "1291" Or pVal.MenuUID = "1288" Or pVal.MenuUID = "1290" Then
            If oForm.Items.Item("59").Specific.selected.value = "C" Then
                oForm.Items.Item("59").Enabled = False
                oForm.Items.Item("txtgrnno").Enabled = False
                oForm.Items.Item("1000008").Enabled = False
                oForm.Items.Item("1000005").Enabled = False
                oForm.Items.Item("46").Enabled = False
                oForm.Items.Item("txtbom").Enabled = False
                oForm.Items.Item("txtrwk").Enabled = False
                oForm.Items.Item("1000015").Enabled = False
                oForm.Items.Item("57").Enabled = False
                oForm.Items.Item("txtsno").Enabled = False
                oForm.Items.Item("txtrfd").Enabled = False
                oForm.Items.Item("1000010").Enabled = False
                oForm.Items.Item("txtrsno").Enabled = False
                oForm.Items.Item("txtecn").Enabled = False
                oForm.Items.Item("txtasstim").Enabled = False
                oForm.Items.Item("txtrmrks").Enabled = False
                oForm.Items.Item("58").Enabled = False
                oMatrix = oForm.Items.Item("mat2").Specific
                oMatrix.Columns.Item("actual").Editable = False
            Else
                oForm.Items.Item("59").Enabled = True
                oForm.Items.Item("txtgrnno").Enabled = True
                oForm.Items.Item("1000008").Enabled = True
                oForm.Items.Item("1000005").Enabled = True
                oForm.Items.Item("46").Enabled = True
                oForm.Items.Item("txtbom").Enabled = True
                oForm.Items.Item("txtrwk").Enabled = True
                oForm.Items.Item("1000015").Enabled = True
                oForm.Items.Item("57").Enabled = True
                oForm.Items.Item("txtsno").Enabled = True
                oForm.Items.Item("txtrfd").Enabled = True
                oForm.Items.Item("1000010").Enabled = True
                oForm.Items.Item("txtrsno").Enabled = True
                oForm.Items.Item("txtecn").Enabled = True
                oForm.Items.Item("txtasstim").Enabled = True
                oForm.Items.Item("txtrmrks").Enabled = True
                oForm.Items.Item("txtrmrks").Enabled = True
                oMatrix = oForm.Items.Item("mat2").Specific
                oMatrix.Columns.Item("actual").Editable = True
            End If
        End If

    End Sub

End Class
