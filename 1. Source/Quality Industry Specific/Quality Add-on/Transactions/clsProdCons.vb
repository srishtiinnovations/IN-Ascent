Public Class clsProdCons
    Private oForm As SAPbouiCOM.Form
    Private oEdit1, oEdit2 As SAPbouiCOM.EditText
    Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
    Dim oCFL As SAPbouiCOM.ChooseFromList
    Dim oCons As SAPbouiCOM.Conditions
    Dim oCon As SAPbouiCOM.Condition
    Dim RS As SAPbobsCOM.Recordset
    Dim oDT As SAPbouiCOM.DataTable
    Dim strSQL As String
    Dim oMatrix As SAPbouiCOM.Matrix
    Dim i As Integer
    Dim PONum As Integer
    Private colitemcod As SAPbouiCOM.Column
    Private colitemnam As SAPbouiCOM.Column
    Private colreceipt As SAPbouiCOM.Column
    Private colacptqty As SAPbouiCOM.Column
    Private colrejqty As SAPbouiCOM.Column
    Private colrework As SAPbouiCOM.Column
    Private colrecode As SAPbouiCOM.Column
    Private colrename As SAPbouiCOM.Column
    Private oColumns As SAPbouiCOM.Columns
    Dim oEdit As SAPbouiCOM.EditText
    Dim InsNo As String
    Public Const formtype As String = "Frm_PrdCons"
    Public Sub LoadScreen()
        oForm = objAddOn.objUIXml.LoadScreenXML("Frm_PrdConsolidate.xml", SST.enuResourceType.Embeded, formtype)
        AddCflCon()

        oEdit1 = oForm.Items.Item("txtdocode").Specific
        oEdit2 = oForm.Items.Item("txtdate").Specific
        oEdit1.Value = objAddOn.objGenFunc.GetDocNum("SST_PCONS")
        oEdit2.String = objAddOn.objGenFunc.GetDateTimeValue(objAddOn.SBO_Application.Company.ServerDate).ToString("dd/MM/yy")

        oMatrix = oForm.Items.Item("mat").Specific
        oColumns = oMatrix.Columns
        colitemcod = oColumns.Item("colitemcod")
        colitemnam = oColumns.Item("colitemnam")
        colreceipt = oColumns.Item("colreceipt")
        colacptqty = oColumns.Item("colacptqty")
        colrejqty = oColumns.Item("colrejqty")
        colrework = oColumns.Item("colrework")
        colrecode = oColumns.Item("colrecode")
        colrename = oColumns.Item("colrename")
        AddresonCombo(colrecode)
        oForm.DataBrowser.BrowseBy = "txtdocode"
    End Sub
    Private Sub AddCflCon()
        oCFLs = oForm.ChooseFromLists
        oCFL = oCFLs.Item("CFL_2")
        oCons = oCFL.GetConditions()
        oCon = oCons.Add()
        oCon.Alias = "U_status"
        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCon.CondVal = "N"
        oCon.Alias = "U_DocStat"
        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCon.CondVal = "O"
        oCFL.SetConditions(oCons)
    End Sub
    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.Before_Action = True Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        oForm = objAddOn.SBO_Application.Forms.Item(FormUID)
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            If Validate(pVal) = False Then
                                BubbleEvent = False
                            Else
                                InsNo = oForm.Items.Item("txtpono").Specific.string
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
                        If pVal.ItemUID = "txtpono" Then
                            Choose(FormUID, pVal)
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        If (pVal.ItemUID = "mat") And (pVal.ColUID = "colrecode") Then
                            Dim oEdit As SAPbouiCOM.EditText
                            Dim oCombo As SAPbouiCOM.ComboBox
                            Dim paracd As String
                            oCombo = colrecode.Cells.Item(pVal.Row).Specific
                            oEdit = colrename.Cells.Item(pVal.Row).Specific
                            paracd = oCombo.Selected.Value
                            oEdit.Value = oCombo.Selected.Description
                        End If
                End Select
            End If
            If pVal.ItemUID = "1" And pVal.Action_Success = True And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oEdit1 = oForm.Items.Item("txtdocode").Specific
                    oEdit2 = oForm.Items.Item("txtdate").Specific
                    oEdit1.Value = objAddOn.objGenFunc.GetDocNum("SST_PCONS")
                    oEdit2.String = objAddOn.objGenFunc.GetDateTimeValue(objAddOn.SBO_Application.Company.ServerDate).ToString("dd/MM/yy")
                End If
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    strSQL = "update [@SST_PRDQCHDR] set U_DocStat = 'C' where DocNum =  '" & InsNo & "'"
                    RS.DoQuery(strSQL)
                End If

            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub Choose(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent)
        Dim objCFLEvent As SAPbouiCOM.ChooseFromListEvent
        Dim strCFL As String
        RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objCFLEvent = pval
        strCFL = objCFLEvent.ChooseFromListUID
        oForm = objAddOn.SBO_Application.Forms.Item(FormUID)
        oDT = objCFLEvent.SelectedObjects
        If objCFLEvent.BeforeAction = False Then
            Try
                If strCFL = "CFL_2" Then
                    If CheckExists(oDT.GetValue("DocNum", 0)) = True Then
                        oForm.DataSources.DBDataSources.Item("@SST_PRDCONHDR").SetValue("U_prodno", 0, oDT.GetValue("DocNum", 0))
                        'oForm.Items.Item("txtpodt").Specific.string = objAddOn.objGenFunc.GetDateTimeValue(oDT.GetValue("CreateDate", 0))

                        strSQL = "select  a.U_SItemCd,a.U_sitmnme,a.U_prdqty from [@SST_PRDQCHDR] a where a.U_insno ='" & oDT.GetValue("DocNum", 0) & "'"
                        RS.DoQuery(strSQL)
                        If RS.RecordCount > 0 Then
                            If oMatrix.RowCount > 0 Then
                                oMatrix.Clear()
                            End If
                            For i = 1 To RS.RecordCount
                                oMatrix.AddRow()
                                oMatrix.GetLineData(i)
                                oForm.DataSources.DBDataSources.Item("@SST_PRDCONDTL").SetValue("LineId", 0, i)
                                oForm.DataSources.DBDataSources.Item("@SST_PRDCONDTL").SetValue("U_itemcode", 0, RS.Fields.Item("U_SItemCd").Value)
                                oForm.DataSources.DBDataSources.Item("@SST_PRDCONDTL").SetValue("U_itemname", 0, RS.Fields.Item("U_sitmnme").Value)
                                oForm.DataSources.DBDataSources.Item("@SST_PRDCONDTL").SetValue("U_recvdqty", 0, RS.Fields.Item("U_prdqty").Value)
                                oMatrix.SetLineData(i)
                                RS.MoveNext()
                            Next

                        End If
                    Else
                        objAddOn.SBO_Application.SetStatusBarMessage("Production inspection Already Exists........", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    End If
                End If
            Catch ex As Exception

            End Try
        End If
        RS = Nothing
    End Sub
    Private Function CheckExists(ByVal Dcn As Integer) As Boolean
        Try
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strSQL = "select * from [@SST_PRDQCHDR] where U_docstat = 'O' and u_insno = " & Dcn
            RS.DoQuery(strSQL)
            If RS.RecordCount > 0 Then
                Return True
            Else
                RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strSQL = "select * from [@SST_PRDCONHDR] where u_prodno = " & Dcn

                'strSQL = "select * from [@SST_PRDqcHDR] where u_insno = " & Dcn
                RS.DoQuery(strSQL)
                If RS.RecordCount > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If

        Catch ex As Exception

        End Try
    End Function
    Private Sub AddresonCombo(ByVal oColumn As SAPbouiCOM.Column)
        Try
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            RS.DoQuery("SELECT Code, U_desc FROM [dbo].[@SST_QCREASON]")
            RS.MoveFirst()
            While RS.EoF = False
                oColumn.ValidValues.Add(RS.Fields.Item("Code").Value, RS.Fields.Item("U_desc").Value)
                RS.MoveNext()
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Function Validate(ByVal pVal) As Boolean

        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            RS.DoQuery("select U_prodno from [@SST_PRDCONHDR]  where U_prodno='" & Trim(oForm.DataSources.DBDataSources.Item("@SST_PRDCONHDR").GetValue("U_prodno", 0)) & "'")
            If RS.RecordCount > 0 Then
                objAddOn.SBO_Application.SetStatusBarMessage("Production No Already Exists")
                Return False
            End If
        End If


        oForm = objAddOn.SBO_Application.Forms.Item(pVal.FormUID)

        '**************Mandatory For Doc No*****************
        oEdit = oForm.Items.Item("txtdocode").Specific
        If oEdit.Value = "" Or oEdit.Value = Nothing Then

            oEdit.Active = True
            objAddOn.SBO_Application.SetStatusBarMessage("Doc No should not be left blank")
            Return False
        End If

        '**************Mandatory For Date*****************
        oEdit1 = oForm.Items.Item("txtdate").Specific
        If oEdit1.Value = "" Or oEdit1.Value = Nothing Then

            oEdit1.Active = True
            objAddOn.SBO_Application.SetStatusBarMessage("Date should not be left blank")
            Return False
        End If

        '****************** Mandatory GRN NO ***************

        oEdit2 = oForm.Items.Item("txtpono").Specific
        If oEdit2.Value = "" Or oEdit2.Value = Nothing Then

            oEdit2.Active = True
            objAddOn.SBO_Application.SetStatusBarMessage("Production No should not be left blank")
            Return False
        End If

        '****************** Mandatory GRN Date ***************

        Dim oEdit3 As SAPbouiCOM.EditText
        oEdit3 = oForm.Items.Item("txtpodt").Specific
        If oEdit3.Value = "" Or oEdit3.Value = Nothing Then

            oEdit3.Active = True
            'objAddOn.SBO_Application.SetStatusBarMessage("Production Date should not be left blank")
            'Return False
        End If
        If oMatrix.RowCount = 0 Then
            objAddOn.SBO_Application.SetStatusBarMessage("No line Items....")
            Return False
        End If
        Try
            For i = 1 To oMatrix.RowCount

                If oMatrix.Columns.Item("colitemcod").Cells.Item(i).Specific.string <> "" Then
                    If CDbl(oMatrix.Columns.Item("colreceipt").Cells.Item(i).Specific.string) <> CDbl(oMatrix.Columns.Item("colacptqty").Cells.Item(i).Specific.string) + CDbl(oMatrix.Columns.Item("colrejqty").Cells.Item(i).Specific.string) + CDbl(oMatrix.Columns.Item("colrework").Cells.Item(i).Specific.string) Then
                        objAddOn.SBO_Application.SetStatusBarMessage("Sum of Quantities should be equal to the received qty")
                        Return False
                    End If
                Else
                    objAddOn.SBO_Application.SetStatusBarMessage("Select ItemCode....")
                    Return False
                End If

            Next
        Catch ex As Exception

        End Try
        Return True

    End Function
    Public Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        If pVal.MenuUID = "1282" Then
            oEdit1 = oForm.Items.Item("txtdocode").Specific
            oEdit2 = oForm.Items.Item("txtdate").Specific
            oEdit1.Value = objAddOn.objGenFunc.GetDocNum("SST_PCONS")
            oEdit2.String = objAddOn.objGenFunc.GetDateTimeValue(objAddOn.SBO_Application.Company.ServerDate).ToString("dd/MM/yy")
        End If
    End Sub
End Class
