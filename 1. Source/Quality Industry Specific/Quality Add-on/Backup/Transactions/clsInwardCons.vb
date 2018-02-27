
Public Class clsInwardCons
    Private oForm As SAPbouiCOM.Form
    Private oEdit1, oEdit2 As SAPbouiCOM.EditText
    Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
    Dim oCFL As SAPbouiCOM.ChooseFromList
    Dim oCons As SAPbouiCOM.Conditions
    Dim oCon As SAPbouiCOM.Condition
    Dim RS As SAPbobsCOM.Recordset
    Dim oDT As SAPbouiCOM.DataTable
    Dim strSQL As String
    Dim oMatrix, oMatrix1 As SAPbouiCOM.Matrix
    Dim i As Integer
    Dim PONum As Integer
    Private colrecode As SAPbouiCOM.Column
    Private oColumns As SAPbouiCOM.Columns
    Public Const formtype As String = "Frm_InwCons"
    Dim InsNo As String
    Dim oCmb As SAPbouiCOM.ComboBox

    Public Sub LoadScreen()
        oForm = objAddOn.objUIXml.LoadScreenXML("Frm_InwardConsdt.xml", SST.enuResourceType.Embeded, formtype)
        AddCflCon()
        oMatrix = oForm.Items.Item("mat").Specific
        oEdit1 = oForm.Items.Item("txtdocode").Specific
        oEdit2 = oForm.Items.Item("txtdate").Specific
        oColumns = oMatrix.Columns
        colrecode = oColumns.Item("colreason")
        AddresonCombo(colrecode)
        oEdit1.Value = objAddOn.objGenFunc.GetDocNum("SST_CONS")
        oEdit2.String = objAddOn.objGenFunc.GetDateTimeValue(objAddOn.SBO_Application.Company.ServerDate).ToString("dd/MM/yy")
        oForm.DataBrowser.BrowseBy = "txtdocode"
        LoadCombo()
    End Sub

    Private Sub LoadCombo()
        oCmb = oForm.Items.Item("14").Specific

        For i = oCmb.ValidValues.Count - 1 To 0 Step -1
            oCmb.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
        Next
        RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strSQL = "SELECT U_insno,docentry FROM [@SST_NQCHDR] WHERE U_DocStat = 'O'"
        RS.DoQuery(strSQL)
        While Not RS.EoF
            oCmb.ValidValues.Add(RS.Fields.Item("U_insno").Value, RS.Fields.Item("docentry").Value)
            RS.MoveNext()
        End While

    End Sub

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

    Private Sub AddCflCon()
        oCFLs = oForm.ChooseFromLists
        oCFL = oCFLs.Item("CFL_2")
        oCons = oCFL.GetConditions()
        oCon = oCons.Add()
        oCon.Alias = "U_PosStat"
        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCon.CondVal = "N"
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

                                PONum = Trim(oForm.DataSources.DBDataSources.Item("@SST_CONSHDR").GetValue("U_geno", 0))
                                InsNo = oForm.Items.Item("14").Specific.selected.value
                                InvenGRN()
                                RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                strSQL = "update [@SST_CONSHDR] set U_DocStat = 'C' where docnum = '" & oForm.Items.Item("txtdocode").Specific.string & "'"
                                RS.DoQuery(strSQL)
                            End If
                        End If
                        If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            strSQL = "select * from pdn1 where U_InsNo ='" & oForm.Items.Item("txtdocode").Specific.string & "' and U_EType = 'C'"
                            RS.DoQuery(strSQL)
                            If RS.RecordCount > 0 Then
                                objAddOn.SBO_Application.SetStatusBarMessage("Cannot be updated")
                                BubbleEvent = False
                            Else
                                If Validate(pVal) = False Then
                                    BubbleEvent = False
                                End If
                            End If

                        End If
                End Select
            Else
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        If pVal.ItemUID = "txtgenum" Then
                            Choose(FormUID, pVal)
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        If pVal.ItemUID = "14" Then
                            LoadItemCode()
                        End If
                End Select
            End If

            If pVal.ItemUID = "1" And pVal.Action_Success = True And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oEdit1 = oForm.Items.Item("txtdocode").Specific
                    oEdit2 = oForm.Items.Item("txtdate").Specific
                    oEdit1.Value = objAddOn.objGenFunc.GetDocNum("SST_CONS")
                    oEdit2.String = objAddOn.objGenFunc.GetDateTimeValue(objAddOn.SBO_Application.Company.ServerDate).ToString("dd/MM/yy")
                End If

                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    strSQL = "update [@SST_NQCHDR] set u_docstat = 'C' where docnum = '" & InsNo & "'"
                    RS.DoQuery(strSQL)

                   
                End If

            End If

        Catch ex As Exception

        End Try
       
       
    End Sub
    
    Private Function Validate(ByVal pVal) As Boolean
        oForm = objAddOn.SBO_Application.Forms.Item(pVal.FormUID)
        oMatrix = oForm.Items.Item("mat").Specific


        If oForm.Items.Item("14").Specific.value = "" Or oForm.Items.Item("14").Specific.value = Nothing Then
            objAddOn.SBO_Application.SetStatusBarMessage("Select Inward Inspection....")
            Return False
        End If

        If oMatrix.RowCount = 0 Then
            objAddOn.SBO_Application.SetStatusBarMessage("No line Items....")
            Return False
        End If

        Try
            For i = 1 To oMatrix.RowCount

                'If oMatrix.Columns.Item("colacptqty").Cells.Item(i).Specific.string = 0 Then
                '    objAddOn.SBO_Application.SetStatusBarMessage("Enter Accepted qty")
                '    Return False
                'End If

                'If oMatrix.Columns.Item("colrejqty").Cells.Item(i).Specific.string = 0 Then
                '    objAddOn.SBO_Application.SetStatusBarMessage("Enter Rejected qty")
                '    Return False
                'End If

                If CDbl(oMatrix.Columns.Item("colreceipt").Cells.Item(i).Specific.string) <> CDbl(oMatrix.Columns.Item("colacptqty").Cells.Item(i).Specific.string) + CDbl(oMatrix.Columns.Item("colrejqty").Cells.Item(i).Specific.string) + CDbl(oMatrix.Columns.Item("colrework").Cells.Item(i).Specific.string) Then
                    objAddOn.SBO_Application.SetStatusBarMessage("Sum of Quantities should be equal to the received qty")
                    Return False
                End If

                'If oMatrix.Columns.Item("colreason").Cells.Item(i).Specific.Selected.Value = Nothing Then
                '    objAddOn.SBO_Application.SetStatusBarMessage("Reason Should not be empty")
                '    Return False
                'End If

            Next

            If oForm.Items.Item("15").Specific.value = "" Or oForm.Items.Item("15").Specific.value = Nothing Then
                objAddOn.SBO_Application.SetStatusBarMessage("Remarks Should not be empty....")
                Return False
            End If

        Catch ex As Exception

        End Try

        Return True
    End Function

    Private Sub Choose(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent)
        Dim objCFLEvent As SAPbouiCOM.ChooseFromListEvent
        Dim strCFL As String
        RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objCFLEvent = pVal
        strCFL = objCFLEvent.ChooseFromListUID
        oForm = objAddOn.SBO_Application.Forms.Item(FormUID)
        oDT = objCFLEvent.SelectedObjects
        If objCFLEvent.BeforeAction = False Then
            Try
                If strCFL = "CFL_2" Then
                    oForm.DataSources.DBDataSources.Item("@SST_CONSHDR").SetValue("U_geno", 0, oDT.GetValue("DocNum", 0))
                    oForm.Items.Item("txtgrdt").Specific.string = objAddOn.objGenFunc.GetDateTimeValue(oDT.GetValue("CreateDate", 0))
                    LoadInspNo()
                    'strSQL = "select a.* from [@SST_GAT1] a "
                    'strSQL = strSQL + " inner join [@SST_NQCHDR] b on a. docentry = b.u_geno and a.u_itmcode = b.U_itemcode"
                    'strSQL = strSQL + " where  a.u_grnno is null and a.docentry = " & oDT.GetValue("DocEntry", 0)
                    'RS.DoQuery(strSQL)
                    'If RS.RecordCount > 0 Then
                    '    If oMatrix.RowCount > 0 Then
                    '        oMatrix.Clear()
                    '    End If
                    '    For i = 1 To RS.RecordCount
                    '        oMatrix.AddRow()
                    '        oMatrix.GetLineData(i)
                    '        oForm.DataSources.DBDataSources.Item("@SST_CONSDTL").SetValue("LineId", 0, i)
                    '        oForm.DataSources.DBDataSources.Item("@SST_CONSDTL").SetValue("U_itemcode", 0, RS.Fields.Item("U_ItmCode").Value)
                    '        oForm.DataSources.DBDataSources.Item("@SST_CONSDTL").SetValue("U_itemname", 0, RS.Fields.Item("U_ItmDesc").Value)
                    '        oForm.DataSources.DBDataSources.Item("@SST_CONSDTL").SetValue("U_recvdqty", 0, RS.Fields.Item("U_Qty").Value)
                    '        oMatrix.SetLineData(i)
                    '        RS.MoveNext()
                    '    Next
                    'End If
                End If
            Catch ex As Exception

            End Try
        End If
        RS = Nothing
    End Sub

    Private Sub LoadInspNo()
        Dim oCombo As SAPbouiCOM.ComboBox
        oCombo = oForm.Items.Item("14").Specific
        Try
            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
            Next

            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strSQL = "select docnum from [@SST_NQCHDR] where U_DocStat = 'O' and U_GENo = " & oForm.DataSources.DBDataSources.Item("@SST_CONSHDR").GetValue("U_geno", 0) & " "
            'strSQL = "select docnum from [@SST_NQCHDR] where U_DocStat = 'O' and U_GENo = " & oForm.DataSources.DBDataSources.Item("@SST_CONSHDR").GetValue("U_geno", 0) & " and U_Status = 'N'"
            strSQL = strSQL + " and docnum not in (select U_insno from [@SST_CONSHDR] where U_GENo = " & oForm.DataSources.DBDataSources.Item("@SST_CONSHDR").GetValue("U_geno", 0) & ")"
            RS.DoQuery(strSQL)

            While RS.EoF = False
                oCombo.ValidValues.Add(RS.Fields.Item("docnum").Value, RS.Fields.Item("docnum").Value)
                RS.MoveNext()
            End While
         
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub LoadItemCode()
       
        Try
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'strSQL = "select U_itemcode,U_itemname,U_GEqty  from [@SST_NQCHDR] where U_DocStat = 'O' and U_GENo = " & oForm.DataSources.DBDataSources.Item("@SST_CONSHDR").GetValue("U_geno", 0) & " and DocNum = " & oForm.Items.Item("14").Specific.selected.value
            strSQL = "select U_itemcode,U_itemname,U_GEqty,U_GENo  from [@SST_NQCHDR] where U_DocStat = 'O' and DocNum = " & oForm.Items.Item("14").Specific.selected.value
            RS.DoQuery(strSQL)
            If RS.RecordCount > 0 Then
                oForm.Items.Item("txtgenum").Specific.Value = RS.Fields.Item("U_GENo").Value
                If oMatrix.RowCount > 0 Then
                    oMatrix.Clear()
                End If
                For i = 1 To RS.RecordCount
                    oMatrix.AddRow()
                    oMatrix.GetLineData(i)
                    oForm.DataSources.DBDataSources.Item("@SST_CONSDTL").SetValue("LineId", 0, i)
                    oForm.DataSources.DBDataSources.Item("@SST_CONSDTL").SetValue("U_itemcode", 0, RS.Fields.Item("U_itemcode").Value)
                    oForm.DataSources.DBDataSources.Item("@SST_CONSDTL").SetValue("U_itemname", 0, RS.Fields.Item("U_itemname").Value)
                    oForm.DataSources.DBDataSources.Item("@SST_CONSDTL").SetValue("U_recvdqty", 0, RS.Fields.Item("U_GEqty").Value)
                    oMatrix.SetLineData(i)
                    RS.MoveNext()
                Next
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        If pVal.MenuUID = "1282" Then
            oEdit1 = oForm.Items.Item("txtdocode").Specific
            oEdit2 = oForm.Items.Item("txtdate").Specific
            oEdit1.Value = objAddOn.objGenFunc.GetDocNum("SST_CONS")
            oEdit2.String = objAddOn.objGenFunc.GetDateTimeValue(objAddOn.SBO_Application.Company.ServerDate).ToString("dd/MM/yy")
            oForm.Items.Item("14").Enabled = True
            oForm.Items.Item("txtgenum").Enabled = True
            oForm.Items.Item("15").Enabled = True
            oMatrix = oForm.Items.Item("mat").Specific
            oMatrix.Columns.Item("colacptqty").Editable = True
            oMatrix.Columns.Item("colrejqty").Editable = True
            oMatrix.Columns.Item("colrework").Editable = True
            oMatrix.Columns.Item("V_0").Editable = True
            oMatrix.Columns.Item("colreason").Editable = True
            oForm.Items.Item("17").Enabled = True
        End If
        If pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291" Then
            If oForm.Items.Item("17").Specific.selected.value = "C" Then
                oForm.Items.Item("14").Enabled = False
                oForm.Items.Item("17").Enabled = False
                oForm.Items.Item("txtgenum").Enabled = False
                oForm.Items.Item("15").Enabled = False
                oMatrix = oForm.Items.Item("mat").Specific
                oMatrix.Columns.Item("colacptqty").Editable = False
                oMatrix.Columns.Item("colrejqty").Editable = False
                oMatrix.Columns.Item("colrework").Editable = False
                oMatrix.Columns.Item("V_0").Editable = False
                oMatrix.Columns.Item("colreason").Editable = False
            Else
                oForm.Items.Item("14").Enabled = True
                oForm.Items.Item("17").Enabled = True
                oForm.Items.Item("txtgenum").Enabled = True
                oForm.Items.Item("15").Enabled = True
                oMatrix = oForm.Items.Item("mat").Specific
                oMatrix.Columns.Item("colacptqty").Editable = True
                oMatrix.Columns.Item("colrejqty").Editable = True
                oMatrix.Columns.Item("colrework").Editable = True
                oMatrix.Columns.Item("V_0").Editable = True
                oMatrix.Columns.Item("colreason").Editable = True
            End If
        End If
    End Sub

    Private Function validate() As Boolean
        oMatrix = oForm.Items.Item("23").Specific
        For i = 1 To oMatrix.RowCount

        Next
        Return True
    End Function

    Private Function InvenGRN() As Boolean
        Try
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oIN As SAPbobsCOM.StockTransfer
            Dim icode, idesc, Excisable, wcode As String
            oIN = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)

            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strSQL = "select ItemCode,Dscription,Quantity,TaxCode,WhsCode,Excisable from PDN1 where DocEntry  = '" & oForm.Items.Item("txtgenum").Specific.string & "'"
            RS.DoQuery(strSQL)
            If RS.RecordCount > 0 Then

                icode = RS.Fields.Item(0).Value
                idesc = RS.Fields.Item(1).Value
                wcode = RS.Fields.Item(5).Value
                Excisable = RS.Fields.Item(5).Value
                ' MsgBox(Excisable)

                oIN.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer
                'oIN.FromWarehouse = "STR"

                If Excisable = "Y" Then
                    oIN.FromWarehouse = "IWD"

                    oMatrix = oForm.Items.Item("mat").Specific
                    Dim acceptedqty, rejectedqty As Integer

                    For i = 1 To oMatrix.RowCount
                        acceptedqty = CDbl(oMatrix.Columns.Item("colacptqty").Cells.Item(i).Specific.String) + CDbl(oMatrix.Columns.Item("colrework").Cells.Item(i).Specific.String)
                        ' MsgBox(acceptedqty)
                        rejectedqty = oMatrix.Columns.Item("colrejqty").Cells.Item(i).Specific.String
                        ' MsgBox(rejectedqty)
                    Next

                    For i = 0 To oIN.Lines.Count - 1
                        oIN.Lines.SetCurrentLine(i)
                        If acceptedqty > 0 Then
                            oIN.Lines.ItemCode = icode
                            oIN.Lines.ItemDescription = idesc
                            oIN.Lines.Quantity = acceptedqty
                            oIN.Lines.WarehouseCode = "STR"
                        End If
                    Next

                    oIN.Lines.Add()

                    For i = 1 To oIN.Lines.Count - 1
                        oIN.Lines.SetCurrentLine(i)
                        If rejectedqty > 0 Then
                            oIN.Lines.ItemCode = icode
                            oIN.Lines.ItemDescription = idesc
                            oIN.Lines.Quantity = rejectedqty
                            oIN.Lines.WarehouseCode = "IRW"
                        End If
                    Next

                    oIN.Lines.Add()

                    If oIN.Add() <> 0 Then
                        MsgBox(objAddOn.oCompany.GetLastErrorDescription)
                        Return False
                    Else
                        RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        strSQL = "select docnum from owtr where docentry = " & objAddOn.oCompany.GetNewObjectKey
                        RS.DoQuery(strSQL)
                        'Dim msg As String = RS.Fields.Item(0).Value
                        MessageBox.Show("Stock Transfer successfully created.DocNum : " & RS.Fields.Item(0).Value)
                        MessageBox.Show("Create Corresponding Excise Invoice..")

                        RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        strSQL = "update wtr1 set Excisable = 'Y' where docentry  = " & objAddOn.oCompany.GetNewObjectKey
                        RS.DoQuery(strSQL)
                    End If

                Else
                    oIN.FromWarehouse = "IWD1"
                    oMatrix = oForm.Items.Item("mat").Specific
                    Dim acceptedqty, rejectedqty As Integer

                    For i = 1 To oMatrix.RowCount
                        acceptedqty = CDbl(oMatrix.Columns.Item("colacptqty").Cells.Item(i).Specific.String) + CDbl(oMatrix.Columns.Item("colrework").Cells.Item(i).Specific.String)
                        rejectedqty = oMatrix.Columns.Item("colrejqty").Cells.Item(i).Specific.String
                    Next

                    For i = 0 To oIN.Lines.Count - 1
                        oIN.Lines.SetCurrentLine(i)
                        If acceptedqty > 0 Then
                            oIN.Lines.ItemCode = icode
                            oIN.Lines.ItemDescription = idesc
                            oIN.Lines.Quantity = acceptedqty
                            oIN.Lines.WarehouseCode = "STR1"
                        End If
                    Next

                    oIN.Lines.Add()

                    For i = 1 To oIN.Lines.Count - 1
                        oIN.Lines.SetCurrentLine(i)
                        If rejectedqty > 0 Then
                            oIN.Lines.ItemCode = icode
                            oIN.Lines.ItemDescription = idesc
                            oIN.Lines.Quantity = rejectedqty
                            oIN.Lines.WarehouseCode = "IRW1"
                        End If
                    Next

                    oIN.Lines.Add()

                    If oIN.Add() <> 0 Then
                        MsgBox(objAddOn.oCompany.GetLastErrorDescription)
                        Return False
                    Else
                        RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        strSQL = "select docnum from owtr where docentry = " & objAddOn.oCompany.GetNewObjectKey
                        RS.DoQuery(strSQL)
                        MessageBox.Show("Stock Transfer successfully created.DocNum : " & RS.Fields.Item(0).Value)
                        'Dim msg As String = RS.Fields.Item(0).Value
                        'MsgBox(msg)
                    End If

                End If

            End If

        Catch ex As Exception
            Return False
        End Try

    End Function
    
End Class
