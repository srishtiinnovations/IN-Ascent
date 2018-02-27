Public Class clsUserCFL
    Private oForm As SAPbouiCOM.Form
    Private oGrid As SAPbouiCOM.Grid
    Private oMatrix As SAPbouiCOM.Matrix
    Dim RS As SAPbobsCOM.Recordset
    Dim oDT As SAPbouiCOM.DataTable
    Dim strSQL As String
    Dim GEno, docen As Int64
    Dim ConsEntry As String
    Public Const formtype As String = "Frm_CFL"

    Public Sub LoadScreen(ByVal VenCode As String)
        oForm = objAddOn.objUIXml.LoadScreenXML("Frm_UserCFL.xml", SST.enuResourceType.Embeded, formtype)
        LoadGrid(oForm.UniqueID, VenCode)
        Try
            oDT = oForm.DataSources.DataTables.Item("CFL")
        Catch ex As Exception
            oDT = oForm.DataSources.DataTables.Add("CFL")
        End Try
        oDT.Columns.Add("Qty", SAPbouiCOM.BoFieldsType.ft_Quantity)
        oDT.Columns.Add("WhsCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 5)
        oDT.Columns.Add("itemcode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50)
        oDT.Columns.Add("InsNo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 10)
        oDT.Columns.Add("rea", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 15)
        oDT.Columns.Add("EType", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 10)
        oDT.Columns.Add("UP", SAPbouiCOM.BoFieldsType.ft_Price)
        oDT.Columns.Add("TC", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 15)

        oForm.Visible = True
    End Sub

    Public Sub ItemEvent(ByVal FormUID As String, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pval.BeforeAction Then

            Else
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pval.ItemUID = "101" Then
                            LoadDatatable(FormUID)
                            LoadItemMatrix(FormUID)
                            oForm.Close()
                        ElseIf pval.ItemUID = "2" Then
                            oForm.Close()
                        End If
                End Select
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub LoadGrid(ByVal FormUID As String, ByVal Vendor As String)
        Try
            oForm = objAddOn.SBO_Application.Forms.Item(FormUID)

            oGrid = oForm.Items.Item("3").Specific
            oForm.DataSources.DataTables.Add("DT")
            strSQL = "Select U_DocNum 'GE_No',U_VCode 'Vendorcode',u_vdesc as 'VendorName',U_PONum 'PO_No',U_DocDate 'GE_Date'FROM [@SST_OGAT] WHERE U_Status = 'O'"

            'strSQL = "select a.DocNum as 'Inspection No',a.U_GENo as 'GateEntry No',a.U_itemcode as 'ItemCode','Inward Entry' as Type "
            'strSQL = strSQL + " from [@SST_NQCHDR] a  "
            ''strSQL = strSQL + " inner join [@SST_OGAT] c on c.docnum = a.U_geno  where a.U_Status = 'A' and a.u_supcode = '" & Vendor & "' and  "
            'strSQL = strSQL + " inner join [@SST_OGAT] c on c.docnum = a.U_geno  where a.u_supcode = '" & Vendor & "' and  "
            'strSQL = strSQL + " c.U_Status = 'O' and a.U_DocStat = 'O'"
            'strSQL = strSQL + " and a.docnum not in (select u_insno from pdn1 where u_insno is not null and u_EType = 'I' group by u_insno)"
            'strSQL = strSQL + " group by a.Docnum,a.U_GENo,a.U_itemcode"
            'strSQL = strSQL + " Union All"
            'strSQL = strSQL + " select a.DocNum as 'Inspection No',a.U_GENo as 'GateEntry No', b.U_itemcode as 'ItemCode','Consolidate Entry' as Type "
            'strSQL = strSQL + " from [@SST_CONSHDR] a "
            'strSQL = strSQL + " inner join [@SST_CONSDTL] b on b.docentry = a.docentry"
            'strSQL = strSQL + " inner join [@SST_OGAT] c on c.docnum = a.U_geno"
            'strSQL = strSQL + " where c.U_VCode = '" & Vendor & "' and   c.U_Status = 'O' and a.U_DocStat = 'O' and a.docnum not in (select u_insno from pdn1 where u_insno is not null and u_EType = 'C' "
            'strSQL = strSQL + " group by u_insno)"
            'strSQL = strSQL + " group by a.U_GENo,a.Docnum,b.U_itemcode"

            oGrid.DataTable = oForm.DataSources.DataTables.Item("DT")

            oGrid.DataTable.ExecuteQuery(strSQL)

        Catch ex As Exception

        End Try

    End Sub
    'Private Sub LoadDatatable(ByVal FormUID As String)
    '    Dim i As Integer = 0
    '    Dim k As Integer = 0
    '    Dim Whscode As String = ""
    '    Dim RWhscode As String = ""
    '    Dim RJWhscode As String = ""
    '    Dim RWWhscode As String = ""


    '    oForm = objAddOn.SBO_Application.Forms.Item(FormUID)
    '    oGrid = oForm.Items.Item("3").Specific
    '    If oGrid.Rows.Count > 0 Then
    '        For i = 0 To oGrid.Rows.Count - 1
    '            If oGrid.Rows.IsSelected(i) = True Then
    '                oDT.Rows.Clear()
    '                If oGrid.DataTable.GetValue("Type", i) = "Inward Entry" Then
    '                    ConsEntry = "N"
    '                    oDT.Rows.Add()
    '                    RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                    strSQL = "select * from [@SST_SETUP]"
    '                    RS.DoQuery(strSQL)
    '                    If RS.RecordCount > 0 Then
    '                        Whscode = RS.Fields.Item("U_RGLWH").Value
    '                    Else
    '                        objAddOn.SBO_Application.SetStatusBarMessage("No Warehouse details found.....", SAPbouiCOM.BoMessageTime.bmt_Short, True)
    '                    End If
    '                    RS = Nothing
    '                    RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                    strSQL = "select a.Price,a.TaxCode, a.docentry,"
    '                    strSQL = strSQL + " (select b.U_Qty from [@SST_OGAT] a"
    '                    strSQL = strSQL + " inner join [@SST_GAT1] b on a.docentry = b.docentry"
    '                    strSQL = strSQL + " where b.u_itmcode = '" & oGrid.DataTable.GetValue("ItemCode", i) & "' and a.docnum = " & oGrid.DataTable.GetValue("GateEntry No", i) & ") U_Qty "
    '                    strSQL = strSQL + " from por1 a "
    '                    strSQL = strSQL + " inner join OPOR b on b.DocEntry = a.docentry"
    '                    strSQL = strSQL + " where a.itemcode = '" & oGrid.DataTable.GetValue("ItemCode", i) & "' and b.DocNum = (select U_PONum  from [@SST_OGAT] where docnum = " & oGrid.DataTable.GetValue("GateEntry No", i) & ")"
    '                    RS.DoQuery(strSQL)
    '                    If RS.RecordCount > 0 Then
    '                        oDT.Columns.Item("itemcode").Cells.Item(0).Value = oGrid.DataTable.GetValue("ItemCode", i)
    '                        oDT.Columns.Item("WhsCode").Cells.Item(0).Value = Whscode
    '                        oDT.Columns.Item("Qty").Cells.Item(0).Value = CStr(RS.Fields.Item("U_Qty").Value)
    '                        oDT.Columns.Item("InsNo").Cells.Item(0).Value = CStr(oGrid.DataTable.GetValue("Inspection No", i))
    '                        oDT.Columns.Item("EType").Cells.Item(0).Value = "I"
    '                        GEno = oGrid.DataTable.GetValue("GateEntry No", i)
    '                        docen = RS.Fields.Item("docentry").Value
    '                        oDT.Columns.Item("UP").Cells.Item(0).Value = CStr(RS.Fields.Item("Price").Value)
    '                        oDT.Columns.Item("TC").Cells.Item(0).Value = CStr(RS.Fields.Item("TaxCode").Value)
    '                    End If
    '                Else
    '                    ConsEntry = "Y"
    '                    RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                    strSQL = "select * from [@SST_SETUP]"
    '                    RS.DoQuery(strSQL)
    '                    If RS.RecordCount > 0 Then
    '                        RWhscode = RS.Fields.Item("U_RGLWH").Value
    '                        RJWhscode = RS.Fields.Item("U_RJTWH").Value
    '                        RWWhscode = RS.Fields.Item("U_RWKWH").Value
    '                    Else
    '                        objAddOn.SBO_Application.SetStatusBarMessage("No Warehouse details found.....", SAPbouiCOM.BoMessageTime.bmt_Short, True)
    '                    End If

    '                    RS = Nothing
    '                    RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                    'strSQL = "select u_itemcode,U_acptqty + ISNULL(U_AccDev,0) U_acptqty,U_rejctqty,U_rwkqty,U_reason from [@SST_CONSDTL] where docentry = " & oGrid.DataTable.GetValue("Inspection No", i)
    '                    strSQL = "select a.u_itemcode,a.U_acptqty + ISNULL(a.U_AccDev,0) U_acptqty,a.U_rejctqty,a.U_rwkqty,a.U_reason,d.Price,d.TaxCode,e.docentry  from "
    '                    strSQL = strSQL + " [@SST_CONSDTL] a inner join [@SST_CONSHDR] b on a.DocEntry = b.DocEntry "
    '                    strSQL = strSQL + " inner join [@SST_OGAT] c on b.U_geno = c.DocNum "
    '                    strSQL = strSQL + " inner join OPOR e on c.U_PONum = e.docnum"
    '                    strSQL = strSQL + " inner join POR1 d on e.docentry = d.docentry"
    '                    strSQL = strSQL + " where a.docentry = " & oGrid.DataTable.GetValue("Inspection No", i) & "  and d.ItemCode = '" & oGrid.DataTable.GetValue("ItemCode", i) & "'"
    '                    strSQL = strSQL + " and b.u_geno = " & oGrid.DataTable.GetValue("GateEntry No", i)
    '                    RS.DoQuery(strSQL)
    '                    For k = 0 To RS.RecordCount - 1
    '                        If CDbl(RS.Fields.Item("U_acptqty").Value) <> 0 Then
    '                            oDT.Rows.Add()
    '                            oDT.Columns.Item("itemcode").Cells.Item(oDT.Rows.Count - 1).Value = CStr(RS.Fields.Item("u_itemcode").Value)
    '                            oDT.Columns.Item("WhsCode").Cells.Item(oDT.Rows.Count - 1).Value = RWhscode
    '                            oDT.Columns.Item("Qty").Cells.Item(oDT.Rows.Count - 1).Value = CStr(RS.Fields.Item("U_acptqty").Value)
    '                            oDT.Columns.Item("InsNo").Cells.Item(oDT.Rows.Count - 1).Value = CStr(oGrid.DataTable.GetValue("Inspection No", i))
    '                            oDT.Columns.Item("rea").Cells.Item(oDT.Rows.Count - 1).Value = ""
    '                            oDT.Columns.Item("EType").Cells.Item(oDT.Rows.Count - 1).Value = "C"
    '                            oDT.Columns.Item("UP").Cells.Item(oDT.Rows.Count - 1).Value = CStr(RS.Fields.Item("Price").Value)
    '                            oDT.Columns.Item("TC").Cells.Item(oDT.Rows.Count - 1).Value = CStr(RS.Fields.Item("TaxCode").Value)
    '                            GEno = oGrid.DataTable.GetValue("GateEntry No", i)
    '                            docen = RS.Fields.Item("docentry").Value

    '                        End If
    '                        If CDbl(RS.Fields.Item("U_rejctqty").Value) <> 0 Then
    '                            oDT.Rows.Add()
    '                            oDT.Columns.Item("itemcode").Cells.Item(oDT.Rows.Count - 1).Value = CStr(RS.Fields.Item("u_itemcode").Value)
    '                            oDT.Columns.Item("WhsCode").Cells.Item(oDT.Rows.Count - 1).Value = RJWhscode
    '                            oDT.Columns.Item("Qty").Cells.Item(oDT.Rows.Count - 1).Value = CStr(RS.Fields.Item("U_rejctqty").Value)
    '                            oDT.Columns.Item("InsNo").Cells.Item(oDT.Rows.Count - 1).Value = CStr(oGrid.DataTable.GetValue("Inspection No", i))
    '                            oDT.Columns.Item("rea").Cells.Item(oDT.Rows.Count - 1).Value = CStr(RS.Fields.Item("U_reason").Value)
    '                            oDT.Columns.Item("EType").Cells.Item(oDT.Rows.Count - 1).Value = "C"
    '                            oDT.Columns.Item("UP").Cells.Item(oDT.Rows.Count - 1).Value = CStr(RS.Fields.Item("Price").Value)
    '                            oDT.Columns.Item("TC").Cells.Item(oDT.Rows.Count - 1).Value = CStr(RS.Fields.Item("TaxCode").Value)
    '                            GEno = oGrid.DataTable.GetValue("GateEntry No", i)
    '                            docen = RS.Fields.Item("docentry").Value
    '                        End If
    '                        If CDbl(RS.Fields.Item("U_rwkqty").Value) <> 0 Then
    '                            oDT.Rows.Add()
    '                            oDT.Columns.Item("itemcode").Cells.Item(oDT.Rows.Count - 1).Value = CStr(RS.Fields.Item("u_itemcode").Value)
    '                            'oDT.Columns.Item("WhsCode").Cells.Item(oDT.Rows.Count - 1).Value = RWWhscode
    '                            oDT.Columns.Item("WhsCode").Cells.Item(oDT.Rows.Count - 1).Value = RJWhscode
    '                            oDT.Columns.Item("Qty").Cells.Item(oDT.Rows.Count - 1).Value = CStr(RS.Fields.Item("U_rwkqty").Value)
    '                            oDT.Columns.Item("InsNo").Cells.Item(oDT.Rows.Count - 1).Value = CStr(oGrid.DataTable.GetValue("Inspection No", i))
    '                            oDT.Columns.Item("rea").Cells.Item(oDT.Rows.Count - 1).Value = ""
    '                            oDT.Columns.Item("UP").Cells.Item(oDT.Rows.Count - 1).Value = CStr(RS.Fields.Item("Price").Value)
    '                            oDT.Columns.Item("TC").Cells.Item(oDT.Rows.Count - 1).Value = CStr(RS.Fields.Item("TaxCode").Value)
    '                            oDT.Columns.Item("EType").Cells.Item(oDT.Rows.Count - 1).Value = "C"
    '                            GEno = oGrid.DataTable.GetValue("GateEntry No", i)
    '                            docen = RS.Fields.Item("docentry").Value
    '                        End If
    '                        RS.MoveNext()
    '                    Next
    '                End If
    '            End If
    '        Next
    '    End If

    'End Sub

    Private Sub LoadDatatable(ByVal FormUID As String)
        Dim i As Integer = 0
        Dim k As Integer = 0
        Dim Whscode As String = ""
        Dim RWhscode As String = ""
        Dim RJWhscode As String = ""
        Dim RWWhscode As String = ""


        oForm = objAddOn.SBO_Application.Forms.Item(FormUID)
        oGrid = oForm.Items.Item("3").Specific
        If oGrid.Rows.Count > 0 Then
            For i = 0 To oGrid.Rows.Count - 1
                If oGrid.Rows.IsSelected(i) = True Then
                    oDT.Rows.Clear()
                    oDT.Rows.Add()
                    RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    RS = Nothing
                    RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    strSQL = "select a.U_VCode,U_VDesc,b.U_ItmCode,b.U_ItmDesc,b.u_qty,a.U_Status,b.U_whs,U_tcode,U_price"
                    strSQL = strSQL + " from [@SST_OGAT] a"
                    strSQL = strSQL + " inner join [@SST_GAT1] b on a.U_DocNum = b.DocEntry"
                    strSQL = strSQL + " where a.U_VCode = '" & oGrid.DataTable.GetValue("Vendorcode", i) & "' and a.U_Status = 'O' and a.U_DocNum = '" & oGrid.DataTable.GetValue("GE_No", i) & "' "
                    RS.DoQuery(strSQL)
                    If RS.RecordCount > 0 Then
                        oDT.Columns.Item("itemcode").Cells.Item(0).Value = RS.Fields.Item("U_ItmCode").Value
                        oDT.Columns.Item("Qty").Cells.Item(0).Value = CStr(RS.Fields.Item("u_qty").Value)
                        'oDT.Columns.Item("InsNo").Cells.Item(0).Value = CStr(oGrid.DataTable.GetValue("Inspection No", i))
                        'oDT.Columns.Item("EType").Cells.Item(0).Value = "I"
                        GEno = oGrid.DataTable.GetValue("GE_No", i)
                        'docen = RS.Fields.Item("docentry").Value
                        oDT.Columns.Item("WhsCode").Cells.Item(0).Value = CStr(RS.Fields.Item("U_whs").Value)
                        'MsgBox(RS.Fields.Item("U_price").Value)
                        oDT.Columns.Item("UP").Cells.Item(0).Value = RS.Fields.Item("U_price").Value
                        oDT.Columns.Item("TC").Cells.Item(0).Value = CStr(RS.Fields.Item("U_tcode").Value)
                    End If
                End If
            Next
        End If

    End Sub
    Private Sub LoadItemMatrix(ByVal FormUID As String)

        Dim GRPOForm As SAPbouiCOM.Form
        Dim p As Integer
        GRPOForm = objAddOn.GRPO.ReturnForm
        oMatrix = GRPOForm.Items.Item("38").Specific
        If oMatrix.RowCount = 1 Then
            'oMatrix.Columns.Item("11").Editable = True
            'oMatrix.Columns.Item("24").Editable = True
            For p = 0 To oDT.Rows.Count - 1
                oMatrix.Columns.Item("1").Cells.Item(oMatrix.RowCount).Specific.string = oDT.Columns.Item("itemcode").Cells.Item(p).Value
                oMatrix.Columns.Item("11").Cells.Item(oMatrix.RowCount - 1).Specific.string = oDT.Columns.Item("Qty").Cells.Item(p).Value
                oMatrix.Columns.Item("24").Cells.Item(oMatrix.RowCount - 1).Specific.string = CStr(oDT.Columns.Item("WhsCode").Cells.Item(p).Value)
                oMatrix.Columns.Item("14").Cells.Item(oMatrix.RowCount - 1).Specific.string = CStr(oDT.Columns.Item("UP").Cells.Item(p).Value)
                oMatrix.Columns.Item("160").Cells.Item(oMatrix.RowCount - 1).Specific.string = CStr(oDT.Columns.Item("TC").Cells.Item(p).Value)
                oMatrix.Columns.Item("U_InsNo").Cells.Item(oMatrix.RowCount - 1).Specific.string = oDT.Columns.Item("InsNo").Cells.Item(p).Value
                oMatrix.Columns.Item("U_Res").Cells.Item(oMatrix.RowCount - 1).Specific.string = CStr(oDT.Columns.Item("rea").Cells.Item(p).Value)
                oMatrix.Columns.Item("U_EType").Cells.Item(oMatrix.RowCount - 1).Specific.string = CStr(oDT.Columns.Item("EType").Cells.Item(p).Value)

            Next
        Else
            objAddOn.SBO_Application.SetStatusBarMessage("Delete the lineitems and then proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End If

        RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strSQL = ""
        ' strSQL = "select a.*,b.DocNum from por1 a inner join opor b on b.docentry = a.docentry where a.docentry = (select docentry from opor where docnum = (select u_ponum from [@SST_OGAT] where docnum =" & GEno & "))"
        strSQL = "select a.*,b.DocNum from por1 a inner join opor b on b.docentry = a.docentry where a.docentry = " & docen & " "
        RS.DoQuery(strSQL)
        If RS.RecordCount > 0 Then
            GRPOForm.Items.Item("16").Specific.string = "Based On Purchase Orders " & RS.Fields.Item("DocNum").Value
        End If

        GRPOForm.Update()
        objAddOn.GRPO.GetNos(GEno, ConsEntry, RS.Fields.Item("DocNum").Value)
    End Sub
End Class
