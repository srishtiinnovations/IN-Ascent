Public Class clsUserCFL1
    Private oForm As SAPbouiCOM.Form
    Private oGrid As SAPbouiCOM.Grid
    Private oMatrix As SAPbouiCOM.Matrix
    Dim RS As SAPbobsCOM.Recordset
    Dim oDT As SAPbouiCOM.DataTable
    Dim strSQL As String
    Dim GEno, docen As Int64
    Dim ConsEntry As String
    Public Const formtype As String = "Frm_CFL1"

    Public Sub LoadScreen(ByVal VenCode As String)
        oForm = objAddOn.objUIXml.LoadScreenXML("Frm_UserCFL1.xml", SST.enuResourceType.Embeded, formtype)
        LoadGrid(oForm.UniqueID, VenCode)
        Try
            oDT = oForm.DataSources.DataTables.Item("CFL1")
        Catch ex As Exception
            oDT = oForm.DataSources.DataTables.Add("CFL1")
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
            strSQL = "Select U_DocNum 'GE_No',U_VCode 'Vendorcode',u_vdesc as 'VendorName',U_PONum 'PO_No',U_DocDate 'GE_Date'FROM [@SST_OGATSUB] WHERE U_Status = 'O'"

            oGrid.DataTable = oForm.DataSources.DataTables.Item("DT")

            oGrid.DataTable.ExecuteQuery(strSQL)

        Catch ex As Exception

        End Try

    End Sub
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

                    strSQL = "select a.U_VCode,U_VDesc,b.U_ItmCode,b.U_ItmDesc,b.u_qty,a.U_Status"
                    strSQL = strSQL + " from [@SST_OGATSUB] a"
                    strSQL = strSQL + " inner join [@SST_SUBGAT1] b on a.U_DocNum = b.DocEntry"
                    strSQL = strSQL + " where a.U_VCode = '" & oGrid.DataTable.GetValue("Vendorcode", i) & "' and a.U_Status = 'O' and a.U_DocNum = '" & oGrid.DataTable.GetValue("GE_No", i) & "' "
                    RS.DoQuery(strSQL)
                    If RS.RecordCount > 0 Then
                        oDT.Columns.Item("itemcode").Cells.Item(0).Value = RS.Fields.Item("U_ItmCode").Value
                        oDT.Columns.Item("Qty").Cells.Item(0).Value = CStr(RS.Fields.Item("u_qty").Value)
                        GEno = oGrid.DataTable.GetValue("GE_No", i)
                        'oDT.Columns.Item("WhsCode").Cells.Item(0).Value = CStr(RS.Fields.Item("U_whs").Value)
                        'oDT.Columns.Item("UP").Cells.Item(0).Value = RS.Fields.Item("U_price").Value
                        'oDT.Columns.Item("TC").Cells.Item(0).Value = CStr(RS.Fields.Item("U_tcode").Value)
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
