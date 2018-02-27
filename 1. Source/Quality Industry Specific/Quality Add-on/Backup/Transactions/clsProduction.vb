
Public Class clsproduction

    Public Const formtype As String = "65211"
    Private oForm As SAPbouiCOM.Form
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oCompany As SAPbobsCOM.Company
    Private GENum As Int64
    Private RS As SAPbobsCOM.Recordset
    Private oMatrix As SAPbouiCOM.Matrix
    Private strSQL As String
    Dim GRNNo, GEQty, qty As Int64
    Dim DocNum, Linenum, DocEntry, PONum, baselinenum, docen, POQuantity, grpoquantity, quan, pdn1qty, linenumber, openqty As Int64
    Dim ItmCode, VCode, CnsEnt, ActItem, ItemCode, InsNo, DocType, icode, docseries, U_Gateentry, Gateentry, Series, doc, OpenCreQty, por1qty, podocentry, Gateentrynumber As String
    Dim oPO As SAPbobsCOM.Documents

    Private Sub DrawForm(ByVal FormUID)
        Try
            oForm = objAddOn.SBO_Application.Forms.Item(FormUID)
            oForm.Freeze(True)
            InitializeFormComponent(FormUID)
            oForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub InitializeFormComponent(ByVal FormUID)
        Try

            Dim objItem As SAPbouiCOM.Item
            Dim BtnTax, BtnTax1 As SAPbouiCOM.Button

            oForm = objAddOn.SBO_Application.Forms.Item(FormUID)
            oForm.Freeze(True)
            objItem = oForm.Items.Add("BtnGRN1", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem.Width = 120
            objItem.Top = oForm.Items.Item("1").Top
            objItem.Visible = True
            objItem.Left = (oForm.Width - (oForm.Width / 2) - 50)
            objItem.Height = 20
            BtnTax = oForm.Items.Item("BtnGRN1").Specific
            BtnTax.Caption = "Copy from Concolidated "
            oForm.Refresh()
            oForm.Freeze(False)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Dim m As Integer
        ActItem = "Test"
        Try
            If pVal.FormType = 65211 Then
                If pVal.Before_Action = True Then
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            DrawForm(FormUID)
                            oForm = objAddOn.SBO_Application.Forms.Item(FormUID)
                            oMatrix = oForm.Items.Item("37").Specific
                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                            'If pVal.ItemUID = "1" Then
                            '    GRNNo = oForm.Items.Item("8").Specific.string
                            '    VCode = oForm.Items.Item("4").Specific.string
                            '    'docnum = oForm.Items.Item("8").Specific.string
                            '    docseries = oForm.Items.Item("88").Specific.selected.value

                            '    ItmCode = oMatrix.Columns.Item("1").Cells.Item(oMatrix.RowCount - 1).Specific.string
                            '    InsNo = oMatrix.Columns.Item("U_InsNo").Cells.Item(oMatrix.RowCount - 1).Specific.string
                            '    DocType = oMatrix.Columns.Item("U_EType").Cells.Item(oMatrix.RowCount - 1).Specific.string
                            'End If


                        Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                            If pVal.ItemUID = "37" Then
                                'If pVal.ColUID = "11" Then
                                '    ItmCode = oMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.string
                                '    RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                '    strSQL = "select QryGroup6 from OITM where itemcode = '" & ItmCode & "'"
                                '    RS.DoQuery(strSQL)
                                '    If RS.RecordCount > 0 Then
                                '        If RS.Fields.Item(0).Value <> "Y" Then
                                '            objAddOn.SBO_Application.SetStatusBarMessage("Cannot Edit")
                                '            BubbleEvent = False
                                '        End If
                                '    Else
                                '        objAddOn.SBO_Application.SetStatusBarMessage("Cannot Edit")
                                '        BubbleEvent = False
                                '    End If

                                'End If
                                'If pVal.ColUID = "24" Or pVal.ColUID = "U_InsNo" Or pVal.ColUID = "U_EType" Then
                                '    objAddOn.SBO_Application.SetStatusBarMessage("Cannot Edit")
                                '    BubbleEvent = False
                                'End If
                            End If
                    End Select
                Else
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                            If pVal.ItemUID = "BtnGRN" Then

                                If oForm.Items.Item("6").Specific.string <> "" Then
                                    objAddOn.CFL.LoadScreen(oForm.Items.Item("6").Specific.string)
                                Else
                                    objAddOn.SBO_Application.SetStatusBarMessage("Select Product No ", SAPbouiCOM.BoMessageTime.bmt_Short)
                                End If
                            End If
                            If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                                RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                strSQL = "select docentry from opdn where docnum = '" & GRNNo & "' and series = '" & docseries & "' "
                                RS.DoQuery(strSQL)
                                Updation(RS.Fields.Item("docentry").Value)

                            End If
                            If pVal.ItemUID = "1" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                                MsgBox("success")
                            End If
                    End Select
                End If

            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Function ReturnForm() As SAPbouiCOM.Form
        oForm = objAddOn.SBO_Application.Forms.GetForm(clsproduction.formtype, 1)
        Return oForm
    End Function

    'last Modified on 27.09.2011

    Private Sub Updation(ByVal docentry As String)
        Try
            'PONum = docentry
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strSQL = "select sum(t2.u_qty) qty,t1.docentry from [@SST_OGAT] t1"
            strSQL = strSQL + " inner join [@SST_GAT1] t2 on t1.docentry = t2.docentry"
            strSQL = strSQL + " where t1.U_PONum =" & PONum & " and t1.U_VCode = '" & VCode & "'"
            strSQL = strSQL + " and t2.U_ItmCode = '" & ItmCode & "' "
            strSQL = strSQL + " group by t1.docentry "
            RS.DoQuery(strSQL)
            If RS.RecordCount > 0 Then
                GEQty = CDbl(RS.Fields.Item(0).Value)
                Gateentrynumber = RS.Fields.Item("docentry").Value
            End If
            Try

                RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strSQL = ""
                strSQL = "select a.docentry as podocentry,a.linenum,b.DocNum,c.U_PONum,c.U_Docentry,c.docentry as Gateentry,d.u_brefno,d.U_ItmCode,b.Series,"
                strSQL = strSQL + " sum(a.Quantity) as Quantity, a.OpenQty, a.OpenCreQty"
                strSQL = strSQL + " from por1 a "
                strSQL = strSQL + " inner join opor b on b.docentry = a.docentry  "
                strSQL = strSQL + " inner join [@sst_ogat] c on a.DocEntry = c.U_Docentry"
                strSQL = strSQL + " inner join [@sst_gat1] d on c.DocEntry = d.docentry"
                strSQL = strSQL + " where a.itemcode = '" & ItmCode & "' and a.docentry = (select distinct u_docentry  from [@SST_OGAT] where u_ponum = " & PONum & " )"
                strSQL = strSQL + " and  c.docentry = '" & Gateentrynumber & " '"
                strSQL = strSQL + " group by a.docentry ,a.linenum,b.DocNum,c.U_PONum,c.U_Docentry,c.docentry,d.u_brefno,d.U_ItmCode,b.Series,"
                strSQL = strSQL + " a.OpenQty, a.OpenCreQty"
                RS.DoQuery(strSQL)

                If RS.RecordCount > 0 Then
                    'qty = RS.Fields.Item("Quantity").Value
                    podocentry = RS.Fields.Item("podocentry").Value
                    Linenum = RS.Fields.Item("linenum").Value
                    docen = RS.Fields.Item("U_Docentry").Value
                    DocNum = RS.Fields.Item("DocNum").Value
                    baselinenum = RS.Fields.Item("u_brefno").Value
                    icode = RS.Fields.Item("U_ItmCode").Value
                    Gateentry = RS.Fields.Item("Gateentry").Value
                    Series = RS.Fields.Item("Series").Value
                    POQuantity = RS.Fields.Item("Quantity").Value
                    OpenCreQty = RS.Fields.Item("OpenCreQty").Value
                End If

            Catch ex As Exception

            End Try

            If GEQty >= qty Then
                RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If Gateentry <> "" Then
                    Dim strSQL1 As String
                    strSQL1 = "update pdn1 set  BaseEntry = " & docen & ", BaseLine = " & baselinenum & ", BaseType = 22"
                    strSQL1 = strSQL1 + " where  docentry = " & docentry & " and  ItemCode = '" & icode & "' " ' and linenum = " & baselinenum & " "
                    RS.DoQuery(strSQL1)

                    'strSQL = "select  a.docentry as podocentry,a.linenum,b.DocNum,b.Series,sum(a.Quantity) as Quantity, a.OpenQty, a.OpenCreQty"
                    'strSQL = strSQL + " from POR1 a inner join OPOR b on a.DocEntry = b.DocEntry where a.DocEntry in "
                    'strSQL = strSQL + " (select distinct u_docentry  from [@SST_OGAT] where u_ponum = " & Gateentrynumber & ") and a.ItemCode = '" & ItmCode & "' "
                    'strSQL = strSQL + " group by a.docentry,a.linenum,b.DocNum,b.Series,a.OpenQty, a.OpenCreQty"
                    'RS.DoQuery(strSQL)
                    Dim strSQL10 As String
                    strSQL10 = "Select  SUM(quantity) as grpoquantity"
                    strSQL10 = strSQL10 + " from pdn1 a inner join OPOR b on a.BaseEntry = b.DocEntry"
                    strSQL10 = strSQL10 + " where a.docentry = " & docentry & " and a.ItemCode = '" & icode & "' "
                    RS.DoQuery(strSQL10)
                    If RS.RecordCount > 0 Then
                        grpoquantity = RS.Fields.Item("grpoquantity").Value
                    End If

                    Try
                        If grpoquantity >= POQuantity And docentry <> "0" Then

                            Dim a As Integer
                            a = OpenCreQty - grpoquantity

                            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim strSQL3 As String
                            strSQL3 = "update por1 set OpenCreQty = " & a & " ,linestatus = 'C' where docentry = " & podocentry & " and linenum = '" & Linenum & "'"
                            RS.DoQuery(strSQL3)

                            'RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'Dim strSQL4 As String
                            'strSQL4 = "select a.docentry,x.docnum,a.BaseEntry , b.U_Docentry,b.Series,c.U_Qty,sum(c.U_Qty) as grpoquty"
                            'strSQL4 = strSQL4 + " from  opdn x "
                            'strSQL4 = strSQL4 + " inner join PDN1 a on x.DocEntry = a.DocEntry"
                            'strSQL4 = strSQL4 + " inner join [@SST_OGAT] b on a.BaseEntry = b.U_Docentry"
                            'strSQL4 = strSQL4 + " inner join [@sst_gat1] c on b.DocEntry = c.docentry "
                            'strSQL4 = strSQL4 + " where x.docnum = " & c & " "
                            'strSQL4 = strSQL4 + " group by a.docentry,x.docnum,a.BaseEntry , b.U_Docentry,b.Series,c.U_Qty"
                            'RS.DoQuery(strSQL4)
                            'doc = RS.Fields.Item("U_Docentry").Value
                            'MsgBox(doc)
                            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim strSQL5 As String
                            'strSQL5 = "select * from por1 where linestatus ='O' and docentry = " & podocentry & ""
                            strSQL5 = "select sum(OpenCreQty)as Openqty from por1  where  DocEntry = " & podocentry & ""
                            RS.DoQuery(strSQL5)
                            openqty = RS.Fields.Item("Openqty").Value
                            'If RS.RecordCount = 0 And a = 0 Then
                            If openqty = 0 Then
                                RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim strSQL6 As String
                                strSQL6 = "Update OPOR set DocStatus = 'C',InvntSttus = 'C' where docentry = " & podocentry & "  "
                                RS.DoQuery(strSQL6)

                            End If

                        End If
                    Catch ex As Exception

                    End Try

                    Try
                        If grpoquantity < POQuantity And docentry <> "0" Then

                            Dim a As Integer
                            a = OpenCreQty - grpoquantity

                            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim strSQL3 As String
                            'strSQL3 = "update por1 set OpenCreQty = " & a & " ,linestatus = 'C' where docentry = " & podocentry & " and linenum = '" & Linenum & "'"
                            strSQL3 = "update por1 set OpenCreQty = " & a & " where docentry = " & podocentry & " and linenum = '" & Linenum & "'"
                            RS.DoQuery(strSQL3)

                            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim strSQL5 As String
                            'strSQL5 = "select * from por1 where linestatus ='O' and docentry = " & podocentry & ""
                            strSQL5 = "select sum(OpenCreQty)as Openqty from por1  where  DocEntry = " & podocentry & ""
                            RS.DoQuery(strSQL5)
                            openqty = RS.Fields.Item("Openqty").Value
                            'If RS.RecordCount = 0 And a = 0 Then
                            If openqty = 0 Then
                                RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim strSQL6 As String
                                strSQL6 = "Update OPOR set DocStatus = 'C',InvntSttus = 'C' where docentry = " & podocentry & "  "
                                RS.DoQuery(strSQL6)

                                Dim strSQL7 As String
                                strSQL7 = "update por1 set linestatus = 'C' where docentry = " & podocentry & " "
                                RS.DoQuery(strSQL7)
                            End If

                        End If
                    Catch ex As Exception
                    End Try

                Else

                    ' RS.DoQuery("Update por1 set linestatus = 'C' where docentry = " & docentry & " and linenum = " & Linenum & "")

                    oPO = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                    RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    strSQL = "select * from por1 where docentry = (select docentry from opor where docnum =" & DocNum & ") and LineStatus = 'O'"
                    RS.DoQuery(strSQL)
                    If RS.RecordCount = 0 Then
                        If oPO.GetByKey(docentry) = True Then
                            oPO.Close()
                            oPO = Nothing
                        End If
                    Else
                        oPO = Nothing
                    End If
                End If
            End If

            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            RS.DoQuery("update [@SST_GAT1] set U_Status = 'A',U_GrnNo = '" & GRNNo & "' where U_ItmCode = '" & ItmCode & "' and docentry = (select docentry from [@SST_OGAT] where docnum =" & GENum & ") ")
            RS = Nothing
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strSQL = "Update [@SST_OGAT] set U_PosStat = 'A' where docnum = " & GENum & ""
            RS.DoQuery(strSQL)
            RS = Nothing
            strSQL = ""

            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strSQL = "Select * from [@SST_GAT1] where isnull(u_status,'') = '' and docentry = (select docentry from [@SST_OGAT] where docnum =" & GENum & ")"
            RS.DoQuery(strSQL)
            If RS.RecordCount = 0 Then
                RS = Nothing
                RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strSQL = ""
                strSQL = "Update [@SST_OGAT] set U_Status = 'C' where docnum = " & GENum & ""
                RS.DoQuery(strSQL)
                RS = Nothing
                strSQL = ""
            End If
            'Next
            If DocType = "I" Then
                RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strSQL = "update [@SST_NQCHDR] set u_docstat = 'C' where docnum = '" & InsNo & "'"
                RS.DoQuery(strSQL)
            ElseIf DocType = "C" Then
                RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strSQL = "update [@SST_CONSHDR] set U_DocStat = 'C' where DocNum =  '" & InsNo & "'"
                RS.DoQuery(strSQL)
            End If
        Catch ex As Exception

        End Try

    End Sub
    Public Sub GetNos(ByVal GEno, ByVal ConsEntry, ByVal PONo)
        GENum = GEno
        CnsEnt = ConsEntry
        PONum = PONo
    End Sub
End Class
