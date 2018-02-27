Public Class clsInwInsp
    Private oForm As SAPbouiCOM.Form
    Private oMatrix As SAPbouiCOM.Matrix
    Dim RS As SAPbobsCOM.Recordset
    Dim oDT As SAPbouiCOM.DataTable
    Dim strSQL As String
    Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
    Dim oCFL As SAPbouiCOM.ChooseFromList
    Dim oCons As SAPbouiCOM.Conditions
    Dim oCon As SAPbouiCOM.Condition
    Dim oCmb, oCmb1, oCmb2 As SAPbouiCOM.ComboBox
    Dim oColumns As SAPbouiCOM.Columns
    Dim oCol1, oCol2, oCol3, oCol4, oCol5, oCol6, oCol7, oCol8, oCol9, oCol10 As SAPbouiCOM.Column
    Private matcol1 As SAPbouiCOM.Column
    Dim oEdit1, oEdit2, oEdit3 As SAPbouiCOM.EditText
    Dim i As Integer
    Dim k As Integer = 0
    Private Status As String
    Dim oCombo As SAPbouiCOM.ComboBox
    Private PONum As Integer
    Public Const formtype As String = "Frm_InwInsp"

    Public Sub LoadScreen()
        Try
            oForm = objAddOn.objUIXml.LoadScreenXML("Frm_InwardInsp.xml", SST.enuResourceType.Embeded, formtype)

            oCmb = oForm.Items.Item("cboitem").Specific
            oEdit1 = oForm.Items.Item("txtinsno").Specific
            oEdit2 = oForm.Items.Item("txtinsdt").Specific
            oEdit1.Value = objAddOn.objGenFunc.GetDocNum("SST_NINSP")
            oEdit2.String = objAddOn.objGenFunc.GetDateTimeValue(objAddOn.SBO_Application.Company.ServerDate).ToString("dd/MM/yy")
            oMatrix = oForm.Items.Item("mat2").Specific
            oColumns = oMatrix.Columns
            'oCol1 = oColumns.Item("tollp")
            'oCol2 = oColumns.Item("tollm")
            'oCol3 = oColumns.Item("actual")
            'oCol4 = oColumns.Item("1000002")
            oCol5 = oColumns.Item("txtaccplvl")
            oCol6 = oColumns.Item("txtrejlvl")
            oCol7 = oColumns.Item("txtaccqty")
            oCol8 = oColumns.Item("txtrejqty")
            oCol9 = oColumns.Item("txtobser")
            oCol10 = oColumns.Item("txtrmks")

            oForm.DataBrowser.BrowseBy = "txtinsno"
            oMatrix = oForm.Items.Item("mat2").Specific
            'oMatrix.Columns.Item("actual").Editable = True

            LoadCombo()
            AddCflbp()
        Catch ex As Exception
        End Try

    End Sub

    Private Sub LoadCombo()
        oCombo = oForm.Items.Item("27").Specific
        Try
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            RS.DoQuery("select docnum,docentry  from opdn where DocStatus = 'O'")
            RS.MoveFirst()
            While RS.EoF = False
                oCombo.ValidValues.Add(RS.Fields.Item("docnum").Value, RS.Fields.Item("docentry").Value)
                RS.MoveNext()
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub loadGEno(ByVal scode As String)
        oCombo = oForm.Items.Item("27").Specific
        Try
            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            RS.DoQuery("select docnum from [@SST_OGAT]  where U_VCode = '" + scode + "' and U_status='O' ")
            While RS.EoF = False
                oCombo.ValidValues.Add(RS.Fields.Item("docnum").Value, RS.Fields.Item("docnum").Value)
                RS.MoveNext()
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub AddCflbp()
        oCFLs = oForm.ChooseFromLists
        oCFL = oCFLs.Item("CFL_4")
        oCons = oCFL.GetConditions()
        oCon = oCons.Add()
        oCon.Alias = "CardType"
        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCon.CondVal = "S"
        oCFL.SetConditions(oCons)
    End Sub

    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.Before_Action = True Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        oForm = objAddOn.SBO_Application.Forms.Item(FormUID)
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        Try
                            '*************Addrow************
                            If pVal.ItemUID = "Addrow" Then
                                oMatrix = oForm.Items.Item("mat2").Specific
                                If oMatrix.RowCount = 0 Then
                                    oMatrix.AddRow()
                                    oMatrix.GetLineData(oMatrix.RowCount)
                                    oForm.DataSources.DBDataSources.Item("@SST_NQCDTL").SetValue("LineId", 0, oMatrix.RowCount)
                                    oMatrix.SetLineData(oMatrix.RowCount)
                                Else
                                    oForm.DataSources.DBDataSources.Item("@SST_NQCDTL").Clear()
                                    oMatrix.AddRow()
                                    oMatrix.GetLineData(oMatrix.RowCount)
                                    oForm.DataSources.DBDataSources.Item("@SST_NQCDTL").SetValue("LineId", 0, oMatrix.RowCount)
                                    oMatrix.SetLineData(oMatrix.RowCount)
                                End If
                            End If

                        Catch ex As Exception
                        End Try
                        '*******************************
                        If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            If Validate(pVal) = False Then
                                BubbleEvent = False
                            Else
                                PONum = Trim(oForm.DataSources.DBDataSources.Item("@SST_NQCHDR").GetValue("U_GENo", 0))

                                If Status = "A" Then
                                Else
                                    RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    strSQL = "Update [@SST_OGAT] set U_PosStat = 'N' where docnum = " & Trim(oForm.DataSources.DBDataSources.Item("@SST_NQCHDR").GetValue("U_GENo", 0)) & ""
                                    RS.DoQuery(strSQL)
                                    RS = Nothing
                                    strSQL = ""
                                End If
                            End If
                        End If
                        If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            If oForm.Items.Item("26").Specific.string = "A" Then
                                RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                strSQL = "select * from pdn1 where U_InsNo ='" & oForm.Items.Item("txtinsno").Specific.string & "' and U_EType = 'I'"
                                RS.DoQuery(strSQL)
                                If RS.RecordCount > 0 Then
                                    objAddOn.SBO_Application.SetStatusBarMessage("Cannot be updated")
                                    BubbleEvent = False
                                Else
                                    If Validate(pVal) = False Then
                                        BubbleEvent = False
                                    End If
                                End If
                            ElseIf oForm.Items.Item("26").Specific.string = "N" Then
                                RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                strSQL = "select * from [@SST_CONSHDR] where U_InsNo ='" & oForm.Items.Item("txtinsno").Specific.string & "' "
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
                        End If
                End Select
            Else
                Select Case pVal.EventType

                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        If pVal.ItemUID = "cboitem" Then
                            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            strSQL = "SELECT  T1.[Dscription] as Dscription, sum(T1.[Quantity]) as Quantity"
                            strSQL = strSQL + " FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry "
                            strSQL = strSQL + " WHERE T1.[ItemCode]  = '" & oCmb.Selected.Description & "' and t0.docentry = '" & oForm.Items.Item("27").Specific.Value & "'"
                            strSQL = strSQL + " group by T1.[Dscription]"
                            RS.DoQuery(strSQL)
                            If RS.RecordCount > 0 Then
                                oForm.DataSources.DBDataSources.Item("@SST_NQCHDR").SetValue("U_GEqty", 0, RS.Fields.Item("Quantity").Value)
                                oForm.DataSources.DBDataSources.Item("@SST_NQCHDR").SetValue("U_itemname", 0, RS.Fields.Item("Dscription").Value)
                            End If
                            RS = Nothing

                            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            strSQL = "select b.U_paracode,b.U_paradesc,b.U_smpsize,b.U_AccNo,b.U_RejNo,b.U_percen,d.U_value,d.U_tollplus as USL, d.U_tollmins as LSL"
                            strSQL = strSQL + " from [@SST_Nplanhdr] a inner join [@SST_Nplandtl] b on a.Code = b.Code"
                            strSQL = strSQL + " inner join [@SST_QCSTANDHDR] c on a.U_itemcode = c.U_itemcode"
                            strSQL = strSQL + " inner join [@SST_QCSTANDDTL] d on c.Code = d.code"
                            strSQL = strSQL + " where a.U_itemcode = '" & oCmb.Selected.Description & "'"
                            RS.DoQuery(strSQL)

                            If RS.RecordCount > 0 Then
                                If oMatrix.RowCount > 0 Then
                                    oMatrix.Clear()
                                End If

                                '*******
                                Dim rowno As Integer
                                rowno = 0
                                RS.MoveFirst()

                                '*******

                                While Not RS.EoF
                                    oMatrix.AddRow()

                                    rowno += 1
                                    oMatrix.GetLineData(oMatrix.RowCount)
                                    oForm.DataSources.DBDataSources.Item("@SST_NQCDTL").SetValue("LineId", 0, oMatrix.RowCount)
                                    oForm.DataSources.DBDataSources.Item("@SST_NQCDTL").SetValue("U_paradesc", 0, RS.Fields.Item("U_paradesc").Value)
                                    oForm.DataSources.DBDataSources.Item("@SST_NQCDTL").SetValue("U_smpsize", 0, RS.Fields.Item("U_smpsize").Value)
                                    oForm.DataSources.DBDataSources.Item("@SST_NQCDTL").SetValue("U_percen", 0, RS.Fields.Item("U_percen").Value)
                                    oForm.DataSources.DBDataSources.Item("@SST_NQCDTL").SetValue("U_acclvl", 0, RS.Fields.Item("U_AccNo").Value)
                                    oForm.DataSources.DBDataSources.Item("@SST_NQCDTL").SetValue("U_rejlvl", 0, RS.Fields.Item("U_RejNo").Value)

                                    oForm.DataSources.DBDataSources.Item("@SST_NQCDTL").SetValue("U_value", 0, RS.Fields.Item("U_value").Value)
                                    oForm.DataSources.DBDataSources.Item("@SST_NQCDTL").SetValue("U_usl", 0, RS.Fields.Item("USL").Value)
                                    oForm.DataSources.DBDataSources.Item("@SST_NQCDTL").SetValue("U_lsl", 0, RS.Fields.Item("LSL").Value)

                                    oMatrix.SetLineData(rowno)
                                    RS.MoveNext()
                                End While
                                'For i = 0 To RS.RecordCount
                                '    oMatrix.AddRow()
                                '    oMatrix.GetLineData(oMatrix.RowCount)
                                '    oForm.DataSources.DBDataSources.Item("@SST_NQCDTL").SetValue("LineId", 0, oMatrix.RowCount)
                                '    oForm.DataSources.DBDataSources.Item("@SST_NQCDTL").SetValue("U_paradesc", 0, RS.Fields.Item("U_paradesc").Value)
                                '    oForm.DataSources.DBDataSources.Item("@SST_NQCDTL").SetValue("U_smpsize", 0, RS.Fields.Item("U_smpsize").Value)
                                '    oForm.DataSources.DBDataSources.Item("@SST_NQCDTL").SetValue("U_percen", 0, RS.Fields.Item("U_percen").Value)
                                '    oForm.DataSources.DBDataSources.Item("@SST_NQCDTL").SetValue("U_acclvl", 0, RS.Fields.Item("U_AccNo").Value)
                                '    oForm.DataSources.DBDataSources.Item("@SST_NQCDTL").SetValue("U_rejlvl", 0, RS.Fields.Item("U_RejNo").Value)
                                '    oMatrix.SetLineData(oMatrix.RowCount)
                                '    i = i + 1
                                'Next

                            Else
                                objAddOn.SBO_Application.SetStatusBarMessage("No matching records....", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                        End If

                        If pVal.ItemUID = "27" Then
                            Try
                                RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                For i = oCmb.ValidValues.Count - 1 To 0 Step -1
                                    oCmb.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                                Next
                                strSQL = "SELECT t0.CardCode,t0.CardName,t1.ItemCode, T1.[Dscription],T1.[Quantity] as Quantity"
                                strSQL = strSQL + " FROM OPDN T0 INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry "
                                strSQL = strSQL + " where  T0.docentry = " & Trim(oForm.DataSources.DBDataSources.Item("@SST_NQCHDR").GetValue("U_GENo", 0)) & ""
                                RS.DoQuery(strSQL)
                                While Not RS.EoF
                                    oForm.Items.Item("txtsup").Specific.String = RS.Fields.Item("CardCode").Value
                                    oForm.Items.Item("txtsupname").Specific.String = RS.Fields.Item("CardName").Value
                                    'oForm.Items.Item("txtqty").Specific.String = RS.Fields.Item("u_qty").Value
                                    oCmb.ValidValues.Add(RS.Fields.Item("ItemCode").Value, RS.Fields.Item("ItemCode").Value)


                                    RS.MoveNext()
                                End While
                            Catch ex As Exception
                            End Try
                        End If



                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        'If pVal.ItemUID = "txtgeno" Then
                        '    Choose(FormUID, pVal)
                        'End If
                        If pVal.ItemUID = "txtsup" Then
                            Choose(FormUID, pVal)
                            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            strSQL = "select sum(t1.U_Qty) Qty,t1.U_itmdesc,t1.u_uom from [@SST_GAT1] t1"
                            strSQL = strSQL + " inner join [@SST_OGAT] t2 on t2.docentry = t1.docentry"
                            strSQL = strSQL + " where t2.docnum = " & Trim(oForm.DataSources.DBDataSources.Item("@SST_NQCHDR").GetValue("U_GENo", 0)) & " and t1.U_ItmCode = '" & oCmb.Selected.Value & "'"
                            strSQL = strSQL + " group by t1.U_itmdesc,t1.u_uom"
                            RS.DoQuery(strSQL)
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        If pVal.ItemUID = "mat2" And pVal.ColUID = "actual" Then
                            If CDbl(oCol1.Cells.Item(pVal.Row).Specific.string) > CDbl(oCol2.Cells.Item(pVal.Row).Specific.string) Then
                                If CDbl(oCol1.Cells.Item(pVal.Row).Specific.string) >= CDbl(oCol3.Cells.Item(pVal.Row).Specific.string) And CDbl(oCol2.Cells.Item(pVal.Row).Specific.string) <= CDbl(oCol3.Cells.Item(pVal.Row).Specific.string) Then
                                    'oMatrix.Columns.Item("status").Cells.Item(pVal.Row).Specific.string = "A"
                                    oMatrix.Columns.Item("status").Cells.Item(pVal.Row).Specific.string = "Accepted"
                                Else
                                    'oMatrix.Columns.Item("status").Cells.Item(pVal.Row).Specific.string = "N"
                                    oMatrix.Columns.Item("status").Cells.Item(pVal.Row).Specific.string = "Rejection"
                                End If
                            End If
                            If CDbl(oCol2.Cells.Item(pVal.Row).Specific.string) > CDbl(oCol1.Cells.Item(pVal.Row).Specific.string) Then
                                If CDbl(oCol1.Cells.Item(pVal.Row).Specific.string) <= CDbl(oCol3.Cells.Item(pVal.Row).Specific.string) And CDbl(oCol2.Cells.Item(pVal.Row).Specific.string) >= CDbl(oCol3.Cells.Item(pVal.Row).Specific.string) Then
                                    'oMatrix.Columns.Item("status").Cells.Item(pVal.Row).Specific.string = "A"
                                    oMatrix.Columns.Item("status").Cells.Item(pVal.Row).Specific.string = "Accepted"
                                Else
                                    'oMatrix.Columns.Item("status").Cells.Item(pVal.Row).Specific.string = "N"
                                    oMatrix.Columns.Item("status").Cells.Item(pVal.Row).Specific.string = "Rejection"
                                End If
                            End If
                        End If
                End Select
            End If

            If pVal.ItemUID = "1" And pVal.Action_Success = True And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oEdit1 = oForm.Items.Item("txtinsno").Specific
                    oEdit2 = oForm.Items.Item("txtinsdt").Specific
                    oEdit1.Value = objAddOn.objGenFunc.GetDocNum("SST_NINSP")
                    oEdit2.String = objAddOn.objGenFunc.GetDateTimeValue(objAddOn.SBO_Application.Company.ServerDate).ToString("dd/MM/yy")
                End If

            End If
        Catch ex As Exception
            oForm.Freeze(False)
        End Try
    End Sub

    Private Function Validate(ByVal pVal) As Boolean
        oForm = objAddOn.SBO_Application.Forms.Item(pVal.FormUID)
        oMatrix = oForm.Items.Item("mat2").Specific


        If oForm.Items.Item("27").Specific.value = "" Then
            objAddOn.SBO_Application.SetStatusBarMessage("Select GRPO ....")
            Return False
        End If
        If oForm.Items.Item("cboitem").Specific.value = "" Then
            objAddOn.SBO_Application.SetStatusBarMessage("Select ItemCode ....")
            Return False
        End If

        'Try
        '    If oForm.Items.Item("txtname").Specific.string = "" Then
        '        objAddOn.SBO_Application.SetStatusBarMessage("Choose ItemCode......")
        '        Return False
        '    End If
        'Catch ex As Exception
        '    objAddOn.SBO_Application.SetStatusBarMessage("Choose ItemCode......")
        '    Return False
        'End Try

        If oMatrix.RowCount = 0 Then
            objAddOn.SBO_Application.SetStatusBarMessage("No line Items....")
            Return False
        End If

        Try
            For i = 1 To oMatrix.RowCount

                If oMatrix.Columns.Item("txtaccqty").Cells.Item(i).Specific.string = "" Then
                    objAddOn.SBO_Application.SetStatusBarMessage("Enter Accepted Qty....")
                    Return False
                End If

                If oMatrix.Columns.Item("txtrejqty").Cells.Item(i).Specific.string = "" Then
                    objAddOn.SBO_Application.SetStatusBarMessage("Enter Rejected Qty....")
                    Return False
                End If

                If CDbl(oForm.Items.Item("txtqty").Specific.string) <> CDbl(oMatrix.Columns.Item("txtaccqty").Cells.Item(i).Specific.string) + CDbl(oMatrix.Columns.Item("txtrejqty").Cells.Item(i).Specific.string) Then
                    objAddOn.SBO_Application.SetStatusBarMessage("Sum of Quantities should be equal to the received qty")
                    Return False
                End If

                If oMatrix.Columns.Item("txtrmrks").Cells.Item(i).Specific.string = "" Then
                    objAddOn.SBO_Application.SetStatusBarMessage("Enter Remarks....")
                    Return False
                End If

            Next

            'For i = 1 To oMatrix.RowCount
            '    'If oMatrix.Columns.Item("status").Cells.Item(i).Specific.string = "N" Then
            '    If oMatrix.Columns.Item("status").Cells.Item(i).Specific.string = "Rejection" Then
            '        oForm.Items.Item("26").Specific.string = "N"
            '        Status = oForm.Items.Item("26").Specific.string
            '        Exit For
            '    Else
            '        oForm.Items.Item("26").Specific.string = "A"
            '        Status = oForm.Items.Item("26").Specific.string
            '    End If
            'Next
        Catch ex As Exception

        End Try

        Return True
    End Function

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
                If strCFL = "CFL_4" Then
                    'oForm.Items.Item("1000001").Specific.string = objAddOn.objGenFunc.GetDateTimeValue(oDT.GetValue("CreateDate", 0))
                    oForm.DataSources.DBDataSources.Item("@SST_NQCHDR").SetValue("U_supcode", 0, oDT.GetValue("CardCode", 0))
                    oForm.DataSources.DBDataSources.Item("@SST_NQCHDR").SetValue("U_supname", 0, oDT.GetValue("CardName", 0))
                    loadGEno(oDT.GetValue("CardCode", 0))

                    oCombo = oForm.Items.Item("cboitem").Specific
                    For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                        oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                    Next
                    RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    strSQL = "SELECT T1.[ItemCode] "
                    strSQL = strSQL + " FROM OPDN T0 INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry "
                    strSQL = strSQL + " WHERE T0.[CardCode]  = '" & oForm.Items.Item("txtsup").Specific.String & "' and T0.[DocStatus] = 'O'"
                    strSQL = strSQL + " group by T1.[ItemCode]"
                    RS.DoQuery(strSQL)
                    While Not RS.EoF
                        k = k + 1
                        oCombo.ValidValues.Add(k, RS.Fields.Item("ItemCode").Value)
                        RS.MoveNext()
                    End While
                End If

            Catch ex As Exception
            End Try
        End If
        RS = Nothing
    End Sub

    Public Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)

        If pVal.MenuUID = "1282" Then
            oEdit1 = oForm.Items.Item("txtinsno").Specific
            oEdit2 = oForm.Items.Item("txtinsdt").Specific
            oEdit1.Value = objAddOn.objGenFunc.GetDocNum("SST_NINSP")
            oEdit2.String = objAddOn.objGenFunc.GetDateTimeValue(objAddOn.SBO_Application.Company.ServerDate).ToString("dd/MM/yy")
            oForm.Items.Item("txtSupCd").Enabled = True
            oForm.Items.Item("27").Enabled = True
            oForm.Items.Item("cboitem").Enabled = True
            oForm.Items.Item("1000002").Enabled = True
            oMatrix = oForm.Items.Item("mat2").Specific
            oMatrix.Columns.Item("actual").Editable = True
            oForm.Items.Item("31").Enabled = True
        End If

        If pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291" Then
            If oForm.Items.Item("31").Specific.selected.value = "C" Then
                oForm.Items.Item("31").Enabled = False
                oForm.Items.Item("txtSupCd").Enabled = False
                oForm.Items.Item("27").Enabled = False
                oForm.Items.Item("cboitem").Enabled = False
                oForm.Items.Item("1000002").Enabled = False
                oMatrix = oForm.Items.Item("mat2").Specific
                oMatrix.Columns.Item("actual").Editable = False
            Else
                oForm.Items.Item("31").Enabled = True
                oForm.Items.Item("txtSupCd").Enabled = True
                oForm.Items.Item("27").Enabled = True
                oForm.Items.Item("cboitem").Enabled = True
                oForm.Items.Item("1000002").Enabled = True
                oMatrix = oForm.Items.Item("mat2").Specific
                oMatrix.Columns.Item("actual").Editable = True
            End If
        End If
    End Sub

End Class
