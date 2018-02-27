
Public Class clsSamPl
    Private oForm As SAPbouiCOM.Form
    Private oMatrix As SAPbouiCOM.Matrix
    Dim Optn1, Optn2 As SAPbouiCOM.OptionBtn
    Dim ChkBx As SAPbouiCOM.CheckBox
    Dim objRS As SAPbobsCOM.Recordset
    Dim strSQL As String
    Dim oDT As SAPbouiCOM.DataTable
    Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
    Dim oCFL As SAPbouiCOM.ChooseFromList
    Dim oCons As SAPbouiCOM.Conditions
    Dim oCon As SAPbouiCOM.Condition
    Private matcol1 As SAPbouiCOM.Column
    Private matcol2 As SAPbouiCOM.Column
    Private matcol3 As SAPbouiCOM.Column
    Private matcol4 As SAPbouiCOM.Column
    Private matcol5 As SAPbouiCOM.Column
    Private matcol6 As SAPbouiCOM.Column
    Private matcol7 As SAPbouiCOM.Column
    Private matcol8 As SAPbouiCOM.Column
    Private matcol9 As SAPbouiCOM.Column
    Private matcol10, matcol11, matcol12, matcol13, matcol14 As SAPbouiCOM.Column
    Private oColumns As SAPbouiCOM.Columns
    Dim oEdit As SAPbouiCOM.EditText
    Dim oCombo, oCombo1, oCombo2, oCombo3, oCombo4, oCombo5, oCombo6, oCombo7, oCombo8 As SAPbouiCOM.ComboBox
    Dim i As Integer
    Private paracd, size, sno, per As String
    Dim RowNo As Integer = 0
    Public Const formtype As String = "Frm_SamPl"
    Public Sub LoadScreen()
        Try
            oForm = objAddOn.objUIXml.LoadScreenXML("Frm_SamPlnInward.xml", SST.enuResourceType.Embeded, formtype)
            AddCflCon()
            Loadreference()
            oForm.Items.Item("17").Enabled = True
            Optn1 = oForm.Items.Item("16").Specific
            Optn2 = oForm.Items.Item("17").Specific
            Optn2.GroupWith("16")
            ChkBx = oForm.Items.Item("18").Specific
            ChkBx.Checked = False
            oForm.Items.Item("1000004").Enabled = False
            oMatrix = oForm.Items.Item("mat1").Specific
            oColumns = oMatrix.Columns

            matcol4 = oColumns.Item("txtparcod")
            matcol5 = oColumns.Item("txtcatcode")
            matcol10 = oColumns.Item("txtsmplvl")
            matcol11 = oColumns.Item("txtsize")
            matcol12 = oColumns.Item("V_0")
            matcol13 = oColumns.Item("V_1")
            matcol14 = oColumns.Item("txtper")

            'LoadCombo(matcol4)
            LoadCombo2(matcol10)
            ' Loadcatcode(matcol5)
            LoadComboPC(matcol5)
            oForm.DataBrowser.BrowseBy = "20"
        Catch ex As Exception
        End Try
    End Sub

    'Private Sub Loadcatcode(ByVal oCombo20 As SAPbouiCOM.Column)
    '    Try
    '        objRS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '        objRS.DoQuery("select Code,name from [@ss_cat]")
    '        objRS.MoveFirst()
    '        While objRS.EoF = False
    '            oCombo20.ValidValues.Add(objRS.Fields.Item("Code").Value, objRS.Fields.Item("name").Value)
    '            objRS.MoveNext()
    '        End While
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '    End Try
    'End Sub

    Private Sub Loadreference()
        Dim RS As SAPbobsCOM.Recordset
        oCombo = oForm.Items.Item("23").Specific
        Try
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            RS.DoQuery("select Code,Name from [@SST_REF]")
            RS.MoveFirst()
            While RS.EoF = False
                oCombo.ValidValues.Add(RS.Fields.Item("Name").Value, RS.Fields.Item("Name").Value)
                RS.MoveNext()
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub LoadCombo2(ByVal oCombo2 As SAPbouiCOM.Column)
        Try
            objRS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            objRS.DoQuery("SELECT Code,Name FROM [@SST_SAMPLINGLEVEL]")
            objRS.MoveFirst()
            While objRS.EoF = False
                oCombo2.ValidValues.Add(objRS.Fields.Item("Code").Value, objRS.Fields.Item("Name").Value)
                objRS.MoveNext()
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    'Private Sub LoadCombo(ByVal oCombo As SAPbouiCOM.Column)
    '    Try
    '        objRS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '        objRS.DoQuery("SELECT Code,Name FROM [@SST_PARACAT]")
    '        objRS.MoveFirst()
    '        While objRS.EoF = False
    '            oCombo.ValidValues.Add(objRS.Fields.Item("Code").Value, objRS.Fields.Item("Name").Value)
    '            objRS.MoveNext()
    '        End While
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '    End Try
    'End Sub

    Private Sub LoadComboPC(ByVal oCombo8 As SAPbouiCOM.Column)
        Try
            objRS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strSQL = "select * from [@SST_QCPARAMETER]"
            objRS.DoQuery(strSQL)
            While objRS.EoF = False
                oCombo8.ValidValues.Add(objRS.Fields.Item("Code").Value, objRS.Fields.Item("U_paradesc").Value)
                objRS.MoveNext()
            End While
        Catch ex As Exception

        End Try
    End Sub

    Private Sub LoadComboUOM(ByVal oCombo As SAPbouiCOM.Column)
        Try
            objRS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strSQL = "select * from [@SST_QCUOM]"
            objRS.DoQuery(strSQL)
            While objRS.EoF = False
                oCombo.ValidValues.Add(objRS.Fields.Item("Code").Value, objRS.Fields.Item("Name").Value)
                objRS.MoveNext()
            End While
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Choose(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent)
        Dim objCFLEvent As SAPbouiCOM.ChooseFromListEvent
        Dim strCFL As String
        objRS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objCFLEvent = pval
        strCFL = objCFLEvent.ChooseFromListUID
        oForm = objAddOn.SBO_Application.Forms.Item(FormUID)
        oMatrix = oForm.Items.Item("mat1").Specific
        oDT = objCFLEvent.SelectedObjects
        If objCFLEvent.BeforeAction = False Then
            Try
                If strCFL = "CFL_2" Then
                    oForm.DataSources.DBDataSources.Item("@SST_NPLANHDR").SetValue("U_itemcode", 0, oDT.GetValue("ItemCode", 0))
                    oForm.DataSources.DBDataSources.Item("@SST_NPLANHDR").SetValue("U_itemname", 0, oDT.GetValue("ItemName", 0))

                ElseIf strCFL = "CFL_4" Then
                    oForm.DataSources.DBDataSources.Item("@SST_NPLANHDR").SetValue("U_ItmGrp", 0, oDT.GetValue("ItmsGrpCod", 0))
                    oForm.DataSources.DBDataSources.Item("@SST_NPLANHDR").SetValue("U_ItmGrpNM", 0, oDT.GetValue("ItmsGrpNam", 0))

                ElseIf strCFL = "CFL_3" Then
                    oForm.DataSources.DBDataSources.Item("@SST_NPLANHDR").SetValue("U_supcode", 0, oDT.GetValue("CardCode", 0))
                    oForm.DataSources.DBDataSources.Item("@SST_NPLANHDR").SetValue("U_supname", 0, oDT.GetValue("CardName", 0))
                End If
            Catch ex As Exception
                'objAddOn.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End If
        objRS = Nothing
    End Sub

    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.Before_Action = True Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        oForm = objAddOn.SBO_Application.Forms.Item(FormUID)

                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "AddRow" Then
                            oMatrix = oForm.Items.Item("mat1").Specific
                            If oMatrix.RowCount = 0 Then
                                oMatrix.AddRow()
                                oMatrix.GetLineData(oMatrix.RowCount)
                                oForm.DataSources.DBDataSources.Item("@SST_NPLANDTL").SetValue("LineId", 0, oMatrix.RowCount)
                                oMatrix.SetLineData(oMatrix.RowCount)
                            Else
                                Try
                                    For i = 1 To oMatrix.RowCount
                                        oCombo = oMatrix.Columns.Item("txtcatcode").Cells.Item(i).Specific
                                        If oCombo.Selected Is Nothing Then
                                            objAddOn.SBO_Application.SetStatusBarMessage("Select Category Code")
                                            Exit Sub
                                        Else
                                            If oCombo.Selected.Value = "" Then
                                                objAddOn.SBO_Application.SetStatusBarMessage("Select Category Code")
                                                Exit Sub
                                            End If
                                        End If
                                        matcol1 = oColumns.Item("txtparname")
                                        If matcol1.Cells.Item(i).Specific.string = "" Then
                                            objAddOn.SBO_Application.SetStatusBarMessage("Select Parameter Code")
                                            Exit Sub
                                        End If
                                        'If CDbl(oMatrix.Columns.Item("txtqtyfrom").Cells.Item(i).Specific.string) = 0 Or CDbl(oMatrix.Columns.Item("txtqtyto").Cells.Item(i).Specific.string) = 0 Or oMatrix.Columns.Item("txtsize").Cells.Item(i).Specific.string = "" Then
                                        '    objAddOn.SBO_Application.SetStatusBarMessage("Enter From Qty/To Qty/Sample Size.....")
                                        '    Exit Sub
                                        'End If
                                    Next
                                Catch ex As Exception
                                    objAddOn.SBO_Application.SetStatusBarMessage("Enter All Mandatory Fields.....")
                                    Exit Sub
                                End Try

                                oForm.DataSources.DBDataSources.Item("@SST_NPLANDTL").Clear()
                                oMatrix.AddRow()
                                oMatrix.GetLineData(oMatrix.RowCount)
                                oForm.DataSources.DBDataSources.Item("@SST_NPLANDTL").SetValue("LineId", 0, oMatrix.RowCount)
                                oMatrix.SetLineData(oMatrix.RowCount)
                            End If
                        End If



                        If pVal.ItemUID = "1" Then
                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                If Validate(pVal) = False Then
                                    BubbleEvent = False
                                Else
                                    oForm.Items.Item("20").Specific.value = objAddOn.objGenFunc.GetDocNum("SST_NPLAN")
                                    oForm.Items.Item("21").Specific.value = objAddOn.objGenFunc.GetDocNum("SST_NPLAN")
                                End If
                            End If
                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                If Validate(pVal) = False Then
                                    BubbleEvent = False
                                End If
                            End If
                        End If
                End Select

            Else
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If pVal.ItemUID = "18" Then
                            ChkBx = oForm.Items.Item("18").Specific
                            If ChkBx.Checked = True Then
                                oForm.Items.Item("1000004").Enabled = False
                                oForm.Items.Item("txtsupname").Enabled = False
                            Else
                                oForm.Items.Item("1000004").Enabled = True
                                oForm.Items.Item("txtsupname").Enabled = False
                            End If
                        End If

                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            If pVal.ItemUID = "16" Then
                                Optn1 = oForm.Items.Item("16").Specific
                                oForm.Items.Item("1000001").Enabled = True
                                oForm.Items.Item("15").Enabled = False
                            End If
                            If pVal.ItemUID = "17" Then
                                Optn2 = oForm.Items.Item("17").Specific
                                oForm.Items.Item("1000001").Enabled = False
                                oForm.Items.Item("15").Enabled = True
                            End If
                        End If

                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            If pVal.ItemUID = "16" Then
                                Optn2 = oForm.Items.Item("17").Specific
                                If Optn2.Selected = True Then
                                    BubbleEvent = False
                                End If
                            End If
                            If pVal.ItemUID = "17" Then
                                Optn1 = oForm.Items.Item("16").Specific
                                If Optn1.Selected = True Then
                                    BubbleEvent = False
                                End If
                            End If
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        If pVal.ItemUID = "15" Or pVal.ItemUID = "1000001" Or pVal.ItemUID = "1000004" Then
                            Choose(FormUID, pVal)
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                        If (pVal.ItemUID = "mat1") And (pVal.ColUID = "txtcatcode") Then
                            Dim catcode As String
                            oCombo8 = matcol5.Cells.Item(pVal.Row).Specific
                            catcode = oCombo8.Selected.Value

                            '*********** For the category the corresponding parameters will be loaded in the Parameters combo 
                            Try
                                Dim oCombo1 As SAPbouiCOM.Column = matcol5
                                objRS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                objRS.DoQuery("select U_catcode,U_paramcat from [@SST_QCPARAMETER] where code ='" & catcode & "'")
                                objRS.MoveFirst()
                                oMatrix.Columns.Item("txtparcod").Cells.Item(pVal.Row).Specific.value = objRS.Fields.Item(0).Value
                                oMatrix.Columns.Item("txtparname").Cells.Item(pVal.Row).Specific.value = objRS.Fields.Item(1).Value
                            Catch ex As Exception
                                MessageBox.Show(ex.Message)
                            End Try
                        End If

                        If (pVal.ItemUID = "mat1") And (pVal.ColUID = "txtsmplvl") Then
                            oCombo2 = matcol10.Cells.Item(pVal.Row).Specific
                            'size = oCombo2.Selected.Value
                            'MsgBox(size)

                            Try
                                Dim oCombo3 As SAPbouiCOM.Column = matcol14
                                objRS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                objRS.DoQuery("select distinct b.u_percent from [@SST_AQLHDR] a inner join [@SST_AQLDTL] b on a.code = b.code   ")

                                objRS.MoveFirst()
                                If oCombo3.ValidValues.Count > 0 Then
                                    For i As Int16 = oCombo3.ValidValues.Count - 1 To 0 Step -1
                                        oCombo3.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Next
                                End If
                                For i As Int16 = 0 To objRS.RecordCount - 1
                                    oCombo3.ValidValues.Add(objRS.Fields.Item(0).Value, objRS.Fields.Item(0).Value)
                                    objRS.MoveNext()
                                Next
                            Catch ex As Exception
                                MessageBox.Show(ex.Message)
                            End Try
                        End If

                        If (pVal.ItemUID = "mat1") And (pVal.ColUID = "txtper") Then
                            oCombo4 = matcol14.Cells.Item(pVal.Row).Specific
                            per = oCombo4.Selected.Description
                            'MsgBox(per)
                            Try
                                Dim oCombo5 As SAPbouiCOM.Column = matcol11
                                objRS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                objRS.DoQuery("select distinct  a.u_smpsize from [@SST_AQLHDR] a inner join [@SST_AQLDTL] b on a.code = b.code  where b.U_percent = '" & per & "'")
                                objRS.MoveFirst()
                                If oCombo5.ValidValues.Count > 0 Then
                                    For i As Int16 = oCombo5.ValidValues.Count - 1 To 0 Step -1
                                        oCombo5.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Next
                                End If
                                For i As Int16 = 0 To objRS.RecordCount - 1
                                    oCombo5.ValidValues.Add(objRS.Fields.Item(0).Value, objRS.Fields.Item(0).Value)
                                    objRS.MoveNext()
                                Next
                            Catch ex As Exception
                                MessageBox.Show(ex.Message)
                            End Try
                        End If

                        If (pVal.ItemUID = "mat1") And (pVal.ColUID = "txtsize") Then
                            oCombo6 = matcol11.Cells.Item(pVal.Row).Specific
                            sno = oCombo6.Selected.Value
                            'MsgBox(sno)
                            Try
                                Dim oCombo7 As SAPbouiCOM.Column = matcol11
                                objRS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                objRS.DoQuery("select  b.U_accepted,b.U_rejected from [@SST_AQLHDR] a inner join [@SST_AQLDTL] b on a.code = b.code  where b.u_percent = '" & per & "' and a.U_smpsize ='" & sno & "'")
                                objRS.MoveFirst()
                                oMatrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.value = objRS.Fields.Item(0).Value
                                oMatrix.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.value = objRS.Fields.Item(1).Value
                            Catch ex As Exception
                                MessageBox.Show(ex.Message)
                            End Try
                        End If


                        'If (pVal.ItemUID = "mat1") And (pVal.ColUID = "txtparcod") Then
                        '    matcol1 = oColumns.Item("txtparname")
                        '    matcol4 = oColumns.Item("txtparcod")
                        '    oCombo = matcol4.Cells.Item(pVal.Row).Specific
                        '    oEdit = matcol1.Cells.Item(pVal.Row).Specific
                        '    oEdit.Value = oCombo.Selected.Description

                        'End If

                End Select
            End If

        Catch ex As Exception
        End Try
    End Sub

    Private Function Validate(ByVal pVal) As Boolean
        oForm = objAddOn.SBO_Application.Forms.Item(pVal.FormUID)
        oMatrix = oForm.Items.Item("mat1").Specific
        Optn1 = oForm.Items.Item("16").Specific
        Optn2 = oForm.Items.Item("17").Specific
        ChkBx = oForm.Items.Item("18").Specific



        If Optn1.Selected = False And Optn2.Selected = False Then
            objAddOn.SBO_Application.SetStatusBarMessage("Select either Items or ItemGroup")
            Return False
        End If

        If Optn1.Selected = True Then
            If oForm.Items.Item("1000001").Specific.string = "" Or oForm.Items.Item("1000001").Specific.string = Nothing Then
                objAddOn.SBO_Application.SetStatusBarMessage("Select Item Group")
                Return False
            Else
                oForm.Items.Item("15").Specific.string = ""
                oForm.Items.Item("txtdesc").Specific.string = ""
            End If
        End If

        If Optn2.Selected = True Then
            If oForm.Items.Item("15").Specific.string = "" Or oForm.Items.Item("15").Specific.string = Nothing Then
                objAddOn.SBO_Application.SetStatusBarMessage("Select Item Code")
                Return False
            Else
                oForm.Items.Item("1000001").Specific.string = ""
                oForm.Items.Item("1000002").Specific.string = ""
            End If
        End If

        If ChkBx.Checked = True Then
            If oForm.Items.Item("1000004").Specific.string = "" Or oForm.Items.Item("1000004").Specific.string = Nothing Then
                objAddOn.SBO_Application.SetStatusBarMessage("Select Supplier Code")
                Return False
            End If
        Else
            oForm.Items.Item("1000004").Specific.string = ""
            oForm.Items.Item("txtsupname").Specific.string = ""
        End If

        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            If oForm.Items.Item("1000001").Specific.string <> "" Then
                objRS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                objRS.DoQuery("select U_ItmGrp from [@SST_NPLANHDR]  where U_ItmGrp ='" & oForm.Items.Item("1000001").Specific.string & "'  ")
                If objRS.RecordCount > 0 Then
                    objAddOn.SBO_Application.SetStatusBarMessage("Item Group Already Exists")
                    Return False
                End If

            End If
            If oForm.Items.Item("15").Specific.string <> "" Then
                objRS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                objRS.DoQuery("select U_itemcode from [@SST_NPLANHDR]  where U_itemcode='" & oForm.Items.Item("15").Specific.string & "'  ")
                If objRS.RecordCount > 0 Then
                    objAddOn.SBO_Application.SetStatusBarMessage("Item Code Already Exists")
                    Return False
                End If

            End If
        End If

        Try
            If oMatrix.RowCount = 0 Then
                objAddOn.SBO_Application.SetStatusBarMessage("Matrix  should not be left blank")
                Return False
            Else
                For i = 1 To oMatrix.RowCount
                    oCombo = oMatrix.Columns.Item("txtcatcode").Cells.Item(i).Specific
                    If oCombo.Selected Is Nothing Then
                        objAddOn.SBO_Application.SetStatusBarMessage("Select Category Code")
                        Return False
                    Else
                        If oCombo.Selected.Value = "" Then
                            objAddOn.SBO_Application.SetStatusBarMessage("Select Category Code")
                            Return False
                        End If
                    End If
                    matcol1 = oColumns.Item("txtparname")
                    If matcol1.Cells.Item(i).Specific.string = "" Then
                        objAddOn.SBO_Application.SetStatusBarMessage("Select Parameter Code")
                        Return False
                    End If
                    'If CDbl(oMatrix.Columns.Item("txtqtyfrom").Cells.Item(i).Specific.string) = 0 Or CDbl(oMatrix.Columns.Item("txtqtyto").Cells.Item(i).Specific.string) = 0 Or oMatrix.Columns.Item("txtsize").Cells.Item(i).Specific.string = "" Then
                    '    objAddOn.SBO_Application.SetStatusBarMessage("Enter From Qty/To Qty/Sample Size.....")
                    '    Return False
                    'End If
                    'If CDbl(oMatrix.Columns.Item("txtqtyfrom").Cells.Item(i).Specific.string) > CDbl(oMatrix.Columns.Item("txtqtyto").Cells.Item(i).Specific.string) Then
                    '    objAddOn.SBO_Application.SetStatusBarMessage("From Qty should be lesser than the To Qty.....")
                    '    Return False
                    'End If
                    'If CDbl(oMatrix.Columns.Item("txtsize").Cells.Item(i).Specific.string) < CDbl(oMatrix.Columns.Item("txtqtyfrom").Cells.Item(i).Specific.string) Or CDbl(oMatrix.Columns.Item("txtsize").Cells.Item(i).Specific.string) > CDbl(oMatrix.Columns.Item("txtqtyto").Cells.Item(i).Specific.string) Then
                    '    objAddOn.SBO_Application.SetStatusBarMessage("Sample No. should be between from Qty and to qty")
                    '    Return False
                    'End If
                    'If CDbl(oMatrix.Columns.Item("txtsize").Cells.Item(i).Specific.string) > CDbl(oMatrix.Columns.Item("txtqtyto").Cells.Item(i).Specific.string) Then
                    '    objAddOn.SBO_Application.SetStatusBarMessage("Sample No. should be less than or equal to the ToQty")
                    '    Return False
                    'End If
                    'If CDbl(oMatrix.Columns.Item("txtsize").Cells.Item(i).Specific.string) <= 0 Then
                    '    objAddOn.SBO_Application.SetStatusBarMessage("Sample No. should not be zero or -ve")
                    '    Return False
                    'End If
                Next
            End If
        Catch ex As Exception
            objAddOn.SBO_Application.SetStatusBarMessage("Enter All Mandatory Fields.....")
            Return False
        End Try

        Return True
    End Function

    Private Sub AddCflCon()
        oCFLs = oForm.ChooseFromLists
        oCFL = oCFLs.Item("CFL_3")
        oCons = oCFL.GetConditions()
        oCon = oCons.Add()
        oCon.Alias = "CardType"
        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCon.CondVal = "S"
        oCFL.SetConditions(oCons)
    End Sub

    Public Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)

        If pVal.MenuUID = "1290" Or pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1291" Then
            Disable()
        End If
        If pVal.MenuUID = "DelRow" Then
            DeleteRow(pVal, objAddOn.SBO_Application.Forms.ActiveForm.UniqueID)
        End If
        If pVal.MenuUID = "1282" Then
            oForm.Items.Item("17").Enabled = True
            'oForm.Items.Item("16").Enabled = True
        End If
    End Sub

    Private Sub Disable()
        Optn1 = oForm.Items.Item("16").Specific
        Optn2 = oForm.Items.Item("17").Specific
        'If Optn1.Selected = True Then
        '    oForm.Items.Item("17").Enabled = False
        'End If
        'If Optn2.Selected = True Then
        '    oForm.Items.Item("16").Enabled = False
        'End If
        oForm.Items.Item("17").Enabled = False
        oForm.Items.Item("16").Enabled = False
        oForm.Items.Item("15").Enabled = False
        oForm.Items.Item("1000001").Enabled = False
        oForm.Items.Item("18").Enabled = False
        oForm.Items.Item("1000004").Enabled = False
    End Sub

    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        oMatrix = oForm.Items.Item("mat1").Specific
        If (eventInfo.BeforeAction = True) Then
            If eventInfo.ItemUID = "mat1" Then
                oMenuItem = objAddOn.SBO_Application.Menus.Item("1280") 'Data'
                oMenus = oMenuItem.SubMenus
                If oMenus.Exists("DelRow") Then
                Else
                    Try
                        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                        oCreationPackage = objAddOn.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                        oCreationPackage.UniqueID = "DelRow"
                        oCreationPackage.String = "Delete Row"
                        oCreationPackage.Enabled = True
                        oMenuItem = objAddOn.SBO_Application.Menus.Item("1280") 'Data'
                        oMenus = oMenuItem.SubMenus
                        oMenus.AddEx(oCreationPackage)

                    Catch ex As Exception

                    End Try
                End If
                RowNo = eventInfo.Row
            End If
        End If
    End Sub

    Private Sub DeleteRow(ByVal pVal As SAPbouiCOM.MenuEvent, ByVal FormUID As String)
        oForm = objAddOn.SBO_Application.Forms.Item(FormUID)
        oMatrix = oForm.Items.Item("mat1").Specific
        If oMatrix.RowCount > 0 Then
            oMatrix.DeleteRow(RowNo)
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
        End If
    End Sub

End Class
