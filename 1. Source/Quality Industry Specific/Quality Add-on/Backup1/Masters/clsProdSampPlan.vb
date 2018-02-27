Public Class clsProdSampPlan
    Private oForm As SAPbouiCOM.Form
    Private oColumns As SAPbouiCOM.Columns
    Private oColumn As SAPbouiCOM.Column
    Private oMatrix As SAPbouiCOM.Matrix
    Private matcol1 As SAPbouiCOM.Column
    Private matcol2 As SAPbouiCOM.Column
    Private matcol3 As SAPbouiCOM.Column
    Private matcol4 As SAPbouiCOM.Column
    Private matcol5, matcol6, matcol7, matcol8 As SAPbouiCOM.Column
    Private txtcode As SAPbouiCOM.EditText
    Private txtdesc As SAPbouiCOM.EditText
    Private txtslno As SAPbouiCOM.EditText
    Private txtstgdesc As SAPbouiCOM.EditText
    Private cbostage As SAPbouiCOM.ComboBox
    Dim oEdit As SAPbouiCOM.EditText
    Dim oEdit1 As SAPbouiCOM.EditText
    Dim oEdit2 As SAPbouiCOM.EditText
    Dim oEdit3 As SAPbouiCOM.EditText
    Dim RS As SAPbobsCOM.Recordset
    Dim oCombo, oCombo2, oCombo4, oCombo5, oCombo7, oCombo10 As SAPbouiCOM.ComboBox
    Dim oDT As SAPbouiCOM.DataTable
    Dim i As Integer
    Dim RowNo As Integer = 0
    Public Const formtype As String = "Frm_PrdPl"
    Private size, sno, per As String
    Dim strSQL As String
    Public Sub LoadScreen()
        Try

            oForm = objAddOn.objUIXml.LoadScreenXML("Frm_ProdSampPlan.xml", SST.enuResourceType.Embeded, formtype)

            txtcode = oForm.Items.Item("txtcode").Specific
            txtdesc = oForm.Items.Item("txtdesc").Specific
            txtslno = oForm.Items.Item("txtslno").Specific
            cbostage = oForm.Items.Item("cbostage").Specific
            txtstgdesc = oForm.Items.Item("txtstgdesc").Specific
            oMatrix = oForm.Items.Item("mat1").Specific
            oColumns = oMatrix.Columns

            matcol1 = oColumns.Item("txtcatcode")
            matcol2 = oColumns.Item("txtparcode")
            matcol3 = oColumns.Item("txtparades")
            matcol4 = oColumns.Item("txtsmplvl")
            matcol5 = oColumns.Item("txtsize")
            matcol6 = oColumns.Item("txtpercent")
            matcol7 = oColumns.Item("txtaccpqty")
            matcol8 = oColumns.Item("txtrejqty")

            LoadCombo1(matcol2)
            LoadCombo2(matcol4)
            LoadCombo()
            LoadRefer()
            Loadcatcode(matcol1)
            oForm.DataBrowser.BrowseBy = "txtslno"

        Catch ex As Exception
        End Try
    End Sub

    Private Sub Loadcatcode(ByVal oCombo20 As SAPbouiCOM.Column)
        Try
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            RS.DoQuery("select Code,name from [@ss_cat]")
            RS.MoveFirst()
            While RS.EoF = False
                oCombo20.ValidValues.Add(RS.Fields.Item("Code").Value, RS.Fields.Item("name").Value)
                RS.MoveNext()
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub LoadCombo1(ByVal oCombo As SAPbouiCOM.Column)
        Try
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            RS.DoQuery("SELECT Code,Name FROM [@SST_PARACAT]")
            RS.MoveFirst()
            While RS.EoF = False
                oCombo.ValidValues.Add(RS.Fields.Item("Code").Value, RS.Fields.Item("Name").Value)
                RS.MoveNext()
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub LoadCombo2(ByVal oCombo4 As SAPbouiCOM.Column)
        Try
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            RS.DoQuery("SELECT Code,Name FROM [@SST_SAMPLINGLEVEL] ")
            RS.MoveFirst()
            While RS.EoF = False
                oCombo4.ValidValues.Add(RS.Fields.Item("Code").Value, RS.Fields.Item("Name").Value)
                RS.MoveNext()
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub LoadRefer()
        Dim RS As SAPbobsCOM.Recordset
        oCombo = oForm.Items.Item("16").Specific
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

    Private Sub LoadCombo()
        oCombo = oForm.Items.Item("cbostage").Specific

        For i = oCombo.ValidValues.Count - 1 To 0 Step -1
            oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
        Next
        RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        RS.DoQuery("select code,name from [@SST_STAGE]")
        While Not RS.EoF
            oCombo.ValidValues.Add(RS.Fields.Item("code").Value, RS.Fields.Item("name").Value)
            RS.MoveNext()
        End While
        oCombo.ValidValues.Add("Define", "Define New")
    End Sub

    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.Before_Action = True Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        oForm = objAddOn.SBO_Application.Forms.Item(FormUID)
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If pVal.ItemUID = "cbostage" Then
                            LoadCombo()
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "1" Then
                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                If Validate(pVal) = False Then
                                    BubbleEvent = False
                                Else
                                    oForm.Items.Item("txtslno").Specific.value = objAddOn.objGenFunc.GetDocNum("SST_PLAN")
                                    oForm.Items.Item("14").Specific.value = objAddOn.objGenFunc.GetDocNum("SST_PLAN")
                                End If
                            End If
                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                If Validate(pVal) = False Then
                                    BubbleEvent = False
                                End If
                            End If
                        End If
                        If pVal.ItemUID = "AddRow" Then
                            oMatrix = oForm.Items.Item("mat1").Specific
                            If oMatrix.RowCount = 0 Then
                                oMatrix.AddRow()
                                oMatrix.GetLineData(oMatrix.RowCount)
                                oForm.DataSources.DBDataSources.Item("@SST_PLANDTL").SetValue("LineId", 0, oMatrix.RowCount)
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
                                    Next
                                Catch ex As Exception
                                    objAddOn.SBO_Application.SetStatusBarMessage("Enter All Mandatory Fields.....")
                                    Exit Sub
                                End Try

                                oForm.DataSources.DBDataSources.Item("@SST_PLANDTL").Clear()
                                oMatrix.AddRow()
                                oMatrix.GetLineData(oMatrix.RowCount)
                                oForm.DataSources.DBDataSources.Item("@SST_PLANDTL").SetValue("LineId", 0, oMatrix.RowCount)
                                oMatrix.SetLineData(oMatrix.RowCount)
                            End If
                        End If
                End Select
            Else
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        If (pVal.ItemUID = "cbostage") Then
                            txtstgdesc.Value = cbostage.Selected.Description
                        End If
                        If (pVal.ItemUID = "mat1") And (pVal.ColUID = "txtuom") Then
                            Dim paracd As String
                            oCombo = matcol3.Cells.Item(pVal.Row).Specific
                            oEdit = matcol4.Cells.Item(pVal.Row).Specific
                            paracd = oCombo.Selected.Value
                            oEdit.Value = oCombo.Selected.Description
                        End If
                        If pVal.ItemUID = "cbostage" Then
                            oCombo = oForm.Items.Item("cbostage").Specific
                            If oCombo.Selected.Value = "Define" Then
                                Dim j As Integer
                                Dim omenus As SAPbouiCOM.MenuItem
                                Dim strSQL As String
                                omenus = objAddOn.SBO_Application.Menus.Item("47616")
                                For j = 0 To omenus.SubMenus.Count - 1
                                    strSQL = omenus.SubMenus.Item(j).String
                                    If strSQL.StartsWith("SST_STG") = True Then
                                        objAddOn.SBO_Application.ActivateMenuItem(omenus.SubMenus.Item(j).UID.ToString)
                                        Exit For
                                    End If
                                Next
                            End If
                        End If

                        Try
                            '***********Parameter code and name**************
                            If (pVal.ItemUID = "mat1") And (pVal.ColUID = "txtparcode") Then
                                matcol3 = oColumns.Item("txtparades")
                                matcol2 = oColumns.Item("txtparcode")
                                oCombo = matcol2.Cells.Item(pVal.Row).Specific
                                oEdit = matcol3.Cells.Item(pVal.Row).Specific
                                oEdit.Value = oCombo.Selected.Description
                            End If
                            '***********Sampling Level**************
                            'If (pVal.ItemUID = "mat1") And (pVal.ColUID = "txtsmplvl") Then
                            '    matcol4 = oColumns.Item("txtsmplvl")
                            '    oCombo4 = matcol4.Cells.Item(pVal.Row).Specific
                            '    oEdit = matcol4.Cells.Item(pVal.Row).Specific
                            '    oEdit.Value = oCombo4.Selected.Description
                            'End If
                            If (pVal.ItemUID = "mat1") And (pVal.ColUID = "txtsmplvl") Then
                                oCombo4 = matcol4.Cells.Item(pVal.Row).Specific
                                'size = oCombo4.Selected.Description
                                'MsgBox(size)
                                Try
                                    Dim oCombo3 As SAPbouiCOM.Column = matcol6
                                    RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    strSQL = "select distinct b.u_percent from [@SST_AQLHDR] a inner join [@SST_AQLDTL] b on a.code = b.code  inner join [@sst_slmdtl] c on b.Code = c.code "
                                    RS.DoQuery(strSQL)
                                    RS.MoveFirst()
                                    If oCombo3.ValidValues.Count > 0 Then
                                        For i As Int16 = oCombo3.ValidValues.Count - 1 To 0 Step -1
                                            oCombo3.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                                        Next
                                    End If
                                    For i As Int16 = 0 To RS.RecordCount - 1
                                        oCombo3.ValidValues.Add(RS.Fields.Item(0).Value, RS.Fields.Item(0).Value)
                                        RS.MoveNext()
                                    Next
                                Catch ex As Exception
                                    MessageBox.Show(ex.Message)
                                End Try
                            End If
                            '***********Percentage**************
                            If (pVal.ItemUID = "mat1") And (pVal.ColUID = "txtpercent") Then
                                oCombo7 = matcol6.Cells.Item(pVal.Row).Specific
                                per = oCombo7.Selected.Value
                                Try
                                    Dim oCombo8 As SAPbouiCOM.Column = matcol5
                                    RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    RS.DoQuery("select distinct  a.u_smpsize from [@SST_AQLHDR] a inner join [@SST_AQLDTL] b on a.code = b.code where b.u_percent = '" & per & "'")

                                    RS.MoveFirst()
                                    If oCombo8.ValidValues.Count > 0 Then
                                        For i As Int16 = oCombo8.ValidValues.Count - 1 To 0 Step -1
                                            oCombo8.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                                        Next
                                    End If
                                    For i As Int16 = 0 To RS.RecordCount - 1
                                        oCombo8.ValidValues.Add(RS.Fields.Item(0).Value, RS.Fields.Item(0).Value)
                                        RS.MoveNext()
                                    Next
                                Catch ex As Exception
                                    MessageBox.Show(ex.Message)
                                End Try
                            End If
                            '***********Size**************
                            If (pVal.ItemUID = "mat1") And (pVal.ColUID = "txtsize") Then
                                oCombo10 = matcol5.Cells.Item(pVal.Row).Specific
                                sno = oCombo10.Selected.Value
                                Try
                                    Dim oCombo9 As SAPbouiCOM.Column = matcol5
                                    RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    RS.DoQuery("select  b.U_accepted,b.U_rejected from [@SST_AQLHDR] a inner join [@SST_AQLDTL] b on a.code = b.code where b.u_percent = '" & per & "' and a.U_smpsize ='" & sno & "'")

                                    RS.MoveFirst()
                                    oMatrix.Columns.Item("txtaccpqty").Cells.Item(pVal.Row).Specific.value = RS.Fields.Item(0).Value
                                    oMatrix.Columns.Item("txtrejqty").Cells.Item(pVal.Row).Specific.value = RS.Fields.Item(1).Value
                                Catch ex As Exception
                                    MessageBox.Show(ex.Message)
                                End Try
                            End If

                          

                        Catch ex As Exception
                        End Try

                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        If pVal.ItemUID = "txtcode" Or pVal.ItemUID = "txtccode" Then
                            Choose(FormUID, pVal)
                        End If
                End Select
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Choose(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent)
        Dim objCFLEvent As SAPbouiCOM.ChooseFromListEvent
        Dim strCFL, strsql As String
        RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objCFLEvent = pval
        strCFL = objCFLEvent.ChooseFromListUID
        oForm = objAddOn.SBO_Application.Forms.Item(FormUID)
        oDT = objCFLEvent.SelectedObjects
        If objCFLEvent.BeforeAction = False Then
            Try
                If strCFL = "CFL_2" Then
                    oForm.DataSources.DBDataSources.Item("@SST_PLANHDR").SetValue("U_itemcode", 0, oDT.GetValue("ItemCode", 0))
                    oForm.DataSources.DBDataSources.Item("@SST_PLANHDR").SetValue("U_itemname", 0, oDT.GetValue("ItemName", 0))
                End If

                If strCFL = "CFL_3" Then
                    oForm.DataSources.DBDataSources.Item("@SST_PLANHDR").SetValue("U_ccode", 0, oDT.GetValue("CardCode", 0))
                    oForm.DataSources.DBDataSources.Item("@SST_PLANHDR").SetValue("U_cname", 0, oDT.GetValue("CardName", 0))
                End If
            Catch ex As Exception
            End Try
        End If
        RS = Nothing
    End Sub

    Private Function Validate(ByVal pVal) As Boolean
        oForm = objAddOn.SBO_Application.Forms.Item(pVal.FormUID)
        oMatrix = oForm.Items.Item("mat1").Specific
        Try

            '**************Mandatory For Code*****************
            oEdit = oForm.Items.Item("txtcode").Specific
            If oEdit.Value = "" Or oEdit.Value = Nothing Then

                oEdit.Active = True
                objAddOn.SBO_Application.SetStatusBarMessage("Code should not be left blank")
                Return False
            End If

            '**************Mandatory For Description*****************
            oEdit1 = oForm.Items.Item("txtdesc").Specific
            If oEdit1.Value = "" Or oEdit1.Value = Nothing Then

                oEdit1.Active = True
                objAddOn.SBO_Application.SetStatusBarMessage("Description should not be left blank")
                Return False
            End If


            '****************** Mandatory Stage Description ***************


            oEdit3 = oForm.Items.Item("txtstgdesc").Specific
            If oEdit3.Value = "" Or oEdit3.Value = Nothing Then

                oEdit3.Active = True
                objAddOn.SBO_Application.SetStatusBarMessage("Stage code should not be left blank")
                Return False
            End If
            If oEdit3.Value = "Define New" Then
                objAddOn.SBO_Application.SetStatusBarMessage("Select Stage code")
                Return False
            End If

            'If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            '    RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '    RS.DoQuery("select U_itemcode, U_stage from [@SST_PLANHDR]  where U_itemcode='" & txtcode.Value & "' and U_stage='" & cbostage.Selected.Value.ToString() & "' ")
            '    If RS.RecordCount > 0 Then
            '        objAddOn.SBO_Application.SetStatusBarMessage("Item Code Already Exists")
            '        Return False
            '    End If
            'End If

            ''****************** Matrix  ***************
            ''****************** Mandatory Quantity From  ***************
            Try

                If oMatrix.RowCount = 0 Then
                    objAddOn.SBO_Application.SetStatusBarMessage("Matrix  should not be left blank")
                    Return False
                Else
                    For i = 1 To oMatrix.RowCount
                        If CDbl(oMatrix.Columns.Item("txtfromqty").Cells.Item(i).Specific.string) = 0 Or CDbl(oMatrix.Columns.Item("txttoqty").Cells.Item(i).Specific.string) = 0 Then
                            objAddOn.SBO_Application.SetStatusBarMessage("Enter From Qty/To Qty.....")
                            Return False
                        End If
                        If CDbl(oMatrix.Columns.Item("txtfromqty").Cells.Item(i).Specific.string) > CDbl(oMatrix.Columns.Item("txttoqty").Cells.Item(i).Specific.string) Then
                            objAddOn.SBO_Application.SetStatusBarMessage("From Qty should be lesser than the To Qty...")
                            Return False
                        End If
                        If oMatrix.Columns.Item("txtsize").Cells.Item(i).Specific.string = "" Then
                            objAddOn.SBO_Application.SetStatusBarMessage("Sample No. should not be empty")
                            Return False
                        End If
                        'If CDbl(oMatrix.Columns.Item("txtsize").Cells.Item(i).Specific.string) < CDbl(oMatrix.Columns.Item("txtfromqty").Cells.Item(i).Specific.string) Or CDbl(oMatrix.Columns.Item("txtsize").Cells.Item(i).Specific.string) > CDbl(oMatrix.Columns.Item("txttoqty").Cells.Item(i).Specific.string) Then
                        '    objAddOn.SBO_Application.SetStatusBarMessage("Sample No. should be between from Qty and to qty")
                        '    Return False
                        'End If
                        If CDbl(oMatrix.Columns.Item("txtsize").Cells.Item(i).Specific.string) > CDbl(oMatrix.Columns.Item("txttoqty").Cells.Item(i).Specific.string) Then
                            objAddOn.SBO_Application.SetStatusBarMessage("Sample No. should be less than or equal to the ToQty")
                            Return False
                        End If
                        If CDbl(oMatrix.Columns.Item("txtsize").Cells.Item(i).Specific.string) <= 0 Then
                            objAddOn.SBO_Application.SetStatusBarMessage("Sample No. should not be zero or -ve")
                            Return False
                        End If
                    Next

                End If
            Catch ex As Exception

            End Try
        Catch ex As Exception

        End Try
        Return True
    End Function

    Private Sub Addcombouom(ByVal oCombo As SAPbouiCOM.Column)
        Try
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            RS.DoQuery("select Code,Name from [@SST_QCUOM]")
            RS.MoveFirst()
            While RS.EoF = False
                oCombo.ValidValues.Add(RS.Fields.Item("Code").Value, RS.Fields.Item("Name").Value)
                RS.MoveNext()
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
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

    Public Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        If pVal.MenuUID = "1290" Or pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1291" Then
            Disable()
        End If

        If pVal.MenuUID = "DelRow" Then
            DeleteRow(pVal, objAddOn.SBO_Application.Forms.ActiveForm.UniqueID)
        End If
    End Sub

    Private Sub Disable()

        oForm.Items.Item("txtcode").Enabled = False

    End Sub

End Class
