' Author                     Created Date
' Manimaran                   20/11/2010
Public Class clsItemParaInward
    Private oForm As SAPbouiCOM.Form
    Private oMatrix As SAPbouiCOM.Matrix
    Private matcol1 As SAPbouiCOM.Column
    Private matcol2 As SAPbouiCOM.Column
    Private matcol3 As SAPbouiCOM.Column
    Private matcol4 As SAPbouiCOM.Column
    Private matcol5 As SAPbouiCOM.Column
    Private matcol6 As SAPbouiCOM.Column
    Private matcol7 As SAPbouiCOM.Column
    Private matcol8 As SAPbouiCOM.Column
    Private matcol9 As SAPbouiCOM.Column
    Private txtcode As SAPbouiCOM.EditText
    Private txtdesc As SAPbouiCOM.EditText
    Private txtslno As SAPbouiCOM.EditText
    Private txtcatcode As SAPbouiCOM.ComboBox
    Private txtparcode As SAPbouiCOM.ComboBox
    Private paracd As String
    Private paracd1 As String
    Private oColumns As SAPbouiCOM.Columns
    Private oColumn As SAPbouiCOM.Column
    Dim oItem, oItem1 As SAPbouiCOM.EditText
    Dim RS As SAPbobsCOM.Recordset
    Dim oEdit, oEdit1 As SAPbouiCOM.EditText
    Dim oCombo, oCombo1, oCombo2, oCombo5, oCombo6, oCombo7 As SAPbouiCOM.ComboBox
    Private oDT As SAPbouiCOM.DataTable
    Dim i As Integer
    Dim RowNo As Integer = 0
    Public Const formtype As String = "Frm_ItmPm"
    Dim strsql As String
    Public Sub LoadScreen()
        oForm = objAddOn.objUIXml.LoadScreenXML("Frm_ItmPRIwd.xml", SST.enuResourceType.Embeded, formtype)

        txtcode = oForm.Items.Item("txtcode").Specific
        txtdesc = oForm.Items.Item("txtdesc").Specific
        txtslno = oForm.Items.Item("txtslno").Specific
        oMatrix = oForm.Items.Item("mat1").Specific
        oColumns = oMatrix.Columns
        matcol1 = oColumns.Item("txtcatcode")
        matcol2 = oColumns.Item("txtcatname")
        matcol3 = oColumns.Item("txtparcode")
        matcol4 = oColumns.Item("txtparname")
        matcol5 = oColumns.Item("txtuomcode")
        matcol6 = oColumns.Item("txtuomname")
        matcol7 = oColumns.Item("txtvalue")
        matcol8 = oColumns.Item("txttollp")
        matcol9 = oColumns.Item("txttollm")
        'Addcombocategory(matcol3)
        Loadcatcode(matcol1)
        'parameter(matcol3)
        'UOM(matcol5)
       
        oForm.DataBrowser.BrowseBy = "txtslno"
    End Sub

    Private Sub Loadcatcode(ByVal oCombo5 As SAPbouiCOM.Column)
        Try
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            RS.DoQuery("select Code,name from [@ss_cat]")
            RS.MoveFirst()
            While RS.EoF = False
                oCombo5.ValidValues.Add(RS.Fields.Item("Code").Value, RS.Fields.Item("name").Value)
                RS.MoveNext()
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    'Private Sub Addcombocategory(ByVal oCombo As SAPbouiCOM.Column)
    '    Try
    '        RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '        RS.DoQuery("SELECT Code,Name FROM [@SST_PARACAT]")
    '        RS.MoveFirst()
    '        While RS.EoF = False
    '            oCombo.ValidValues.Add(RS.Fields.Item("Code").Value, RS.Fields.Item("Name").Value)
    '            RS.MoveNext()
    '        End While
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '    End Try
    'End Sub

    'Private Sub parameter(ByVal oCombo6 As SAPbouiCOM.Column)
    '    Try
    '        RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '        RS.DoQuery("SELECT Code,u_paradesc FROM [@SST_QCPARAMETER]")
    '        RS.MoveFirst()
    '        While RS.EoF = False
    '            oCombo6.ValidValues.Add(RS.Fields.Item("Code").Value, RS.Fields.Item("u_paradesc").Value)
    '            RS.MoveNext()
    '        End While
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '    End Try
    'End Sub

    'Private Sub UOM(ByVal oCombo7 As SAPbouiCOM.Column)
    '    Try
    '        RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '        RS.DoQuery("SELECT Code,Name FROM [@SST_QCUOM]")
    '        RS.MoveFirst()
    '        While RS.EoF = False
    '            oCombo7.ValidValues.Add(RS.Fields.Item("Code").Value, RS.Fields.Item("Name").Value)
    '            RS.MoveNext()
    '        End While
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '    End Try
    'End Sub

    Private Function Validate(ByVal pVal) As Boolean
        oForm = objAddOn.SBO_Application.Forms.Item(pVal.FormUID)
        oMatrix = oForm.Items.Item("mat1").Specific


        ''**************Mandatory For Code*****************
        'oEdit = oForm.Items.Item("txtcode").Specific
        'If oEdit.Value = "" Or oEdit.Value = Nothing Then
        '    oEdit.Active = True
        '    objAddOn.SBO_Application.SetStatusBarMessage("Code should not be left blank")
        '    Return False
        'End If

        ''**************Mandatory For Description*****************
        'oEdit1 = oForm.Items.Item("txtdesc").Specific
        'If oEdit1.Value = "" Or oEdit1.Value = Nothing Then
        '    oEdit1.Active = True
        '    objAddOn.SBO_Application.SetStatusBarMessage("Description should not be left blank")
        '    Return False
        'End If

        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            RS.DoQuery("select U_itemcode from [@SST_QCSTANDHDR]  where U_itemcode='" & txtcode.Value & "'  ")
            If RS.RecordCount > 0 Then
                objAddOn.SBO_Application.SetStatusBarMessage("Item Code Already Exists")
                Return False
            End If
        End If

        Try
            If oMatrix.RowCount = 0 Then
                objAddOn.SBO_Application.SetStatusBarMessage("Matrix should not be left blank")
                Return False
            Else
                For i = 1 To oMatrix.RowCount
                    If CDbl(oMatrix.Columns.Item("txttollp").Cells.Item(i).Specific.value) = 0 Or CDbl(oMatrix.Columns.Item("txttollm").Cells.Item(i).Specific.value) = 0 Or CDbl(oMatrix.Columns.Item("txtvalue").Cells.Item(i).Specific.value) = 0 Then
                        objAddOn.SBO_Application.SetStatusBarMessage("Enter the ULS/LLS/Values.....")
                        Return False
                    End If

                    If CDbl(oMatrix.Columns.Item("txttollp").Cells.Item(i).Specific.value) < 0 Or CDbl(oMatrix.Columns.Item("txttollm").Cells.Item(i).Specific.value) < 0 Or CDbl(oMatrix.Columns.Item("txtvalue").Cells.Item(i).Specific.value) < 0 Then
                        objAddOn.SBO_Application.SetStatusBarMessage("ULS/LLS/Values should not be Negative")
                        Return False
                    End If

                    If CDbl(oMatrix.Columns.Item("txttollp").Cells.Item(i).Specific.value) < CDbl(oMatrix.Columns.Item("txtvalue").Cells.Item(i).Specific.value) Then
                        objAddOn.SBO_Application.SetStatusBarMessage("ULS should be equal or greater than the value")
                        Return False
                    End If
                    If CDbl(oMatrix.Columns.Item("txttollm").Cells.Item(i).Specific.value) > CDbl(oMatrix.Columns.Item("txtvalue").Cells.Item(i).Specific.value) Then
                        objAddOn.SBO_Application.SetStatusBarMessage("LLS should be equal or lesser than the value")
                        Return False
                    End If
                    'Try
                    '    oCombo = matcol1.Cells.Item(i).Specific
                    '    oCombo1 = matcol3.Cells.Item(i).Specific
                    '    oCombo2 = matcol5.Cells.Item(i).Specific

                    '    If oCombo.Selected Is Nothing Then
                    '        objAddOn.SBO_Application.SetStatusBarMessage("Select Category Code")
                    '        Return False
                    '    Else
                    '        If oCombo.Selected.Value = "" Then
                    '            objAddOn.SBO_Application.SetStatusBarMessage("Select Category Code")
                    '            Return False
                    '        End If
                    '    End If

                    '    If matcol4.Cells.Item(i).Specific.string = "" Then
                    '        objAddOn.SBO_Application.SetStatusBarMessage("Select Parameter Code")
                    '        Return False

                    '    End If

                    '    If matcol6.Cells.Item(i).Specific.string = "" Then
                    '        objAddOn.SBO_Application.SetStatusBarMessage("Select UOM Code")
                    '        Return False
                    '    End If

                    'Catch ex As Exception
                    '    objAddOn.SBO_Application.SetStatusBarMessage("Enter All Mandatory Fields.....")
                    '    Return False
                    'End Try
                Next
            End If
        Catch ex As Exception

        End Try
        Return True

    End Function

    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.Before_Action = True Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        oForm = objAddOn.SBO_Application.Forms.Item(FormUID)
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "1" Then
                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                If Validate(pVal) = False Then
                                    BubbleEvent = False
                                Else
                                    oForm.Items.Item("txtslno").Specific.value = objAddOn.objGenFunc.GetDocNum("SST_STAND")
                                    oForm.Items.Item("txttemp").Specific.value = objAddOn.objGenFunc.GetDocNum("SST_STAND")
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
                                oForm.DataSources.DBDataSources.Item("@SST_QCSTANDDTL").SetValue("LineId", 0, oMatrix.RowCount)
                                oMatrix.SetLineData(oMatrix.RowCount)
                            Else
                                Try
                                    For i = 1 To oMatrix.RowCount
                                        oCombo = matcol1.Cells.Item(i).Specific
                                        oCombo1 = matcol3.Cells.Item(i).Specific
                                        oCombo2 = matcol5.Cells.Item(i).Specific

                                        If oCombo.Selected Is Nothing Then
                                            objAddOn.SBO_Application.SetStatusBarMessage("Select Category Code")
                                            Exit Sub
                                        Else
                                            If oCombo.Selected.Value = "" Then
                                                objAddOn.SBO_Application.SetStatusBarMessage("Select Category Code")
                                                Exit Sub
                                            End If
                                        End If

                                        If matcol4.Cells.Item(i).Specific.string = "" Then
                                            objAddOn.SBO_Application.SetStatusBarMessage("Select Parameter Code")
                                            Exit Sub
                                        End If

                                        If matcol6.Cells.Item(i).Specific.string = "" Then
                                            objAddOn.SBO_Application.SetStatusBarMessage("Select UOM Code")
                                            Exit Sub
                                        End If

                                        If CDbl(oMatrix.Columns.Item("txttollp").Cells.Item(i).Specific.value) = 0 Or CDbl(oMatrix.Columns.Item("txttollm").Cells.Item(i).Specific.value) = 0 Or CDbl(oMatrix.Columns.Item("txtvalue").Cells.Item(i).Specific.value) = 0 Then
                                            objAddOn.SBO_Application.SetStatusBarMessage("Enter the ULS/LLS/Values.....")
                                            Exit Sub
                                        End If
                                    Next
                                Catch ex As Exception
                                    objAddOn.SBO_Application.SetStatusBarMessage("Enter All Mandatory Fields.....")
                                    Exit Sub
                                End Try

                                oForm.DataSources.DBDataSources.Item("@SST_QCSTANDDTL").Clear()
                                oMatrix.AddRow()
                                oMatrix.GetLineData(oMatrix.RowCount)
                                oForm.DataSources.DBDataSources.Item("@SST_QCSTANDDTL").SetValue("LineId", 0, oMatrix.RowCount)
                                oMatrix.SetLineData(oMatrix.RowCount)
                            End If
                        End If
                End Select
            Else
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                        If (pVal.ItemUID = "mat1") And (pVal.ColUID = "txtparcode") Then
                            Dim paracd3 As String
                            oCombo6 = matcol3.Cells.Item(pVal.Row).Specific
                            oEdit = matcol4.Cells.Item(pVal.Row).Specific
                            paracd3 = oCombo6.Selected.Value
                            oEdit.Value = oCombo6.Selected.Description
                            MsgBox(paracd3)
                            'oCombo = matcol3.Cells.Item(pVal.Row).Specific
                            'oEdit = matcol4.Cells.Item(pVal.Row).Specific
                            'paracd1 = oCombo.Selected.Value
                            'oEdit.Value = oCombo.Selected.Description
                            ''*********** For the Parameter the corresponding parameters Uom will be loaded in the Parameters Uom combo 
                            Try
                                Dim oCombo1 As SAPbouiCOM.Column = matcol5
                                RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                RS.DoQuery("select U_uomcode,U_uomdesc from [@SST_QCPARAMETER] where Code='" & paracd3 & "'")
                                RS.MoveFirst()
                                If oCombo1.ValidValues.Count > 0 Then
                                    For i As Int16 = oCombo1.ValidValues.Count - 1 To 0 Step -1
                                        oCombo1.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Next
                                End If
                                For i As Int16 = 0 To RS.RecordCount - 1
                                    oCombo1.ValidValues.Add(RS.Fields.Item(0).Value, RS.Fields.Item(1).Value)
                                    RS.MoveNext()
                                Next
                            Catch ex As Exception
                                MessageBox.Show(ex.Message)
                            End Try
                        End If
                        If (pVal.ItemUID = "mat1") And (pVal.ColUID = "txtcatcode") Then

                            oCombo = matcol1.Cells.Item(pVal.Row).Specific
                            oEdit = matcol2.Cells.Item(pVal.Row).Specific
                            paracd = oCombo.Selected.Value
                            oEdit.Value = oCombo.Selected.Description

                            '*********** For the category the corresponding parameters will be loaded in the Parameters combo 
                            'Try
                            '    Dim oCombo1 As SAPbouiCOM.Column = matcol3
                            '    RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            '    RS.DoQuery("select Code,U_paradesc,U_uomcode,U_uomdesc from [@SST_QCPARAMETER] where U_catcode='" & paracd & "'")
                            '    RS.MoveFirst()
                            '    If oCombo1.ValidValues.Count > 0 Then
                            '        For i As Int16 = oCombo1.ValidValues.Count - 1 To 0 Step -1
                            '            oCombo1.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                            '        Next
                            '    End If
                            '    For i As Int16 = 0 To RS.RecordCount - 1
                            '        oCombo1.ValidValues.Add(RS.Fields.Item(0).Value, RS.Fields.Item(1).Value)
                            '        RS.MoveNext()
                            '    Next
                            'Catch ex As Exception
                            '    MessageBox.Show(ex.Message)
                            'End Try
                        End If


                        If (pVal.ItemUID = "mat1") And (pVal.ColUID = "txtuomcode") Then
                            Dim paracd2 As String
                            oCombo = matcol5.Cells.Item(pVal.Row).Specific
                            oEdit = matcol6.Cells.Item(pVal.Row).Specific
                            paracd2 = oCombo.Selected.Value
                            oEdit.Value = oCombo.Selected.Description
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        If pVal.ItemUID = "txtcode" Then
                            Choose(FormUID, pVal)
                        End If
                End Select
            End If
        Catch ex As Exception

        End Try
    End Sub

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
                Dim k As Integer
                If strCFL = "CFL_3" Then
                    oForm.DataSources.DBDataSources.Item("@SST_QCSTANDHDR").SetValue("U_itemcode", 0, oDT.GetValue("ItemCode", 0))
                    oForm.DataSources.DBDataSources.Item("@SST_QCSTANDHDR").SetValue("U_itemname", 0, oDT.GetValue("ItemName", 0))

                    RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    strsql = "select a.U_itemcode,a.u_itemname,c.Code,c.U_paradesc,c.U_catcode,c.U_paramcat,c.U_uomcode,c.U_uomdesc from [@sst_nplanhdr] a inner join [@sst_nplandtl] b on a.Code = b.Code inner join [@sst_qcparameter] c on b.U_catcode = c.Code where a.U_itemcode = '" & oForm.Items.Item("txtcode").Specific.string & "'"
                    RS.DoQuery(strsql)
                    If RS.RecordCount > 0 Then
                        If oMatrix.RowCount > 0 Then
                            oMatrix.Clear()
                        End If
                        k = 0
                        For i = 1 To RS.RecordCount
                            oMatrix.AddRow()
                            oMatrix.GetLineData(oMatrix.RowCount)
                            oForm.DataSources.DBDataSources.Item("@SST_QCSTANDDTL").SetValue("LineId", 0, oMatrix.RowCount)
                            oForm.DataSources.DBDataSources.Item("@SST_QCSTANDDTL").SetValue("U_catcode", 0, RS.Fields.Item("U_catcode").Value)
                            oForm.DataSources.DBDataSources.Item("@SST_QCSTANDDTL").SetValue("U_paramcat", 0, RS.Fields.Item("U_paramcat").Value)
                            oForm.DataSources.DBDataSources.Item("@SST_QCSTANDDTL").SetValue("U_paracode", 0, RS.Fields.Item("Code").Value)
                            oForm.DataSources.DBDataSources.Item("@SST_QCSTANDDTL").SetValue("U_paradesc", 0, RS.Fields.Item("U_paradesc").Value)
                            oForm.DataSources.DBDataSources.Item("@SST_QCSTANDDTL").SetValue("U_uomcode", 0, RS.Fields.Item("U_uomcode").Value)
                            oForm.DataSources.DBDataSources.Item("@SST_QCSTANDDTL").SetValue("U_uomdesc", 0, RS.Fields.Item("U_uomdesc").Value)
                            oMatrix.SetLineData(oMatrix.RowCount)
                        Next
                    Else
                        objAddOn.SBO_Application.SetStatusBarMessage("No matching records....", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    End If
                End If
            Catch ex As Exception
                'objAddOn.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End If
        RS = Nothing
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

        If pVal.MenuUID = "DelRow" Then
            DeleteRow(pVal, objAddOn.SBO_Application.Forms.ActiveForm.UniqueID)
        End If
        If pVal.MenuUID = "1290" Or pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1291" Then
            Disable()
        End If
    End Sub

    Private Sub Disable()
      
        oForm.Items.Item("txtcode").Enabled = False
      
    End Sub

End Class
