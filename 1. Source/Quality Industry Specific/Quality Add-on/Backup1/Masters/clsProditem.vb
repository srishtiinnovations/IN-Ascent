' Author                     Created Date
' Manimaran                   20/11/2010
Public Class clsProditem
    Private oForm As SAPbouiCOM.Form
    Dim RS As SAPbobsCOM.Recordset
    Private oMatrix As SAPbouiCOM.Matrix
    Dim i As Integer
    Private oColumns As SAPbouiCOM.Columns
    Private oColumn As SAPbouiCOM.Column
    Private matcol1 As SAPbouiCOM.Column
    Private matcol2 As SAPbouiCOM.Column
    Private matcol3 As SAPbouiCOM.Column
    Private matcol4 As SAPbouiCOM.Column
    Private matcol5 As SAPbouiCOM.Column
    Private matcol6 As SAPbouiCOM.Column
    Private matcol7 As SAPbouiCOM.Column
    Private matcol8 As SAPbouiCOM.Column
    Private matcol9 As SAPbouiCOM.Column
    Dim oDT As SAPbouiCOM.DataTable
    Dim oCombo, oCombo1, oCombo2 As SAPbouiCOM.ComboBox
    Private txtcode As SAPbouiCOM.EditText
    Private txtdesc As SAPbouiCOM.EditText
    Private txtslno As SAPbouiCOM.EditText
    Private txtstgdesc As SAPbouiCOM.EditText
    Private cbostage As SAPbouiCOM.ComboBox
    Dim RowNo As Integer = 0
    Private paracd As String
    Private paracd1 As String

    Public Const formtype As String = "Frm_PrdItm"
    Public Sub LoadScreen()
        oForm = objAddOn.objUIXml.LoadScreenXML("Frm_Proditem.xml", SST.enuResourceType.Embeded, formtype)

        txtcode = oForm.Items.Item("txtcode").Specific
        txtdesc = oForm.Items.Item("txtdesc").Specific
        txtslno = oForm.Items.Item("txtslno").Specific
        cbostage = oForm.Items.Item("cbostage").Specific
        txtstgdesc = oForm.Items.Item("txtstgdesc").Specific
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
        Addcombocategory(matcol1)
        Addcomboparam(matcol3)
        oForm.DataBrowser.BrowseBy = "txtslno"
    End Sub
    Private Sub Addcomboparam(ByVal oCombo As SAPbouiCOM.Column)
        Try
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            RS.DoQuery("select Code,U_paradesc from [@SST_QCPARAMETER]")
            RS.MoveFirst()
            While RS.EoF = False
                oCombo.ValidValues.Add(RS.Fields.Item("Code").Value, RS.Fields.Item("U_paradesc").Value)
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
                                    oForm.Items.Item("txtslno").Specific.value = objAddOn.objGenFunc.GetDocNum("SST_PSTAN")
                                    oForm.Items.Item("14").Specific.value = objAddOn.objGenFunc.GetDocNum("SST_PSTAN")
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
                                oForm.DataSources.DBDataSources.Item("@SST_PRDSTDDTL").SetValue("LineId", 0, oMatrix.RowCount)
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
                                oForm.DataSources.DBDataSources.Item("@SST_PRDSTDDTL").Clear()
                                oMatrix.AddRow()
                                oMatrix.GetLineData(oMatrix.RowCount)
                                oForm.DataSources.DBDataSources.Item("@SST_PRDSTDDTL").SetValue("LineId", 0, oMatrix.RowCount)
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
                        If (pVal.ItemUID = "mat1") And (pVal.ColUID = "txtcatcode") Then
                            Dim oEdit As SAPbouiCOM.EditText
                            oCombo = matcol1.Cells.Item(pVal.Row).Specific
                            oEdit = matcol2.Cells.Item(pVal.Row).Specific
                            paracd = oCombo.Selected.Value
                            oEdit.Value = oCombo.Selected.Description


                            '*********** For the category the corresponding parameters will be loaded in the Parameters combo 
                            Try
                                Dim oCombo1 As SAPbouiCOM.Column = matcol3

                                RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                RS.DoQuery("select Code,U_paradesc,U_uomcode,U_uomdesc from [@SST_QCPARAMETER] where U_catcode='" & paracd & "'")
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


                        '********* Displays the Description of the Parameter selected from the combo ********
                        If (pVal.ItemUID = "mat1") And (pVal.ColUID = "txtparcode") Then

                            Dim oEdit As SAPbouiCOM.EditText

                            oCombo = matcol3.Cells.Item(pVal.Row).Specific
                            oEdit = matcol4.Cells.Item(pVal.Row).Specific

                            paracd1 = oCombo.Selected.Value
                            oEdit.Value = oCombo.Selected.Description


                            '*********** For the Parameter the corresponding parameters will be loaded in the Parameter Uom combo 
                            Try
                             

                                Dim oCombo1 As SAPbouiCOM.Column = matcol5

                                RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                               
                                RS.DoQuery("select U_uomcode,U_uomdesc from [@SST_QCPARAMETER] where Code='" & paracd1 & "'")
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

                        '********* Displays the Description of the Parameter Uom selected from the combo ********
                        If (pVal.ItemUID = "mat1") And (pVal.ColUID = "txtuomcode") Then
                            Dim oEdit As SAPbouiCOM.EditText
                            Dim paracd2 As String
                            oCombo = matcol5.Cells.Item(pVal.Row).Specific
                            oEdit = matcol6.Cells.Item(pVal.Row).Specific
                            paracd2 = oCombo.Selected.Value
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
        Dim strCFL, strsql As String
        RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objCFLEvent = pVal
        strCFL = objCFLEvent.ChooseFromListUID
        oForm = objAddOn.SBO_Application.Forms.Item(FormUID)
        oDT = objCFLEvent.SelectedObjects
        If objCFLEvent.BeforeAction = False Then
            Try
                If strCFL = "CFL_3" Then
                    oForm.DataSources.DBDataSources.Item("@SST_PRDSTDHDR").SetValue("U_itemcode", 0, oDT.GetValue("ItemCode", 0))
                    oForm.DataSources.DBDataSources.Item("@SST_PRDSTDHDR").SetValue("U_itemname", 0, oDT.GetValue("ItemName", 0))
                    'strsql = " select ItemName from oitm where itemcode = '" & oDT.GetValue("Code", 0) & "'"
                    'RS.DoQuery(strsql)
                    'If RS.RecordCount > 0 Then
                    '    oForm.DataSources.DBDataSources.Item("@SST_PRDSTDHDR").SetValue("U_itemname", 0, RS.Fields.Item(0).Value)
                    'End If
                End If

            Catch ex As Exception
                'objAddOn.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End If
        RS = Nothing
    End Sub
    Private Function Validate(ByVal pVal) As Boolean
        oForm = objAddOn.SBO_Application.Forms.Item(pVal.FormUID)
        oMatrix = oForm.Items.Item("mat1").Specific

        Dim oEdit As SAPbouiCOM.EditText
        Dim oEdit1 As SAPbouiCOM.EditText
        Dim oEdit2 As SAPbouiCOM.EditText

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


            Dim oEdit3 As SAPbouiCOM.EditText
            oEdit3 = oForm.Items.Item("txtstgdesc").Specific
            If oEdit3.Value = "" Or oEdit3.Value = Nothing Then

                oEdit3.Active = True
                objAddOn.SBO_Application.SetStatusBarMessage("Stage Code should not be left blank")
                Return False
            End If
            If oEdit3.Value = "Define New" Then
                objAddOn.SBO_Application.SetStatusBarMessage("Select Stage Code ")
                Return False
            End If

            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                Dim RS As SAPbobsCOM.Recordset
                RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                RS.DoQuery("select U_itemcode, U_stage from [@SST_PRDSTDHDR]  where U_itemcode='" & txtcode.Value & "' and U_stage='" & cbostage.Selected.Value.ToString() & "' ")
                If RS.RecordCount > 0 Then
                    objAddOn.SBO_Application.SetStatusBarMessage("Item Code Already Exists")
                    Return False
                End If
            End If


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
                    Try

                        oCombo = matcol1.Cells.Item(i).Specific
                        oCombo1 = matcol3.Cells.Item(i).Specific
                        oCombo2 = matcol5.Cells.Item(i).Specific

                        If oCombo.Selected Is Nothing Then
                            objAddOn.SBO_Application.SetStatusBarMessage("Select Category Code")
                            Return False
                        Else
                            If oCombo.Selected.Value = "" Then
                                objAddOn.SBO_Application.SetStatusBarMessage("Select Category Code")
                                Return False
                            End If
                        End If

                        If matcol4.Cells.Item(i).Specific.string = "" Then
                            objAddOn.SBO_Application.SetStatusBarMessage("Select Parameter Code")
                            Return False

                        End If

                        If matcol6.Cells.Item(i).Specific.string = "" Then
                            objAddOn.SBO_Application.SetStatusBarMessage("Select UOM Code")
                            Return False
                        End If
                    Catch ex As Exception
                        objAddOn.SBO_Application.SetStatusBarMessage("Enter All Mandatory Fields.....")
                        Return False
                    End Try
                Next
            End If

        Catch ex As Exception

        End Try
        Return True
    End Function
    Private Sub Addcombocategory(ByVal oCombo As SAPbouiCOM.Column)
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
        If pVal.MenuUID = "1282" Then
            oForm.Items.Item("txtcode").Enabled = True
        End If
    End Sub
    Private Sub Disable()

        oForm.Items.Item("txtcode").Enabled = False

    End Sub
End Class
