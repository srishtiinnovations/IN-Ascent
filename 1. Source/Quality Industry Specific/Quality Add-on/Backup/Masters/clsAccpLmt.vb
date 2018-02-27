' Author                    Created Date
' Sankar                    28/01/2012
Public Class clsAccpLmt
    Private oForm As SAPbouiCOM.Form
    Public Const formtype As String = "Frm_AccpLmt"
    Dim rs As SAPbobsCOM.Recordset
    Dim strSQL As String
    Dim oDT As SAPbouiCOM.DataTable
    Dim objRS As SAPbobsCOM.Recordset
    Dim oCons As SAPbouiCOM.Conditions
    Dim oCon As SAPbouiCOM.Condition
    Private oMatrix As SAPbouiCOM.Matrix
    Private oColumns As SAPbouiCOM.Columns
    Private matcol1 As SAPbouiCOM.Column
    Private matcol2 As SAPbouiCOM.Column
    Private matcol3 As SAPbouiCOM.Column
    Private matcol4 As SAPbouiCOM.Column
    Dim ChkBx As SAPbouiCOM.CheckBox
    Dim oEdit As SAPbouiCOM.EditText
    Dim oCombo1 As SAPbouiCOM.ComboBox
    Dim i As Integer
    Private lot As String
    Dim RowNo As Integer = 0
    Private oEdit1, oEdit2 As SAPbouiCOM.EditText

    Public Sub LoadScreen()
        'oForm = objAddOn.objUIXml.LoadScreenXML("Frm_AccLmt.xml", SST.enuResourceType.Embeded, formtype)
        oForm = objAddOn.objUIXml.LoadScreenXML("Frm_AccLmt_test.xml", SST.enuResourceType.Embeded, formtype)
        ChkBx = oForm.Items.Item("6").Specific
        oMatrix = oForm.Items.Item("7").Specific
        oColumns = oMatrix.Columns
        matcol1 = oColumns.Item("txtsmp")
        LoadCombo1(matcol1)

        oForm.DataBrowser.BrowseBy = "8"
    End Sub

    Private Sub LoadCombo1(ByVal oCombo1 As SAPbouiCOM.Column)
        Try
            objRS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            strSQL = "SELECT Code,Name FROM [@SST_ACCEPETLIMIT]"
            objRS.DoQuery(strSQL)

            If objRS.RecordCount > 0 Then
                While Not objRS.EoF
                    oCombo1.ValidValues.Add(objRS.Fields.Item(0).Value, objRS.Fields.Item(1).Value)
                    objRS.MoveNext()
                End While
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

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
                                    oForm.Items.Item("8").Specific.value = objAddOn.objGenFunc.GetDocNum("SST_AQL")
                                    oEdit1 = oForm.Items.Item("8").Specific
                                    oEdit1.Value = objAddOn.objGenFunc.GetDocNum("SST_AQL")
                                End If
                            End If
                        End If

                        If pVal.ItemUID = "12" Then
                            oMatrix = oForm.Items.Item("7").Specific
                            If oMatrix.RowCount = 0 Then
                                oMatrix.AddRow()
                                oMatrix.GetLineData(oMatrix.RowCount)
                                oForm.DataSources.DBDataSources.Item("@SST_AQLDTL").SetValue("LineId", 0, oMatrix.RowCount)
                                oMatrix.SetLineData(oMatrix.RowCount)
                            Else
                                oForm.DataSources.DBDataSources.Item("@SST_AQLDTL").Clear()
                                oMatrix.AddRow()
                                oMatrix.GetLineData(oMatrix.RowCount)
                                oForm.DataSources.DBDataSources.Item("@SST_AQLDTL").SetValue("LineId", 0, oMatrix.RowCount)
                                oMatrix.SetLineData(oMatrix.RowCount)
                            End If
                        End If

                End Select
            Else
                Select Case pVal.EventType

                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                        'If (pVal.ItemUID = "7") And (pVal.ColUID = "txtsmpsize") Then
                        '    matcol4 = oColumns.Item("txtsmpsize")
                        '    oCombo1 = matcol4.Cells.Item(pVal.Row).Specific
                        '    'oEdit = matcol4.Cells.Item(pVal.Row).Specific
                        '    oEdit.Value = oCombo1.Selected.Description
                        'End If

                End Select
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Function Validate(ByVal pVal) As Boolean
        oForm = objAddOn.SBO_Application.Forms.Item(pVal.FormUID)
        Try

            If oForm.Items.Item("4").Specific.string = "" Then
                objAddOn.SBO_Application.SetStatusBarMessage("Enter Sampling Size..")
                Return False
            End If
            If oMatrix.RowCount = 0 Then
                objAddOn.SBO_Application.SetStatusBarMessage("Matrix  should not be left blank")
                Return False
            Else
                For i = 1 To oMatrix.RowCount
                    'If oCombo1.Selected.Value = "" Then
                    '    objAddOn.SBO_Application.SetStatusBarMessage("Select Sample Size")
                    '    Exit Function
                    'End If
                    matcol2 = oColumns.Item("txtaccp")
                    If matcol2.Cells.Item(i).Specific.string = "" Then
                        objAddOn.SBO_Application.SetStatusBarMessage("Enter Accepted")
                        Exit Function
                    End If
                    matcol3 = oColumns.Item("txtrejec")
                    If matcol3.Cells.Item(i).Specific.string = "" Then
                        objAddOn.SBO_Application.SetStatusBarMessage("Enter Rejected")
                        Exit Function
                    End If
                Next
            End If

            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                objRS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRS.DoQuery("select * from [@SST_aqlhdr] where U_smpsize = '" & oForm.Items.Item("4").Specific.string & "' ")
                If objRS.RecordCount > 0 Then
                    objAddOn.SBO_Application.SetStatusBarMessage("Data Already Exists")
                    Return False
                End If
            End If

            Return True

        Catch ex As Exception
        End Try
    End Function

    Public Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        If pVal.MenuUID = "1290" Or pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1291" Then
            Disable()
        End If

        If pVal.MenuUID = "DelRow" Then
            DeleteRow(pVal, objAddOn.SBO_Application.Forms.ActiveForm.UniqueID)
        End If

    End Sub

    Private Sub Disable()
        'oForm.Items.Item("4").Enabled = False
    End Sub

    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        oMatrix = oForm.Items.Item("7").Specific
        If (eventInfo.BeforeAction = True) Then
            If eventInfo.ItemUID = "7" Then
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
        Try
            oForm = objAddOn.SBO_Application.Forms.Item(FormUID)

            oMatrix = oForm.Items.Item("7").Specific
            If oMatrix.RowCount > 0 Then
                oMatrix.DeleteRow(RowNo)
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

End Class


