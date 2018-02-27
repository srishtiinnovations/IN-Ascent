
' Author                   Created Date
' Sankar                   31/01/2012
Public Class clsSamplingLevel
    Private oForm As SAPbouiCOM.Form
    Private oMatrix As SAPbouiCOM.Matrix
    Dim Optn1, Optn2 As SAPbouiCOM.OptionBtn
    Dim ChkBx As SAPbouiCOM.CheckBox
    Dim objRS As SAPbobsCOM.Recordset
    Dim strSQL As String
    Dim oDT As SAPbouiCOM.DataTable
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
    Private matcol10 As SAPbouiCOM.Column
    Private oColumns As SAPbouiCOM.Columns
    Dim oEdit As SAPbouiCOM.EditText
    Dim oCombo, oCombo1 As SAPbouiCOM.ComboBox
    Dim i As Integer
    Private paracd As String
    Dim RowNo As Integer = 0
    Public Const formtype As String = "Frm_SamplingLevel"
    Private oEdit1, oEdit2 As SAPbouiCOM.EditText

    Public Sub LoadScreen()
        Try
            oForm = objAddOn.objUIXml.LoadScreenXML("Frm_SmplLvl.xml", SST.enuResourceType.Embeded, formtype)
            ChkBx = oForm.Items.Item("24").Specific
            oMatrix = oForm.Items.Item("11").Specific
            oColumns = oMatrix.Columns
            matcol1 = oColumns.Item("txtlot")
            LoadCombo(matcol1)
            oForm.DataBrowser.BrowseBy = "8"
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadCombo(ByVal oCombo As SAPbouiCOM.Column)
        Try
            objRS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRS.DoQuery("SELECT Code,Name FROM [@SST_SAMPLINGLEVEL] order by  docentry asc")
            objRS.MoveFirst()
            While objRS.EoF = False
                oCombo.ValidValues.Add(objRS.Fields.Item("Code").Value, objRS.Fields.Item("Name").Value)
                objRS.MoveNext()
            End While
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
                                    oForm.Items.Item("8").Specific.value = objAddOn.objGenFunc.GetDocNum("SST_SLM")
                                    oEdit1 = oForm.Items.Item("8").Specific
                                    oEdit1.Value = objAddOn.objGenFunc.GetDocNum("SST_SLM")
                                End If
                            End If
                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                If Validate(pVal) = False Then
                                    BubbleEvent = False
                                End If
                            End If
                        End If

                        If pVal.ItemUID = "12" Then
                            oMatrix = oForm.Items.Item("11").Specific
                            If oMatrix.RowCount = 0 Then
                                oMatrix.AddRow()
                                oMatrix.GetLineData(oMatrix.RowCount)
                                oForm.DataSources.DBDataSources.Item("@SST_SLMDTL").SetValue("LineId", 0, oMatrix.RowCount)
                                oMatrix.SetLineData(oMatrix.RowCount)
                            Else
                                Try
                                    For i = 1 To oMatrix.RowCount
                                        oCombo = oMatrix.Columns.Item("txtlot").Cells.Item(i).Specific
                                        If oCombo.Selected Is Nothing Then
                                            objAddOn.SBO_Application.SetStatusBarMessage("Select Lot or Batch")
                                            Exit Sub
                                        Else
                                            If oCombo.Selected.Value = "" Then
                                                objAddOn.SBO_Application.SetStatusBarMessage("Select Lot or Batch")
                                                Exit Sub
                                            End If
                                        End If
                                        matcol1 = oColumns.Item("txtsmpsize")
                                        If matcol1.Cells.Item(i).Specific.string = "" Then
                                            objAddOn.SBO_Application.SetStatusBarMessage("Enter Sample Size")
                                            Exit Sub
                                        End If

                                    Next
                                Catch ex As Exception
                                    objAddOn.SBO_Application.SetStatusBarMessage("Enter All Mandatory Fields.....")
                                    Exit Sub
                                End Try

                                oForm.DataSources.DBDataSources.Item("@SST_SLMDTL").Clear()
                                oMatrix.AddRow()
                                oMatrix.GetLineData(oMatrix.RowCount)
                                oForm.DataSources.DBDataSources.Item("@SST_SLMDTL").SetValue("LineId", 0, oMatrix.RowCount)
                                oMatrix.SetLineData(oMatrix.RowCount)
                            End If
                        End If
                End Select

            Else
                Select Case pVal.EventType

                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        If (pVal.ItemUID = "11") And (pVal.ColUID = "txtlot") Then
                            matcol1 = oColumns.Item("txtlot")
                            oCombo = matcol1.Cells.Item(pVal.Row).Specific
                            oEdit = matcol1.Cells.Item(pVal.Row).Specific
                            oEdit.Value = oCombo.Selected.Description
                        End If

                End Select
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Function Validate(ByVal pVal) As Boolean
        oForm = objAddOn.SBO_Application.Forms.Item(pVal.FormUID)

        If oForm.Items.Item("4").Specific.string = "" Or oForm.Items.Item("6").Specific.string = "" Then
            objAddOn.SBO_Application.SetStatusBarMessage("Enter From and To Fields..")
            Return False
        End If
        If oMatrix.RowCount = 0 Then
            objAddOn.SBO_Application.SetStatusBarMessage("Matrix  should not be left blank")
            Return False
        Else
            For i = 1 To oMatrix.RowCount
                oCombo = oMatrix.Columns.Item("txtlot").Cells.Item(i).Specific
                If oCombo.Selected Is Nothing Then
                    objAddOn.SBO_Application.SetStatusBarMessage("Select Lot or Batch")
                    Exit Function
                Else
                    If oCombo.Selected.Value = "" Then
                        objAddOn.SBO_Application.SetStatusBarMessage("Select Lot or Batch")
                        Exit Function
                    End If

                    matcol1 = oColumns.Item("txtsmpsize")
                    If matcol1.Cells.Item(i).Specific.string = "" Then
                        objAddOn.SBO_Application.SetStatusBarMessage("Enter Sample Size")
                        Exit Function
                    End If
                End If
            Next
        End If

        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            objRS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRS.DoQuery("select * from [@SST_slmhdr] where U_frmbth = '" & oForm.Items.Item("4").Specific.string & "' and U_tobth = '" & oForm.Items.Item("6").Specific.string & "' ")
            If objRS.RecordCount > 0 Then
                objAddOn.SBO_Application.SetStatusBarMessage("Data Already Exists")
                Return False
            End If
        End If

        Return True
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
        'oForm.Items.Item("6").Enabled = False
    End Sub

    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        oMatrix = oForm.Items.Item("11").Specific
        If (eventInfo.BeforeAction = True) Then
            If eventInfo.ItemUID = "11" Then
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

            oMatrix = oForm.Items.Item("11").Specific
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
