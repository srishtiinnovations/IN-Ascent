' Author                     Created Date
' Manimaran                   19/11/2010
Public Class clsParaMst
    Private oForm As SAPbouiCOM.Form
    Public Const formtype As String = "Frm_PMst"
    Dim ocatdesc As SAPbouiCOM.EditText
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim ouomdesc As SAPbouiCOM.EditText
    Dim oCombo1 As SAPbouiCOM.ComboBox
    Dim oEdit As SAPbouiCOM.EditText
    Dim oEdit1 As SAPbouiCOM.EditText
    Dim oEdit3 As SAPbouiCOM.EditText
    Dim oEdit4 As SAPbouiCOM.EditText
    Dim RS As SAPbobsCOM.Recordset
    Public Sub LoadScreen()
        oForm = objAddOn.objUIXml.LoadScreenXML("Frm_PrmMaster.xml", SST.enuResourceType.Embeded, formtype)
        oForm.DataBrowser.BrowseBy = "txtcode"
        Addcombo()
        Addcombo1()
    End Sub
    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try

            ocatdesc = oForm.Items.Item("txtcatdesc").Specific
            ouomdesc = oForm.Items.Item("txtuomdesc").Specific
            oCombo = oForm.Items.Item("cbocat").Specific
            oCombo1 = oForm.Items.Item("cbouom").Specific
            If (pVal.Before_Action = True) Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        oForm = objAddOn.SBO_Application.Forms.Item(pVal.FormUID)
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            If validate(pVal) = False Then
                                BubbleEvent = False
                            End If
                        End If
                End Select
            End If


            If (pVal.Before_Action = False) Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        If (pVal.ItemUID = "cbocat") Then
                            ocatdesc.Value = oCombo.Selected.Description
                        End If
                        If (pVal.ItemUID = "cbouom") Then
                            ouomdesc.Value = oCombo1.Selected.Description
                        End If
                End Select
            End If

        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        End Try

    End Sub
    Private Function validate(ByVal pVal) As Boolean
        '*************Code Validation **********
        oEdit = oForm.Items.Item("txtcode").Specific


        '**************Mandatory For Code*****************

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

        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            RS.DoQuery("select Code from [@SST_QCPARAMETER]  where Code='" & oEdit.Value & "'  ")
            If RS.RecordCount > 0 Then
                objAddOn.SBO_Application.SetStatusBarMessage("Code Already Exists")
                Return False
            End If
        End If

        '****************** Mandatory Category Code ***************
        oCombo = oForm.Items.Item("cbocat").Specific
        If oCombo.Selected Is Nothing Then
            oCombo.Active = True
            objAddOn.SBO_Application.SetStatusBarMessage("Category Code should not be left blank")
            Return False
        Else
            If oCombo.Selected.Value Is "" Then
                oCombo.Active = True
                objAddOn.SBO_Application.SetStatusBarMessage("Category Code should not be left blank")
                Return False
            End If
        End If

        '****************** Mandatory Category Description ***************

        oEdit3 = oForm.Items.Item("txtcatdesc").Specific
        If oEdit3.Value = "" Or oEdit3.Value = Nothing Then
            oEdit3.Active = True
            objAddOn.SBO_Application.SetStatusBarMessage("Category Description should not be left blank")
            Return False
        End If

        '****************** Mandatory Category Code ***************
        oCombo1 = oForm.Items.Item("cbouom").Specific
        If oCombo1.Selected Is Nothing Then
            oCombo.Active = True
            objAddOn.SBO_Application.SetStatusBarMessage("UOM Code should not be left blank")
            Return False
        Else
            If oCombo1.Selected.Value Is "" Then
                oCombo.Active = True
                objAddOn.SBO_Application.SetStatusBarMessage("UOM Code should not be left blank")
                Return False
            End If
        End If

        '****************** Mandatory Category Description ***************

        oEdit4 = oForm.Items.Item("txtuomdesc").Specific
        If oEdit4.Value = "" Or oEdit3.Value = Nothing Then
            oEdit4.Active = True
            objAddOn.SBO_Application.SetStatusBarMessage("UOM Description should not be left blank")
            Return False
        End If
        Return True
    End Function
    Private Sub Addcombo()
        Dim RS As SAPbobsCOM.Recordset
        oCombo = oForm.Items.Item("cbocat").Specific
        Try
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            RS.DoQuery("select Code,Name from [@SST_PARACAT]")
            RS.MoveFirst()
            While RS.EoF = False
                oCombo.ValidValues.Add(RS.Fields.Item("Code").Value, RS.Fields.Item("Name").Value)
                RS.MoveNext()
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Addcombo1()
        Dim RS As SAPbobsCOM.Recordset
        oCombo1 = oForm.Items.Item("cbouom").Specific
        Try
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            RS.DoQuery("select Code,Name from [@SST_QCUOM]")
            RS.MoveFirst()
            While RS.EoF = False
                oCombo1.ValidValues.Add(RS.Fields.Item("Code").Value, RS.Fields.Item("Name").Value)
                RS.MoveNext()
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
End Class
