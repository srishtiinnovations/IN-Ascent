' Author                     Created Date
' Manimaran                   19/11/2010
Public Class ClsReaMst
    Private oForm As SAPbouiCOM.Form
    Public Const formtype As String = "Frm_Reas"
    Dim ocatDesc As SAPbouiCOM.EditText
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim RS As SAPbobsCOM.Recordset
    Dim oEdit As SAPbouiCOM.EditText
    Dim oEdit1 As SAPbouiCOM.EditText
    Dim oEdit3 As SAPbouiCOM.EditText

    Public Sub LoadScreen()
        oForm = objAddOn.objUIXml.LoadScreenXML("Frm_Reason.xml", SST.enuResourceType.Embeded, formtype)
        oForm.DataBrowser.BrowseBy = "txtcode"
        Addcombo()
    End Sub
    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try

            ocatDesc = oForm.Items.Item("txtcatdesc").Specific
            oCombo = oForm.Items.Item("cbocate").Specific

            If (pVal.Before_Action = True) Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        oForm = objAddOn.SBO_Application.Forms.Item(pVal.FormUID)
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                            '*************ItemCode Validation **********
                            If Validate(pVal) = False Then
                                BubbleEvent = False
                            End If
                        End If
                End Select
               
            End If



            '******** This is used to display the description of the Category Selected


            If pVal.Before_Action = False Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        If (pVal.ItemUID = "cbocate") Then
                            ocatDesc.Value = oCombo.Selected.Description
                        End If
                End Select
               
            End If
        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub Addcombo()

        oCombo = oForm.Items.Item("cbocate").Specific
        Try
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            RS.DoQuery("select Code,Name from [@SST_PRDCAT]")
            RS.MoveFirst()
            While RS.EoF = False
                oCombo.ValidValues.Add(RS.Fields.Item("Code").Value, RS.Fields.Item("Name").Value)
                RS.MoveNext()
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Function Validate(ByVal pVal) As Boolean

        oForm = objAddOn.SBO_Application.Forms.Item(pVal.FormUID)
        oEdit3 = oForm.Items.Item("txtcatdesc").Specific
        oEdit = oForm.Items.Item("txtcode").Specific
        oEdit1 = oForm.Items.Item("txtdesc").Specific
        oCombo = oForm.Items.Item("cbocate").Specific

        '**************Mandatory For Code*****************

        If oEdit.Value = "" Or oEdit.Value = Nothing Then
            oEdit.Active = True
            objAddOn.SBO_Application.SetStatusBarMessage("Code should not be left blank")
            Return False
        End If

        '**************Mandatory For Description*****************

        If oEdit1.Value = "" Or oEdit1.Value = Nothing Then
            oEdit1.Active = True
            objAddOn.SBO_Application.SetStatusBarMessage("Description should not be left blank")
            Return False
        End If

        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            RS = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            RS.DoQuery("select Code from [@SST_QCREASON]  where Code='" & oEdit.Value & "'  ")
            If RS.RecordCount > 0 Then
                objAddOn.SBO_Application.SetStatusBarMessage("Code Already Exists")
                Return False
            End If
        End If

       

        '****************** Mandatory Category Code ***************

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


        If oEdit3.Value = "" Or oEdit3.Value = Nothing Then
            oEdit3.Active = True
            objAddOn.SBO_Application.SetStatusBarMessage("Category Description should not be left blank")
            Return False
        End If
        Return True
    End Function
End Class
