' Author                     Created Date
' Manimaran                   01/12/2010
Public Class clsSetUp
    Private oForm As SAPbouiCOM.Form
    Public Const formtype As String = "Frm_STP"
    Dim rs As SAPbobsCOM.Recordset
    Dim strSQL As String
    Dim oDT As SAPbouiCOM.DataTable
    Public Sub LoadScreen()
        oForm = objAddOn.objUIXml.LoadScreenXML("Frm_WHSetUp.xml", SST.enuResourceType.Embeded, formtype)

        oForm.DataBrowser.BrowseBy = "1000002"
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
                                If Check() = False Then
                                    If Validate(pVal) = False Then
                                        BubbleEvent = False
                                    Else
                                        oForm.Items.Item("1000002").Specific.value = objAddOn.objGenFunc.GetDocNum("SST_SET")
                                    End If
                                Else
                                    objAddOn.SBO_Application.SetStatusBarMessage("Record Already Exists...")
                                    BubbleEvent = False
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
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        If pVal.ItemUID = "6" Or pVal.ItemUID = "1000001" Or pVal.ItemUID = "8" Then
                            Choose(FormUID, pVal)
                        End If
                End Select
            End If
        Catch ex As Exception
           
        End Try
    End Sub
    Private Function Check() As Boolean
        rs = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strsql = "select * from [@SST_SETUP]"
        rs.DoQuery(strSQL)
        If rs.RecordCount > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Function Validate(ByVal pVal) As Boolean
        oForm = objAddOn.SBO_Application.Forms.Item(pVal.FormUID)

        If oForm.Items.Item("6").Specific.string = "" Or oForm.Items.Item("1000001").Specific.string = "" Or oForm.Items.Item("8").Specific.string = "" Then
            objAddOn.SBO_Application.SetStatusBarMessage("Enter All Mandatory Fields..")
            Return False
        End If
        Return True
    End Function

    Private Sub Choose(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent)
        Dim objCFLEvent As SAPbouiCOM.ChooseFromListEvent
        Dim strCFL As String
        objCFLEvent = pval
        strCFL = objCFLEvent.ChooseFromListUID
        oForm = objAddOn.SBO_Application.Forms.Item(FormUID)
        oDT = objCFLEvent.SelectedObjects
        If objCFLEvent.BeforeAction = False Then
            Try
                If strCFL = "CFL_2" Then
                    oForm.DataSources.DBDataSources.Item("@SST_SETUP").SetValue("U_RGLWH", 0, oDT.GetValue("WhsCode", 0))
                ElseIf strCFL = "CFL_3" Then
                    oForm.DataSources.DBDataSources.Item("@SST_SETUP").SetValue("U_RJTWH", 0, oDT.GetValue("WhsCode", 0))
                ElseIf strCFL = "CFL_4" Then
                    oForm.DataSources.DBDataSources.Item("@SST_SETUP").SetValue("U_RWKWH", 0, oDT.GetValue("WhsCode", 0))
                End If
            Catch ex As Exception
                'objAddOn.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End If

    End Sub
End Class
