'''' <summary>
'''' Author                     Created Date
'''' Suresh                      19/12/2008
'''' <remarks> This class is used for entering the Stoppage Details.</remarks>
Public Class Stoppage
    Inherits GeneralLib
    ''' <summary>
    ''' Variable Declaration
    ''' </summary>
    ''' <remarks></remarks>
#Region "Variable Declaration"
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private SboGuiApi As New SAPbouiCOM.SboGuiApi
    Private oCompany As SAPbobsCOM.Company
    '**************************DataSource************************************
    Private oParentDB As SAPbouiCOM.DBDataSource
    '**************************Form************************************
    Private oForm As SAPbouiCOM.Form
    '**************************Items - EditText************************************
    Private oSGCodeTxt, oSGDescTxt, oCatNameTxt, oPlnTimeeTxt, oInfo1Txt, oInfo2Txt, oRemarksTxt As SAPbouiCOM.EditText
    '**************************Items - Combo************************************
    Private oCatCodeCombo As SAPbouiCOM.ComboBox
    '**************************Items - CheckBox************************************
    Private oActiveCheck As SAPbouiCOM.CheckBox

    Private BoolCatType As Boolean = True
#End Region
    ''' <summary>
    ''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmStoppage.srf") method is called to load the Machine Group form.
    ''' Drawform() method is called to Initialize the form,Datasources and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company)
        MyBase.New(aSBO_Application, aCompany)
        SBO_Application = aSBO_Application
        oCompany = aCompany
        LoadFromXML("FrmStoppage.srf")
        DrawForm()
        oForm.DataBrowser.BrowseBy = "txtscode"
    End Sub
    ''' <summary>
    ''' Connecting the application through connection string.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetApplication()
        Dim sConnectionString As String
        SboGuiApi = New SAPbouiCOM.SboGuiApi
        sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"
        SboGuiApi.Connect(sConnectionString)
        SboGuiApi.AddonIdentifier = "5645523035446576656C6F706D656E743A453038373933323333343581F0D8D8C45495472FC628EF425AD5AC2AEDC411"
        SBO_Application = SboGuiApi.GetApplication()
    End Sub
    ''' <summary>
    ''' Initializing the instance of the active form to the form object.
    ''' Initializing the Datasources.
    ''' InitializeFormComponent() method is called to initialize the items.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub DrawForm()
        Try
            oForm = SBO_Application.Forms.Item(SBO_Application.Forms.ActiveForm.UniqueID)
            oParentDB = oForm.DataSources.DBDataSources.Item("@PSSIT_OSGE")
            oForm.Freeze(True)
            InitializeFormComponent()
            oForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Configuring the items/controls in the form(.srf) by bounding to the object and to the DBDatasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitializeFormComponent()
        Try
            oSGCodeTxt = oForm.Items.Item("txtscode").Specific
            oSGCodeTxt.DataBind.SetBound(True, "@PSSIT_OSGE", "Code")

            oSGDescTxt = oForm.Items.Item("txtsname").Specific
            oSGDescTxt.DataBind.SetBound(True, "@PSSIT_OSGE", "U_Stopname")

            oCatCodeCombo = oForm.Items.Item("cmbccode").Specific
            oCatCodeCombo.DataBind.SetBound(True, "@PSSIT_OSGE", "U_Catcode")
            CatCombo()

            oCatNameTxt = oForm.Items.Item("txtcname").Specific
            oCatNameTxt.DataBind.SetBound(True, "@PSSIT_OSGE", "U_Catname")

            oPlnTimeeTxt = oForm.Items.Item("txtplantim").Specific
            oPlnTimeeTxt.DataBind.SetBound(True, "@PSSIT_OSGE", "U_Plantime")

            oInfo1Txt = oForm.Items.Item("txtadnl1").Specific
            oForm.Items.Item("lbladnl1").Visible = False
            oForm.Items.Item("txtadnl1").Visible = False
            oInfo1Txt.DataBind.SetBound(True, "@PSSIT_OSGE", "U_Adnl1")

            oInfo2Txt = oForm.Items.Item("txtadnl2").Specific
            oForm.Items.Item("lbladnl2").Visible = False
            oForm.Items.Item("txtadnl2").Visible = False
            oInfo2Txt.DataBind.SetBound(True, "@PSSIT_OSGE", "U_Adnl2")

            oRemarksTxt = oForm.Items.Item("txtremark").Specific
            oRemarksTxt.DataBind.SetBound(True, "@PSSIT_OSGE", "U_Remarks")

            oActiveCheck = oForm.Items.Item("chkactive").Specific
            oActiveCheck.DataBind.SetBound(True, "@PSSIT_OSGE", "U_Active")
            oActiveCheck.Checked = True

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This is used to Load the Category in the Combo
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CatCombo()
        Dim oRs As SAPbobsCOM.Recordset
        Try
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery("select Code,U_Catname from [@PSSIT_OSCY] where code is not null and U_Catname is not Null")
            oRs.MoveFirst()
            If oCatCodeCombo.ValidValues.Count > 0 Then
                For i As Int16 = oCatCodeCombo.ValidValues.Count - 1 To 0 Step -1
                    oCatCodeCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            For i As Int16 = 0 To oRs.RecordCount - 1
                oCatCodeCombo.ValidValues.Add(oRs.Fields.Item(0).Value, oRs.Fields.Item(1).Value)
                oRs.MoveNext()
            Next
            oCatCodeCombo.ValidValues.Add("Define New", "")
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' Loads the last entered Value
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CatDFN()
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim StrSql As String
        Try
            If BoolCatType = False Then
                If Not oCatCodeCombo Is Nothing Then
                    CatCombo()
                    StrSql = "select * from [@PSSIT_OSCY] where DocEntry=(Select IsNull(Max(DocEntry),0) as Code from [@PSSIT_OSCY])"
                    oRs.DoQuery(StrSql)
                    If oRs.RecordCount > 0 Then
                        oRs.MoveFirst()
                        oCatCodeCombo.Select(oRs.Fields.Item("U_Catname").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                        BoolCatType = True
                    End If
                End If
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        Finally
            oRs = Nothing
            GC.Collect()
        End Try

    End Sub
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            If pVal.FormUID = "FST" Then
                '*****************************Releasing the Com Object*******************************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                    SBO_Application = Nothing
                    GC.Collect()
                End If

                '******** Validation() method is called for validating the values in the edit text **********
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                        Try
                            Validation()
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End Try
                    End If
                End If

                '***** Reloads the Combo's if Define New is selected and data added in the Forms *****
                If (pVal.FormTypeEx = "FST") And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE) Then
                    CatDFN()

                End If

                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT) Then
                    '******** Work Center Type Combo Select *********
                    If (pVal.ItemUID = "cmbccode") And (pVal.BeforeAction = False) Then
                        oParentDB.SetValue("U_Catname", oParentDB.Offset, oCatCodeCombo.Selected.Description)
                        '**** Work Center Type Combo Define New Selection *****
                        If oCatCodeCombo.Selected.Value = "Define New" Then
                            LoadDefaultForm("PSSIT_SCY")
                            BubbleEvent = False
                            oParentDB.SetValue("U_Catname", oParentDB.Offset, "")
                            BoolCatType = False
                        End If

                    End If
                End If
                '********** Add Button Press ***********
                If pVal.ItemUID = "1" And (pVal.BeforeAction = False) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    Try
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            oForm.Items.Item("txtscode").Enabled = False
                            oForm.Items.Item("txtsname").Enabled = True
                        End If
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            oForm.Refresh()
                            oForm.Freeze(True)
                            oActiveCheck.Checked = True
                            oForm.Freeze(False)
                            oSGCodeTxt.Active = True
                        End If
                    Catch ex As Exception
                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    End Try
                End If
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try

    End Sub

    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Dim FormID As String
        Try
            FormID = SBO_Application.Forms.ActiveForm.UniqueID
            If pVal.MenuUID = "1281" And FormID = "FST" Then
                If pVal.BeforeAction = True Then
                    oForm.Items.Item("txtscode").Enabled = True
                    oForm.Items.Item("txtsname").Enabled = True
                End If
                If pVal.BeforeAction = False Then
                    oSGCodeTxt.Active = True
                End If
            End If
            If pVal.MenuUID = "1282" And pVal.BeforeAction = False And FormID = "FST" Then
                oForm.Freeze(True)
                oParentDB.Offset = oParentDB.Size - 1
                oActiveCheck.Checked = True
                oForm.Items.Item("txtscode").Enabled = True
                oForm.Items.Item("txtsname").Enabled = True
                oForm.Freeze(False)
                oSGCodeTxt.Active = True

            End If
            '*****************************Navigation*******************************
            If (pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") And pVal.BeforeAction = False And FormID = "FST" Then
                Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Try
                    oRs.DoQuery("Select * from [@PSSIT_OSGE]")
                    If oRs.RecordCount > 0 Then
                        oForm.Items.Item("txtscode").Enabled = False
                        oForm.Items.Item("txtsname").Enabled = True
                    Else
                        oForm.Items.Item("txtscode").Enabled = True
                        oForm.Items.Item("txtsname").Enabled = True
                        oSGCodeTxt.Active = True
                    End If
                Catch ex As Exception
                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Finally
                    oRs = Nothing
                    GC.Collect()
                End Try
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    ''' <summary>
    ''' This method is used for validating the values in the EditText.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Validation()
        Dim oRs, oRs1 As SAPbobsCOM.Recordset
        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            If oSGCodeTxt.Value.Length = 0 Then
                oSGCodeTxt.Active = True
                Throw New Exception("Stoppage Code should not be Empty")
            End If
            If oSGDescTxt.Value.Length = 0 Then
                oSGDescTxt.Active = True
                Throw New Exception("Stoppage Name should not be Empty")
            End If
            oRs.DoQuery("select Code from [@PSSIT_OSGE]  where Code= '" & oSGCodeTxt.Value & "' ")
            If oRs.RecordCount > 0 Then
                oSGCodeTxt.Active = True
                Throw New Exception("Stoppage Code Already Exist")
            End If

            oRs1.DoQuery("select U_Stopname from [@PSSIT_OSGE]  where U_Stopname= '" & oSGDescTxt.Value & "' ")
            If oRs1.RecordCount > 0 Then
                oSGDescTxt.Active = True
                Throw New Exception("Stoppage Name Already Exist")
            End If

        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            oRs1 = Nothing
            GC.Collect()
        End Try
    End Sub
End Class
