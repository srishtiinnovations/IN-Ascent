'''' <summary>
'''' Author                     Created Date
'''' Suresh                      17/12/2008
'''' <remarks> This class is used for entering the Skill Group Details.</remarks>
Public Class SkillGroups
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
    '**************************ChooseFromList************************************
    Private oChWCList, oChWCBtnList, oChLAccList, oChLABtnList As SAPbouiCOM.ChooseFromList
    Private Event ChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable)
    '**************************Items - EditText************************************
    Private oActAcCodeTxt, oSGCodeTxt, oSGDescTxt, oWCNoTxt, oTWCNameTxt, oLbrRateTxt, oLBAcctCodeTxt, oLBAcctNameTxt, oInfo1Txt, oInfo2Txt, oRemarksTxt As SAPbouiCOM.EditText
    '**************************Items - Combo************************************
    Private oLbrCurrCombo As SAPbouiCOM.ComboBox
    '**************************Items - CheckBox************************************
    Private oActiveCheck As SAPbouiCOM.CheckBox
    '**************************Items - LinkedButton************************************
    Private oWCLink, oLBAcctCodeLink As SAPbouiCOM.LinkedButton
    '**************************Items - Button************************************
    Private BtnWC, BtnLAcct As SAPbouiCOM.Button
    Private sPath As String = IO.Directory.GetParent(Application.ExecutablePath).ToString

    Private WithEvents WorkCentreClass As WorkCentre

    Private oSkillCode As String
    Private oFormName As String
#End Region
    ''' <summary>
    ''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmLabourSkillGroups.srf") method is called to load the Skill Group form.
    ''' Drawform() method is called to Initialize the form,Datasources and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company, Optional ByVal aSkillCode As String = Nothing, Optional ByVal aFormName As String = Nothing)
        MyBase.New(aSBO_Application, aCompany)
        SBO_Application = aSBO_Application
        oCompany = aCompany
        oSkillCode = aSkillCode
        oFormName = aFormName
        LoadFromXML("FrmLabourSkillGroups.srf")
        DrawForm()
        If oFormName = "Labour" Or oFormName = "OprRouting" Or oFormName = "Operation" Then
            oParentDB.SetValue("Code", oParentDB.Offset, oSkillCode)
            oSGCodeTxt.Value = oSkillCode
            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        End If
        oForm.DataBrowser.BrowseBy = "txtgcode"
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
            oParentDB = oForm.DataSources.DBDataSources.Item("@PSSIT_OLGP")
            oForm.Freeze(True)
            LoadLookups()
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
            oSGCodeTxt = oForm.Items.Item("txtgcode").Specific
            oSGCodeTxt.DataBind.SetBound(True, "@PSSIT_OLGP", "Code")

            oSGDescTxt = oForm.Items.Item("txtgname").Specific
            oSGDescTxt.DataBind.SetBound(True, "@PSSIT_OLGP", "U_LGname")

            oWCNoTxt = oForm.Items.Item("txtwcode").Specific
            oWCNoTxt.DataBind.SetBound(True, "@PSSIT_OLGP", "U_WCcode")
            oWCNoTxt.ChooseFromListUID = "WCLst"
            oWCNoTxt.ChooseFromListAlias = "Code"
            oForm.Items.Item("txtwcode").LinkTo = "lnkwc"
            oForm.Items.Add("lnkwc", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnkwc").Visible = True
            oForm.Items.Item("lnkwc").LinkTo = "txtwcode"
            oForm.Items.Item("lnkwc").Top = 21
            oForm.Items.Item("lnkwc").Left = 108
            oForm.Items.Item("lnkwc").Description = "Link to" & vbNewLine & "Work Centre"
            oWCLink = oForm.Items.Item("lnkwc").Specific

            oTWCNameTxt = oForm.Items.Item("txtwcnam").Specific
            oTWCNameTxt.DataBind.SetBound(True, "@PSSIT_OLGP", "U_WCName")

            oLbrRateTxt = oForm.Items.Item("txtlabrat").Specific
            oLbrRateTxt.DataBind.SetBound(True, "@PSSIT_OLGP", "U_Labrate")

            oLbrCurrCombo = oForm.Items.Item("cmbcurncy").Specific
            oLbrCurrCombo.DataBind.SetBound(True, "@PSSIT_OLGP", "U_Currncy")
            oForm.Items.Item("cmbcurncy").Enabled = False

            oLBAcctCodeTxt = oForm.Items.Item("txtaccod").Specific
            oLBAcctCodeTxt.DataBind.SetBound(True, "@PSSIT_OLGP", "U_Accode")
            oLBAcctCodeTxt.ChooseFromListUID = "LAccLst"
            oLBAcctCodeTxt.ChooseFromListAlias = "AcctCode"
            oForm.Items.Item("txtaccod").LinkTo = "lnkaccod"
            oForm.Items.Add("lnkaccod", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnkaccod").Visible = True
            oForm.Items.Item("lnkaccod").LinkTo = "txtaccod"
            oForm.Items.Item("lnkaccod").Top = 51
            oForm.Items.Item("lnkaccod").Left = 235
            oForm.Items.Item("lnkaccod").Description = "Link to" & vbNewLine & "Chart Of Accounts"
            oLBAcctCodeLink = oForm.Items.Item("lnkaccod").Specific
            oLBAcctCodeLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_GLAccounts

            oLBAcctNameTxt = oForm.Items.Item("txtacnam").Specific
            oLBAcctNameTxt.DataBind.SetBound(True, "@PSSIT_OLGP", "U_Acname")

            oInfo1Txt = oForm.Items.Item("txtadnl1").Specific
            oForm.Items.Item("lbladnl1").Visible = False
            oForm.Items.Item("txtadnl1").Visible = False
            oInfo1Txt.DataBind.SetBound(True, "@PSSIT_OLGP", "U_Adnl1")

            oInfo2Txt = oForm.Items.Item("txtadnl2").Specific
            oForm.Items.Item("lbladnl2").Visible = False
            oForm.Items.Item("txtadnl2").Visible = False
            oInfo2Txt.DataBind.SetBound(True, "@PSSIT_OLGP", "U_Adnl2")

            oRemarksTxt = oForm.Items.Item("txtremark").Specific
            oRemarksTxt.DataBind.SetBound(True, "@PSSIT_OLGP", "U_Remarks")

            oActiveCheck = oForm.Items.Item("chkactive").Specific
            oActiveCheck.DataBind.SetBound(True, "@PSSIT_OLGP", "U_Active")
            oActiveCheck.Checked = True

            BtnWC = oForm.Items.Item("btnwc").Specific
            oForm.Items.Item("btnwc").Description = "Choose from List" & vbNewLine & "Work Centre List View"
            BtnWC.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            BtnWC.Image = sPath & "\Resources\CFL.bmp"
            BtnWC = oForm.Items.Item("btnwc").Specific
            BtnWC.ChooseFromListUID = "BtWCLst"

            BtnLAcct = oForm.Items.Item("btnlacct").Specific
            oForm.Items.Item("btnlacct").Description = "Choose from List" & vbNewLine & "Account Detail List View"
            BtnLAcct.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            BtnLAcct.Image = sPath & "\Resources\CFL.bmp"
            BtnLAcct = oForm.Items.Item("btnlacct").Specific
            BtnLAcct.ChooseFromListUID = "BtLAccLst"

            oActAcCodeTxt = oForm.Items.Item("txtaccode").Specific
            oForm.Items.Item("txtaccode").Enabled = False
            'oForm.Items.Item("txtaccode").Visible = False
            'oForm.Items.Item("lblaccode").Visible = False
            oActAcCodeTxt.DataBind.SetBound(True, "@PSSIT_OLGP", "U_ActAcCode")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#Region "CFL"
    Private ReadOnly Property ChooseFromLists() As SAPbouiCOM.ChooseFromListCollection
        Get
            Return oForm.ChooseFromLists
        End Get
    End Property
    Private Function CreateNewChooseFromListParams(ByVal MultiSelection As Boolean, ByVal ObjectType As String, ByVal UniqueID As String) As SAPbouiCOM.ChooseFromListCreationParams
        Dim ChooseFromListParameters As SAPbouiCOM.ChooseFromListCreationParams
        Try
            ChooseFromListParameters = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            ChooseFromListParameters.MultiSelection = MultiSelection
            ChooseFromListParameters.ObjectType = ObjectType
            ChooseFromListParameters.UniqueID = UniqueID
            Return ChooseFromListParameters
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Sub CreateNewConditions(ByRef ChooseFromList As SAPbouiCOM.ChooseFromList, ByVal ConditionAlias As String, ByVal ConditionOperation As SAPbouiCOM.BoConditionOperation, ByVal ConditionValue As String, Optional ByVal ConditionOpenBracket As Integer = Nothing, Optional ByVal ConditionCloseBracket As Integer = Nothing, Optional ByVal ConditionRelation As SAPbouiCOM.BoConditionRelationship = Nothing)
        Dim oConditions As SAPbouiCOM.Conditions = ChooseFromList.GetConditions()
        Dim oCondition As SAPbouiCOM.Condition = oConditions.Add
        Try
            oCondition.BracketOpenNum = ConditionOpenBracket
            oCondition.Alias = ConditionAlias
            oCondition.Operation = ConditionOperation
            oCondition.CondVal = ConditionValue
            oCondition.BracketCloseNum = ConditionCloseBracket
            oCondition.Relationship = ConditionRelation
            ChooseFromList.SetConditions(oConditions)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function CreateNewChoosefromList(ByVal ChooseFromListParameters As SAPbouiCOM.ChooseFromListCreationParams) As SAPbouiCOM.ChooseFromList
        Dim ChooseFromList As SAPbouiCOM.ChooseFromList
        Try
            ChooseFromList = ChooseFromLists.Add(ChooseFromListParameters)
            Return ChooseFromList
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ChooseFromListParameters)
        End Try
    End Function
#End Region
    ''' <summary>
    ''' Creating ChooseFromList and Setting Conditions
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadLookups()
        Try
            '***************************Work Centre-CFL**********************
            oChWCList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_WCR", "WCLst"))
            CreateNewConditions(oChWCList, "U_Active", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            oChWCBtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_WCR", "BtWCLst"))
            CreateNewConditions(oChWCBtnList, "U_Active", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            '***********************************Account-CFL****************************
            oChLAccList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "1", "LAccLst"))
            CreateNewConditions(oChLAccList, "Postable", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            oChLABtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "1", "BtLAccLst"))
            CreateNewConditions(oChLABtnList, "Postable", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Choosing  from the CFL and setting the values to the corresponding field.
    ''' </summary>
    ''' <param name="ControlName"></param>
    ''' <param name="ColumnUID"></param>
    ''' <param name="CurrentRow"></param>
    ''' <param name="ChoosefromListUID"></param>
    ''' <param name="ChooseFromListSelectedObjects"></param>
    ''' <remarks></remarks>
    Private Sub SkillGroup_ChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable) Handles Me.ChooseFromList
        Dim oDataTable As SAPbouiCOM.DataTable = ChooseFromListSelectedObjects
        Dim oCurrentRow As Integer
        Dim oWCNo, oWCName, oLAccCode, oLAccName As String
        Dim oRs As SAPbobsCOM.Recordset
        Try
            oCurrentRow = CurrentRow
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '*************Work Centre CFL**************
            If (ControlName = "txtwcode" Or ControlName = "btnwc") And (ChoosefromListUID = "WCLst" Or ChoosefromListUID = "BtWCLst") Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    If Not oDataTable Is Nothing Then
                        oWCNo = oDataTable.GetValue("Code", 0)
                        oWCName = oDataTable.GetValue("U_WCname", 0)
                        oParentDB.SetValue("U_WCcode", oParentDB.Offset, oWCNo)
                        oParentDB.SetValue("U_WCName", oParentDB.Offset, oWCName)
                    End If
                Else
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        If Not oDataTable Is Nothing Then
                            oWCNo = oDataTable.GetValue("Code", 0)
                            oWCName = oDataTable.GetValue("U_WCname", 0)
                            oParentDB.SetValue("U_WCcode", oParentDB.Offset, oWCNo)
                            oParentDB.SetValue("U_WCName", oParentDB.Offset, oWCName)
                        End If
                    End If
                End If
            End If
            '***********Labour  Account CFL********************
            If (ControlName = "txtaccod" Or ControlName = "btnlacct") And (ChoosefromListUID = "LAccLst" Or ChoosefromListUID = "BtLAccLst") Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    If Not oDataTable Is Nothing Then
                        oLAccCode = oDataTable.GetValue("FormatCode", 0)
                        oLAccName = oDataTable.GetValue("AcctName", 0)
                        oParentDB.SetValue("U_Accode", oParentDB.Offset, FormatAccountCode(oLAccCode))
                        oParentDB.SetValue("U_Acname", oParentDB.Offset, oLAccName)
                        oParentDB.SetValue("U_ActAcCode", oParentDB.Offset, oLAccCode.ToString().Replace("-", ""))
                    End If
                Else
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        If Not oDataTable Is Nothing Then
                            oLAccCode = oDataTable.GetValue("FormatCode", 0)
                            oLAccName = oDataTable.GetValue("AcctName", 0)
                            oParentDB.SetValue("U_Accode", oParentDB.Offset, FormatAccountCode(oLAccCode))
                            oParentDB.SetValue("U_Acname", oParentDB.Offset, oLAccName)
                            oParentDB.SetValue("U_ActAcCode", oParentDB.Offset, oLAccCode.ToString().Replace("-", ""))
                        End If
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
    ''' <summary>
    ''' This is used to Load the Running Rate Currency in the Combo
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CurrCombo()
        Dim oRs, oRs1 As SAPbobsCOM.Recordset
        Try
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery("select CurrCode from OCRN where CurrCode is not null")
            oRs1.DoQuery("select a.CurrCode from OCRN a,OADM b where a.CurrCode=b.MainCurncy")
            oRs.MoveFirst()
            If oLbrCurrCombo.ValidValues.Count > 0 Then
                For i As Int16 = oLbrCurrCombo.ValidValues.Count - 1 To 0 Step -1
                    oLbrCurrCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            For i As Int16 = 0 To oRs.RecordCount - 1
                oLbrCurrCombo.ValidValues.Add(oRs.Fields.Item(0).Value, "")
                oRs.MoveNext()
            Next
            oLbrCurrCombo.Select(oRs1.Fields.Item(0).Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        Finally
            oRs = Nothing
            oRs1 = Nothing
            GC.Collect()
        End Try
    End Sub

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Dim ChooseFromListEvent As SAPbouiCOM.ChooseFromListEvent
        Try
            If pVal.FormUID = "FLSG" Then
                '*****************************Releasing the Com Object*******************************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                    SBO_Application = Nothing
                    GC.Collect()
                End If
                '**********ChooseFromList Event is called using the raiseevent*********
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST And pVal.BeforeAction = False Then
                    ChooseFromListEvent = pVal
                    RaiseEvent ChooseFromList(pVal.ItemUID, pVal.ColUID, pVal.Row, ChooseFromListEvent.ChooseFromListUID, ChooseFromListEvent.SelectedObjects)
                End If
                '******** Validation() method is called for validating the values in the edit text **********
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.ItemUID = "1" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = True Then
                        Try
                            Validation()
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End Try
                    End If

                    
                End If

            End If

            If pVal.ItemUID = "lnkwc" And pVal.BeforeAction = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                Dim oWCCode As String
                oWCCode = oWCNoTxt.Value
                WorkCentreClass = New WorkCentre(SBO_Application, oCompany, oWCCode, "SkillGroups")
            End If
            '********** Add Button Press ***********
            If pVal.ItemUID = "1" And (pVal.BeforeAction = False) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                Try
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        oForm.Items.Item("txtgcode").Enabled = False
                        oForm.Items.Item("txtgname").Enabled = True
                    End If
                    If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        oForm.Refresh()
                        oForm.Freeze(True)
                        oActiveCheck.Checked = True
                        CurrCombo()
                        oForm.Items.Item("cmbcurncy").Enabled = False
                        oForm.Freeze(False)
                        oSGCodeTxt.Active = True
                    End If
                Catch ex As Exception
                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
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
                Throw New Exception("Skill Group Code should not be Empty")
            End If
            If oSGDescTxt.Value.Length = 0 Then
                oSGDescTxt.Active = True
                Throw New Exception("Skill Group Name should not be Empty")
            End If
            If oWCNoTxt.Value.Length = 0 Then
                oWCNoTxt.Active = True
                Throw New Exception("Work Centre Time should not be Empty")
            End If
            If oLbrRateTxt.Value.Length = 0 Or oLbrRateTxt.Value = "0.0" Then
                oLbrRateTxt.Active = True
                Throw New Exception("Labour Rate should be greater than zero")
            Else
                If CInt(oLbrRateTxt.Value) < 0 Then
                    oLbrRateTxt.Active = True
                    Throw New Exception("Labour Rate should be greater than zero")
                End If
            End If
            'Modified by Manimaran------s
            'If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            '    oRs.DoQuery("select Code from [@PSSIT_OLGP]  where Code= '" & oSGCodeTxt.Value & "' ")
            '    If oRs.RecordCount > 0 Then
            '        oSGCodeTxt.Active = True
            '        Throw New Exception("Skill Group Code Already Exist")
            '    End If
            '    oRs1.DoQuery("select U_LGname from [@PSSIT_OLGP]  where U_LGname= '" & oSGDescTxt.Value & "' ")
            '    If oRs1.RecordCount > 0 Then
            '        oSGDescTxt.Active = True
            '        Throw New Exception("Skill Group Name Already Exist")
            '    End If
            'End If
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                oRs.DoQuery("select Code from [@PSSIT_OLGP]  where Code= '" & oSGCodeTxt.Value & "' and U_LGname= '" & oSGDescTxt.Value & "'")
                If oRs.RecordCount > 0 Then
                    oSGDescTxt.Active = True
                    Throw New Exception("This Combination Already Exist")
                End If
            End If
            If oForm.Items.Item("txtaccod").Specific.value.length = 0 Then
                Throw New Exception("Account Code should not be empty")
            End If
            'Modified by Manimaran------e
            AccKeyCheck()
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            oRs1 = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' Checking the Account Key based on the production configuration form.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AccKeyCheck()
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            oRs.DoQuery("Select * from [@PSSIT_OCON] where U_AccKey = 'Y'")
            If oRs.RecordCount > 0 Then
                If oLBAcctCodeTxt.Value.Length = 0 Then
                    oLBAcctCodeTxt.Active = True
                    Throw New Exception("Select Account from the List")
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Dim FormID As String
        Try
            FormID = SBO_Application.Forms.ActiveForm.UniqueID
            If pVal.MenuUID = "1281" And FormID = "FLSG" Then
                If pVal.BeforeAction = True Then
                    oForm.Items.Item("txtgcode").Enabled = True
                    oForm.Items.Item("txtgname").Enabled = True
                End If
                If pVal.BeforeAction = False Then
                    oSGCodeTxt.Active = True
                End If
            End If
            If pVal.MenuUID = "1282" And pVal.BeforeAction = False And FormID = "FLSG" Then
                oForm.Freeze(True)
                oParentDB.Offset = oParentDB.Size - 1
                oActiveCheck.Checked = True
                oForm.Items.Item("txtgcode").Enabled = True
                oForm.Items.Item("txtgname").Enabled = True
                oForm.Items.Item("cmbcurncy").Enabled = False
                CurrCombo()
                oForm.Freeze(False)
                oSGCodeTxt.Active = True
            End If
            If pVal.MenuUID = "1283" And FormID = "FLSG" Then
                If pVal.BeforeAction = True Then
                    Dim oStrSql As String
                    Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Try
                        oStrSql = "Select (sum(a.cnt) + Sum (b.cnt) + Sum (c.cnt)) as ReferredCount " _
                                & "from (select count(*) as cnt from [@PSSIT_OLBR]  Where U_LGCode = '" & oSGCodeTxt.Value & "' ) as a, " _
                                & "(Select count(*) as cnt from [@PSSIT_RTE2]  Where U_Skilgrp = '" & oSGCodeTxt.Value & "') as b, " _
                                & "(Select count(*) as cnt from [@PSSIT_PRN2]  Where U_Skilgrp = '" & oSGCodeTxt.Value & "') as c "
                        oRs.DoQuery(oStrSql)
                        If oRs.RecordCount > 0 Then
                            oRs.MoveFirst()
                            If oRs.Fields.Item("ReferredCount").Value > 0 Then
                                SBO_Application.SetStatusBarMessage("Cannot be removed. Transactions are linked to an object, '" & oSGCodeTxt.Value & "'", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                            End If
                        End If
                    Catch ex As Exception
                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        BubbleEvent = False
                    Finally
                        oRs = Nothing
                        GC.Collect()
                    End Try
                End If
            End If
            '*****************************Navigation*******************************
            If (pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") And pVal.BeforeAction = False And FormID = "FLSG" Then
                Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Try
                    oRs.DoQuery("Select * from [@PSSIT_OLGP]")
                    If oRs.RecordCount > 0 Then
                        oForm.Items.Item("txtgcode").Enabled = False
                        oForm.Items.Item("txtgname").Enabled = True
                    Else
                        oForm.Items.Item("txtgcode").Enabled = True
                        oForm.Items.Item("txtgname").Enabled = True
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
    
End Class
