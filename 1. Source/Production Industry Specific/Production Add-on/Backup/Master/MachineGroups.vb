'''' <summary>
'''' Author                     Created Date
'''' Suresh                      17/12/2008
'''' <remarks> This class is used for entering the Machine Group Details.</remarks>
Public Class MachineGroups
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
    Private oChWCList, oChWCBtnList, oChRAccList, oChRABtnList, oChSAccList, oChSABtnList As SAPbouiCOM.ChooseFromList
    Private Event ChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable)
    Private Event BeforeChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable)
    '**************************Items - EditText************************************
    Private oMGCodeTxt, oMGDescTxt, oWCNoTxt, oTWCNameTxt, oRunRateTxt, oSetUpRateTxt, oRAcctCodeTxt, oRAcctNameTxt, oSAcctCodeTxt, oSAcctNameTxt, oInfo1Txt, oInfo2Txt, oRemarksTxt, oRActAcCodeTxt, oSActAcCodeTxt As SAPbouiCOM.EditText
    '**************************Items - Combo************************************
    Private oRCurrCombo, oSCurrCombo As SAPbouiCOM.ComboBox
    '**************************Items - CheckBox************************************
    Private oActiveCheck As SAPbouiCOM.CheckBox
    '**************************Link Button************************************
    Private oWCCodeLink, oRAcctCodeLink, oSAcctCodeLink As SAPbouiCOM.LinkedButton
    '**************************Items - Button************************************
    Private BtnWC, BtnRAcct, BtnSAcct As SAPbouiCOM.Button
    Private sPath As String = IO.Directory.GetParent(Application.ExecutablePath).ToString

    Private WithEvents WorkCentreClass As WorkCentre

    Private oMGCode As String
    Private oFormName As String
#End Region
    ''' <summary>
    ''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmMachineGroups.srf") method is called to load the Machine Group form.
    ''' Drawform() method is called to Initialize the form,Datasources and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company, Optional ByVal aMGCode As String = Nothing, Optional ByVal aFormName As String = Nothing)
        MyBase.New(aSBO_Application, aCompany)
        SBO_Application = aSBO_Application
        oCompany = aCompany
        oMGCode = aMGCode
        oFormName = aFormName
        LoadFromXML("FrmMachineGroups.srf")
        DrawForm()
        If oFormName = "MachineMaster" Then
            oParentDB.SetValue("Code", oParentDB.Offset, oMGCode)
            oMGCodeTxt.Value = oMGCode
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
            oParentDB = oForm.DataSources.DBDataSources.Item("@PSSIT_OMGP")
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
            oMGCodeTxt = oForm.Items.Item("txtgcode").Specific
            oMGCodeTxt.DataBind.SetBound(True, "@PSSIT_OMGP", "Code")

            oMGDescTxt = oForm.Items.Item("txtgname").Specific
            oMGDescTxt.DataBind.SetBound(True, "@PSSIT_OMGP", "U_MGname")

            oWCNoTxt = oForm.Items.Item("txtwcode").Specific
            oWCNoTxt.DataBind.SetBound(True, "@PSSIT_OMGP", "U_WCcode")
            oWCNoTxt.ChooseFromListUID = "WCLst"
            oWCNoTxt.ChooseFromListAlias = "Code"
            oForm.Items.Item("txtwcode").LinkTo = "lnkwccod"
            oForm.Items.Add("lnkwccod", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnkwccod").Visible = True
            oForm.Items.Item("lnkwccod").LinkTo = "txtwcode"
            oForm.Items.Item("lnkwccod").Top = 21
            oForm.Items.Item("lnkwccod").Left = 110
            oForm.Items.Item("lnkwccod").Description = "Link to" & vbNewLine & "Work Centre"
            oWCCodeLink = oForm.Items.Item("lnkwccod").Specific

            oTWCNameTxt = oForm.Items.Item("txtwcnam").Specific
            oTWCNameTxt.DataBind.SetBound(True, "@PSSIT_OMGP", "U_WCName")

            oRunRateTxt = oForm.Items.Item("txtrunrat").Specific
            oRunRateTxt.DataBind.SetBound(True, "@PSSIT_OMGP", "U_Runrate")

            oSetUpRateTxt = oForm.Items.Item("txtdsetrat").Specific
            oSetUpRateTxt.DataBind.SetBound(True, "@PSSIT_OMGP", "U_Setrate")

            oRCurrCombo = oForm.Items.Item("cmbcurncyr").Specific
            oRCurrCombo.DataBind.SetBound(True, "@PSSIT_OMGP", "U_RCurrncy")
            oForm.Items.Item("cmbcurncyr").Enabled = False
            RCurrCombo()
          
            oSCurrCombo = oForm.Items.Item("cmbcurncys").Specific
            oSCurrCombo.DataBind.SetBound(True, "@PSSIT_OMGP", "U_SCurrncy")
            oForm.Items.Item("cmbcurncys").Enabled = False
            SCurrCombo()

            oRAcctCodeTxt = oForm.Items.Item("txtaccodr").Specific
            oRAcctCodeTxt.DataBind.SetBound(True, "@PSSIT_OMGP", "U_RAccode")
            oRAcctCodeTxt.ChooseFromListUID = "RAccLst"
            oRAcctCodeTxt.ChooseFromListAlias = "FormatCode"
            oForm.Items.Item("txtaccodr").LinkTo = "lnkraccod"
            oForm.Items.Add("lnkraccod", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnkraccod").Visible = True
            oForm.Items.Item("lnkraccod").LinkTo = "txtaccodr"
            oForm.Items.Item("lnkraccod").Top = 51
            oForm.Items.Item("lnkraccod").Left = 233
            oForm.Items.Item("lnkraccod").Description = "Link to" & vbNewLine & "Chart Of Accounts"
            oRAcctCodeLink = oForm.Items.Item("lnkraccod").Specific
            oRAcctCodeLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_GLAccounts


            oRAcctNameTxt = oForm.Items.Item("txtacnamr").Specific
            oRAcctNameTxt.DataBind.SetBound(True, "@PSSIT_OMGP", "U_RAcname")

            oSAcctCodeTxt = oForm.Items.Item("txtaccods").Specific
            oSAcctCodeTxt.DataBind.SetBound(True, "@PSSIT_OMGP", "U_SAccode")
            oSAcctCodeTxt.ChooseFromListUID = "SAccLst"
            oSAcctCodeTxt.ChooseFromListAlias = "FormatCode"
            oForm.Items.Item("txtaccods").LinkTo = "lnksaccod"
            oForm.Items.Add("lnksaccod", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnksaccod").Visible = True
            oForm.Items.Item("lnksaccod").LinkTo = "txtaccods"
            oForm.Items.Item("lnksaccod").Top = 66
            oForm.Items.Item("lnksaccod").Left = 233
            oForm.Items.Item("lnksaccod").Description = "Link to" & vbNewLine & "Chart Of Accounts"
            oSAcctCodeLink = oForm.Items.Item("lnksaccod").Specific
            oSAcctCodeLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_GLAccounts

           
            oSAcctNameTxt = oForm.Items.Item("txtacnams").Specific
            oSAcctNameTxt.DataBind.SetBound(True, "@PSSIT_OMGP", "U_SAcname")

            oInfo1Txt = oForm.Items.Item("txtadnl1").Specific
            oForm.Items.Item("lbladnl1").Visible = False
            oForm.Items.Item("txtadnl1").Visible = False
            oInfo1Txt.DataBind.SetBound(True, "@PSSIT_OMGP", "U_Adnl1")

            oInfo2Txt = oForm.Items.Item("txtadnl2").Specific
            oForm.Items.Item("lbladnl2").Visible = False
            oForm.Items.Item("txtadnl2").Visible = False
            oInfo2Txt.DataBind.SetBound(True, "@PSSIT_OMGP", "U_Adnl2")

            oRemarksTxt = oForm.Items.Item("txtremark").Specific
            oRemarksTxt.DataBind.SetBound(True, "@PSSIT_OMGP", "U_Remarks")

            oActiveCheck = oForm.Items.Item("chkactive").Specific
            oActiveCheck.DataBind.SetBound(True, "@PSSIT_OMGP", "U_Active")
            oActiveCheck.Checked = True

            BtnWC = oForm.Items.Item("btnwc").Specific
            oForm.Items.Item("btnwc").Description = "Choose from List" & vbNewLine & "Work Centre List View"
            BtnWC.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            BtnWC.Image = sPath & "\Resources\CFL.bmp"
            BtnWC = oForm.Items.Item("btnwc").Specific
            BtnWC.ChooseFromListUID = "BtWCLst"

            BtnRAcct = oForm.Items.Item("btnracct").Specific
            oForm.Items.Item("btnracct").Description = "Choose from List" & vbNewLine & "Account Detail List View"
            BtnRAcct.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            BtnRAcct.Image = sPath & "\Resources\CFL.bmp"
            BtnRAcct = oForm.Items.Item("btnracct").Specific
            BtnRAcct.ChooseFromListUID = "BtRAccLst"

            BtnSAcct = oForm.Items.Item("btnsacct").Specific
            oForm.Items.Item("btnsacct").Description = "Choose from List" & vbNewLine & "Account Detail List View"
            BtnSAcct.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            BtnSAcct.Image = sPath & "\Resources\CFL.bmp"
            BtnSAcct = oForm.Items.Item("btnsacct").Specific
            BtnSAcct.ChooseFromListUID = "BtSAccLst"

            oRActAcCodeTxt = oForm.Items.Item("txtraccode").Specific
            oForm.Items.Item("txtraccode").Enabled = False
            'oForm.Items.Item("txtraccode").Visible = False
            oRActAcCodeTxt.DataBind.SetBound(True, "@PSSIT_OMGP", "U_RActAcCode")

            oSActAcCodeTxt = oForm.Items.Item("txtsaccode").Specific
            oForm.Items.Item("txtsaccode").Enabled = False
            'oForm.Items.Item("txtsaccode").Visible = False
            oSActAcCodeTxt.DataBind.SetBound(True, "@PSSIT_OMGP", "U_SActAcCode")
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
            '***************************Work Centre-CFL*********************
            oChWCList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_WCR", "WCLst"))
            CreateNewConditions(oChWCList, "U_Active", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            oChWCBtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_WCR", "BtWCLst"))
            CreateNewConditions(oChWCBtnList, "U_Active", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            '***********************************Running Rate-CFL*****************************
            oChRAccList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "1", "RAccLst"))
            CreateNewConditions(oChRAccList, "Postable", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            oChRABtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "1", "BtRAccLst"))
            CreateNewConditions(oChRABtnList, "Postable", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            '*******************************Setup Rate-CFL********************************
            oChSAccList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "1", "SAccLst"))
            CreateNewConditions(oChSAccList, "Postable", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            oChSABtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "1", "BtSAccLst"))
            CreateNewConditions(oChSABtnList, "Postable", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub MachineGroups_BeforeChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable) Handles Me.BeforeChooseFromList
        Try
            If (ControlName = "txtaccodr" Or ControlName = "btnracct") And (ChoosefromListUID = "RAccLst" Or ChoosefromListUID = "BtRAccLst") Then
                If oRAcctCodeTxt.Value.Length > 0 Then
                    oParentDB.SetValue("U_RAccode", oParentDB.Offset, Replace(oRAcctCodeTxt.Value, "-", ""))
                End If
            End If
            If (ControlName = "txtaccods" Or ControlName = "btnsacct") And (ChoosefromListUID = "SAccLst" Or ChoosefromListUID = "BtSAccLst") Then
                If oSAcctCodeTxt.Value.Length > 0 Then
                    oParentDB.SetValue("U_SAccode", oParentDB.Offset, Replace(oSAcctCodeTxt.Value, "-", ""))
                End If
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
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
    Private Sub MachineGroup_ChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable) Handles Me.ChooseFromList
        Dim oDataTable As SAPbouiCOM.DataTable = ChooseFromListSelectedObjects
        Dim oCurrentRow As Integer
        Dim oWCNo, oWCName, oRAccCode, oRAccName, oSAccCode, oSAccName As String
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
            '***********Running Rate Account CFL********************
            If (ControlName = "txtaccodr" Or ControlName = "btnracct") And (ChoosefromListUID = "RAccLst" Or ChoosefromListUID = "BtRAccLst") Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    If Not oDataTable Is Nothing Then
                        oRAccCode = oDataTable.GetValue("FormatCode", 0)
                        oRAccName = oDataTable.GetValue("AcctName", 0)
                        oParentDB.SetValue("U_RAccode", oParentDB.Offset, FormatAccountCode(oRAccCode))
                        oParentDB.SetValue("U_RAcname", oParentDB.Offset, oRAccName)
                        oParentDB.SetValue("U_RActAcCode", oParentDB.Offset, oRAccCode.ToString().Replace("-", ""))
                    End If
                Else
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        If Not oDataTable Is Nothing Then
                            oRAccCode = oDataTable.GetValue("FormatCode", 0)
                            oRAccName = oDataTable.GetValue("AcctName", 0)
                            oParentDB.SetValue("U_RAccode", oParentDB.Offset, FormatAccountCode(oRAccCode))
                            oParentDB.SetValue("U_RAcname", oParentDB.Offset, oRAccName)
                            oParentDB.SetValue("U_RActAcCode", oParentDB.Offset, oRAccCode.ToString().Replace("-", ""))
                        End If
                    End If
                End If
            End If
            '***********SetUp Rate Account CFL********************
            If (ControlName = "txtaccods" Or ControlName = "btnsacct") And (ChoosefromListUID = "SAccLst" Or ChoosefromListUID = "BtSAccLst") Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    If Not oDataTable Is Nothing Then
                        oSAccCode = oDataTable.GetValue("FormatCode", 0)
                        oSAccName = oDataTable.GetValue("AcctName", 0)
                        oParentDB.SetValue("U_SAccode", oParentDB.Offset, FormatAccountCode(oSAccCode))
                        oParentDB.SetValue("U_SAcname", oParentDB.Offset, oSAccName)
                        oParentDB.SetValue("U_SActAcCode", oParentDB.Offset, oSAccCode.ToString().Replace("-", ""))
                    End If
                Else
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        If Not oDataTable Is Nothing Then
                            oSAccCode = oDataTable.GetValue("FormatCode", 0)
                            oSAccName = oDataTable.GetValue("AcctName", 0)
                            oParentDB.SetValue("U_SAccode", oParentDB.Offset, FormatAccountCode(oSAccCode))
                            oParentDB.SetValue("U_SAcname", oParentDB.Offset, oSAccName)
                            oParentDB.SetValue("U_SActAcCode", oParentDB.Offset, oSAccCode.ToString().Replace("-", ""))
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
    Private Sub RCurrCombo()
        Dim oRs, oRs1 As SAPbobsCOM.Recordset
        Try
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery("select CurrCode from OCRN where CurrCode is not null")
            oRs1.DoQuery("select a.CurrCode from OCRN a,OADM b where a.CurrCode=b.MainCurncy")
            oRs.MoveFirst()
            If oRCurrCombo.ValidValues.Count > 0 Then
                For i As Int16 = oRCurrCombo.ValidValues.Count - 1 To 0 Step -1
                    oRCurrCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            For i As Int16 = 0 To oRs.RecordCount - 1
                oRCurrCombo.ValidValues.Add(oRs.Fields.Item(0).Value, "")
                oRs.MoveNext()
            Next
            oRCurrCombo.Select(oRs1.Fields.Item(0).Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            oRs1 = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' This is used to Load the SetUp Rate Currency in the Combo
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SCurrCombo()
        Dim oRs, oRs1 As SAPbobsCOM.Recordset
        Try
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery("select CurrCode from OCRN where CurrCode is not null")
            oRs1.DoQuery("select a.CurrCode from OCRN a,OADM b where a.CurrCode=b.MainCurncy")
            oRs.MoveFirst()
            If oSCurrCombo.ValidValues.Count > 0 Then
                For i As Int16 = oSCurrCombo.ValidValues.Count - 1 To 0 Step -1
                    oSCurrCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            For i As Int16 = 0 To oRs.RecordCount - 1
                oSCurrCombo.ValidValues.Add(oRs.Fields.Item(0).Value, "")
                oRs.MoveNext()
            Next
            oSCurrCombo.Select(oRs1.Fields.Item(0).Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            oRs1 = Nothing
            GC.Collect()
        End Try
    End Sub


    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Dim ChooseFromListEvent As SAPbouiCOM.ChooseFromListEvent
        Try
            If pVal.FormUID = "FMG" Then
                '*****************************Releasing the Com Object*******************************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                    SBO_Application = Nothing
                    GC.Collect()
                End If
                '**********ChooseFromList Event is called using the raiseevent*********
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                    ChooseFromListEvent = pVal
                    If pVal.BeforeAction = True Then
                        RaiseEvent BeforeChooseFromList(pVal.ItemUID, pVal.ColUID, pVal.Row, ChooseFromListEvent.ChooseFromListUID, ChooseFromListEvent.SelectedObjects)
                    Else
                        RaiseEvent ChooseFromList(pVal.ItemUID, pVal.ColUID, pVal.Row, ChooseFromListEvent.ChooseFromListUID, ChooseFromListEvent.SelectedObjects)
                    End If
                End If
                
                'If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST And pVal.BeforeAction = False Then
                '    ChooseFromListEvent = pVal
                '    RaiseEvent ChooseFromList(pVal.ItemUID, pVal.ColUID, pVal.Row, ChooseFromListEvent.ChooseFromListUID, ChooseFromListEvent.SelectedObjects)
                'End If
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

                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And pVal.BeforeAction = True Then
                        Try
                            validation1()
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End Try
                    End If



                    If pVal.ItemUID = "lnkwccod" And pVal.BeforeAction = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        Dim oWCCode As String
                        oWCCode = oWCNoTxt.Value
                        WorkCentreClass = New WorkCentre(SBO_Application, oCompany, oWCCode, "MachineGroups")
                    End If
                End If

                '********** Add Button Press ***********
                If pVal.ItemUID = "1" And (pVal.BeforeAction = False) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    Try
                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            oForm.Items.Item("txtgcode").Enabled = False
                            oForm.Items.Item("txtgname").Enabled = True
                        End If
                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            oForm.Refresh()
                            oForm.Freeze(True)
                            oActiveCheck.Checked = True
                            RCurrCombo()
                            SCurrCombo()
                            oForm.Items.Item("cmbcurncys").Enabled = False
                            oForm.Items.Item("cmbcurncyr").Enabled = False
                            oForm.Freeze(False)
                            oMGCodeTxt.Active = True
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
            If pVal.MenuUID = "1281" And FormID = "FMG" Then
                If pVal.BeforeAction = True Then
                    oForm.Items.Item("txtgcode").Enabled = True
                    oForm.Items.Item("txtgname").Enabled = True
                End If
                If pVal.BeforeAction = False Then
                    oMGCodeTxt.Active = True
                End If
            End If
            If pVal.MenuUID = "1282" And pVal.BeforeAction = False And FormID = "FMG" Then
                oForm.Freeze(True)
                oParentDB.Offset = oParentDB.Size - 1
                oActiveCheck.Checked = True
                oForm.Items.Item("txtgcode").Enabled = True
                oForm.Items.Item("txtgname").Enabled = True
                oForm.Items.Item("cmbcurncys").Enabled = False
                oForm.Items.Item("cmbcurncyr").Enabled = False
                RCurrCombo()
                SCurrCombo()
                oForm.Freeze(False)
                oMGCodeTxt.Active = True
            End If
            If pVal.MenuUID = "1283" And FormID = "FMG" Then
                If pVal.BeforeAction = True Then
                    Dim oStrSql As String
                    Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Try
                        oStrSql = "Select IsNull(Count(*),0) as ReferredCount from [@PSSIT_PMWCHDR] where U_MGCode = '" & oMGCodeTxt.Value & "'"
                        oRs.DoQuery(oStrSql)
                        If oRs.RecordCount > 0 Then
                            oRs.MoveFirst()
                            If oRs.Fields.Item("ReferredCount").Value > 0 Then
                                SBO_Application.SetStatusBarMessage("Cannot be removed. Transactions are linked to an object, '" & oMGCodeTxt.Value & "'", SAPbouiCOM.BoMessageTime.bmt_Short, True)
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
            If (pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") And pVal.BeforeAction = False And FormID = "FMG" Then
                Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Try
                    oRs.DoQuery("Select * from [@PSSIT_OMGP]")
                    If oRs.RecordCount > 0 Then
                        oForm.Items.Item("txtgcode").Enabled = False
                        oForm.Items.Item("txtgname").Enabled = True
                    Else
                        oForm.Items.Item("txtgcode").Enabled = True
                        oForm.Items.Item("txtgname").Enabled = True
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
            If oMGCodeTxt.Value.Length = 0 Then
                oMGCodeTxt.Active = True
                Throw New Exception("Machine Group Code should not be Empty")
            End If
            If oMGDescTxt.Value.Length = 0 Then
                oMGDescTxt.Active = True
                Throw New Exception("Machine Group Name should not be Empty")
            End If
            If oWCNoTxt.Value.Length = 0 Then
                oWCNoTxt.Active = True
                Throw New Exception("Work Centre Time should not be Empty")
            End If
            If CInt(oRunRateTxt.Value) < 0 Then
                oRunRateTxt.Active = True
                Throw New Exception("Running Rate should be greater than zero")
            End If
            If CInt(oSetUpRateTxt.Value) < 0 Then
                oSetUpRateTxt.Active = True
                Throw New Exception("SetUp Rate should be greater than zero")
            End If

            oRs.DoQuery("select Code from [@PSSIT_OMGP]  where Code= '" & oMGCodeTxt.Value & "' ")
            If oRs.RecordCount > 0 Then
                Throw New Exception("Machine Group Code Already Exist")
            End If

            oRs1.DoQuery("select U_MGname from [@PSSIT_OMGP]  where U_MGname= '" & oMGDescTxt.Value & "' ")
            If oRs1.RecordCount > 0 Then
                oMGCodeTxt.Active = True
                Throw New Exception("Machine Group Name Already Exist")
            End If
            'Added by Manimaran------s
            If oForm.Items.Item("txtaccodr").Specific.value.length = 0 Then
                Throw New Exception("Account Code should not be empty")
            End If
            If oForm.Items.Item("txtaccods").Specific.value.length = 0 Then
                Throw New Exception("Account Code should not be empty")
            End If
            'Added by Manimaran------e
            AccKeyCheck()
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            oRs1 = Nothing
            GC.Collect()
        End Try
    End Sub
    Private Sub validation1()
        Dim oRs, oRs1 As SAPbobsCOM.Recordset
        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            If oMGCodeTxt.Value.Length = 0 Then
                oMGCodeTxt.Active = True
                Throw New Exception("Machine Group Code should not be Empty")
            End If
            If oMGDescTxt.Value.Length = 0 Then
                oMGDescTxt.Active = True
                Throw New Exception("Machine Group Name should not be Empty")
            End If
            If oWCNoTxt.Value.Length = 0 Then
                oWCNoTxt.Active = True
                Throw New Exception("Work Centre Time should not be Empty")
            End If
            If CInt(oRunRateTxt.Value) < 0 Then
                oRunRateTxt.Active = True
                Throw New Exception("Running Rate should be greater than zero")
            End If
            If CInt(oSetUpRateTxt.Value) < 0 Then
                oSetUpRateTxt.Active = True
                Throw New Exception("SetUp Rate should be greater than zero")
            End If
            'Added by Manimaran------s
            If oForm.Items.Item("txtaccodr").Specific.value.length = 0 Then
                Throw New Exception("Account Code should not be empty")
            End If
            If oForm.Items.Item("txtaccods").Specific.value.length = 0 Then
                Throw New Exception("Account Code should not be empty")
            End If
            'Added by Manimaran------e
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
                If oRAcctCodeTxt.Value.Length = 0 Then
                    oRAcctCodeTxt.Active = True
                    Throw New Exception("Running Rate Account Code should not be Empty")
                End If
                If oSAcctCodeTxt.Value.Length = 0 Then
                    oSAcctCodeTxt.Active = True
                    Throw New Exception("SetUp Rate Account Code should not be Empty")
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
End Class
