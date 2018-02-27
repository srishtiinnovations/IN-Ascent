'''' <summary>
'''' Author                     Created Date
'''' Suresh                      18/12/2008
'''' <remarks> This class is used for entering the Tools Details.</remarks>
Public Class Tools
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
    Private oParentDB, oReCondDB As SAPbouiCOM.DBDataSource
    '**************************Form************************************
    Private oForm As SAPbouiCOM.Form
    '**************************ChooseFromList************************************
    Private oChItmList, oChItmBtnList, oChWCList, oChWCBtnList, oChAccList, oChABtnList As SAPbouiCOM.ChooseFromList
    Private Event ChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable)
    '**************************Items - EditText************************************
    Private oTLCodeTxt, oTLDescTxt, oItmCodeTxt, oItmNameTxt, oWCNoTxt, oWCNameTxt, oDtPurTxt, oLdCostTxt, oExpStrkTxt, oCmpltStrkTxt, oSTimeTxt, oCstStrkTxt, oAcctCodeTxt, oAcctNameTxt, oTechSpecTxt, oInfo1Txt, oInfo2Txt, oRemarksTxt, oActAcCodeTxt, oNot As SAPbouiCOM.EditText
    '**************************Items - Combo************************************
    Private oPrtToolCombo As SAPbouiCOM.ComboBox
    '**************************Items - CheckBox************************************
    Private oReCondCheck, oActiveCheck As SAPbouiCOM.CheckBox
    '**************************Items - LinkedButton************************************
    Private oWCCodeLink, oItmCodeLink, oAcctCodeLink As SAPbouiCOM.LinkedButton
    '**************************Items - Button************************************
    Private BtnItm, BtnWC, BtnAcct As SAPbouiCOM.Button
    Private sPath As String = IO.Directory.GetParent(Application.ExecutablePath).ToString
    Private oToolCode, oFormName As String
    Private WithEvents WorkCentreClass As WorkCentre
#End Region
    ''' <summary>
    ''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmTools.srf") method is called to load the Labour form.
    ''' Drawform() method is called to Initialize the form,Datasources and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company, Optional ByVal aToolCode As String = Nothing, Optional ByVal aFormName As String = Nothing)
        MyBase.New(aSBO_Application, aCompany)
        SBO_Application = aSBO_Application
        oCompany = aCompany
        oFormName = aFormName
        oToolCode = aToolCode
        LoadFromXML("FrmTools.srf")
        DrawForm()
        If oFormName = "Production Entry" Or oFormName = "OprRouting" Or oFormName = "Operation" Then
            oParentDB.SetValue("Code", oParentDB.Offset, oToolCode)
            oTLCodeTxt.Value = oToolCode
            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        End If
        oForm.DataBrowser.BrowseBy = "txttcode"
    End Sub
    ''' <summary>
    ''' Connecting the application through connection string.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetApplication()
        'Dim sConnectionString As String
        'SboGuiApi = New SAPbouiCOM.SboGuiApi
        'sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"
        'SboGuiApi.Connect(sConnectionString)
        'SboGuiApi.AddonIdentifier = "5645523035446576656C6F706D656E743A453038373933323333343581F0D8D8C45495472FC628EF425AD5AC2AEDC411"
        'SBO_Application = SboGuiApi.GetApplication()
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
            oParentDB = oForm.DataSources.DBDataSources.Item("@PSSIT_OTLS")
            oReCondDB = oForm.DataSources.DBDataSources.Item("@PSSIT_OTLS")
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
            oTLCodeTxt = oForm.Items.Item("txttcode").Specific
            oTLCodeTxt.DataBind.SetBound(True, "@PSSIT_OTLS", "Code")

            oTLDescTxt = oForm.Items.Item("txttname").Specific
            oTLDescTxt.DataBind.SetBound(True, "@PSSIT_OTLS", "U_TLname")

            oItmCodeTxt = oForm.Items.Item("txtitmcod").Specific
            oItmCodeTxt.DataBind.SetBound(True, "@PSSIT_OTLS", "U_Itemcode")
            oItmCodeTxt.ChooseFromListUID = "ItmLst"
            oItmCodeTxt.ChooseFromListAlias = "ItemCode"
            oForm.Items.Item("txtitmcod").LinkTo = "lnkitem"
            oForm.Items.Add("lnkitem", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnkitem").Visible = True
            oForm.Items.Item("lnkitem").LinkTo = "txtitmcod"
            oForm.Items.Item("lnkitem").Top = 21
            oForm.Items.Item("lnkitem").Left = 92
            oForm.Items.Item("lnkitem").Description = "Link to" & vbNewLine & "Item Master"
            oItmCodeLink = oForm.Items.Item("lnkitem").Specific
            oItmCodeLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Items


            oItmNameTxt = oForm.Items.Item("txtitmnam").Specific
            oItmNameTxt.DataBind.SetBound(True, "@PSSIT_OTLS", "U_Itemname")

            oWCNoTxt = oForm.Items.Item("txtwcode").Specific
            oWCNoTxt.DataBind.SetBound(True, "@PSSIT_OTLS", "U_WCcode")
            oWCNoTxt.ChooseFromListUID = "WCLst"
            oWCNoTxt.ChooseFromListAlias = "Code"
            oForm.Items.Item("txtwcode").LinkTo = "lnkwccod"
            oForm.Items.Add("lnkwccod", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnkwccod").Visible = True
            oForm.Items.Item("lnkwccod").LinkTo = "txtwcode"
            oForm.Items.Item("lnkwccod").Top = 36
            oForm.Items.Item("lnkwccod").Left = 92
            oForm.Items.Item("lnkwccod").Description = "Link to" & vbNewLine & "Work Centre"
            oWCCodeLink = oForm.Items.Item("lnkwccod").Specific


            oWCNameTxt = oForm.Items.Item("txtwcnam").Specific
            oWCNameTxt.DataBind.SetBound(True, "@PSSIT_OTLS", "U_WCname")

            oDtPurTxt = oForm.Items.Item("txtpdate").Specific
            oDtPurTxt.DataBind.SetBound(True, "@PSSIT_OTLS", "U_Purdate")

            oLdCostTxt = oForm.Items.Item("txtlct").Specific
            oLdCostTxt.DataBind.SetBound(True, "@PSSIT_OTLS", "U_Lcost")

            oExpStrkTxt = oForm.Items.Item("txtenou").Specific
            oExpStrkTxt.DataBind.SetBound(True, "@PSSIT_OTLS", "U_Enou")

            oCmpltStrkTxt = oForm.Items.Item("txtcnou").Specific
            oCmpltStrkTxt.DataBind.SetBound(True, "@PSSIT_OTLS", "U_Cnou")

            oSTimeTxt = oForm.Items.Item("txttst").Specific
            oSTimeTxt.DataBind.SetBound(True, "@PSSIT_OTLS", "U_Tstime")

            oCstStrkTxt = oForm.Items.Item("txtcpn").Specific
            oCstStrkTxt.DataBind.SetBound(True, "@PSSIT_OTLS", "U_Cpno")

            oAcctCodeTxt = oForm.Items.Item("txtaccode").Specific
            oAcctCodeTxt.DataBind.SetBound(True, "@PSSIT_OTLS", "U_Accode")
            oAcctCodeTxt.ChooseFromListUID = "AccLst"
            oAcctCodeTxt.ChooseFromListAlias = "AcctCode"
            oForm.Items.Item("txtaccode").LinkTo = "lnkaccod"
            oForm.Items.Add("lnkaccod", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnkaccod").Visible = True
            oForm.Items.Item("lnkaccod").LinkTo = "txtaccode"
            oForm.Items.Item("lnkaccod").Top = 96
            oForm.Items.Item("lnkaccod").Left = 92
            oForm.Items.Item("lnkaccod").Description = "Link to" & vbNewLine & "Chart Of Accounts"
            oAcctCodeLink = oForm.Items.Item("lnkaccod").Specific
            oAcctCodeLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_GLAccounts

            oAcctNameTxt = oForm.Items.Item("txtacname").Specific
            oAcctNameTxt.DataBind.SetBound(True, "@PSSIT_OTLS", "U_Acname")

            oPrtToolCombo = oForm.Items.Item("cmbpartool").Specific
            oPrtToolCombo.DataBind.SetBound(True, "@PSSIT_OTLS", "U_Partool")
            oForm.Items.Item("cmbpartool").Enabled = False

            oTechSpecTxt = oForm.Items.Item("txttsp").Specific
            oTechSpecTxt.DataBind.SetBound(True, "@PSSIT_OTLS", "U_Techspec")

            oInfo1Txt = oForm.Items.Item("txtadnl1").Specific
            oForm.Items.Item("lbladnl1").Visible = False
            oForm.Items.Item("txtadnl1").Visible = False
            oInfo1Txt.DataBind.SetBound(True, "@PSSIT_OTLS", "U_Adnl1")

            oInfo2Txt = oForm.Items.Item("txtadnl2").Specific
            oForm.Items.Item("lbladnl2").Visible = False
            oForm.Items.Item("txtadnl2").Visible = False
            oInfo2Txt.DataBind.SetBound(True, "@PSSIT_OTLS", "U_Adnl2")

            oRemarksTxt = oForm.Items.Item("txtremark").Specific
            oRemarksTxt.DataBind.SetBound(True, "@PSSIT_OTLS", "U_Remarks")

            oActiveCheck = oForm.Items.Item("chkactive").Specific
            oActiveCheck.DataBind.SetBound(True, "@PSSIT_OTLS", "U_Active")
            oActiveCheck.Checked = True

            oReCondCheck = oForm.Items.Item("chkrecon").Specific
            oReCondCheck.DataBind.SetBound(True, "@PSSIT_OTLS", "U_Recond")

            BtnItm = oForm.Items.Item("btnitm").Specific
            oForm.Items.Item("btnitm").Description = "Choose from List" & vbNewLine & "Item List View"
            BtnItm.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            BtnItm.Image = sPath & "\Resources\CFL.bmp"
            BtnItm = oForm.Items.Item("btnitm").Specific
            BtnItm.ChooseFromListUID = "BtItmLst"

            BtnWC = oForm.Items.Item("btnwc").Specific
            oForm.Items.Item("btnwc").Description = "Choose from List" & vbNewLine & "Work Centre List View"
            BtnWC.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            BtnWC.Image = sPath & "\Resources\CFL.bmp"
            BtnWC = oForm.Items.Item("btnwc").Specific
            BtnWC.ChooseFromListUID = "BtWCLst"

            BtnAcct = oForm.Items.Item("btnacct").Specific
            oForm.Items.Item("btnacct").Description = "Choose from List" & vbNewLine & "Account Detail List View"
            BtnAcct.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            BtnAcct.Image = sPath & "\Resources\CFL.bmp"
            BtnAcct = oForm.Items.Item("btnacct").Specific
            BtnAcct.ChooseFromListUID = "BtAccLst"

            oActAcCodeTxt = oForm.Items.Item("txtacaccod").Specific
            oForm.Items.Item("txtacaccod").Enabled = False
            'oForm.Items.Item("txtacaccod").Visible = False
            'oForm.Items.Item("lblacaccod").Visible = False
            oActAcCodeTxt.DataBind.SetBound(True, "@PSSIT_OTLS", "U_ActAcCode")

            'Added by Manimaran------s
            oNot = oForm.Items.Item("50").Specific
            oForm.Items.Item("50").Enabled = True
            oNot.DataBind.SetBound(True, "@PSSIT_OTLS", "U_TypOfItm")
            'Added by Manimaran------e

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
            '*****************************Item-CFL****************************
            oChItmList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "4", "ItmLst"))
            oChItmBtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "4", "BtItmLst"))
            '*********************************Work Centre-CFL*********************************
            oChWCList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_WCR", "WCLst"))
            CreateNewConditions(oChWCList, "U_Active", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            oChWCBtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_WCR", "BtWCLst"))
            CreateNewConditions(oChWCBtnList, "U_Active", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            '**************************Accounts-CFL***************************
            oChAccList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "1", "AccLst"))
            CreateNewConditions(oChAccList, "Postable", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            oChABtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "1", "BtAccLst"))
            CreateNewConditions(oChABtnList, "Postable", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
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
    Private Sub Labour_ChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable) Handles Me.ChooseFromList
        Dim oDataTable As SAPbouiCOM.DataTable = ChooseFromListSelectedObjects
        Dim oCurrentRow As Integer
        Dim oItmId, oItmName, oWCNo, oWCName, oAccCode, oAccName As String
        Dim oRs As SAPbobsCOM.Recordset
        Try
            oCurrentRow = CurrentRow
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '*************Employee CFL**************
            If (ControlName = "txtitmcod" Or ControlName = "btnitm") And (ChoosefromListUID = "ItmLst" Or ChoosefromListUID = "BtItmLst") Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    If Not oDataTable Is Nothing Then
                        oItmId = oDataTable.GetValue("ItemCode", 0)
                        oItmName = oDataTable.GetValue("ItemName", 0)
                        oParentDB.SetValue("U_Itemcode", oParentDB.Offset, oItmId)
                        oParentDB.SetValue("U_Itemname", oParentDB.Offset, oItmName)
                    End If
                Else
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        If Not oDataTable Is Nothing Then
                            oItmId = oDataTable.GetValue("ItemCode", 0)
                            oItmName = oDataTable.GetValue("ItemName", 0)
                            oParentDB.SetValue("U_Itemcode", oParentDB.Offset, oItmId)
                            oParentDB.SetValue("U_Itemname", oParentDB.Offset, oItmName)
                        End If
                    End If
                End If
            End If
            '*************Skill Group CFL**************
            If (ControlName = "txtwcode" Or ControlName = "btnwc") And (ChoosefromListUID = "WCLst" Or ChoosefromListUID = "BtWCLst") Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    If Not oDataTable Is Nothing Then
                        oWCNo = oDataTable.GetValue("Code", 0)
                        oWCName = oDataTable.GetValue("U_WCname", 0)
                        oParentDB.SetValue("U_WCcode", oParentDB.Offset, oWCNo)
                        oParentDB.SetValue("U_WCname", oParentDB.Offset, oWCName)
                    End If
                Else
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        If Not oDataTable Is Nothing Then
                            oWCNo = oDataTable.GetValue("Code", 0)
                            oWCName = oDataTable.GetValue("U_WCname", 0)
                            oParentDB.SetValue("U_WCcode", oParentDB.Offset, oWCNo)
                            oParentDB.SetValue("U_WCname", oParentDB.Offset, oWCName)
                        End If
                    End If
                End If
            End If
            '***********Labour  Account CFL********************
            If (ControlName = "txtaccode" Or ControlName = "btnacct") And (ChoosefromListUID = "AccLst" Or ChoosefromListUID = "BtAccLst") Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    If Not oDataTable Is Nothing Then
                        oAccCode = oDataTable.GetValue("FormatCode", 0)
                        oAccName = oDataTable.GetValue("AcctName", 0)
                        oParentDB.SetValue("U_Accode", oParentDB.Offset, FormatAccountCode(oAccCode))
                        oParentDB.SetValue("U_Acname", oParentDB.Offset, oAccName)
                        oParentDB.SetValue("U_ActAcCode", oParentDB.Offset, oAccCode.ToString().Replace("-", ""))
                    End If
                Else
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        If Not oDataTable Is Nothing Then
                            oAccCode = oDataTable.GetValue("FormatCode", 0)
                            oAccName = oDataTable.GetValue("AcctName", 0)
                            oParentDB.SetValue("U_Accode", oParentDB.Offset, FormatAccountCode(oAccCode))
                            oParentDB.SetValue("U_Acname", oParentDB.Offset, oAccName)
                            oParentDB.SetValue("U_ActAcCode", oParentDB.Offset, oAccCode.ToString().Replace("-", ""))
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
    ''' This is used to Load the Parent Tools in the Combo based on the Condition
    ''' </summary>
    ''' <param name="oCombo"></param>
    ''' <remarks></remarks>
    Private Sub ReCondCombo(ByVal oCombo As SAPbouiCOM.ComboBox)
        If oReCondCheck.Checked = True Then
            Dim rs As SAPbobsCOM.Recordset
            Try
                rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oForm.Items.Item("cmbpartool").Enabled = True
                rs.DoQuery("select Code,U_TLname from [@PSSIT_OTLS] Group by Code,U_TLname")
                rs.MoveFirst()
                If oCombo.ValidValues.Count > 0 Then
                    For i As Int16 = oCombo.ValidValues.Count - 1 To 0 Step -1
                        oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                    Next
                End If
                For i As Int16 = 0 To rs.RecordCount - 1
                    oCombo.ValidValues.Add(rs.Fields.Item(0).Value, rs.Fields.Item(1).Value)
                    rs.MoveNext()
                Next
            Catch ex As Exception
                Throw ex
            Finally
                rs = Nothing
                GC.Collect()
            End Try
        End If
    End Sub

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Dim ChooseFromListEvent As SAPbouiCOM.ChooseFromListEvent
        Try
            If pVal.FormUID = "FTM" Then
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
                    If pVal.ItemUID = "1" Then
                        If (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = True Then
                            Try
                                Validation()
                                UpdateParentTool()
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                            End Try
                        End If
                        '********** Add Button Press ***********
                        If pVal.BeforeAction = False Then
                            Try
                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    SetItemEnabled()
                                End If
                                If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    oForm.Refresh()
                                    oForm.Freeze(True)
                                    oTLCodeTxt.Active = True
                                    SetItemEnabled()
                                    oActiveCheck.Checked = True
                                    oReCondCheck.Checked = False
                                    oExpStrkTxt.Value = 0
                                    oForm.Freeze(False)
                                End If
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End Try
                        End If
                    End If
                    '**********************WorkCentre Link Button**********************
                    If pVal.ItemUID = "lnkwccod" And pVal.BeforeAction = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        Dim oWCCode As String
                        oWCCode = oWCNoTxt.Value
                        WorkCentreClass = New WorkCentre(SBO_Application, oCompany, oWCCode, "Tools")
                    End If
                    '****** If ReConditioned is Checked then Corresponding Data will be loaded in the Group Under ****** 
                    If (pVal.ItemUID = "chkrecon") And (pVal.BeforeAction = False) Then
                        Dim oRs As SAPbobsCOM.Recordset
                        Try
                            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oForm.Items.Item("cmbpartool").Enabled = True
                            oRs.DoQuery("select Code,U_TLname from [@PSSIT_OTLS] Group by Code,U_TLname")
                            oRs.MoveFirst()
                            If oPrtToolCombo.ValidValues.Count > 0 Then
                                For i As Int16 = oPrtToolCombo.ValidValues.Count - 1 To 0 Step -1
                                    oPrtToolCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                                Next
                            End If
                            For i As Int16 = 0 To oRs.RecordCount - 1
                                oPrtToolCombo.ValidValues.Add(oRs.Fields.Item(0).Value, oRs.Fields.Item(1).Value)
                                oRs.MoveNext()
                            Next
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        Finally
                            ors = Nothing
                            GC.Collect()
                        End Try

                        '****** If Group is UnChecked **********
                        If oReCondCheck.Checked = False Then
                            For i As Int16 = oPrtToolCombo.ValidValues.Count - 1 To 0 Step -1
                                oPrtToolCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                            Next
                            oParentDB.SetValue("U_Partool", oParentDB.Offset, "")
                            oParentDB.SetValue("U_Recond", oParentDB.Offset, "N")
                            oPrtToolCombo.Active = False
                            oForm.Items.Item("cmbpartool").Enabled = False
                        End If
                    End If
                End If
                '******** Cost/Stroke Calculation function **********
                If pVal.CharPressed = Keys.Tab And pVal.BeforeAction = False Then
                    '******************Landed Cost******************
                    If pVal.ItemUID = "txtlct" Then
                        Try
                            If (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                If oLdCostTxt.Value = "" Then
                                    oLdCostTxt.Value = "0.00"
                                ElseIf oExpStrkTxt.Value = "" Then
                                    oExpStrkTxt.Value = 0
                                Else
                                    oCstStrkTxt.Value = CDbl(oLdCostTxt.Value) / CDbl(oExpStrkTxt.Value)
                                End If

                            End If
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End Try
                    End If
                    '****************Expected Stroke******************** 
                    If pVal.ItemUID = "txtenou" Then
                        Try
                            If (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                If oLdCostTxt.Value = "" Then
                                    oLdCostTxt.Value = "0.00"
                                ElseIf oExpStrkTxt.Value = "" Then
                                    oExpStrkTxt.Value = 0
                                Else
                                    oCstStrkTxt.Value = CDbl(oLdCostTxt.Value) / CDbl(oExpStrkTxt.Value)
                                End If

                            End If
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End Try
                    End If
                End If
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    ''' <summary>
    ''' This method is used for validating the values in the EditText.
    ''' </summary>
    ''' <remarks></remarks>
#Region "Modified by senthil"
    Private Function getdatetime(ByVal dateString As String) As Date
        Dim objBridge As SAPbobsCOM.SBObob
        objBridge = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Return objBridge.Format_StringToDate(dateString).Fields.Item(0).Value

    End Function
#End Region

    Private Sub Validation()
        Dim oRs, oRs1 As SAPbobsCOM.Recordset
        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try

            If oTLCodeTxt.Value.Length = 0 Then
                oTLCodeTxt.Active = True
                Throw New Exception("Tools Code should not be Empty")
            End If
            If oTLDescTxt.Value.Length = 0 Then
                oTLDescTxt.Active = True
                Throw New Exception("Tools Name should not be Empty")
            End If
            If oWCNoTxt.Value.Length = 0 Then
                oWCNoTxt.Active = True
                Throw New Exception("Work Centre should not be Empty")
            End If
            If oLdCostTxt.Value.Length > 0 Then
                If oLdCostTxt.Value < 0 Then
                    oLdCostTxt.Active = True
                    oLdCostTxt.Value = ""
                    Throw New Exception("Landed cost can't be negative")
                End If
            End If
            If oExpStrkTxt.Value.Length > 0 Then
                If oExpStrkTxt.Value < 0 Then
                    oExpStrkTxt.Active = True
                    oExpStrkTxt.Value = ""
                    Throw New Exception("Expected Sttroke Can't be negative value")
                End If
            End If
            If oSTimeTxt.Value.Length > 0 Then
                If oSTimeTxt.Value < 0 Then
                    oSTimeTxt.Active = True
                    oSTimeTxt.Value = ""
                    Throw New Exception("Setting time can't be negative value")
                End If
            End If
            'Modified by Manimaran-----s
            If oDtPurTxt.String <> "" Then
                Dim dtDate, dtServerdate As Date
                dtDate = getdatetime(oDtPurTxt.String)
                dtServerdate = getdatetime(SBO_Application.Company.ServerDate.ToString)
                '   If DateDiff(DateInterval.Day, CDate(oDtPurTxt.String), CDate(SBO_Application.Company.ServerDate)) < 0 Then
                If DateDiff(DateInterval.Day, dtDate, dtServerdate) < 0 Then
                    ' oDtPurTxt.Value = ""
                    Throw New Exception("PO date should not be greater than the current date")
                End If
            Else
                Throw New Exception("PO date should not be empty")
            End If
            'Modified by Manimaran-----e

            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                oRs.DoQuery("select Code from [@PSSIT_OTLS]  where Code= '" & oTLCodeTxt.Value & "' ")
                If oRs.RecordCount > 0 Then
                    oTLCodeTxt.Active = True
                    Throw New Exception("Tools Code Already Exist")
                End If
                oRs1.DoQuery("select U_TLname from [@PSSIT_OTLS]  where U_TLname= '" & oTLDescTxt.Value & "' ")
                If oRs1.RecordCount > 0 Then
                    oTLDescTxt.Active = True
                    Throw New Exception("Tools Name Already Exist")
                End If

                oRs1.DoQuery("select U_ItemCode from [@PSSIT_OTLS]  where U_ItemCode= '" & oItmCodeTxt.Value & "' ")
                If oRs1.RecordCount > 0 Then
                    oItmCodeTxt.Active = True
                    Throw New Exception("Item Code Already Exist")
                End If

            End If
            'Added by Manimaran------s
            If oForm.Items.Item("txtaccode").Specific.value.length = 0 Then
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
                If oAcctCodeTxt.Value.Length = 0 Then
                    oAcctCodeTxt.Active = True
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
            If pVal.MenuUID = "1281" And FormID = "FTM" Then
                If pVal.BeforeAction = False Then
                    SetItemEnabled()
                    oTLCodeTxt.Active = True
                End If
            End If
            If pVal.MenuUID = "1282" And pVal.BeforeAction = False And FormID = "FTM" Then
                oForm.Freeze(True)
                oParentDB.Offset = oParentDB.Size - 1
                oActiveCheck.Checked = True
                oParentDB.SetValue("U_Recond", oParentDB.Offset, "N")
                'oReCondCheck.Checked = False
                oExpStrkTxt.Value = 0
                SetItemEnabled()
                oForm.Freeze(False)
                oTLCodeTxt.Active = True
            End If
            If pVal.MenuUID = "1283" And FormID = "FTM" Then
                If pVal.BeforeAction = True Then
                    Dim oStrSql As String
                    Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Try
                        oStrSql = "Select (sum(a.cnt) + Sum (b.cnt)) as ReferredCount " _
                                & "from (Select count(*) as cnt from [@PSSIT_RTE3]  Where U_Toolcode = '" & oTLCodeTxt.Value & "') as a, " _
                                & "(Select count(*) as cnt from [@PSSIT_PRN3]  Where U_Toolcode = '" & oTLCodeTxt.Value & "') as b "
                        oRs.DoQuery(oStrSql)
                        If oRs.RecordCount > 0 Then
                            oRs.MoveFirst()
                            If oRs.Fields.Item("ReferredCount").Value > 0 Then
                                SBO_Application.SetStatusBarMessage("Cannot be removed. Transactions are linked to an object, '" & oTLCodeTxt.Value & "'", SAPbouiCOM.BoMessageTime.bmt_Short, True)
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
            If (pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") And pVal.BeforeAction = False And FormID = "FTM" Then
                Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Try
                    oRs.DoQuery("Select * from [@PSSIT_OTLS]")
                    If oRs.RecordCount > 0 Then
                        SetItemEnabled()
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
    Private Sub UpdateParentTool()
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            If Not oPrtToolCombo Is Nothing Then
                If oReCondCheck.Checked = True Then
                    If oParentDB.GetValue("U_Partool", oParentDB.Offset).Trim().Length > 0 Then
                        oRs.DoQuery("Update [@PSSIT_OTLS] Set U_Active = 'N' where Code = '" & oParentDB.GetValue("U_Partool", oParentDB.Offset).Trim() & "'")
                    End If
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    Private Sub SetItemEnabled()
        Try
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                oForm.Items.Item("txttcode").Enabled = False
                oForm.Items.Item("txttname").Enabled = True
                oForm.Items.Item("cmbpartool").Enabled = False
                oForm.Items.Item("chkrecon").Enabled = False
            Else
                oForm.Items.Item("txttcode").Enabled = True
                oForm.Items.Item("txttname").Enabled = True
                oForm.Items.Item("cmbpartool").Enabled = False
                oForm.Items.Item("chkrecon").Enabled = True
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
