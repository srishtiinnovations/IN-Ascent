'''' <summary>
'''' Author                     Created Date
'''' Suresh                      19/12/2008
'''' <remarks> This class is used for entering the Operations details.</remarks>
Public Class Operations
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
    Private UFolderDs As SAPbouiCOM.UserDataSource
    Private oParentDB, oMachineDB, oLabourDB, oToolsDB As SAPbouiCOM.DBDataSource
    Private PSSIT_PRN3 As SAPbobsCOM.UserTable
    '**************************Form************************************
    Private oForm As SAPbouiCOM.Form
    '**************************ChooseFromList************************************
    Private oChAccList, oChMacList, oChToolsList, oChSkGroupList As SAPbouiCOM.ChooseFromList
    Private oChAccBtnList As SAPbouiCOM.ChooseFromList
    Private Event ChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable)
    '**************************Items - EditText************************************
    Private oOperIdTxt, oOperNameTxt, oRewAcctCodeTxt, oRewAcctNameTxt, oInfo1Txt, oInfo2Txt, oRemarksTxt, oActAcCodeTxt As SAPbouiCOM.EditText
    '**************************Items - ComboBox************************************
    Private oOperTypeCombo As SAPbouiCOM.ComboBox
    '**************************Items - CheckBox************************************
    Private oActiveCheck, oReWorkCheck As SAPbouiCOM.CheckBox
    '**************************Items - Button************************************
    Private oRewAcctBtn As SAPbouiCOM.Button
    '**************************Items - Matrix************************************
    Private oMacMatrix, oLabMatrix, oToolsMatrix As SAPbouiCOM.Matrix
    Private oMacColumns, oLabColumns, oToolsColumns As SAPbouiCOM.Columns
    Private oMacCodeCol, oMacNameCol, oMacGroupCol As SAPbouiCOM.Column
    Private oSkGroupCodeCol, oSkGroupNameCol, oReqNosCol As SAPbouiCOM.Column
    Private oToolCodeCol, oToolDescCol, oTOperCodeCol, oTMacCodeCol, oTCodeCol As SAPbouiCOM.Column
    Private oAcctCodeLink As SAPbouiCOM.LinkedButton
    '**************************Folder************************************
    Private oMacFldr, oLabFldr As SAPbouiCOM.Folder
    '**************************Variables************************************
    Private fSettings As SAPbouiCOM.FormSettings
    Private oSerialNo As Integer
    Private oToolsUID, oMacUID, oLabUID As String
    Private BoolResize As Boolean
    Private WithEvents MachineMasterClass As MachineMaster
    Private WithEvents oSkillGroupClass As SkillGroups
    Private WithEvents oToolsClass As Tools
    Private oOperationCode As String
    Private oFormName As String
#End Region
    ''' <summary>
    '''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmOperations.srf") method is called to load the Labour form.
    ''' Drawform() method is called to Initialize the form,Datasources and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company, Optional ByVal aOperationCode As String = Nothing, Optional ByVal aFormName As String = Nothing)
        MyBase.New(aSBO_Application, aCompany)
        SBO_Application = aSBO_Application
        oCompany = aCompany
        oOperationCode = aOperationCode
        oFormName = aFormName
        LoadFromXML("FrmOperations.srf")
        DrawForm()
        If oFormName = "OprRouting" Then
            oParentDB.SetValue("Code", oParentDB.Offset, oOperationCode)
            oOperIdTxt.Value = oOperationCode
            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        End If
        oForm.DataBrowser.BrowseBy = "txtocode"
        oForm.EnableMenu("1293", True)
        oForm.EnableMenu("1292", True)
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
            fSettings = oForm.Settings
            oForm.Freeze(True)
            oParentDB = oForm.DataSources.DBDataSources.Item("@PSSIT_OPRN")
            oMachineDB = oForm.DataSources.DBDataSources.Item("@PSSIT_PRN1")
            oLabourDB = oForm.DataSources.DBDataSources.Item("@PSSIT_PRN2")
            oToolsDB = oForm.DataSources.DBDataSources.Add("@PSSIT_PRN3")
            Initialize()
            AddBSDUserDataSources()
            InitializeFormComponent()
            LoadLookups()
            ConfigureMachineMatrix()
            ConfigureLabourMatrix()
            ConfigureToolsMatrix()
            oForm.Freeze(False)
            oForm.Update()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Intitializing user table PSSIT_PRN3.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Initialize()
        Try
            PSSIT_PRN3 = UserTables.Item("PSSIT_PRN3")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Initializing UserDataSources By setting UniqueID,DataType of the field and Length of the Field.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AddBSDUserDataSources()
        Try
            UFolderDs = oForm.DataSources.UserDataSources.Add("UFol", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Configuring the items/controls in the form(.srf) by bounding to the object and to the DBDatasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitializeFormComponent()
        Dim IntICount As Integer
        Dim oRs As SAPbobsCOM.Recordset
        Dim sPath As String = IO.Directory.GetParent(Application.ExecutablePath).ToString
        Try
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '**************************Folder Initialization********************************************
            For IntICount = 1 To 2
                If IntICount = 1 Then
                    oMacFldr = oForm.Items.Item("FolMachine").Specific
                    oForm.Items.Item("FolMachine").AffectsFormMode = False
                    oMacFldr.DataBind.SetBound(True, "", "UFol")
                    oMacFldr.Select()
                    fSettings.MatrixUID = "matmachine"
                    'fSettings.Enabled = True
                ElseIf IntICount = 2 Then
                    oLabFldr = oForm.Items.Item("FolLabour").Specific
                    oForm.Items.Item("FolLabour").AffectsFormMode = False
                    oLabFldr.DataBind.SetBound(True, "", "UFol")
                    oLabFldr.GroupWith("FolMachine")
                End If
            Next
            '**************************Header Data******************************************
            oOperIdTxt = oForm.Items.Item("txtocode").Specific
            oOperIdTxt.DataBind.SetBound(True, "@PSSIT_OPRN", "Code")

            oOperNameTxt = oForm.Items.Item("txtoname").Specific
            oOperNameTxt.DataBind.SetBound(True, "@PSSIT_OPRN", "U_Oprname")

            oOperTypeCombo = oForm.Items.Item("cmboprtyp").Specific
            oOperTypeCombo.DataBind.SetBound(True, "@PSSIT_OPRN", "U_Oprtype")
            oOperTypeCombo.ValidValues.Add("Internal", "Internal")
            oOperTypeCombo.ValidValues.Add("SubContract", "SubContract")
            oOperTypeCombo.ValidValues.Add("Both", "Internal-SubContract")

            oReWorkCheck = oForm.Items.Item("chkrewrk").Specific
            oReWorkCheck.DataBind.SetBound(True, "@PSSIT_OPRN", "U_Rework")

            oRewAcctCodeTxt = oForm.Items.Item("txtaccod").Specific
            oRewAcctCodeTxt.DataBind.SetBound(True, "@PSSIT_OPRN", "U_Accode")
            oForm.Items.Item("txtaccod").Enabled = False
            oForm.Items.Item("txtaccod").LinkTo = "lnkaccod"
            oForm.Items.Add("lnkaccod", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnkaccod").Visible = True
            oForm.Items.Item("lnkaccod").LinkTo = "txtaccod"
            oForm.Items.Item("lnkaccod").Top = 11
            oForm.Items.Item("lnkaccod").Left = 379
            oForm.Items.Item("lnkaccod").Description = "Link to" & vbNewLine & "Chart Of Accounts"
            oAcctCodeLink = oForm.Items.Item("lnkaccod").Specific
            oAcctCodeLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_GLAccounts

            oRewAcctBtn = oForm.Items.Item("btnacct").Specific
            oForm.Items.Item("btnacct").Description = "Choose from List" & vbNewLine & "Accounts List View"
            oRewAcctBtn.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            oRewAcctBtn.Image = sPath & "\Resources\CFL.bmp"

            oRewAcctNameTxt = oForm.Items.Item("txtacnam").Specific
            oRewAcctNameTxt.DataBind.SetBound(True, "@PSSIT_OPRN", "U_Acname")

            oInfo1Txt = oForm.Items.Item("txtadnl1").Specific
            oForm.Items.Item("lbladnl1").Visible = False
            oForm.Items.Item("txtadnl1").Visible = False
            oInfo1Txt.DataBind.SetBound(True, "@PSSIT_OPRN", "U_Adnl1")

            oInfo2Txt = oForm.Items.Item("txtadnl2").Specific
            oForm.Items.Item("lbladnl2").Visible = False
            oForm.Items.Item("txtadnl2").Visible = False
            oInfo2Txt.DataBind.SetBound(True, "@PSSIT_OPRN", "U_Adnl2")

            oRemarksTxt = oForm.Items.Item("txtremark").Specific
            oRemarksTxt.DataBind.SetBound(True, "@PSSIT_OPRN", "U_Remarks")

            oActiveCheck = oForm.Items.Item("chkactive").Specific
            oActiveCheck.DataBind.SetBound(True, "@PSSIT_OPRN", "U_Active")
            oActiveCheck.Checked = True

            oActAcCodeTxt = oForm.Items.Item("txtaccode").Specific
            oForm.Items.Item("txtaccode").Enabled = False
            'oForm.Items.Item("txtaccode").Visible = False
            'oForm.Items.Item("lblaccode").Visible = False
            oActAcCodeTxt.DataBind.SetBound(True, "@PSSIT_OPRN", "U_ActAcCode")
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' Creating ChooseFromList and Setting Conditions
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadLookups()
        Try
            oChAccBtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "1", "ACBtnLst"))
            CreateNewConditions(oChAccBtnList, "Postable", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            oRewAcctBtn = oForm.Items.Item("btnacct").Specific
            oRewAcctBtn.ChooseFromListUID = "ACBtnLst"

            oChAccList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "1", "ACTxtLst"))
            CreateNewConditions(oChAccList, "Postable", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            oRewAcctCodeTxt.ChooseFromListUID = "ACTxtLst"
            oRewAcctCodeTxt.ChooseFromListAlias = "AcctCode"

            oChMacList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_WCHDR", "MacLst"))
            CreateNewConditions(oChMacList, "U_Active", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")

            oChSkGroupList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_LGP", "SkGrpLst"))
            CreateNewConditions(oChSkGroupList, "U_Active", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")

            oChToolsList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_TLS", "ToolsLst"))
            CreateNewConditions(oChToolsList, "U_Active", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
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
    ''' Configuring the Machine Matrix items/controls in the form(.srf) by bounding to the object and to the DBDatasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ConfigureMachineMatrix()
        Try
            oMacMatrix = oForm.Items.Item("matmachine").Specific
            oMacMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oMacColumns = oMacMatrix.Columns

            'oMacRowNumCol = oColumns.Item("#")
            'oMacRowNumCol.Editable = False

            oMacCodeCol = oMacColumns.Item("colwcno")
            oMacCodeCol.Editable = True
            oMacCodeCol.DataBind.SetBound(True, "@PSSIT_PRN1", "U_wcno")
            oMacCodeCol.ChooseFromListUID = "MacLst"
            oMacCodeCol.ChooseFromListAlias = "Code"

            oMacNameCol = oMacColumns.Item("colwcnam")
            oMacNameCol.Editable = False
            oMacNameCol.DataBind.SetBound(True, "@PSSIT_PRN1", "U_wcname")

            oMacGroupCol = oMacColumns.Item("colmgnam")
            oMacGroupCol.Editable = False
            oMacGroupCol.DataBind.SetBound(True, "@PSSIT_PRN1", "U_MGname")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Configuring the Labour Matrix items/controls in the form(.srf) by bounding to the object and to the DBDatasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ConfigureLabourMatrix()
        Try
            oLabMatrix = oForm.Items.Item("matlabour").Specific
            oLabMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oLabColumns = oLabMatrix.Columns

            'oLabRowNumCol = oColumns.Item("#")
            'oLabRowNumCol.Editable = False

            'oSkGroupCodeCol = oLabColumns.Item("colcode")
            'oSkGroupCodeCol.Editable = False
            'oSkGroupCodeCol.Visible = False
            'oSkGroupCodeCol.DataBind.SetBound(True, "@PSSIT_PRN3", "Code")

            oSkGroupCodeCol = oLabColumns.Item("colskgrp")
            oSkGroupCodeCol.Editable = True
            oSkGroupCodeCol.DataBind.SetBound(True, "@PSSIT_PRN2", "U_Skilgrp")
            oSkGroupCodeCol.ChooseFromListUID = "SkGrpLst"
            oSkGroupCodeCol.ChooseFromListAlias = "Code"

            oSkGroupNameCol = oLabColumns.Item("colgnam")
            oSkGroupNameCol.Editable = False
            oSkGroupNameCol.DataBind.SetBound(True, "@PSSIT_PRN2", "U_LGname")

            oReqNosCol = oLabColumns.Item("colreqno")
            oReqNosCol.Editable = True
            oReqNosCol.DataBind.SetBound(True, "@PSSIT_PRN2", "U_Reqno")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Configuring the Tools Matrix items/controls in the form(.srf) by bounding to the object and to the DBDatasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ConfigureToolsMatrix()
        Try
            oToolsMatrix = oForm.Items.Item("mattools").Specific
            oToolsMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oToolsColumns = oToolsMatrix.Columns

            'oLabRowNumCol = oColumns.Item("#")
            'oLabRowNumCol.Editable = False

            oTCodeCol = oToolsColumns.Item("colcode")
            oTCodeCol.Editable = False
            oTCodeCol.Visible = False
            oTCodeCol.DataBind.SetBound(True, "@PSSIT_PRN3", "Code")

            oTOperCodeCol = oToolsColumns.Item("colopercod")
            oTOperCodeCol.Editable = False
            oTOperCodeCol.Visible = False
            oTOperCodeCol.DataBind.SetBound(True, "@PSSIT_PRN3", "U_Oprcode")

            oTMacCodeCol = oToolsColumns.Item("colmaccod")
            oTMacCodeCol.Editable = False
            oTMacCodeCol.DataBind.SetBound(True, "@PSSIT_PRN3", "U_wcno")

            oToolCodeCol = oToolsColumns.Item("coltolcod")
            oToolCodeCol.Editable = True
            oToolCodeCol.DataBind.SetBound(True, "@PSSIT_PRN3", "U_Toolcode")
            oToolCodeCol.ChooseFromListUID = "ToolsLst"
            oToolCodeCol.ChooseFromListAlias = "Code"

            oToolDescCol = oToolsColumns.Item("coltolnam")
            oToolDescCol.Editable = False
            oToolDescCol.DataBind.SetBound(True, "@PSSIT_PRN3", "U_TLname")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub SBO_Application_FormDataEvent1(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent
        If BusinessObjectInfo.FormUID = "FSO" Then
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD Then
                LoadToolsData()
            End If

        End If
    End Sub
    ''' <summary>
    ''' Handles all the SBO_Application event and executes as per the the event fired.
    ''' </summary>
    ''' <param name="FormUID"></param>
    ''' <param name="pVal"></param>
    ''' <param name="BubbleEvent"></param>
    ''' <remarks></remarks>
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Dim ChooseFromListEvent As SAPbouiCOM.ChooseFromListEvent
        Try
            If pVal.FormUID = "FSO" Then
                '*****************************ChooseFromList Event is called using the raiseevent*******************************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST And pVal.BeforeAction = False Then
                    ChooseFromListEvent = pVal
                    RaiseEvent ChooseFromList(pVal.ItemUID, pVal.ColUID, pVal.Row, ChooseFromListEvent.ChooseFromListUID, ChooseFromListEvent.SelectedObjects)
                End If
                '****************Matrix Link Pressed**********************************
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED) Then
                    If pVal.BeforeAction = False Then
                        If pVal.ItemUID = "matmachine" And pVal.ColUID = "colwcno" And pVal.Row > 0 Then
                            Dim oMachineNo As String
                            Dim oMachineNoEdit As SAPbouiCOM.EditText
                            oMachineNoEdit = oMacCodeCol.Cells.Item(pVal.Row).Specific
                            oMachineNo = oMachineNoEdit.Value
                            MachineMasterClass = New MachineMaster(SBO_Application, oCompany, oMachineNo, "Operation")
                        End If
                        If pVal.ItemUID = "matlabour" And pVal.ColUID = "colskgrp" And pVal.Row > 0 Then
                            Dim oSkGrpCodeEdit As SAPbouiCOM.EditText
                            oSkGrpCodeEdit = oSkGroupCodeCol.Cells.Item(pVal.Row).Specific
                            oSkillGroupClass = New SkillGroups(SBO_Application, oCompany, oSkGrpCodeEdit.Value, "Operation")
                        End If
                        If pVal.ItemUID = "mattools" And pVal.ColUID = "coltolcod" And pVal.Row > 0 Then
                            Dim oToolCodeEdit As SAPbouiCOM.EditText
                            Dim oCurrentRow As Integer
                            Try
                                oCurrentRow = pVal.Row
                                oToolsMatrix.GetLineData(pVal.Row)
                                oToolCodeEdit = oToolCodeCol.Cells.Item(oCurrentRow).Specific
                                oToolsClass = New Tools(SBO_Application, oCompany, oToolCodeEdit.Value, "Operation")
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End Try
                        End If
                    End If
                End If

                If pVal.ItemUID = "txtocode" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN Then
                    If pVal.Before_Action = True And pVal.CharPressed <> 9 Then
                        oForm = SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            BubbleEvent = False
                        End If
                    End If

                End If
                '**********************Item Pressed Event********************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.BeforeAction = True Then
                        Try
                            '***************Account Items are enabled as per the rework checbox is selected********************
                            If pVal.ItemUID = "chkrewrk" Then
                                Try
                                    If oReWorkCheck.Checked = True Then
                                        oForm.Items.Item("txtaccod").Enabled = True
                                        oForm.Items.Item("btnacct").Enabled = True
                                    ElseIf oReWorkCheck.Checked = False Then
                                        If oRewAcctCodeTxt.Value.Length > 0 Then
                                            oParentDB.SetValue("U_Accode", oParentDB.Offset, "")
                                            oParentDB.SetValue("U_Acname", oParentDB.Offset, "")
                                        End If
                                        oOperIdTxt.Active = True
                                        oForm.Items.Item("txtaccod").Enabled = False
                                        oForm.Items.Item("btnacct").Enabled = False
                                    End If
                                Catch ex As Exception
                                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    BubbleEvent = False
                                End Try
                            End If
                            '**********************Adding the child data to the database table********************
                            If pVal.ItemUID = "1" Then
                                Dim oTransaction As Boolean
                                Dim IntICount, ITools As Integer
                                Dim oCodeEdit, oToolCodeEdit, oToolDescEdit, oOprCodeEdit, oMacCodeEdit As SAPbouiCOM.EditText
                                Try
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        oForm.Freeze(True)
                                        LoadToolsData()
                                        oForm.Freeze(False)
                                    End If
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        Validation()
                                        MachinesDeleteEmptyRow()
                                        LabourDeleteEmptyRow()
                                        ToolsDeleteEmptyRow()
                                        Try
                                            If Not oCompany.InTransaction Then
                                                oCompany.StartTransaction()
                                            End If
                                            If oToolsMatrix.RowCount > 0 Then
                                                oTransaction = True
                                                For IntICount = 1 To oToolsMatrix.RowCount
                                                    oCodeEdit = oTCodeCol.Cells.Item(IntICount).Specific
                                                    oOprCodeEdit = oTOperCodeCol.Cells.Item(IntICount).Specific
                                                    oMacCodeEdit = oTMacCodeCol.Cells.Item(IntICount).Specific
                                                    oToolCodeEdit = oToolCodeCol.Cells.Item(IntICount).Specific
                                                    oToolDescEdit = oToolDescCol.Cells.Item(IntICount).Specific
                                                    oToolsMatrix.GetLineData(IntICount)
                                                    If PSSIT_PRN3.GetByKey(oCodeEdit.Value) = True Then
                                                        PSSIT_PRN3.Code = oCodeEdit.Value
                                                        PSSIT_PRN3.Name = oCodeEdit.Value
                                                        PSSIT_PRN3.UserFields.Fields.Item("U_Oprcode").Value = oOprCodeEdit.Value
                                                        PSSIT_PRN3.UserFields.Fields.Item("U_wcno").Value = oMacCodeEdit.Value
                                                        PSSIT_PRN3.UserFields.Fields.Item("U_Toolcode").Value = oToolCodeEdit.Value
                                                        PSSIT_PRN3.UserFields.Fields.Item("U_TLname").Value = oToolDescEdit.Value
                                                        ITools = PSSIT_PRN3.Update()
                                                    Else
                                                        PSSIT_PRN3.Code = oCodeEdit.Value
                                                        PSSIT_PRN3.Name = oCodeEdit.Value
                                                        PSSIT_PRN3.UserFields.Fields.Item("U_Oprcode").Value = oOprCodeEdit.Value
                                                        PSSIT_PRN3.UserFields.Fields.Item("U_wcno").Value = oMacCodeEdit.Value
                                                        PSSIT_PRN3.UserFields.Fields.Item("U_Toolcode").Value = oToolCodeEdit.Value
                                                        PSSIT_PRN3.UserFields.Fields.Item("U_TLname").Value = oToolDescEdit.Value
                                                        ITools = PSSIT_PRN3.Add()
                                                    End If
                                                Next
                                            ElseIf oToolsMatrix.RowCount = 0 Then
                                                oTransaction = True
                                            End If
                                            If oTransaction = True Then
                                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                            End If
                                        Catch ex As Exception
                                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            BubbleEvent = False
                                        Finally
                                            If oTransaction = False Then
                                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                BubbleEvent = False
                                            End If
                                        End Try
                                    End If
                                Catch ex As Exception
                                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    BubbleEvent = False
                                End Try
                            End If
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End Try
                    End If
                    '**********************If pval.beforeAction = False********************
                    If pVal.BeforeAction = False Then
                        Try
                            '**********************Refreshing the form to initiate default values********************
                            If pVal.ItemUID = "1" Then
                                If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    SetItemEnabled()
                                    If oMacMatrix.RowCount > 0 Then
                                        oMacMatrix.SelectRow(1, True, False)
                                    End If
                                End If
                                If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    DeleteRowFromDB()
                                    oForm.Refresh()
                                    oForm.Freeze(True)
                                    SetItemEnabled()
                                    oOperIdTxt.Active = True
                                    oActiveCheck.Checked = True
                                    oForm.Freeze(False)
                                End If
                            End If
                            '**********************Setting the pane level as per the folder selected********************
                            oForm.Freeze(True)
                            If pVal.ItemUID = "FolMachine" Then
                                oForm.PaneLevel = 1
                                oForm.Freeze(True)
                                SetItemEnabled()
                                oForm.Freeze(False)
                                If Not oMacMatrix Is Nothing Then
                                    fSettings.MatrixUID = "matmachine"
                                    'fSettings.Enabled = True
                                End If
                            End If
                            If pVal.ItemUID = "FolLabour" Then
                                oForm.PaneLevel = 2
                                SetItemEnabled()
                                If Not oLabMatrix Is Nothing Then
                                    fSettings.MatrixUID = "matlabour"
                                    'fSettings.Enabled = True
                                End If
                            End If
                            oForm.Freeze(False)
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        End Try
                    End If
                End If
                '**********************Validating the values in the items********************
                If pVal.CharPressed = Keys.Tab And pVal.ItemUID = "txtocode" Then
                    If pVal.BeforeAction = True Then
                        Dim oStrSql As String
                        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Try
                            If oOperIdTxt.Value <> "" Then
                                If oOperIdTxt.Value.Length > 0 Then
                                    oStrSql = "Select * from [@PSSIT_OPRN] where Code = '" & oOperIdTxt.Value & "'"
                                    oRs.DoQuery(oStrSql)
                                    If oRs.RecordCount > 0 Then
                                        SBO_Application.SetStatusBarMessage("Operation Id '" & oOperIdTxt.Value & "' already exists", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        BubbleEvent = False
                                    End If
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
                    '**********************Adding a row to the machine matrix and labour matrix********************
                    If pVal.BeforeAction = False Then
                        Try
                            If oOperIdTxt.Value <> "" Then
                                If oOperIdTxt.Value.Length > 0 Then
                                    oMachineDB.InsertRecord(oMachineDB.Size)
                                    oMachineDB.Offset = oMachineDB.Size - 1
                                    oMacMatrix.Clear()
                                    oMacMatrix.FlushToDataSource()
                                    oMacMatrix.AddRow(1, oMacMatrix.RowCount)

                                    oLabourDB.InsertRecord(oLabourDB.Size)
                                    oLabourDB.Offset = oLabourDB.Size - 1
                                    oLabMatrix.Clear()
                                    oLabMatrix.FlushToDataSource()
                                    oLabMatrix.AddRow(1, oLabMatrix.RowCount)
                                End If
                            End If
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        End Try
                    End If
                End If
                '**********************Selecting a row in th machine matrix********************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.BeforeAction = False Then
                    Try
                        If pVal.ItemUID = "matmachine" And pVal.Row > 0 Then
                            Try
                                If (pVal.ColUID = "#" Or pVal.ColUID = "colwcno" Or pVal.ColUID = "colwcnam" Or pVal.ColUID = "colmgnam") Then
                                    oMacMatrix.SelectRow(pVal.Row, True, False)
                                End If
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End Try
                        End If
                        If (pVal.ItemUID = "mattools") And pVal.ColUID = "#" Then
                            Try
                                oMacUID = ""
                                oToolsUID = pVal.ItemUID
                                oLabUID = ""
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End Try
                        End If
                        If (pVal.ItemUID = "matmachine") And pVal.ColUID = "#" Then
                            Try
                                oMacUID = pVal.ItemUID
                                oToolsUID = ""
                                oLabUID = ""
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End Try
                        End If
                        If (pVal.ItemUID = "matlabour") And pVal.ColUID = "#" Then
                            Try
                                oMacUID = ""
                                oToolsUID = ""
                                oLabUID = pVal.ItemUID
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End Try
                        End If
                    Catch ex As Exception
                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    End Try
                End If
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_RESIZE And pVal.BeforeAction = False Then
                    Try
                        'Form_Resize()
                    Catch ex As Exception
                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    End Try
                End If
                '*****************************Releasing the Com Object*******************************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                    SBO_Application = Nothing
                    GC.Collect()
                End If
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    Private Sub DeleteRowFromDB()
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            oRs.DoQuery("Delete From [@PSSIT_PRN1] Where U_wcno is null and U_Wcname is null and U_MGname is null")
            oRs.DoQuery("Delete From [@PSSIT_PRN2] Where U_Skilgrp is null and U_LGname is null")
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' Choosing AccountCode,MachineCode,LabourCode,ToolCode from the CFL and setting the values 
    ''' to the corresponding field.
    ''' </summary>
    ''' <param name="ControlName"></param>
    ''' <param name="ColumnUID"></param>
    ''' <param name="CurrentRow"></param>
    ''' <param name="ChoosefromListUID"></param>
    ''' <param name="ChooseFromListSelectedObjects"></param>
    ''' <remarks></remarks>
    Private Sub Operations_ChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable) Handles Me.ChooseFromList
        Dim oDataTable As SAPbouiCOM.DataTable = ChooseFromListSelectedObjects
        Dim oAcctCode, oAcctName, oMacCode, oMacName, oMacGroup, oSkGrpCode, oSkGrpName, oToolCode, oToolName As String
        Dim oCurrentRow As Integer
        Try
            oCurrentRow = CurrentRow

            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                '*********************************Account Code**********************************
                If (ControlName = "btnacct" Or ControlName = "txtaccod") And (ChoosefromListUID = "ACBtnLst" Or ChoosefromListUID = "ACTxtLst") Then
                    If Not oDataTable Is Nothing Then
                        oAcctCode = oDataTable.GetValue("FormatCode", 0)
                        oAcctName = oDataTable.GetValue("AcctName", 0)
                        oParentDB.Offset = oParentDB.Size - 1
                        oParentDB.SetValue("U_Accode", oParentDB.Offset, FormatAccountCode(oAcctCode))
                        oParentDB.SetValue("U_Acname", oParentDB.Offset, oAcctName)
                        oParentDB.SetValue("U_ActAcCode", oParentDB.Offset, oAcctCode.ToString().Replace("-", ""))
                    End If
                End If
                '*********************************Machines**********************************
                If ControlName = "matmachine" And ChoosefromListUID = "MacLst" Then
                    If Not oDataTable Is Nothing Then
                        'Added by Manimaran-----s
                        Dim code As String
                        Dim i As Integer
                        If oMacMatrix.RowCount > 0 Then
                            For i = 1 To oMacMatrix.RowCount
                                oMacMatrix.GetLineData(i)
                                code = oMacMatrix.Columns.Item("colwcno").Cells.Item(i).Specific.string
                                If code = oDataTable.GetValue("U_wcno", 0) Then
                                    Throw New Exception("Selected Machine code found already exists")
                                End If
                            Next
                        End If
                        'Added by Manimaran-----e
                        oMacCode = oDataTable.GetValue("U_wcno", 0)
                        oMacName = oDataTable.GetValue("U_wcname", 0)
                        oMacGroup = oDataTable.GetValue("U_MGcode", 0)
                        If CurrentRow = oMacMatrix.VisualRowCount Then
                            oMachineDB.Offset = oMachineDB.Size - 1
                            SetMachineDefaultValue()
                            oMacMatrix.SetLineData(CurrentRow)
                            oMacMatrix.FlushToDataSource()
                        End If
                        oMachineDB.SetValue("U_wcno", oMachineDB.Offset, oMacCode)
                        oMachineDB.SetValue("U_wcname", oMachineDB.Offset, oMacName)
                        oMachineDB.SetValue("U_MGname", oMachineDB.Offset, oMacGroup)
                        oMacMatrix.SetLineData(CurrentRow)
                        oMacMatrix.Columns.Item("#").Cells.Item(oMacMatrix.RowCount).Specific.string = oMacMatrix.RowCount
                        oMacMatrix.FlushToDataSource()
                    End If
                    If Len(oMacCodeCol.Cells.Item(oMacMatrix.RowCount).Specific.value) > 0 Then
                        'oMacMatrix.GetLineData(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                        'oToolsDB.InsertRecord(oToolsDB.Size)
                        'AddToolsRow(oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                        'oToolsDB.SetValue("Code", oToolsDB.Offset, oSerialNo)
                        'oToolsMatrix.SetLineData(oToolsMatrix.RowCount)
                        '*********************************Machines**********************************
                        oMachineDB.InsertRecord(oMachineDB.Size)
                        oMachineDB.Offset = oMachineDB.Size - 1
                        SetMachineDefaultValue()

                        oMacMatrix.AddRow(1, oMacMatrix.RowCount)
                        '****************************************************************************
                    End If
                End If
                '*********************************Labour**********************************
                If ControlName = "matlabour" And ChoosefromListUID = "SkGrpLst" Then
                    If Not oDataTable Is Nothing Then
                        oSkGrpCode = oDataTable.GetValue("Code", 0)
                        oSkGrpName = oDataTable.GetValue("U_LGname", 0)
                        If CurrentRow = oLabMatrix.VisualRowCount Then
                            oLabourDB.Offset = oLabourDB.Size - 1
                            SetLabourDefaultValue()
                            oLabMatrix.SetLineData(CurrentRow)
                            oLabMatrix.FlushToDataSource()
                        End If
                        oLabourDB.SetValue("U_Skilgrp", oLabourDB.Offset, oSkGrpCode)
                        oLabourDB.SetValue("U_LGname", oLabourDB.Offset, oSkGrpName)
                        oLabMatrix.SetLineData(CurrentRow)
                        oLabMatrix.Columns.Item("#").Cells.Item(oLabMatrix.RowCount).Specific.string = oLabMatrix.RowCount
                        oLabMatrix.FlushToDataSource()
                    End If

                    If Len(oSkGroupCodeCol.Cells.Item(oLabMatrix.RowCount).Specific.value) > 0 Then
                        oLabourDB.InsertRecord(oLabourDB.Size)
                        oLabourDB.Offset = oLabourDB.Size - 1
                        SetLabourDefaultValue()
                        If oLabMatrix.Columns.Item("colskgrp").Cells.Item(oLabMatrix.RowCount).Specific.value <> "" Then
                            oLabMatrix.AddRow(1, oLabMatrix.RowCount)
                            oLabMatrix.Columns.Item("#").Cells.Item(oLabMatrix.RowCount).Specific.string = oLabMatrix.RowCount
                        End If
                    End If
                End If

                'If ControlName = "matlabour" And ChoosefromListUID = "SkGrpLst" Then
                '    If Not oDataTable Is Nothing Then
                '        oSkGrpCode = oDataTable.GetValue("Code", 0)
                '        oSkGrpName = oDataTable.GetValue("U_LGname", 0)
                '        If CurrentRow = oLabMatrix.VisualRowCount Then
                '            oLabourDB.Offset = oLabourDB.Size - 1
                '            SetLabourDefaultValue(oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value, _
                '            oMOprCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                '            oLabMatrix.SetLineData(CurrentRow)
                '            oLabMatrix.FlushToDataSource()
                '        End If
                '        oLabMatrix.GetLineData(CurrentRow)
                '        oLabourDB.SetValue("U_Skilgrp", oLabourDB.Offset, oSkGrpCode)
                '        oLabourDB.SetValue("U_LGname", oLabourDB.Offset, oSkGrpName)
                '        oLabourDB.SetValue("U_wcno", oLabourDB.Offset, oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                '        oLabMatrix.SetLineData(CurrentRow)
                '        oLabMatrix.FlushToDataSource()
                '    End If
                'End If
                ' '*********************************Tools**********************************
                If ControlName = "mattools" And ChoosefromListUID = "ToolsLst" Then
                    If Not oDataTable Is Nothing Then
                        oToolCode = oDataTable.GetValue("Code", 0)
                        oToolName = oDataTable.GetValue("U_TLname", 0)
                        If CurrentRow = oToolsMatrix.VisualRowCount Then
                            'oMacMatrix.GetLineData(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                            oToolsDB.Offset = oToolsDB.Size - 1
                            SetToolsDefaultValue(oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.value)
                            oToolsMatrix.SetLineData(CurrentRow)
                            oToolsMatrix.FlushToDataSource()
                        End If
                        oToolsMatrix.GetLineData(CurrentRow)
                        oToolsDB.SetValue("U_Oprcode", oToolsDB.Offset, oOperIdTxt.Value)
                        oToolsDB.SetValue("U_wcno", oToolsDB.Offset, oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                        oToolsDB.SetValue("U_Toolcode", oToolsDB.Offset, oToolCode)
                        oToolsDB.SetValue("U_TLname", oToolsDB.Offset, oToolName)
                        oToolsMatrix.SetLineData(CurrentRow)
                        oToolsMatrix.FlushToDataSource()
                    End If
                    'If Len(oToolCodeCol.Cells.Item(oToolsMatrix.RowCount).Specific.value) > 0 Then
                    '    oToolsDB.InsertRecord(oToolsDB.Size)
                    '    'oToolsDB.Offset = oToolsDB.Size - 1
                    '    'SetToolsDefaultValue(oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.value)
                    '    'oToolsMatrix.AddRow(1, oToolsMatrix.RowCount)
                    '    AddToolsRow(oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                    '    oToolsDB.SetValue("Code", oToolsDB.Offset, oSerialNo)
                    '    oToolsMatrix.SetLineData(oToolsMatrix.RowCount)
                    'End If
                End If
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    ''' <summary>
    ''' SetItemEnabled() method is called to set the item enabled as per the form mode.
    ''' setting the focus to the OperationId EditText.
    ''' </summary>
    ''' <param name="pVal"></param>
    ''' <param name="BubbleEvent"></param>
    ''' <remarks></remarks>
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Dim FormID As String
        Dim IntICount As Integer
        Dim oMacCode As String
        Dim oToolsDelCode As SAPbouiCOM.EditText
        Try
            FormID = SBO_Application.Forms.ActiveForm.UniqueID
            If pVal.MenuUID = "1281" And FormID = "FSO" Then
                If pVal.BeforeAction = False Then
                    oOperIdTxt.Active = True
                End If
            End If
            If pVal.MenuUID = "1282" And pVal.BeforeAction = False And FormID = "FSO" Then
                SetItemEnabled()
                oOperIdTxt.Active = True
                oActiveCheck.Checked = True
            End If
            If pVal.MenuUID = "1283" And FormID = "FSO" Then
                If pVal.BeforeAction = True Then
                    Dim oStrSql As String
                    Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Try
                        oStrSql = "Select IsNull(Count(*),0) as ReferredCount from [@PSSIT_RTE4] where U_OprCode = '" & oOperIdTxt.Value & "'"
                        oRs.DoQuery(oStrSql)
                        If oRs.RecordCount > 0 Then
                            oRs.MoveFirst()
                            If oRs.Fields.Item("ReferredCount").Value > 0 Then
                                SBO_Application.SetStatusBarMessage("Cannot be removed. Transactions are linked to an object, '" & oOperIdTxt.Value & "'", SAPbouiCOM.BoMessageTime.bmt_Short, True)
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
            '*****************************Adding a row to the toolsmatrix*******************************
            If pVal.MenuUID = "1292" And pVal.BeforeAction = True Then
                If oMacUID = "matmachine" Then
                    If oMacMatrix.RowCount = 0 Then
                        If oOperIdTxt.Value <> "" Then
                            If oOperIdTxt.Value.Length > 0 Then
                                oMachineDB.InsertRecord(oMachineDB.Size)
                                oMachineDB.Offset = oMachineDB.Size - 1
                                oMacMatrix.Clear()
                                SetMachineDefaultValue()
                                oMacMatrix.FlushToDataSource()
                                oMacMatrix.AddRow(1, oMacMatrix.RowCount)
                            End If
                        End If
                    ElseIf oMacMatrix.RowCount > 0 Then
                        If oOperIdTxt.Value <> "" Then
                            If oOperIdTxt.Value.Length > 0 Then
                                oMachineDB.Offset = oMachineDB.Size - 1
                                SetMachineDefaultValue()
                                If oMacMatrix.Columns.Item("colwcno").Cells.Item(oMacMatrix.RowCount).Specific.value <> "" Then
                                    oMacMatrix.AddRow(1, oMacMatrix.RowCount)
                                    oMacMatrix.Columns.Item("#").Cells.Item(oMacMatrix.RowCount).Specific.string = oMacMatrix.RowCount

                                End If

                            End If
                        End If
                    End If
                End If
                If oLabUID = "matlabour" Then
                    If oLabMatrix.RowCount = 0 Then
                        If oOperIdTxt.Value <> "" Then
                            If oOperIdTxt.Value.Length > 0 Then
                                oLabourDB.InsertRecord(oLabourDB.Size)
                                oLabourDB.Offset = oLabourDB.Size - 1
                                oLabMatrix.Clear()
                                SetLabourDefaultValue()
                                oLabMatrix.FlushToDataSource()
                                'Modified by Manimaran-----s
                                'If oLabMatrix.Columns.Item("colskgrp").Cells.Item(oLabMatrix.RowCount).Specific.value <> "" Then
                                '    oLabMatrix.AddRow(1, oLabMatrix.RowCount)
                                'End If
                                oLabMatrix.AddRow(1, oLabMatrix.RowCount)
                                'Modified by Manimaran-----e
                            End If
                        End If
                    ElseIf oLabMatrix.RowCount > 0 Then
                        If oOperIdTxt.Value <> "" Then
                            If oOperIdTxt.Value.Length > 0 Then
                                oLabourDB.Offset = oLabourDB.Size - 1
                                SetLabourDefaultValue()
                                ' oLabMatrix.AddRow(1, oLabMatrix.RowCount)
                                If oLabMatrix.Columns.Item("colskgrp").Cells.Item(oLabMatrix.RowCount).Specific.value <> "" Then
                                    oLabMatrix.AddRow(1, oLabMatrix.RowCount)
                                    oLabMatrix.Columns.Item("#").Cells.Item(oLabMatrix.RowCount).Specific.string = oLabMatrix.RowCount
                                End If
                            End If
                        End If
                    End If
                End If
                If oToolsUID = "mattools" And FormID = "FSO" Then
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                    If oToolsMatrix.RowCount = 0 Then
                        oToolsDB.InsertRecord(oToolsDB.Size)
                        If oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder) > 0 Then
                            AddToolsRow(oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                            oToolsDB.SetValue("Code", oToolsDB.Offset, oSerialNo)
                            oToolsMatrix.SetLineData(oToolsMatrix.RowCount)
                        Else
                            SBO_Application.SetStatusBarMessage("Select the Machine for which the Tools to be added", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                    ElseIf oToolsMatrix.RowCount > 0 Then
                        If oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder) > 0 Then
                            If Len(oToolCodeCol.Cells.Item(oToolsMatrix.RowCount).Specific.value) <= 0 Then
                                SBO_Application.SetStatusBarMessage("Tool Details should not be empty", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                            End If
                            If Len(oToolCodeCol.Cells.Item(oToolsMatrix.RowCount).Specific.value) > 0 Then
                                AddToolsRow(oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                                oToolsDB.SetValue("Code", oToolsDB.Offset, oSerialNo)
                                oToolsMatrix.SetLineData(oToolsMatrix.RowCount)
                            End If
                        Else
                            SBO_Application.SetStatusBarMessage("Select the Machine for which the Tools to be added", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                    End If
                End If
            End If
            If pVal.MenuUID = "1293" And pVal.BeforeAction = True And FormID = "FSO" Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End If
                If oMacUID = "matmachine" Then
                    If oMacMatrix.RowCount > 0 Then
                        oMacCode = oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value
                        For IntICount = oToolsMatrix.RowCount To 1 Step -1
                            oToolsMatrix.GetLineData(IntICount)
                            If oMacCode = oTMacCodeCol.Cells.Item(IntICount).Specific.Value Then
                                oToolsDelCode = oTCodeCol.Cells.Item(IntICount).Specific
                                If PSSIT_PRN3.GetByKey(oToolsDelCode.Value) = True Then
                                    Dim I As Integer = PSSIT_PRN3.Remove()
                                    oToolsMatrix.DeleteRow(IntICount)
                                    oToolsMatrix.FlushToDataSource()
                                ElseIf PSSIT_PRN3.GetByKey(oToolsDelCode.Value) = False Then
                                    oToolsMatrix.DeleteRow(IntICount)
                                    oToolsMatrix.FlushToDataSource()
                                    oToolsMatrix.LoadFromDataSource()
                                End If
                            End If
                        Next
                        oMacMatrix.DeleteRow(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                        oMacMatrix.FlushToDataSource()
                    End If
                End If
                If oLabUID = "matlabour" Then
                    If oLabMatrix.RowCount > 0 Then
                        oLabMatrix.DeleteRow(oLabMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                        oLabMatrix.FlushToDataSource()
                    End If
                End If
                If oToolsUID = "mattools" Then
                    oToolsDelCode = oTCodeCol.Cells.Item(oToolsMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific
                    If PSSIT_PRN3.GetByKey(oToolsDelCode.Value) = True Then
                        Dim I As Integer = PSSIT_PRN3.Remove()
                        oToolsMatrix.DeleteRow(oToolsMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                        oToolsMatrix.FlushToDataSource()
                    Else
                        oToolsMatrix.DeleteRow(oToolsMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                        oToolsMatrix.FlushToDataSource()
                    End If
                End If
                SetItemEnabled()
                BubbleEvent = False
            End If
            '*****************************LoadToolsData() is called.*******************************
            If (pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") And pVal.BeforeAction = False And FormID = "FSO" Then
                Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Try
                    oRs.DoQuery("Select * from [@PSSIT_OPRN]")
                    If oRs.RecordCount > 0 Then
                        oForm.Freeze(True)
                        If oMacMatrix.RowCount > 0 Then
                            oMacMatrix.SelectRow(1, True, False)
                        End If
                        LoadToolsData()
                        SetItemEnabled()
                        oForm.Freeze(False)
                    Else
                        oOperIdTxt.Active = True
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
    ''' Resizing the form 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Form_Resize()
        Try
            If BoolResize = False Then
                oForm.Freeze(True)

                oForm.Items.Item("chkrewrk").Left = oForm.Items.Item("lblaccod").Left
                oForm.Items.Item("rectrew").Left = oForm.Items.Item("lblaccod").Left - 5
                oForm.Items.Item("rectrew").Height = 55
                oForm.Items.Item("rectrew").Top = 5
                'oForm.Items.Item("rectrew").Width = 282

                oForm.Items.Item("RectMac").Height = 400
                oForm.Items.Item("RectMac").Top = 78
                oForm.Items.Item("RectMac").Width = oForm.Width - 20
                oForm.Items.Item("RectLab").Height = 400
                oForm.Items.Item("RectLab").Top = 78
                oForm.Items.Item("RectLab").Width = oForm.Width - 20

                oForm.Items.Item("FolMachine").Top = 59
                oForm.Items.Item("FolLabour").Top = 59

                oForm.Items.Item("matmachine").Left = 10
                oForm.Items.Item("matmachine").Height = 235
                oForm.Items.Item("matmachine").Top = oForm.Items.Item("RectMac").Top + 5
                oForm.Items.Item("matmachine").Width = oForm.Items.Item("RectMac").Width - 10

                oForm.Items.Item("LblTools").Left = 10
                oForm.Items.Item("LblTools").Top = oForm.Items.Item("matmachine").Top + oForm.Items.Item("matmachine").Height + 5

                oForm.Items.Item("mattools").Left = 10
                oForm.Items.Item("mattools").Height = 135
                oForm.Items.Item("mattools").Top = oForm.Items.Item("LblTools").Top + oForm.Items.Item("LblTools").Height + 5
                oForm.Items.Item("mattools").Width = oForm.Items.Item("RectMac").Width - 10

                oForm.Items.Item("matlabour").Left = 10
                oForm.Items.Item("matlabour").Height = oForm.Items.Item("RectLab").Height - 10
                oForm.Items.Item("matlabour").Top = oForm.Items.Item("RectLab").Top + 5
                oForm.Items.Item("matlabour").Width = oForm.Items.Item("RectLab").Width - 10
                oForm.Freeze(False)
                oForm.Update()
                BoolResize = True
            ElseIf BoolResize = True Then
                oForm.Freeze(True)

                oForm.Items.Item("chkrewrk").Left = 275
                oForm.Items.Item("rectrew").Left = 270
                oForm.Items.Item("rectrew").Height = 55
                oForm.Items.Item("rectrew").Top = 5
                'oForm.Items.Item("rectrew").Width = 282

                oForm.Items.Item("RectMac").Height = 235
                oForm.Items.Item("RectMac").Top = 78
                oForm.Items.Item("RectMac").Width = 547
                oForm.Items.Item("RectLab").Height = 235
                oForm.Items.Item("RectLab").Top = 78
                oForm.Items.Item("RectLab").Width = 547

                oForm.Items.Item("FolMachine").Top = 59
                oForm.Items.Item("FolLabour").Top = 59

                oForm.Items.Item("matmachine").Left = 10
                oForm.Items.Item("matmachine").Height = 111
                oForm.Items.Item("matmachine").Top = 83
                oForm.Items.Item("matmachine").Width = 537

                oForm.Items.Item("LblTools").Left = 10
                oForm.Items.Item("LblTools").Top = 197

                oForm.Items.Item("mattools").Left = 10
                oForm.Items.Item("mattools").Height = 92
                oForm.Items.Item("mattools").Top = 215
                oForm.Items.Item("mattools").Width = 537

                oForm.Items.Item("matlabour").Left = 10
                oForm.Items.Item("matlabour").Height = 225
                oForm.Items.Item("matlabour").Top = 83
                oForm.Items.Item("matlabour").Width = 537
                BoolResize = False
                oForm.Freeze(False)
                oForm.Update()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Setting default value to the newly inserting column in the datasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetMachineDefaultValue()
        Try
            oMachineDB.SetValue("U_wcno", oMachineDB.Offset, "")
            oMachineDB.SetValue("U_wcname", oMachineDB.Offset, "")
            oMachineDB.SetValue("U_MGname", oMachineDB.Offset, "")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Setting default value to the newly inserting column in the datasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetLabourDefaultValue()
        Try
            oLabourDB.SetValue("U_Skilgrp", oLabourDB.Offset, "")
            oLabourDB.SetValue("U_LGname", oLabourDB.Offset, "")
            oLabourDB.SetValue("U_Reqno", oLabourDB.Offset, "0")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This method is used to add a row in the ToolsDB.
    ''' </summary>
    ''' <param name="oMacCode"></param>
    ''' <remarks></remarks>
    Private Sub AddToolsRow(ByVal oMacCode As String)
        Try
            oToolsDB.Offset = oToolsDB.Size - 1
            SetToolsDefaultValue(oMacCode)
            oToolsMatrix.AddRow(1, oToolsMatrix.RowCount)
            If oToolsMatrix.RowCount = 1 Then
                oSerialNo = GenerateSerialNo("PSSIT_PRN3")
            ElseIf oToolsMatrix.RowCount > 1 Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    oSerialNo = oSerialNo + 1
                ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    oSerialNo = oSerialNo + 1
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This method is used to set default value to the newly inserting column in the datasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetToolsDefaultValue(ByVal oMacCode As String)
        Try
            'oToolsDB.SetValue("Code", oToolsDB.Offset, "")
            oToolsDB.SetValue("U_Oprcode", oToolsDB.Offset, oOperIdTxt.Value)
            oToolsDB.SetValue("U_wcno", oToolsDB.Offset, oMacCode)
            oToolsDB.SetValue("U_Toolcode", oToolsDB.Offset, "")
            oToolsDB.SetValue("U_TLname", oToolsDB.Offset, "")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This method is used to delete the empty rows in the Machine Matrix.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub MachinesDeleteEmptyRow()
        Dim oMacCodeEdit, oMacNameEdit As SAPbouiCOM.EditText
        Dim IntICount As Integer
        Try
            For IntICount = oMacMatrix.RowCount To 1 Step -1
                oMacMatrix.GetLineData(IntICount)
                oMacCodeEdit = oMacCodeCol.Cells.Item(IntICount).Specific
                oMacNameEdit = oMacNameCol.Cells.Item(IntICount).Specific
                If oMacCodeEdit.Value.Length = 0 And oMacNameEdit.Value.Length = 0 Then
                    oMacMatrix.DeleteRow(IntICount)
                    oMacMatrix.FlushToDataSource()
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This method is used to delete the empty rows in the Labour Matrix.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LabourDeleteEmptyRow()
        Dim oSkGrpCodeEdit, oSkGrpNameEdit As SAPbouiCOM.EditText
        Dim IntICount As Integer
        Try
            For IntICount = oLabMatrix.RowCount To 1 Step -1
                oLabMatrix.GetLineData(IntICount)
                oSkGrpCodeEdit = oSkGroupCodeCol.Cells.Item(IntICount).Specific
                oSkGrpNameEdit = oSkGroupNameCol.Cells.Item(IntICount).Specific
                If oSkGrpCodeEdit.Value.Length = 0 And oSkGrpNameEdit.Value.Length = 0 Then
                    oLabMatrix.DeleteRow(IntICount)
                    oLabMatrix.FlushToDataSource()
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This method is used to delete the empty rows in the Tools Matrix.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ToolsDeleteEmptyRow()
        Dim oToolCodeEdit, oToolNameEdit As SAPbouiCOM.EditText
        Dim IntICount As Integer
        Try
            For IntICount = oToolsMatrix.RowCount To 1 Step -1
                oToolsMatrix.GetLineData(IntICount)
                oToolCodeEdit = oToolCodeCol.Cells.Item(IntICount).Specific
                oToolNameEdit = oToolDescCol.Cells.Item(IntICount).Specific
                If oToolCodeEdit.Value.Length = 0 And oToolNameEdit.Value.Length = 0 Then
                    oToolsMatrix.DeleteRow(IntICount)
                    oToolsMatrix.FlushToDataSource()
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    '''  This method is used for validating the values in the EditText.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Validation()
        Dim IntICount As Integer
        Dim oSkGrpCode, oReqNos As SAPbouiCOM.EditText
        Try
            If oOperIdTxt.Value.Length = 0 Then
                oOperIdTxt.Active = True
                Throw New Exception("Operation Id should not be empty")
            End If
            If oOperNameTxt.Value.Length = 0 Then
                oOperNameTxt.Active = True
                Throw New Exception("Operation Name should not be empty")
            ElseIf oOperNameTxt.Value.Length > 30 Then
                oOperNameTxt.Active = True
                Throw New Exception("Operation Name should be less than 30 charactor")
            End If
            If oParentDB.GetValue("U_Oprtype", oParentDB.Offset).Length = 0 Then
                oOperTypeCombo.Active = True
                Throw New Exception("Operation Type should not be empty")
            End If
            'If oToolsMatrix.RowCount <= 0 Then
            '    Throw New Exception("Atleast one Tools should be added")
            'End If
            If oParentDB.GetValue("U_Oprtype", oParentDB.Offset).Trim() <> "SubContract" Then
                If oMacMatrix.RowCount = 1 Then
                    oMacMatrix.GetLineData(oMacMatrix.RowCount)
                    If oMacCodeCol.Cells.Item(oMacMatrix.RowCount).Specific.Value.Length = 0 And oMacNameCol.Cells.Item(oMacMatrix.RowCount).Specific.Value.Length = 0 Then
                        oMacFldr.Select()
                        oMacCodeCol.Cells.Item(oMacMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Throw New Exception("Atleast one Machine should be added")
                    End If
                End If
                If oLabMatrix.RowCount = 1 Then
                    oLabMatrix.GetLineData(oLabMatrix.RowCount)
                    If oSkGroupCodeCol.Cells.Item(oLabMatrix.RowCount).Specific.Value.Length = 0 And oSkGroupNameCol.Cells.Item(oLabMatrix.RowCount).Specific.Value.Length = 0 Then
                        oLabFldr.Select()
                        oSkGroupCodeCol.Cells.Item(oLabMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Throw New Exception("Atleast one Labour should be added")
                    End If
                End If
            End If
            If oLabMatrix.RowCount > 1 Then
                For IntICount = 1 To oLabMatrix.RowCount
                    oLabMatrix.GetLineData(IntICount)
                    oReqNos = oReqNosCol.Cells.Item(IntICount).Specific
                    oSkGrpCode = oSkGroupCodeCol.Cells.Item(IntICount).Specific
                    If oSkGrpCode.Value.Length > 0 Then
                        If oReqNos.Value.Length = 0 Or oReqNos.Value = "0" Then
                            oLabFldr.Select()
                            oReqNosCol.Cells.Item(IntICount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Throw New Exception("Required Nos. should be entered")
                        ElseIf oReqNos.Value < 0 Then
                            oLabFldr.Select()
                            oReqNosCol.Cells.Item(IntICount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Throw New Exception("Required Nos.can't be negative value")

                        End If
                    End If
                Next
            End If
            'Added by Manimaran------s
            If oForm.Items.Item("txtaccod").Specific.value.length = 0 Then
                Throw New Exception("Account Code should not be empty")
            End If
            'Added by Manimaran------e
            AccKeyCheck()
        Catch ex As Exception
            Throw ex
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
                If oRewAcctCodeTxt.Value.Length = 0 Then
                    oRewAcctCodeTxt.Active = True
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
    ''' <summary>
    ''' Loading the tools from the database as per the conditions.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadToolsData()
        Dim oConditions As SAPbouiCOM.Conditions
        Dim oCondition As SAPbouiCOM.Condition
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            oConditions = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions)
            oCondition = oConditions.Add
            oCondition.Alias = "U_Oprcode"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = oOperIdTxt.Value
            oToolsDB.Query(oConditions)
            oToolsMatrix.LoadFromDataSource()
            oToolsMatrix.FlushToDataSource()
            oRs.DoQuery("Select IsNull(Max(Code),0) as Code from [@PSSIT_PRN3]")
            oSerialNo = oRs.Fields.Item("Code").Value
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' Enabling th Items in th form as per the form mode.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetItemEnabled()
        Try
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                ' oForm.Items.Item("txtocode").Enabled = False
                oForm.Items.Item("txtoname").Enabled = True
                oForm.Items.Item("cmboprtyp").Enabled = False
                ''If oReWorkCheck.Checked = True Then
                ''    oForm.Items.Item("txtaccod").Enabled = True
                ''    oForm.Items.Item("btnacct").Enabled = True
                ''ElseIf oReWorkCheck.Checked = False Then
                oForm.Items.Item("txtaccod").Enabled = False
                oForm.Items.Item("btnacct").Enabled = False
                'End If
            Else
            oForm.Items.Item("txtocode").Enabled = True
            oForm.Items.Item("txtoname").Enabled = True
            oForm.Items.Item("cmboprtyp").Enabled = True
                'If Not oReWorkCheck Is Nothing Then
                '    If oReWorkCheck.Checked = True Then
                oForm.Items.Item("txtaccod").Enabled = True
                oForm.Items.Item("btnacct").Enabled = True
                '    ElseIf oReWorkCheck.Checked = False Then
                '        oForm.Items.Item("txtaccod").Enabled = False
                '        oForm.Items.Item("btnacct").Enabled = False
                '    End If
                'End If
            End If
        Catch ex As Exception
            'Throw ex
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
