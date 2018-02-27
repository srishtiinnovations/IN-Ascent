''' <summary>
''' Author                      Created Date
''' Suresh                       03/12/2008
''' </summary>
''' <remarks>This class is used for adding menus to the main menu and redirecting to the form.</remarks>
Public Class Menus
    Inherits ConnectionLib
#Region "Menus"
    ''' <summary>
    ''' Variable Declaration
    ''' </summary>
    ''' <remarks></remarks>
#Region "Variable Declaration"
    Private formCmdCenter As SAPbouiCOM.Form
    'Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private Shift As Shift
    Private WorkCentre As WorkCentre
    Private MachineGroups As MachineGroups
    Private MachineMaster As MachineMaster
    Private SkillGroups As SkillGroups
    Private Labour As Labour
    Private Tools As Tools
    Private Stoppage As Stoppage
    Private Operations As Operations
    Private OperationsRouting As OperationsRouting
    Private ProductionSetup As ProductionSetup
    Private ProductionEntry As ProductionEntry
    Private MachineDownTime As MachineDownTime
    Private MasterReports As MasterReports
    Private ProdCostReport As ProdCostReport
    Private ProductionReport As ProductionReport
    Private MCUtilReport As MCUtilReport
    Private MCPerfReport As MCPerfReport
    Private LabrPerfReport As LabrPerfReport
    Private ImagePath As String = ""
    Private sPath As String = IO.Directory.GetParent(Application.ExecutablePath).ToString
    Private oGeneralMenu
    Private oMaster
    Private oTransaction
    Private oReports
    Private IssProdNo As Integer
#End Region
#Region "Variable Declaration Production Order"
    '**************************Form************************************
    Private oForm As SAPbouiCOM.Form
    '**************************UserTable************************************
    Private PSSIT_WOR2, PSSIT_WOR3, PSSIT_WOR4 As SAPbobsCOM.UserTable
    '**************************ChooseFromList************************************
    Private oChRoutList As SAPbouiCOM.ChooseFromList
    Private Event ChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable)
    '**************************Items - Matrix************************************
    Private oRoutMatrix, oCostMatrix As SAPbouiCOM.Matrix
    Private oRColumns, oCColumns As SAPbouiCOM.Columns
    'Private oRColumn, oCColumn As SAPbouiCOM.Column
    '**************************Items - Route Matrix Colums************************************
    Private oRRowNoCol, oRCodeCol, oRNameCol, oRPrdSerCol, oRPrdNoCol, oRBsLnNoCol, oROprSeqCol, oRPrntIdCol, oROprCodeCol, oROprNameCol, oRRewrkCol, oRRoutCodCol, oRSeqBsLnIdCol, oRPrdQtyCol, oRPassQtyCol, oRRewrkQtyCol, oRSrpQtyCol, oRLbrCostCol, oRMCCostCol, oRTLCostCol, oRSubCtCostCol, oRSrpCostCol, oRWODocCol, oOthrQty1Col, oOthrQty2Col, oOthrCost1Col, oOthrCost2Col As SAPbouiCOM.Column
    '**************************Items - Cost Matrix Colums************************************
    Private oCRowNoCol, oCDocEntCol, oCCodeCol, oCPrdSerCol, oCPrdNoCol, oCFxdCostCol, oCUnitCostCol, oCAbsMthdCol, oCAcctCodeCol, oCAcctNameCol, oCTotCostCol, oOthrCostCol As SAPbouiCOM.Column
    '**************************UserDataSource************************************
    '*************************Route UserDataSource*******************************
    Private URRowNo, URCode, URRework, URName, URPrdSer, URPrdNo, URBsLnNo, UROprSeq, URPrntId, UROprCode, UROprName, URRewrk, URRoutCod, URSqBsLnId, URPrdQty, URPassQty, URewrkQty, URSrpQty, URLbrCost, URMCCost, URTLCost, URSubCtCst, URSrpCost, URWODoc, UROthrQty1, UROthrQty2, UROthrCst1, UROthrCst2 As SAPbouiCOM.UserDataSource
    '*************************Cost Header UserDataSource*******************************
    Private UCHCode, UProdSer, UProdNo, UCompCost, ULabCost, UMCCost, UToolCost, USCCost, UTotCost, UPrdSer, UPrdNo, UCOthrCst1, UCOthrCst2, UCOthrCst3, UCOthrCst4 As SAPbouiCOM.UserDataSource
    '*************************Cost UserDataSource*******************************
    Private UCLineid, UCDocEntry, UCCode, UCPrdSer, UCPrdNo, UCFxdCost, UCUnitCst, UCAbsMthd, UCActCod, UCActNam, UCTotCost, UOthrCost As SAPbouiCOM.UserDataSource
    '**************************Items************************************
    Private oItem As SAPbouiCOM.Item
    Private oRoutItem, oCostItem, oRMItem, oCMItem, oCLItem, oTLCItem, oTMCItem, oTTCItem, oTSCItem, oTCItem, oTCLItem, oTTLCItem, oTTMCItem, oTTTCItem, oTTSCItem, oTCTItem, oCHCodeItem, oPrdSerItem, oPrdNoItem, oLTCstItem1, oLTCstItem2, oLTCstItem3, oLTCstItem4, oTTCstItem1, oTTCstItem2, oTTCstItem3, oTTCstItem4 As SAPbouiCOM.Item
    '**************************Items - StaticText************************************
    Private oCodeLbl, oCompCostLbl, oLabCostLbl, oMCCostLbl, oToolCostLbl, oSCCostLbl, oTotCostLbl, oOthrCost1Lbl, oOthrCost2Lbl, oOthrCost3Lbl, oOthrCost4Lbl As SAPbouiCOM.StaticText
    '**************************Items - EditText************************************
    Private oProdNoTxt, oProdTxt, oCodeTxt, oCompCostTxt, oLabCostTxt, oMCCostTxt, oToolCostTxt, oSCCostTxt, oTotCostTxt, oCHCodeTxt, oPrdSerTxt, oPrdNoTxt, oOthrCost1Txt, oOthrCost2Txt, oOthrCost3Txt, oOthrCost4Txt As SAPbouiCOM.EditText
    '**************************Items - Combo************************************
    Private oProdSerCombo As SAPbouiCOM.ComboBox
    '**************************Items - Folder************************************
    Private oItemFolder, oRoutFolder, oCostFolder As SAPbouiCOM.Folder
    '**************************Items - Button************************************
    Private oRoutBtnItem, oAddBtnItem, oDelBtnItem, oPrntBtnItem As SAPbouiCOM.Item
    Private oRoutBtn, oAddBtn, oDelBtn, oPrntBtn As SAPbouiCOM.Button
    '***********************Other Variables****************************
    Private IntICount As Integer
    Private BoolResize As Boolean
    Private oRDSerialNo = 0, oCDHSerialNo = 0, oCDDSerialNo = 0
    Private oRRowNo As Integer = 0
    Private oRSerNo As Integer = 0
    Private oProductionOrderNo As String

    Private WithEvents ProcessSheetClass As ProcessSheetReport
    Private oBoolIssueComponents As Boolean = True

    Private TestForm As FrmTest
    Private oThread As System.Threading.Thread
#End Region
    ''' <summary>
    ''' SetApplication() and  LoadMenus() methods are called.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        MyBase.New()
        Dim expdate As Date
        expdate = Now()
        'MsgBox(expdate)
        ' If expdate <= "6/30/2012" Then
        LoadMenus()
        SBO_Application.SetStatusBarMessage("Engineering addon connected successfully", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        'Else
        ' MsgBox("No Sap license")
        'End If
    End Sub
   
    ''' <summary>
    ''' Creating the menus.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadMenus()
        If SBO_Application.Menus.Item("43520").SubMenus.Exists("Engineering") Then Return
        formCmdCenter = SBO_Application.Forms.GetFormByTypeAndCount(169, 1)
        formCmdCenter.Freeze(True)
        oGeneralMenu = CreateMenu(sPath & "\Resources\gear4.bmp", 12, "Engineering", SAPbouiCOM.BoMenuType.mt_POPUP, "Engineering", SBO_Application.Menus.Item("43520"))
        oMaster = CreateMenu(ImagePath, 1, "Master", SAPbouiCOM.BoMenuType.mt_POPUP, "Master", SBO_Application.Menus.Item("Engineering"))
        oTransaction = CreateMenu(ImagePath, 2, "Transaction", SAPbouiCOM.BoMenuType.mt_POPUP, "Transaction", SBO_Application.Menus.Item("Engineering"))
        oReports = CreateMenu(ImagePath, 3, "Reports", SAPbouiCOM.BoMenuType.mt_POPUP, "Reports", SBO_Application.Menus.Item("Engineering"))
        '*****************************Master Forms*************************************
        Call CreateMenu(ImagePath, 1, "Production Setup", SAPbouiCOM.BoMenuType.mt_STRING, "ProdSetup", oMaster)
        Call CreateMenu(ImagePath, 2, "Shift", SAPbouiCOM.BoMenuType.mt_STRING, "Shift", oMaster)
        Call CreateMenu(ImagePath, 3, "Work Centre", SAPbouiCOM.BoMenuType.mt_STRING, "WorkCentre", oMaster)
        Call CreateMenu(ImagePath, 4, "Machine Groups", SAPbouiCOM.BoMenuType.mt_STRING, "MachineGroups", oMaster)
        Call CreateMenu(ImagePath, 5, "Machine Master", SAPbouiCOM.BoMenuType.mt_STRING, "MachineMaster", oMaster)
        Call CreateMenu(ImagePath, 6, "Skill Groups", SAPbouiCOM.BoMenuType.mt_STRING, "SkillGroups", oMaster)
        Call CreateMenu(ImagePath, 7, "Labour", SAPbouiCOM.BoMenuType.mt_STRING, "Labour", oMaster)
        Call CreateMenu(ImagePath, 8, "Tools", SAPbouiCOM.BoMenuType.mt_STRING, "Tools", oMaster)
        ' Call CreateMenu(ImagePath, 9, "Stoppage", SAPbouiCOM.BoMenuType.mt_STRING, "Stoppage", oMaster)
        Call CreateMenu(ImagePath, 10, "Operations", SAPbouiCOM.BoMenuType.mt_STRING, "Operations", oMaster)
        Call CreateMenu(ImagePath, 11, "Operations Routing", SAPbouiCOM.BoMenuType.mt_STRING, "OprRouting", oMaster)
        '*****************************Transaction Forms*************************************
        Call CreateMenu(ImagePath, 1, "Production Entry", SAPbouiCOM.BoMenuType.mt_STRING, "ProductionEntry", oTransaction)
        Call CreateMenu(ImagePath, 2, "Stoppage Entry", SAPbouiCOM.BoMenuType.mt_STRING, "StoppageEntry", oTransaction)
        '*****************************Report Forms*************************************
        Call CreateMenu(ImagePath, 1, "Master Reports", SAPbouiCOM.BoMenuType.mt_STRING, "MasterReports", oReports)
        Call CreateMenu(ImagePath, 2, "Production Order Cost", SAPbouiCOM.BoMenuType.mt_STRING, "ProductionOrderCost", oReports)
        Call CreateMenu(ImagePath, 3, "Production Report", SAPbouiCOM.BoMenuType.mt_STRING, "ProductionReport", oReports)
        Call CreateMenu(ImagePath, 4, "Machine Utilization Report", SAPbouiCOM.BoMenuType.mt_STRING, "MachineUtilizationReport", oReports)
        Call CreateMenu(ImagePath, 5, "Machine Performance Report", SAPbouiCOM.BoMenuType.mt_STRING, "MachinePerformanceReport", oReports)
        Call CreateMenu(ImagePath, 6, "Labour Performance Report", SAPbouiCOM.BoMenuType.mt_STRING, "LabourPerformanceReport", oReports)
        SBO_Application.SetStatusBarMessage("Engineering Add On connected successfully.....", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        formCmdCenter.Freeze(False)
        formCmdCenter.Update()
    End Sub
    ''' <summary>
    ''' Creating the menus.
    ''' </summary>
    ''' <param name="ImagePath"></param>
    ''' <param name="Position"></param>
    ''' <param name="DisplayName"></param>
    ''' <param name="MenuType"></param>
    ''' <param name="UniqueID"></param>
    ''' <param name="ParentMenu"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CreateMenu(ByVal ImagePath As String, ByVal Position As Int32, ByVal DisplayName As String, ByVal MenuType As SAPbouiCOM.BoMenuType, ByVal UniqueID As String, ByVal ParentMenu As SAPbouiCOM.MenuItem) As SAPbouiCOM.MenuItem
        Try
            Dim oMenuPackage As SAPbouiCOM.MenuCreationParams
            oMenuPackage = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oMenuPackage.Image = ImagePath
            oMenuPackage.Position = Position
            oMenuPackage.Type = MenuType
            oMenuPackage.UniqueID = UniqueID
            oMenuPackage.String = DisplayName
            ParentMenu.SubMenus.AddEx(oMenuPackage)
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
        End Try
        Return ParentMenu.SubMenus.Item(UniqueID)
    End Function
    ''' <summary>
    ''' Loading the form when the menu is clicked.
    ''' </summary>
    ''' <param name="pVal"></param>
    ''' <param name="BubbleEvent"></param>
    ''' <remarks></remarks>
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Dim FormType As String
        FormType = SBO_Application.Forms.ActiveForm.Type
        Dim oExRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '******************Menus****************
        Try
            'If pVal.BeforeAction = True Then
            'If pVal.MenuUID = "ProductionEntry" Then
            '     oExRs.DoQuery("Select * From ORTT Where RateDate = '" & Date.Today & "'")
            ' If oExRs.RecordCount = 0 Then
            '        SBO_Application.SetStatusBarMessage("Exchange rate not updated [Message 131-6]", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            '      SBO_Application.ActivateMenuItem("3333")
            '       BubbleEvent = False
            ' End If
            'End If
            'End If
            If (pVal.MenuUID = "1281" Or pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") And FormType = "65211" And pVal.BeforeAction = True Then
                oForm = SBO_Application.Forms.ActiveForm()
                oForm.Freeze(True)
                oForm.Items.Item("35").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                oForm.Freeze(False)
            End If
            If pVal.BeforeAction = False Then
                If pVal.MenuUID = "Shift" Then
                    Shift = New Shift(SBO_Application, oCompany)
                ElseIf pVal.MenuUID = "WorkCentre" Then
                    WorkCentre = New WorkCentre(SBO_Application, oCompany)
                ElseIf pVal.MenuUID = "MachineGroups" Then
                    MachineGroups = New MachineGroups(SBO_Application, oCompany)
                ElseIf pVal.MenuUID = "MachineMaster" Then
                    MachineMaster = New MachineMaster(SBO_Application, oCompany)
                ElseIf pVal.MenuUID = "SkillGroups" Then
                    SkillGroups = New SkillGroups(SBO_Application, oCompany)
                ElseIf pVal.MenuUID = "Labour" Then
                    Labour = New Labour(SBO_Application, oCompany)
                ElseIf pVal.MenuUID = "Tools" Then
                    Tools = New Tools(SBO_Application, oCompany)
                ElseIf pVal.MenuUID = "Stoppage" Then
                    Stoppage = New Stoppage(SBO_Application, oCompany)
                ElseIf pVal.MenuUID = "Operations" Then
                    Operations = New Operations(SBO_Application, oCompany)
                ElseIf pVal.MenuUID = "OprRouting" Then
                    OperationsRouting = New OperationsRouting(SBO_Application, oCompany)
                ElseIf pVal.MenuUID = "ProdSetup" Then
                    ProductionSetup = New ProductionSetup(SBO_Application, oCompany)
                ElseIf pVal.MenuUID = "ProductionEntry" Then
                    ProductionEntry = New ProductionEntry(SBO_Application, oCompany)
                ElseIf pVal.MenuUID = "StoppageEntry" Then
                    MachineDownTime = New MachineDownTime(SBO_Application, oCompany)
                ElseIf pVal.MenuUID = "MasterReports" Then
                    MasterReports = New MasterReports(SBO_Application, oCompany)
                ElseIf pVal.MenuUID = "ProductionOrderCost" Then
                    ProdCostReport = New ProdCostReport(SBO_Application, oCompany)
                ElseIf pVal.MenuUID = "ProductionReport" Then
                    ProductionReport = New ProductionReport(SBO_Application, oCompany)
                ElseIf pVal.MenuUID = "MachineUtilizationReport" Then
                    MCUtilReport = New MCUtilReport(SBO_Application, oCompany)
                ElseIf pVal.MenuUID = "MachinePerformanceReport" Then
                    MCPerfReport = New MCPerfReport(SBO_Application, oCompany)
                ElseIf pVal.MenuUID = "LabourPerformanceReport" Then
                    LabrPerfReport = New LabrPerfReport(SBO_Application, oCompany)
                End If
                '*****************Production Order******************************
                If pVal.MenuUID = "1282" And FormType = "65211" Then
                    oItemFolder.Select()
                    ClearRoutMatrix()
                    ClearCostHeaderUDS()

                End If
               
                If (pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") And FormType = "65211" Then
                    oForm.Freeze(True)
                    SetItemEnabled()
                    ClearRoutMatrix()
                    LoadRoutMatrixData()
                    oForm.Items.Item("URoutFol").Enabled = True
                    ClearCostHeaderUDS()
                    LoadCostDataFromDB()
                    ClearCostMatrix()
                    LoadCostDetailsFromDB()
                    oForm.Items.Item("UCostFol").Enabled = True
                    '*************Add Print Process Sheet Button****************
                    Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Try
                        oRs.DoQuery("select * from OWOR where DocNum='" & oProdNoTxt.Value & "' and status ='R'")
                        If oRs.RecordCount > 0 Then
                            oPrntBtnItem.Visible = True
                            oRoutBtnItem.Enabled = False
                            'oForm.Items.Item("UAddBtn").Enabled = False
                            'oForm.Items.Item("UDelBtn").Enabled = False

                        Else
                            oPrntBtnItem.Visible = False
                            oRoutBtnItem.Enabled = True
                            'oForm.Items.Item("UAddBtn").Enabled = True
                            'oForm.Items.Item("UDelBtn").Enabled = True
                        End If
                    Catch ex As Exception
                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    End Try

                    oForm.Freeze(False)
                End If
            End If
            If pVal.MenuUID = "5923" And pVal.BeforeAction = True And FormType = "65211" Then
                Try
                    oBoolIssueComponents = False
                Catch ex As Exception
                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End Try
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
        End Try
    End Sub
    Private Sub SetItemEnabled()
        Try
            oForm.Items.Item("txtlrcost").Enabled = False
            oForm.Items.Item("txtccost").Enabled = False
            oForm.Items.Item("txtmcst").Enabled = False
            oForm.Items.Item("txttlcst").Enabled = False
            oForm.Items.Item("txtsucst").Enabled = False
            oForm.Items.Item("txtcst1").Enabled = False
            oForm.Items.Item("txtcst2").Enabled = False
            oForm.Items.Item("txtcst3").Enabled = False
            oForm.Items.Item("txtcst4").Enabled = False
            oForm.Items.Item("txttocst").Enabled = False
            oForm.Items.Item("txtcode").Enabled = False
            oForm.Items.Item("txtprdser").Enabled = False
            oForm.Items.Item("txtprdno").Enabled = False
            oForm.Items.Item("txtcode").Visible = False
            oForm.Items.Item("txtprdser").Visible = False
            oForm.Items.Item("txtprdno").Visible = False
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#Region "Production Order"
#Region "Generate Serial No"
    ''' <summary>
    ''' This function is used to generate the serial No from the table
    ''' </summary>
    ''' <param name="aTableName"></param>
    ''' <param name="aCriteriaSqlStr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GenerateSerialNo(ByVal aTableName As String, Optional ByVal aCriteriaSqlStr As String = "") As Integer
        Dim StrSql As String
        Dim oRs As SAPbobsCOM.Recordset
        Dim oCode As Integer
        Try

            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If aCriteriaSqlStr.Length = 0 Then
                StrSql = "Select IsNull(Max(Convert(Float,Code)),0) as Code From [@" & aTableName & "]"
            Else
                StrSql = aCriteriaSqlStr
            End If
            oRs.DoQuery(StrSql)
            If oRs.RecordCount > 0 Then
                oCode = oRs.Fields.Item("Code").Value + 1
            Else
                oCode = 1
            End If
            Return oCode
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region
#Region "Properties"
    ''' <summary>
    ''' This property is used to create the UserTables
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private ReadOnly Property UserTables() As SAPbobsCOM.UserTables
        Get
            Return oCompany.UserTables
        End Get
    End Property
#End Region
#Region "CFL"
#Region "CFL Creation"
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
            oChRoutList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_RTE", "RoutLst"))
            oRoutBtn.ChooseFromListUID = "RoutLst"
            'CreateNewConditions(oChRoutList, "U_Itemcode", SAPbouiCOM.BoConditionOperation.co_EQUAL, oProdTxt.Value)

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
    Private Sub Route_ChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable) Handles Me.ChooseFromList
        Dim oDataTable As SAPbouiCOM.DataTable = ChooseFromListSelectedObjects
        Dim oCurrentRow As Integer
        Dim oRoutID As String
        Dim oRs As SAPbobsCOM.Recordset
        Dim StrSql As String
        Dim IntICount As Integer
        Try
            oCurrentRow = CurrentRow
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If (ControlName = "6") And (ChoosefromListUID = "1") Then
                oForm.Items.Item("URoutBtn").Enabled = True
            End If
            If ControlName = "70" And ChoosefromListUID = "6" Then
                If Not oDataTable Is Nothing Then

                    IssProdNo = oDataTable.GetValue("DocNum", 0)
                    
                End If

            End If


            If (ControlName = "URoutBtn") And (ChoosefromListUID = "RoutLst") Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    If Not oDataTable Is Nothing Then
                        ClearRoutMatrix()
                        oRoutID = oDataTable.GetValue("Code", 0)
                        StrSql = "select b.LineId,b.U_Seqnce,b.U_Parid,b.U_Oprcode,b.U_Oprname,a.code,b.U_Adnl1,b.U_Adnl2 from [@PSSIT_ORTE] a,[@PSSIT_RTE4] b where a.code=b.code and a.Code='" & oRoutID & "'"
                        oRs.DoQuery(StrSql)
                        AddRouteRow(False)
                        If oRs.RecordCount > 0 Then
                            oRs.MoveFirst()
                            For IntICount = 0 To oRs.RecordCount - 1
                                ClearRoutMatrixUDS()
                                URCode.Value = oRDSerialNo
                                URName.Value = oRDSerialNo
                                UROprSeq.Value = oRs.Fields.Item("U_Seqnce").Value
                                URPrntId.Value = oRs.Fields.Item("U_Parid").Value
                                UROprCode.Value = oRs.Fields.Item("U_Oprcode").Value
                                UROprName.Value = oRs.Fields.Item("U_Oprname").Value
                                URRoutCod.Value = oRs.Fields.Item("code").Value
                                URSqBsLnId.Value = oRs.Fields.Item("LineId").Value
                                'UROthrQty1.Value = oRs.Fields.Item("U_Adnl1").Value
                                'UROthrQty2.Value = oRs.Fields.Item("U_Adnl2").Value
                                OperationCombo(oRoutID)
                                oRoutMatrix.AddRow(1, oRoutMatrix.RowCount)
                                If IntICount <> oRs.RecordCount - 1 Then
                                    AddRouteRow(True)
                                End If
                                oRs.MoveNext()
                            Next
                        End If
                        LoadCostDetailsData()
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
    Private Sub LoadCostDetailsData()
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oStrSql As String
        Try
            oStrSql = "Select T1.U_FCost,T1.U_UnitCost,T1.U_Absmthd,T1.U_Accode,T1.U_Acname " _
            & "from [@PSSIT_WCR1] T1 " _
            & "Inner Join [@PSSIT_OWCR] T0 On T1.Code = T0.Code "
            oRs.DoQuery(oStrSql)
            AddCostDetailsRow(False)
            If oRs.RecordCount > 0 Then
                oRs.MoveFirst()
                For IntICount = 0 To oRs.RecordCount - 1
                    ClearCostMatrixUDS()
                    UCLineid.Value = IntICount + 1
                    UCDocEntry.Value = UCHCode.Value
                    UCCode.Value = oCDDSerialNo
                    UCPrdSer.Value = oProdSerCombo.Selected.Value
                    UCPrdNo.Value = oProdNoTxt.Value
                    Dim STR As String
                    STR = oRs.Fields.Item("U_FCost").Value
                    UCFxdCost.Value = STR 'oRs.Fields.Item("U_FCost").Value


                    UCUnitCst.Value = oRs.Fields.Item("U_UnitCost").Value
                    UCAbsMthd.Value = oRs.Fields.Item("U_Absmthd").Value
                    UCActCod.Value = oRs.Fields.Item("U_Accode").Value
                    UCActNam.Value = oRs.Fields.Item("U_Acname").Value
                    UCTotCost.Value = 0
                    UOthrCost.Value = 0
                    oCostMatrix.AddRow(1, oCostMatrix.RowCount)
                    oCostMatrix.Columns.Item("colfcost").Cells.Item(oCostMatrix.RowCount).Specific.value = oRs.Fields.Item("U_FCost").Value
                    If IntICount <> oRs.RecordCount - 1 Then
                        AddCostDetailsRow(True)
                    End If
                    oRs.MoveNext()
                Next
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' This is used to Add one Row in the matrix
    ''' </summary>
    ''' <param name="oBoolStatus"></param>
    ''' <remarks></remarks>
    Private Sub AddRouteRow(ByVal oBoolStatus As Boolean)
        Try
            If oBoolStatus = False Then
                oRDSerialNo = GenerateSerialNo("PSSIT_WOR2")
            ElseIf oBoolStatus = True Then
                If oRoutMatrix.RowCount > 0 Then
                    oRDSerialNo = oRDSerialNo + 1
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This is used to Add one Row in the matrix
    ''' </summary>
    ''' <param name="oBoolStatus"></param>
    ''' <remarks></remarks>
    Private Sub AddCostDetailsRow(ByVal oBoolStatus As Boolean)
        Try
            If oBoolStatus = False Then
                oCDDSerialNo = GenerateSerialNo("PSSIT_WOR4")
            ElseIf oBoolStatus = True Then
                If oCostMatrix.RowCount > 0 Then
                    oCDDSerialNo = oCDDSerialNo + 1
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This condition is used to load the route details in the CFL
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetCFLConditions()
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        Dim oRs As SAPbobsCOM.Recordset

        Dim StrSql As String
        Dim i As Integer
        Try
            oCFLs = oForm.ChooseFromLists
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            For Each oCFL As SAPbouiCOM.ChooseFromList In oCFLs
                If (oCFL.UniqueID.Equals("RoutLst")) Then

                    StrSql = "select Code from [@PSSIT_ORTE] where U_Itemcode= '" & oProdTxt.Value & "'"

                    oRs.DoQuery(StrSql)
                    oCFL.SetConditions(Nothing)
                    '************** Adding Conditions to Item List ***************************
                    oCons = oCFL.GetConditions()
                    '************** Condition 1: ItemCode = oVenCodeTxt.Value *********
                    oCon = oCons.Add()
                    oCon.Alias = "Code"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL

                    For i = 1 To oRs.RecordCount
                        If oRs.EoF = False Then
                            'MsgBox(oRs.Fields.Item("Code").Value)
                            oCon.CondVal = oRs.Fields.Item("Code").Value
                            If Not i = oRs.RecordCount Then
                                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                                oCon = oCons.Add()
                                oCon.Alias = "Code"
                                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                            End If
                        End If
                        oRs.MoveNext()
                    Next
                    oCFL.SetConditions(oCons)
                End If
            Next
        Catch ex As Exception
        Finally
            oCon = Nothing
            oCons = Nothing
            oCFLs = Nothing
        End Try
    End Sub

#Region "Validate duplicate Route"
    Private Function ValidateDuplicateRoute() As Boolean
        Dim strOperation, strsequence, strparentid, stroperationame, stroprSeq, strroute As String
        Dim strOperation1, strsequence1, strparentid1, stroperationame1, stroprSeq1, strroute1 As String
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim oCheckbox As SAPbouiCOM.CheckBox
        Dim strRework, strRework1 As String
        For intRow As Integer = 1 To oRoutMatrix.RowCount
            strOperation = oROprSeqCol.Cells.Item(intRow).Specific.value
            strroute = oRRoutCodCol.Cells.Item(intRow).Specific.value
            'stroprSeq = oROprSeqCol.Cells.Item(intRow).Specific.value
            'strsequence = oROprCodeCol.Cells.Item(intRow).Specific.value
            strparentid = oRoutMatrix.Columns.Item("colparid").Cells.Item(intRow).Specific.value
            oCombo = oROprNameCol.Cells.Item(intRow).Specific
            stroperationame = oROprNameCol.Cells.Item(intRow).Specific.value 'oCombo.Selected.Value
            '   strRework = oRRewrkCol.Cells.Item(intRow).Specific.value
            oCheckbox = oRRewrkCol.Cells.Item(intRow).Specific
            If oCheckbox.Checked = True Then
                strRework = "Y"
            Else
                strRework = "N"
            End If

            For intLoop As Integer = intRow + 1 To oRoutMatrix.RowCount
                strOperation1 = oROprSeqCol.Cells.Item(intLoop).Specific.value
                strroute1 = oRRoutCodCol.Cells.Item(intLoop).Specific.value
                stroprSeq1 = oROprSeqCol.Cells.Item(intLoop).Specific.value
                strsequence1 = oROprCodeCol.Cells.Item(intLoop).Specific.value
                strparentid1 = oRoutMatrix.Columns.Item("colparid").Cells.Item(intLoop).Specific.value
                oCheckbox = oRRewrkCol.Cells.Item(intLoop).Specific
                If oCheckbox.Checked = True Then
                    strRework1 = "Y"
                Else
                    strRework1 = "N"
                End If

                oCombo = oROprNameCol.Cells.Item(intLoop).Specific
                stroperationame1 = oROprNameCol.Cells.Item(intLoop).Specific.value
                '   If strOperation = strOperation1 And stroprSeq = stroprSeq1 And strsequence = strsequence1 And strparentid = strparentid1 And stroperationame = stroperationame1 Then
                If strroute = strroute1 And stroperationame = stroperationame1 And strRework = strRework1 Then
                    SBO_Application.SetStatusBarMessage("Duplicate route details not allowed", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return False
                End If
            Next

        Next
        ' Return False
        Return True
    End Function
#End Region
#End Region
    ''' <summary>
    ''' Item Event
    ''' </summary>
    ''' <param name="FormUID"></param>
    ''' <param name="pVal"></param>
    ''' <param name="BubbleEvent"></param>
    ''' <remarks></remarks>
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Dim ChooseFromListEvent As SAPbouiCOM.ChooseFromListEvent
        Try
            If pVal.FormType = 65213 And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST And pVal.BeforeAction = False Then
                    ChooseFromListEvent = pVal
                    RaiseEvent ChooseFromList(pVal.ItemUID, pVal.ColUID, pVal.Row, ChooseFromListEvent.ChooseFromListUID, ChooseFromListEvent.SelectedObjects)
                End If
                If (pVal.BeforeAction = True) Then
                    If ((pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD)) Then
                        AddCostHDRUserDataSources()
                    End If
                End If
               
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False And pVal.ItemUID = "1" Then
                    Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim oStrSql As String
                    Try
                        'oStrSql = "Update [@PSSIT_WOR3] Set U_Totcmpcst = " _
                        '& "(Select Sum(T0.IssuedQty * T2.AvgPrice) as TotCompCost from WOR1 T0 " _
                        '& "Inner Join OWOR T1 On T1.DocEntry = T0.DocEntry " _
                        '& "Inner Join OITW T2 On T2.ItemCode = T0.ItemCode and T2.WhsCode = T0.warehouse " _
                        '& "Where T1.DocNum = " & oProdNoTxt.Value & "),U_TotCst =  (Select Sum(T0.IssuedQty * T2.AvgPrice) as TotCompCost from WOR1 T0 " _
                        '& "Inner Join OWOR T1 On T1.DocEntry = T0.DocEntry " _
                        '& "Inner Join OITW T2 On T2.ItemCode = T0.ItemCode and T2.WhsCode = T0.warehouse " _
                        '& "Where T1.DocNum = " & oProdNoTxt.Value & ") from [@PSSIT_WOR3] Where U_Pordno = " & oProdNoTxt.Value
                        oStrSql = "Update [@PSSIT_WOR3] Set U_Totcmpcst = " _
                        & "(Select Sum(T0.IssuedQty * T2.AvgPrice) as TotCompCost from WOR1 T0 " _
                        & "Inner Join OWOR T1 On T1.DocEntry = T0.DocEntry " _
                        & "Inner Join OITW T2 On T2.ItemCode = T0.ItemCode and T2.WhsCode = T0.warehouse " _
                        & "Where T1.DocNum = " & IssProdNo & "),U_TotCst =  (Select Sum(T0.IssuedQty * T2.AvgPrice) as TotCompCost from WOR1 T0 " _
                        & "Inner Join OWOR T1 On T1.DocEntry = T0.DocEntry " _
                        & "Inner Join OITW T2 On T2.ItemCode = T0.ItemCode and T2.WhsCode = T0.warehouse " _
                        & "Where T1.DocNum = " & IssProdNo & ") from [@PSSIT_WOR3] Where U_Pordno = " & IssProdNo
                        oRs.DoQuery(oStrSql)

                        oStrSql = ""
                        oRs = Nothing
                        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRs.DoQuery("Select * from [@PSSIT_WOR3] where U_Pordno = " & IssProdNo)
                        If oRs.RecordCount > 0 Then
                            oRs.MoveFirst()
                            UCompCost.Value = oRs.Fields.Item("U_Totcmpcst").Value
                            UTotCost.Value = oRs.Fields.Item("U_Totcst").Value
                        End If
                        oBoolIssueComponents = True
                    Catch ex As Exception
                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    End Try
                End If
            End If
            If ((pVal.FormType = 65211 And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)) Then

                oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                '**********ChooseFromList Event is called using the raiseevent*********
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST And pVal.BeforeAction = False Then
                    ChooseFromListEvent = pVal
                    RaiseEvent ChooseFromList(pVal.ItemUID, pVal.ColUID, pVal.Row, ChooseFromListEvent.ChooseFromListUID, ChooseFromListEvent.SelectedObjects)
                End If
                If (pVal.BeforeAction = True) Then
                    If ((pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD) And (pVal.BeforeAction = True)) Then
                        '*****************UserTables**************************
                        PSSIT_WOR2 = UserTables.Item("PSSIT_WOR2")
                        PSSIT_WOR3 = UserTables.Item("PSSIT_WOR3")
                        PSSIT_WOR4 = UserTables.Item("PSSIT_WOR4")
                        '***************Generate Serial No********************
                        oCDHSerialNo = GenerateSerialNo("PSSIT_WOR3")
                        oCDDSerialNo = GenerateSerialNo("PSSIT_WOR4")

                        oForm.Freeze(True)
                        AddFolders()
                        AddRoutUserDataSources()
                        AddCostHDRUserDataSources()
                        AddCostDTLUserDataSources()
                        AddRoutMatrix()
                        AddButton()
                        AddCostHdr()
                        AddCostMatrix()
                        LoadLookups()
                        oForm.Items.Item("URoutBtn").Enabled = False
                        oPrntBtnItem.Visible = False
                        oForm.Freeze(False)
                        oForm.Freeze(True)
                        LoadRoutMatrixData()
                        If oRoutMatrix.RowCount > 0 Then
                            oForm.Items.Item("URoutFol").Enabled = True
                        End If
                        oForm.Items.Item("UCostFol").Enabled = True
                        LoadCostDataFromDB()
                        LoadCostDetailsFromDB()
                        oForm.Freeze(False)
                    End If
                    '*****************Item Pressed*****************
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        If pVal.BeforeAction = True Then
                            If pVal.ItemUID = "1" Then
                                Try
                                     If oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        oForm.Freeze(True)
                                        LoadRoutMatrixData()
                                        LoadCostDataFromDB()
                                        LoadCostDetailsFromDB()
                                        SetItemEnabled()
                                        oForm.Freeze(False)
                                    End If
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        '****************Adding the child data to the database table***********
                                        Dim oRTransaction As Boolean
                                        Try
                                            If ValidateDuplicateRoute() = False Then
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                            RoutDelEmptyRow()
                                            Try
                                                If Not oCompany.InTransaction Then
                                                    oCompany.StartTransaction()
                                                End If
                                                If oRoutMatrix.RowCount > 0 Then
                                                    oRTransaction = True
                                                    AddRoutChildTable()
                                                End If
                                                AddCostHeaderTable()
                                                If oCostMatrix.RowCount > 0 Then
                                                    AddCostDetailsTable()
                                                End If
                                                If oRTransaction = True Then
                                                    '  If oCompany.InTransaction() Then
                                                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                                    'End If

                                                End If
                                            Catch ex As Exception
                                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                BubbleEvent = False
                                            Finally
                                                If oRTransaction = False Then
                                                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                    BubbleEvent = False
                                                    'ElseIf oCHTransaction = False Then
                                                    '    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                    '    BubbleEvent = False
                                                    'ElseIf oRTransaction = False And oCHTransaction = False Then
                                                    '    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                    '    BubbleEvent = False
                                                End If
                                            End Try
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
                            '**************Setting PaneLevel To the Folders********************
                            If pVal.ItemUID = "URoutFol" And pVal.BeforeAction = True Then

                                oForm.PaneLevel = 3
                                If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                    ' LoadRoutMatrixData()
                                End If
                            End If
                            If pVal.ItemUID = "UCostFol" And pVal.BeforeAction = True Then
                                oForm.PaneLevel = 4
                                SetItemEnabled()
                            End If
                            '******************LoadRoute Button Pres***************************
                            If pVal.ItemUID = "URoutBtn" And pVal.BeforeAction = True Then
                                Try
                                    If Len(oProdTxt.Value) > 0 Then
                                        Dim oCombo As SAPbouiCOM.ComboBox
                                        oCombo = oForm.Items.Item("10").Specific
                                        If oCombo.Selected.Value = "R" Then
                                            ' SBO_Application.SetStatusBarMessage("Can not delete the Route Details in Release Mode", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                        SetCFLConditions()
                                    End If
                                Catch ex As Exception
                                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                End Try
                            End If
                            '******************Process Sheet Button Press***************************
                            If pVal.ItemUID = "UPrntBtn" And pVal.BeforeAction = True Then
                                TestForm = New FrmTest(SBO_Application, oCompany, "ProcessSheetReport", oProdNoTxt.Value)
                                oThread = New Threading.Thread(AddressOf TestForm.StartThread)
                                oThread.Start()
                            End If
                            '*****************Delete Row Button Press************************
                            If pVal.ItemUID = "UDelBtn" And pVal.BeforeAction = True Then
                                If oRoutMatrix.RowCount > 0 Then
                                    Dim oCombo As SAPbouiCOM.ComboBox
                                    oCombo = oForm.Items.Item("10").Specific
                                    If oCombo.Selected.Value = "R" Then
                                        SBO_Application.SetStatusBarMessage("Can not delete the Route Details in Release Mode", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        Exit Sub
                                    End If

                                    If PSSIT_WOR2.GetByKey(URCode.Value) = True Then
                                        Dim I As Integer = PSSIT_WOR2.Remove()
                                        Dim oTemprs As SAPbobsCOM.Recordset
                                        oTemprs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oTemprs.DoQuery("Update [@PSSIT_WOR2] set Name=Name +'N' where code='" & URCode.Value & "'")
                                        oRoutMatrix.DeleteRow(oRoutMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                                        'oRoutMatrix.FlushToDataSource()
                                    ElseIf PSSIT_WOR2.GetByKey(URCode.Value) = False Then
                                        oRoutMatrix.DeleteRow(oRoutMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                                    End If
                                    BubbleEvent = False
                                End If
                            End If
                        End If
                    End If
                End If
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.BeforeAction = False And pVal.ItemUID = "22" Then
                    Try
                        If Not oProdSerCombo Is Nothing Then
                            URPrdSer.Value = oProdSerCombo.Selected.Value
                            UProdSer.Value = oProdSerCombo.Selected.Value
                        End If
                    Catch ex As Exception
                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    End Try
                ElseIf pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.Before_Action = True And pVal.ItemUID = "matrout" And pVal.ColUID = "colopnam" Then
                    'OperationCombo()
                End If
                '*****************Add Row Button Press************************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then
                    If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        '*************Add Print Process Sheet Button****************
                        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Try
                            oRs.DoQuery("select * from OWOR where DocNum='" & oProdNoTxt.Value & "' and status ='R'")
                            If oRs.RecordCount > 0 Then
                                oPrntBtnItem.Visible = True
                            Else
                                oPrntBtnItem.Visible = False
                            End If
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        End Try
                    End If
                    If pVal.ItemUID = "UAddBtn" Then
                        Try
                            If oRoutMatrix.RowCount > 0 Then
                                'Dim oCombo As SAPbouiCOM.ComboBox
                                'oCombo = oForm.Items.Item("10").Specific
                                'If oCombo.Selected.Value = "R" Then
                                '    SBO_Application.SetStatusBarMessage("Can not Add New Route Details in Release Mode", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                '    Exit Sub
                                'End If
                                oRoutMatrix.GetLineData(oRoutMatrix.RowCount)
                                AddRouteRow(True)
                                ClearRoutMatrixUDS()
                                oRoutMatrix.AddRow(1, oRoutMatrix.RowCount)
                                'OperationCombo()
                                URCode.Value = oRDSerialNo
                                URName.Value = oRDSerialNo
                                URPrdSer.Value = oProdSerCombo.Selected.Value
                                URPrdNo.Value = oProdNoTxt.Value
                                oRoutMatrix.SetLineData(oRoutMatrix.RowCount)
                            End If
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        End Try
                    End If
                    If pVal.ItemUID = "35" Then
                        oForm.PaneLevel = 1
                    End If
                    If pVal.ItemUID = "36" Then
                        oForm.PaneLevel = 2
                    End If
                    If pVal.ItemUID = "URoutFol" Then
                        oForm.PaneLevel = 3
                    End If
                    If pVal.ItemUID = "UCostFol" Then
                        oForm.PaneLevel = 4
                    End If
                End If
                '*********************Operation Name Combo Select************************
                If pVal.ItemUID = "matrout" And pVal.ColUID = "colopnam" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.BeforeAction = False Then
                    Dim oOperNamCombo As SAPbouiCOM.ComboBox
                    Dim oRoutEdit As SAPbouiCOM.EditText
                    Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim ORtID As String
                    Try
                        oOperNamCombo = oROprNameCol.Cells.Item(pVal.Row).Specific
                        oRoutMatrix.GetLineData(oRoutMatrix.RowCount)
                        UROprCode.Value = oOperNamCombo.Selected.Value
                        oRoutMatrix.SetLineData(oRoutMatrix.RowCount)
                        '***************Setting the Route ID to the row Added***************
                        If pVal.Row > 1 Then
                            oRoutEdit = oRRoutCodCol.Cells.Item(pVal.Row - 1).Specific
                            ORtID = oRoutEdit.Value
                            oRoutMatrix.GetLineData(oRoutMatrix.RowCount)
                            URRoutCod.Value = ORtID
                            oRoutMatrix.SetLineData(oRoutMatrix.RowCount)

                        End If
                        'oRoutEdit = oRRoutCodCol.Cells.Item(pVal.Row - 1).Specific
                        'ORtID = oRoutEdit.Value
                        'StrSql = "select distinct b.Code,b.LineId,b.U_Seqnce,b.U_Parid from [@PSSIT_ORTE] a,[@PSSIT_RTE4] b where(a.code = b.code And U_Oprname Is Not Null) and a.U_Itemcode='" & oProdTxt.Value & "' and b.U_Oprname='" & oOperNamCombo.Selected.Value & "' and  b.code='" & ORtID & "' group by b.Code,b.LineId,b.U_Seqnce,b.U_Parid"
                        'oRs.DoQuery(StrSql)
                        'If oRs.RecordCount > 0 Then
                        '    oRs.MoveFirst()
                        '    oRoutMatrix.GetLineData(oRoutMatrix.RowCount)
                        '    URRoutCod.Value = oRs.Fields.Item("Code").Value
                        '    UROprSeq.Value = oRs.Fields.Item("U_Seqnce").Value
                        '    URPrntId.Value = oRs.Fields.Item("U_Parid").Value
                        '    URSqBsLnId.Value = oRs.Fields.Item("LineId").Value
                        '    oRoutMatrix.SetLineData(oRoutMatrix.RowCount)
                        'End If
                    Catch ex As Exception
                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    End Try
                End If
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    ''' <summary>
    ''' Initializing UserDataSources By setting UniqueID,DataType of the field and Length of the Field.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AddRoutUserDataSources()
        Try
            URRowNo = oForm.DataSources.UserDataSources.Add("URRowNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            URCode = oForm.DataSources.UserDataSources.Add("URCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
            ' URRewrk = oForm.DataSources.UserDataSources.Add("URCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
            URName = oForm.DataSources.UserDataSources.Add("URName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
            URPrdSer = oForm.DataSources.UserDataSources.Add("URPrdSer", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6)
            URPrdNo = oForm.DataSources.UserDataSources.Add("URPrdNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50)
            URBsLnNo = oForm.DataSources.UserDataSources.Add("URBsLnNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            UROprSeq = oForm.DataSources.UserDataSources.Add("UROprSeq", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            URPrntId = oForm.DataSources.UserDataSources.Add("URPrntId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            UROprCode = oForm.DataSources.UserDataSources.Add("UROprCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
            UROprName = oForm.DataSources.UserDataSources.Add("UROprName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            URRewrk = oForm.DataSources.UserDataSources.Add("URRewrk", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            URRoutCod = oForm.DataSources.UserDataSources.Add("URRoutCod", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
            URSqBsLnId = oForm.DataSources.UserDataSources.Add("URSqBsLnId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 3)
            URPrdQty = oForm.DataSources.UserDataSources.Add("URPrdQty", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            URPassQty = oForm.DataSources.UserDataSources.Add("URPassQty", SAPbouiCOM.BoDataType.dt_QUANTITY, 10)
            URewrkQty = oForm.DataSources.UserDataSources.Add("URewrkQty", SAPbouiCOM.BoDataType.dt_QUANTITY, 10)
            URSrpQty = oForm.DataSources.UserDataSources.Add("URSrpQty", SAPbouiCOM.BoDataType.dt_QUANTITY, 10)
            URLbrCost = oForm.DataSources.UserDataSources.Add("URLbrCost", SAPbouiCOM.BoDataType.dt_PRICE, 10)
            URMCCost = oForm.DataSources.UserDataSources.Add("URMCCost", SAPbouiCOM.BoDataType.dt_PRICE, 10)
            URTLCost = oForm.DataSources.UserDataSources.Add("URTLCost", SAPbouiCOM.BoDataType.dt_PRICE, 10)
            URSubCtCst = oForm.DataSources.UserDataSources.Add("URSubCtCst", SAPbouiCOM.BoDataType.dt_PRICE, 10)
            URSrpCost = oForm.DataSources.UserDataSources.Add("URSrpCost", SAPbouiCOM.BoDataType.dt_PRICE, 10)
            URWODoc = oForm.DataSources.UserDataSources.Add("URWODoc", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 40)
            UROthrQty1 = oForm.DataSources.UserDataSources.Add("UROthrQty1", SAPbouiCOM.BoDataType.dt_QUANTITY, 10)
            UROthrQty2 = oForm.DataSources.UserDataSources.Add("UROthrQty2", SAPbouiCOM.BoDataType.dt_QUANTITY, 10)
            UROthrCst1 = oForm.DataSources.UserDataSources.Add("UROthrCst1", SAPbouiCOM.BoDataType.dt_PRICE, 10)
            UROthrCst2 = oForm.DataSources.UserDataSources.Add("UROthrCst2", SAPbouiCOM.BoDataType.dt_PRICE, 10)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Initializing UserDataSources By setting UniqueID,DataType of the field and Length of the Field.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AddCostHDRUserDataSources()
        Try
            UCHCode = oForm.DataSources.UserDataSources.Add("UCHCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
            UProdSer = oForm.DataSources.UserDataSources.Add("UProdSer", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 6)
            UProdNo = oForm.DataSources.UserDataSources.Add("UProdNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50)
            UCompCost = oForm.DataSources.UserDataSources.Add("UCompCost", SAPbouiCOM.BoDataType.dt_PRICE, 19.6)
            ULabCost = oForm.DataSources.UserDataSources.Add("ULabCost", SAPbouiCOM.BoDataType.dt_PRICE, 19.6)
            UMCCost = oForm.DataSources.UserDataSources.Add("UMCCost", SAPbouiCOM.BoDataType.dt_PRICE, 19.6)
            UToolCost = oForm.DataSources.UserDataSources.Add("UToolCost", SAPbouiCOM.BoDataType.dt_PRICE, 19.6)
            USCCost = oForm.DataSources.UserDataSources.Add("USCCost", SAPbouiCOM.BoDataType.dt_PRICE, 19.6)
            UTotCost = oForm.DataSources.UserDataSources.Add("UTotCost", SAPbouiCOM.BoDataType.dt_PRICE, 19.6)
            UCOthrCst1 = oForm.DataSources.UserDataSources.Add("UCOthrCst1", SAPbouiCOM.BoDataType.dt_PRICE, 19.6)
            UCOthrCst2 = oForm.DataSources.UserDataSources.Add("UCOthrCst2", SAPbouiCOM.BoDataType.dt_PRICE, 19.6)
            UCOthrCst3 = oForm.DataSources.UserDataSources.Add("UCOthrCst3", SAPbouiCOM.BoDataType.dt_PRICE, 19.6)
            UCOthrCst4 = oForm.DataSources.UserDataSources.Add("UCOthrCst4", SAPbouiCOM.BoDataType.dt_PRICE, 19.6)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Initializing UserDataSources By setting UniqueID,DataType of the field and Length of the Field.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AddCostDTLUserDataSources()
        Try
            UCDocEntry = oForm.DataSources.UserDataSources.Add("UCDocEnt", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 10)
            UCLineid = oForm.DataSources.UserDataSources.Add("UCLineid", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 10)
            UCCode = oForm.DataSources.UserDataSources.Add("UCCode", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 10)
            UCPrdSer = oForm.DataSources.UserDataSources.Add("UCPrdSer", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 6)
            UCPrdNo = oForm.DataSources.UserDataSources.Add("UCPrdNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50)
            UCFxdCost = oForm.DataSources.UserDataSources.Add("UCFxdCost", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50)
            UCUnitCst = oForm.DataSources.UserDataSources.Add("UCUnitCst", SAPbouiCOM.BoDataType.dt_PRICE, 19.6)
            UCAbsMthd = oForm.DataSources.UserDataSources.Add("UCAbsMthd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 60)
            UCActCod = oForm.DataSources.UserDataSources.Add("UCActCod", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
            UCActNam = oForm.DataSources.UserDataSources.Add("UCActNam", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            UCTotCost = oForm.DataSources.UserDataSources.Add("UCTotCost", SAPbouiCOM.BoDataType.dt_PRICE, 19.6)
            UOthrCost = oForm.DataSources.UserDataSources.Add("UOthrCost", SAPbouiCOM.BoDataType.dt_PRICE, 19.6)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' ClearMatrixDataSources() function is Called to clear the UserDataSources Value.
    ''' Matrix is Cleared and Flushing the data to the UserDatasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClearRoutMatrix()
        Try
            ClearRoutMatrixUDS()
            oRoutMatrix.Clear()
            oRoutMatrix.FlushToDataSource()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Clearing the Values in the Rout User Data Sources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClearRoutMatrixUDS()
        Try
            URRowNo.Value = oRoutMatrix.RowCount + 1
            URCode.Value = ""
            URName.Value = ""
            URPrdSer.Value = oProdSerCombo.Selected.Value
            URPrdNo.Value = oProdNoTxt.Value
            URBsLnNo.Value = ""
            UROprSeq.Value = ""
            URPrntId.Value = ""
            UROprCode.Value = ""
            UROprName.Value = ""
            URRewrk.Value = ""
            URRoutCod.Value = ""
            URSqBsLnId.Value = ""
            URPrdQty.Value = ""
            URPassQty.Value = ""
            URewrkQty.Value = ""
            URSrpQty.Value = ""
            URLbrCost.Value = ""
            URMCCost.Value = ""
            URTLCost.Value = ""
            URSubCtCst.Value = ""
            URSrpCost.Value = ""
            URWODoc.Value = ""
            UROthrQty1.Value = ""
            UROthrQty2.Value = ""
            UROthrCst1.Value = ""
            UROthrCst2.Value = ""
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This is used to delete the empty rows in the Route Matrix
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub RoutDelEmptyRow()
        Dim IntICount As Integer
        Try
            For IntICount = 1 To oRoutMatrix.VisualRowCount
                oRoutMatrix.GetLineData(IntICount)
                If URRoutCod.Value.Length = 0 Then
                    oRoutMatrix.DeleteRow(IntICount)
                    oRoutMatrix.FlushToDataSource()
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' ClearMatrixDataSources() function is Called to clear the UserDataSources Value.
    ''' Matrix is Cleared and Flushing the data to the UserDatasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClearCostMatrix()
        Try
            ClearCostMatrixUDS()
            oCostMatrix.Clear()
            oCostMatrix.FlushToDataSource()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Clearing the Values in the Rout User Data Sources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClearCostMatrixUDS()
        Try
            UCLineid.Value = "0"
            UCDocEntry.Value = UCHCode.Value
            UCCode.Value = oCDDSerialNo
            UCPrdSer.Value = oProdSerCombo.Selected.Value
            UCPrdNo.Value = oProdNoTxt.Value
            UCFxdCost.Value = ""
            UCUnitCst.Value = ""
            UCAbsMthd.Value = ""
            UCActCod.Value = ""
            UCActNam.Value = ""
            UCTotCost.Value = ""
            UOthrCost.Value = ""
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Clearing the Values in the Rout User Data Sources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClearCostHeaderUDS()
        Try
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                UCHCode.Value = GenerateSerialNo("PSSIT_WOR3")
            Else
                UCHCode.Value = ""
            End If
            UProdSer.Value = oProdSerCombo.Selected.Value
            UProdNo.Value = oProdNoTxt.Value
            UCompCost.Value = "0.00"
            ULabCost.Value = "0.00"
            UMCCost.Value = "0.00"
            USCCost.Value = "0.00"
            UToolCost.Value = "0.00"
            UTotCost.Value = "0.00"
            UCOthrCst1.Value = "0.00"
            UCOthrCst3.Value = "0.00"
            UCOthrCst4.Value = "0.00"
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    
    ''' <summary>
    ''' Configuring the Folder items/controls
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AddFolders()
        Try
            '***************Route Folder*************
            oRoutItem = oForm.Items.Add("URoutFol", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem = oForm.Items.Item("36")
            oItemFolder = oForm.Items.Item("35").Specific
            oRoutItem.Top = oItem.Top
            oRoutItem.Height = oItem.Height
            oRoutItem.Width = oItem.Width
            oRoutItem.Left = oItem.Left + oItem.Width
            oRoutFolder = oRoutItem.Specific
            oForm.Items.Item("URoutFol").AffectsFormMode = False
            oRoutFolder.Caption = "Routes Details"
            oRoutFolder.GroupWith("36")
            oForm.PaneLevel = 1
            '***************Cost Summary Folder*************
            oCostItem = oForm.Items.Add("UCostFol", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem = oForm.Items.Item("URoutFol")
            oCostItem.Top = oItem.Top
            oCostItem.Height = oItem.Height
            oCostItem.Width = oItem.Width
            oCostItem.Left = oItem.Left + oItem.Width
            oCostFolder = oCostItem.Specific
            oForm.Items.Item("UCostFol").AffectsFormMode = False
            oCostFolder.Caption = "Cost Details"
            oCostFolder.GroupWith("URoutFol")
            oForm.PaneLevel = 1
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub AddButton()
        Dim oTItem As SAPbouiCOM.Item
        Dim oLItem As SAPbouiCOM.Item
        Dim oRItem As SAPbouiCOM.Item
        Dim oBItem As SAPbouiCOM.Item
        Try
            oTItem = oForm.Items.Item("83")
            oLItem = oForm.Items.Item("57")
            oRItem = oForm.Items.Item("55")
            oBItem = oForm.Items.Item("54")
            '**************Load Route Button****************
            oRoutBtnItem = oForm.Items.Add("URoutBtn", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem = oForm.Items.Item("2")
            oRoutBtnItem.Top = oItem.Top
            oRoutBtnItem.Height = oItem.Height
            oRoutBtnItem.Width = oItem.Width
            oRoutBtnItem.Left = oItem.Left + oItem.Width + 3
            oRoutBtn = oRoutBtnItem.Specific
            oRoutBtn.Caption = "Load Route"
            '**************Add Row Button****************
            oAddBtnItem = oForm.Items.Add("UAddBtn", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oAddBtnItem.Top = oTItem.Top + 180
            oAddBtnItem.Height = 20
            oAddBtnItem.Width = 65
            oAddBtnItem.Left = oLItem.Left + 5
            oAddBtnItem.FromPane = 3
            oAddBtnItem.ToPane = 3
            oAddBtn = oAddBtnItem.Specific
            oAddBtn.Caption = "Add Row"
            '**************Delete Row Button****************
            oDelBtnItem = oForm.Items.Add("UDelBtn", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oDelBtnItem.Top = oTItem.Top + 180
            oDelBtnItem.Height = 20
            oDelBtnItem.Width = 65
            oDelBtnItem.Left = oAddBtnItem.Left + oAddBtnItem.Width + 3
            oDelBtnItem.FromPane = 3
            oDelBtnItem.ToPane = 3
            oDelBtn = oDelBtnItem.Specific
            oDelBtn.Caption = "Delete Row"
            '**************PrintButton****************
            oPrntBtnItem = oForm.Items.Add("UPrntBtn", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oPrntBtnItem.Top = oRoutBtnItem.Top
            oPrntBtnItem.Height = oRoutBtnItem.Height
            oPrntBtnItem.Width = 110
            oPrntBtnItem.Left = oRoutBtnItem.Left + oRoutBtnItem.Width + 3
            oPrntBtn = oPrntBtnItem.Specific
            oPrntBtn.Caption = "Print Process Sheet"
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    
    ''' <summary>
    ''' Configuring the Matrix items/controls
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AddRoutMatrix()
        Try
            '// we will use the following object to set a linked button
            Dim oTItem As SAPbouiCOM.Item
            Dim oLItem As SAPbouiCOM.Item
            Dim oRItem As SAPbouiCOM.Item
            Dim oBItem As SAPbouiCOM.Item

            oRMItem = oForm.Items.Add("matrout", SAPbouiCOM.BoFormItemTypes.it_MATRIX)
            oTItem = oForm.Items.Item("83")
            oLItem = oForm.Items.Item("57")
            oRItem = oForm.Items.Item("55")
            oBItem = oForm.Items.Item("54")

            oRMItem.Top = oTItem.Top + 5
            oRMItem.Height = 170
            oRMItem.Width = 567
            oRMItem.Left = oLItem.Left + 5
            oRMItem.FromPane = 3
            oRMItem.ToPane = 3


            oRoutMatrix = oRMItem.Specific
            oRoutMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oRColumns = oRoutMatrix.Columns

            '// Adding Culomn items to the matrix

            oRRowNoCol = oRColumns.Add("#", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oRRowNoCol.TitleObject.Caption = "#"
            oRRowNoCol.Width = 20
            oRRowNoCol.Editable = False
            oRRowNoCol.DataBind.SetBound(True, "", "URRowNo")

            oRCodeCol = oRColumns.Add("colcode", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oRCodeCol.TitleObject.Caption = "Code"
            oRCodeCol.Width = 40
            oRCodeCol.Visible = False
            oRCodeCol.DataBind.SetBound(True, "", "URCode")

            oRNameCol = oRColumns.Add("colname", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oRNameCol.TitleObject.Caption = "Name"
            oRNameCol.Width = 100
            oRNameCol.Visible = False
            oRNameCol.DataBind.SetBound(True, "", "URName")

            oRPrdSerCol = oRColumns.Add("colpordser", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oRPrdSerCol.TitleObject.Caption = "Production Series"
            oRPrdSerCol.Width = 150
            oRPrdSerCol.Visible = False
            oRPrdSerCol.DataBind.SetBound(True, "", "URPrdSer")

            oRPrdNoCol = oRColumns.Add("colpordno", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oRPrdNoCol.TitleObject.Caption = "Production Order No"
            oRPrdNoCol.Width = 40
            oRPrdNoCol.DataBind.SetBound(True, "", "URPrdNo")
            oRPrdNoCol.Visible = False

            oRBsLnNoCol = oRColumns.Add("colbaslino", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oRBsLnNoCol.TitleObject.Caption = "Base Line No"
            oRBsLnNoCol.Visible = False
            oRBsLnNoCol.Width = 40
            oRBsLnNoCol.Editable = True
            oRBsLnNoCol.DataBind.SetBound(True, "", "URBsLnNo")

            oROprSeqCol = oRColumns.Add("colseqn", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oROprSeqCol.TitleObject.Caption = "Operation Sequence"
            oROprSeqCol.Width = 150
            oROprSeqCol.Editable = True
            oROprSeqCol.DataBind.SetBound(True, "", "UROprSeq")

            oRPrntIdCol = oRColumns.Add("colparid", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oRPrntIdCol.TitleObject.Caption = "Parent Id"
            oRPrntIdCol.Width = 100
            oRPrntIdCol.Editable = True
            oRPrntIdCol.DataBind.SetBound(True, "", "URPrntId")

            oROprCodeCol = oRColumns.Add("coloprcd", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oROprCodeCol.TitleObject.Caption = "Operation"
            oROprCodeCol.Width = 125
            oROprCodeCol.Editable = True
            oROprCodeCol.DataBind.SetBound(True, "", "UROprCode")

            oROprNameCol = oRColumns.Add("colopnam", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oROprNameCol.TitleObject.Caption = "Operation Name"
            oROprNameCol.DisplayDesc = True
            oROprNameCol.Width = 150
            oROprNameCol.Editable = True
            oROprNameCol.DataBind.SetBound(True, "", "UROprName")
            'oROprCodeCol.DisplayDesc = True

            oRRewrkCol = oRColumns.Add("colrewrk", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oRRewrkCol.TitleObject.Caption = "Rework"
            oRRewrkCol.Visible = True
            oRRewrkCol.Editable = False
            oRRewrkCol.DataBind.SetBound(True, "", "URRewrk")

            oRRoutCodCol = oRColumns.Add("colroutid", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oRRoutCodCol.TitleObject.Caption = "Route ID"
            oRRoutCodCol.Width = 40
            oRRoutCodCol.Visible = False
            oRRoutCodCol.DataBind.SetBound(True, "", "URRoutCod")

            oRSeqBsLnIdCol = oRColumns.Add("colseqlnid", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oRSeqBsLnIdCol.TitleObject.Caption = "Sequence Base Line ID"
            oRSeqBsLnIdCol.Width = 40
            'oRSeqBsLnIdCol.Visible = False
            oRSeqBsLnIdCol.DataBind.SetBound(True, "", "URSqBsLnId")

            oRPrdQtyCol = oRColumns.Add("colprqty", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oRPrdQtyCol.TitleObject.Caption = "Produced Qty"
            oRPrdQtyCol.Width = 40
            oRPrdQtyCol.Editable = False
            oRPrdQtyCol.DataBind.SetBound(True, "", "URPrdQty")

            oRPassQtyCol = oRColumns.Add("colpasqy", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oRPassQtyCol.TitleObject.Caption = "Passed Qty"
            oRPassQtyCol.Width = 40
            oRPassQtyCol.Editable = False
            oRPassQtyCol.DataBind.SetBound(True, "", "URPassQty")

            oRRewrkQtyCol = oRColumns.Add("colrwqty", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oRRewrkQtyCol.TitleObject.Caption = "Rework Qty"
            oRRewrkQtyCol.Width = 40
            oRRewrkQtyCol.Editable = False
            oRRewrkQtyCol.DataBind.SetBound(True, "", "URewrkQty")

            oRSrpQtyCol = oRColumns.Add("colscqty", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oRSrpQtyCol.TitleObject.Caption = "Scrap Qty"
            oRSrpQtyCol.Width = 40
            oRSrpQtyCol.Editable = False
            oRSrpQtyCol.DataBind.SetBound(True, "", "URSrpQty")

            oRLbrCostCol = oRColumns.Add("colbrcst", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oRLbrCostCol.TitleObject.Caption = "Labour Cost"
            oRLbrCostCol.Width = 40
            oRLbrCostCol.Editable = False
            oRLbrCostCol.DataBind.SetBound(True, "", "URLbrCost")

            oRMCCostCol = oRColumns.Add("colmcst", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oRMCCostCol.TitleObject.Caption = "Machine Cost"
            oRMCCostCol.Width = 40
            oRMCCostCol.Editable = False
            oRMCCostCol.DataBind.SetBound(True, "", "URMCCost")

            oRTLCostCol = oRColumns.Add("coltlcst", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oRTLCostCol.TitleObject.Caption = "Tool Cost"
            oRTLCostCol.Width = 40
            oRTLCostCol.Editable = False
            oRTLCostCol.DataBind.SetBound(True, "", "URTLCost")

            oRSubCtCostCol = oRColumns.Add("colsccst", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oRSubCtCostCol.TitleObject.Caption = "Sub Contracting Cost"
            oRSubCtCostCol.Width = 40
            oRSubCtCostCol.Editable = False
            oRSubCtCostCol.DataBind.SetBound(True, "", "URSubCtCst")

            oRSrpCostCol = oRColumns.Add("colspcst", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oRSrpCostCol.TitleObject.Caption = "Scrap Cost"
            oRSrpCostCol.Width = 40
            oRSrpCostCol.Editable = False
            oRSrpCostCol.DataBind.SetBound(True, "", "URSrpCost")

            oRWODocCol = oRColumns.Add("colword", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oRWODocCol.TitleObject.Caption = "Work Order Doc. Entry"
            oRWODocCol.Width = 40
            oRWODocCol.Editable = False
            oRWODocCol.DataBind.SetBound(True, "", "URWODoc")

            oOthrQty1Col = oRColumns.Add("colotrqty1", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oOthrQty1Col.TitleObject.Caption = "Other Qty1"
            oOthrQty1Col.Width = 40
            oOthrQty1Col.Editable = False
            oOthrQty1Col.DataBind.SetBound(True, "", "UROthrQty1")

            oOthrQty2Col = oRColumns.Add("colotrqty2", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oOthrQty2Col.TitleObject.Caption = "Other Qty2"
            oOthrQty2Col.Width = 40
            oOthrQty2Col.Editable = False
            oOthrQty2Col.DataBind.SetBound(True, "", "UROthrQty2")

            oOthrCost1Col = oRColumns.Add("colotrcst1", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oOthrCost1Col.TitleObject.Caption = "Other Cost1"
            oOthrCost1Col.Width = 40
            oOthrCost1Col.Editable = False
            oOthrCost1Col.DataBind.SetBound(True, "", "UROthrCst1")

            oOthrCost2Col = oRColumns.Add("colotrcst2", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oOthrCost2Col.TitleObject.Caption = "Other Cost2"
            oOthrCost2Col.Width = 40
            oOthrCost2Col.Editable = False
            oOthrCost2Col.DataBind.SetBound(True, "", "UROthrCst2")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This is used to Load the Operation Name in the Combo
    ''' </summary>
    ''' <remarks></remarks>
    ''' 
    Private Function getOperatioName(ByVal oprcode As String) As String
        Dim oRS As SAPbobsCOM.Recordset
        '& "Where T1.U_OprName is not Null and T0.U_ItemCode = '" & ItemCode & "' and T1.U_OprCode='" & oprcode & "'" _
        Dim strsql As String
        oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strsql = "Select isnull(T1.U_OprName,'') From [@PSSIT_ORTE] T0 " _
                   & "Inner Join [@PSSIT_RTE4] T1 On T1.Code = T0.Code " _
                   & "Where T1.U_OprName is not Null and T1.U_OprCode='" & oprcode & "'" _
                   & "Group by T1.U_OprName,T1.U_OprCode "
        oRS.DoQuery(strsql)
        Return oRS.Fields.Item(0).Value
    End Function
    Private Sub OperationCombo(ByVal opCode)
        Dim oRs As SAPbobsCOM.Recordset
        Dim StrSql As String
        Try
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            StrSql = "Select T1.U_OprCode,left(T1.U_OprName,30) From [@PSSIT_ORTE] T0 " _
            & "Inner Join [@PSSIT_RTE4] T1 On T1.Code = T0.Code " _
            & "Where T1.U_OprName is not Null and T0.U_ItemCode = '" & oProdTxt.Value & "' and  t0.code = '" & opCode & "'" _
            & "Group by T1.U_OprName,T1.U_OprCode "
            If oROprNameCol.ValidValues.Count > 0 Then
                For i As Int16 = oROprNameCol.ValidValues.Count - 1 To 0 Step -1
                    oROprNameCol.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            oRs.DoQuery(StrSql)
            If oRs.RecordCount > 0 Then
                oRs.MoveFirst()
                While oRs.EoF = False
                    oROprNameCol.ValidValues.Add(oRs.Fields.Item(0).Value, oRs.Fields.Item(1).Value)
                    oRs.MoveNext()
                End While
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' Configuring Cost Details Header Items
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AddCostHdr()
        '**************************Items************************************
        Dim oTItem As SAPbouiCOM.Item
        Dim oLItem As SAPbouiCOM.Item
        Dim oRItem As SAPbouiCOM.Item
        Dim oBItem As SAPbouiCOM.Item
        Try

            oTItem = oForm.Items.Item("83")
            oLItem = oForm.Items.Item("57")
            oRItem = oForm.Items.Item("55")
            oBItem = oForm.Items.Item("54")

            oCLItem = oForm.Items.Add("lbltccost", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oCLItem.FromPane = 4
            oCLItem.ToPane = 4
            oCLItem.Top = oTItem.Top + 5
            oCLItem.Height = 14
            oCLItem.Width = 100
            oCLItem.Left = oLItem.Left + 5
            oCompCostLbl = oCLItem.Specific
            oCompCostLbl.Caption = "Component Cost"

            oTLCItem = oForm.Items.Add("lbltlrcst", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oTLCItem.FromPane = 4
            oTLCItem.ToPane = 4
            oTLCItem.Top = oTItem.Top + 5
            oTLCItem.Height = 14
            oTLCItem.Width = 100
            oTLCItem.Left = oLItem.Left + 330
            oLabCostLbl = oTLCItem.Specific
            oLabCostLbl.Caption = "Total Labour Cost"

            oTMCItem = oForm.Items.Add("lbltmcst", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oTMCItem.FromPane = 4
            oTMCItem.ToPane = 4
            oTMCItem.Top = oTItem.Top + 20
            oTMCItem.Height = 14
            oTMCItem.Width = 100
            oTMCItem.Left = oLItem.Left + 5
            oMCCostLbl = oTMCItem.Specific
            oMCCostLbl.Caption = "Total Machine Cost"

            oTTCItem = oForm.Items.Add("lbltlcst", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oTTCItem.FromPane = 4
            oTTCItem.ToPane = 4
            oTTCItem.Top = oTItem.Top + 20
            oTTCItem.Height = 14
            oTTCItem.Width = 100
            oTTCItem.Left = oLItem.Left + 330
            oToolCostLbl = oTTCItem.Specific
            oToolCostLbl.Caption = "Total Tool Cost"

            oTSCItem = oForm.Items.Add("lblsucst", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oTSCItem.FromPane = 4
            oTSCItem.ToPane = 4
            oTSCItem.Top = oTItem.Top + 35
            oTSCItem.Height = 14
            oTSCItem.Width = 100
            oTSCItem.Left = oLItem.Left + 5
            oSCCostLbl = oTSCItem.Specific
            oSCCostLbl.Caption = "Total S.C. Cost"

            oLTCstItem1 = oForm.Items.Add("lblcst1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oLTCstItem1.FromPane = 4
            oLTCstItem1.ToPane = 4
            oLTCstItem1.Top = oTItem.Top + 35
            oLTCstItem1.Height = 14
            oLTCstItem1.Width = 100
            oLTCstItem1.Left = oLItem.Left + 330
            oOthrCost1Lbl = oLTCstItem1.Specific
            oOthrCost1Lbl.Caption = "Other Cost1"

            oLTCstItem2 = oForm.Items.Add("lblcst2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oLTCstItem2.FromPane = 4
            oLTCstItem2.ToPane = 4
            oLTCstItem2.Top = oTItem.Top + 50
            oLTCstItem2.Height = 14
            oLTCstItem2.Width = 100
            oLTCstItem2.Left = oLItem.Left + 5
            oOthrCost2Lbl = oLTCstItem2.Specific
            oOthrCost2Lbl.Caption = "Other Cost2"

            oLTCstItem3 = oForm.Items.Add("lblcst3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oLTCstItem3.FromPane = 4
            oLTCstItem3.ToPane = 4
            oLTCstItem3.Top = oTItem.Top + 50
            oLTCstItem3.Height = 14
            oLTCstItem3.Width = 100
            oLTCstItem3.Left = oLItem.Left + 330
            oOthrCost3Lbl = oLTCstItem3.Specific
            oOthrCost3Lbl.Caption = "Other Cost3"

            oLTCstItem4 = oForm.Items.Add("lblcst4", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oLTCstItem4.FromPane = 4
            oLTCstItem4.ToPane = 4
            oLTCstItem4.Top = oTItem.Top + 65
            oLTCstItem4.Height = 14
            oLTCstItem4.Width = 100
            oLTCstItem4.Left = oLItem.Left + 5
            oOthrCost4Lbl = oLTCstItem4.Specific
            oOthrCost4Lbl.Caption = "Other Cost4"

            oTCItem = oForm.Items.Add("lbltocst", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oTCItem.FromPane = 4
            oTCItem.ToPane = 4
            oTCItem.Top = oTItem.Top + 190
            oTCItem.Height = 14
            oTCItem.Width = 100
            oTCItem.Left = oLItem.Left + 330
            oTotCostLbl = oTCItem.Specific
            oTotCostLbl.Caption = "Total Cost"

            '***********Edit Text************************
            oProdSerCombo = oForm.Items.Item("22").Specific
            oProdNoTxt = oForm.Items.Item("18").Specific
            oProdTxt = oForm.Items.Item("6").Specific

            oTCLItem = oForm.Items.Add("txtccost", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oTCLItem.FromPane = 4
            oTCLItem.ToPane = 4
            oTCLItem.Top = oTItem.Top + 5
            oTCLItem.Height = 14
            oTCLItem.Width = 90
            oTCLItem.Left = oLItem.Left + 115
            oCompCostTxt = oTCLItem.Specific
            oForm.Items.Item("txtccost").Enabled = False
            oCompCostTxt.DataBind.SetBound(True, "", "UCompCost")

            oTTLCItem = oForm.Items.Add("txtlrcost", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oTTLCItem.FromPane = 4
            oTTLCItem.ToPane = 4
            oTTLCItem.Top = oTItem.Top + 5
            oTTLCItem.Height = 14
            oTTLCItem.Width = 105
            oTTLCItem.Left = oLItem.Left + 463
            oLabCostTxt = oTTLCItem.Specific
            oForm.Items.Item("txtlrcost").Enabled = False
            oLabCostTxt.DataBind.SetBound(True, "", "ULabCost")

            oTTMCItem = oForm.Items.Add("txtmcst", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oTTMCItem.FromPane = 4
            oTTMCItem.ToPane = 4
            oTTMCItem.Top = oTItem.Top + 20
            oTTMCItem.Height = 14
            oTTMCItem.Width = 90
            oTTMCItem.Left = oLItem.Left + 115
            oMCCostTxt = oTTMCItem.Specific
            oForm.Items.Item("txtmcst").Enabled = False
            oMCCostTxt.DataBind.SetBound(True, "", "UMCCost")

            oTTTCItem = oForm.Items.Add("txttlcst", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oTTTCItem.FromPane = 4
            oTTTCItem.ToPane = 4
            oTTTCItem.Top = oTItem.Top + 20
            oTTTCItem.Height = 14
            oTTTCItem.Width = 105
            oTTTCItem.Left = oLItem.Left + 463
            oToolCostTxt = oTTTCItem.Specific
            oForm.Items.Item("txttlcst").Enabled = False
            oToolCostTxt.DataBind.SetBound(True, "", "UToolCost")

            oTTSCItem = oForm.Items.Add("txtsucst", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oTTSCItem.FromPane = 4
            oTTSCItem.ToPane = 4
            oTTSCItem.Top = oTItem.Top + 35
            oTTSCItem.Height = 14
            oTTSCItem.Width = 90
            oTTSCItem.Left = oLItem.Left + 115
            oSCCostTxt = oTTSCItem.Specific
            oForm.Items.Item("txtsucst").Enabled = False
            oSCCostTxt.DataBind.SetBound(True, "", "USCCost")

            oTTCstItem1 = oForm.Items.Add("txtcst1", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oTTCstItem1.FromPane = 4
            oTTCstItem1.ToPane = 4
            oTTCstItem1.Top = oTItem.Top + 35
            oTTCstItem1.Height = 14
            oTTCstItem1.Width = 105
            oTTCstItem1.Left = oLItem.Left + 463
            oOthrCost1Txt = oTTCstItem1.Specific
            oForm.Items.Item("txtcst1").Enabled = False
            oOthrCost1Txt.DataBind.SetBound(True, "", "UCOthrCst1")

            oTTCstItem2 = oForm.Items.Add("txtcst2", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oTTCstItem2.FromPane = 4
            oTTCstItem2.ToPane = 4
            oTTCstItem2.Top = oTItem.Top + 50
            oTTCstItem2.Height = 14
            oTTCstItem2.Width = 90
            oTTCstItem2.Left = oLItem.Left + 115
            oOthrCost2Txt = oTTCstItem2.Specific
            oForm.Items.Item("txtcst2").Enabled = False
            oOthrCost2Txt.DataBind.SetBound(True, "", "UCOthrCst2")

            oTTCstItem3 = oForm.Items.Add("txtcst3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oTTCstItem3.FromPane = 4
            oTTCstItem3.ToPane = 4
            oTTCstItem3.Top = oTItem.Top + 50
            oTTCstItem3.Height = 14
            oTTCstItem3.Width = 105
            oTTCstItem3.Left = oLItem.Left + 463
            oOthrCost3Txt = oTTCstItem3.Specific
            oForm.Items.Item("txtcst3").Enabled = False
            oOthrCost3Txt.DataBind.SetBound(True, "", "UCOthrCst3")

            oTTCstItem4 = oForm.Items.Add("txtcst4", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oTTCstItem4.FromPane = 4
            oTTCstItem4.ToPane = 4
            oTTCstItem4.Top = oTItem.Top + 65
            oTTCstItem4.Height = 14
            oTTCstItem4.Width = 90
            oTTCstItem4.Left = oLItem.Left + 115
            oOthrCost4Txt = oTTCstItem4.Specific
            oForm.Items.Item("txtcst4").Enabled = False
            oOthrCost4Txt.DataBind.SetBound(True, "", "UCOthrCst4")

            oTCTItem = oForm.Items.Add("txttocst", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oTCTItem.FromPane = 4
            oTCTItem.ToPane = 4
            oTCTItem.Top = oTItem.Top + 190
            oTCTItem.Height = 14
            oTCTItem.Width = 105
            oTCTItem.Left = oLItem.Left + 463
            oTotCostTxt = oTCTItem.Specific
            oForm.Items.Item("txttocst").Enabled = False
            oTotCostTxt.DataBind.SetBound(True, "", "UTotCost")

            oCHCodeItem = oForm.Items.Add("txtcode", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oCHCodeItem.FromPane = 4
            oCHCodeItem.ToPane = 4
            oCHCodeItem.Top = oTItem.Top + 5
            oCHCodeItem.Height = 14
            oCHCodeItem.Width = 105
            oCHCodeItem.Left = oLItem.Left + 220
            oCHCodeTxt = oCHCodeItem.Specific
            oCHCodeTxt.DataBind.SetBound(True, "", "UCHCode")
            oForm.Items.Item("txtcode").Enabled = False
            oForm.Items.Item("txtcode").Visible = False
            UCHCode.Value = GenerateSerialNo("PSSIT_WOR3")

            oPrdSerItem = oForm.Items.Add("txtprdser", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oPrdSerItem.FromPane = 4
            oPrdSerItem.ToPane = 4
            oPrdSerItem.Top = oTItem.Top + 20
            oPrdSerItem.Height = 14
            oPrdSerItem.Width = 105
            oPrdSerItem.Left = oLItem.Left + 220
            oPrdSerTxt = oPrdSerItem.Specific
            oForm.Items.Item("txtprdser").Enabled = False
            oForm.Items.Item("txtprdser").Visible = False
            oPrdSerTxt.DataBind.SetBound(True, "", "UProdSer")
            oPrdSerTxt.Value = oProdSerCombo.Selected.Value

            oPrdNoItem = oForm.Items.Add("txtprdno", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oPrdNoItem.FromPane = 4
            oPrdNoItem.ToPane = 4
            oPrdNoItem.Top = oTItem.Top + 35
            oPrdNoItem.Height = 14
            oPrdNoItem.Width = 105
            oPrdNoItem.Left = oLItem.Left + 220
            oPrdNoTxt = oPrdNoItem.Specific
            oForm.Items.Item("txtprdno").Enabled = False
            oForm.Items.Item("txtprdno").Visible = False
            oPrdNoTxt.DataBind.SetBound(True, "", "UProdNo")
            UProdNo.Value = oProdNoTxt.Value

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Configuring the Matrix items/controls
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AddCostMatrix()
        Dim oTItem As SAPbouiCOM.Item
        Dim oLItem As SAPbouiCOM.Item
        Dim oRItem As SAPbouiCOM.Item
        Dim oBItem As SAPbouiCOM.Item
        Try

            oCMItem = oForm.Items.Add("matcost", SAPbouiCOM.BoFormItemTypes.it_MATRIX)
            oCMItem.FromPane = 4
            oCMItem.ToPane = 4
            oTItem = oForm.Items.Item("83")
            oLItem = oForm.Items.Item("57")
            oRItem = oForm.Items.Item("55")
            oBItem = oForm.Items.Item("54")

            oCMItem.Top = oTItem.Top + 90
            oCMItem.Height = 100
            oCMItem.Width = 567
            oCMItem.Left = oLItem.Left + 5

            oCostMatrix = oCMItem.Specific
            oCColumns = oCostMatrix.Columns

            oCRowNoCol = oCColumns.Add("#", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oCRowNoCol.TitleObject.Caption = "#"
            oCRowNoCol.Width = 20
            oCRowNoCol.Editable = False
            oCRowNoCol.DataBind.SetBound(True, "", "UCLineid")

            oCCodeCol = oCColumns.Add("colcode", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oCCodeCol.TitleObject.Caption = "Code"
            oCCodeCol.Width = 100
            oCCodeCol.Editable = False
            oCCodeCol.DataBind.SetBound(True, "", "UCCode")

            oCDocEntCol = oCColumns.Add("coldocent", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oCDocEntCol.TitleObject.Caption = "DocEntry"
            oCDocEntCol.Width = 80
            oCDocEntCol.Editable = False
            oCDocEntCol.DataBind.SetBound(True, "", "UCDocEnt")

            oCPrdSerCol = oCColumns.Add("colpordser", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oCPrdSerCol.TitleObject.Caption = "Production Series"
            oCPrdSerCol.Width = 80
            oCPrdSerCol.Editable = False
            oCPrdSerCol.DataBind.SetBound(True, "", "UCPrdSer")

            oCPrdNoCol = oCColumns.Add("colpordno", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oCPrdNoCol.TitleObject.Caption = "Production Order No"
            oCPrdNoCol.Width = 80
            oCPrdNoCol.Editable = False
            oCPrdNoCol.DataBind.SetBound(True, "", "UCPrdNo")

            oCFxdCostCol = oCColumns.Add("colfcost", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oCFxdCostCol.TitleObject.Caption = "Fixed Cost"
            oCFxdCostCol.Width = 100
            oCFxdCostCol.Editable = False
            oCFxdCostCol.DataBind.SetBound(True, "", "UCFxdCost")

            oCUnitCostCol = oCColumns.Add("colutcost", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oCUnitCostCol.TitleObject.Caption = "Unit Cost"
            oCUnitCostCol.Width = 100
            oCUnitCostCol.Editable = False
            oCUnitCostCol.DataBind.SetBound(True, "", "UCUnitCst")

            oCAbsMthdCol = oCColumns.Add("colabsmd", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oCAbsMthdCol.TitleObject.Caption = "Absorbtion Method"
            oCAbsMthdCol.Width = 100
            oCAbsMthdCol.Editable = False
            oCAbsMthdCol.DataBind.SetBound(True, "", "UCAbsMthd")

            oCAcctCodeCol = oCColumns.Add("colaccode", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oCAcctCodeCol.TitleObject.Caption = "Account Code"
            oCAcctCodeCol.Width = 100
            oCAcctCodeCol.Editable = False
            oCAcctCodeCol.DataBind.SetBound(True, "", "UCActCod")

            oCAcctNameCol = oCColumns.Add("colacname", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oCAcctNameCol.TitleObject.Caption = "Account Name"
            oCAcctNameCol.Width = 150
            oCAcctNameCol.Editable = False
            oCAcctNameCol.DataBind.SetBound(True, "", "UCActNam")

            oCTotCostCol = oCColumns.Add("coltfcst", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oCTotCostCol.TitleObject.Caption = "Total Fixed Cost"
            oCTotCostCol.Width = 150
            oCTotCostCol.Editable = False
            oCTotCostCol.DataBind.SetBound(True, "", "UCTotCost")

            oOthrCostCol = oCColumns.Add("colothrcst", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oOthrCostCol.TitleObject.Caption = "Other Cost1"
            oOthrCostCol.Width = 150
            oOthrCostCol.Editable = False
            oOthrCostCol.DataBind.SetBound(True, "", "UOthrCost")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#Region "NoObject Table Data"
    Private Sub AddRoutChildTable()
        '****************Adding the child data to the database table***********
        Dim IntICount, RoutChild As Integer
        Dim oTemprs As SAPbobsCOM.Recordset
        oTemprs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '   oTemprs.DoQuery("Delete from [@PSSIT_WOR2] where U_Pordser='" & URPrdSer.Value & "' and U_PordNo=" & URPrdNo.Value)

        Try
            '************** Records Added in Tools Bom Matrix **********
            For IntICount = 1 To oRoutMatrix.RowCount
                oRoutMatrix.GetLineData(IntICount)
                If PSSIT_WOR2.GetByKey(URCode.Value) = True Then
                    PSSIT_WOR2.Code = URCode.Value
                    PSSIT_WOR2.Name = URName.Value
                    PSSIT_WOR2.UserFields.Fields.Item("U_Pordser").Value = URPrdSer.Value
                    PSSIT_WOR2.UserFields.Fields.Item("U_Pordno").Value = URPrdNo.Value
                    PSSIT_WOR2.UserFields.Fields.Item("U_Baslino").Value = URRowNo.Value
                    PSSIT_WOR2.UserFields.Fields.Item("U_Seqnce").Value = UROprSeq.Value
                    PSSIT_WOR2.UserFields.Fields.Item("U_Parid").Value = URPrntId.Value
                    PSSIT_WOR2.UserFields.Fields.Item("U_Rework").Value = URRewrk.Value
                    PSSIT_WOR2.UserFields.Fields.Item("U_Oprcode").Value = UROprCode.Value
                    PSSIT_WOR2.UserFields.Fields.Item("U_Oprname").Value = getOperatioName(UROprCode.Value) 'UROprName.Value
                    PSSIT_WOR2.UserFields.Fields.Item("U_Rteid").Value = URRoutCod.Value
                    PSSIT_WOR2.UserFields.Fields.Item("U_Seqbaslino").Value = URSqBsLnId.Value
                    RoutChild = PSSIT_WOR2.Update()
                Else
                    PSSIT_WOR2.Code = URCode.Value
                    PSSIT_WOR2.Name = URName.Value
                    PSSIT_WOR2.UserFields.Fields.Item("U_Pordser").Value = URPrdSer.Value
                    PSSIT_WOR2.UserFields.Fields.Item("U_Pordno").Value = URPrdNo.Value
                    PSSIT_WOR2.UserFields.Fields.Item("U_Baslino").Value = URRowNo.Value
                    PSSIT_WOR2.UserFields.Fields.Item("U_Seqnce").Value = UROprSeq.Value
                    PSSIT_WOR2.UserFields.Fields.Item("U_Rework").Value = "N"
                    PSSIT_WOR2.UserFields.Fields.Item("U_Parid").Value = URPrntId.Value
                    PSSIT_WOR2.UserFields.Fields.Item("U_Oprcode").Value = UROprCode.Value
                    PSSIT_WOR2.UserFields.Fields.Item("U_Oprname").Value = getOperatioName(UROprCode.Value) 'UROprName.Value
                    PSSIT_WOR2.UserFields.Fields.Item("U_Rteid").Value = URRoutCod.Value
                    PSSIT_WOR2.UserFields.Fields.Item("U_Seqbaslino").Value = URSqBsLnId.Value
                    RoutChild = PSSIT_WOR2.Add()
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub AddCostDetailsTable()
        '****************Adding the Cost Details data to the database table***********
        Dim IntICount, CostChild As Integer

        '   Dim oTemprs As SAPbobsCOM.Recordset
        '  oTemprs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '  oTemprs.DoQuery("Delete from [@PSSIT_WOR3] where U_Pordser='" & URPrdSer.Value & "' and U_PordNo=" & URPrdNo.Value)

        Try
            For IntICount = 1 To oCostMatrix.RowCount
                oCostMatrix.GetLineData(IntICount)
                If PSSIT_WOR4.GetByKey(UCCode.Value) = True Then
                    PSSIT_WOR4.Code = UCCode.Value
                    PSSIT_WOR4.Name = UCCode.Value
                    PSSIT_WOR4.UserFields.Fields.Item("U_Pordser").Value = UCPrdSer.Value
                    PSSIT_WOR4.UserFields.Fields.Item("U_Pordno").Value = UCPrdNo.Value
                    PSSIT_WOR4.UserFields.Fields.Item("U_DocEntry").Value = UCDocEntry.Value
                    PSSIT_WOR4.UserFields.Fields.Item("U_Lineid").Value = UCLineid.Value
                    PSSIT_WOR4.UserFields.Fields.Item("U_FCost").Value = UCFxdCost.Value
                    PSSIT_WOR4.UserFields.Fields.Item("U_UnitCost").Value = UCUnitCst.Value
                    PSSIT_WOR4.UserFields.Fields.Item("U_Absmthd").Value = UCAbsMthd.Value
                    PSSIT_WOR4.UserFields.Fields.Item("U_Accode").Value = UCActCod.Value
                    PSSIT_WOR4.UserFields.Fields.Item("U_Acname").Value = UCActNam.Value
                    PSSIT_WOR4.UserFields.Fields.Item("U_Totfcst").Value = UCTotCost.Value
                    PSSIT_WOR4.UserFields.Fields.Item("U_Adnl1").Value = UOthrCost.Value
                    CostChild = PSSIT_WOR4.Update()
                Else
                    PSSIT_WOR4.Code = UCCode.Value
                    PSSIT_WOR4.Name = UCCode.Value
                    PSSIT_WOR4.UserFields.Fields.Item("U_Pordser").Value = UCPrdSer.Value
                    PSSIT_WOR4.UserFields.Fields.Item("U_Pordno").Value = UCPrdNo.Value
                    PSSIT_WOR4.UserFields.Fields.Item("U_DocEntry").Value = UCDocEntry.Value
                    PSSIT_WOR4.UserFields.Fields.Item("U_Lineid").Value = UCLineid.Value
                    PSSIT_WOR4.UserFields.Fields.Item("U_FCost").Value = UCFxdCost.Value
                    PSSIT_WOR4.UserFields.Fields.Item("U_UnitCost").Value = UCUnitCst.Value
                    PSSIT_WOR4.UserFields.Fields.Item("U_Absmthd").Value = UCAbsMthd.Value
                    PSSIT_WOR4.UserFields.Fields.Item("U_Accode").Value = UCActCod.Value
                    PSSIT_WOR4.UserFields.Fields.Item("U_Acname").Value = UCActNam.Value
                    PSSIT_WOR4.UserFields.Fields.Item("U_Totfcst").Value = UCTotCost.Value
                    PSSIT_WOR4.UserFields.Fields.Item("U_Adnl1").Value = UOthrCost.Value
                    CostChild = PSSIT_WOR4.Add()
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub AddCostHeaderTable()
        Try
            '****************Adding the child data to the database table***********
            Dim CostChild As Integer
            '************** Records Added in Tools Bom Matrix **********
            If PSSIT_WOR3.GetByKey(UCHCode.Value) = True Then
                PSSIT_WOR3.Code = UCHCode.Value
                PSSIT_WOR3.Name = UCHCode.Value
                PSSIT_WOR3.UserFields.Fields.Item("U_Pordser").Value = UProdSer.Value
                PSSIT_WOR3.UserFields.Fields.Item("U_Pordno").Value = UProdNo.Value
                PSSIT_WOR3.UserFields.Fields.Item("U_Totcmpcst").Value = UCompCost.Value
                PSSIT_WOR3.UserFields.Fields.Item("U_Totlbrcst").Value = ULabCost.Value
                PSSIT_WOR3.UserFields.Fields.Item("U_Totmccst").Value = UMCCost.Value
                PSSIT_WOR3.UserFields.Fields.Item("U_Tottoolcst").Value = UToolCost.Value
                PSSIT_WOR3.UserFields.Fields.Item("U_Totsubcst").Value = USCCost.Value
                PSSIT_WOR3.UserFields.Fields.Item("U_Totcst").Value = UTotCost.Value
                PSSIT_WOR3.UserFields.Fields.Item("U_Adnl1").Value = UCOthrCst1.Value
                PSSIT_WOR3.UserFields.Fields.Item("U_Adnl2").Value = UCOthrCst2.Value
                PSSIT_WOR3.UserFields.Fields.Item("U_Adnl3").Value = UCOthrCst3.Value
                PSSIT_WOR3.UserFields.Fields.Item("U_Adnl4").Value = UCOthrCst4.Value
                CostChild = PSSIT_WOR3.Update()
            Else
                PSSIT_WOR3.Code = UCHCode.Value
                PSSIT_WOR3.Name = UCHCode.Value
                PSSIT_WOR3.UserFields.Fields.Item("U_Pordser").Value = UProdSer.Value
                PSSIT_WOR3.UserFields.Fields.Item("U_Pordno").Value = UProdNo.Value
                PSSIT_WOR3.UserFields.Fields.Item("U_Totcmpcst").Value = UCompCost.Value
                PSSIT_WOR3.UserFields.Fields.Item("U_Totlbrcst").Value = ULabCost.Value
                PSSIT_WOR3.UserFields.Fields.Item("U_Totmccst").Value = UMCCost.Value
                PSSIT_WOR3.UserFields.Fields.Item("U_Tottoolcst").Value = UToolCost.Value
                PSSIT_WOR3.UserFields.Fields.Item("U_Totsubcst").Value = USCCost.Value
                PSSIT_WOR3.UserFields.Fields.Item("U_Totcst").Value = UTotCost.Value
                PSSIT_WOR3.UserFields.Fields.Item("U_Adnl1").Value = UCOthrCst1.Value
                PSSIT_WOR3.UserFields.Fields.Item("U_Adnl2").Value = UCOthrCst2.Value
                PSSIT_WOR3.UserFields.Fields.Item("U_Adnl3").Value = UCOthrCst3.Value
                PSSIT_WOR3.UserFields.Fields.Item("U_Adnl4").Value = UCOthrCst4.Value
                CostChild = PSSIT_WOR3.Add()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Load the Tools Bom Matrix for the correspondin Schedule and Department
    ''' </summary>
    ''' <remarks></remarks>
    ''' 
    Private Sub SBO_Application_FormDataEvent1(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent

        If BusinessObjectInfo.FormTypeEx = "65211" And BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD And BusinessObjectInfo.BeforeAction = False Then
            ClearRoutMatrix()
            LoadRoutMatrixData()
        End If

    End Sub
    Private Sub LoadRoutMatrixData()
        Dim StrSql As String
        Dim oRs As SAPbobsCOM.Recordset
        Dim IntICount As Integer
        Try
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oProdNoTxt.String = "*" Then
                StrSql = "select * from [@PSSIT_WOR2] where U_Pordno = ''"
            Else
                StrSql = "select * from [@PSSIT_WOR2] where U_Pordno = '" & oProdNoTxt.Value & "'"
            End If

            oRs.DoQuery(StrSql)
            If oRs.RecordCount > 0 Then
                oRs.MoveFirst()
                For IntICount = 0 To oRs.RecordCount - 1
                    URCode.Value = oRs.Fields.Item("Code").Value
                    URName.Value = oRs.Fields.Item("Name").Value
                    URRowNo.Value = oRoutMatrix.RowCount + 1
                    'URRowNo.Value = oRs.Fields.Item("U_Baslino").Value
                    URPrdSer.Value = oRs.Fields.Item("U_Pordser").Value
                    URPrdNo.Value = oRs.Fields.Item("U_Pordno").Value
                    UROprSeq.Value = oRs.Fields.Item("U_Seqnce").Value
                    URPrntId.Value = oRs.Fields.Item("U_Parid").Value
                    UROprCode.Value = oRs.Fields.Item("U_Oprcode").Value
                    UROprName.Value = oRs.Fields.Item("U_Oprname").Value
                    URRewrk.Value = oRs.Fields.Item("U_Rework").Value
                    URSqBsLnId.Value = oRs.Fields.Item("U_Seqbaslino").Value
                    URRoutCod.Value = oRs.Fields.Item("U_Rteid").Value
                    URPrdQty.Value = oRs.Fields.Item("U_ProdQty").Value
                    URPassQty.Value = oRs.Fields.Item("U_PassQty").Value
                    URewrkQty.Value = oRs.Fields.Item("U_RewrkQty").Value
                    URSrpQty.Value = oRs.Fields.Item("U_ScrapQty").Value
                    URLbrCost.Value = oRs.Fields.Item("U_LbrCst").Value
                    URMCCost.Value = oRs.Fields.Item("U_Mccst").Value
                    URTLCost.Value = oRs.Fields.Item("U_Toolcst").Value
                    URSubCtCst.Value = oRs.Fields.Item("U_Subcst").Value
                    URSrpCost.Value = oRs.Fields.Item("U_ScrapCst").Value
                    URWODoc.Value = oRs.Fields.Item("U_WoDoc").Value
                    UROthrQty1.Value = oRs.Fields.Item("U_Adnl1").Value
                    UROthrQty2.Value = oRs.Fields.Item("U_Adnl2").Value
                    UROthrCst1.Value = oRs.Fields.Item("U_Adnl3").Value
                    UROthrCst2.Value = oRs.Fields.Item("U_Adnl4").Value
                    oRoutMatrix.AddRow(1, oRoutMatrix.RowCount)
                    oRs.MoveNext()
                Next
            End If
            oRs.DoQuery("Select IsNull(Max(Code),0) as Code from [@PSSIT_WOR2]")
            oRDSerialNo = oRs.Fields.Item("Code").Value
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    Private Sub LoadCostDataFromDB()
        Dim StrSql As String
        Dim oRs As SAPbobsCOM.Recordset
        Dim IntICount As Integer
        Try
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oProdNoTxt.String = "*" Then
                StrSql = "select * from [@PSSIT_WOR3] where U_Pordno = ''"
            Else
                StrSql = "select * from [@PSSIT_WOR3] where U_Pordno = '" & oProdNoTxt.Value & "'"
            End If

            oRs.DoQuery(StrSql)
            If oRs.RecordCount > 0 Then
                oRs.MoveFirst()
                For IntICount = 0 To oRs.RecordCount - 1
                    UCHCode.Value = oRs.Fields.Item("Code").Value
                    UProdSer.Value = oRs.Fields.Item("U_Pordser").Value
                    UProdNo.Value = oRs.Fields.Item("U_Pordno").Value
                    UCompCost.Value = oRs.Fields.Item("U_Totcmpcst").Value
                    ULabCost.Value = oRs.Fields.Item("U_Totlbrcst").Value
                    UMCCost.Value = oRs.Fields.Item("U_Totmccst").Value
                    UToolCost.Value = oRs.Fields.Item("U_Tottoolcst").Value
                    USCCost.Value = oRs.Fields.Item("U_Totsubcst").Value
                    UTotCost.Value = oRs.Fields.Item("U_Totcst").Value
                    UCOthrCst1.Value = oRs.Fields.Item("U_Adnl1").Value
                    UCOthrCst2.Value = oRs.Fields.Item("U_Adnl2").Value
                    UCOthrCst3.Value = oRs.Fields.Item("U_Adnl3").Value
                    UCOthrCst4.Value = oRs.Fields.Item("U_Adnl4").Value
                    oRs.MoveNext()
                Next
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    Private Sub LoadCostDetailsFromDB()
        Dim StrSql As String
        Dim oRs As SAPbobsCOM.Recordset
        Dim IntICount As Integer
        Try
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oProdNoTxt.String = "*" Then
                StrSql = "select * from [@PSSIT_WOR4] where U_Pordno = ''"
            Else
                StrSql = "select * from [@PSSIT_WOR4] where U_Pordno = '" & oProdNoTxt.Value & "'"
            End If

            oRs.DoQuery(StrSql)
            If oRs.RecordCount > 0 Then
                oRs.MoveFirst()
                For IntICount = 0 To oRs.RecordCount - 1
                    UCPrdNo.Value = oRs.Fields.Item("U_Pordno").Value
                    UCPrdSer.Value = oRs.Fields.Item("U_Pordser").Value
                    UCDocEntry.Value = oRs.Fields.Item("U_DocEntry").Value
                    UCCode.Value = oRs.Fields.Item("Code").Value
                    UCLineid.Value = oRs.Fields.Item("U_Lineid").Value
                    UCFxdCost.Value = oRs.Fields.Item("U_Fcost").Value
                    UCUnitCst.Value = oRs.Fields.Item("U_UnitCost").Value
                    UCAbsMthd.Value = oRs.Fields.Item("U_Absmthd").Value
                    UCActCod.Value = oRs.Fields.Item("U_Accode").Value
                    UCActNam.Value = oRs.Fields.Item("U_Acname").Value
                    UCTotCost.Value = oRs.Fields.Item("U_Totfcst").Value
                    UOthrCost.Value = oRs.Fields.Item("U_Adnl1").Value
                    oCostMatrix.AddRow(1, oCostMatrix.RowCount)
                    oRs.MoveNext()
                Next
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
#End Region

#End Region
#Region "Production Order Report"
    Private Shared Sub SetCrystalLogin(ByVal sUser As String, ByVal sPassword As String, ByVal sServer As String, ByVal sCompanyDB As String, _
              ByRef oRpt As CrystalDecisions.CrystalReports.Engine.ReportDocument)

        Dim oDB As CrystalDecisions.CrystalReports.Engine.Database = oRpt.Database
        Dim oTables As CrystalDecisions.CrystalReports.Engine.Tables = oDB.Tables
        Dim oLogonInfo As CrystalDecisions.Shared.TableLogOnInfo
        Dim oConnectInfo As CrystalDecisions.Shared.ConnectionInfo = New CrystalDecisions.Shared.ConnectionInfo()
        oConnectInfo.DatabaseName = sCompanyDB
        oConnectInfo.ServerName = sServer
        oConnectInfo.UserID = sUser
        oConnectInfo.Password = sPassword
        ' Set the logon credentials for all tables
        For Each oTable As CrystalDecisions.CrystalReports.Engine.Table In oTables
            oLogonInfo = oTable.LogOnInfo
            oLogonInfo.ConnectionInfo = oConnectInfo
            oTable.ApplyLogOnInfo(oLogonInfo)
        Next
        ' Check for subreports
        Dim oSections As CrystalDecisions.CrystalReports.Engine.Sections
        Dim oSection As CrystalDecisions.CrystalReports.Engine.Section
        Dim oRptObjs As CrystalDecisions.CrystalReports.Engine.ReportObjects
        Dim oRptObj As CrystalDecisions.CrystalReports.Engine.ReportObject
        Dim oSubRptObj As CrystalDecisions.CrystalReports.Engine.SubreportObject
        Dim oSubRpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        oSections = oRpt.ReportDefinition.Sections
        For Each oSection In oSections
            oRptObjs = oSection.ReportObjects
            For Each oRptObj In oRptObjs

                If oRptObj.Kind = CrystalDecisions.Shared.ReportObjectKind.SubreportObject Then

                    ' This is a subreport so set the logon credentials for this report's tables
                    oSubRptObj = CType(oRptObj, CrystalDecisions.CrystalReports.Engine.SubreportObject)
                    ' Open the subreport
                    oSubRpt = oSubRptObj.OpenSubreport(oSubRptObj.SubreportName)

                    oDB = oSubRpt.Database
                    oTables = oDB.Tables

                    For Each oTable As CrystalDecisions.CrystalReports.Engine.Table In oTables
                        oLogonInfo = oTable.LogOnInfo
                        oLogonInfo.ConnectionInfo = oConnectInfo
                        oTable.ApplyLogOnInfo(oLogonInfo)
                    Next
                End If
            Next
        Next
    End Sub
#End Region
End Class
