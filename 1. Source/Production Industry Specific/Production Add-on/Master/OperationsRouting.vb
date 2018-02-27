'''' <summary>
'''' Author                     Created Date
'''' Suresh                      22/12/2008
'''' <remarks> This class is used for entering the Operations Routing details.</remarks>
Public Class OperationsRouting
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
    Private oParentDB, oRouteDB, oMachinesDB, oLabourDB, oToolsDB As SAPbouiCOM.DBDataSource
    Private PSSIT_RTE1, PSSIT_RTE2, PSSIT_RTE3 As SAPbobsCOM.UserTable
    '**************************Form************************************
    Private oForm As SAPbouiCOM.Form
    '**************************ChooseFromList************************************
    Private oChItemList, oChOprList, oChMacList, oChToolsList, oChSkGroupList As SAPbouiCOM.ChooseFromList
    Private oChItemBtnList As SAPbouiCOM.ChooseFromList
    Private Event ChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable)
    '**************************Items - EditText************************************
    Private oRouteIdTxt, oItemCodeTxt, oItemNameTxt, oDrawingNoTxt, oRevNoTxt, oRevDateTxt, oInfo1Txt, oInfo2Txt, oRemarksTxt As SAPbouiCOM.EditText
    '**************************Items - CheckBox************************************
    Private oActiveCheck, oDefRouteCheck As SAPbouiCOM.CheckBox
    '**************************Items - Button************************************
    Private oItemBtn As SAPbouiCOM.Button
    '**************************Items - Matrix************************************
    Private oRouteMatrix, oMacMatrix, oLabMatrix, oToolsMatrix As SAPbouiCOM.Matrix
    Private oRouteColumns, oMacColumns, oLabColumns, oToolsColumns As SAPbouiCOM.Columns
    Private oRBaseLineNoCol, oRCodeCol, oLineIdCol, oLogInstCol, oObjCol, oSeqCol, oParIdCol, oOprIdCol, oOprNameCol, oMileStoneCol, oOprTypeCol, oSubConRateCol, oQtyCol, oRInfo1Col, oRInfo2Col As SAPbouiCOM.Column
    Private oMCodeCol, oMRouteIdCol, oMOprCodeCol, oMacCodeCol, oMacNameCol, oMacGroupCol, oSetupTimeCol, oOprTimeCol, oOtherTime1Col, oOtherTime2Col, oPerQtyCol, oMInfo1Col, oMInfo2Col, oMInfo3Col, oMInfo4Col As SAPbouiCOM.Column
    Private oTCodeCol, oTRouteIdCol, oTOprCodeCol, oTMacNoCol, oToolCodeCol, oToolDescCol, oNoOfStrokesCol, oTInfo1Col, oTInfo2Col As SAPbouiCOM.Column
    Private oLCodeCol, oLRouteIdCol, oLOprCodeCol, oLMacNoCol, oSkGroupCodeCol, oSkGroupNameCol, oReqTimeCol, oReqNosCol, oLInfo1Col, oLInfo2Col As SAPbouiCOM.Column
    '**************************Folder************************************
    Private oToolsFldr, oLabFldr As SAPbouiCOM.Folder
    '**************************Items - Matrix************************************
    Private oItemCodeLink As SAPbouiCOM.LinkedButton
    '**************************Variables************************************
    Private oMacSerialNo, oToolsSerialNo, oLabSerialNo, oRouteBaseLineNo As Integer
    Private oMachineUID, oToolsUID, oLabourUID, oRouteUID As String
    Private oBoolResize As Boolean
    Private oRouteCode As String
    Private oFormName As String
    Private WithEvents oToolsClass As Tools
    Private WithEvents oSkillGroupClass As SkillGroups
    Private WithEvents oMachineClass As MachineMaster
    Private WithEvents oOperationsClass As Operations
#End Region
    ''' <summary>
    ''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmOperations.srf") method is called to load the Labour form.
    ''' Drawform() method is called to Initialize the form,Datasources and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company, Optional ByVal aRouteCode As String = Nothing, Optional ByVal aFormName As String = Nothing)
        MyBase.New(aSBO_Application, aCompany)
        SBO_Application = aSBO_Application
        oCompany = aCompany
        oRouteCode = aRouteCode
        oFormName = aFormName
        LoadFromXML("FrmRouteCard.srf")
        DrawForm()
        If oFormName = "OprRouting" Then
            oParentDB.SetValue("Code", oParentDB.Offset, oRouteCode)
            oRouteIdTxt.Value = oRouteCode
            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        End If
        oForm.DataBrowser.BrowseBy = "txtocode"
        oForm.EnableMenu("1292", True)
        oForm.EnableMenu("1293", True)
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
            oForm.Freeze(True)
            oParentDB = oForm.DataSources.DBDataSources.Item("@PSSIT_ORTE")
            oRouteDB = oForm.DataSources.DBDataSources.Item("@PSSIT_RTE4")
            oMachinesDB = oForm.DataSources.DBDataSources.Add("@PSSIT_RTE1")
            oLabourDB = oForm.DataSources.DBDataSources.Add("@PSSIT_RTE2")
            oToolsDB = oForm.DataSources.DBDataSources.Add("@PSSIT_RTE3")
            Initialize()
            AddBSDUserDataSources()
            InitializeFormComponent()
            LoadLookups()
            ConfigureRouteMatrix()
            ConfigureMachineMatrix()
            ConfigureToolsMatrix()
            ConfigureLabourMatrix()
            oForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Intitializing user table PSSIT_RTE1, PSSIT_RTE2, PSSIT_RTE3.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Initialize()
        Try
            PSSIT_RTE1 = UserTables.Item("PSSIT_RTE1")
            PSSIT_RTE2 = UserTables.Item("PSSIT_RTE2")
            PSSIT_RTE3 = UserTables.Item("PSSIT_RTE3")
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
        Dim sPath As String = IO.Directory.GetParent(Application.ExecutablePath).ToString
        Try
            '**************************Folder Initialization********************************************
            For IntICount = 1 To 2
                If IntICount = 1 Then
                    oToolsFldr = oForm.Items.Item("foltools").Specific
                    oForm.Items.Item("foltools").AffectsFormMode = False
                    oToolsFldr.DataBind.SetBound(True, "", "UFol")
                    oToolsFldr.GroupWith("follabour")
                    oToolsFldr.Select()
                ElseIf IntICount = 2 Then
                    oLabFldr = oForm.Items.Item("follabour").Specific
                    oForm.Items.Item("follabour").AffectsFormMode = False
                    oLabFldr.DataBind.SetBound(True, "", "UFol")
                    oLabFldr.GroupWith("foltools")
                End If
            Next
            '**************************Header Data******************************************
            oRouteIdTxt = oForm.Items.Item("txtocode").Specific
            oRouteIdTxt.DataBind.SetBound(True, "@PSSIT_ORTE", "Code")

            oDefRouteCheck = oForm.Items.Item("chkdefrte").Specific
            oDefRouteCheck.DataBind.SetBound(True, "@PSSIT_ORTE", "U_Defrte")

            oItemCodeTxt = oForm.Items.Item("txtitmcod").Specific
            oForm.Items.Item("txtitmcod").LinkTo = "lnkItm"
            oItemCodeTxt.DataBind.SetBound(True, "@PSSIT_ORTE", "U_Itemcode")
            oForm.Items.Add("lnkItm", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnkItm").Visible = True
            oForm.Items.Item("lnkItm").LinkTo = "txtitmcod"
            oForm.Items.Item("lnkItm").Top = 21
            oForm.Items.Item("lnkItm").Left = 104
            oForm.Items.Item("lnkItm").Description = "Link to" & vbNewLine & "Item Master"
            oItemCodeLink = oForm.Items.Item("lnkItm").Specific
            oItemCodeLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Items

            oItemBtn = oForm.Items.Item("btnitmcod").Specific
            oForm.Items.Item("btnitmcod").Description = "Choose from List" & vbNewLine & "Items List View"
            oItemBtn.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            oItemBtn.Image = sPath & "\Resources\CFL.bmp"

            oItemNameTxt = oForm.Items.Item("txtitmnam").Specific
            oItemNameTxt.DataBind.SetBound(True, "@PSSIT_ORTE", "U_Itemname")
         
            oDrawingNoTxt = oForm.Items.Item("txtdrgno").Specific
            oDrawingNoTxt.DataBind.SetBound(True, "@PSSIT_ORTE", "U_drgno")

            oRevNoTxt = oForm.Items.Item("txtrevno").Specific
            oRevNoTxt.DataBind.SetBound(True, "@PSSIT_ORTE", "U_Revno")

            oRevDateTxt = oForm.Items.Item("txtrevdt").Specific
            oRevDateTxt.DataBind.SetBound(True, "@PSSIT_ORTE", "U_Revdt")
            ' oRevDateTxt.String = System.DateTime.Today.Date.ToString("dd/MM/yyyy")

            oInfo1Txt = oForm.Items.Item("txtadnl1").Specific
            oForm.Items.Item("lbladnl1").Visible = False
            oForm.Items.Item("txtadnl1").Visible = False
            oInfo1Txt.DataBind.SetBound(True, "@PSSIT_ORTE", "U_Adnl1")

            oInfo2Txt = oForm.Items.Item("txtadnl2").Specific
            oForm.Items.Item("lbladnl2").Visible = False
            oForm.Items.Item("txtadnl2").Visible = False
            oInfo2Txt.DataBind.SetBound(True, "@PSSIT_ORTE", "U_Adnl2")

            oRemarksTxt = oForm.Items.Item("txtremark").Specific
            oRemarksTxt.DataBind.SetBound(True, "@PSSIT_ORTE", "U_Remarks")

            oActiveCheck = oForm.Items.Item("chkactive").Specific
            oActiveCheck.DataBind.SetBound(True, "@PSSIT_ORTE", "U_Active")
            oActiveCheck.Checked = True
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Creating ChooseFromList and Setting Conditions
    ''' </summary>
    ''' <remarks></remarks>
    ''' 
#Region "AddCFL Condition"
    Private Sub addCFLCondition(ByVal afrom As SAPbouiCOM.Form, ByVal strId As String)
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        Dim oCFl As SAPbouiCOM.ChooseFromList
        oCFLs = afrom.ChooseFromLists

        If strId = "ItmTxtLst" Then
            oCFl = oCFLs.Item("ItmTxtLst")
            oCons = oCFl.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "TreeType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "P"
            oCFl.SetConditions(oCons)
            oCon = oCons.Add()

        End If
    End Sub
#End Region
    Private Sub LoadLookups()
        Try
            oChItemBtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "4", "ItmBtnLst"))
            oItemBtn = oForm.Items.Item("btnitmcod").Specific
            oItemBtn.ChooseFromListUID = "ItmBtnLst"

            oChItemList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "4", "ItmTxtLst"))
            addCFLCondition(oForm, "ItmTxtLst")
            oItemCodeTxt.ChooseFromListUID = "ItmTxtLst"
            oItemCodeTxt.ChooseFromListAlias = "ItemCode"

            oChOprList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_PRN", "OprLst"))
            CreateNewConditions(oChOprList, "U_Active", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")

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
    ''' Configuring the Route Matrix items/controls in the form(.srf) by bounding to the object and to the DBDatasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ConfigureRouteMatrix()
        Try
            oRouteMatrix = oForm.Items.Item("matroutes").Specific
            oRouteMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oRouteColumns = oRouteMatrix.Columns

            oRBaseLineNoCol = oRouteColumns.Item("#")
            oRBaseLineNoCol.Editable = False
            oRBaseLineNoCol.DataBind.SetBound(True, "@PSSIT_RTE4", "LineId")

            oRCodeCol = oRouteColumns.Item("colcode")
            oRCodeCol.Editable = False
            oRCodeCol.Visible = False
            oRCodeCol.DataBind.SetBound(True, "@PSSIT_RTE4", "Code")

            'oLineIdCol = oRouteColumns.Item("collineid")
            'oLineIdCol.Editable = False
            'oLineIdCol.Visible = False
            'oLineIdCol.DataBind.SetBound(True, "@PSSIT_RTE4", "LineId")

            oObjCol = oRouteColumns.Item("colobj")
            oObjCol.Editable = False
            oObjCol.Visible = False
            oObjCol.DataBind.SetBound(True, "@PSSIT_RTE4", "Object")

            oLogInstCol = oRouteColumns.Item("colloginst")
            oLogInstCol.Editable = False
            oLogInstCol.Visible = False
            oLogInstCol.DataBind.SetBound(True, "@PSSIT_RTE4", "LogInst")

            oSeqCol = oRouteColumns.Item("colseq")
            oSeqCol.Editable = True
            oSeqCol.DataBind.SetBound(True, "@PSSIT_RTE4", "U_Seqnce")

            oParIdCol = oRouteColumns.Item("colparid")
            oParIdCol.Editable = True
            oParIdCol.DataBind.SetBound(True, "@PSSIT_RTE4", "U_Parid")

            oOprIdCol = oRouteColumns.Item("coloprcod")
            oOprIdCol.Editable = True
            oOprIdCol.DataBind.SetBound(True, "@PSSIT_RTE4", "U_Oprcode")
            oOprIdCol.ChooseFromListUID = "OprLst"
            oOprIdCol.ChooseFromListAlias = "Code"

            oOprNameCol = oRouteColumns.Item("coloprname")
            oOprNameCol.Editable = False
            oOprNameCol.DataBind.SetBound(True, "@PSSIT_RTE4", "U_Oprname")

            oMileStoneCol = oRouteColumns.Item("colmile")
            oMileStoneCol.Editable = True
            oMileStoneCol.DataBind.SetBound(True, "@PSSIT_RTE4", "U_Milestne")

            oOprTypeCol = oRouteColumns.Item("coloprtyp")
            oOprTypeCol.Editable = True
            oOprTypeCol.DataBind.SetBound(True, "@PSSIT_RTE4", "U_Oprtype")
            oOprTypeCol.ValidValues.Add("Internal", "Internal")
            oOprTypeCol.ValidValues.Add("SubContract", "SubContract")
            oOprTypeCol.ValidValues.Add("Both", "SubContract")

            oSubConRateCol = oRouteColumns.Item("colscrate")
            oSubConRateCol.Editable = True
            oSubConRateCol.DataBind.SetBound(True, "@PSSIT_RTE4", "U_SCRate")

            oQtyCol = oRouteColumns.Item("colqty")
            oQtyCol.Editable = True
            oQtyCol.DataBind.SetBound(True, "@PSSIT_RTE4", "U_Qty")

            oRInfo1Col = oRouteColumns.Item("coladnl1")
            oRInfo1Col.Editable = True
            '  oRInfo1Col.Visible = False
            oRInfo1Col.DataBind.SetBound(True, "@PSSIT_RTE4", "U_Adnl1")

            oRInfo2Col = oRouteColumns.Item("coladnl2")
            oRInfo2Col.Editable = True
            '  oRInfo2Col.Visible = False
            oRInfo2Col.DataBind.SetBound(True, "@PSSIT_RTE4", "U_Adnl2")

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
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
            oMCodeCol = oMacColumns.Item("code")
            oMCodeCol.Editable = False
            oMCodeCol.Visible = False
            oMCodeCol.DataBind.SetBound(True, "@PSSIT_RTE1", "Code")

            oMRouteIdCol = oMacColumns.Item("colrouteid")
            oMRouteIdCol.Editable = False
            oMRouteIdCol.Visible = False
            oMRouteIdCol.DataBind.SetBound(True, "@PSSIT_RTE1", "U_Rteid")

            oMOprCodeCol = oMacColumns.Item("coloprcod")
            oMOprCodeCol.Editable = False
            oMOprCodeCol.DataBind.SetBound(True, "@PSSIT_RTE1", "U_OprCode")

            oMacCodeCol = oMacColumns.Item("colwcno")
            oMacCodeCol.Editable = True
            oMacCodeCol.DataBind.SetBound(True, "@PSSIT_RTE1", "U_wcno")
            oMacCodeCol.ChooseFromListUID = "MacLst"
            oMacCodeCol.ChooseFromListAlias = "Code"

            oMacNameCol = oMacColumns.Item("colwcnam")
            oMacNameCol.Editable = False
            oMacNameCol.DataBind.SetBound(True, "@PSSIT_RTE1", "U_wcname")

            oMacGroupCol = oMacColumns.Item("colmgnam")
            oMacGroupCol.Editable = False
            oMacGroupCol.DataBind.SetBound(True, "@PSSIT_RTE1", "U_MGname")

            oSetupTimeCol = oMacColumns.Item("colSetime")
            oSetupTimeCol.Editable = True
            oSetupTimeCol.DataBind.SetBound(True, "@PSSIT_RTE1", "U_Setime")

            oOprTimeCol = oMacColumns.Item("colpertime")
            oOprTimeCol.Editable = True
            oOprTimeCol.DataBind.SetBound(True, "@PSSIT_RTE1", "U_Opertime")

            oOtherTime1Col = oMacColumns.Item("colothtim1")
            oOtherTime1Col.Editable = True
            oOtherTime1Col.DataBind.SetBound(True, "@PSSIT_RTE1", "U_Othetime1")

            oOtherTime2Col = oMacColumns.Item("colothtim2")
            oOtherTime2Col.Editable = True
            oOtherTime2Col.DataBind.SetBound(True, "@PSSIT_RTE1", "U_Othetime2")

            oPerQtyCol = oMacColumns.Item("colperqty")
            oPerQtyCol.Editable = True
            oPerQtyCol.DataBind.SetBound(True, "@PSSIT_RTE1", "U_perqty")

            oMInfo1Col = oMacColumns.Item("coladnl1")
            oMInfo1Col.Editable = True
            ' oMInfo1Col.Visible = False
            oMInfo1Col.DataBind.SetBound(True, "@PSSIT_RTE1", "U_Adnl1")

            oMInfo2Col = oMacColumns.Item("coladnl2")
            oMInfo2Col.Editable = True
            '  oMInfo2Col.Visible = False
            oMInfo2Col.DataBind.SetBound(True, "@PSSIT_RTE1", "U_Adnl2")

            oMInfo3Col = oMacColumns.Item("coladnl3")
            oMInfo3Col.Editable = True
            ' oMInfo3Col.Visible = False
            oMInfo3Col.DataBind.SetBound(True, "@PSSIT_RTE1", "U_Adnl3")

            oMInfo4Col = oMacColumns.Item("coladnl4")
            oMInfo4Col.Editable = True
            '  oMInfo4Col.Visible = False
            oMInfo4Col.DataBind.SetBound(True, "@PSSIT_RTE1", "U_Adnl4")
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
            oToolsMatrix = oForm.Items.Item("mattool").Specific
            oToolsMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oToolsColumns = oToolsMatrix.Columns

            'oToolsRowNumCol = oColumns.Item("#")
            'oToolsRowNumCol.Editable = False 

            oTCodeCol = oToolsColumns.Item("colcode")
            oTCodeCol.Editable = False
            oTCodeCol.Visible = False
            oTCodeCol.DataBind.SetBound(True, "@PSSIT_RTE3", "Code")

            oTRouteIdCol = oToolsColumns.Item("colrouteid")
            oTRouteIdCol.Editable = False
            oTRouteIdCol.Visible = False
            oTRouteIdCol.DataBind.SetBound(True, "@PSSIT_RTE3", "U_Rteid")

            oTOprCodeCol = oToolsColumns.Item("coloprcod")
            oTOprCodeCol.Editable = False
            oTOprCodeCol.Visible = False
            oTOprCodeCol.DataBind.SetBound(True, "@PSSIT_RTE3", "U_OprCode")

            oTMacNoCol = oToolsColumns.Item("colwcno")
            oTMacNoCol.Editable = False
            oTMacNoCol.DataBind.SetBound(True, "@PSSIT_RTE3", "U_wcno")

            oToolCodeCol = oToolsColumns.Item("coltolcod")
            oToolCodeCol.Editable = True
            oToolCodeCol.DataBind.SetBound(True, "@PSSIT_RTE3", "U_Toolcode")
            oToolCodeCol.ChooseFromListUID = "ToolsLst"
            oToolCodeCol.ChooseFromListAlias = "Code"

            oToolDescCol = oToolsColumns.Item("coltolnam")
            oToolDescCol.Editable = False
            oToolDescCol.DataBind.SetBound(True, "@PSSIT_RTE3", "U_TLname")

            oNoOfStrokesCol = oToolsColumns.Item("colstroke")
            oNoOfStrokesCol.Editable = True
            oNoOfStrokesCol.DataBind.SetBound(True, "@PSSIT_RTE3", "U_Strokes")

            oTInfo1Col = oToolsColumns.Item("coladnl1")
            oTInfo1Col.Editable = True
            oTInfo1Col.Visible = False
            oTInfo1Col.DataBind.SetBound(True, "@PSSIT_RTE3", "U_Adnl1")

            oTInfo2Col = oToolsColumns.Item("coladnl2")
            oTInfo2Col.Editable = True
            oTInfo2Col.Visible = False
            oTInfo2Col.DataBind.SetBound(True, "@PSSIT_RTE3", "U_Adnl2")
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

            oLCodeCol = oLabColumns.Item("colcode")
            oLCodeCol.Editable = False
            oLCodeCol.Visible = False
            oLCodeCol.DataBind.SetBound(True, "@PSSIT_RTE2", "Code")

            oLRouteIdCol = oLabColumns.Item("colrouteid")
            oLRouteIdCol.Editable = False
            oLRouteIdCol.Visible = False
            oLRouteIdCol.DataBind.SetBound(True, "@PSSIT_RTE2", "U_Rteid")

            oLOprCodeCol = oLabColumns.Item("coloprcod")
            oLOprCodeCol.Editable = False
            oLOprCodeCol.Visible = False
            oLOprCodeCol.DataBind.SetBound(True, "@PSSIT_RTE2", "U_OprCode")

            oLMacNoCol = oLabColumns.Item("colwcno")
            oLMacNoCol.Editable = False
            oLMacNoCol.DataBind.SetBound(True, "@PSSIT_RTE2", "U_wcno")

            oSkGroupCodeCol = oLabColumns.Item("colskgrp")
            oSkGroupCodeCol.Editable = True
            oSkGroupCodeCol.DataBind.SetBound(True, "@PSSIT_RTE2", "U_Skilgrp")
            oSkGroupCodeCol.ChooseFromListUID = "SkGrpLst"
            oSkGroupCodeCol.ChooseFromListAlias = "Code"

            oSkGroupNameCol = oLabColumns.Item("colgnam")
            oSkGroupNameCol.Editable = False
            oSkGroupNameCol.DataBind.SetBound(True, "@PSSIT_RTE2", "U_LGname")

            oReqTimeCol = oLabColumns.Item("colreqtime")
            oReqTimeCol.Editable = True
            oReqTimeCol.DataBind.SetBound(True, "@PSSIT_RTE2", "U_Reqtime")

            oReqNosCol = oLabColumns.Item("colreqno")
            oReqNosCol.Editable = True
            oReqNosCol.DataBind.SetBound(True, "@PSSIT_RTE2", "U_Reqno")

            oLInfo1Col = oLabColumns.Item("coladnl1")
            oLInfo1Col.Editable = True
            oLInfo1Col.Visible = False
            oLInfo1Col.DataBind.SetBound(True, "@PSSIT_RTE2", "U_Adnl1")

            oLInfo2Col = oLabColumns.Item("coladnl2")
            oLInfo2Col.Editable = True
            oLInfo2Col.Visible = False
            oLInfo2Col.DataBind.SetBound(True, "@PSSIT_RTE2", "U_Adnl2")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub SBO_Application_FormDataEvent1(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent
        If BusinessObjectInfo.FormUID = "FRC" Then
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD Then
                LoadMachineData()
                LoadToolsData()
                LoadLabourData()
            End If
        End If
    End Sub

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Dim ChooseFromListEvent As SAPbouiCOM.ChooseFromListEvent
        Try
            If pVal.FormUID = "FRC" Then
                '*****************************ChooseFromList Event is called using the raiseevent*******************************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST And pVal.BeforeAction = False Then
                    ChooseFromListEvent = pVal
                    RaiseEvent ChooseFromList(pVal.ItemUID, pVal.ColUID, pVal.Row, ChooseFromListEvent.ChooseFromListUID, ChooseFromListEvent.SelectedObjects)
                End If
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED Then
                    If pVal.BeforeAction = False Then
                        If pVal.ItemUID = "matroutes" And pVal.ColUID = "coloprcod" And pVal.Row > 0 Then
                            Dim oOprCodeEdit As SAPbouiCOM.EditText
                            Dim oCurrentRow As Integer
                            Try
                                oCurrentRow = pVal.Row
                                oRouteMatrix.GetLineData(pVal.Row)
                                oOprCodeEdit = oOprIdCol.Cells.Item(oCurrentRow).Specific
                                oOperationsClass = New Operations(SBO_Application, oCompany, oOprCodeEdit.Value, "OprRouting")
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End Try
                        End If
                        If pVal.ItemUID = "matmachine" And pVal.ColUID = "colwcno" And pVal.Row > 0 Then
                            Dim oWCCodeEdit As SAPbouiCOM.EditText
                            Dim oCurrentRow As Integer
                            Try
                                oCurrentRow = pVal.Row
                                oMacMatrix.GetLineData(pVal.Row)
                                oWCCodeEdit = oMacCodeCol.Cells.Item(oCurrentRow).Specific
                                oMachineClass = New MachineMaster(SBO_Application, oCompany, oWCCodeEdit.Value, "OprRouting")
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End Try
                        End If
                        If pVal.ItemUID = "mattool" And pVal.ColUID = "coltolcod" And pVal.Row > 0 Then
                            Dim oToolCodeEdit As SAPbouiCOM.EditText
                            Dim oCurrentRow As Integer
                            Try
                                oCurrentRow = pVal.Row
                                oToolsMatrix.GetLineData(pVal.Row)
                                oToolCodeEdit = oToolCodeCol.Cells.Item(oCurrentRow).Specific
                                oToolsClass = New Tools(SBO_Application, oCompany, oToolCodeEdit.Value, "OprRouting")
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End Try
                        End If
                        If pVal.ItemUID = "matlabour" And pVal.ColUID = "colskgrp" And pVal.Row > 0 Then
                            Dim oLabSkGrpCodeEdit As SAPbouiCOM.EditText
                            Dim oCurrentRow As Integer
                            Try
                                oCurrentRow = pVal.Row
                                oToolsMatrix.GetLineData(pVal.Row)
                                oLabSkGrpCodeEdit = oSkGroupCodeCol.Cells.Item(oCurrentRow).Specific
                                oSkillGroupClass = New SkillGroups(SBO_Application, oCompany, oLabSkGrpCodeEdit.Value, "OprRouting")
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End Try
                        End If
                    End If
                End If
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.BeforeAction = True Then
                        Try
                            '**********************Adding the child data to the database table********************
                            If pVal.ItemUID = "1" Then
                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    oForm.Freeze(True)
                                    If oMacMatrix.RowCount > 0 Then
                                        oMacMatrix.SelectRow(1, True, False)
                                    End If
                                    LoadMachineData()
                                    LoadToolsData()
                                    LoadLabourData()
                                    oForm.Freeze(False)
                                End If
                                Dim oMTransaction, oTTransaction, oLTransaction As Boolean
                                Dim IntICount, IMac, ILab, ITools As Integer
                                Dim oMCodeEdit, oMacCodeEdit, oMacNameEdit, oMacGroupEdit, oSetTimeEdit, oOprTimeEdit, oOtherTime1Edit, oOtherTime2Edit, oPerQtyEdit, oMInfo1Edit, oMInfo2Edit, oMInfo3Edit, oMInfo4Edit, oMRouteIDEdit, oMOprCodeEdit As SAPbouiCOM.EditText
                                Dim oTCodeEdit, oToolCodeEdit, oToolDescEdit, oNoOfStrokesEdit, oTInfo1Edit, oTInfo2Edit, oTRouteIDEdit, oTOprCodeEdit, oTMacCodeEdit As SAPbouiCOM.EditText
                                Dim oLCodeEdit, oSkGroupCodeEdit, oSkGroupNameEdit, oReqTimeEdit, oReqNoEdit, oLInfo1Edit, oLInfo2Edit, oLRouteIdEdit, oLOprCodeEdit, oLMacCodeEdit As SAPbouiCOM.EditText
                                Try
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        RouteDeleteEmptyRow()
                                        MachinesDeleteEmptyRow()
                                        LabourDeleteEmptyRow()
                                        ToolsDeleteEmptyRow()
                                        Validation()
                                        Try
                                            If Not oCompany.InTransaction Then
                                                oCompany.StartTransaction()
                                            End If
                                            '****************************Machine Details********************************
                                            If oMacMatrix.RowCount > 0 Then
                                                oMTransaction = True
                                                For IntICount = 1 To oMacMatrix.VisualRowCount
                                                    oMCodeEdit = oMCodeCol.Cells.Item(IntICount).Specific
                                                    oMacCodeEdit = oMacCodeCol.Cells.Item(IntICount).Specific
                                                    oMacNameEdit = oMacNameCol.Cells.Item(IntICount).Specific
                                                    oMacGroupEdit = oMacGroupCol.Cells.Item(IntICount).Specific
                                                    oSetTimeEdit = oSetupTimeCol.Cells.Item(IntICount).Specific
                                                    oOprTimeEdit = oOprTimeCol.Cells.Item(IntICount).Specific
                                                    oOtherTime1Edit = oOtherTime1Col.Cells.Item(IntICount).Specific
                                                    oOtherTime2Edit = oOtherTime2Col.Cells.Item(IntICount).Specific
                                                    oPerQtyEdit = oPerQtyCol.Cells.Item(IntICount).Specific
                                                    oMInfo1Edit = oMInfo1Col.Cells.Item(IntICount).Specific
                                                    oMInfo2Edit = oMInfo2Col.Cells.Item(IntICount).Specific
                                                    oMInfo3Edit = oMInfo3Col.Cells.Item(IntICount).Specific
                                                    oMInfo4Edit = oMInfo4Col.Cells.Item(IntICount).Specific
                                                    oMRouteIDEdit = oMRouteIdCol.Cells.Item(IntICount).Specific
                                                    oMOprCodeEdit = oMOprCodeCol.Cells.Item(IntICount).Specific
                                                    oMacMatrix.GetLineData(IntICount)
                                                    If PSSIT_RTE1.GetByKey(oMCodeEdit.Value) = True Then
                                                        PSSIT_RTE1.Code = oMCodeEdit.Value
                                                        PSSIT_RTE1.Name = oMCodeEdit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_wcno").Value = oMacCodeEdit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_wcname").Value = oMacNameEdit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_MGname").Value = oMacGroupEdit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_Setime").Value = oSetTimeEdit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_Opertime").Value = oOprTimeEdit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_Othetime1").Value = oOtherTime1Edit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_Othetime2").Value = oOtherTime2Edit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_perqty").Value = oPerQtyEdit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_Adnl1").Value = oMInfo1Edit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_Adnl2").Value = oMInfo2Edit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_Adnl3").Value = oMInfo3Edit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_Adnl4").Value = oMInfo4Edit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_Rteid").Value = oMRouteIDEdit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_OprCode").Value = oMOprCodeEdit.Value
                                                        IMac = PSSIT_RTE1.Update()
                                                    Else
                                                        PSSIT_RTE1.Code = oMCodeEdit.Value
                                                        PSSIT_RTE1.Name = oMCodeEdit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_wcno").Value = oMacCodeEdit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_wcname").Value = oMacNameEdit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_MGname").Value = oMacGroupEdit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_Setime").Value = oSetTimeEdit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_Opertime").Value = oOprTimeEdit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_Othetime1").Value = oOtherTime1Edit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_Othetime2").Value = oOtherTime2Edit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_perqty").Value = oPerQtyEdit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_Adnl1").Value = oMInfo1Edit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_Adnl2").Value = oMInfo2Edit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_Adnl3").Value = oMInfo3Edit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_Adnl4").Value = oMInfo4Edit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_Rteid").Value = oMRouteIDEdit.Value
                                                        PSSIT_RTE1.UserFields.Fields.Item("U_OprCode").Value = oMOprCodeEdit.Value
                                                        IMac = PSSIT_RTE1.Add()
                                                    End If
                                                Next
                                            End If
                                            '****************************Tool Details********************************
                                            If oToolsMatrix.RowCount > 0 Then
                                                oTTransaction = True
                                                For IntICount = 1 To oToolsMatrix.VisualRowCount
                                                    oTCodeEdit = oTCodeCol.Cells.Item(IntICount).Specific
                                                    oToolCodeEdit = oToolCodeCol.Cells.Item(IntICount).Specific
                                                    oToolDescEdit = oToolDescCol.Cells.Item(IntICount).Specific
                                                    oNoOfStrokesEdit = oNoOfStrokesCol.Cells.Item(IntICount).Specific
                                                    oTInfo1Edit = oTInfo1Col.Cells.Item(IntICount).Specific
                                                    oTInfo2Edit = oTInfo2Col.Cells.Item(IntICount).Specific
                                                    oTRouteIDEdit = oTRouteIdCol.Cells.Item(IntICount).Specific
                                                    oTOprCodeEdit = oTOprCodeCol.Cells.Item(IntICount).Specific
                                                    oTMacCodeEdit = oTMacNoCol.Cells.Item(IntICount).Specific
                                                    oToolsMatrix.GetLineData(IntICount)
                                                    If PSSIT_RTE3.GetByKey(oTCodeEdit.Value) = True Then
                                                        PSSIT_RTE3.Code = oTCodeEdit.Value
                                                        PSSIT_RTE3.Name = oTCodeEdit.Value
                                                        PSSIT_RTE3.UserFields.Fields.Item("U_Toolcode").Value = oToolCodeEdit.Value
                                                        PSSIT_RTE3.UserFields.Fields.Item("U_TLname").Value = oToolDescEdit.Value
                                                        PSSIT_RTE3.UserFields.Fields.Item("U_Strokes").Value = oNoOfStrokesEdit.Value
                                                        PSSIT_RTE3.UserFields.Fields.Item("U_Adnl1").Value = oTInfo1Edit.Value
                                                        PSSIT_RTE3.UserFields.Fields.Item("U_Adnl2").Value = oTInfo2Edit.Value
                                                        PSSIT_RTE3.UserFields.Fields.Item("U_Rteid").Value = oTRouteIDEdit.Value
                                                        PSSIT_RTE3.UserFields.Fields.Item("U_OprCode").Value = oTOprCodeEdit.Value
                                                        PSSIT_RTE3.UserFields.Fields.Item("U_wcno").Value = oTMacCodeEdit.Value
                                                        ITools = PSSIT_RTE3.Update()
                                                    Else
                                                        PSSIT_RTE3.Code = oTCodeEdit.Value
                                                        PSSIT_RTE3.Name = oTCodeEdit.Value
                                                        PSSIT_RTE3.UserFields.Fields.Item("U_Toolcode").Value = oToolCodeEdit.Value
                                                        PSSIT_RTE3.UserFields.Fields.Item("U_TLname").Value = oToolDescEdit.Value
                                                        PSSIT_RTE3.UserFields.Fields.Item("U_Strokes").Value = oNoOfStrokesEdit.Value
                                                        PSSIT_RTE3.UserFields.Fields.Item("U_Adnl1").Value = oTInfo1Edit.Value
                                                        PSSIT_RTE3.UserFields.Fields.Item("U_Adnl2").Value = oTInfo2Edit.Value
                                                        PSSIT_RTE3.UserFields.Fields.Item("U_Rteid").Value = oTRouteIDEdit.Value
                                                        PSSIT_RTE3.UserFields.Fields.Item("U_OprCode").Value = oTOprCodeEdit.Value
                                                        PSSIT_RTE3.UserFields.Fields.Item("U_wcno").Value = oTMacCodeEdit.Value
                                                        ITools = PSSIT_RTE3.Add()
                                                    End If
                                                Next
                                            End If
                                            '****************************Labour Details********************************
                                            If oLabMatrix.RowCount > 0 Then
                                                oLTransaction = True
                                                For IntICount = 1 To oLabMatrix.VisualRowCount
                                                    oLCodeEdit = oLCodeCol.Cells.Item(IntICount).Specific
                                                    oSkGroupCodeEdit = oSkGroupCodeCol.Cells.Item(IntICount).Specific
                                                    oSkGroupNameEdit = oSkGroupNameCol.Cells.Item(IntICount).Specific
                                                    oReqTimeEdit = oReqTimeCol.Cells.Item(IntICount).Specific
                                                    oReqNoEdit = oReqNosCol.Cells.Item(IntICount).Specific
                                                    oLInfo1Edit = oLInfo1Col.Cells.Item(IntICount).Specific
                                                    oLInfo2Edit = oLInfo2Col.Cells.Item(IntICount).Specific
                                                    oLRouteIdEdit = oLRouteIdCol.Cells.Item(IntICount).Specific
                                                    oLOprCodeEdit = oLOprCodeCol.Cells.Item(IntICount).Specific
                                                    oLMacCodeEdit = oLMacNoCol.Cells.Item(IntICount).Specific
                                                    oLabMatrix.GetLineData(IntICount)
                                                    If PSSIT_RTE2.GetByKey(oLCodeEdit.Value) = True Then
                                                        PSSIT_RTE2.Code = oLCodeEdit.Value
                                                        PSSIT_RTE2.Name = oLCodeEdit.Value
                                                        PSSIT_RTE2.UserFields.Fields.Item("U_Skilgrp").Value = oSkGroupCodeEdit.Value
                                                        PSSIT_RTE2.UserFields.Fields.Item("U_LGname").Value = oSkGroupNameEdit.Value
                                                        PSSIT_RTE2.UserFields.Fields.Item("U_Reqtime").Value = oReqTimeEdit.Value
                                                        PSSIT_RTE2.UserFields.Fields.Item("U_Reqno").Value = oReqNoEdit.Value
                                                        PSSIT_RTE2.UserFields.Fields.Item("U_Adnl1").Value = oLInfo1Edit.Value
                                                        PSSIT_RTE2.UserFields.Fields.Item("U_Adnl2").Value = oLInfo2Edit.Value
                                                        PSSIT_RTE2.UserFields.Fields.Item("U_Rteid").Value = oLRouteIdEdit.Value
                                                        PSSIT_RTE2.UserFields.Fields.Item("U_OprCode").Value = oLOprCodeEdit.Value
                                                        PSSIT_RTE2.UserFields.Fields.Item("U_wcno").Value = oLMacCodeEdit.Value
                                                        ILab = PSSIT_RTE2.Update()
                                                    Else
                                                        PSSIT_RTE2.Code = oLCodeEdit.Value
                                                        PSSIT_RTE2.Name = oLCodeEdit.Value
                                                        PSSIT_RTE2.UserFields.Fields.Item("U_Skilgrp").Value = oSkGroupCodeEdit.Value
                                                        PSSIT_RTE2.UserFields.Fields.Item("U_LGname").Value = oSkGroupNameEdit.Value
                                                        PSSIT_RTE2.UserFields.Fields.Item("U_Reqtime").Value = oReqTimeEdit.Value
                                                        PSSIT_RTE2.UserFields.Fields.Item("U_Reqno").Value = oReqNoEdit.Value
                                                        PSSIT_RTE2.UserFields.Fields.Item("U_Adnl1").Value = oLInfo1Edit.Value
                                                        PSSIT_RTE2.UserFields.Fields.Item("U_Adnl2").Value = oLInfo2Edit.Value
                                                        PSSIT_RTE2.UserFields.Fields.Item("U_Rteid").Value = oLRouteIdEdit.Value
                                                        PSSIT_RTE2.UserFields.Fields.Item("U_OprCode").Value = oLOprCodeEdit.Value
                                                        PSSIT_RTE2.UserFields.Fields.Item("U_wcno").Value = oLMacCodeEdit.Value
                                                        ILab = PSSIT_RTE2.Add()
                                                    End If
                                                Next
                                            End If
                                            If oMTransaction = True Or oTTransaction = True Or oLTransaction = True Then
                                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                            End If
                                        Catch ex As Exception
                                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            BubbleEvent = False
                                        Finally
                                            If oMTransaction = False And oTTransaction = False And oLTransaction = False Then
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
                    If pVal.BeforeAction = False Then
                        Try
                            '**********************Refreshing the form to initiate default values********************
                            If pVal.ItemUID = "1" Then
                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oForm.Freeze(True)
                                    SetItemEnabled()
                                    If oRouteMatrix.RowCount > 0 Then
                                        oRouteMatrix.SelectRow(1, True, False)
                                    End If
                                    If oMacMatrix.RowCount > 0 Then
                                        oMacMatrix.SelectRow(1, True, False)
                                    End If
                                    oForm.Freeze(False)
                                End If
                                If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    oForm.Refresh()
                                    oForm.Freeze(True)
                                    SetItemEnabled()
                                    oRouteIdTxt.Active = True
                                    oRevDateTxt.String = "t" 'System.DateTime.Today.Date.ToString("dd/MM/yyyy")
                                    SBO_Application.SendKeys("{TAB}")
                                    oActiveCheck.Checked = True
                                    oForm.Freeze(False)
                                End If
                            End If
                            '**********************Setting the pane level as per the folder selected********************
                            oForm.Freeze(True)
                            If pVal.ItemUID = "foltools" Then
                                oForm.PaneLevel = 1
                                SetItemEnabled()
                            End If
                            If pVal.ItemUID = "follabour" Then
                                oForm.PaneLevel = 2
                                SetItemEnabled()
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
                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                If oRouteIdTxt.Value <> "" Then
                                    If oRouteIdTxt.Value.Length > 0 Then
                                        oStrSql = "Select * from [@PSSIT_ORTE] where Code = '" & oRouteIdTxt.Value & "'"
                                        oRs.DoQuery(oStrSql)
                                        If oRs.RecordCount > 0 Then
                                            SBO_Application.SetStatusBarMessage("Route Id '" & oRouteIdTxt.Value & "' already exists", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            BubbleEvent = False
                                        End If
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
                    '**********************Adding a row to the Route matrix********************
                    If pVal.BeforeAction = False Then
                        Try
                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                If oRouteIdTxt.Value <> "" Then
                                    If oRouteIdTxt.Value.Length > 0 Then
                                        oRouteDB.InsertRecord(oRouteDB.Size)
                                        oRouteDB.Offset = oRouteDB.Size - 1
                                        oRouteMatrix.Clear()
                                        SetRouteDefaultValue()
                                        oRouteMatrix.AddRow(1, oRouteMatrix.RowCount)
                                        'AddBaseLineNo()
                                        oRouteDB.SetValue("LineId", oRouteDB.Offset, oRouteMatrix.RowCount)
                                        oRouteMatrix.SetLineData(oRouteMatrix.RowCount)
                                        oRouteMatrix.FlushToDataSource()
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        End Try
                    End If
                End If
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN And pVal.CharPressed <> Keys.Tab Then
                    If pVal.BeforeAction = True Then
                        Try
                            If pVal.ItemUID = "matroutes" And pVal.Row > 0 Then
                                If pVal.ColUID = "colscrate" Or pVal.ColUID = "colqty" Then
                                    Dim oOprTypeCombo As SAPbouiCOM.ComboBox
                                    Dim oCurrentRow As Integer
                                    Try
                                        oCurrentRow = pVal.Row
                                        oOprTypeCombo = oOprTypeCol.Cells.Item(oCurrentRow).Specific
                                        oRouteMatrix.GetLineData(pVal.Row)
                                        If oRouteDB.GetValue("U_Oprtype", oRouteDB.Offset).Trim().Length > 0 Then
                                            If oOprTypeCombo.Selected.Value = "Internal" Then
                                                '  BubbleEvent = False
                                            Else
                                                BubbleEvent = True
                                            End If
                                        End If
                                    Catch ex As Exception
                                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        BubbleEvent = False
                                    End Try
                                End If
                            End If
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End Try
                    End If
                End If
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                    If pVal.BeforeAction = True Then
                        Try
                            If pVal.ItemUID = "matroutes" And pVal.Row > 0 Then
                                If pVal.ColUID = "colscrate" Or pVal.ColUID = "colqty" Then
                                    Dim oOprTypeCombo As SAPbouiCOM.ComboBox
                                    Dim oCurrentRow As Integer
                                    Try
                                        oCurrentRow = pVal.Row
                                        oOprTypeCombo = oOprTypeCol.Cells.Item(oCurrentRow).Specific
                                        oRouteMatrix.GetLineData(pVal.Row)
                                        If oRouteDB.GetValue("U_Oprtype", oRouteDB.Offset).Trim().Length > 0 Then
                                            If oOprTypeCombo.Selected.Value = "Internal" Then
                                                BubbleEvent = False
                                            Else
                                                BubbleEvent = True
                                            End If
                                        End If
                                    Catch ex As Exception
                                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        BubbleEvent = False
                                    End Try
                                End If
                            End If
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End Try
                    End If
                    If pVal.BeforeAction = False Then
                        Try
                            If pVal.ItemUID = "matroutes" And pVal.Row > 0 Then
                                If (pVal.ColUID = "#" Or pVal.ColUID = "colseq" Or pVal.ColUID = "colparid" Or pVal.ColUID = "coloprcod" Or pVal.ColUID = "coloprname" Or pVal.ColUID = "colmile" Or pVal.ColUID = "coloprtyp" Or pVal.ColUID = "colscrate" Or pVal.ColUID = "colqty" Or pVal.ColUID = "coladnl1" Or pVal.ColUID = "coladnl2") Then
                                    oRouteMatrix.SelectRow(pVal.Row, True, False)
                                End If
                            End If
                            If (pVal.ItemUID = "matroutes") And pVal.ColUID = "#" Then
                                Try
                                    oRouteUID = pVal.ItemUID
                                    oToolsUID = ""
                                    oLabourUID = ""
                                    oMachineUID = ""
                                Catch ex As Exception
                                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                End Try
                            End If
                            If (pVal.ItemUID = "matmachine") And pVal.ColUID = "#" Then
                                Try
                                    oRouteUID = ""
                                    oToolsUID = ""
                                    oLabourUID = ""
                                    oMachineUID = pVal.ItemUID
                                Catch ex As Exception
                                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                End Try
                            End If
                            If (pVal.ItemUID = "mattool") And pVal.ColUID = "#" Then
                                Try
                                    oRouteUID = ""
                                    oMachineUID = ""
                                    oLabourUID = ""
                                    oToolsUID = pVal.ItemUID
                                Catch ex As Exception
                                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                End Try
                            End If
                            If (pVal.ItemUID = "matlabour") And pVal.ColUID = "#" Then
                                Try
                                    oRouteUID = ""
                                    oMachineUID = ""
                                    oToolsUID = ""
                                    oLabourUID = pVal.ItemUID
                                Catch ex As Exception
                                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                End Try
                            End If
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        End Try
                    End If
                End If
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_RESIZE And pVal.BeforeAction = False Then
                    Try
                        '  Form_Resize()
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
    Private Sub Validation()
        Dim IntICount As Integer
        Try
            If oRouteIdTxt.Value.Length = 0 Then
                oRouteIdTxt.Active = True
                Throw New Exception("Operation Id should not be empty")
            End If
            If oItemCodeTxt.Value.Length = 0 Then
                oItemCodeTxt.Active = True
                Throw New Exception("ItemCode Should not be empty")
            End If
            If oRouteMatrix.RowCount = 0 Then
                Throw New Exception("Route Details should be entered")
            ElseIf oRouteMatrix.VisualRowCount >= 1 Then
                For IntICount = 1 To oRouteMatrix.VisualRowCount
                    oRouteMatrix.GetLineData(IntICount)
                    If oSeqCol.Cells.Item(IntICount).Specific.Value = "0" Or oSeqCol.Cells.Item(IntICount).Specific.Value = "" Or oSeqCol.Cells.Item(IntICount).Specific.Value.Length = 0 Then
                        oSeqCol.Cells.Item(IntICount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Throw New Exception("Sequence should be greater than zero")
                    End If
                    If oOprIdCol.Cells.Item(IntICount).Specific.Value.Length = 0 Then
                        oOprIdCol.Cells.Item(IntICount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Throw New Exception("Operation Code should not be empty")
                    End If
                Next
            End If
            If oMacMatrix.RowCount = 0 Then
                Throw New Exception("Atleast one Machine should be added")
            ElseIf oMacMatrix.VisualRowCount >= 1 Then
                For IntICount = 1 To oMacMatrix.VisualRowCount
                    oMacMatrix.GetLineData(IntICount)
                    If oMacCodeCol.Cells.Item(IntICount).Specific.Value.Length = 0 Or oMacNameCol.Cells.Item(IntICount).Specific.Value.Length = 0 Then
                        oMacCodeCol.Cells.Item(IntICount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Throw New Exception("Machine Details should be entered")
                    End If
                    If oOprTimeCol.Cells.Item(IntICount).Specific.Value.Length = 0 Then
                        oOprTimeCol.Cells.Item(IntICount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Throw New Exception("Operation Time should not be empty")
                    End If
                    If oPerQtyCol.Cells.Item(IntICount).Specific.Value = 0 Then
                        oPerQtyCol.Cells.Item(IntICount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Throw New Exception("Per Qty should be greater than zero")
                    End If
                Next
            End If
            If oToolsMatrix.VisualRowCount > 0 Then
                For IntICount = 1 To oToolsMatrix.VisualRowCount
                    If oToolCodeCol.Cells.Item(IntICount).Specific.Value.Length = 0 Or oToolDescCol.Cells.Item(IntICount).Specific.Value.Length = 0 Then
                        oToolCodeCol.Cells.Item(IntICount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Throw New Exception("Tool Details should be entered")
                    End If
                    If oToolCodeCol.Cells.Item(IntICount).Specific.Value.Length > 0 Then
                        If oNoOfStrokesCol.Cells.Item(IntICount).Specific.Value = "0" Or oNoOfStrokesCol.Cells.Item(IntICount).Specific.Value = "" Or oNoOfStrokesCol.Cells.Item(IntICount).Specific.Value.Length = 0 Then
                            oNoOfStrokesCol.Cells.Item(IntICount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Throw New Exception("No Of Strokes should be entered")
                        End If
                    End If
                Next
            End If
            If oLabMatrix.RowCount = 0 Then
                oLabFldr.Select()
                Throw New Exception("Atleast one Labour should be added")
            ElseIf oLabMatrix.VisualRowCount >= 1 Then
                For IntICount = 1 To oLabMatrix.VisualRowCount
                    oLabMatrix.GetLineData(IntICount)
                    If oSkGroupCodeCol.Cells.Item(IntICount).Specific.Value.Length = 0 And oSkGroupNameCol.Cells.Item(IntICount).Specific.Value.Length = 0 Then
                        oLabFldr.Select()
                        oSkGroupCodeCol.Cells.Item(IntICount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Throw New Exception("Labour Details should be entered")
                    End If
                    If oReqTimeCol.Cells.Item(IntICount).Specific.Value.Length = 0 Then
                        oLabFldr.Select()
                        oReqTimeCol.Cells.Item(IntICount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Throw New Exception("Required Time should not be empty")
                    End If
                    If oReqNosCol.Cells.Item(IntICount).Specific.Value = 0 Then
                        oLabFldr.Select()
                        oReqNosCol.Cells.Item(IntICount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Throw New Exception("Required Persons should be greater than zero")
                    End If
                    'For IntKCount = 1 To oMacMatrix.RowCount
                    '    oMacMatrix.GetLineData(IntKCount)
                    '    oMacCode = oMacCodeCol.Cells.Item(IntKCount).Specific.Value
                    '    For IntJCount = 1 To oLabMatrix.RowCount
                    '        oLabMatrix.GetLineData(IntJCount)
                    '        If oMacCode <> oLMacNoCol.Cells.Item(IntJCount).Specific.Value Then
                    '            Throw New Exception("Atleast one labour should be added for each machine")
                    '        End If
                    '    Next
                    'Next
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Dim FormID As String
        Dim IntICount, IntJCount, IntKCount As Integer
        Dim oMacDelCode, oToolsDelCode, oLabDelCode As SAPbouiCOM.EditText
        Dim oMOprCode, oMMacCode As String
        Try
            FormID = SBO_Application.Forms.ActiveForm.UniqueID
            If pVal.MenuUID = "1281" And FormID = "FRC" Then
                If pVal.BeforeAction = False Then
                    SetItemEnabled()
                    oRouteIdTxt.Active = True
                End If
            End If
            If pVal.MenuUID = "1282" And pVal.BeforeAction = False And FormID = "FRC" Then
                SetItemEnabled()
                oForm.Freeze(True)
                oRouteIdTxt.Active = True
                oRevDateTxt.String = "t" 'System.DateTime.Today.Date.ToString("dd/MM/yyyy")
                SBO_Application.SendKeys("{TAB}")
                oActiveCheck.Checked = True
                oForm.Freeze(False)
            End If
            '*****************************Adding a row to the Route Matrix*******************************
            If pVal.MenuUID = "1292" And pVal.BeforeAction = True And FormID = "FRC" Then
                If oRouteUID = "matroutes" Then
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                    If oRouteMatrix.RowCount = 0 Then
                        If oRouteIdTxt.Value <> "" Then
                            If oRouteIdTxt.Value.Length > 0 Then
                                oRouteDB.InsertRecord(oRouteDB.Size)
                                oRouteDB.Offset = oRouteDB.Size - 1
                                oRouteMatrix.Clear()
                                SetRouteDefaultValue()
                                oRouteMatrix.FlushToDataSource()
                                oRouteMatrix.AddRow(1, oRouteMatrix.RowCount)
                                oRouteDB.SetValue("LineId", oRouteDB.Offset, oRouteMatrix.RowCount)
                                'AddBaseLineNo()
                                'oRouteDB.SetValue("U_Bselino", oRouteDB.Offset, oRouteBaseLineNo)
                                oRouteMatrix.SetLineData(oRouteMatrix.RowCount)
                            End If
                        End If
                    ElseIf oRouteMatrix.RowCount > 0 Then
                        If oRouteIdTxt.Value <> "" Then
                            If oRouteIdTxt.Value.Length > 0 Then
                                If oRouteMatrix.Columns.Item("colseq").Cells.Item(oRouteMatrix.RowCount).Specific.value <> "" Then
                                    oRouteDB.Offset = oRouteDB.Size - 1
                                    SetRouteDefaultValue()
                                    oRouteDB.SetValue("U_Seqnce", oRouteDB.Offset, "")
                                    oRouteDB.SetValue("U_Parid", oRouteDB.Offset, "")

                                    oRouteMatrix.AddRow(1, oRouteMatrix.RowCount)

                                    oRouteDB.SetValue("LineId", oRouteDB.Offset, oRouteMatrix.RowCount)
                                    oRouteMatrix.SetLineData(oRouteMatrix.RowCount)
                                End If



                                'AddBaseLineNo()
                                'oRouteDB.SetValue("U_Bselino", oRouteDB.Offset, oRouteBaseLineNo)
                                'If oRouteMatrix.Columns.Item("colseq").Cells.Item(oRouteMatrix.RowCount).Specific.value <> "" Then
                                '    ' oRouteMatrix.SetLineData(oRouteMatrix.RowCount)
                                'End If

                            End If
                        End If
                    End If
                End If
                '*****************************Adding Rows to Machine Matrix*******************************
                If oMachineUID = "matmachine" Then
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                    If oMacMatrix.RowCount = 0 Then
                        If oRouteMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder) > 0 Then
                            oMachinesDB.InsertRecord(oMachinesDB.Size)
                            AddMachineRow(oOprIdCol.Cells.Item(oRouteMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                            oMachinesDB.SetValue("Code", oMachinesDB.Offset, oMacSerialNo)
                            oMacMatrix.SetLineData(oMacMatrix.RowCount)
                        Else
                            SBO_Application.SetStatusBarMessage("Select the Operation for which the machine to be added", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                    ElseIf oMacMatrix.RowCount > 0 Then
                        If oRouteMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder) > 0 Then
                            If Len(oMacCodeCol.Cells.Item(oMacMatrix.RowCount).Specific.value) <= 0 Then
                                SBO_Application.SetStatusBarMessage("Machine Details should not be empty", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                            End If
                            If Len(oMacCodeCol.Cells.Item(oMacMatrix.RowCount).Specific.value) > 0 Then
                                AddMachineRow(oOprIdCol.Cells.Item(oRouteMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                                oMachinesDB.SetValue("Code", oMachinesDB.Offset, oMacSerialNo)
                                oMacMatrix.SetLineData(oMacMatrix.RowCount)
                            End If
                        Else
                            SBO_Application.SetStatusBarMessage("Select the Operation for which the machine to be added", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                    End If
                End If
                '*****************************Adding Rows to Tools Matrix*******************************
                If oToolsUID = "mattool" Then
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                    If oToolsMatrix.RowCount = 0 Then
                        If oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder) > 0 Then
                            oToolsDB.InsertRecord(oToolsDB.Size)
                            AddToolsRow(oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value, _
                            oMOprCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                            oToolsDB.SetValue("Code", oToolsDB.Offset, oToolsSerialNo)
                            oToolsMatrix.SetLineData(oToolsMatrix.RowCount)
                        Else
                            SBO_Application.SetStatusBarMessage("Select the machine for which the tools to be added", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                    ElseIf oToolsMatrix.RowCount > 0 Then
                        If oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder) > 0 Then
                            If Len(oToolCodeCol.Cells.Item(oToolsMatrix.RowCount).Specific.value) <= 0 Then
                                SBO_Application.SetStatusBarMessage("Tool Details should not be empty", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                            End If
                            If Len(oToolCodeCol.Cells.Item(oToolsMatrix.RowCount).Specific.value) > 0 Then
                                AddToolsRow(oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value, _
                                oMOprCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                                oToolsDB.SetValue("Code", oToolsDB.Offset, oToolsSerialNo)
                                oToolsMatrix.SetLineData(oToolsMatrix.RowCount)
                            End If
                        End If
                    End If
                End If
                '*****************************Adding Rows to Labour Matrix*******************************
                If oLabourUID = "matlabour" Then
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                    If oLabMatrix.RowCount = 0 Then
                        If oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder) > 0 Then
                            oLabourDB.InsertRecord(oLabourDB.Size)
                            AddLabourRow(oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value, _
                            oMOprCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                            oLabourDB.SetValue("Code", oLabourDB.Offset, oLabSerialNo)
                            oLabMatrix.SetLineData(oLabMatrix.RowCount)
                        Else
                            SBO_Application.SetStatusBarMessage("Select the machine for which the labour to be added", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                    ElseIf oLabMatrix.RowCount > 0 Then
                        If oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder) > 0 Then
                            If Len(oSkGroupCodeCol.Cells.Item(oLabMatrix.RowCount).Specific.value) <= 0 Then
                                SBO_Application.SetStatusBarMessage("Labour Details should not be empty", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                            End If
                            If Len(oSkGroupCodeCol.Cells.Item(oLabMatrix.RowCount).Specific.value) > 0 Then
                                AddLabourRow(oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value, _
                                oMOprCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                                oLabourDB.SetValue("Code", oLabourDB.Offset, oLabSerialNo)
                                oLabMatrix.SetLineData(oLabMatrix.RowCount)
                            Else
                                SBO_Application.SetStatusBarMessage("Select the machine for which the labour to be added", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                            End If
                            '------------------
                            'If Len(oSkGroupCodeCol.Cells.Item(oLabMatrix.RowCount).Specific.value) > 0 Then
                            '    oLabourDB.InsertRecord(oLabourDB.Size)
                            '    oLabourDB.Offset = oLabourDB.Size - 1
                            '    SetLabourDefaultValue()
                            '    If oLabMatrix.Columns.Item("colskgrp").Cells.Item(oLabMatrix.RowCount).Specific.value <> "" Then
                            '        oLabMatrix.AddRow(1, oLabMatrix.RowCount)
                            '    End If

                            'End If
                            '-------------
                        End If
                    End If
                End If
            End If
            If pVal.MenuUID = "1293" And pVal.BeforeAction = True And FormID = "FRC" Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End If
                If oRouteUID = "matroutes" Then
                    If oRouteMatrix.RowCount > 0 Then
                        Dim oMOperationID As String = oOprIdCol.Cells.Item(oRouteMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value
                        For IntICount = oMacMatrix.RowCount To 1 Step -1
                            oMacMatrix.GetLineData(IntICount)
                            If oMOperationID = oMOprCodeCol.Cells.Item(IntICount).Specific.Value Then
                                oMOprCode = oMOprCodeCol.Cells.Item(IntICount).Specific.Value
                                oMMacCode = oMacCodeCol.Cells.Item(IntICount).Specific.Value
                                For IntJCount = oToolsMatrix.RowCount To 1 Step -1
                                    oToolsMatrix.GetLineData(IntJCount)
                                    If oMOprCode = oTOprCodeCol.Cells.Item(IntJCount).Specific.Value And oMMacCode = oTMacNoCol.Cells.Item(IntJCount).Specific.Value Then
                                        oToolsDelCode = oTCodeCol.Cells.Item(IntJCount).Specific
                                        If PSSIT_RTE3.GetByKey(oToolsDelCode.Value) = True Then
                                            Dim I As Integer = PSSIT_RTE3.Remove()
                                            oToolsMatrix.DeleteRow(IntJCount)
                                            oToolsMatrix.FlushToDataSource()
                                        ElseIf PSSIT_RTE3.GetByKey(oToolsDelCode.Value) = False Then
                                            oToolsMatrix.DeleteRow(IntJCount)
                                            oToolsMatrix.FlushToDataSource()
                                            oToolsMatrix.LoadFromDataSource()
                                        End If
                                    End If
                                Next
                                For IntKCount = oLabMatrix.RowCount To 1 Step -1
                                    oLabMatrix.GetLineData(IntKCount)
                                    If oMOprCode = oLOprCodeCol.Cells.Item(IntKCount).Specific.Value And oMMacCode = oLMacNoCol.Cells.Item(IntKCount).Specific.Value Then
                                        oLabDelCode = oLCodeCol.Cells.Item(IntKCount).Specific
                                        If PSSIT_RTE2.GetByKey(oLabDelCode.Value) = True Then
                                            Dim I As Integer = PSSIT_RTE2.Remove()
                                            oLabMatrix.DeleteRow(IntKCount)
                                            oLabMatrix.FlushToDataSource()
                                        ElseIf PSSIT_RTE2.GetByKey(oLabDelCode.Value) = False Then
                                            oLabMatrix.DeleteRow(IntKCount)
                                            oLabMatrix.FlushToDataSource()
                                            oLabMatrix.LoadFromDataSource()
                                        End If
                                    End If
                                Next
                                oMacDelCode = oMCodeCol.Cells.Item(IntICount).Specific
                                If PSSIT_RTE1.GetByKey(oMacDelCode.Value) = True Then
                                    Dim I As Integer = PSSIT_RTE1.Remove()
                                    oMacMatrix.DeleteRow(IntICount)
                                    oMacMatrix.FlushToDataSource()
                                ElseIf PSSIT_RTE1.GetByKey(oMacDelCode.Value) = False Then
                                    oMacMatrix.DeleteRow(IntICount)
                                    oMacMatrix.FlushToDataSource()
                                    oMacMatrix.LoadFromDataSource()
                                End If
                            End If
                        Next
                        oRouteMatrix.DeleteRow(oRouteMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                        oRouteMatrix.FlushToDataSource()
                        'oRouteMatrix.LoadFromDataSource()
                        For IntICount = 1 To oRouteMatrix.VisualRowCount
                            oRouteMatrix.GetLineData(IntICount)
                            oForm.Freeze(True)
                            oRouteDB.SetValue("LineId", oRouteDB.Offset, IntICount)
                            oRouteMatrix.SetLineData(IntICount)
                            oForm.Freeze(False)
                        Next
                    End If
                End If
                If oMachineUID = "matmachine" Then
                    If oMacMatrix.RowCount > 0 Then
                        oMacMatrix.GetLineData(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                        oMOprCode = oMOprCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value
                        oMMacCode = oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value
                        For IntJCount = oToolsMatrix.RowCount To 1 Step -1
                            oToolsMatrix.GetLineData(IntJCount)
                            If oMOprCode = oTOprCodeCol.Cells.Item(IntJCount).Specific.Value And oMMacCode = oTMacNoCol.Cells.Item(IntJCount).Specific.Value Then
                                oToolsDelCode = oTCodeCol.Cells.Item(IntJCount).Specific
                                If PSSIT_RTE3.GetByKey(oToolsDelCode.Value) = True Then
                                    Dim I As Integer = PSSIT_RTE3.Remove()
                                    oToolsMatrix.DeleteRow(IntJCount)
                                    oToolsMatrix.FlushToDataSource()
                                Else
                                    oToolsMatrix.DeleteRow(IntJCount)
                                    oToolsMatrix.FlushToDataSource()
                                End If
                            End If
                        Next
                        For IntKCount = oLabMatrix.RowCount To 1 Step -1
                            oLabMatrix.GetLineData(IntKCount)
                            If oMOprCode = oLOprCodeCol.Cells.Item(IntKCount).Specific.Value And oMMacCode = oLMacNoCol.Cells.Item(IntKCount).Specific.Value Then
                                oLabDelCode = oLCodeCol.Cells.Item(IntKCount).Specific
                                If PSSIT_RTE2.GetByKey(oLabDelCode.Value) = True Then
                                    Dim I As Integer = PSSIT_RTE2.Remove()
                                    oLabMatrix.DeleteRow(IntKCount)
                                    oLabMatrix.FlushToDataSource()
                                Else
                                    oLabMatrix.DeleteRow(IntKCount)
                                    oLabMatrix.FlushToDataSource()
                                End If
                            End If
                        Next
                        oMacDelCode = oMCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific
                        If PSSIT_RTE1.GetByKey(oMacDelCode.Value) = True Then
                            Dim I As Integer = PSSIT_RTE1.Remove()
                            oMacMatrix.DeleteRow(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                            oMacMatrix.FlushToDataSource()
                        Else
                            oMacMatrix.DeleteRow(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                            oMacMatrix.FlushToDataSource()
                        End If
                    End If
                End If
                If oToolsUID = "mattool" Then
                    oToolsDelCode = oTCodeCol.Cells.Item(oToolsMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific
                    If PSSIT_RTE3.GetByKey(oToolsDelCode.Value) = True Then
                        Dim I As Integer = PSSIT_RTE3.Remove()
                        oToolsMatrix.DeleteRow(oToolsMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                        oToolsMatrix.FlushToDataSource()
                    Else
                        oToolsMatrix.DeleteRow(oToolsMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                        oToolsMatrix.FlushToDataSource()
                    End If
                End If
                If oLabourUID = "matlabour" Then
                    oLabDelCode = oLCodeCol.Cells.Item(oLabMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific
                    If PSSIT_RTE2.GetByKey(oLabDelCode.Value) = True Then
                        Dim I As Integer = PSSIT_RTE2.Remove()
                        oLabMatrix.DeleteRow(oLabMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                        oLabMatrix.FlushToDataSource()
                    Else
                        oLabMatrix.DeleteRow(oLabMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                        oLabMatrix.FlushToDataSource()
                    End If
                End If
                BubbleEvent = False
            End If
            '*****************************LoadToolsData() is called.*******************************
            If (pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") And pVal.BeforeAction = False And FormID = "FRC" Then
                Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Try
                    oRs.DoQuery("Select * from [@PSSIT_ORTE]")
                    If oRs.RecordCount > 0 Then
                        oForm.Freeze(True)
                        LoadMachineData()
                        LoadToolsData()
                        LoadLabourData()
                        SetItemEnabled()
                        If oRouteMatrix.RowCount > 0 Then
                            oRouteMatrix.SelectRow(1, True, False)
                        End If
                        If oMacMatrix.RowCount > 0 Then
                            oMacMatrix.SelectRow(1, True, False)
                        End If
                        oForm.Freeze(False)
                    Else
                        oRouteIdTxt.Active = True
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
    ''' Choosing ItemCode,OperationsCode,MachineCode,LabourCode,ToolCode from the CFL and setting the values 
    ''' to the corresponding field.
    ''' </summary>
    ''' <param name="ControlName"></param>
    ''' <param name="ColumnUID"></param>
    ''' <param name="CurrentRow"></param>
    ''' <param name="ChoosefromListUID"></param>
    ''' <param name="ChooseFromListSelectedObjects"></param>
    ''' <remarks></remarks>
    Private Sub OperationsRouting_ChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable) Handles Me.ChooseFromList
        Dim oDataTable As SAPbouiCOM.DataTable = ChooseFromListSelectedObjects
        Dim oItemCode, oItemName, oOperCode, oOperName, oOperType, oMacCode, oMacName, oMacGroup, oSkGrpCode, oSkGrpName, oToolCode, oToolName As String
        Dim oCurrentRow, IntICount As Integer
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oStrSql As String
        Try
            oCurrentRow = CurrentRow
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                '*********************************Account Code**********************************
                If (ControlName = "btnitmcod" Or ControlName = "txtitmcod") And (ChoosefromListUID = "ItmBtnLst" Or ChoosefromListUID = "ItmTxtLst") Then
                    If Not oDataTable Is Nothing Then
                        oItemCode = oDataTable.GetValue("ItemCode", 0)
                        oItemName = oDataTable.GetValue("ItemName", 0)
                        oParentDB.Offset = oParentDB.Size - 1
                        oParentDB.SetValue("U_Itemcode", oParentDB.Offset, oItemCode)
                        oParentDB.SetValue("U_Itemname", oParentDB.Offset, oItemName)
                    End If
                End If
                '*********************************Route**********************************
                If ControlName = "matroutes" And ChoosefromListUID = "OprLst" Then
                    If Not oDataTable Is Nothing Then
                        
                        oOperCode = oDataTable.GetValue("Code", 0)
                        oOperName = oDataTable.GetValue("U_Oprname", 0)
                        oOperType = oDataTable.GetValue("U_Oprtype", 0)
                        oRouteMatrix.GetLineData(CurrentRow)
                        'oRouteDB.SetValue("U_Seqnce", oRouteDB.Offset, "0")
                        'oRouteDB.SetValue("U_Parid", oRouteDB.Offset, "0")
                        'oRouteMatrix.SetLineData(CurrentRow)
                        If CurrentRow = oRouteMatrix.VisualRowCount Then
                            oRouteDB.Offset = oRouteDB.Size - 1
                            SetRouteDefaultValue()
                            oRouteMatrix.SetLineData(CurrentRow)
                            oRouteMatrix.FlushToDataSource()
                        End If
                        'oRouteMatrix.GetLineData(CurrentRow)
                        oRouteDB.SetValue("U_Oprcode", oRouteDB.Offset, oOperCode)
                        oRouteDB.SetValue("U_Oprname", oRouteDB.Offset, oOperName)
                        oRouteDB.SetValue("U_Oprtype", oRouteDB.Offset, oOperType)
                        oRouteMatrix.SetLineData(CurrentRow)
                        oRouteMatrix.FlushToDataSource()
                        '********************************Load Machines********************************************************
                        Try
                            oStrSql = "Select * From [@PSSIT_PRN1] where Code = '" & oOperCode & "'"
                            oRs.DoQuery(oStrSql)
                            'oMachinesDB.Clear()
                            If oRs.RecordCount > 0 Then
                                oMacMatrix.AddRow(1, oMacMatrix.RowCount)
                                If oMacMatrix.RowCount = 1 Then
                                    oMacSerialNo = GenerateSerialNo("PSSIT_RTE1")
                                ElseIf oMacMatrix.RowCount > 1 Then
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        oMacSerialNo = oMacSerialNo + 1
                                    ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        oMacSerialNo = GenerateSerialNo("PSSIT_RTE1")
                                    End If
                                End If
                                ' If oRs.RecordCount > 0 Then
                                oRs.MoveFirst()
                                For IntICount = 0 To oRs.RecordCount - 1
                                    If CurrentRow = oMacMatrix.VisualRowCount Then
                                        oMachinesDB.Offset = oMachinesDB.Size - 1
                                        SetMachineDefaultValue(oOprIdCol.Cells.Item(oRouteMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                                        oMacMatrix.SetLineData(CurrentRow)
                                        oMacMatrix.FlushToDataSource()
                                    End If
                                    oMachinesDB.SetValue("Code", oMachinesDB.Offset, oMacSerialNo)
                                    oMachinesDB.SetValue("U_Rteid", oMachinesDB.Offset, oRouteIdTxt.Value)
                                    oMachinesDB.SetValue("U_OprCode", oMachinesDB.Offset, oOprIdCol.Cells.Item(oRouteMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                                    oMachinesDB.SetValue("U_wcno", oMachinesDB.Offset, oRs.Fields.Item("U_wcno").Value)
                                    oMachinesDB.SetValue("U_wcname", oMachinesDB.Offset, oRs.Fields.Item("U_wcname").Value)
                                    oMachinesDB.SetValue("U_MGname", oMachinesDB.Offset, oRs.Fields.Item("U_MGname").Value)

                                    '----- commented kabilahan begin -----
                                    'oMachinesDB.SetValue("U_Setime", oMachinesDB.Offset, FormatDateTime(Now(), DateFormat.ShortTime))
                                    'oMachinesDB.SetValue("U_Opertime", oMachinesDB.Offset, FormatDateTime(Now(), DateFormat.ShortTime))
                                    'oMachinesDB.SetValue("U_Othetime1", oMachinesDB.Offset, FormatDateTime(Now(), DateFormat.ShortTime))
                                    'oMachinesDB.SetValue("U_Othetime2", oMachinesDB.Offset, FormatDateTime(Now(), DateFormat.ShortTime))

                                    '-----commented kablahan end
                                    oMachinesDB.SetValue("U_perqty", oMachinesDB.Offset, "0")
                                    oMachinesDB.SetValue("U_Adnl1", oMachinesDB.Offset, "")
                                    oMachinesDB.SetValue("U_Adnl2", oMachinesDB.Offset, "")
                                    oMachinesDB.SetValue("U_Adnl3", oMachinesDB.Offset, "")
                                    oMachinesDB.SetValue("U_Adnl4", oMachinesDB.Offset, "")
                                    oMacMatrix.SetLineData(oMacMatrix.RowCount)
                                    If IntICount <> oRs.RecordCount - 1 Then
                                        'If Len(oMacCodeCol.Cells.Item(IntICount + 1).Specific.Value) > 0 Then
                                        oMachinesDB.InsertRecord(oMachinesDB.Size)
                                        oMachinesDB.Offset = oMachinesDB.Size - 1
                                        SetMachineDefaultValue(oSeqCol.Cells.Item(oRouteMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                                        oMacMatrix.AddRow(1, oMacMatrix.RowCount)
                                        oMacSerialNo = oMacSerialNo + 1
                                    End If
                                    oRs.MoveNext()
                                Next
                            End If
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        Finally
                            oRs = Nothing
                            oStrSql = Nothing
                            GC.Collect()
                        End Try
                        If oMacMatrix.RowCount > 0 Then
                            oMacMatrix.SelectRow(1, True, False)
                        End If
                        '********************************Load Tools********************************************************
                        Try
                            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oStrSql = "Select T0.* from [@PSSIT_PRN3] T0 " _
                            & "Inner Join [@PSSIT_PRN1] T1 On T1.U_wcno = T0.U_wcno " _
                            & "Inner Join [@PSSIT_OPRN] T2 On T2.Code = T1.Code " _
                            & "Where T0.U_Oprcode = '" & oOperCode & "' " _
                            & "Group by T0.Code,T0.Name,T0.U_ToolCode,T0.U_TLName,T0.U_OprCode,T0.U_wcno"
                            oRs.DoQuery(oStrSql)
                            'oToolsDB.Clear()
                            If oRs.RecordCount > 0 Then
                                oToolsMatrix.AddRow(1, oToolsMatrix.RowCount)
                                If oToolsMatrix.RowCount = 1 Then
                                    oToolsSerialNo = GenerateSerialNo("PSSIT_RTE3")
                                ElseIf oToolsMatrix.RowCount > 1 Then
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        oToolsSerialNo = oToolsSerialNo + 1
                                    ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        oToolsSerialNo = GenerateSerialNo("PSSIT_RTE3")
                                    End If
                                End If
                                'If oRs.RecordCount > 0 Then
                                oRs.MoveFirst()
                                For IntICount = 0 To oRs.RecordCount - 1
                                    If CurrentRow = oToolsMatrix.VisualRowCount Then
                                        oToolsDB.Offset = oToolsDB.Size - 1
                                        SetToolsDefaultValue(oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value, oOprIdCol.Cells.Item(oRouteMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                                        oToolsMatrix.SetLineData(CurrentRow)
                                        oToolsMatrix.FlushToDataSource()
                                    End If
                                    oToolsDB.SetValue("Code", oToolsDB.Offset, oToolsSerialNo)
                                    oToolsDB.SetValue("U_Rteid", oToolsDB.Offset, oRouteIdTxt.Value)
                                    oToolsDB.SetValue("U_OprCode", oToolsDB.Offset, oOprIdCol.Cells.Item(oRouteMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                                    oToolsDB.SetValue("U_wcno", oToolsDB.Offset, oRs.Fields.Item("U_wcno").Value)
                                    oToolsDB.SetValue("U_Toolcode", oToolsDB.Offset, oRs.Fields.Item("U_Toolcode").Value)
                                    oToolsDB.SetValue("U_TLname", oToolsDB.Offset, oRs.Fields.Item("U_TLname").Value)
                                    oToolsDB.SetValue("U_Strokes", oToolsDB.Offset, "")
                                    oToolsDB.SetValue("U_Adnl1", oToolsDB.Offset, "")
                                    oToolsDB.SetValue("U_Adnl2", oToolsDB.Offset, "")
                                    oToolsMatrix.SetLineData(oToolsMatrix.RowCount)
                                    If IntICount <> oRs.RecordCount - 1 Then
                                        oToolsDB.InsertRecord(oToolsDB.Size)
                                        oToolsDB.Offset = oToolsDB.Size - 1
                                        SetToolsDefaultValue(oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value, oOprIdCol.Cells.Item(oRouteMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                                        oToolsMatrix.AddRow(1, oToolsMatrix.RowCount)
                                        oToolsSerialNo = oToolsSerialNo + 1
                                        'oMacMatrix.SelectRow(oMacMatrix.RowCount, True, False)
                                    End If
                                    oRs.MoveNext()
                                Next
                            End If
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        Finally
                            oRs = Nothing
                            oStrSql = Nothing
                            GC.Collect()
                        End Try

                        'Added by Manimaran ------------S ----Labour details
                        Try
                            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oStrSql = "Select T0.* from [@PSSIT_PRN2] T0 "
                            oStrSql = oStrSql + " Inner Join [@PSSIT_OPRN] T2 On T2.Code = T0.Code "
                            oStrSql = oStrSql + " Where T0.code = '" & oOperCode & "' "
                            oRs.DoQuery(oStrSql)
                            'oToolsDB.Clear()
                            If oRs.RecordCount > 0 Then
                                oLabMatrix.AddRow(1, oLabMatrix.RowCount)
                                If oLabMatrix.RowCount = 1 Then
                                    oLabSerialNo = GenerateSerialNo("PSSIT_RTE2")
                                ElseIf oLabMatrix.RowCount > 1 Then
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        oLabSerialNo = oLabSerialNo + 1
                                    ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        oLabSerialNo = GenerateSerialNo("PSSIT_RTE2")
                                    End If
                                End If
                                'If oRs.RecordCount > 0 Then
                                oRs.MoveFirst()
                                For IntICount = 0 To oRs.RecordCount - 1
                                    If CurrentRow = oLabMatrix.VisualRowCount Then
                                        oLabourDB.Offset = oLabourDB.Size - 1
                                        SetLabourDefaultValue(oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value, _
                                        oMOprCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                                        oLabMatrix.SetLineData(CurrentRow)
                                        oLabMatrix.FlushToDataSource()
                                    End If
                                    oLabourDB.SetValue("Code", oLabourDB.Offset, oLabSerialNo)
                                    oLabourDB.SetValue("U_Skilgrp", oLabourDB.Offset, oRs.Fields.Item("U_Skilgrp").Value)
                                    oLabourDB.SetValue("U_LGname", oLabourDB.Offset, oRs.Fields.Item("U_LGname").Value)
                                    oLabourDB.SetValue("U_Reqno", oLabourDB.Offset, oRs.Fields.Item("U_Reqno").Value)
                                    'Added by Manimaran------s
                                    oLabourDB.SetValue("U_wcno", oLabourDB.Offset, oRs.Fields.Item("Code").Value)
                                    'Added by Manimaran------e
                                    oLabMatrix.SetLineData(oLabMatrix.RowCount)
                                    If IntICount <> oRs.RecordCount - 1 Then
                                        oLabourDB.InsertRecord(oLabourDB.Size)
                                        oLabourDB.Offset = oLabourDB.Size - 1
                                        SetLabourDefaultValue(oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value, _
                                        oMOprCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                                        oLabMatrix.AddRow(1, oLabMatrix.RowCount)
                                        oLabSerialNo = oLabSerialNo + 1
                                    End If
                                    oRs.MoveNext()
                                Next
                            End If
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        Finally
                            oRs = Nothing
                            oStrSql = Nothing
                            GC.Collect()
                        End Try
                        'Added by Manimaran ------------E

                    End If
                    If Len(oOprIdCol.Cells.Item(oRouteMatrix.RowCount).Specific.value) > 0 Then
                        oRouteDB.InsertRecord(oRouteDB.Size)
                        oRouteDB.Offset = oRouteDB.Size - 1
                        SetRouteDefaultValue()
                        oRouteMatrix.AddRow(1, oRouteMatrix.RowCount)
                        'AddBaseLineNo()
                        oRouteDB.SetValue("LineId", oRouteDB.Offset, oRouteMatrix.RowCount)
                        'oRouteDB.SetValue("U_Bselino", oRouteDB.Offset, oRouteBaseLineNo)
                        oRouteMatrix.SetLineData(oRouteMatrix.RowCount)
                    End If

                End If
                '*********************************Machines**********************************
                If ControlName = "matmachine" And ChoosefromListUID = "MacLst" Then
                    If Not oDataTable Is Nothing Then
                        oMacCode = oDataTable.GetValue("U_wcno", 0)
                        oMacName = oDataTable.GetValue("U_wcname", 0)
                        oMacGroup = oDataTable.GetValue("U_MGcode", 0)
                        oMacMatrix.GetLineData(CurrentRow)
                        oMachinesDB.SetValue("U_wcno", oMachinesDB.Offset, oMacCode)
                        oMachinesDB.SetValue("U_wcname", oMachinesDB.Offset, oMacName)
                        oMachinesDB.SetValue("U_MGname", oMachinesDB.Offset, oMacGroup)
                        oMacMatrix.SetLineData(CurrentRow)
                        oMacMatrix.FlushToDataSource()
                    End If
                End If
                '*********************************Labour**********************************
                If ControlName = "matlabour" And ChoosefromListUID = "SkGrpLst" Then
                    If Not oDataTable Is Nothing Then
                        oSkGrpCode = oDataTable.GetValue("Code", 0)
                        oSkGrpName = oDataTable.GetValue("U_LGname", 0)
                        If CurrentRow = oLabMatrix.VisualRowCount Then
                            oLabourDB.Offset = oLabourDB.Size - 1
                            SetLabourDefaultValue(oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value, _
                            oMOprCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                            oLabMatrix.SetLineData(CurrentRow)
                            oLabMatrix.FlushToDataSource()
                        End If
                        oLabMatrix.GetLineData(CurrentRow)
                        oLabourDB.SetValue("U_Skilgrp", oLabourDB.Offset, oSkGrpCode)
                        oLabourDB.SetValue("U_LGname", oLabourDB.Offset, oSkGrpName)
                        oLabourDB.SetValue("U_wcno", oLabourDB.Offset, oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                        oLabMatrix.SetLineData(CurrentRow)
                        oLabMatrix.FlushToDataSource()
                    End If
                End If
                ' '*********************************Tools**********************************
                If ControlName = "mattool" And ChoosefromListUID = "ToolsLst" Then
                    If Not oDataTable Is Nothing Then
                        'Added by Manimaran-----s
                        Dim code, Maccode As String
                        Dim i As Integer
                        If oToolsMatrix.RowCount > 0 Then
                            For i = 1 To oToolsMatrix.RowCount
                                oToolsMatrix.GetLineData(i)
                                code = oToolsMatrix.Columns.Item("coltolcod").Cells.Item(i).Specific.string
                                Maccode = oToolsMatrix.Columns.Item("colwcno").Cells.Item(i).Specific.string
                                If code = oDataTable.GetValue("Code", 0) And oToolsMatrix.Columns.Item("colwcno").Cells.Item(oToolsMatrix.RowCount).Specific.string = Maccode Then
                                    Throw New Exception("Selected code found already exists")
                                End If
                            Next
                        End If
                        'Added by Manimaran-----e
                        oToolCode = oDataTable.GetValue("Code", 0)
                        oToolName = oDataTable.GetValue("U_TLname", 0)
                        oToolsMatrix.GetLineData(CurrentRow)
                        oToolsDB.SetValue("U_Oprcode", oToolsDB.Offset, oMOprCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                        oToolsDB.SetValue("U_wcno", oToolsDB.Offset, oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.Value)
                        oToolsDB.SetValue("U_Toolcode", oToolsDB.Offset, oToolCode)
                        oToolsDB.SetValue("U_TLname", oToolsDB.Offset, oToolName)
                        oToolsMatrix.SetLineData(CurrentRow)
                        oToolsMatrix.FlushToDataSource()
                    End If
                End If
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    ''' <summary>
    ''' Setting default value to the newly inserting column in the datasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetRouteDefaultValue()
        Try
            'oRouteDB.SetValue("U_Seqnce", oRouteDB.Offset, "0")
            'oRouteDB.SetValue("U_Parid", oRouteDB.Offset, "0")
            oRouteDB.SetValue("U_Oprcode", oRouteDB.Offset, "")
            oRouteDB.SetValue("U_Oprname", oRouteDB.Offset, "")
            oRouteDB.SetValue("U_Milestne", oRouteDB.Offset, "N")
            oRouteDB.SetValue("U_Oprtype", oRouteDB.Offset, "")
            oRouteDB.SetValue("U_SCRate", oRouteDB.Offset, "0.00")
            oRouteDB.SetValue("U_Qty", oRouteDB.Offset, "0.00")
            oRouteDB.SetValue("U_Adnl1", oRouteDB.Offset, "")
            oRouteDB.SetValue("U_Adnl2", oRouteDB.Offset, "")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Setting default value to the newly inserting column in the datasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetMachineDefaultValue(ByVal aOperationCode As String)
        Try
            oMachinesDB.SetValue("U_Rteid", oMachinesDB.Offset, oRouteIdTxt.Value)
            oMachinesDB.SetValue("U_Oprcode", oMachinesDB.Offset, aOperationCode)
            oMachinesDB.SetValue("U_wcno", oMachinesDB.Offset, "")
            oMachinesDB.SetValue("U_wcname", oMachinesDB.Offset, "")
            oMachinesDB.SetValue("U_MGname", oMachinesDB.Offset, "")

            ' commented by kabilahan b

            'oMachinesDB.SetValue("U_Setime", oMachinesDB.Offset, FormatDateTime(Now(), DateFormat.ShortTime))
            'oMachinesDB.SetValue("U_Opertime", oMachinesDB.Offset, FormatDateTime(Now(), DateFormat.ShortTime))
            'oMachinesDB.SetValue("U_Othetime1", oMachinesDB.Offset, FormatDateTime(Now(), DateFormat.ShortTime))
            'oMachinesDB.SetValue("U_Othetime2", oMachinesDB.Offset, FormatDateTime(Now(), DateFormat.ShortTime))

            ' commented by kabilahan e

            oMachinesDB.SetValue("U_perqty", oMachinesDB.Offset, "0")
            oMachinesDB.SetValue("U_Adnl1", oMachinesDB.Offset, "")
            oMachinesDB.SetValue("U_Adnl2", oMachinesDB.Offset, "")
            oMachinesDB.SetValue("U_Adnl3", oMachinesDB.Offset, "")
            oMachinesDB.SetValue("U_Adnl4", oMachinesDB.Offset, "")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Setting default value to the newly inserting column in the datasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetLabourDefaultValue(ByVal aMachineCode As String, ByVal aOperationCode As String)
        Try
            oLabourDB.SetValue("U_Rteid", oLabourDB.Offset, oRouteIdTxt.Value)
            oLabourDB.SetValue("U_Oprcode", oLabourDB.Offset, aOperationCode)
            'Modified by Manimaran-----s
            'oLabourDB.SetValue("U_wcno", oLabourDB.Offset, aMachineCode)
            oLabourDB.SetValue("U_wcno", oLabourDB.Offset, "")
            'Modified by Manimaran-----e
            oLabourDB.SetValue("U_Skilgrp", oLabourDB.Offset, "")
            oLabourDB.SetValue("U_LGname", oLabourDB.Offset, "")
            oLabourDB.SetValue("U_Reqtime", oLabourDB.Offset, FormatDateTime(Now(), DateFormat.ShortTime))
            oLabourDB.SetValue("U_Reqno", oLabourDB.Offset, "0")
            oLabourDB.SetValue("U_Adnl1", oLabourDB.Offset, "")
            oLabourDB.SetValue("U_Adnl2", oLabourDB.Offset, "")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This method is used to add a row in the ToolsDB.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AddBaseLineNo()
        Dim oStrSql As String
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            If oRouteMatrix.RowCount = 1 Then
                oRouteBaseLineNo = 1
            ElseIf oRouteMatrix.RowCount > 1 Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    oRouteBaseLineNo = oRouteBaseLineNo + 1
                ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    oStrSql = "Select IsNull(Max(U_Bselino),0) as '#' from [@PSSIT_RTE4] Where Code = '" & oRouteIdTxt.Value & "'"
                    oRs.DoQuery(oStrSql)
                    If oRs.RecordCount > 0 Then
                        oRouteBaseLineNo = oRs.Fields.Item("#").Value + 1
                    Else
                        oRouteBaseLineNo = 1
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
    ''' <summary>
    ''' This method is used to add a row in the ToolsDB.
    ''' </summary>
    ''' <param name="aMachineCode"></param>
    ''' <remarks></remarks>
    Private Sub AddLabourRow(ByVal aMachineCode As String, ByVal aOperationCode As String)
        Try
            oLabourDB.Offset = oLabourDB.Size - 1
            SetLabourDefaultValue(aMachineCode, aOperationCode)
            oLabMatrix.AddRow(1, oLabMatrix.RowCount)
            If oLabMatrix.RowCount = 1 Then
                oLabSerialNo = GenerateSerialNo("PSSIT_RTE2")
            ElseIf oLabMatrix.RowCount > 1 Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    oLabSerialNo = oLabSerialNo + 1
                ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    oLabSerialNo = oLabSerialNo + 1
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This method is used to add a row in the ToolsDB.
    ''' </summary>
    ''' <param name="aOperationCode"></param>
    ''' <remarks></remarks>
    Private Sub AddMachineRow(ByVal aOperationCode As String)
        Try
            oMachinesDB.Offset = oMachinesDB.Size - 1
            SetMachineDefaultValue(aOperationCode)
            oMacMatrix.AddRow(1, oMacMatrix.RowCount)
            If oMacMatrix.RowCount = 1 Then
                oMacSerialNo = GenerateSerialNo("PSSIT_RTE1")
            ElseIf oMacMatrix.RowCount > 1 Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    oMacSerialNo = oMacSerialNo + 1
                ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    oMacSerialNo = oMacSerialNo + 1
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This method is used to add a row in the ToolsDB.
    ''' </summary>
    ''' <param name="aOperationCode"></param>
    ''' <remarks></remarks>
    Private Sub AddToolsRow(ByVal aMacCode As String, ByVal aOperationCode As String)
        Try
            oToolsDB.Offset = oToolsDB.Size - 1
            SetToolsDefaultValue(aMacCode, aOperationCode)
            oToolsMatrix.AddRow(1, oToolsMatrix.RowCount)
            If oToolsMatrix.RowCount = 1 Then
                oToolsSerialNo = GenerateSerialNo("PSSIT_RTE3")
            ElseIf oToolsMatrix.RowCount > 1 Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    oToolsSerialNo = oToolsSerialNo + 1
                ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    oToolsSerialNo = oToolsSerialNo + 1
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
    Private Sub SetToolsDefaultValue(ByVal aMacCode As String, ByVal aOperationCode As String)
        Try
            oToolsDB.SetValue("U_Rteid", oToolsDB.Offset, oRouteIdTxt.Value)
            oToolsDB.SetValue("U_Oprcode", oToolsDB.Offset, aOperationCode)
            oToolsDB.SetValue("U_wcno", oToolsDB.Offset, aMacCode)
            oToolsDB.SetValue("U_Toolcode", oToolsDB.Offset, "")
            oToolsDB.SetValue("U_TLname", oToolsDB.Offset, "")
            oToolsDB.SetValue("U_Strokes", oToolsDB.Offset, "")
            oToolsDB.SetValue("U_Adnl1", oToolsDB.Offset, "")
            oToolsDB.SetValue("U_Adnl2", oToolsDB.Offset, "")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This method is used to delete the empty rows in the Route Matrix.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub RouteDeleteEmptyRow()
        Dim oSeqEdit, oOprCodeEdit, oOprNameEdit As SAPbouiCOM.EditText
        Dim IntICount As Integer
        Try
            If oRouteMatrix.RowCount > 1 Then
                For IntICount = oRouteMatrix.RowCount To 1 Step -1
                    oRouteMatrix.GetLineData(IntICount)
                    oSeqEdit = oSeqCol.Cells.Item(IntICount).Specific
                    oOprCodeEdit = oOprIdCol.Cells.Item(IntICount).Specific
                    oOprNameEdit = oOprNameCol.Cells.Item(IntICount).Specific
                    If oSeqEdit.Value = "" And oOprCodeEdit.Value.Length = 0 And oOprNameEdit.Value.Length = 0 Then
                        oRouteMatrix.DeleteRow(IntICount)
                        oRouteMatrix.FlushToDataSource()
                    End If
                Next
            End If
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
    Private Sub SetItemEnabled()
        Try
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                '  oForm.Items.Item("txtocode").Enabled = False
            Else
                oForm.Items.Item("txtocode").Enabled = True
            End If
        Catch ex As Exception
            Throw ex
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
            oCondition.Alias = "U_Rteid"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = oRouteIdTxt.Value
            oToolsDB.Query(oConditions)
            oToolsMatrix.LoadFromDataSource()
            oToolsMatrix.FlushToDataSource()
            oRs.DoQuery("Select IsNull(Max(Code),0) as Code from [@PSSIT_RTE3]")
            oToolsSerialNo = oRs.Fields.Item("Code").Value
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' Loading the Machine from the database as per the conditions.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadMachineData()
        Dim oConditions As SAPbouiCOM.Conditions
        Dim oCondition As SAPbouiCOM.Condition
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            oConditions = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions)
            oCondition = oConditions.Add
            oCondition.Alias = "U_Rteid"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = oRouteIdTxt.Value
            oMachinesDB.Query(oConditions)
            oMacMatrix.LoadFromDataSource()
            oMacMatrix.FlushToDataSource()
            oRs.DoQuery("Select IsNull(Max(Code),0) as Code from [@PSSIT_RTE1]")
            oMacSerialNo = oRs.Fields.Item("Code").Value
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' Loading the Labour from the database as per the conditions.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadLabourData()
        Dim oConditions As SAPbouiCOM.Conditions
        Dim oCondition As SAPbouiCOM.Condition
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            oConditions = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions)
            oCondition = oConditions.Add
            oCondition.Alias = "U_Rteid"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = oRouteIdTxt.Value
            oLabourDB.Query(oConditions)
            oLabMatrix.LoadFromDataSource()
            oLabMatrix.FlushToDataSource()
            oRs.DoQuery("Select IsNull(Max(Code),0) as Code from [@PSSIT_RTE2]")
            oLabSerialNo = oRs.Fields.Item("Code").Value
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' Resizing the form 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Form_Resize()
        Try
            If oBoolResize = False Then
                oForm.Freeze(True)

                oForm.Items.Item("chkdefrte").Left = oForm.Items.Item("lblitmnam").Left

                oForm.Items.Item("matroutes").Height = 120
                oForm.Items.Item("matroutes").Top = 70
                oForm.Items.Item("matroutes").Width = oForm.Width - 20
                oForm.Items.Item("1000001").Top = oForm.Items.Item("matroutes").Top + oForm.Items.Item("matroutes").Height + 10
                oForm.Items.Item("matmachine").Height = 90
                oForm.Items.Item("matmachine").Top = oForm.Items.Item("1000001").Top + oForm.Items.Item("1000001").Height + 10
                oForm.Items.Item("matmachine").Width = oForm.Width - 20

                oForm.Items.Item("foltools").Top = oForm.Items.Item("matmachine").Top + oForm.Items.Item("matmachine").Height + 20
                oForm.Items.Item("follabour").Top = oForm.Items.Item("matmachine").Top + oForm.Items.Item("matmachine").Height + 20

                oForm.Items.Item("recttools").Left = 5
                oForm.Items.Item("recttools").Height = 120
                oForm.Items.Item("recttools").Top = oForm.Items.Item("foltools").Top + oForm.Items.Item("foltools").Height
                oForm.Items.Item("recttools").Width = oForm.Width - 20

                oForm.Items.Item("rectlab").Left = 5
                oForm.Items.Item("rectlab").Height = 120
                oForm.Items.Item("rectlab").Top = oForm.Items.Item("follabour").Top + oForm.Items.Item("follabour").Height
                oForm.Items.Item("rectlab").Width = oForm.Width - 20

                oForm.Items.Item("mattool").Left = 10
                oForm.Items.Item("mattool").Height = oForm.Items.Item("recttools").Height - 10
                oForm.Items.Item("mattool").Top = oForm.Items.Item("recttools").Top + 5
                oForm.Items.Item("mattool").Width = oForm.Items.Item("recttools").Width - 10

                oForm.Items.Item("matlabour").Left = 10
                oForm.Items.Item("matlabour").Height = oForm.Items.Item("rectlab").Height - 10
                oForm.Items.Item("matlabour").Top = oForm.Items.Item("rectlab").Top + 5
                oForm.Items.Item("matlabour").Width = oForm.Items.Item("rectlab").Width - 10
                oForm.Freeze(False)
                oForm.Update()
                oBoolResize = True
            ElseIf oBoolResize = True Then
                oForm.Freeze(True)

                oForm.Items.Item("chkdefrte").Left = 257

                oForm.Items.Item("matroutes").Height = 91
                oForm.Items.Item("matroutes").Top = 70
                oForm.Items.Item("matroutes").Width = 547

                oForm.Items.Item("1000001").Top = 167
                oForm.Items.Item("1000001").Left = 5

                oForm.Items.Item("matmachine").Height = 86
                oForm.Items.Item("matmachine").Top = 186
                oForm.Items.Item("matmachine").Width = 547

                oForm.Items.Item("foltools").Top = 279
                oForm.Items.Item("foltools").Left = 5

                oForm.Items.Item("follabour").Top = 279
                oForm.Items.Item("follabour").Left = 84

                oForm.Items.Item("recttools").Left = 5
                oForm.Items.Item("recttools").Height = 96
                oForm.Items.Item("recttools").Top = 298
                oForm.Items.Item("recttools").Width = 547

                oForm.Items.Item("rectlab").Left = 5
                oForm.Items.Item("rectlab").Height = 96
                oForm.Items.Item("rectlab").Top = 298
                oForm.Items.Item("rectlab").Width = 547

                oForm.Items.Item("mattool").Left = 10
                oForm.Items.Item("mattool").Height = 86
                oForm.Items.Item("mattool").Top = 303
                oForm.Items.Item("mattool").Width = 537

                oForm.Items.Item("matlabour").Left = 10
                oForm.Items.Item("matlabour").Height = 86
                oForm.Items.Item("matlabour").Top = 303
                oForm.Items.Item("matlabour").Width = 537
                oBoolResize = False
                oForm.Freeze(False)
                oForm.Update()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
