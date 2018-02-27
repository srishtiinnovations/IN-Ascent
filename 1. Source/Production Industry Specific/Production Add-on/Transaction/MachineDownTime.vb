''' <summary>
''' Author                        Created Date
''' Suresh                       17/01/2009
''' </summary>
''' <remarks>This class is used for entering the Machine Down Time Details. </remarks>
Public Class MachineDownTime
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
    Private oParentDB, oRSDB, oRMDB, oMRDB As SAPbouiCOM.DBDataSource
    '**************************UserDataSource************************************
    Private oUD As SAPbouiCOM.UserDataSource
    Private oForm As SAPbouiCOM.Form
    '**************************ChooseFromList************************************
    Private oChMCList, oChMCBtnList, oChMRList, oShiftList, oShiftBtnList, oPECFL, oPECFLbtn As SAPbouiCOM.ChooseFromList
    Private Event ChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable)
    '**************************Items - EditText************************************
    Private oPETxt, oShiftCodeTxt, oShiftNameTxt, oDocEntryTxt, oDocNoTxt, oDocDtTxt, oDeptCodeTxt, oDeptNameTxt, oWCCodeTxt, oWCNameTxt, oIntDtTxt, oIntTimeTxt, oBDWDTxt, oAttDtTxt, oAttTimeTxt, oCmpltDtTxt, oCmpltTimeTxt, oStpMtTxt, oEmpIdTxt, oEmpNameTxt, oSerTxt As SAPbouiCOM.EditText
    '**************************Items - Combo************************************
    Private oWCCombo, oEmpCombo, oSerCombo As SAPbouiCOM.ComboBox
    '**************************Items - Button************************************
    Private BtnMC, oShiftBtn, oPEbtn As SAPbouiCOM.Button
    '**************************Items - LinkedButton************************************
    Private oMCLink, oShiftLink As SAPbouiCOM.LinkedButton
    '**************************Items - Matrix************************************
    Private oRSMatrix, oRMMatrix, oMRMatrix As SAPbouiCOM.Matrix
    Private oRSColumns, oRMColumns, oMRColumns As SAPbouiCOM.Columns
    Private oRSColumn, oREColumn, oMRColumn As SAPbouiCOM.Column
    '*************************Reasons-MAtrix Column*********************************
    Private oRsCodeCol, oRsNameCol, oRsStTime, oRsEndTime, oRsStpTime As SAPbouiCOM.Column
    '*************************Remedies-Matrix Column***************************
    Private oRMCodeCol, oRMNameCol As SAPbouiCOM.Column
    '************************Materials-Matrix***********************************
    Private oMRItmCodeCol, oMRItmNameCol, oMRQtyCol, oMRUomCol, oMRRateCol, oMRValueCol As SAPbouiCOM.Column
    '*********************Variables**************************
    Dim oDocNoSerialNo As Integer
    Private BoolResize As Boolean
    Private BoolRes As Boolean = True
    Private BoolRem As Boolean = True
    Private oRSComboRow, oRMComboRow As Integer
    Private ReaUID, ResUID, MatUID As String
    Private sPath As String = IO.Directory.GetParent(Application.ExecutablePath).ToString
    Private WithEvents MachineMasterClass As MachineMaster
    Private WithEvents ShiftMaster As Shift
    Private sQry As String
    Private Rs As SAPbobsCOM.Recordset
    Dim oToTime As Integer
    Dim ofrTime As Integer
    Dim ofldr As SAPbouiCOM.Folder

#End Region
    ''' <summary>
    ''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmDownTimeEntry.srf") method is called to load the Labour form.
    ''' Drawform() method is called to Initialize the form,Datasources and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company)
        MyBase.New(aSBO_Application, aCompany)
        SBO_Application = aSBO_Application
        oCompany = aCompany
        LoadFromXML("FrmDownTimeEntry.srf")
        DrawForm()
        'Added by Manimaran-----s
        ofldr = oForm.Items.Item("1000001").Specific
        ofldr.Select()
        'Added by Manimaran-----e
    End Sub

    Private Sub DrawForm()
        Try
            oForm = SBO_Application.Forms.Item("FrmDownTimeEntry")

            oParentDB = oForm.DataSources.DBDataSources.Item("@PSSIT_PMWCBREAKHDR") 'Header Datasource
            oRSDB = oForm.DataSources.DBDataSources.Item("@PSSIT_PMWCREASONDTL") 'Reason For BreakDown Datasource
            oRMDB = oForm.DataSources.DBDataSources.Item("@PSSIT_PMWCREMEDTL") 'Remedies for Break Down Datasource
            oMRDB = oForm.DataSources.DBDataSources.Item("@PSSIT_PMWCITEMSDTL") 'Materials Replaced Datasource

            oForm.Freeze(True)

            LoadLookups()
            InitTxtComp()             'Add EditText
            InitCbopComp()            'Add ComboBox 
            InitMatrix()              'Add Matrix 

            EmpCombo(oEmpCombo)   'Employee Combo Load
            SetToolBarEnabled()
            oForm.Freeze(False)
            oForm.DataBrowser.BrowseBy = "TxtDocNo"

        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    ''' <summary>
    ''' Add EditText Items
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitTxtComp()
        Try
            oDocEntryTxt = oForm.Items.Item("TxtDocEnt").Specific
            oDocEntryTxt.DataBind.SetBound(True, "@PSSIT_PMWCBREAKHDR", "DocEntry")

            'oDocNoTxt = oForm.Items.Item("TxtDocNo").Specific
            'oDocNoTxt.DataBind.SetBound(True, "@PSSIT_PMWCBREAKHDR", "DocNum")
            'LoadDocNumber()
            oDocNoTxt = oForm.Items.Item("TxtDocNo").Specific
            oForm.Items.Item("TxtDocNo").Enabled = False
            oDocNoTxt.DataBind.SetBound(True, "@PSSIT_PMWCBREAKHDR", "DocNum")
            With oForm.DataSources.DBDataSources.Item("@PSSIT_PMWCBREAKHDR")
                .SetValue("DocNum", .Offset, oForm.BusinessObject.GetNextSerialNumber(Trim(.GetValue("Series", .Offset))).ToString)
            End With

            oDocDtTxt = oForm.Items.Item("TxtDocDt").Specific
            oDocDtTxt.DataBind.SetBound(True, "@PSSIT_PMWCBREAKHDR", "U_docdate")
            oDocDtTxt.String = System.DateTime.Today.Date.ToString("dd/MM/yyyy")

            oDeptCodeTxt = oForm.Items.Item("txtmcno").Specific
            oDeptCodeTxt.DataBind.SetBound(True, "@PSSIT_PMWCBREAKHDR", "U_deptcode")
            oDeptCodeTxt.ChooseFromListUID = "MCLst"
            oDeptCodeTxt.ChooseFromListAlias = "U_wcno"
            oForm.Items.Item("txtmcno").LinkTo = "lnkmc"
            oForm.Items.Add("lnkmc", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnkmc").Visible = True
            oForm.Items.Item("lnkmc").LinkTo = "txtmcno"
            oForm.Items.Item("lnkmc").Top = 21
            oForm.Items.Item("lnkmc").Left = 106
            oForm.Items.Item("lnkmc").Description = "Link to" & vbNewLine & "Machine Master"
            oMCLink = oForm.Items.Item("lnkmc").Specific

            oDeptNameTxt = oForm.Items.Item("TxtDept").Specific
            oDeptNameTxt.DataBind.SetBound(True, "@PSSIT_PMWCBREAKHDR", "U_deptdesc")

            oWCCodeTxt = oForm.Items.Item("TxtWC").Specific
            oWCCodeTxt.DataBind.SetBound(True, "@PSSIT_PMWCBREAKHDR", "U_wccode")

            oWCNameTxt = oForm.Items.Item("txtwcname").Specific
            oWCNameTxt.DataBind.SetBound(True, "@PSSIT_PMWCBREAKHDR", "U_wcname")

            oIntDtTxt = oForm.Items.Item("TxtIntDt").Specific
            oIntDtTxt.DataBind.SetBound(True, "@PSSIT_PMWCBREAKHDR", "U_indate")
            oIntDtTxt.String = System.DateTime.Today.Date.ToString("dd/MM/yyyy")

            'oIntTimeTxt = oForm.Items.Item("TxtIntTime").Specific
            'oIntTimeTxt.DataBind.SetBound(True, "@PSSIT_PMWCBREAKHDR", "U_intime")
            'oIntTimeTxt.Value = System.DateTime.Now.ToShortTimeString
            'oIntTimeTxt.String = FormatDateTime(Now(), DateFormat.ShortTime)

            'Dim dt As DateTime
            'dt = DateTime.Parse(System.DateTime.Now.ToShortTimeString)
            'oForm.Items.Item("TxtIntTime").Specific.value = dt.ToString("HH:mm:ss tt")

            oBDWDTxt = oForm.Items.Item("TxtBDWD").Specific
            oBDWDTxt.DataBind.SetBound(True, "@PSSIT_PMWCBREAKHDR", "U_natuwork")

            'oAttDtTxt = oForm.Items.Item("TxtAttDt").Specific
            'oAttDtTxt.DataBind.SetBound(True, "@PSSIT_PMWCBREAKHDR", "U_attdate")
            'oAttDtTxt.String = System.DateTime.Today.Date.ToString("dd/MM/yyyy")

            'oAttTimeTxt = oForm.Items.Item("TxtAttTime").Specific
            'oAttTimeTxt.DataBind.SetBound(True, "@PSSIT_PMWCBREAKHDR", "U_atttime")
            'oAttTimeTxt.String = FormatDateTime(Now(), DateFormat.ShortTime)

            'oCmpltDtTxt = oForm.Items.Item("TxtCmpDt").Specific
            'oCmpltDtTxt.DataBind.SetBound(True, "@PSSIT_PMWCBREAKHDR", "U_compdate")
            'oCmpltDtTxt.String = System.DateTime.Today.Date.ToString("dd/MM/yyyy")

            'oCmpltTimeTxt = oForm.Items.Item("TxtCmpTime").Specific
            'oCmpltTimeTxt.DataBind.SetBound(True, "@PSSIT_PMWCBREAKHDR", "U_comptime")
            'oCmpltTimeTxt.String = FormatDateTime(Now(), DateFormat.ShortTime)

            'oStpMtTxt = oForm.Items.Item("TxtTotStp").Specific
            'oStpMtTxt.DataBind.SetBound(True, "@PSSIT_PMWCBREAKHDR", "U_totmts")

            'oStpMtTxt.enable = False

            'oEmpIdTxt = oForm.Items.Item("TxtEmpId").Specific
            'oEmpIdTxt.DataBind.SetBound(True, "@PSSIT_PMWCBREAKHDR", "U_empid")

            oEmpNameTxt = oForm.Items.Item("TxtEmpName").Specific
            oEmpNameTxt.DataBind.SetBound(True, "@PSSIT_PMWCBREAKHDR", "U_empname")

            oSerTxt = oForm.Items.Item("TxtSer").Specific
            oSerTxt.DataBind.SetBound(True, "@PSSIT_PMWCBREAKHDR", "Series")

            BtnMC = oForm.Items.Item("btnmc").Specific
            oForm.Items.Item("btnmc").Description = "Choose from List" & vbNewLine & "Machine Master List View"
            BtnMC.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            BtnMC.Image = sPath & "\Resources\CFL.bmp"
            BtnMC = oForm.Items.Item("btnmc").Specific
            BtnMC.ChooseFromListUID = "BtMCLst"

            oShiftCodeTxt = oForm.Items.Item("txtSftCode").Specific
            oForm.Items.Item("txtSftCode").Enabled = True
            oShiftCodeTxt.DataBind.SetBound(True, "@PSSIT_PMWCBREAKHDR", "U_Scode")
            oShiftCodeTxt.ChooseFromListUID = "SftLst"
            oShiftCodeTxt.ChooseFromListAlias = "Code"
            oForm.Items.Item("txtSftCode").LinkTo = "lnksft"
            oForm.Items.Add("lnksft", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnksft").Visible = True
            oForm.Items.Item("lnksft").LinkTo = "txtSftCode"
            oForm.Items.Item("lnksft").Height = 12
            oForm.Items.Item("lnksft").Width = 9
            oForm.Items.Item("lnksft").Top = 67
            oForm.Items.Item("lnksft").Left = 112
            oForm.Items.Item("lnksft").Description = "Link to" & vbNewLine & "Shift"
            oShiftLink = oForm.Items.Item("lnksft").Specific

            oShiftBtn = oForm.Items.Item("btnscode").Specific
            oForm.Items.Item("btnscode").Enabled = True
            oForm.Items.Item("btnscode").Description = "Choose from List" & vbNewLine & "Shift List View"
            oShiftBtn.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            oShiftBtn.Image = sPath & "\Resources\CFL.bmp"
            oShiftBtn.ChooseFromListUID = "BtSftLst"

            oShiftNameTxt = oForm.Items.Item("txtSftDesc").Specific
            oForm.Items.Item("txtSftDesc").Enabled = False
            oShiftNameTxt.DataBind.SetBound(True, "@PSSIT_PMWCBREAKHDR", "U_Sdesc")

            'Added by Manimaran------s
            oPETxt = oForm.Items.Item("41").Specific
            oForm.Items.Item("41").Enabled = True
            oPETxt.DataBind.SetBound(True, "@PSSIT_PMWCBREAKHDR", "U_PENum")
            oPETxt.ChooseFromListUID = "PeLst"
            oPETxt.ChooseFromListAlias = "DocNum"
            'Added by Manimaran------e

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Add ComboBox Items
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitCbopComp()
        Dim oForm As SAPbouiCOM.Form
        Try
            oForm = SBO_Application.Forms.Item(SBO_Application.Forms.ActiveForm.UniqueID)

            oSerCombo = oForm.Items.Item("CmbSer").Specific
            oForm.Items.Item("CmbSer").DisplayDesc = True
            oSerCombo.DataBind.SetBound(True, "@PSSIT_PMWCBREAKHDR", "Series")
            oSerCombo.ValidValues.LoadSeries(oForm.BusinessObject.Type, SAPbouiCOM.BoSeriesMode.sf_Add)
            'oSerCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            Dim pa As Integer
            For pa = oSerCombo.ValidValues.Count - 1 To 0 Step -1
                oSerCombo.Select(pa, SAPbouiCOM.BoSearchKey.psk_Index)
            Next


            oEmpCombo = oForm.Items.Item("CmbEmpId").Specific
            oEmpCombo.DataBind.SetBound(True, "@PSSIT_PMWCBREAKHDR", "U_empid")

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' Adding Matrix Items
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitMatrix()
        Dim oItem As SAPbouiCOM.Item
        Try
            ' Add a matrix
            'Reasons For Breakdown
            oItem = oForm.Items.Item("MatReaBD")
            oRSMatrix = oItem.Specific
            oRSMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oRSColumns = oRSMatrix.Columns

            oRsCodeCol = oRSColumns.Item("ReaCode")
            oRsCodeCol.DataBind.SetBound(True, "@PSSIT_PMWCREASONDTL", "U_reascode")
            oRSColumns.Item("ReaCode").Width = 100
            ReaCombo()

            oRsNameCol = oRSColumns.Item("ReaName")
            oRsNameCol.DataBind.SetBound(True, "@PSSIT_PMWCREASONDTL", "U_reasdesc")
            oRsNameCol.Editable = False
            oRSColumns.Item("ReaName").Width = 200
            'Added by Manimaran--------------s
            oRsStTime = oRSColumns.Item("ReaStTime")
            oRsStTime.DataBind.SetBound(True, "@PSSIT_PMWCREASONDTL", "U_StTime")

            oRsEndTime = oRSColumns.Item("ReaEndTime")
            oRsEndTime.DataBind.SetBound(True, "@PSSIT_PMWCREASONDTL", "U_EndTime")

            oRsStpTime = oRSColumns.Item("ReaStp")
            oRsStpTime.DataBind.SetBound(True, "@PSSIT_PMWCREASONDTL", "U_StpgTime")
            oRsStpTime.Editable = False
            'Added by Manimaran--------------e

            'Remedies for Break Down
            oItem = oForm.Items.Item("MatResBD")
            oRMMatrix = oItem.Specific
            oRMMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oRMColumns = oRMMatrix.Columns

            oRMCodeCol = oRMColumns.Item("RemCode")
            oRMCodeCol.DataBind.SetBound(True, "@PSSIT_PMWCREMEDTL", "U_remecode")
            oRMColumns.Item("RemCode").Width = 100
            ResCombo()

            oRMNameCol = oRMColumns.Item("RemName")
            oRMNameCol.DataBind.SetBound(True, "@PSSIT_PMWCREMEDTL", "U_remedesc")
            oRMNameCol.Editable = False
            oRMColumns.Item("RemName").Width = 200

            'Materials Replaced
            oItem = oForm.Items.Item("MatMR")
            oMRMatrix = oItem.Specific
            oMRMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oMRColumns = oMRMatrix.Columns

            oMRItmCodeCol = oMRColumns.Item("ItemCode")
            oMRItmCodeCol.DataBind.SetBound(True, "@PSSIT_PMWCITEMSDTL", "U_itemcode")
            oMRItmCodeCol.Editable = True
            oMRItmCodeCol.ChooseFromListUID = "MRLst"
            oMRItmCodeCol.ChooseFromListAlias = "ItemCode"

            oMRItmNameCol = oMRColumns.Item("ItemName")
            oMRItmNameCol.DataBind.SetBound(True, "@PSSIT_PMWCITEMSDTL", "U_itemdesc")
            oMRItmNameCol.Editable = False

            oMRQtyCol = oMRColumns.Item("Qty")
            oMRQtyCol.DataBind.SetBound(True, "@PSSIT_PMWCITEMSDTL", "U_itemqty")

            oMRUomCol = oMRColumns.Item("Uom")
            oMRUomCol.DataBind.SetBound(True, "@PSSIT_PMWCITEMSDTL", "U_itemuom")
            oMRUomCol.Editable = False

            oMRRateCol = oMRColumns.Item("Rate")
            oMRRateCol.DataBind.SetBound(True, "@PSSIT_PMWCITEMSDTL", "U_itemrate")
            oMRRateCol.Editable = False

            oMRValueCol = oMRColumns.Item("Value")
            oMRValueCol.DataBind.SetBound(True, "@PSSIT_PMWCITEMSDTL", "U_itemval")
            oMRValueCol.Editable = False
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub SetApplication()
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String
        SboGuiApi = New SAPbouiCOM.SboGuiApi
        Try
            sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1))
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show("AddOn must start in SAP Business One")
            System.Environment.Exit(0)
            Throw ex
        End Try
        Try
            SboGuiApi.Connect(sConnectionString)
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show("No SAP Business One Application was found")
            System.Environment.Exit(0)
        End Try
        SboGuiApi.AddonIdentifier = "5645523035446576656C6F706D656E743A453038373933323333343581F0D8D8C45495472FC628EF425AD5AC2AEDC411"
        SBO_Application = SboGuiApi.GetApplication(-1)
    End Sub
    Private Sub SetToolBarEnabled()
        oForm.EnableMenu("1288", True)
        oForm.EnableMenu("1289", True)
        oForm.EnableMenu("1290", True)
        oForm.EnableMenu("1291", True)
        oForm.EnableMenu("1292", True)
        oForm.EnableMenu("1293", True)
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
            '***************************Machine-CFL************************
            oChMCList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_WCHDR", "MCLst"))
            oChMCBtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_WCHDR", "BtMCLst"))

            oShiftList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_SFT", "SftLst"))
            oShiftBtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_SFT", "BtSftLst"))
            '***************************Item-CFL****************************
            oChMRList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "4", "MRLst"))
            'Added by Manimaran-----s
            oPECFL = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "202", "PeLst"))
            SetPOCFLConditions()
            'oPECFLbtn = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_PEY", "BtnPECfl"))
            'Added by Manimaran-----e
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub SetPOCFLConditions()
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oStrSql As String
        Try
            oCFLs = oForm.ChooseFromLists
            For Each oCFL As SAPbouiCOM.ChooseFromList In oCFLs
                If (oCFL.UniqueID.Equals("PeLst")) Then
                    oStrSql = "Select T0.U_Pordno from [@PSSIT_WOR2] T0 " _
                    & "Inner Join OWOR T1 On T1.DocNum = T0.U_POrdno and T1.PlannedQty > T1.CmpltQty " _
                    & "left outer Join IGE1 T2 On T2.BaseRef = T1.DocNum " _
                    & "Where T1.Status = 'R' Group by T0.U_Pordno"
                    oRs.DoQuery(oStrSql)
                    oCFL.SetConditions(Nothing)
                    '************** Adding Conditions to Item List ***************************
                    oCons = oCFL.GetConditions()
                    '************** Condition 1: ItemCode = oVenCodeTxt.Value *********
                    oCon = oCons.Add()
                    oCon.Alias = "DocNum"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    For i As Integer = 1 To oRs.RecordCount
                        If oRs.EoF = False Then
                            oCon.CondVal = oRs.Fields.Item("U_Pordno").Value
                            If Not i = oRs.RecordCount Then
                                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                                oCon = oCons.Add()
                                oCon.Alias = "DocNum"
                                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                            End If
                        End If
                        oRs.MoveNext()
                    Next
                    oCFL.SetConditions(oCons)
                End If
            Next
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
    Private Sub WCSMapping_ChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable) Handles Me.ChooseFromList
        Dim oDataTable As SAPbouiCOM.DataTable = ChooseFromListSelectedObjects
        Dim oCurrentRow As Integer
        Dim oMCNo, oMCName, oWCNo, oWCName, oItemCode, oItemName, oSftCode, oSftDesc As String
        Dim oRs As SAPbobsCOM.Recordset
        Dim StrSql As String
        Try
            oCurrentRow = CurrentRow
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '*************Work Centre CFL**************
            If (ControlName = "txtmcno" Or ControlName = "btnmc") And (ChoosefromListUID = "MCLst" Or ChoosefromListUID = "BtMCLst") Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    If Not oDataTable Is Nothing Then
                        oMCNo = oDataTable.GetValue("U_wcno", 0)
                        oMCName = oDataTable.GetValue("U_wcname", 0)
                        oWCNo = oDataTable.GetValue("U_deptcode", 0)
                        oWCName = oDataTable.GetValue("U_deptdesc", 0)
                        oParentDB.SetValue("U_deptcode", oParentDB.Offset, oMCNo)
                        oParentDB.SetValue("U_deptdesc", oParentDB.Offset, oMCName)
                        oParentDB.SetValue("U_wccode", oParentDB.Offset, oWCNo)
                        oParentDB.SetValue("U_wcname", oParentDB.Offset, oWCName)
                    End If
                Else
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        If Not oDataTable Is Nothing Then
                            oMCNo = oDataTable.GetValue("U_wcno", 0)
                            oMCName = oDataTable.GetValue("U_wcname", 0)
                            oWCNo = oDataTable.GetValue("U_deptcode", 0)
                            oWCName = oDataTable.GetValue("U_deptdesc", 0)
                            oParentDB.SetValue("U_deptcode", oParentDB.Offset, oMCNo)
                            oParentDB.SetValue("U_deptdesc", oParentDB.Offset, oMCName)
                            oParentDB.SetValue("U_wccode", oParentDB.Offset, oWCNo)
                            oParentDB.SetValue("U_wcname", oParentDB.Offset, oWCName)
                        End If
                    End If
                End If
            End If
            'Added by Manimaran------s
            If ControlName = "41" And ChoosefromListUID = "PeLst" Then
                If Not oDataTable Is Nothing Then
                    oParentDB.SetValue("U_PENum", oParentDB.Offset, oDataTable.GetValue("DocNum", 0))
                End If
            End If
            'Added by Manimaran------e
            If (ControlName = "txtSftCode" Or ControlName = "btnscode") And (ChoosefromListUID = "SftLst" Or ChoosefromListUID = "BtSftLst") Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    If Not oDataTable Is Nothing Then
                        oSftCode = oDataTable.GetValue("Code", 0)
                        oSftDesc = oDataTable.GetValue("U_Sdescr", 0)

                        oParentDB.SetValue("U_SCode", oParentDB.Offset, oSftCode)
                        oParentDB.SetValue("U_SDesc", oParentDB.Offset, oSftDesc)

                    End If
                Else
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        If Not oDataTable Is Nothing Then
                            oSftCode = oDataTable.GetValue("Code", 0)
                            oSftDesc = oDataTable.GetValue("U_Sdescr", 0)

                            oParentDB.SetValue("U_SCode", oParentDB.Offset, oSftCode)
                            oParentDB.SetValue("U_SDesc", oParentDB.Offset, oSftDesc)

                        End If
                    End If
                End If
            End If

            If (ControlName = "MatMR") And (ChoosefromListUID = "MRLst") Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    If Not oDataTable Is Nothing Then
                        oItemCode = oDataTable.GetValue("ItemCode", 0)
                        oItemName = oDataTable.GetValue("ItemName", 0)

                        ' ******* Add Next Row If the Item Code is Selected **********
                        If CurrentRow = oMRMatrix.VisualRowCount Then
                            oMRDB.Offset = oMRDB.Size - 1
                            MRSetValue()
                            oMRMatrix.SetLineData(CurrentRow)
                            oMRMatrix.FlushToDataSource()
                        End If
                        oMRDB.SetValue("U_itemcode", oMRDB.Offset, oItemCode)
                        oMRDB.SetValue("U_itemdesc", oMRDB.Offset, oItemName)

                        'StrSql = "select Invntryuom,avgprice from OITM where Itemcode='" & oItemCode & "'"
                        StrSql = "select a.Invntryuom,b.avgprice from OITM a, OITW b ,OWHS c where a.ItemCode=b.ItemCode and b.WhsCode = c.WhsCode and a.Itemcode='" & oItemCode & "'"

                        oRs.DoQuery(StrSql)
                        If oRs.RecordCount > 0 Then
                            oRs.MoveFirst()
                            oMRDB.SetValue("U_itemqty", oMRDB.Offset, "")
                            oMRDB.SetValue("U_itemuom", oMRDB.Offset, oRs.Fields.Item("Invntryuom").Value)
                            oMRDB.SetValue("U_itemrate", oMRDB.Offset, oRs.Fields.Item("avgprice").Value)
                            oMRDB.SetValue("U_itemval", oMRDB.Offset, "")
                        End If
                        oMRMatrix.SetLineData(CurrentRow)
                        oMRMatrix.FlushToDataSource()

                        If Len(oMRColumns.Item("ItemCode").Cells.Item(oMRMatrix.RowCount).Specific.value) > 0 Then
                            oMRDB.InsertRecord(oMRDB.Size)
                            oMRDB.Offset = oMRDB.Size - 1
                            MRSetValue()
                            oMRMatrix.AddRow(1, oMRMatrix.RowCount)
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
    Private Sub SetCFLConditions()
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim StrSql As String
        Try
            oCFLs = oForm.ChooseFromLists
            For Each oCFL As SAPbouiCOM.ChooseFromList In oCFLs
                If (oCFL.UniqueID.Equals("MRLst")) Then
                    StrSql = "select a.Itemcode,a.Dscription from IGE1 a, OIGE b where a.docentry = b.docentry and b.U_machno='" & oDeptCodeTxt.Value & "' and U_wcno='" & oWCCodeTxt.Value & "' and U_status=1"
                    oRs.DoQuery(StrSql)
                    oCFL.SetConditions(Nothing)
                    '************** Adding Conditions to Item List ***************************
                    oCons = oCFL.GetConditions()
                    '************** Condition 1: ItemCode = oVenCodeTxt.Value *********
                    oCon = oCons.Add()
                    oCon.Alias = "ItemCode"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    For i As Integer = 1 To oRs.RecordCount
                        If oRs.EoF = False Then
                            oCon.CondVal = oRs.Fields.Item("ItemCode").Value
                            If Not i = oRs.RecordCount Then
                                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                                oCon = oCons.Add()
                                oCon.Alias = "ItemCode"
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
            oRs = Nothing
            oCon = Nothing
            oCons = Nothing
            oCFLs = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' This is used to Load the Department in the Combo
    ''' </summary>
    ''' <param name="oCombo"></param>
    ''' <remarks></remarks>
    Private Sub DeptCombo(ByVal oCombo2 As SAPbouiCOM.ComboBox)
        Dim rs As SAPbobsCOM.Recordset
        Try
            rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rs.DoQuery("select Code,U_deptdesc from [@PSSIT_PRDEPT] where code is not null")
            rs.MoveFirst()
            If oCombo2.ValidValues.Count > 0 Then
                For i As Int16 = oCombo2.ValidValues.Count - 1 To 0 Step -1
                    oCombo2.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            For i As Int16 = 0 To rs.RecordCount - 1
                oCombo2.ValidValues.Add(rs.Fields.Item(1).Value, rs.Fields.Item(0).Value)
                rs.MoveNext()
            Next
        Catch ex As Exception
            Throw ex
        Finally
            rs = Nothing
            GC.Collect()
        End Try
    End Sub

    ''' <summary>
    ''' This is used to Load the Work Center in the Combo
    ''' </summary>
    ''' <param name="oCombo"></param>
    ''' <remarks></remarks>
    Private Sub WCCombo(ByVal oCombo As SAPbouiCOM.ComboBox)
        Dim rs As SAPbobsCOM.Recordset
        Try
            rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rs.DoQuery("select U_wcno,U_wcname from [@PSSIT_PMWCHDR] where U_deptcode='" & oDeptCodeTxt.Value & "' and U_wcno is not null and U_wcname is not null")
            rs.MoveFirst()
            If oCombo.ValidValues.Count > 0 Then
                For i As Int16 = oCombo.ValidValues.Count - 1 To 0 Step -1
                    oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            For i As Int16 = 0 To rs.RecordCount - 1
                oCombo.ValidValues.Add(rs.Fields.Item(1).Value, rs.Fields.Item(0).Value)
                rs.MoveNext()
            Next
        Catch ex As Exception
            Throw ex
        Finally
            rs = Nothing
            GC.Collect()
        End Try
    End Sub
    Private Sub EmpCombo(ByVal oCombo1 As SAPbouiCOM.ComboBox)
        Dim RS As SAPbobsCOM.Recordset
        Try
            RS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            RS.DoQuery("select empID,firstName from OHEM")
            RS.MoveFirst()
            If oCombo1.ValidValues.Count > 0 Then
                For i As Int16 = oCombo1.ValidValues.Count - 1 To 0 Step -1
                    oCombo1.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            For i As Int16 = 0 To RS.RecordCount - 1
                oCombo1.ValidValues.Add(RS.Fields.Item(0).Value, RS.Fields.Item(1).Value)
                RS.MoveNext()
            Next
        Catch ex As Exception
            Throw ex
        Finally
            RS = Nothing
            GC.Collect()
        End Try
    End Sub
    Private Sub ReaCombo()
        Dim RS As SAPbobsCOM.Recordset
        Try
            RS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            RS.DoQuery("select Code, Name from [@PSSIT_ORES]")

            If oRsCodeCol.ValidValues.Count > 0 Then
                For i As Int16 = oRsCodeCol.ValidValues.Count - 1 To 0 Step -1
                    oRsCodeCol.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            If RS.RecordCount > 0 Then
                RS.MoveFirst()
                For i As Int16 = 0 To RS.RecordCount - 1
                    oRsCodeCol.ValidValues.Add(RS.Fields.Item(0).Value, RS.Fields.Item(1).Value)
                    RS.MoveNext()
                Next
            End If
            oRsCodeCol.ValidValues.Add("Define New", "")
        Catch ex As Exception
            Throw ex
        Finally
            RS = Nothing
            GC.Collect()
        End Try
    End Sub
    Private Sub ReaDFN()
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim StrSql As String
        Dim oRSCombo As SAPbouiCOM.ComboBox
        Try
            If BoolRes = False Then

                If Not oRsCodeCol Is Nothing Then
                    If oRSMatrix.RowCount > 0 Then
                        oRSCombo = oRsCodeCol.Cells.Item(oRSComboRow).Specific()
                        ReaCombo()
                        'StrSql = "Select IsNull(Max(code),0) as Code From [@PSSIT_ORES]"
                        StrSql = "select Name from [@PSSIT_ORES] where code=(Select IsNull(Max(Code),0) as Code from [@PSSIT_ORES])"
                        oRs.DoQuery(StrSql)
                        If oRs.RecordCount > 0 Then
                            oRs.MoveFirst()
                            Dim Val As String = oRs.Fields.Item("Name").Value
                            oRSCombo.Select(Val, SAPbouiCOM.BoSearchKey.psk_ByDescription)
                            BoolRes = True
                        End If
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

    Private Sub ResCombo()
        Dim RS As SAPbobsCOM.Recordset
        Try
            RS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            RS.DoQuery("select Code,U_remedesc from [@PSSIT_PMREMEDIES]")


            If oRMCodeCol.ValidValues.Count > 0 Then
                For i As Int16 = oRMCodeCol.ValidValues.Count - 1 To 0 Step -1
                    oRMCodeCol.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If

            If RS.RecordCount > 0 Then
                RS.MoveFirst()
                For i As Int16 = 0 To RS.RecordCount - 1
                    oRMCodeCol.ValidValues.Add(RS.Fields.Item(0).Value, RS.Fields.Item(1).Value)
                    RS.MoveNext()
                Next
            End If
            oRMCodeCol.ValidValues.Add("Define New", "")
        Catch ex As Exception
            Throw ex
        Finally
            RS = Nothing
            GC.Collect()
        End Try
    End Sub
    Private Sub ResDFN()
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim StrSql As String
        Dim oRMCombo As SAPbouiCOM.ComboBox
        Try
            If BoolRem = False Then

                If Not oRMCodeCol Is Nothing Then
                    If oRMMatrix.RowCount > 0 Then
                        oRMCombo = oRMCodeCol.Cells.Item(oRMComboRow).Specific()
                        ResCombo()
                        'StrSql = "Select IsNull(Max(code),0) as Code From [@PSSIT_PMREMEDIES]"
                        StrSql = "select U_remedesc from [@PSSIT_PMREMEDIES] where code=(Select IsNull(Max(Code),0) as Code from [@PSSIT_PMREMEDIES])"
                        oRs.DoQuery(StrSql)
                        If oRs.RecordCount > 0 Then
                            oRs.MoveFirst()
                            Dim Val As String = oRs.Fields.Item("U_remedesc").Value()
                            oRMCombo.Select(Val, SAPbouiCOM.BoSearchKey.psk_ByDescription)
                            BoolRem = True
                        End If
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
    ''' This is used to Set the values in the Production Parameter matrix while Adding the empty row
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub RSSetValue()
        Try
            oRSDB.SetValue("U_reascode", oRSDB.Offset, "")
            oRSDB.SetValue("U_reasdesc", oRSDB.Offset, "")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This is used to delete the empty rows in the ManPower Requirement Matrix
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub RSDeleteEmptyRow()
        Dim oRSName As SAPbouiCOM.EditText
        Dim IntICount As Integer
        Try
            For IntICount = oRSMatrix.RowCount To 1 Step -1
                oRSName = oRsNameCol.Cells.Item(IntICount).Specific
                If oRSName.Value.Length = 0 And oRSMatrix.RowCount > 1 Then
                    oRSMatrix.DeleteRow(IntICount)
                    oRSMatrix.FlushToDataSource()
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This is used to Set the values in the Production Parameter matrix while Adding the empty row
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub RMSetValue()
        Try
            oRMDB.SetValue("U_remecode", oRMDB.Offset, "")
            oRMDB.SetValue("U_remedesc", oRMDB.Offset, "")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This is used to delete the empty rows in the ManPower Requirement Matrix
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub RMDeleteEmptyRow()
        Dim oRMName As SAPbouiCOM.EditText
        Dim IntICount As Integer
        Try
            For IntICount = oRMMatrix.RowCount To 1 Step -1
                oRMName = oRMNameCol.Cells.Item(IntICount).Specific
                If oRMName.Value.Length = 0 And oRMMatrix.RowCount > 1 Then
                    oRMMatrix.DeleteRow(IntICount)
                    oRMMatrix.FlushToDataSource()
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This is used to Set the values in the Production Parameter matrix while Adding the empty row
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub MRSetValue()
        Try
            oMRDB.SetValue("U_itemcode", oMRDB.Offset, "")
            oMRDB.SetValue("U_itemdesc", oMRDB.Offset, "")
            oMRDB.SetValue("U_itemqty", oMRDB.Offset, "")
            oMRDB.SetValue("U_itemuom", oMRDB.Offset, "")
            oMRDB.SetValue("U_itemrate", oMRDB.Offset, "")
            oMRDB.SetValue("U_itemval", oMRDB.Offset, "")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This is used to delete the empty rows in the Critical Item Matrix
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub MRDeleteEmptyRow()
        Dim oItemCodeEdit, oItemDescEdit As SAPbouiCOM.EditText
        Dim IntICount As Integer
        Try
            For IntICount = oMRMatrix.RowCount To 1 Step -1
                oItemCodeEdit = oMRItmCodeCol.Cells.Item(IntICount).Specific
                oItemDescEdit = oMRItmNameCol.Cells.Item(IntICount).Specific
                If oItemCodeEdit.Value.Length = 0 And oItemDescEdit.Value.Length = 0 And oMRMatrix.RowCount > 1 Then
                    oMRMatrix.DeleteRow(IntICount)
                    oMRMatrix.FlushToDataSource()
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub Form_Resize()
        Try
            If BoolResize = False Then
                'oForm.Freeze(True)
                'oForm.Items.Item("RectRR").Height = 115
                'oForm.Items.Item("RectRR").Top = 85
                'oForm.Items.Item("RectRR").Width = oForm.Width - 20
                'oForm.Items.Item("RectMR").Height = 145
                'oForm.Items.Item("RectMR").Top = 220
                'oForm.Items.Item("RectMR").Width = oForm.Width - 20

                'oForm.Items.Item("MatReaBD").Top = 90
                'oForm.Items.Item("MatReaBD").Height = 110
                'oForm.Items.Item("MatReaBD").Width = 470
                'oRSColumns.Item("ReaCode").Width = 100
                'oRSColumns.Item("ReaName").Width = 150


                'oForm.Items.Item("LblResBD").Left = 500
                'oForm.Items.Item("MatResBD").Left = 500
                'oForm.Items.Item("MatResBD").Top = 90
                'oForm.Items.Item("MatResBD").Height = 110
                'oForm.Items.Item("MatResBD").Width = 505
                'oRMColumns.Item("RemCode").Width = 100
                'oRMColumns.Item("RemName").Width = 150

                'oForm.Freeze(False)
                'oForm.Update()
                'BoolResize = True
            ElseIf BoolResize = True Then
                'oForm.Freeze(True)
                'oForm.Items.Item("RectRR").Height = 115
                'oForm.Items.Item("RectRR").Top = 85
                'oForm.Items.Item("RectRR").Width = 570
                'oForm.Items.Item("RectMR").Height = 110
                'oForm.Items.Item("RectMR").Top = 220
                'oForm.Items.Item("RectMR").Width = 570

                'oForm.Items.Item("MatReaBD").Top = 90
                'oForm.Items.Item("MatReaBD").Height = 110
                'oForm.Items.Item("MatReaBD").Width = 275
                'oRSColumns.Item("ReaCode").Width = 50
                'oRSColumns.Item("ReaName").Width = 100

                'oForm.Items.Item("LblResBD").Left = 290
                'oForm.Items.Item("MatResBD").Left = 290
                'oForm.Items.Item("MatResBD").Top = 90
                'oForm.Items.Item("MatResBD").Height = 110
                'oForm.Items.Item("MatResBD").Width = 275
                'oRMColumns.Item("RemCode").Width = 50
                'oRMColumns.Item("RemName").Width = 100

                'BoolResize = False
                'oForm.Freeze(False)
                'oForm.Update()
            End If
        Catch ex As Exception

        End Try
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="FormUID"></param>
    ''' <param name="pVal"></param>
    ''' <param name="BubbleEvent"></param>
    ''' <remarks></remarks>
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Dim ChooseFromListEvent As SAPbouiCOM.ChooseFromListEvent
        Try
            If (FormUID = "FrmDownTimeEntry") And (pVal.BeforeAction = False) Then
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE) Then
                    SBO_Application = Nothing
                    GC.Collect()
                End If
                '**********ChooseFromList Event is called using the raiseevent*********
                If (FormUID = "FrmDownTimeEntry") And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST And pVal.BeforeAction = False Then
                    ChooseFromListEvent = pVal
                    RaiseEvent ChooseFromList(pVal.ItemUID, pVal.ColUID, pVal.Row, ChooseFromListEvent.ChooseFromListUID, ChooseFromListEvent.SelectedObjects)
                End If
                '**********Realligning the values when the form is resized***********
                If pVal.FormUID = "FrmDownTimeEntry" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_RESIZE And pVal.BeforeAction = False And pVal.InnerEvent = False Then
                    Form_Resize()
                End If
                '************** Employee Combo Select **************
                If pVal.ItemUID = "CmbEmpId" And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT) Then
                    oParentDB.SetValue("U_empname", oParentDB.Offset, oEmpCombo.Selected.Description)
                End If
                '********* Work Center Combo Select **********
                If (pVal.ItemUID = "txtwcname") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT) Then
                    oRSMatrix.Clear()
                    oRMMatrix.Clear()
                    ' oWCCodeTxt.Value = oWCCombo.Selected.Description
                    oParentDB.SetValue("U_wccode", oParentDB.Offset, oWCCombo.Selected.Description)

                    ''******* Adding Empty Row to the Specification Matrix *******
                    If Len(oWCCombo.Selected.Value) > 0 Then
                        oMRDB.InsertRecord(oMRDB.Size)
                        oMRDB.Offset = oMRDB.Size - 1
                        MRSetValue()
                        oMRMatrix.AddRow(1, oMRMatrix.RowCount)
                        SetCFLConditions()
                    End If
                End If
                'Modified by Manimaran------s
                'If (pVal.ItemUID = "TxtCmpTime") And pVal.CharPressed = Keys.Tab And pVal.BeforeAction = False Then
                '    Dim StpInMins As String
                '    Try
                '        StpInMins = StPInMinsCalculation()
                '        oStpMtTxt.Value = StpInMins
                '        oParentDB.SetValue("U_totmts", oParentDB.Offset, oStpMtTxt.Value)
                '    Catch ex As Exception
                '        Throw ex
                '    End Try
                'End If
                If (pVal.ColUID = "ReaStTime" Or pVal.ColUID = "ReaEndTime") And pVal.ItemUID = "MatReaBD" And pVal.CharPressed = Keys.Tab And pVal.BeforeAction = False Then
                    Dim StpInMins As String
                    Try
                        Dim oFromTime, oToTime As DateTime
                        Dim oIntDt, oStDt As Date
                        Dim oIntTime, oStTime As DateTime
                        Dim oStpTime As SAPbouiCOM.EditText
                        oIntDt = Convert.ToDateTime(Date.Parse(oIntDtTxt.String)) ' String2Date(oIntDtTxt.String, "DD/MM/YY")
                        If shiftTimeValidation(pVal) = True Then
                            If pVal.ColUID = "ReaEndTime" = True Then
                                If validateTime(oIntDt) = True Then
                                    oIntTime = CDate(oRSMatrix.Columns.Item("ReaStTime").Cells.Item(oRSMatrix.RowCount).Specific.string)
                                    oFromTime = New Date(oIntDt.Year, oIntDt.Month, oIntDt.Day, oIntTime.Hour, oIntTime.Minute, oIntTime.Second)
                                    oStDt = Convert.ToDateTime(Date.Parse(oIntDtTxt.String)) 'String2Date(oCmpltDtTxt.String, "DD/MM/YY") 
                                    oStTime = CDate(oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.string)
                                    oToTime = New Date(oStDt.Year, oStDt.Month, oStDt.Day, oStTime.Hour, oStTime.Minute, oStTime.Second)
                                    StpInMins = DateDiff(DateInterval.Minute, oFromTime, oToTime)
                                    oStpTime = oRSMatrix.Columns.Item("ReaStp").Cells.Item(oRSMatrix.RowCount).Specific

                                    Try
                                        oStpTime.String = StpInMins
                                    Catch ex As Exception
                                    End Try
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        Throw ex
                    End Try
                End If
                'Modified by Manimaran------e
                '***** Reloads the Combo's if Define New is selected and data added in the Forms *****
                If (pVal.FormTypeEx = "FrmDownTimeEntry") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE) Then
                    Try
                        ReaDFN()
                        ResDFN()
                    Catch ex As Exception
                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    End Try
                End If

                '******* Add Next Row If the Reason Code is Selected **********
                If (pVal.ItemUID = "MatReaBD") And (pVal.ColUID = "ReaCode") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT) And pVal.Row > 0 Then
                    Dim oRSCombo As SAPbouiCOM.ComboBox
                    Dim oRSEdit As SAPbouiCOM.EditText
                    Dim oRSCode, oRSName As String
                    Dim CurrentRow As Integer
                    oRSCombo = oRsCodeCol.Cells.Item(pVal.Row).Specific
                    oRSEdit = oRsNameCol.Cells.Item(pVal.Row).Specific
                    oRSCode = oRSCombo.Selected.Value
                    oRSEdit.Value = oRSCombo.Selected.Description
                    oRSName = oRSEdit.Value
                    '****  Reason Code Combo Define New Selection *****
                    If oRSCombo.Selected.Value = "Define New" Then
                        LoadDefaultForm("PSSIT_RES")
                        BubbleEvent = False
                        BoolRes = False
                        oRSComboRow = pVal.Row
                    End If
                    CurrentRow = pVal.Row
                    If CurrentRow = oRSMatrix.VisualRowCount Then
                        oRSDB.Offset = oRSDB.Size - 1
                        RSSetValue()
                        oRSMatrix.SetLineData(CurrentRow)
                        oRSMatrix.FlushToDataSource()
                    End If
                    oRSDB.SetValue("U_reascode", oRSDB.Offset, oRSCode)
                    oRSDB.SetValue("U_reasdesc", oRSDB.Offset, oRSName)
                    oRSMatrix.SetLineData(CurrentRow)
                    oRSMatrix.FlushToDataSource()
                End If

                '******* Add Next Row If the Remedies Code is Selected **********
                If (pVal.ItemUID = "MatResBD") And (pVal.ColUID = "RemCode") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT) And pVal.Row > 0 Then
                    Dim oRMCombo As SAPbouiCOM.ComboBox
                    Dim oRMEdit As SAPbouiCOM.EditText
                    Dim oRMCode, oRMName As String
                    Dim CurrentRow As Integer
                    oRMCombo = oRMCodeCol.Cells.Item(pVal.Row).Specific
                    oRMEdit = oRMNameCol.Cells.Item(pVal.Row).Specific
                    oRMCode = oRMCombo.Selected.Value
                    oRMEdit.Value = oRMCombo.Selected.Description
                    oRMName = oRMEdit.Value
                    '****  Remedies Code Combo Define New Selection *****
                    If oRMCombo.Selected.Value = "Define New" Then
                        LoadDefaultForm("PSSIT_REMEDIES")
                        BubbleEvent = False
                        BoolRem = False
                        oRMComboRow = pVal.Row
                    End If
                    CurrentRow = pVal.Row
                    If CurrentRow = oRMMatrix.VisualRowCount Then
                        oRMDB.Offset = oRMDB.Size - 1
                        RMSetValue()
                        oRMMatrix.SetLineData(CurrentRow)
                        oRMMatrix.FlushToDataSource()
                    End If
                    oRMDB.SetValue("U_remecode", oRMDB.Offset, oRMCode)
                    oRMDB.SetValue("U_remedesc", oRMDB.Offset, oRMName)
                    oRMMatrix.SetLineData(CurrentRow)
                    oRMMatrix.FlushToDataSource()
                End If
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK) And pVal.BeforeAction = False Then
                    If (pVal.ItemUID = "MatReaBD") And pVal.ColUID = "#" Then
                        ReaUID = pVal.ItemUID
                        ResUID = ""
                        MatUID = ""
                    End If
                    If (pVal.ItemUID = "MatResBD") And pVal.ColUID = "#" Then
                        ResUID = pVal.ItemUID
                        ReaUID = ""
                        MatUID = ""
                    End If
                    If (pVal.ItemUID = "MatMR") And pVal.ColUID = "#" Then
                        MatUID = pVal.ItemUID
                        ResUID = ""
                        ReaUID = ""
                    End If
                End If
                '******* Value Calculation in Materials Replaced Matrix **********
                If (pVal.ItemUID = "MatMR") And (pVal.ColUID = "Qty") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) Then
                    Dim oQtyEdit, oRateEdit, oValEdit As SAPbouiCOM.EditText
                    Dim oQty, oRate, oVal As String
                    Dim CurrentRow As Integer
                    oQtyEdit = oMRQtyCol.Cells.Item(pVal.Row).Specific
                    oRateEdit = oMRRateCol.Cells.Item(pVal.Row).Specific
                    oValEdit = oMRValueCol.Cells.Item(pVal.Row).Specific
                    oQty = oQtyEdit.Value
                    oRate = oRateEdit.Value
                    oValEdit.Value = oQty * oRate
                    oVal = oValEdit.Value

                    CurrentRow = pVal.Row
                    ' For CurrentRow = 1 To oMRMatrix.RowCount
                    oMRMatrix.GetLineData(CurrentRow)
                    oMRDB.SetValue("U_itemval", oMRDB.Offset, oVal)
                    oForm.Update()
                    ' Next
                End If
                If pVal.ItemUID = "lnkmc" And pVal.BeforeAction = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    Dim oMCCode As String
                    oMCCode = oDeptCodeTxt.Value
                    MachineMasterClass = New MachineMaster(SBO_Application, oCompany, oMCCode, "DownTime")
                End If
                If pVal.ItemUID = "lnksft" And pVal.BeforeAction = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    Dim oSftCode As String
                    oSftCode = oShiftCodeTxt.Value
                    ShiftMaster = New Shift(SBO_Application, oCompany, oSftCode, "DownTime")
                End If
                '********** Add Button Press ***********
                If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    ' LoadDocNumber()
                    oForm.Refresh()
                    oSerCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                    With oForm.DataSources.DBDataSources.Item("@PSSIT_PMWCBREAKHDR")
                        .SetValue("DocNum", .Offset, oForm.BusinessObject.GetNextSerialNumber(Trim(.GetValue("Series", .Offset))).ToString)
                    End With
                    oDocDtTxt.String = System.DateTime.Today.Date.ToString("dd/MM/yyyy")
                    oIntDtTxt.String = System.DateTime.Today.Date.ToString("dd/MM/yyyy")
                    'Commented by Manimaran------s
                    'oIntTimeTxt.String = FormatDateTime(Now(), DateFormat.ShortTime)
                    'oAttDtTxt.String = System.DateTime.Today.Date.ToString("dd/MM/yyyy")
                    'oAttTimeTxt.String = FormatDateTime(Now(), DateFormat.ShortTime)
                    'oCmpltDtTxt.String = System.DateTime.Today.Date.ToString("dd/MM/yyyy")
                    'oCmpltTimeTxt.String = FormatDateTime(Now(), DateFormat.ShortTime)
                    'Commented by Manimaran------e
                End If
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
        '************ Add Mode Item Press **********
        If (FormUID = "FrmDownTimeEntry") And (pVal.BeforeAction = True) And pVal.ItemUID = "1" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
            Try
                '********* Delete Empty Rows In the matrix **********
                If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    Try
                        RSDeleteEmptyRow()
                        RMDeleteEmptyRow()
                        MRDeleteEmptyRow()
                        'DeleteEmptyRows()
                        Validation()
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
        'Added by Manimaran----------S
        If pVal.Before_Action = False Then
            If pVal.ItemUID = "1000001" Then
                oForm.PaneLevel = 1
            End If
            If pVal.ItemUID = "40" Then
                oForm.PaneLevel = 2
            End If
        End If
        If (pVal.ItemUID = "TxtIntTime" Or pVal.ItemUID = "TxtCmpTime") And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS And pVal.BeforeAction = False Then
            'Dim QToTime As Integer
            'Dim QFrtime As Integer
            'If CInt(oParentDB.GetValue("U_intime", oParentDB.Offset).Trim().Length) > 0 Then
            '    ofrTime = CInt(oParentDB.GetValue("U_intime", oParentDB.Offset).Trim())
            'End If
            'If CInt(oParentDB.GetValue("U_comptime", oParentDB.Offset).Trim().Length) > 0 Then
            '    oToTime = CInt(oParentDB.GetValue("U_comptime", oParentDB.Offset).Trim())
            'End If

            'sQry = "Select * from [@PSSIT_OSFT] where code = '" & oForm.Items.Item("txtSftCode").Specific.string & "'"
            'Rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'Rs.DoQuery(sQry)
            'If Rs.RecordCount > 0 Then
            '    QFrtime = Integer.Parse(Rs.Fields.Item("U_Sftime").Value.ToString)
            '    QToTime = Integer.Parse(Rs.Fields.Item("U_Sttime").Value.ToString)
            'End If
            'If pVal.ItemUID = "TxtIntTime" Then
            '    If QFrtime < QToTime Then
            '        If ofrTime < QFrtime Then
            '            SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            '            oParentDB.SetValue("U_intime", oParentDB.Offset, "0")
            '        ElseIf ofrTime > QToTime Then
            '            SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            '            oParentDB.SetValue("U_comptime", oParentDB.Offset, "0")
            '        End If
            '    Else
            '        If QToTime <= ofrTime Then
            '            If ofrTime < QFrtime Then
            '                SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range  ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            '                oParentDB.SetValue("U_intime", oParentDB.Offset, "0")
            '            ElseIf ofrTime < QToTime Then
            '                SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            '                oParentDB.SetValue("U_comptime", oParentDB.Offset, "0")
            '            End If
            '        End If
            '    End If
            'End If
            'If pVal.ItemUID = "TxtCmpTime" Then
            '    If QFrtime < QToTime Then
            '        If oToTime < QFrtime Then
            '            SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range  ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            '            oParentDB.SetValue("U_comptime", oParentDB.Offset, "0")
            '        ElseIf oToTime > QToTime Then
            '            SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            '            oParentDB.SetValue("U_comptime", oParentDB.Offset, "0")
            '        End If
            '    Else
            '        If ofrTime > oToTime Then
            '            If oToTime > QFrtime Then
            '                SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            '                oParentDB.SetValue("U_comptime", oParentDB.Offset, "0")
            '            ElseIf oToTime > QToTime Then
            '                SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            '                oParentDB.SetValue("U_comptime", oParentDB.Offset, "0")
            '            End If
            '        Else
            '            If ofrTime < oToTime And QToTime < oToTime And QFrtime > ofrTime Then
            '                SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range  ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            '                oParentDB.SetValue("U_comptime", oParentDB.Offset, "0")
            '            End If
            '        End If
            '    End If
            'End If
        End If
        'Added by Manimaran----------E
    End Sub
    'Added by Manimaran -------s
    Private Function validateTime(ByVal dt As Date) As Boolean
        If oForm.Items.Item("txtmcno").Specific.string = "" Or oForm.Items.Item("txtSftCode").Specific.string = "" Or oForm.Items.Item("41").Specific.string = "" Then
            SBO_Application.SetStatusBarMessage("Enter Mandatory fields", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return False
        End If
        Dim qry As String
        Dim rs As SAPbobsCOM.Recordset
        rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'If dt.Day <= 12 Then
        '    qry = "select t2.u_frtime,t2.u_totime from [@PSSIT_OPEY] t1"
        '    qry = qry + " inner join [@PSSIT_PEY1] t2 on t1.docentry = t2.docentry"
        '    qry = qry + " where t1.u_scode = '" & oForm.Items.Item("txtSftCode").Specific.string & "' and t2.u_wcno = '" & oForm.Items.Item("txtmcno").Specific.string & "' and  convert(varchar,t1.u_docdt,101) = '" & Left(String.Format(dt, "dd/mm/yyyy"), 6) & dt.Year & "'"
        'Else
        '    qry = "select t2.u_frtime,t2.u_totime from [@PSSIT_OPEY] t1"
        '    qry = qry + " inner join [@PSSIT_PEY1] t2 on t1.docentry = t2.docentry"
        '    qry = qry + " where t1.u_scode = '" & oForm.Items.Item("txtSftCode").Specific.string & "' and t2.u_wcno = '" & oForm.Items.Item("txtmcno").Specific.string & "' and  convert(varchar,t1.u_docdt,103) = '" & Left(String.Format(dt, "dd/mm/yyyy"), 6) & dt.Year & "'"
        'End If
        'Modified by Kabilahan-------s
        'qry = "select t2.u_frtime,t2.u_totime from [@PSSIT_OPEY] t1"
        'qry = qry + " inner join [@PSSIT_PEY1] t2 on t1.docentry = t2.docentry"
        'qry = qry + " where t1.u_scode = '" & oForm.Items.Item("txtSftCode").Specific.string & "' and t2.u_wcno = '" & oForm.Items.Item("txtmcno").Specific.string & "'"

        qry = "set dateformat dmy select t2.u_frtime,t2.u_totime from [@PSSIT_OPEY] t1"
        qry = qry + " inner join [@PSSIT_PEY1] t2 on t1.docentry = t2.docentry"
        qry = qry + " where t1.u_scode = '" & oForm.Items.Item("txtSftCode").Specific.string & "' and t2.u_wcno = '" & oForm.Items.Item("txtmcno").Specific.string & "' and t1.U_Docdt = convert(varchar,'" & oForm.Items.Item("TxtIntDt").Specific.string & "',102)"
        'Added by manimaran-----s
        qry = qry + " and t1.u_pnordno = '" & oForm.Items.Item("41").Specific.string & "'"
        'added by Manimaran-----e
        'Modified by Kabilahan-------e
        rs.DoQuery(qry)
        If rs.RecordCount > 0 Then
            While Not rs.EoF
                If rs.Fields.Item(0).Value < rs.Fields.Item(1).Value Then
                    If (oRSMatrix.Columns.Item("ReaStTime").Cells.Item(oRSMatrix.RowCount).Specific.value >= rs.Fields.Item(0).Value Or oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.value >= rs.Fields.Item(0).Value) Or (oRSMatrix.Columns.Item("ReaStTime").Cells.Item(oRSMatrix.RowCount).Specific.value >= rs.Fields.Item(1).Value Or oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.value >= rs.Fields.Item(1).Value) Then
                        If (rs.Fields.Item(0).Value > oRSMatrix.Columns.Item("ReaStTime").Cells.Item(oRSMatrix.RowCount).Specific.value) Or (rs.Fields.Item(1).Value > oRSMatrix.Columns.Item("ReaStTime").Cells.Item(oRSMatrix.RowCount).Specific.value) Or (rs.Fields.Item(0).Value > oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.value) Or (rs.Fields.Item(1).Value > oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.value) Then
                            SBO_Application.SetStatusBarMessage("Overlapping with the Machine run time", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Return False
                        End If
                    ElseIf (oRSMatrix.Columns.Item("ReaStTime").Cells.Item(oRSMatrix.RowCount).Specific.value <= rs.Fields.Item(0).Value Or oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.value <= rs.Fields.Item(0).Value) Or (oRSMatrix.Columns.Item("ReaStTime").Cells.Item(oRSMatrix.RowCount).Specific.value <= rs.Fields.Item(1).Value Or oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.value <= rs.Fields.Item(1).Value) Then
                        If (rs.Fields.Item(0).Value < oRSMatrix.Columns.Item("ReaStTime").Cells.Item(oRSMatrix.RowCount).Specific.value) Or (rs.Fields.Item(1).Value < oRSMatrix.Columns.Item("ReaStTime").Cells.Item(oRSMatrix.RowCount).Specific.value) Or (rs.Fields.Item(0).Value < oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.value) Or (rs.Fields.Item(1).Value < oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.value) Then
                            SBO_Application.SetStatusBarMessage("Overlapping with the Machine run time", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Return False
                        End If
                    End If
                Else
                    If (oRSMatrix.Columns.Item("ReaStTime").Cells.Item(oRSMatrix.RowCount).Specific.value <= rs.Fields.Item(0).Value Or oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.value <= rs.Fields.Item(0).Value) Or (oRSMatrix.Columns.Item("ReaStTime").Cells.Item(oRSMatrix.RowCount).Specific.value <= rs.Fields.Item(1).Value Or oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.value <= rs.Fields.Item(1).Value) Then
                        If (rs.Fields.Item(0).Value < oRSMatrix.Columns.Item("ReaStTime").Cells.Item(oRSMatrix.RowCount).Specific.value) Or (rs.Fields.Item(1).Value > oRSMatrix.Columns.Item("ReaStTime").Cells.Item(oRSMatrix.RowCount).Specific.value) Or (rs.Fields.Item(0).Value < oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.value) Or (rs.Fields.Item(1).Value > oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.value) Then
                            SBO_Application.SetStatusBarMessage("Overlapping with the Machine run time", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Return False
                        End If
                    ElseIf (oRSMatrix.Columns.Item("ReaStTime").Cells.Item(oRSMatrix.RowCount).Specific.value >= rs.Fields.Item(0).Value Or oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.value >= rs.Fields.Item(0).Value) Or (oRSMatrix.Columns.Item("ReaStTime").Cells.Item(oRSMatrix.RowCount).Specific.value >= rs.Fields.Item(1).Value Or oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.value >= rs.Fields.Item(1).Value) Then
                        If (rs.Fields.Item(0).Value > oRSMatrix.Columns.Item("ReaStTime").Cells.Item(oRSMatrix.RowCount).Specific.value) Or (rs.Fields.Item(1).Value < oRSMatrix.Columns.Item("ReaStTime").Cells.Item(oRSMatrix.RowCount).Specific.value) Or (rs.Fields.Item(0).Value > oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.value) Or (rs.Fields.Item(1).Value < oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.value) Then
                            SBO_Application.SetStatusBarMessage("Overlapping with the Machine run time", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Return False
                        End If
                    End If
                End If
                rs.MoveNext()
            End While
        Else
            SBO_Application.SetStatusBarMessage("Machine Master Record not found", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return False
        End If

        Return True
    End Function
    Private Function shiftTimeValidation(ByRef pVal As SAPbouiCOM.ItemEvent) As Boolean
        Dim QToTime As Integer
        Dim QFrtime As Integer
        If oRSMatrix.Columns.Item("ReaStTime").Cells.Item(oRSMatrix.RowCount).Specific.string <> "" Then
            ofrTime = CInt(oRSMatrix.Columns.Item("ReaStTime").Cells.Item(oRSMatrix.RowCount).Specific.value)
        End If
        If oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.string <> "" Then
            oToTime = CInt(oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.value)
        End If

        sQry = "Select * from [@PSSIT_OSFT] where code = '" & oForm.Items.Item("txtSftCode").Specific.string & "'"
        Rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Rs.DoQuery(sQry)
        If Rs.RecordCount > 0 Then
            QFrtime = Integer.Parse(Rs.Fields.Item("U_Sftime").Value.ToString)
            QToTime = Integer.Parse(Rs.Fields.Item("U_Sttime").Value.ToString)
        End If
        If pVal.ColUID = "ReaStTime" Then
            If QFrtime < QToTime Then
                If ofrTime < QFrtime Then
                    SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    oRSMatrix.Columns.Item("ReaStTime").Cells.Item(oRSMatrix.RowCount).Specific.value = 0
                    Return False
                ElseIf ofrTime > QToTime Then
                    SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.value = 0
                    Return False
                End If
            Else
                If QToTime <= ofrTime Then
                    If ofrTime < QFrtime Then
                        SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range  ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oRSMatrix.Columns.Item("ReaStTime").Cells.Item(oRSMatrix.RowCount).Specific.value = 0
                        Return False
                    ElseIf ofrTime < QToTime Then
                        SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.value = 0
                        Return False
                    End If
                End If
            End If
        End If
        If pVal.ColUID = "ReaEndTime" Then
            If QFrtime < QToTime Then
                If oToTime < QFrtime Then
                    SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range  ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.value = 0
                    Return False
                ElseIf oToTime > QToTime Then
                    SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.value = 0
                    Return False
                End If
            Else
                If ofrTime > oToTime Then
                    If oToTime > QFrtime Then
                        SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.value = 0
                        Return False
                    ElseIf oToTime > QToTime Then
                        SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.value = 0
                        Return False
                    End If
                Else
                    If ofrTime < oToTime And QToTime < oToTime And QFrtime > ofrTime Then
                        SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range  ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oRSMatrix.Columns.Item("ReaEndTime").Cells.Item(oRSMatrix.RowCount).Specific.value = 0
                        Return False
                    End If
                End If
            End If
        End If
        Return True
    End Function
    'Added by Manimaran -------e
    Private Sub Validation()
        Try
            '**************** Mandatory **************
            If oWCNameTxt.Value = "" Or oWCNameTxt.Value = Nothing Then
                oWCNameTxt.Active = True
                Throw New Exception("Enter Machine Detail")
            End If
            ' *************Intimation Date Validation **********
            If oIntDtTxt.String > System.DateTime.Today.Date.ToString("dd/MM/yyyy") Then
                oIntDtTxt.Active = True
                Throw New Exception("Intimation Date Should Not be greater than the Current Date")
            End If
            'Added by Manimaran----------s
            Dim i As Integer
            If oRSMatrix.RowCount > 0 Then
                For i = 1 To oRSMatrix.RowCount
                    If oRSMatrix.Columns.Item("ReaStTime").Cells.Item(i).Specific.string <> "" And oRSMatrix.Columns.Item("ReaStp").Cells.Item(i).Specific.string = "" Then
                        Throw New Exception("Stoppage Time is missing")
                    End If
                Next
            Else
                Throw New Exception("Enter atleast one reason")
            End If

            'Added by Manimaran----------e
            'Commented by Manimaran------s
            ' *************Attended Date Validation **********
            'If oAttDtTxt.String > System.DateTime.Today.Date.ToString("dd/MM/yyyy") Then
            '    oAttDtTxt.Active = True
            '    Throw New Exception("Attended Date Should Not be greater than the Current Date")
            'End If
            'If oAttDtTxt.String > oCmpltDtTxt.String Then
            '    oAttDtTxt.Active = True
            '    Throw New Exception("Attended Date Should Not be greater than the Completed Date")
            'End If
            ' *************Completed Date Validation **********
            'If oCmpltDtTxt.String > System.DateTime.Today.Date.ToString("dd/MM/yyyy") Then
            '    oCmpltDtTxt.Active = True
            '    Throw New Exception("Completed Date Should Not be greater than the Current Date")
            'End If
            ' *************Attended Time Validation **********
            'If oAttTimeTxt.String > oCmpltTimeTxt.String Then
            '    oAttTimeTxt.Active = True
            '    Throw New Exception("Attended Time Should Not be greater than the Completed Time")
            'End If
            'ofrTime = CInt(oParentDB.GetValue("U_intime", oParentDB.Offset).Trim())
            'oToTime = CInt(oParentDB.GetValue("U_comptime", oParentDB.Offset).Trim())
            'If ofrTime = 0 Or oToTime = 0 Then
            '    Throw New Exception("Stoppage time or Started time should not be 0")
            'End If
            'Commented by Manimaran------e
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' Loading the document number from the document numbering tables.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadDocNumber()
        Dim StrSql As String
        Dim oRs As SAPbobsCOM.Recordset
        Try
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            StrSql = "Select ObjectCode,Series,SeriesName,InitialNum,NextNumber,LastNum From NNM1 " _
                       & "Where NNM1.ObjectCode='PSSIT_WCBREAK'"
            oRs.DoQuery(StrSql)
            If oRs.RecordCount > 0 Then
                oRs.MoveFirst()
                oDocNoTxt.Value = oRs.Fields.Item("NextNumber").Value
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' This is used to refresh the form
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub RefreshForm()
        Try
            Dim f As SAPbouiCOM.Form
            f = SBO_Application.Forms.Item("FrmDownTimeEntry")
            f.Refresh()
        Catch ex As Exception
        End Try
    End Sub
    ''' <summary>
    ''' Stopage In Mins Calulation
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function StPInMinsCalculation() As String
        Dim oFromTime, oToTime As DateTime
        Dim oIntDt, oStDt As Date
        Dim oIntTime, oStTime As DateTime
        Dim min As Integer
        Try

            oIntDt = Convert.ToDateTime(Date.Parse(oIntDtTxt.String)) ' String2Date(oIntDtTxt.String, "DD/MM/YY")
            oIntTime = Convert.ToDateTime(Date.Parse(oIntTimeTxt.String))
            oFromTime = New Date(oIntDt.Year, oIntDt.Month, oIntDt.Day, oIntTime.Hour, oIntTime.Minute, oIntTime.Second)
            oStDt = Convert.ToDateTime(Date.Parse(oIntDtTxt.String)) 'String2Date(oCmpltDtTxt.String, "DD/MM/YY") 
            oStTime = Convert.ToDateTime(Date.Parse(oCmpltTimeTxt.String))
            oToTime = New Date(oStDt.Year, oStDt.Month, oStDt.Day, oStTime.Hour, oStTime.Minute, oStTime.Second)

            min = DateDiff(DateInterval.Minute, oFromTime, oToTime)
        Catch ex As Exception
            Throw ex
        End Try
        Return min
    End Function
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Dim FormID As String
        FormID = SBO_Application.Forms.ActiveForm.UniqueID
        Try
            If pVal.BeforeAction = False Then
                If pVal.MenuUID = "1282" Then
                    'LoadDocNumber()
                    oForm.Refresh()
                    oSerCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                    With oForm.DataSources.DBDataSources.Item("@PSSIT_PMWCBREAKHDR")
                        .SetValue("DocNum", .Offset, oForm.BusinessObject.GetNextSerialNumber(Trim(.GetValue("Series", .Offset))).ToString)
                    End With
                    oDocDtTxt.String = System.DateTime.Today.Date.ToString("dd/MM/yyyy")
                    oIntDtTxt.String = System.DateTime.Today.Date.ToString("dd/MM/yyyy")
                    'Commented by Manimaran------s
                    'oIntTimeTxt.String = FormatDateTime(Now(), DateFormat.ShortTime)
                    'oAttDtTxt.String = System.DateTime.Today.Date.ToString("dd/MM/yyyy")
                    'oAttTimeTxt.String = FormatDateTime(Now(), DateFormat.ShortTime)
                    'oCmpltDtTxt.String = System.DateTime.Today.Date.ToString("dd/MM/yyyy")
                    'oCmpltTimeTxt.String = FormatDateTime(Now(), DateFormat.ShortTime)
                    'Commented by Manimaran------e
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try

        '*****************************LoadParameterData() is called.*******************************
        If (pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") And pVal.BeforeAction = False And FormID = "FrmDownTimeEntry" Then
            oForm.Freeze(True)
            '  LoadParameterData()
            oForm.Freeze(False)
        End If
        '*************** Delete Row in Matrix ******************
        If pVal.MenuUID = "1293" And pVal.BeforeAction = True Then
            '*************Specification Matrix**************
            If ReaUID = "MatReaBD" Then
                Try
                    'oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    oRSDB = oForm.DataSources.DBDataSources.Item("@PSSIT_PMWCREASONDTL")

                    oRSMatrix.DeleteRow(oRSMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                    oRSMatrix.FlushToDataSource()

                    If (oRSMatrix.RowCount = 0) Then
                        oRSDB.RemoveRecord(0)
                    End If
                    BubbleEvent = False
                Catch ex As Exception
                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End Try
            ElseIf ResUID = "MatResBD" Then
                '*************Specification Matrix**************
                ' If ResUID = "MatResBD" Then
                Try
                    'oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    oRMDB = oForm.DataSources.DBDataSources.Item("@PSSIT_PMWCREMEDTL")

                    oRMMatrix.DeleteRow(oRMMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                    oRMMatrix.FlushToDataSource()

                    If (oRMMatrix.RowCount = 0) Then
                        oRMDB.RemoveRecord(0)
                    End If
                    BubbleEvent = False
                Catch ex As Exception
                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End Try
            ElseIf MatUID = "MatMR" Then
                '*************Specification Matrix**************
                'If MatUID = "MatMR" Then
                Try
                    'oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    oMRDB = oForm.DataSources.DBDataSources.Item("@PSSIT_PMWCITEMSDTL")

                    oMRMatrix.DeleteRow(oMRMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                    oMRMatrix.FlushToDataSource()

                    If (oMRMatrix.RowCount = 0) Then
                        oMRDB.RemoveRecord(0)
                    End If
                    BubbleEvent = False
                Catch ex As Exception
                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End Try
            End If
            End If
        '********************************Add Row************************88

        If pVal.MenuUID = "1292" And pVal.BeforeAction = True Then
            If ReaUID = "MatReaBD" Then
                Try
                    ' oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    oRSDB.InsertRecord(oRSDB.Size)
                    oRSDB.Offset = oRSDB.Size - 1
                    RSSetValue()
                    oRSMatrix.AddRow(1, oRSMatrix.RowCount)
                    oForm.Update()
                    ReaUID = ""
                Catch ex As Exception
                    Throw ex
                End Try
            ElseIf ResUID = "MatResBD" Then
                Try
                    ' oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    oRMDB.InsertRecord(oRMDB.Size)
                    oRMDB.Offset = oRMDB.Size - 1
                    RMSetValue()
                    oRMMatrix.AddRow(1, oRMMatrix.RowCount)
                    oForm.Update()
                    ResUID = ""
                Catch ex As Exception
                    Throw ex
                End Try
            ElseIf MatUID = "MatMR" Then
                Try
                    'oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    oMRDB.InsertRecord(oMRDB.Size)
                    oMRDB.Offset = oMRDB.Size - 1
                    MRSetValue()
                    oMRMatrix.AddRow(1, oMRMatrix.RowCount)
                    SetCFLConditions()
                    oForm.Update()
                    MatUID = ""
                Catch ex As Exception
                    Throw ex
                End Try
            End If
        End If
    End Sub
    Private Sub DeleteEmptyRows()
        Dim oRS, oRS1, oRS2 As SAPbobsCOM.Recordset
        Try
            oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("Delete from [@PSSIT_PMWCREASONDTL] where U_reascode is null")
            oRS1.DoQuery("Delete from [@PSSIT_PMWCREMEDTL] where U_remecode is null")
            oRS2.DoQuery("Delete from [@PSSIT_PMWCITEMSDTL] where U_itemcode is null")
        Catch ex As Exception
            Throw ex
        Finally
            oRS = Nothing
            oRS1 = Nothing
            oRS2 = Nothing
            GC.Collect()
        End Try
    End Sub
    Function String2Date(ByVal S As String, _
                           ByVal Fmt As String) As Object
        Select Case Fmt
            Case "MMDDYY", "MMDDYYYY"      '052793   05271993
                String2Date = CDate(Left(S, 2) & "/" & Mid(S, 3, 2) & "/" & _
                                    Mid(S, 5))
            Case "DDMMYY", "DDMMYYYY"      '270593   27051993
                String2Date = CDate(Mid(S, 3, 2) & "/" & Left(S, 2) & "/" & _
                                    Mid(S, 5))
            Case "YYMMDD"                  '930527
                String2Date = CDate(Mid(S, 3, 2) & "/" & Right(S, 2) & "/" & _
                                    Left(S, 2))
            Case "YYYYMMDD"                '19930527
                String2Date = CDate(Mid(S, 5, 2) & "/" & Right(S, 2) & "/" & _
                                    Left(S, 4))
            Case "MM/DD/YY", "MM/DD/YYYY", "M/D/Y", "M/D/YY", "M/D/YYYY", _
                 "DD-MMM-YY", "DD-MMM-YYYY"
                String2Date = CDate(S)
            Case "DD/MM/YY", "DD/MM/YYYY"  '27/05/93   27/05/1993
                String2Date = CDate(Mid(S, 4, 3) & Left(S, 3) & Mid(S, 7))
            Case "YY/MM/DD"                '93/05/27
                String2Date = CDate(Mid(S, 4, 3) & Right(S, 2) & _
                                    "/" & Left(S, 2))
            Case "YYYY/MM/DD"              '1993/05/27
                String2Date = CDate(Mid(S, 6, 3) & Right(S, 2) & _
                                    "/" & Left(S, 4))
            Case Else
                String2Date = Nothing
        End Select
    End Function
End Class