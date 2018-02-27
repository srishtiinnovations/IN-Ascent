'''' <summary>
'''' Author                     Created Date
'''' Suresh                    16/11/2009
'''' <remarks> This class is used for entering the Work Centre details.</remarks>
Public Class WorkCentre
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
    Private oParentDB, oChildDB, oFCostDB As SAPbouiCOM.DBDataSource
    '**************************Form************************************
    Private oForm As SAPbouiCOM.Form
    '**************************ChooseFromList************************************
    Private oChCurList, oChAccList As SAPbouiCOM.ChooseFromList
    Private Event ChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable)
    '**************************Items - EditText************************************
    Private oWCCodeTxt, oWCNameTxt, oInfo1Txt, oInfo2Txt, oRemarksTxt As SAPbouiCOM.EditText
    '**************************Items - ComboBox************************************
    Private oWCTypeCombo As SAPbouiCOM.ComboBox
    '**************************Items - CheckBox************************************
    Private oActiveCheck As SAPbouiCOM.CheckBox
    '**************************Items - LinkButton************************************
    Private oAcctLinkBtn As SAPbouiCOM.LinkedButton
    '**************************Items - Matrix************************************
    Private oCostMatrix As SAPbouiCOM.Matrix
    Private oColumns As SAPbouiCOM.Columns
    Private oRowNoCol, oFxdCostCol, oCurrencyCol, oUnitCostCol, oAbsMethodCol, oAcctCodeCol, oAcctNameCol, oActAcCodeCol, oInfo1Col As SAPbouiCOM.Column
    '**************************Boolean Variables for Define New************************************
    Private BoolWCType As Boolean = True
    Private BoolCostTyp As Boolean = True
    Private oCostComboRow As Integer
    Private FCostUID As String

    Private oWCCode As String
    Private oFormName As String
#End Region
    ''' <summary>
    ''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmWorkCenter.srf") method is called to load the Work Centre form.
    ''' Drawform() method is called to Initialize the form,Datasources and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company, Optional ByVal aWCCode As String = Nothing, Optional ByVal aFormName As String = Nothing)
        MyBase.New(aSBO_Application, aCompany)
        SBO_Application = aSBO_Application
        oCompany = aCompany
        oWCCode = aWCCode
        oFormName = aFormName
        LoadFromXML("FrmWorkCenter.srf")
        DrawForm()
        If oFormName = "SkillGroups" Or oFormName = "MachineGroups" Or oFormName = "Tools" Or oFormName = "MachineMaster" Then
            oParentDB.SetValue("Code", oParentDB.Offset, oWCCode)
            oWCCodeTxt.Value = oWCCode
            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        End If
        oForm.DataBrowser.BrowseBy = "txtwcode"
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
            oParentDB = oForm.DataSources.DBDataSources.Item("@PSSIT_OWCR")
            oFCostDB = oForm.DataSources.DBDataSources.Item("@PSSIT_WCR1")
            oForm.Freeze(True)
            InitializeFormComponent()
            LoadLookups()
            ConfigureMatrix()
            FxdCostTypCombo()
            AbsMethodCombo()
            oForm.EnableMenu("1292", True)
            oForm.EnableMenu("1293", True)
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
            oWCCodeTxt = oForm.Items.Item("txtwcode").Specific
            oWCCodeTxt.DataBind.SetBound(True, "@PSSIT_OWCR", "Code")

            oWCNameTxt = oForm.Items.Item("txtwname").Specific
            oWCNameTxt.DataBind.SetBound(True, "@PSSIT_OWCR", "U_WCname")

            oWCTypeCombo = oForm.Items.Item("cmbwtype").Specific
            oWCTypeCombo.DataBind.SetBound(True, "@PSSIT_OWCR", "U_WCtype")
            WCTypeCombo()

            oInfo1Txt = oForm.Items.Item("txtadnl1").Specific
            oForm.Items.Item("lbladnl1").Visible = False
            oForm.Items.Item("txtadnl1").Visible = False
            oInfo1Txt.DataBind.SetBound(True, "@PSSIT_OWCR", "U_Adnl1")

            oInfo2Txt = oForm.Items.Item("txtadnl2").Specific
            oForm.Items.Item("lbladnl2").Visible = False
            oForm.Items.Item("txtadnl2").Visible = False
            oInfo2Txt.DataBind.SetBound(True, "@PSSIT_OWCR", "U_Adnl2")

            oRemarksTxt = oForm.Items.Item("txtremark").Specific
            oRemarksTxt.DataBind.SetBound(True, "@PSSIT_OWCR", "U_Remarks")

            oActiveCheck = oForm.Items.Item("chkactive").Specific
            oActiveCheck.DataBind.SetBound(True, "@PSSIT_OWCR", "U_Active")
            oActiveCheck.Checked = True

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Configuring the Matrix items/controls in the form(.srf) by bounding to the object and to the DBDatasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ConfigureMatrix()
        Try
            oCostMatrix = oForm.Items.Item("matcost").Specific
            oCostMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oColumns = oCostMatrix.Columns

            oRowNoCol = oColumns.Item("#")
            oRowNoCol.Editable = False

            oFxdCostCol = oColumns.Item("colfcost")
            oFxdCostCol.Editable = True
            oFxdCostCol.DataBind.SetBound(True, "@PSSIT_WCR1", "U_Fcost")

            oCurrencyCol = oColumns.Item("colcurr")
            oCurrencyCol.Editable = False
            oCurrencyCol.DataBind.SetBound(True, "@PSSIT_WCR1", "U_Currency")
            FxtCstCurrCombo()

            oUnitCostCol = oColumns.Item("colucst")
            oUnitCostCol.Editable = True
            oUnitCostCol.Visible = True
            oUnitCostCol.DataBind.SetBound(True, "@PSSIT_WCR1", "U_UnitCost")

            oAbsMethodCol = oColumns.Item("colabmd")
            oAbsMethodCol.Editable = True
            oAbsMethodCol.Visible = False
            oAbsMethodCol.DataBind.SetBound(True, "@PSSIT_WCR1", "U_Absmthd")

            oAcctCodeCol = oColumns.Item("colaccod")
            oAcctCodeCol.Editable = True
            oAcctCodeCol.Visible = True
            oAcctCodeCol.DataBind.SetBound(True, "@PSSIT_WCR1", "U_Accode")
            oAcctCodeCol.ChooseFromListUID = "AccLst"
            oAcctCodeCol.ChooseFromListAlias = "AcctCode"
            oAcctLinkBtn = oAcctCodeCol.ExtendedObject
            oAcctLinkBtn.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_GLAccounts

            oAcctNameCol = oColumns.Item("colacnam")
            oAcctNameCol.Editable = False
            oAcctNameCol.Visible = True
            oAcctNameCol.DataBind.SetBound(True, "@PSSIT_WCR1", "U_Acname")

            oAcctCodeCol = oColumns.Item("accode")
            oAcctCodeCol.Editable = False
            oAcctCodeCol.Visible = True
            oAcctCodeCol.DataBind.SetBound(True, "@PSSIT_WCR1", "U_ActAcCode")

            oInfo1Col = oColumns.Item("coladnl1")
            oInfo1Col.Editable = True
            oInfo1Col.Visible = True
            oInfo1Col.DataBind.SetBound(True, "@PSSIT_WCR1", "U_Adnl1")

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
            oChAccList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "1", "AccLst"))
            CreateNewConditions(oChAccList, "Postable", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
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
    Private Sub WorkCenter_ChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable) Handles Me.ChooseFromList
        Dim oDataTable As SAPbouiCOM.DataTable = ChooseFromListSelectedObjects
        Dim oCurrentRow As Integer
        Dim oAccCode, oAccName As String
        Dim oRs As SAPbobsCOM.Recordset
        Try
            oCurrentRow = CurrentRow
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
           
            If (ControlName = "matcost") And (ChoosefromListUID = "AccLst") Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        If Not oDataTable Is Nothing Then
                            oAccCode = oDataTable.GetValue("FormatCode", 0)
                            oAccName = oDataTable.GetValue("AcctName", 0)
                            oCostMatrix.GetLineData(CurrentRow)
                            ' ******* Add Next Row If the Item Code is Selected **********
                            If CurrentRow = oCostMatrix.VisualRowCount Then
                                oFCostDB.Offset = oFCostDB.Size - 1
                                FxdCostSetValue()
                                oCostMatrix.SetLineData(CurrentRow)
                                oCostMatrix.FlushToDataSource()
                            End If
                            oFCostDB.SetValue("U_Accode", oFCostDB.Offset, FormatAccountCode(oAccCode))
                            oFCostDB.SetValue("U_Acname", oFCostDB.Offset, oAccName)
                            oFCostDB.SetValue("U_ActAcCode", oFCostDB.Offset, oAccCode.ToString().Replace("-", ""))
                            oCostMatrix.SetLineData(CurrentRow)
                            oCostMatrix.FlushToDataSource()
                          
                        End If
                    End If
                Else

                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        If Not oDataTable Is Nothing Then
                            oAccCode = oDataTable.GetValue("FormatCode", 0)
                            oAccName = oDataTable.GetValue("AcctName", 0)
                            ' ******* Add Next Row If the Item Code is Selected **********
                            If CurrentRow = oCostMatrix.VisualRowCount Then
                                oFCostDB.Offset = oFCostDB.Size - 1
                                FxdCostSetValue()
                                oCostMatrix.SetLineData(CurrentRow)
                                oCostMatrix.FlushToDataSource()
                            End If
                            oCostMatrix.GetLineData(CurrentRow)
                            oFCostDB.SetValue("U_Accode", oFCostDB.Offset, FormatAccountCode(oAccCode))
                            oFCostDB.SetValue("U_Acname", oFCostDB.Offset, oAccName)
                            oFCostDB.SetValue("U_ActAcCode", oFCostDB.Offset, oAccCode.ToString().Replace("-", ""))
                            oCostMatrix.SetLineData(CurrentRow)
                            oCostMatrix.FlushToDataSource()
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
    ''' This is used to Load the Work Centre Type in the Combo
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WCTypeCombo()
        Dim oRs As SAPbobsCOM.Recordset
        Try
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery("select Code,U_Wctypnam from [@PSSIT_OTYP] where code is not null") ''and U_Wctypnam is not Null")
            oRs.MoveFirst()
            If oWCTypeCombo.ValidValues.Count > 0 Then
                For i As Int16 = oWCTypeCombo.ValidValues.Count - 1 To 0 Step -1
                    oWCTypeCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            For i As Int16 = 0 To oRs.RecordCount - 1
                oWCTypeCombo.ValidValues.Add(oRs.Fields.Item(1).Value, oRs.Fields.Item(0).Value)
                oRs.MoveNext()
            Next
            oWCTypeCombo.ValidValues.Add("Define New", "")
        Catch ex As Exception
            Throw ex
        Finally
            ors = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' Loads the last entered Value
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WCTypeDFN()
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim StrSql As String
        Try
            If BoolWCType = False Then
                If Not oWCTypeCombo Is Nothing Then
                    WCTypeCombo()
                    StrSql = "select * from [@PSSIT_OTYP] where Docentry=(Select IsNull(Max(Docentry),0) as Code from [@PSSIT_OTYP])"
                    oRs.DoQuery(StrSql)
                    If oRs.RecordCount > 0 Then
                        oRs.MoveFirst()
                        oWCTypeCombo.Select(oRs.Fields.Item("U_Wctypnam").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                        BoolWCType = True
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
    ''' This is used to Load the Fixed Cost Type in the Combo
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FxdCostTypCombo()
        Dim oRs As SAPbobsCOM.Recordset
        Try
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery("select Code, U_Wcstyp from [@PSSIT_OCST]")


            If oFxdCostCol.ValidValues.Count > 0 Then
                For i As Int16 = oFxdCostCol.ValidValues.Count - 1 To 0 Step -1
                    oFxdCostCol.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            If oRs.RecordCount > 0 Then
                oRs.MoveFirst()
                For i As Int16 = 0 To oRs.RecordCount - 1
                    oFxdCostCol.ValidValues.Add(oRs.Fields.Item(1).Value, oRs.Fields.Item(0).Value)
                    oRs.MoveNext()
                Next
            End If
            oFxdCostCol.ValidValues.Add("Define New", "")
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    '''  Loads the last entered Value
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FxdCostTypDFN()
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim StrSql As String
        Dim oCostCombo As SAPbouiCOM.ComboBox
        Try
            If BoolCostTyp = False Then

                If Not oFxdCostCol Is Nothing Then
                    If oCostMatrix.RowCount > 0 Then
                        oCostCombo = oFxdCostCol.Cells.Item(oCostComboRow).Specific()
                        FxdCostTypCombo()
                        StrSql = "select U_Wcstyp from [@PSSIT_OCST] where Docentry=(Select IsNull(Max(Docentry),0) as Code from [@PSSIT_OCST])"
                        oRs.DoQuery(StrSql)
                        If oRs.RecordCount > 0 Then
                            oRs.MoveFirst()
                            Dim Val As String = oRs.Fields.Item("U_Wcstyp").Value()
                            oCostCombo.Select(Val, SAPbouiCOM.BoSearchKey.psk_ByDescription)
                            BoolCostTyp = True
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
    ''' This is used to delete the empty rows in the Fixed Cost Matrix
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FxdCostDeleteEmptyRow()
        Dim IntICount As Integer
        Dim oCostType As String
        Try
            For IntICount = oCostMatrix.RowCount To 1 Step -1
                oCostType = oFCostDB.GetValue("U_Fcost", oFCostDB.Offset).Trim()
                If oCostType.Length = 0 Then
                    oCostMatrix.DeleteRow(IntICount)
                    oCostMatrix.FlushToDataSource()
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    '''  This is used to Load the Absorption Method in the Combo
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AbsMethodCombo()
        Try
            If oAbsMethodCol.ValidValues.Count > 0 Then
                For i As Int16 = oAbsMethodCol.ValidValues.Count - 1 To 0 Step -1
                    oAbsMethodCol.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            oAbsMethodCol.ValidValues.Add("Finished Goods", "1")
            oAbsMethodCol.ValidValues.Add("Machine", "2")
            oAbsMethodCol.ValidValues.Add("Labour", "3")
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    ''' <summary>
    ''' This is used to Set the values in the Fixed Cost matrix while Adding the empty row
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FxdCostSetValue()
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs.DoQuery("select a.CurrCode from OCRN a,OADM b where a.CurrCode=b.MainCurncy")
        Try
            'oFCostDB.SetValue("U_Fcost", oFCostDB.Offset, "")
            oFCostDB.SetValue("U_Currency", oFCostDB.Offset, oRs.Fields.Item("CurrCode").Value)
            'oFCostDB.SetValue("U_UnitCost", oFCostDB.Offset, "")
            'oFCostDB.SetValue("U_Absmthd", oFCostDB.Offset, "")
            oFCostDB.SetValue("U_Accode", oFCostDB.Offset, "")
            oFCostDB.SetValue("U_Acname", oFCostDB.Offset, "")
            oFCostDB.SetValue("U_ActAcCode", oFCostDB.Offset, "")
            oFCostDB.SetValue("U_Adnl1", oFCostDB.Offset, "")
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' This is used to Load the Fixed Cost matrix Currency in the Combo
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FxtCstCurrCombo()
        Dim oRs, oRs1 As SAPbobsCOM.Recordset
        Try
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery("select CurrCode from OCRN where CurrCode is not null")
            oRs1.DoQuery("select a.CurrCode from OCRN a,OADM b where a.CurrCode=b.MainCurncy")
            oRs.MoveFirst()
            If oCurrencyCol.ValidValues.Count > 0 Then
                For i As Int16 = oCurrencyCol.ValidValues.Count - 1 To 0 Step -1
                    oCurrencyCol.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            For i As Int16 = 0 To oRs.RecordCount - 1
                oCurrencyCol.ValidValues.Add(oRs.Fields.Item(0).Value, "")
                oRs.MoveNext()
            Next
            'oCurrencyCol.Select(oRs1.Fields.Item(0).Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            oRs1 = Nothing
            GC.Collect()
        End Try
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
            If pVal.FormUID = "FWC" Then
                '*****************************Releasing the Com Object*******************************
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE) Then
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
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True Then
                        Try
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                FxdCostDeleteEmptyRow()
                                Validation()
                            End If
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End Try
                    End If

                    '********** Add Button Press ***********
                    If pVal.ItemUID = "1" And (pVal.BeforeAction = False) Then
                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            oForm.Items.Item("txtwcode").Enabled = False
                            oForm.Items.Item("txtwname").Enabled = True
                        End If
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Try
                                DeleteEmptyRows()
                                oForm.Refresh()
                                oForm.Freeze(True)
                                oActiveCheck.Checked = True
                                oCurrencyCol.DataBind.SetBound(True, "@PSSIT_WCR1", "U_Currency")
                                '*******Empty Row in Fixed Cost Matrix ********
                                oFCostDB.InsertRecord(oFCostDB.Size)
                                oFCostDB.Offset = oFCostDB.Size - 1
                                FxdCostSetValue()
                                oCostMatrix.AddRow(1, oCostMatrix.RowCount)
                                FxtCstCurrCombo()
                                oForm.Freeze(False)
                                oWCCodeTxt.Active = True
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End Try
                        End If
                    End If
                    '********** Update Button Press ***********
                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And (pVal.BeforeAction = True) Then
                        Try
                            DeleteEmptyRows()
                            FxdCostDeleteEmptyRow()
                            oForm.Update()
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        End Try
                    End If

                End If
                '***** Reloads the Combo's if Define New is selected and data added in the Forms *****
                If (pVal.FormTypeEx = "FWC") And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE) Then
                    WCTypeDFN()
                    FxdCostTypDFN()
                End If

                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT) Then
                    '******** Work Center Type Combo Select *********
                    If (pVal.ItemUID = "cmbwtype") And (pVal.BeforeAction = False) Then
                        '**** Work Center Type Combo Define New Selection *****
                        If oWCTypeCombo.Selected.Value = "Define New" Then
                            LoadDefaultForm("PSSIT_WCT")
                            BubbleEvent = False
                            oParentDB.SetValue("U_WCtype", oParentDB.Offset, "")
                            BoolWCType = False
                        End If
                    End If

                    '******* Add Next Row If the Cost Type is Selected **********
                    If (pVal.ItemUID = "matcost") And (pVal.ColUID = "colfcost") And (pVal.BeforeAction = False) And pVal.Row > 0 Then
                        Dim oCstCombo As SAPbouiCOM.ComboBox
                        Dim oCstTyp As String
                        Dim CurrentRow As Integer

                        CurrentRow = pVal.Row
                        oCstCombo = oFxdCostCol.Cells.Item(pVal.Row).Specific
                        oCstTyp = oCstCombo.Selected.Value
                        '****  Cost Type Combo Define New Selection *****
                        If oCstCombo.Selected.Value = "Define New" Then
                            oCostMatrix.GetLineData(CurrentRow)
                            oFCostDB.SetValue("U_Fcost", oFCostDB.Offset, "")
                            oCostMatrix.SetLineData(CurrentRow)
                            oCstTyp = ""
                            LoadDefaultForm("PSSIT_CST")
                            BubbleEvent = False
                            BoolCostTyp = False
                            oCostComboRow = pVal.Row
                        End If


                        If CurrentRow = oCostMatrix.VisualRowCount Then
                            oFCostDB.Offset = oFCostDB.Size - 1
                            FxdCostSetValue()
                            oCostMatrix.SetLineData(CurrentRow)
                            oCostMatrix.FlushToDataSource()
                        End If
                        oFCostDB.SetValue("U_Fcost", oFCostDB.Offset, oCstTyp)
                        oCostMatrix.SetLineData(CurrentRow)
                        oCostMatrix.FlushToDataSource()
                        If oCstTyp <> "Define New" And Len(oCstTyp) > 0 Then
                            oFCostDB.InsertRecord(oFCostDB.Size)
                            oFCostDB.Offset = oFCostDB.Size - 1
                            FxdCostSetValue()
                            If pVal.Row = oCostMatrix.RowCount Then
                                oCostMatrix.AddRow(1, oCostMatrix.RowCount)
                            End If

                            'End If
                        End If

                    End If
                End If
                '************** Add Row in Matrix *************
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK) And pVal.BeforeAction = False Then
                    If (pVal.ItemUID = "matcost") And pVal.ColUID = "#" Then
                        FCostUID = pVal.ItemUID
                    End If

                End If

            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="pVal"></param>
    ''' <param name="BubbleEvent"></param>
    ''' <remarks></remarks>
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Dim FormID As String
        Try
            FormID = SBO_Application.Forms.ActiveForm.UniqueID
            If pVal.MenuUID = "1281" And FormID = "FWC" Then
                If pVal.BeforeAction = True Then
                    oForm.Items.Item("txtwcode").Enabled = True
                    oForm.Items.Item("txtwname").Enabled = True
                End If
                If pVal.BeforeAction = False Then
                    oWCCodeTxt.Active = True
                End If
            End If
            If pVal.MenuUID = "1282" And pVal.BeforeAction = False And FormID = "FWC" Then
                oForm.Freeze(True)
                oParentDB.Offset = oParentDB.Size - 1
                oActiveCheck.Checked = True
                oForm.Items.Item("txtwcode").Enabled = True
                oForm.Items.Item("txtwname").Enabled = True
                oCurrencyCol.DataBind.SetBound(True, "@PSSIT_WCR1", "U_Currency")
                '*******Empty Row in Fixed Cost Matrix ********
                oFCostDB.InsertRecord(oFCostDB.Size)
                oFCostDB.Offset = oFCostDB.Size - 1
                FxdCostSetValue()
                oCostMatrix.AddRow(1, oCostMatrix.RowCount)
                FxtCstCurrCombo()
                oForm.Freeze(False)
                oWCCodeTxt.Active = True
            End If
            If pVal.MenuUID = "1283" And FormID = "FWC" Then
                If pVal.BeforeAction = True Then
                    Dim oStrSql As String
                    Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Try
                        oStrSql = "Select (sum(a.cnt) + Sum (b.cnt) + Sum (c.cnt) + Sum (d.cnt)) as ReferredCount " _
                        & "from (select count(*) as cnt from [@PSSIT_OLGP]  Where U_WCCode = '" & oWCCodeTxt.Value & "' ) as a, " _
                        & "(Select count(*) as cnt from [@PSSIT_PMWCHDR]  Where U_deptcode = '" & oWCCodeTxt.Value & "') as b, " _
                        & "(Select count(*) as cnt from [@PSSIT_OTLS]  Where U_WCCode = '" & oWCCodeTxt.Value & "') as c, " _
                        & "(Select count(*) as cnt from [@PSSIT_OMGP]  Where U_WCCode = '" & oWCCodeTxt.Value & "') as d "
                        oRs.DoQuery(oStrSql)
                        If oRs.RecordCount > 0 Then
                            oRs.MoveFirst()
                            If oRs.Fields.Item("ReferredCount").Value > 0 Then
                                SBO_Application.SetStatusBarMessage("Cannot be removed. Transactions are linked to an object, '" & oWCCodeTxt.Value & "'", SAPbouiCOM.BoMessageTime.bmt_Short, True)
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
            If (pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") And pVal.BeforeAction = False And FormID = "FWC" Then
                Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Try
                    oRs.DoQuery("Select * from [@PSSIT_OWCR]")
                    If oRs.RecordCount > 0 Then
                        oForm.Items.Item("txtwcode").Enabled = False
                        oForm.Items.Item("txtwname").Enabled = True
                    Else
                        oForm.Items.Item("txtwcode").Enabled = True
                        oForm.Items.Item("txtwname").Enabled = True
                    End If
                Catch ex As Exception
                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Finally
                    oRs = Nothing
                    GC.Collect()
                End Try
            End If
            '*************** Delete Row in Matrix ******************
            If pVal.MenuUID = "1293" And pVal.BeforeAction = True And FCostUID = "matcost" And FormID = "FWC" Then
                Try
                    oFCostDB = oForm.DataSources.DBDataSources.Item("@PSSIT_WCR1")
                    oCostMatrix.FlushToDataSource()
                    oCostMatrix.DeleteRow(oCostMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                    If (oCostMatrix.RowCount = 0) Then
                        oFCostDB.RemoveRecord(0)
                    End If
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    BubbleEvent = False
                Catch ex As Exception
                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    BubbleEvent = False
                End Try
            End If
            '*************** Add Row in Matrix ******************
            If pVal.MenuUID = "1292" And pVal.BeforeAction = True And FCostUID = "matcost" Then
                Try
                    'oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                    If oCostMatrix.Columns.Item("colfcost").Cells.Item(oCostMatrix.RowCount).Specific.value <> "" Then
                        oFCostDB.InsertRecord(oFCostDB.Size)
                        oFCostDB.Offset = oFCostDB.Size - 1
                        FxdCostSetValue()
                        oCostMatrix.AddRow(1, oCostMatrix.RowCount)
                        oForm.Update()
                    End If
                    
                Catch ex As Exception
                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    BubbleEvent = False
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
        Dim IntICount As Integer
        Dim oCostType, oAbsMthd As String
        Try
            If oWCCodeTxt.Value.Length = 0 Then
                oWCCodeTxt.Active = True
                Throw New Exception("Work Centre Code should not be Empty")
            End If
            If oWCNameTxt.Value.Length = 0 Then
                oWCNameTxt.Active = True
                Throw New Exception("Work Centre Name should not be Empty")
            End If
            'If oParentDB.GetValue("U_WCtype", oParentDB.Offset).Length = 0 Then
            '    oWCTypeCombo.Active = True
            '    Throw New Exception("Work Centre Type should not be empty")
            'End If
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                oRs.DoQuery("select Code from [@PSSIT_OWCR]  where Code= '" & oWCCodeTxt.Value & "' ")
                If oRs.RecordCount > 0 Then
                    oWCCodeTxt.Active = True
                    Throw New Exception("Work Centre Code Already Exist")
                End If
                oRs1.DoQuery("select U_WCname from [@PSSIT_OWCR]  where U_WCname= '" & oWCNameTxt.Value & "' ")
                If oRs1.RecordCount > 0 Then
                    oWCNameTxt.Active = True
                    Throw New Exception("Work Centre Name Already Exist")
                End If

            End If
            FixedCostKeyCheck()
            AccKeyCheck()
            If oCostMatrix.RowCount > 0 Then
                For IntICount = 1 To oCostMatrix.VisualRowCount
                    oCostMatrix.GetLineData(IntICount)
                    oCostType = oFCostDB.GetValue("U_Fcost", IntICount - 1).Trim()
                    oAbsMthd = oFCostDB.GetValue("U_Absmthd", IntICount - 1).Trim()
                    If oCostType.Length > 0 Then
                        If oUnitCostCol.Cells.Item(IntICount).Specific.Value < 0 Then
                            oUnitCostCol.Cells.Item(IntICount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Throw New Exception("Unit Cost should be greater than zero")
                        End If
                        'Modified by Manimaran-------s
                        'If oAbsMthd.Length = 0 Then
                        '    Throw New Exception("Absorption method should not be empty")
                        'End If
                        If oAcctCodeCol.Cells.Item(IntICount).Specific.Value.Length = 0 Or oCostMatrix.Columns.Item("colaccod").Cells.Item(IntICount).Specific.value.length = 0 Then
                            oAcctCodeCol.Cells.Item(IntICount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Throw New Exception("Account Code should not be empty")
                        End If
                        'Modified by Manimaran-------e
                    End If
                Next
            End If
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
        Dim IntICount As Integer
        Try
            oRs.DoQuery("Select * from [@PSSIT_OCON] where U_AccKey = 'Y'")
            If oRs.RecordCount > 0 Then
                For IntICount = 1 To oCostMatrix.RowCount
                    oCostMatrix.GetLineData(IntICount)
                    If oFCostDB.GetValue("U_Accode", oFCostDB.Offset).Trim().Length = 0 Then
                        oAcctCodeCol.Cells.Item(IntICount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Throw New Exception("Account Details Mandatory")
                    End If
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
    ''' Checking the Account Key based on the production configuration form.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FixedCostKeyCheck()
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim IntICount As Integer
        Dim oFxdCostCombo, oCurrcombo As SAPbouiCOM.ComboBox
        Dim oUnitCostEdit As SAPbouiCOM.EditText
        Try
            oRs.DoQuery("Select * from [@PSSIT_OCON] where U_Fcman = 'Y'")
            'If oRs.RecordCount > 0 Then
            '    If oCostMatrix.RowCount = 0 Then
            '        Throw New Exception("Fixed Cost is Mandatory")
            '    ElseIf oCostMatrix.RowCount > 0 Then
            '        For IntICount = 1 To oCostMatrix.RowCount
            '            oFxdCostCombo = oFxdCostCol.Cells.Item(IntICount).Specific
            '            oCurrcombo = oCurrencyCol.Cells.Item(IntICount).Specific
            '            oUnitCostEdit = oUnitCostCol.Cells.Item(IntICount).Specific
            '            oCostMatrix.GetLineData(IntICount)
            '            If Not oFxdCostCombo Is Nothing Then
            '                'If oFxdCostCombo.Selected.Description.Length = 0 Then
            '                If oFCostDB.GetValue("U_Fcost", oFCostDB.Offset).Trim().Length = 0 Then
            '                    oFxdCostCol.Cells.Item(IntICount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            '                    Throw New Exception("Fixed Cost is Mandatory")
            '                End If
            '            End If
            '            If oCurrcombo.Selected.Value.Length = 0 Then
            '                Throw New Exception("Currency is Mandatory")
            '            End If
            '            If oUnitCostEdit.Value = 0 Then
            '                oUnitCostCol.Cells.Item(IntICount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            '                Throw New Exception("Unit Cost should be greater than zero")
            '            End If
            '        Next
            '    End If
            'End If
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' Delete the empty rows added in the database
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub DeleteEmptyRows()
        Dim oRS As SAPbobsCOM.Recordset
        Try
            oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("Delete from [@PSSIT_WCR1] where U_Fcost is null")
        Catch ex As Exception
            Throw ex
        Finally
            oRS = Nothing
            GC.Collect()
        End Try
      
    End Sub
    ''' <summary>
    ''' Load the data in the matrix 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadMatrixData()
        Dim oConditions As SAPbouiCOM.Conditions
        Dim oCondition As SAPbouiCOM.Condition
        Try

            oConditions = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions)
            oCondition = oConditions.Add

            oCondition.BracketOpenNum = 1
            oCondition.Alias = "Code"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = oWCCode
            oCondition.BracketCloseNum = 1

            oFCostDB.Query(oConditions)
            oCostMatrix.LoadFromDataSource()
            oCostMatrix.FlushToDataSource()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
