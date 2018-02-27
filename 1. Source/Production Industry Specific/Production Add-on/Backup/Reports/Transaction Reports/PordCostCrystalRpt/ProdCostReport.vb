'''' <summary>
'''' Author                     Created Date
'''' Suresh                      08/01/2009
'''' <remarks> This class is used for entering the Parameters for the Production Order Cost Report.</remarks>
Public Class ProdCostReport
    Inherits GeneralLib
    ''' <summary>
    ''' Variable Declaration
    ''' </summary>
    ''' <remarks></remarks>
#Region "Variable Declaration"
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private SboGuiApi As New SAPbouiCOM.SboGuiApi
    Private oCompany As SAPbobsCOM.Company
    '**************************Form************************************
    Private oForm As SAPbouiCOM.Form
    '**************************ChooseFromList************************************
    Private oChPOFList, oChPOFBtnList, oChPOTList, oChPOTBtnList As SAPbouiCOM.ChooseFromList
    Private Event ChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable)
    '**************************UserDataSource************************************
    Private UFPordNo, UTPordNo As SAPbouiCOM.UserDataSource
    '**************************Items - EditText************************************
    Private oPordSerNoTxt, oFromPordNoTxt, oToPordNoTxt, oPordStatusTxt As SAPbouiCOM.EditText
    '**************************Items - Combo************************************
    Private oPordSeriesCombo, oPordStatusCombo As SAPbouiCOM.ComboBox
    '**************************Items - Button************************************
    Private BtnPrint, BtnFrom, BtnTo As SAPbouiCOM.Button
    Private sPath As String = IO.Directory.GetParent(Application.ExecutablePath).ToString
    Private TestForm As FrmTest
    Private oThread As System.Threading.Thread
#End Region
    ''' <summary>
    ''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmProdCostReport.srf") method is called to load the Skill Group form.
    ''' Drawform() method is called to Initialize the form,Datasources and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company)
        MyBase.New(aSBO_Application, aCompany)
        SBO_Application = aSBO_Application
        oCompany = aCompany
        LoadFromXML("FrmProdCostReport.srf")
        DrawForm()
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
            AddUserDataSources()
            InitializeFormComponent()
            LoadLookups()
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            oForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Initializing UserDataSources By setting UniqueID,DataType of the field and Length of the Field.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AddUserDataSources()
        Try
            UFPordNo = oForm.DataSources.UserDataSources.Add("UFPordNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            UTPordNo = oForm.DataSources.UserDataSources.Add("UTPordNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
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
            oPordSeriesCombo = oForm.Items.Item("cmbprodno").Specific
            ProdSeriesCombo()

            oPordSerNoTxt = oForm.Items.Item("txtprdsno").Specific

            oFromPordNoTxt = oForm.Items.Item("txtfrom").Specific
            oFromPordNoTxt.DataBind.SetBound(True, "", "UFPordNo")
            ' oSGCodeTxt.DataBind.SetBound(True, "@PSSIT_OLGP", "Code")
           


            oToPordNoTxt = oForm.Items.Item("txtto").Specific
            oToPordNoTxt.DataBind.SetBound(True, "", "UTPordNo")
           

            '  oSGDescTxt.DataBind.SetBound(True, "@PSSIT_OLGP", "U_LGname")

            oPordStatusCombo = oForm.Items.Item("cmbsts").Specific
            ProdStatusCombo()

            oPordStatusTxt = oForm.Items.Item("txtsts").Specific
          

            BtnFrom = oForm.Items.Item("btnfrm").Specific
            oForm.Items.Item("btnfrm").Description = "Choose from List" & vbNewLine & "Production Order Details List View"
            BtnFrom.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            BtnFrom.Image = sPath & "\Resources\CFL.bmp"
            BtnFrom = oForm.Items.Item("btnfrm").Specific


            BtnTo = oForm.Items.Item("btnto").Specific
            oForm.Items.Item("btnto").Description = "Choose from List" & vbNewLine & "Production Order Details List View"
            BtnTo.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            BtnTo.Image = sPath & "\Resources\CFL.bmp"
            BtnTo = oForm.Items.Item("btnto").Specific


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
            oChPOFList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "202", "POFLst"))
            '  CreateNewConditions(oChPOFList, "Series", SAPbouiCOM.BoConditionOperation.co_EQUAL, oPordSerNoTxt.Value)
            oFromPordNoTxt.ChooseFromListUID = "POFLst"
            oFromPordNoTxt.ChooseFromListAlias = "DocNum"

            oChPOFBtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "202", "BtPOFLst"))
            ' CreateNewConditions(oChPOFBtnList, "Series", SAPbouiCOM.BoConditionOperation.co_EQUAL, oPordSerNoTxt.Value)
            BtnFrom.ChooseFromListUID = "BtPOFLst"

            oChPOTList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "202", "POTLst"))
            ' CreateNewConditions(oChPOTList, "Series", SAPbouiCOM.BoConditionOperation.co_EQUAL, oPordSerNoTxt.Value)
            oToPordNoTxt.ChooseFromListUID = "POTLst"
            oToPordNoTxt.ChooseFromListAlias = "DocNum"

            oChPOTBtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "202", "BtPOTLst"))
           ' CreateNewConditions(oChPOTBtnList, "Series", SAPbouiCOM.BoConditionOperation.co_EQUAL, oPordSerNoTxt.Value)
            BtnTo.ChooseFromListUID = "BtPOTLst"

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
        Dim oFPONo, oTPONo As String
        Dim oRs As SAPbobsCOM.Recordset
        Try
            oCurrentRow = CurrentRow
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '**********************From PO No CFL**************************
            If (ControlName = "txtfrom" Or ControlName = "btnfrm") And (ChoosefromListUID = "POFLst" Or ChoosefromListUID = "BtPOFLst") Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    If Not oDataTable Is Nothing Then
                        oFPONo = oDataTable.GetValue("DocNum", 0)
                        UFPordNo.Value = oFPONo
                    End If
                End If
            End If
            '**********************To PO No CFL**************************
            If (ControlName = "txtto" Or ControlName = "btnto") And (ChoosefromListUID = "POTLst" Or ChoosefromListUID = "BtPOTLst") Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    If Not oDataTable Is Nothing Then
                        oTPONo = oDataTable.GetValue("DocNum", 0)
                        UTPordNo.Value = oTPONo
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
    ''' This is used to Production Order Series in the Combo
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ProdSeriesCombo()
        Try
            Dim oRs As SAPbobsCOM.Recordset
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery("select SeriesName,Series from NNM1 where ObjectCode='202'")
            oRs.MoveFirst()
            If oPordSeriesCombo.ValidValues.Count > 0 Then
                For i As Int16 = oPordSeriesCombo.ValidValues.Count - 1 To 0 Step -1
                    oPordSeriesCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            For i As Int16 = 0 To oRs.RecordCount - 1
                oPordSeriesCombo.ValidValues.Add(oRs.Fields.Item(0).Value, oRs.Fields.Item(1).Value)
                oRs.MoveNext()
            Next
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    ''' <summary>
    ''' This is used to Production Order Status in the Combo
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ProdStatusCombo()
        Try
            Dim oRs As SAPbobsCOM.Recordset
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery("select distinct Status,case when Status='R' then 'Released' when Status ='C' then 'Closed' End as StatusDesc from OWOR where (Status='R' or Status='C') Group By Status")
            oRs.MoveFirst()
            If oPordStatusCombo.ValidValues.Count > 0 Then
                For i As Int16 = oPordSeriesCombo.ValidValues.Count - 1 To 0 Step -1
                    oPordStatusCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            For i As Int16 = 0 To oRs.RecordCount - 1
                oPordStatusCombo.ValidValues.Add(oRs.Fields.Item(1).Value, oRs.Fields.Item(0).Value)
                oRs.MoveNext()
            Next
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Dim ChooseFromListEvent As SAPbouiCOM.ChooseFromListEvent
        Try
            If pVal.FormUID = "FPOCR" Then
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
                '**************Combo Select**************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT Then
                    'If pVal.ItemUID = "cmbprodno" And pVal.BeforeAction = False Then
                    '    oPordSerNoTxt.Value = oPordSeriesCombo.Selected.Description
                    '    Try
                    '        If Len(oPordSerNoTxt.Value) > 0 Then
                    '            If (oChPOFList Is Nothing Or oChPOFBtnList Is Nothing Or oChPOTList Is Nothing Or oChPOTBtnList Is Nothing) Then
                    '                LoadLookups()
                    '            End If
                    '        End If
                    '    Catch ex As Exception
                    '        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    '    End Try
                    'End If
                    If pVal.ItemUID = "cmbsts" And pVal.BeforeAction = False Then
                        oPordStatusTxt.Value = oPordStatusCombo.Selected.Description
                    End If
                End If
                '**************Item Pressed**************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.ItemUID = "btnprint" Then
                        '***** Validation() method is called for validating the values in the edit text *****
                        If (pVal.BeforeAction = True) Then
                            Try
                                Validation()
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                            End Try
                        End If
                        If (pVal.BeforeAction = False) Then
                            TestForm = New FrmTest(SBO_Application, oCompany, "ProdCostCrystalRpt", oFromPordNoTxt.Value, oToPordNoTxt.Value, oPordStatusTxt.Value)
                            oThread = New Threading.Thread(AddressOf TestForm.StartThread)
                            oThread.Start()
                        End If
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
    Private Sub Validation()

        Try
            If oFromPordNoTxt.Value.Length = 0 Then
                oFromPordNoTxt.Active = True
                Throw New Exception("From Date should not be Empty")
            End If
            If oToPordNoTxt.Value.Length = 0 Then
                oToPordNoTxt.Active = True
                Throw New Exception("To Date should not be Empty")
            End If
            'If oPordStatusTxt.Value.Length = 0 Then
            '    Throw New Exception("Status should not be Empty")
            'End If
            
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
