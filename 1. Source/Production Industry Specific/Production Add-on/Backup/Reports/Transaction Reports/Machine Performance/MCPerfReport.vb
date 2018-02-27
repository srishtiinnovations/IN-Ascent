'''' <summary>
'''' Author                     Created Date
'''' Suresh                      21/01/2009
'''' <remarks> This class is used for entering the Parameters for the Machine Performance Report.</remarks>
Public Class MCPerfReport
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
    Private oChMCFList, oChMCFBtnList, oChMCTList, oChMCTBtnList As SAPbouiCOM.ChooseFromList
    Private Event ChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable)
    '**************************UserDataSource************************************
    Private UFPordDate, UTPordDate, UFMCNo, UTMCNo As SAPbouiCOM.UserDataSource
    '**************************Items - EditText************************************
    Private oFromDateTxt, oToDateTxt, oFromMachineTxt, oToMachineTxt As SAPbouiCOM.EditText
    '**************************Items - Button************************************
    Private BtnPrint, BtnFMc, BtnTMc As SAPbouiCOM.Button
    Private oMachNo As String
    Private sPath As String = IO.Directory.GetParent(Application.ExecutablePath).ToString
    Private WithEvents MCPerfChildReportClass As MCPerfChildReport
#End Region
    ''' <summary>
    ''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmMCPerfReport.srf") method is called to load the Machine Performance Report form.
    ''' Drawform() method is called to Initialize the form,Datasources and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company)
        MyBase.New(aSBO_Application, aCompany)
        SBO_Application = aSBO_Application
        oCompany = aCompany
        LoadFromXML("FrmMCPerfReport.srf")
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
            LoadLookups()
            InitializeFormComponent()
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
            UFPordDate = oForm.DataSources.UserDataSources.Add("UFPordDate", SAPbouiCOM.BoDataType.dt_DATE, 10)
            UTPordDate = oForm.DataSources.UserDataSources.Add("UTPordDate", SAPbouiCOM.BoDataType.dt_DATE, 10)
            UFMCNo = oForm.DataSources.UserDataSources.Add("UFMCNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
            UTMCNo = oForm.DataSources.UserDataSources.Add("UTMCNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
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
            oFromDateTxt = oForm.Items.Item("txtfdate").Specific
            oFromDateTxt.DataBind.SetBound(True, "", "UFPordDate")

            oToDateTxt = oForm.Items.Item("txttdate").Specific
            oToDateTxt.DataBind.SetBound(True, "", "UTPordDate")

            oFromMachineTxt = oForm.Items.Item("txtfmc").Specific
            oFromMachineTxt.DataBind.SetBound(True, "", "UFMCNo")
            oFromMachineTxt.ChooseFromListUID = "MCFLst"
            oFromMachineTxt.ChooseFromListAlias = "U_wcno"

            oToMachineTxt = oForm.Items.Item("txttmc").Specific
            oToMachineTxt.DataBind.SetBound(True, "", "UTMCNo")
            oToMachineTxt.ChooseFromListUID = "MCTLst"
            oToMachineTxt.ChooseFromListAlias = "U_wcno"


            BtnFMc = oForm.Items.Item("btnfmc").Specific
            oForm.Items.Item("btnfmc").Description = "Choose from List" & vbNewLine & "Machine Utilization List View"
            BtnFMc.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            BtnFMc.Image = sPath & "\Resources\CFL.bmp"
            BtnFMc = oForm.Items.Item("btnfmc").Specific
            BtnFMc.ChooseFromListUID = "BtMCFLst"


            BtnTMc = oForm.Items.Item("btntmc").Specific
            oForm.Items.Item("btntmc").Description = "Choose from List" & vbNewLine & "Machine Utilization List View"
            BtnTMc.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            BtnTMc.Image = sPath & "\Resources\CFL.bmp"
            BtnTMc = oForm.Items.Item("btntmc").Specific
            BtnTMc.ChooseFromListUID = "BtMCTLst"

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
            oChMCFList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_WCHDR", "MCFLst"))
            oChMCFBtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_WCHDR", "BtMCFLst"))

            oChMCTList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_WCHDR", "MCTLst"))

            oChMCTBtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_WCHDR", "BtMCTLst"))

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
    Private Sub MachineUtilization_ChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable) Handles Me.ChooseFromList
        Dim oDataTable As SAPbouiCOM.DataTable = ChooseFromListSelectedObjects
        Dim oCurrentRow As Integer
        Dim oFMCNo, oTMCNo As String
        Dim oRs As SAPbobsCOM.Recordset
        Try
            oCurrentRow = CurrentRow
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '**********************From MC No CFL**************************
            If (ControlName = "txtfmc" Or ControlName = "btnfmc") And (ChoosefromListUID = "MCFLst" Or ChoosefromListUID = "BtMCFLst") Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    If Not oDataTable Is Nothing Then
                        oFMCNo = oDataTable.GetValue("U_wcno", 0)
                        UFMCNo.Value = oFMCNo
                    End If
                End If
            End If
            '**********************To MC No CFL**************************
            If (ControlName = "txttmc" Or ControlName = "btntmc") And (ChoosefromListUID = "MCTLst" Or ChoosefromListUID = "BtMCTLst") Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    If Not oDataTable Is Nothing Then
                        oTMCNo = oDataTable.GetValue("U_wcno", 0)
                        UTMCNo.Value = oTMCNo
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
        Dim ChooseFromListEvent As SAPbouiCOM.ChooseFromListEvent
        Try
            If pVal.FormUID = "FMPR" Then
                '*****************************Releasing the Com Object*******************************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                    SBO_Application = Nothing
                    GC.Collect()
                End If
                '**********ChooseFromList Event is called using the raiseevent******************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST And pVal.BeforeAction = False Then
                    ChooseFromListEvent = pVal
                    RaiseEvent ChooseFromList(pVal.ItemUID, pVal.ColUID, pVal.Row, ChooseFromListEvent.ChooseFromListUID, ChooseFromListEvent.SelectedObjects)
                End If
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) Then
                    '*************Shift Button Press*****************
                    If (pVal.ItemUID = "btnprint") Then
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
                            Dim StrSql, StrSql1 As String
                            Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            Try
                                'Modified by Manimaran----s
                                '                            StrSql1 = "select T0.U_wcno from [@PSSIT_PEY1] T0 " _
                                '& "inner join [@PSSIT_OPEY] T1 on T1.Docentry=T0.Docentry " _
                                '& "inner join [@PSSIT_PMWCHDR] T2 on T2.U_wcno=T0.U_wcno " _
                                '& "where T1.U_Docdt >= '" & oFromDateTxt.Value & "' " _
                                '& "and T1.U_Docdt <= '" & oToDateTxt.Value & "' " _
                                '& "and T2.U_wcno >= '" & oFromMachineTxt.Value & "' and T2.U_wcno <= '" & oToMachineTxt.Value & "' " _
                                '& "Group by T0.U_wcno"
                                StrSql1 = "select T0.U_wcno from [@PSSIT_PEY1] T0 " _
& "inner join [@PSSIT_OPEY] T1 on T1.Docentry=T0.Docentry " _
& "inner join [@PSSIT_PMWCHDR] T2 on T2.U_wcno=T0.U_wcno " _
& "where T1.U_Docdt >= '" & oFromDateTxt.Value & "' " _
& "and T1.U_Docdt <= '" & oToDateTxt.Value & "' " _
& "and T2.U_wcno in ( '" & oFromMachineTxt.Value & "' ,'" & oToMachineTxt.Value & "') " _
& "Group by T0.U_wcno"
                                'Modified by Manimaran----e
                                oRs1.DoQuery(StrSql1)
                                If oRs1.RecordCount > 0 Then
                                    oRs1.MoveFirst()
                                    Dim i As Integer

                                    For i = 0 To oRs1.RecordCount - 1
                                        oMachNo = oRs1.Fields.Item("U_wcno").Value
                                        'Modified by Manimaran----s
                                        '                                    StrSql = "select Tbl2.Machine,Tbl2.Item,Tbl2.PlanTime,Tbl2.PlanQty,Tbl2.WrkTime,Tbl2.OutQty , " _
                                        '& "Case When Tbl2.Perf <= 100 then Tbl2.Perf Else '100' End from " _
                                        '& "(select T0.U_wcno as 'Machine',Tbl.Item as 'Item',Tbl.PlanTime as 'PlanTime', " _
                                        '& "Tbl.PlanQty as 'PlanQty',Tbl1.WrkTime as 'WrkTime',Tbl1.OutQty as 'OutQty' " _
                                        '& ",Convert(float,(convert(float,Tbl1.WrkTime) / convert(float,Tbl.PlanTime)) * 100) as 'Perf' " _
                                        '& "from [@PSSIT_PMWCHDR] T0 " _
                                        '& "inner join (select T3.U_wcno,T1.ItemCode as 'Item',T1.PlannedQty as 'PlanQty' , " _
                                        '& "((T1.PlannedQty * (Sum(T3.U_Opertime)/sum(T3.U_perqty)))/60) as 'PlanTime' from OWOR T1 " _
                                        '& "inner join [@PSSIT_WOR2] T2 on T2.U_Pordno=T1.Docnum " _
                                        '& "inner join [@PSSIT_RTE1] T3 on T3.U_Rteid=T2.U_Rteid " _
                                        '& "where T3.U_wcno='" & oMachNo & "' and T3.U_Opertime is not null and T3.U_perqty is not null " _
                                        '& "group by T3.U_wcno,T1.ItemCode,T1.PlannedQty ) Tbl on Tbl.U_wcno=T0.U_wcno " _
                                        '& "inner join (select T4.U_wcno,(Sum(T4.U_Rntime)/60) as 'WrkTime', " _
                                        '& "Sum(T4.U_Qty) as 'OutQty' from [@PSSIT_PEY1] T4 " _
                                        '& "inner join [@PSSIT_OPEY] T5 on T5.Docentry=T4.Docentry " _
                                        '& "inner join OWOR T6 on T6.Docnum=T5.U_WORNo " _
                                        '& "where T4.U_wcno='" & oMachNo & "' and T5.U_Rework='N' " _
                                        '& "Group by T4.U_wcno) Tbl1 on Tbl.U_wcno=T0.U_wcno " _
                                        '& "where T0.U_wcno='" & oMachNo & "'  " _
                                        '& "group by T0.U_wcno,Tbl.Item,Tbl.PlanTime,Tbl.PlanQty,Tbl1.WrkTime,Tbl1.OutQty) Tbl2"
                                        StrSql = "select Tbl2.Machine,Tbl2.Item,Tbl2.PlanTime,Tbl2.PlanQty,Tbl2.WrkTime,Tbl2.OutQty , " _
                                           & "Case When Tbl2.Perf <= 100 then Tbl2.Perf Else '100' End from " _
                                           & "(select T0.U_wcno as 'Machine',Tbl.Item as 'Item',Tbl.PlanTime as 'PlanTime', " _
                                           & "Tbl.PlanQty as 'PlanQty',Tbl1.WrkTime as 'WrkTime',Tbl1.OutQty as 'OutQty' " _
                                           & ",convert(numeric,Convert(float,(convert(float,Tbl1.WrkTime) / convert(float,Tbl.PlanTime)) * 100),3) as 'Perf' " _
                                           & "from [@PSSIT_PMWCHDR] T0 " _
                                           & "inner join (select T3.U_wcno,T1.ItemCode as 'Item',T1.PlannedQty as 'PlanQty' , " _
                                           & "((T1.PlannedQty * (Sum(T3.U_Opertime)/sum(T3.U_perqty)))) as 'PlanTime' from OWOR T1 " _
                                           & "inner join [@PSSIT_WOR2] T2 on T2.U_Pordno=T1.Docnum " _
                                           & "inner join [@PSSIT_RTE1] T3 on T3.U_Rteid=T2.U_Rteid " _
                                           & "where T3.U_wcno='" & oMachNo & "' and T3.U_Opertime is not null and T3.U_perqty is not null " _
                                           & "group by T3.U_wcno,T1.ItemCode,T1.PlannedQty ) Tbl on Tbl.U_wcno=T0.U_wcno " _
                                           & "inner join (select T4.U_wcno,(Sum(T4.U_Rntime)) as 'WrkTime', " _
                                           & "Sum(T4.U_Qty) as 'OutQty' from [@PSSIT_PEY1] T4 " _
                                           & "inner join [@PSSIT_OPEY] T5 on T5.Docentry=T4.Docentry " _
                                           & "inner join OWOR T6 on T6.Docnum=T5.U_WORNo " _
                                           & "where T4.U_wcno='" & oMachNo & "' and T5.U_Rework='N' " _
                                           & "Group by T4.U_wcno) Tbl1 on Tbl.U_wcno=T0.U_wcno " _
                                           & "where T0.U_wcno='" & oMachNo & "'  " _
                                           & "group by T0.U_wcno,Tbl.Item,Tbl.PlanTime,Tbl.PlanQty,Tbl1.WrkTime,Tbl1.OutQty) Tbl2"
                                        'Modified by Manimaran--------e
                                        MCPerfChildReportClass = New MCPerfChildReport(SBO_Application, oCompany, StrSql)
                                        oRs1.MoveNext()
                                        If oRs1.EoF = True Then
                                            Exit For
                                        End If
                                    Next i
                                Else
                                    SBO_Application.StatusBar.SetText("No Records Found", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                End If

                            Catch ex As Exception
                                If ex.Message.Contains("Form - already exists") Then
                                Else
                                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                End If

                            End Try
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
            If oFromDateTxt.Value.Length = 0 Then
                oFromDateTxt.Active = True
                Throw New Exception("From Date should not be Empty")
            End If
            If oToDateTxt.Value.Length = 0 Then
                oToDateTxt.Active = True
                Throw New Exception("To Date should not be Empty")
            End If
            If oFromMachineTxt.Value.Length = 0 Then
                oFromMachineTxt.Active = True
                Throw New Exception("From Machine should not be Empty")
            End If
            If oToMachineTxt.Value.Length = 0 Then
                oToMachineTxt.Active = True
                Throw New Exception("To Machine should not be Empty")
            End If
        Catch ex As Exception
            Throw ex
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
