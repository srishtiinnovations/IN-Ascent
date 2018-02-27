'''' <summary>
'''' Author                     Created Date
'''' Suresh                      18/12/2008
'''' <remarks> This class is used for entering the Labour Details.</remarks>
Public Class Labour
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
    Private oChEmpList, oChEmpBtnList, oChSGList, oChSGBtnList, oChLAccList, oChLABtnList As SAPbouiCOM.ChooseFromList
    Private Event ChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable)
    '**************************Items - EditText************************************
    Private oLBActAcCodeTxt, oLGCodeTxt, oLGDescTxt, oEmpNoTxt, oEmpNameTxt, oSGNoTxt, oSGNameTxt, oLbrRateTxt, oLBAcctCodeTxt, oLBAcctNameTxt, oInfo1Txt, oInfo2Txt, oRemarksTxt As SAPbouiCOM.EditText
    '**************************Items - Combo************************************
    Private oLbrCurrCombo As SAPbouiCOM.ComboBox
    '**************************Items - CheckBox************************************
    Private oActiveCheck As SAPbouiCOM.CheckBox
    '**************************Items - LinkedButton************************************
    Private oEmpCodeLink, oSkillGrpLink, oLBAcctCodeLink As SAPbouiCOM.LinkedButton
    '**************************Items - Button************************************
    Private BtnEmp, BtnSG, BtnLAcct As SAPbouiCOM.Button
    Private sPath As String = IO.Directory.GetParent(Application.ExecutablePath).ToString
    Private WithEvents SkillGroupClass As SkillGroups
    Private oLabourCode, oFormName As String
#End Region
    ''' <summary>
    ''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmLabour.srf") method is called to load the Labour form.
    ''' Drawform() method is called to Initialize the form,Datasources and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company, Optional ByVal aLabourCode As String = Nothing, Optional ByVal aFormName As String = Nothing)
        MyBase.New(aSBO_Application, aCompany)
        SBO_Application = aSBO_Application
        oLabourCode = aLabourCode
        oFormName = aFormName
        oCompany = aCompany
        LoadFromXML("FrmLabour.srf")
        DrawForm()
        If oFormName = "Production Entry" Or oFormName = "LbrPerfRpt" Then
            oParentDB.SetValue("Code", oParentDB.Offset, oLabourCode)
            oLGCodeTxt.Value = oLabourCode
            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        End If
        oForm.DataBrowser.BrowseBy = "txtlrcode"
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
            oParentDB = oForm.DataSources.DBDataSources.Item("@PSSIT_OLBR")
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
            oLGCodeTxt = oForm.Items.Item("txtlrcode").Specific
            oLGCodeTxt.DataBind.SetBound(True, "@PSSIT_OLBR", "Code")

            oEmpNoTxt = oForm.Items.Item("txtempid").Specific
            oEmpNoTxt.DataBind.SetBound(True, "@PSSIT_OLBR", "U_Empid")
            oEmpNoTxt.ChooseFromListUID = "EmpLst"
            oEmpNoTxt.ChooseFromListAlias = "empID"
            oForm.Items.Item("txtempid").LinkTo = "lnkempid"
            oForm.Items.Add("lnkempid", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnkempid").Visible = True
            oForm.Items.Item("lnkempid").LinkTo = "txtempid"
            oForm.Items.Item("lnkempid").Top = 21
            oForm.Items.Item("lnkempid").Left = 99
            oForm.Items.Item("lnkempid").Description = "Link to" & vbNewLine & "Employee Master Data"
            oEmpCodeLink = oForm.Items.Item("lnkempid").Specific
            oEmpCodeLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Employee

            oEmpNameTxt = oForm.Items.Item("txtempnam").Specific
            oEmpNameTxt.DataBind.SetBound(True, "@PSSIT_OLBR", "U_Empnam")

            oSGNoTxt = oForm.Items.Item("txtgcode").Specific
            oSGNoTxt.DataBind.SetBound(True, "@PSSIT_OLBR", "U_LGCode")
            oSGNoTxt.ChooseFromListUID = "SGLst"
            oSGNoTxt.ChooseFromListAlias = "Code"
            oForm.Items.Item("txtgcode").LinkTo = "lnkskill"
            oForm.Items.Add("lnkskill", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnkskill").Visible = True
            oForm.Items.Item("lnkskill").LinkTo = "txtgcode"
            oForm.Items.Item("lnkskill").Top = 36
            oForm.Items.Item("lnkskill").Left = 99
            oForm.Items.Item("lnkskill").Description = "Link to" & vbNewLine & "Skill Group"
            oSkillGrpLink = oForm.Items.Item("lnkskill").Specific

            oSGNameTxt = oForm.Items.Item("txtgname").Specific
            oSGNameTxt.DataBind.SetBound(True, "@PSSIT_OLBR", "U_LGname")

            oLbrRateTxt = oForm.Items.Item("txtlabrat").Specific
            oLbrRateTxt.DataBind.SetBound(True, "@PSSIT_OLBR", "U_Labrate")

            oLbrCurrCombo = oForm.Items.Item("cmbcurncy").Specific
            oLbrCurrCombo.DataBind.SetBound(True, "@PSSIT_OLBR", "U_Currncy")
            oForm.Items.Item("cmbcurncy").Enabled = False


            oLBAcctCodeTxt = oForm.Items.Item("txtaccod").Specific
            oLBAcctCodeTxt.DataBind.SetBound(True, "@PSSIT_OLBR", "U_Accode")
            oLBAcctCodeTxt.ChooseFromListUID = "LAccLst"
            oLBAcctCodeTxt.ChooseFromListAlias = "AcctCode"
            oForm.Items.Item("txtaccod").LinkTo = "lnkaccod"
            oForm.Items.Add("lnkaccod", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnkaccod").Visible = True
            oForm.Items.Item("lnkaccod").LinkTo = "txtaccod"
            oForm.Items.Item("lnkaccod").Top = 66
            oForm.Items.Item("lnkaccod").Left = 226
            oForm.Items.Item("lnkaccod").Description = "Link to" & vbNewLine & "Chart Of Accounts"
            oLBAcctCodeLink = oForm.Items.Item("lnkaccod").Specific
            oLBAcctCodeLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_GLAccounts

            oLBAcctNameTxt = oForm.Items.Item("txtacnam").Specific
            oLBAcctNameTxt.DataBind.SetBound(True, "@PSSIT_OLBR", "U_Acname")

            oInfo1Txt = oForm.Items.Item("txtadnl1").Specific
            oForm.Items.Item("lbladnl1").Visible = False
            oForm.Items.Item("txtadnl1").Visible = False
            oInfo1Txt.DataBind.SetBound(True, "@PSSIT_OLBR", "U_Adnl1")

            oInfo2Txt = oForm.Items.Item("txtadnl2").Specific
            oForm.Items.Item("lbladnl2").Visible = False
            oForm.Items.Item("txtadnl2").Visible = False
            oInfo2Txt.DataBind.SetBound(True, "@PSSIT_OLBR", "U_Adnl2")

            oRemarksTxt = oForm.Items.Item("txtremark").Specific
            oRemarksTxt.DataBind.SetBound(True, "@PSSIT_OLBR", "U_Remarks")

            oActiveCheck = oForm.Items.Item("chkactive").Specific
            oActiveCheck.DataBind.SetBound(True, "@PSSIT_OLBR", "U_Active")
            oActiveCheck.Checked = True

            BtnEmp = oForm.Items.Item("btnemp").Specific
            oForm.Items.Item("btnemp").Description = "Choose from List" & vbNewLine & "Employee List View"
            BtnEmp.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            BtnEmp.Image = sPath & "\Resources\CFL.bmp"
            BtnEmp = oForm.Items.Item("btnemp").Specific
            BtnEmp.ChooseFromListUID = "BtEmpLst"

            BtnSG = oForm.Items.Item("btnskill").Specific
            oForm.Items.Item("btnskill").Description = "Choose from List" & vbNewLine & "Skill Group List View"
            BtnSG.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            BtnSG.Image = sPath & "\Resources\CFL.bmp"
            BtnSG = oForm.Items.Item("btnskill").Specific
            BtnSG.ChooseFromListUID = "BtSGLst"

            BtnLAcct = oForm.Items.Item("btnlacct").Specific
            oForm.Items.Item("btnlacct").Description = "Choose from List" & vbNewLine & "Account Detail List View"
            BtnLAcct.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            BtnLAcct.Image = sPath & "\Resources\CFL.bmp"
            BtnLAcct = oForm.Items.Item("btnlacct").Specific
            BtnLAcct.ChooseFromListUID = "BtLAccLst"

            oLBActAcCodeTxt = oForm.Items.Item("txtaccode").Specific
            oForm.Items.Item("txtaccode").Enabled = False
            'oForm.Items.Item("txtaccode").Visible = False
            'oForm.Items.Item("lblaccode").Visible = False
            oLBActAcCodeTxt.DataBind.SetBound(True, "@PSSIT_OLBR", "U_ActAcCode")
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
            '*****************Employee-CFL**************************
            oChEmpList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "171", "EmpLst"))
            oChEmpBtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "171", "BtEmpLst"))
            '**************************Skill Group-CFL****************************
            oChSGList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_LGP", "SGLst"))
            CreateNewConditions(oChSGList, "U_Active", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            oChSGBtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_LGP", "BtSGLst"))
            CreateNewConditions(oChSGBtnList, "U_Active", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            '***********************************Accounts-CFL********************************
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
    Private Sub Labour_ChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable) Handles Me.ChooseFromList
        Dim oDataTable As SAPbouiCOM.DataTable = ChooseFromListSelectedObjects
        Dim oCurrentRow As Integer
        Dim oEmpId, oEmpName, oSGNo, oSGName, oLAccCode, oLAccName As String
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim StrSql, StrSql1 As String
        Try
            oCurrentRow = CurrentRow
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '*************Employee CFL**************
            If (ControlName = "txtempid" Or ControlName = "btnemp") And (ChoosefromListUID = "EmpLst" Or ChoosefromListUID = "BtEmpLst") Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    If Not oDataTable Is Nothing Then
                        oEmpId = oDataTable.GetValue("empID", 0)
                        oEmpName = oDataTable.GetValue("firstName", 0)
                        oParentDB.SetValue("U_Empid", oParentDB.Offset, oEmpId)
                        oParentDB.SetValue("U_Empnam", oParentDB.Offset, oEmpName)
                    End If
                Else
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        If Not oDataTable Is Nothing Then
                            If Not oDataTable Is Nothing Then
                                oEmpId = oDataTable.GetValue("empID", 0)
                                oEmpName = oDataTable.GetValue("firstName", 0)
                                oParentDB.SetValue("U_Empid", oParentDB.Offset, oEmpId)
                                oParentDB.SetValue("U_Empnam", oParentDB.Offset, oEmpName)
                            End If
                        End If
                    End If
                End If
            End If
            '*************Skill Group CFL**************
            If (ControlName = "txtgcode" Or ControlName = "btnskill") And (ChoosefromListUID = "SGLst" Or ChoosefromListUID = "BtSGLst") Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    If Not oDataTable Is Nothing Then
                        oSGNo = oDataTable.GetValue("Code", 0)
                        oSGName = oDataTable.GetValue("U_LGname", 0)
                        oParentDB.SetValue("U_LGCode", oParentDB.Offset, oSGNo)
                        oParentDB.SetValue("U_LGname", oParentDB.Offset, oSGName)
                        StrSql = "select U_Labrate,U_Accode,U_Acname from [@PSSIT_OLGP] where code='" & oSGNo & "'"
                        oRs.DoQuery(StrSql)
                        If oRs.RecordCount > 0 Then
                            oRs.MoveFirst()
                            oParentDB.SetValue("U_Labrate", oParentDB.Offset, oRs.Fields.Item("U_Labrate").Value)
                            oParentDB.SetValue("U_Accode", oParentDB.Offset, oRs.Fields.Item("U_Accode").Value)
                            oParentDB.SetValue("U_Acname", oParentDB.Offset, oRs.Fields.Item("U_Acname").Value)
                            oParentDB.SetValue("U_ActAcCode", oParentDB.Offset, oRs1.Fields.Item("U_Accode").Value.ToString().Replace("-", ""))
                        End If
                    End If
                Else
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        If Not oDataTable Is Nothing Then
                            oSGNo = oDataTable.GetValue("Code", 0)
                            oSGName = oDataTable.GetValue("U_LGname", 0)
                            oParentDB.SetValue("U_LGCode", oParentDB.Offset, oSGNo)
                            oParentDB.SetValue("U_LGname", oParentDB.Offset, oSGName)
                            StrSql1 = "select U_Labrate,U_Accode,U_Acname from [@PSSIT_OLGP] where code='" & oSGNo & "'"
                            oRs1.DoQuery(StrSql1)
                            If oRs1.RecordCount > 0 Then
                                oRs1.MoveFirst()
                                oParentDB.SetValue("U_Labrate", oParentDB.Offset, oRs1.Fields.Item("U_Labrate").Value)
                                oParentDB.SetValue("U_Accode", oParentDB.Offset, oRs1.Fields.Item("U_Accode").Value)
                                oParentDB.SetValue("U_Acname", oParentDB.Offset, oRs1.Fields.Item("U_Acname").Value)
                                oParentDB.SetValue("U_ActAcCode", oParentDB.Offset, oRs1.Fields.Item("U_Accode").Value.ToString().Replace("-", ""))
                            End If
                        End If
                    End If
                End If
            End If
            '***********Labour Account CFL********************
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
            oRs1 = Nothing
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
            oLbrCurrCombo = oForm.Items.Item("cmbcurncy").Specific
            'oForm.Items.Item("cmbcurncy").Enabled = True

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
            oForm.Items.Item("cmbcurncy").Enabled = False
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
            If pVal.FormUID = "FLM" Then
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
                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                        Try
                            Validation()
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End Try
                    End If
                End If
                '*** Update mode ****
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And pVal.BeforeAction = True Then
                        Try
                            val1()
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End Try
                    End If
                End If
                '********

                '********** Add Button Press ***********
                If pVal.ItemUID = "1" And (pVal.BeforeAction = False) And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    Try
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            oForm.Items.Item("txtlrcode").Enabled = False
                        End If
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            oForm.Refresh()
                            oForm.Freeze(True)
                            oActiveCheck.Checked = True
                            CurrCombo()
                            oForm.Items.Item("cmbcurncy").Enabled = False
                            oForm.Freeze(False)
                            oLGCodeTxt.Active = True
                        End If
                    Catch ex As Exception
                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    End Try
                End If
                If pVal.ItemUID = "lnkskill" And pVal.BeforeAction = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    Dim oSkillCode As String
                    oSkillCode = oSGNoTxt.Value
                    SkillGroupClass = New SkillGroups(SBO_Application, oCompany, oSkillCode, "Labour")
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
        Dim oRs, oRs1 As SAPbobsCOM.Recordset
        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            oRs.DoQuery("select Code from [@PSSIT_OLBR]  where Code= '" & oLGCodeTxt.Value & "' ")
            If oRs.RecordCount > 0 Then
                oLGCodeTxt.Active = True
                Throw New Exception("Labour Code Already Exist")
            End If
            If oLGCodeTxt.Value.Length = 0 Then
                oLGCodeTxt.Active = True
                Throw New Exception("Labour Code should not be Empty")
            End If
            'Added by Manimaran------s
            If oEmpNoTxt.Value.Length = 0 Then
                oEmpNoTxt.Active = True
                Throw New Exception("Employee Code should not be Empty")
            End If
            If oForm.Items.Item("txtaccod").Specific.value.length = 0 Then
                Throw New Exception("Account Code should not be empty")
            End If
            'Added by Manimaran------e
            oRs1.DoQuery("select Code from [@PSSIT_OLBR]  where U_Empid= '" & oEmpNoTxt.Value & "' ")
            If oRs1.RecordCount > 0 Then
                oEmpNoTxt.Active = True
                Throw New Exception("Employee Code Already Exist")
            End If
            If oSGNoTxt.Value.Length = 0 Then
                oSGNoTxt.Active = True
                Throw New Exception("Skill Group should not be Empty")
            End If
            If oLbrRateTxt.Value.Length = 0 Or oLbrRateTxt.Value = "0.0" Then
                oLbrRateTxt.Active = True
                Throw New Exception("Labour Rate should be greater than zero")
            End If
            If CInt(oLbrRateTxt.Value) < 0 Then
                oLbrRateTxt.Active = True
                Throw New Exception("Labour Rate should be greater than zero")
            End If

            AccKeyCheck()
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            oRs1 = Nothing
            GC.Collect()
        End Try
    End Sub
    Private Sub val1()
        Dim oRs, oRs1 As SAPbobsCOM.Recordset
        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            If CInt(oLbrRateTxt.Value) < 0 Then
                oLbrRateTxt.Active = True
                Throw New Exception("Labour Rate should be greater than zero")
            End If
            'Added by Manimaran------s
            If oEmpNoTxt.Value.Length = 0 Then
                oEmpNoTxt.Active = True
                Throw New Exception("Employee Code should not be Empty")
            End If
            If oSGNoTxt.Value.Length = 0 Then
                oSGNoTxt.Active = True
                Throw New Exception("Skill Group should not be Empty")
            End If
            If oForm.Items.Item("txtaccod").Specific.value.length = 0 Then
                Throw New Exception("Account Code should not be empty")
            End If
            'Added by Manimaran------e
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
            If pVal.MenuUID = "1281" And FormID = "FLM" Then
                If pVal.BeforeAction = True Then
                    oForm.Items.Item("txtlrcode").Enabled = True
                End If
                If pVal.BeforeAction = False Then
                    oLGCodeTxt.Active = True
                End If
            End If
            If pVal.MenuUID = "1282" And pVal.BeforeAction = False And FormID = "FLM" Then
                oForm.Freeze(True)
                oParentDB.Offset = oParentDB.Size - 1
                oActiveCheck.Checked = True
                oForm.Items.Item("txtlrcode").Enabled = True
                oForm.Items.Item("cmbcurncy").Enabled = False
                CurrCombo()
                oForm.Freeze(False)
                oLGCodeTxt.Active = True
            End If
            '*****************************Navigation*******************************
            If (pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") And pVal.BeforeAction = False And FormID = "FLM" Then
                Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Try
                    oRs.DoQuery("Select * from [@PSSIT_OLBR]")
                    If oRs.RecordCount > 0 Then
                        oForm.Items.Item("txtlrcode").Enabled = False
                    Else
                        oForm.Items.Item("txtlrcode").Enabled = True
                        oLGCodeTxt.Active = True
                    End If
                Catch ex As Exception
                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End Try
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
End Class
