'''' <summary>
'''' Author                     Created Date
'''' Suresh                      19/12/2008
'''' <remarks> This class is used for entering the shift details.</remarks>
Public Class MachineMaster
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
    '**************************DataSource************************************
    Private oParentDB As SAPbouiCOM.DBDataSource
    Private oCriticalDB, oSpecDB, oParamDB, oSftDB As SAPbouiCOM.DBDataSource
    '**************************UserDataSource************************************
    Private oUD, oITmUD As SAPbouiCOM.UserDataSource
    '**************************ChooseFromList************************************
    Private oChWCList, oChWCBtnList, oChMGList, oChMGBtnList, oChAccList, oChABtnList, oChSAccList, oChSABtnList, oChBPList, oChBPBtnList, oChItemList, oChSftList, oChSftBtnList, oPOCFL As SAPbouiCOM.ChooseFromList
    Private Event ChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable)
    '**************************Items - EditText************************************
    '**************************Header-EditText**************************************
    Private oCodeTxt, oWCTCodeTxt, oModCodeTxt, oMakCodeTxt, oWcNamTxt, oSrtNamTxt, oWCNoTxt, oManNoTxt, oDeptCodTxt, oDeptNoTxt, oOprtHrsTxt, oMsrCodeTxt, oSpcLtTxt, oInstlDtTxt, oWCCapTxt, oSpcBrthTxt, oInstlCpTxt, oMachineGrpCodeTxt, oMachineGrpNameTxt, oInfo1Txt, oInfo2Txt, oRemarksTxt, oRActAcCodeTxt, oSActAcCodeTxt As SAPbouiCOM.EditText
    '*************************Purchase Info-EditText********************************
    Private oBpTxt, oPONoTxt, oPODtTxt, oWrtyExpTxt As SAPbouiCOM.EditText
    '*************************Cost Details-EditText*********************************
    Private oOprCostTxt, oPwrCostTxt, oSetupCost, oCost1Txt, oCost2Txt As SAPbouiCOM.EditText
    '************************Account Details Details-EditText***********************
    Private oRAcctCodeTxt, oRAcctNameTxt, oSAcctCodeTxt, oSAcctNameTxt As SAPbouiCOM.EditText
    '************************Shift Details Details-EditText***********************
    Private oShiftCodeTxt, oShiftNameTxt, oDurationTxt As SAPbouiCOM.EditText
    '**************************Items - ComboBox************************************
    '**************************Header-ComboBox**********************************
    Private oWCTypeCombo, oModlCombo, oMakCombo, oYrMakCombo, oDeptCombo, oGrpUndrCombo, oMsrUntCombo, oStsCombo As SAPbouiCOM.ComboBox
    '**************************Items -CheckBox*********************************** 
    Private oGrpCheck As SAPbouiCOM.CheckBox
    '**************************Items -Button*********************************
    Private oWCBtn, oMGBtn, oAcctBtn, oSAcctBtn, oBPBtn, oSftBtn As SAPbouiCOM.Button
    '**************************Items -Folders*************************************
    Private oPIFldr, oSpecFldr, oPPFldr, oCIFldr, oCDFldr, oACFldr, oSFTFldr As SAPbouiCOM.Folder
    '**************************Items - Matrix************************************
    Private oSMatrix, oPMatrix, oCMatrix, oSFMatrix As SAPbouiCOM.Matrix
    Private oSColumns, oPColumns, oCColumns, oSFcolumns As SAPbouiCOM.Columns
    Private oSColumn, oPColumn, oCColumn, oSFColoumn As SAPbouiCOM.Column
    '**************************Specification Matix Columns*************************
    Private oSpecIdCol, oSpecValCol As SAPbouiCOM.Column
    '**************************Production Parameter******************************
    Private oParIdCol, oParNameCol, oParValCol As SAPbouiCOM.Column
    '***************************Critical Item Matrix Column************************
    Private oItemIdCol, oItemDescCol, oInstlDtCol, oLfSpanCol, oQtyCol, oUnitsCol As SAPbouiCOM.Column
    '**************************Shift Details******************************
    Private oSftCodeCol, oSftNameCol, oSftDurCol As SAPbouiCOM.Column
    '**************************Items - CheckBox************************************
    Private oReCondCheck, oActiveCheck As SAPbouiCOM.CheckBox
    '**************************Link Button************************************
    Private oWCCodeLink, oMGCodeLink, oBPCodeLink, oAcctCodeLink, oSAcctCodeLink, oShiftCodeLink As SAPbouiCOM.LinkedButton
    '************************Variables*************************************
    Private sPath As String = IO.Directory.GetParent(Application.ExecutablePath).ToString
    Private BoolResize As Boolean
    Private BoolWCType As Boolean = True
    Private BoolModel As Boolean = True
    Private BoolMake As Boolean = True
    Private BoolDept As Boolean = True
    Private BoolUOM As Boolean = True
    Private SpecUID, CritUID, ParamUID, ShiftUID As String
    Private fSettings As SAPbouiCOM.FormSettings
    Private oMachineNo, oFormName As String
    Private WithEvents WorkCentreClass As WorkCentre
    Private WithEvents MachineGroupsClass As MachineGroups
    Private WithEvents ShiftClass As Shift
#End Region
    ''' <summary>
    ''' SetApplication(),LoadFromXML(),DrawForm() methods are called. 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company, Optional ByVal aMachineNo As String = Nothing, Optional ByVal aFormName As String = Nothing)
        MyBase.New(aSBO_Application, aCompany)
        SBO_Application = aSBO_Application
        oCompany = aCompany
        oMachineNo = aMachineNo
        oFormName = aFormName
        LoadFromXML("FrmMachine.srf")
        DrawForm()
        If oFormName = "Operation" Or oFormName = "DownTime" Or oFormName = "OprRouting" Or oFormName = "MCUtilRpt" Or oFormName = "MCPerfRpt" Then
            oParentDB.SetValue("U_wcno", oParentDB.Offset, oMachineNo)
            oWCNoTxt.Value = oMachineNo
            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        End If
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
            oForm = SBO_Application.Forms.Item("FM")
            fSettings = oForm.Settings
            oParentDB = oForm.DataSources.DBDataSources.Item("@PSSIT_PMWCHDR")
            oSpecDB = oForm.DataSources.DBDataSources.Item("@PSSIT_PMWCSPEC")     ' Specification Datasource
            oParamDB = oForm.DataSources.DBDataSources.Item("@PSSIT_PMWCPARA")    'Production Parameter Datasource      
            oCriticalDB = oForm.DataSources.DBDataSources.Item("@PSSIT_PMWCITEM") 'Critical Item Datasource
            oSftDB = oForm.DataSources.DBDataSources.Item("@PSSIT_PMWCSFT") 'Shift Detail Datasource
            oForm.Freeze(True)
            LoadLookups()             'Loading CFL    
            InitTxtComp()             'Add EditText
            InitCbopComp()            'Add ComboBox 
            InitOthrComp()            'Add OtherComponents 
            InitMatrix()              'Add Matrix 

            ModelCombo()       'Model Combo Load
            MakeCombo()         'Make Combo Load  
            YearMake()        'Year Make Combo Load
            MeasUnitCombo()  'Measurement Unit Combo Load
            UnitsCombo()
            StatusCombo(oStsCombo)

            oForm.Freeze(False)
            SetToolBarEnabled()

            oForm.DataBrowser.BrowseBy = "txtmcno"
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
            '***************************Header***********************
            oCodeTxt = oForm.Items.Item("txtcode").Specific
            oCodeTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "Code")

            oWCNoTxt = oForm.Items.Item("txtmcno").Specific
            oWCNoTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_wcno")

            oWcNamTxt = oForm.Items.Item("txtmcname").Specific
            oWcNamTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_wcname")

            oSrtNamTxt = oForm.Items.Item("txtsrtname").Specific
            oSrtNamTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_wcshname")

            oManNoTxt = oForm.Items.Item("txtmanno").Specific
            oManNoTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_mfserial")

            oWCTCodeTxt = oForm.Items.Item("txtmctpcod").Specific
            oWCTCodeTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_typecode")
            oForm.Items.Item("txtmctpcod").Visible = False
            oForm.Items.Item("lblmactype").Visible = False

            oModCodeTxt = oForm.Items.Item("txtmodcode").Specific
            oModCodeTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_modecode")

            oMakCodeTxt = oForm.Items.Item("txtmakcode").Specific
            oMakCodeTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_makecode")

            oDeptNoTxt = oForm.Items.Item("txtdeptno").Specific
            oDeptNoTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_deptcode")

            oDeptCodTxt = oForm.Items.Item("txtdept").Specific
            oDeptCodTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_deptdesc")
            oDeptCodTxt.ChooseFromListUID = "WCLst"
            oDeptCodTxt.ChooseFromListAlias = "Code"
            oForm.Items.Item("txtdept").LinkTo = "lnkwccod"
            oForm.Items.Add("lnkwccod", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnkwccod").Visible = True
            oForm.Items.Item("lnkwccod").LinkTo = "txtdept"
            oForm.Items.Item("lnkwccod").Top = 66
            oForm.Items.Item("lnkwccod").Left = 135
            oForm.Items.Item("lnkwccod").Description = "Link to" & vbNewLine & "Work Centre"
            oWCCodeLink = oForm.Items.Item("lnkwccod").Specific

            oOprtHrsTxt = oForm.Items.Item("txtoprhrs").Specific
            oOprtHrsTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_ohrsday")

            oMsrCodeTxt = oForm.Items.Item("txtmsrcode").Specific
            oMsrCodeTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_uomcode")

            oSpcLtTxt = oForm.Items.Item("txtspreqlt").Specific
            oSpcLtTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_spacelen")

            oInstlDtTxt = oForm.Items.Item("txtInsDate").Specific
            oInstlDtTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_insdate")

            oWCCapTxt = oForm.Items.Item("txtmaccap").Specific
            oWCCapTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_wccapa")

            oSpcBrthTxt = oForm.Items.Item("txtspreqbt").Specific
            oSpcBrthTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_spacebre")

            oInstlCpTxt = oForm.Items.Item("txtinscap").Specific
            oInstlCpTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_inskw")

            oMachineGrpCodeTxt = oForm.Items.Item("txtmcgroup").Specific
            oMachineGrpCodeTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_MGcode")
            oMachineGrpCodeTxt.ChooseFromListUID = "MGLst"
            oMachineGrpCodeTxt.ChooseFromListAlias = "Code"
            oForm.Items.Item("txtmcgroup").LinkTo = "lnkmgcod"
            oForm.Items.Add("lnkmgcod", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnkmgcod").Visible = True
            oForm.Items.Item("lnkmgcod").LinkTo = "txtmcgroup"
            oForm.Items.Item("lnkmgcod").Top = 96
            oForm.Items.Item("lnkmgcod").Left = 135
            oForm.Items.Item("lnkmgcod").Description = "Link to" & vbNewLine & "Machine Group"
            oMGCodeLink = oForm.Items.Item("lnkmgcod").Specific

            oMachineGrpNameTxt = oForm.Items.Item("txtmcgpcod").Specific
            oMachineGrpNameTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_MGname")
            oForm.Items.Item("txtmcgpcod").Enabled = False

            '*************************Purchase Info********************************
            oBpTxt = oForm.Items.Item("txtbp").Specific
            oBpTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_bpcode")
            oBpTxt.ChooseFromListUID = "BPLst"
            oBpTxt.ChooseFromListAlias = "CardCode"
            oForm.Items.Item("txtbp").LinkTo = "lnkbpcod"
            oForm.Items.Add("lnkbpcod", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnkbpcod").Visible = True
            oForm.Items.Item("lnkbpcod").LinkTo = "txtbp"
            oForm.Items.Item("lnkbpcod").Top = 198
            oForm.Items.Item("lnkbpcod").Left = 135
            oForm.Items.Item("lnkbpcod").FromPane = 1
            oForm.Items.Item("lnkbpcod").ToPane = 1
            oForm.Items.Item("lnkbpcod").Description = "Link to" & vbNewLine & "Business Partnere List View"
            oBPCodeLink = oForm.Items.Item("lnkbpcod").Specific
            oBPCodeLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner

            oPONoTxt = oForm.Items.Item("txtpono").Specific
            oPONoTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_pono")
            'Added by Manimaran-----s
            oPONoTxt.ChooseFromListUID = "POLst"
            oPONoTxt.ChooseFromListAlias = "DocEntry"
            'Added by Manimaran-----e


            oPODtTxt = oForm.Items.Item("txtpodt").Specific
            oPODtTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_podate")

            oWrtyExpTxt = oForm.Items.Item("txtwarrexp").Specific
            oWrtyExpTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_wardate")

            oOprCostTxt = oForm.Items.Item("txtoprcost").Specific
            oOprCostTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_opercost")

            oPwrCostTxt = oForm.Items.Item("txtpwrcost").Specific
            oPwrCostTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_powecost")

            oSetupCost = oForm.Items.Item("txtsetup").Specific
            oSetupCost.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_Setupcost")

            oCost1Txt = oForm.Items.Item("txtcost1").Specific
            oCost1Txt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_cost1")

            oCost2Txt = oForm.Items.Item("txtcost2").Specific
            oCost2Txt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_cost2")

            '************************Running Rate Accounts Info****************************
            oRAcctCodeTxt = oForm.Items.Item("txtacccode").Specific
            oRAcctCodeTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_Accode")
            oRAcctCodeTxt.ChooseFromListUID = "AccLst"
            oRAcctCodeTxt.ChooseFromListAlias = "AcctCode"
            oForm.Items.Item("txtacccode").LinkTo = "lnkraccod"
            oForm.Items.Add("lnkraccod", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnkraccod").Visible = True
            oForm.Items.Item("lnkraccod").LinkTo = "txtacccode"
            oForm.Items.Item("lnkraccod").Top = 198
            oForm.Items.Item("lnkraccod").Left = 135
            oForm.Items.Item("lnkraccod").FromPane = 6
            oForm.Items.Item("lnkraccod").ToPane = 6
            oForm.Items.Item("lnkraccod").Description = "Link to" & vbNewLine & "Chart Of Accounts"
            oAcctCodeLink = oForm.Items.Item("lnkraccod").Specific
            oAcctCodeLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_GLAccounts

            oRAcctNameTxt = oForm.Items.Item("txtaccdesc").Specific
            oRAcctNameTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_Acname")

            '************************Setup Rate Accounts Info******************************
            oSAcctCodeTxt = oForm.Items.Item("txtsetupac").Specific
            oSAcctCodeTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_SAccode")
            oSAcctCodeTxt.ChooseFromListUID = "SAccLst"
            oSAcctCodeTxt.ChooseFromListAlias = "AcctCode"
            oForm.Items.Item("txtsetupac").LinkTo = "lnksaccod"
            oForm.Items.Add("lnksaccod", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnksaccod").Visible = True
            oForm.Items.Item("lnksaccod").LinkTo = "txtsetupac"
            oForm.Items.Item("lnksaccod").Top = 228
            oForm.Items.Item("lnksaccod").Left = 135
            oForm.Items.Item("lnksaccod").FromPane = 6
            oForm.Items.Item("lnksaccod").ToPane = 6
            oForm.Items.Item("lnksaccod").Description = "Link to" & vbNewLine & "Chart Of Accounts"
            oSAcctCodeLink = oForm.Items.Item("lnksaccod").Specific
            oSAcctCodeLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_GLAccounts

            oSAcctNameTxt = oForm.Items.Item("txtstaccod").Specific
            oSAcctNameTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_SAcname")

            ''************************Shift Info******************************
            'oShiftCodeTxt = oForm.Items.Item("txtsftno").Specific
            'oShiftCodeTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_SCode")
            'oShiftCodeTxt.ChooseFromListUID = "SftLst"
            'oShiftCodeTxt.ChooseFromListAlias = "Code"
            'oForm.Items.Item("txtsftno").LinkTo = "lnksft"
            'oForm.Items.Add("lnksft", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            'oForm.Items.Item("lnksft").Visible = True
            'oForm.Items.Item("lnksft").LinkTo = "txtsftno"
            'oForm.Items.Item("lnksft").Top = 198
            'oForm.Items.Item("lnksft").Left = 135
            'oForm.Items.Item("lnksft").FromPane = 7
            'oForm.Items.Item("lnksft").ToPane = 7
            'oForm.Items.Item("lnksft").Description = "Link to" & vbNewLine & "Shift Details"
            'oShiftCodeLink = oForm.Items.Item("lnksft").Specific

            'oShiftNameTxt = oForm.Items.Item("txtsftname").Specific
            'oShiftNameTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_Sdescr")

            'oDurationTxt = oForm.Items.Item("txtdur").Specific
            'oDurationTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_Duratmin")

            '***********************Footer Info*********************************
            oInfo1Txt = oForm.Items.Item("txtadnl1").Specific
            oForm.Items.Item("lbloi1").Visible = False
            oForm.Items.Item("txtadnl1").Visible = False
            oInfo1Txt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_Adnl1")

            oInfo2Txt = oForm.Items.Item("txtadnl2").Specific
            oForm.Items.Item("lbloi2").Visible = False
            oForm.Items.Item("txtadnl2").Visible = False
            oInfo2Txt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_Adnl2")

            oRemarksTxt = oForm.Items.Item("txtremark").Specific
            oRemarksTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_Remarks")

            oRActAcCodeTxt = oForm.Items.Item("txtaccode").Specific
            oForm.Items.Item("txtaccode").Enabled = False
            'oForm.Items.Item("txtaccode").Visible = False
            'oForm.Items.Item("lblaccode").Visible = False
            oRActAcCodeTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_ActAcCode")

            oSActAcCodeTxt = oForm.Items.Item("txtsaccode").Specific
            oForm.Items.Item("txtsaccode").Enabled = False
            'oForm.Items.Item("txtsaccode").Visible = False
            'oForm.Items.Item("lblsaccode").Visible = False
            oSActAcCodeTxt.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_SActAcCode")

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Add ComboBox Items
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitCbopComp()
        Try
            oWCTypeCombo = oForm.Items.Item("cmbmactype").Specific
            oWCTypeCombo.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_typedesc")
            oForm.Items.Item("cmbmactype").Visible = False

            oModlCombo = oForm.Items.Item("cmbmodel").Specific
            oModlCombo.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_modedesc")

            oMakCombo = oForm.Items.Item("cmbmake").Specific
            oMakCombo.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_makedesc")

            oYrMakCombo = oForm.Items.Item("cmbyrmake").Specific
            oYrMakCombo.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_yearmake")
            oYrMakCombo.Select(Date.Today, SAPbouiCOM.BoSearchKey.psk_ByValue)

            oGrpUndrCombo = oForm.Items.Item("cmbgrpund").Specific
            oGrpUndrCombo.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_undergrp")
            oForm.Items.Item("cmbgrpund").Visible = False
            oForm.Items.Item("lblgrpund").Visible = False

            oMsrUntCombo = oForm.Items.Item("cmbmsrunit").Specific
            oMsrUntCombo.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_uomdesc")

            oStsCombo = oForm.Items.Item("cmbsts").Specific
            oStsCombo.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_status")
            oForm.Items.Item("cmbsts").Visible = False
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Adding Other Components Items
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitOthrComp()
        Try
            '****************Add CheckBox Items**************************
            oGrpCheck = oForm.Items.Item("chkisgrp").Specific
            oGrpCheck.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_isgrp")
            oForm.Items.Item("chkisgrp").Visible = False

            oActiveCheck = oForm.Items.Item("chkactive").Specific
            oActiveCheck.DataBind.SetBound(True, "@PSSIT_PMWCHDR", "U_Active")
            oActiveCheck.Checked = True

            '*********************Add Button Item***************************
            '*************Work Centre Button******************
            oWCBtn = oForm.Items.Item("btndept").Specific
            oForm.Items.Item("btndept").Description = "Choose from List" & vbNewLine & "Work Centre List View"
            oWCBtn.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            oWCBtn.Image = sPath & "\Resources\CFL.bmp"
            oWCBtn = oForm.Items.Item("btndept").Specific
            oWCBtn.ChooseFromListUID = "BtWCLst"

            '*************Machine Group Button******************
            oMGBtn = oForm.Items.Item("btnmcgrp").Specific
            oForm.Items.Item("btnmcgrp").Description = "Choose from List" & vbNewLine & "Machine Group List View"
            oMGBtn.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            oMGBtn.Image = sPath & "\Resources\CFL.bmp"
            oMGBtn = oForm.Items.Item("btnmcgrp").Specific
            oMGBtn.ChooseFromListUID = "BtMGLst"

            '*************Accounts Button******************
            oAcctBtn = oForm.Items.Item("btnacct").Specific
            oForm.Items.Item("btnacct").Description = "Choose from List" & vbNewLine & "Accounts List View"
            oAcctBtn.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            oAcctBtn.Image = sPath & "\Resources\CFL.bmp"
            oAcctBtn = oForm.Items.Item("btnacct").Specific
            oAcctBtn.ChooseFromListUID = "BtAccLst"

            '*************Setup Accounts Button******************
            oSAcctBtn = oForm.Items.Item("btnsacct").Specific
            oForm.Items.Item("btnsacct").Description = "Choose from List" & vbNewLine & "Accounts List View"
            oSAcctBtn.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            oSAcctBtn.Image = sPath & "\Resources\CFL.bmp"
            oSAcctBtn = oForm.Items.Item("btnsacct").Specific
            oSAcctBtn.ChooseFromListUID = "BtSAccLst"

            '*************Business Partner Button******************
            oBPBtn = oForm.Items.Item("btnbp").Specific
            oForm.Items.Item("btnbp").Description = "Choose from List" & vbNewLine & "Business Partnere List View"
            oBPBtn.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            oBPBtn.Image = sPath & "\Resources\CFL.bmp"
            oBPBtn = oForm.Items.Item("btnbp").Specific
            oBPBtn.ChooseFromListUID = "BtBPLst"

            ''*************Shift Button******************
            'oSftBtn = oForm.Items.Item("btnsft").Specific
            'oForm.Items.Item("btnsft").Description = "Choose from List" & vbNewLine & "Shift List View"
            'oSftBtn.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            'oSftBtn.Image = sPath & "\Resources\CFL.bmp"
            'oSftBtn = oForm.Items.Item("btnsft").Specific
            'oSftBtn.ChooseFromListUID = "BtSftLst"

            '************************Folders*****************************
            oUD = oForm.DataSources.UserDataSources.Add("FolderPM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            Dim I As Integer
            Try
                For I = 1 To 7
                    If I = 1 Then
                        oPIFldr = oForm.Items.Item("folpi").Specific
                        oForm.Items.Item("folpi").AffectsFormMode = False
                        oPIFldr.DataBind.SetBound(True, "", "FolderPM")
                        oPIFldr.Select()
                    ElseIf I = 2 Then
                        oSpecFldr = oForm.Items.Item("folspec").Specific
                        oForm.Items.Item("folspec").AffectsFormMode = False
                        oSpecFldr.DataBind.SetBound(True, "", "FolderPM")
                        fSettings.MatrixUID = "matspec"
                        oSpecFldr.GroupWith("folpi")
                    ElseIf I = 3 Then
                        oPPFldr = oForm.Items.Item("folpp").Specific
                        oForm.Items.Item("folpp").AffectsFormMode = False
                        oPPFldr.DataBind.SetBound(True, "", "FolderPM")
                        fSettings.MatrixUID = "matpp"
                        oPPFldr.GroupWith("folspec")
                    ElseIf I = 4 Then
                        oCIFldr = oForm.Items.Item("folci").Specific
                        oForm.Items.Item("folci").AffectsFormMode = False
                        oCIFldr.DataBind.SetBound(True, "", "FolderPM")
                        fSettings.MatrixUID = "matci"
                        oCIFldr.GroupWith("folpp")
                    ElseIf I = 5 Then
                        oCDFldr = oForm.Items.Item("folcd").Specific
                        oForm.Items.Item("folcd").AffectsFormMode = False
                        oCDFldr.DataBind.SetBound(True, "", "FolderPM")
                        oCDFldr.GroupWith("folci")
                    ElseIf I = 6 Then
                        oACFldr = oForm.Items.Item("folacc").Specific
                        oForm.Items.Item("folacc").AffectsFormMode = False
                        oACFldr.DataBind.SetBound(True, "", "FolderPM")
                        oACFldr.GroupWith("folcd")
                    ElseIf I = 7 Then
                        oSFTFldr = oForm.Items.Item("folsft").Specific
                        oForm.Items.Item("folsft").AffectsFormMode = False
                        oSFTFldr.DataBind.SetBound(True, "", "FolderPM")
                        fSettings.MatrixUID = "matsft"
                        oSFTFldr.GroupWith("folacc")
                    End If
                Next
            Catch ex As Exception
                Throw ex
            End Try
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

        'fSettings = oForm.Settings
        'fSettings.MatrixUID = "MyGrid"
        Try
            '*******************Add a matrix************************
            oItem = oForm.Items.Item("matspec")
            oSMatrix = oItem.Specific
            oSMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oSColumns = oSMatrix.Columns

            '**********************Count Parameter*********************
            '********************Specification**********************
            oSpecIdCol = oSColumns.Item("specid")
            oSColumns.Item("specid").Width = 100
            oSpecIdCol.DataBind.SetBound(True, "@PSSIT_PMWCSPEC", "U_speccode")

            oSpecValCol = oSColumns.Item("specval")
            oSColumns.Item("specval").Width = 100
            oSpecValCol.DataBind.SetBound(True, "@PSSIT_PMWCSPEC", "U_specval")

            '*******************Production Parameter**********************
            oItem = oForm.Items.Item("matpp")
            oPMatrix = oItem.Specific
            oPMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oPColumns = oPMatrix.Columns

            oParIdCol = oPColumns.Item("paramid")
            oPColumns.Item("paramid").Width = 100
            oParIdCol.DataBind.SetBound(True, "@PSSIT_PMWCPARA", "U_paracode")

            oParNameCol = oPColumns.Item("paramname")
            oParNameCol.Editable = True
            oPColumns.Item("paramname").Width = 135
            oParNameCol.DataBind.SetBound(True, "@PSSIT_PMWCPARA", "U_paradesc")

            oParValCol = oPColumns.Item("paramval")
            oPColumns.Item("paramval").Width = 100
            oParValCol.DataBind.SetBound(True, "@PSSIT_PMWCPARA", "U_paraval")

            '********************Critical Item********************************
            oItem = oForm.Items.Item("matci")
            oCMatrix = oItem.Specific
            oCMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oCColumns = oCMatrix.Columns

            oItemIdCol = oCColumns.Item("itemid")
            oCColumns.Item("itemid").Width = 100
            oItemIdCol.DataBind.SetBound(True, "@PSSIT_PMWCITEM", "U_itemcode")
            oItemIdCol.ChooseFromListUID = "ILst"
            oItemIdCol.ChooseFromListAlias = "ItemCode"

            oItemDescCol = oCColumns.Item("itemdesc")
            oCColumns.Item("itemdesc").Width = 135
            oItemDescCol.DataBind.SetBound(True, "@PSSIT_PMWCITEM", "U_itemdesc")


            oInstlDtCol = oCColumns.Item("insdate")
            oCColumns.Item("insdate").Width = 75
            oInstlDtCol.DataBind.SetBound(True, "@PSSIT_PMWCITEM", "U_insdate")

            oLfSpanCol = oCColumns.Item("lifespan")
            oCColumns.Item("lifespan").Width = 75
            oLfSpanCol.DataBind.SetBound(True, "@PSSIT_PMWCITEM", "U_lifeday")

            oQtyCol = oCColumns.Item("qty")
            oCColumns.Item("units").Width = 75
            oQtyCol.DataBind.SetBound(True, "@PSSIT_PMWCITEM", "U_qty")

            oUnitsCol = oCColumns.Item("units")
            oCColumns.Item("units").Width = 75
            oUnitsCol.DataBind.SetBound(True, "@PSSIT_PMWCITEM", "U_units")

            '*******************Shift Detail**********************
            oItem = oForm.Items.Item("matsft")
            oSFMatrix = oItem.Specific
            oSFMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oSFcolumns = oSFMatrix.Columns

            oSftCodeCol = oSFcolumns.Item("sftcode")
            oSFcolumns.Item("sftcode").Width = 100
            oSftCodeCol.DataBind.SetBound(True, "@PSSIT_PMWCSFT", "U_SCode")
            oSftCodeCol.ChooseFromListUID = "SftLst"
            oSftCodeCol.ChooseFromListAlias = "Code"

            oSftNameCol = oSFcolumns.Item("shiftname")
            oSftNameCol.Editable = False
            oSFcolumns.Item("shiftname").Width = 135
            oSftNameCol.DataBind.SetBound(True, "@PSSIT_PMWCSFT", "U_Sdescr")

            oSftDurCol = oSFcolumns.Item("durmins")
            oSftDurCol.Editable = False
            oSFcolumns.Item("durmins").Width = 135
            oSftDurCol.DataBind.SetBound(True, "@PSSIT_PMWCSFT", "U_Duratmin")
        Catch ex As Exception
            Throw ex
        End Try
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
            '****************************Machine-CFL************************************
            oChWCList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_WCR", "WCLst"))
            CreateNewConditions(oChWCList, "U_Active", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            oChWCBtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_WCR", "BtWCLst"))
            CreateNewConditions(oChWCBtnList, "U_Active", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            '**********************Machine Group-CFL***************************** 
            oChMGList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_MGP", "MGLst"))
            CreateNewConditions(oChMGList, "U_Active", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            oChMGBtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_MGP", "BtMGLst"))
            CreateNewConditions(oChMGBtnList, "U_Active", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            '********************************Running Rate Accounts-CFL**************************
            oChAccList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "1", "AccLst"))
            CreateNewConditions(oChAccList, "Postable", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            oChABtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "1", "BtAccLst"))
            CreateNewConditions(oChABtnList, "Postable", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            '********************************Setup Rate Accounts-CFL**************************
            oChSAccList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "1", "SAccLst"))
            CreateNewConditions(oChSAccList, "Postable", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            oChSABtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "1", "BtSAccLst"))
            CreateNewConditions(oChSABtnList, "Postable", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            '*********************Business Partner-CFL***********************
            oChBPBtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "2", "BtBPLst"))
            CreateNewConditions(oChBPBtnList, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S")
            oChBPList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "2", "BPLst"))
            CreateNewConditions(oChBPList, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S")
            '************************Critical Item-CFL************************************
            oChItemList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "4", "ILst"))
            ' CreateNewConditions(oChItemList, "itmsgrpcod", SAPbouiCOM.BoConditionOperation.co_EQUAL, 101)
            '************************Shift-CFL************************************
            oChSftList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_SFT", "SftLst"))
            CreateNewConditions(oChSftList, "U_Active", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            'oChSftBtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_SFT", "BtSftLst"))
            'Added by Manimaran----s
            oPOCFL = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "22", "POLst"))
            'Added by Manimaran----e
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
        Dim oWCNo, oWCName, oMGCode, oMGName, oAccCode, oAccName, oSAccCode, oSAccName, oBPCode, oItemCode, oItemName, oSftCode, oSftName, oSftDur, PONum As String
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim StrSql, StrSql1 As String
        Try
            oCurrentRow = CurrentRow
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '**********************Work Centre CFL**************************
            If (ControlName = "txtdept" Or ControlName = "btndept") And (ChoosefromListUID = "WCLst" Or ChoosefromListUID = "BtWCLst") Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    If Not oDataTable Is Nothing Then
                        oWCNo = oDataTable.GetValue("Code", 0)
                        oWCName = oDataTable.GetValue("U_WCname", 0)
                        oParentDB.SetValue("U_deptcode", oParentDB.Offset, oWCNo)
                        oParentDB.SetValue("U_deptdesc", oParentDB.Offset, oWCName)
                    End If
                Else
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        If Not oDataTable Is Nothing Then
                            oWCNo = oDataTable.GetValue("Code", 0)
                            oWCName = oDataTable.GetValue("U_WCname", 0)
                            oParentDB.SetValue("U_deptcode", oParentDB.Offset, oWCNo)
                            oParentDB.SetValue("U_deptdesc", oParentDB.Offset, oWCName)
                        End If
                    End If
                End If
            End If
            '**********************Machine Group CFL**************************
            If (ControlName = "txtmcgroup" Or ControlName = "btnmcgrp") And (ChoosefromListUID = "MGLst" Or ChoosefromListUID = "BtMGLst") Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    If Not oDataTable Is Nothing Then
                        oMGCode = oDataTable.GetValue("Code", 0)
                        oMGName = oDataTable.GetValue("U_Mgname", 0)
                        oParentDB.SetValue("U_MGcode", oParentDB.Offset, oMGCode)
                        oParentDB.SetValue("U_MGname", oParentDB.Offset, oMGName)
                        StrSql = "select  U_Runrate,U_Setrate,U_RAccode,U_RAcname,U_RActAcCode,U_SAccode,U_SAcname,U_SActAcCode from [@PSSIT_OMGP]  where Code='" & oMGCode & "'"
                        oRs.DoQuery(StrSql)
                        If oRs.RecordCount > 0 Then
                            oRs.MoveFirst()
                            oParentDB.SetValue("U_opercost", oParentDB.Offset, oRs.Fields.Item("U_Runrate").Value)
                            oParentDB.SetValue("U_Setupcost", oParentDB.Offset, oRs.Fields.Item("U_Setrate").Value)
                            oParentDB.SetValue("U_Accode", oParentDB.Offset, oRs.Fields.Item("U_RAccode").Value)
                            oParentDB.SetValue("U_Acname", oParentDB.Offset, oRs.Fields.Item("U_RAcname").Value)
                            oParentDB.SetValue("U_ActAcCode", oParentDB.Offset, oRs.Fields.Item("U_RActAcCode").Value)
                            oParentDB.SetValue("U_SAccode", oParentDB.Offset, oRs.Fields.Item("U_SAccode").Value)
                            oParentDB.SetValue("U_SAcname", oParentDB.Offset, oRs.Fields.Item("U_SAcname").Value)
                            oParentDB.SetValue("U_SActAcCode", oParentDB.Offset, oRs.Fields.Item("U_SActAcCode").Value)
                        End If
                    End If
                Else
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        If Not oDataTable Is Nothing Then
                            oMGCode = oDataTable.GetValue("Code", 0)
                            oMGName = oDataTable.GetValue("U_Mgname", 0)
                            oParentDB.SetValue("U_MGcode", oParentDB.Offset, oMGCode)
                            oParentDB.SetValue("U_MGname", oParentDB.Offset, oMGName)
                            StrSql1 = "select  U_Runrate,U_Setrate,U_RAccode,U_RAcname,U_RActAcCode,U_SAccode,U_SAcname,U_SActAcCode from [@PSSIT_OMGP]  where Code='" & oMGCode & "'"
                            oRs1.DoQuery(StrSql1)
                            If oRs1.RecordCount > 0 Then
                                oRs1.MoveFirst()
                                oParentDB.SetValue("U_opercost", oParentDB.Offset, oRs1.Fields.Item("U_Runrate").Value)
                                oParentDB.SetValue("U_Setupcost", oParentDB.Offset, oRs1.Fields.Item("U_Setrate").Value)
                                oParentDB.SetValue("U_Accode", oParentDB.Offset, oRs1.Fields.Item("U_RAccode").Value)
                                oParentDB.SetValue("U_Acname", oParentDB.Offset, oRs1.Fields.Item("U_RAcname").Value)
                                oParentDB.SetValue("U_ActAcCode", oParentDB.Offset, oRs1.Fields.Item("U_RActAcCode").Value)
                                oParentDB.SetValue("U_SAccode", oParentDB.Offset, oRs1.Fields.Item("U_SAccode").Value)
                                oParentDB.SetValue("U_SAcname", oParentDB.Offset, oRs1.Fields.Item("U_SAcname").Value)
                                oParentDB.SetValue("U_SActAcCode", oParentDB.Offset, oRs1.Fields.Item("U_SActAcCode").Value)
                            End If
                        End If
                    End If
                End If
            End If
            '**********************Accounts CFL**************************
            If (ControlName = "txtacccode" Or ControlName = "btnacct") And (ChoosefromListUID = "AccLst" Or ChoosefromListUID = "BtAccLst") Then
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
            '**********************SetUp Accounts CFL**************************
            If (ControlName = "txtsetupac" Or ControlName = "btnsacct") And (ChoosefromListUID = "SAccLst" Or ChoosefromListUID = "BtSAccLst") Then
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

            '**********************Business Partner CFL**************************
            If (ControlName = "txtbp" Or ControlName = "btnbp") And (ChoosefromListUID = "BtBPLst" Or ChoosefromListUID = "BPLst") Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    If Not oDataTable Is Nothing Then
                        oBPCode = oDataTable.GetValue("CardCode", 0)
                        oParentDB.SetValue("U_bpcode", oParentDB.Offset, oBPCode)
                    End If
                Else
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        If Not oDataTable Is Nothing Then
                            oBPCode = oDataTable.GetValue("CardCode", 0)
                            oParentDB.SetValue("U_bpcode", oParentDB.Offset, oBPCode)
                        End If
                    End If
                End If
            End If
            'Added by Manimaran----s
            If (ControlName = "txtpono" And ChoosefromListUID = "POLst") Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    If Not oDataTable Is Nothing Then
                        PONum = oDataTable.GetValue("DocEntry", 0)
                        oParentDB.SetValue("U_pono", oParentDB.Offset, PONum)
                    End If
                Else
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        If Not oDataTable Is Nothing Then
                            PONum = oDataTable.GetValue("DocEntry", 0)
                            oParentDB.SetValue("U_pono", oParentDB.Offset, PONum)
                        End If
                    End If
                End If
            End If
            'Added by Manimaran----e
            '**********************Shift CFL**************************
            If (ControlName = "matsft") And (ChoosefromListUID = "SftLst") Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    If Not oDataTable Is Nothing Then
                        oSftCode = oDataTable.GetValue("Code", 0)
                        oSftName = oDataTable.GetValue("U_Sdescr", 0)
                        oSftDur = oDataTable.GetValue("U_Duratmin", 0)

                        ' ******* Add Next Row If the Shift is Selected **********
                        If CurrentRow = oSFMatrix.VisualRowCount Then
                            oSftDB.Offset = oSftDB.Size - 1
                            ShiftSetValue()
                            oSFMatrix.SetLineData(CurrentRow)
                            oSFMatrix.FlushToDataSource()
                        End If
                        oSftDB.SetValue("U_SCode", oSftDB.Offset, oSftCode)
                        oSftDB.SetValue("U_Sdescr", oSftDB.Offset, oSftName)
                        oSftDB.SetValue("U_Duratmin", oSftDB.Offset, oSftDur)
                        oSFMatrix.SetLineData(CurrentRow)
                        oSFMatrix.FlushToDataSource()
                        If Len(oSFcolumns.Item("sftcode").Cells.Item(oSFMatrix.RowCount).Specific.value) > 0 Then
                            oSftDB.InsertRecord(oSftDB.Size)
                            oSftDB.Offset = oSftDB.Size - 1
                            ShiftSetValue()
                            oSFMatrix.AddRow(1, oSFMatrix.RowCount)
                        End If
                    End If
                Else
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        If Not oDataTable Is Nothing Then
                            oSftCode = oDataTable.GetValue("Code", 0)
                            oSftName = oDataTable.GetValue("U_Sdescr", 0)
                            oSftDur = oDataTable.GetValue("U_Duratmin", 0)
                            ' ******* Add Next Row If the Shift is Selected **********
                            If CurrentRow = oSFMatrix.VisualRowCount Then
                                oSftDB.Offset = oSftDB.Size - 1
                                ShiftSetValue()
                                oSFMatrix.SetLineData(CurrentRow)
                                oSFMatrix.FlushToDataSource()
                            End If
                            oSftDB.SetValue("U_SCode", oSftDB.Offset, oSftCode)
                            oSftDB.SetValue("U_Sdescr", oSftDB.Offset, oSftName)
                            oSftDB.SetValue("U_Duratmin", oSftDB.Offset, oSftDur)
                            oSFMatrix.SetLineData(CurrentRow)
                            oSFMatrix.FlushToDataSource()
                            If Len(oSFcolumns.Item("sftcode").Cells.Item(oSFMatrix.RowCount).Specific.value) > 0 Then
                                oSftDB.InsertRecord(oSftDB.Size)
                                oSftDB.Offset = oSftDB.Size - 1
                                ShiftSetValue()
                                oSFMatrix.AddRow(1, oSFMatrix.RowCount)
                            End If
                        End If
                    End If
                End If
            End If
            '**********************Item CFL**************************
            If (ControlName = "matci") And (ChoosefromListUID = "ILst") Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        If Not oDataTable Is Nothing Then
                            oItemCode = oDataTable.GetValue("ItemCode", 0)
                            oItemName = oDataTable.GetValue("ItemName", 0)
                            ' ******* Add Next Row If the Item Code is Selected **********
                            If CurrentRow = oCMatrix.VisualRowCount Then
                                oCriticalDB.Offset = oCriticalDB.Size - 1
                                CriticalSetValue()
                                oCMatrix.SetLineData(CurrentRow)
                                oCMatrix.FlushToDataSource()
                            End If
                            oCriticalDB.SetValue("U_ItemCode", oCriticalDB.Offset, oItemCode)
                            oCriticalDB.SetValue("U_Itemdesc", oCriticalDB.Offset, oItemName)
                            oCMatrix.SetLineData(CurrentRow)
                            oCMatrix.FlushToDataSource()
                            SetInstalledDate(CurrentRow)
                            If Len(oCColumns.Item("itemid").Cells.Item(oCMatrix.RowCount).Specific.value) > 0 Then
                                oCriticalDB.InsertRecord(oCriticalDB.Size)
                                oCriticalDB.Offset = oCriticalDB.Size - 1
                                CriticalSetValue()
                                oCMatrix.AddRow(1, oCMatrix.RowCount)
                                SetInstalledDate(oCMatrix.RowCount)
                            End If
                        End If
                    End If
                Else
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        If Not oDataTable Is Nothing Then
                            oItemCode = oDataTable.GetValue("ItemCode", 0)
                            oItemName = oDataTable.GetValue("ItemName", 0)
                            ' ******* Add Next Row If the Item Code is Selected **********
                            If CurrentRow = oCMatrix.VisualRowCount Then
                                oCriticalDB.Offset = oCriticalDB.Size - 1
                                CriticalSetValue()
                                oCMatrix.SetLineData(CurrentRow)
                                oCMatrix.FlushToDataSource()
                            End If
                            oCriticalDB.SetValue("U_ItemCode", oCriticalDB.Offset, oItemCode)
                            oCriticalDB.SetValue("U_Itemdesc", oCriticalDB.Offset, oItemName)
                            oCMatrix.SetLineData(CurrentRow)
                            oCMatrix.FlushToDataSource()
                            SetInstalledDate(CurrentRow)
                            If Len(oCColumns.Item("itemid").Cells.Item(oCMatrix.RowCount).Specific.value) > 0 Then
                                oCriticalDB.InsertRecord(oCriticalDB.Size)
                                oCriticalDB.Offset = oCriticalDB.Size - 1
                                CriticalSetValue()
                                oCMatrix.AddRow(1, oCMatrix.RowCount)
                                SetInstalledDate(oCMatrix.RowCount)
                            End If
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
    Private Sub SetInstalledDate(ByVal oRowNo As Integer)
        Dim oedit As SAPbouiCOM.EditText
        Try
            oedit = oInstlDtCol.Cells.Item(oRowNo).Specific
            oCMatrix.GetLineData(oRowNo)
            oedit.String = "T" 'Date.Today.ToString("dd/MM/yyyy")
            SBO_Application.SendKeys("{TAB}")
            'oCriticalDB.SetValue("U_insdate", oCriticalDB.Offset, oedit.String)
        Catch ex As Exception
            'Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This is used to Load the Work Center Type in the Combo
    ''' </summary>
    ''' <param name="oCombo"></param>
    ''' <remarks></remarks>
    Private Sub StatusCombo(ByVal oCombo As SAPbouiCOM.ComboBox)
        Try
            If oCombo.ValidValues.Count > 0 Then
                For i As Int16 = oCombo.ValidValues.Count - 1 To 0 Step -1
                    oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            oCombo.ValidValues.Add("Opened", "")
            oCombo.ValidValues.Add("Retired", "")
            oCombo.Select("Opened", SAPbouiCOM.BoSearchKey.psk_ByValue)
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    ''' <summary>
    ''' This is used to Load the Units in the Combo
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub UnitsCombo()
        Try
            If oUnitsCol.ValidValues.Count > 0 Then
                For i As Int16 = oUnitsCol.ValidValues.Count - 1 To 0 Step -1
                    oUnitsCol.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If

            oUnitsCol.ValidValues.Add("Days", "1")
            oUnitsCol.ValidValues.Add("Months", "2")
            oUnitsCol.ValidValues.Add("Years", "3")
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    ''' <summary>
    ''' This is used to Load the Model in the Combo
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ModelCombo()
        Dim rs As SAPbobsCOM.Recordset
        Try
            rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rs.DoQuery("select Code,U_modedesc from [@PSSIT_PMWCMODEL] where code is not null and U_modedesc is not Null")
            rs.MoveFirst()
            If oModlCombo.ValidValues.Count > 0 Then
                For i As Int16 = oModlCombo.ValidValues.Count - 1 To 0 Step -1
                    oModlCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            For i As Int16 = 0 To rs.RecordCount - 1
                oModlCombo.ValidValues.Add(rs.Fields.Item(1).Value, rs.Fields.Item(0).Value)
                rs.MoveNext()
            Next
            oModlCombo.ValidValues.Add("Define New", "")
        Catch ex As Exception
            Throw ex
        Finally
            rs = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' Define New
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ModelDFN()
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim StrSql As String
        Try
            If BoolModel = False Then
                If Not oModlCombo Is Nothing Then
                    ModelCombo()
                    StrSql = "select * from [@PSSIT_PMWCMODEL] where DocEntry=(Select IsNull(Max(DocEntry),0) as Code from [@PSSIT_PMWCMODEL])"
                    oRs.DoQuery(StrSql)
                    If oRs.RecordCount > 0 Then
                        oRs.MoveFirst()
                        oModlCombo.Select(oRs.Fields.Item("U_modedesc").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                        BoolModel = True
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
    ''' This is used to Load the Make in the Combo
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub MakeCombo()
        Dim rs As SAPbobsCOM.Recordset
        Try
            rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rs.DoQuery("select Code,U_makedesc from [@PSSIT_PMWCMAKE] where code is not null and U_makedesc is not Null")
            rs.MoveFirst()
            If oMakCombo.ValidValues.Count > 0 Then
                For i As Int16 = oMakCombo.ValidValues.Count - 1 To 0 Step -1
                    oMakCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            For i As Int16 = 0 To rs.RecordCount - 1
                oMakCombo.ValidValues.Add(rs.Fields.Item(1).Value, rs.Fields.Item(0).Value)
                rs.MoveNext()
            Next
            oMakCombo.ValidValues.Add("Define New", "")
        Catch ex As Exception
            Throw ex
        Finally
            rs = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' Define New
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub MakeDFN()
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim StrSql As String
        Try
            If BoolMake = False Then
                If Not oMakCombo Is Nothing Then
                    MakeCombo()
                    StrSql = "select * from [@PSSIT_PMWCMAKE] where DocEntry=(Select IsNull(Max(DocEntry),0) as Code from [@PSSIT_PMWCMAKE])"
                    oRs.DoQuery(StrSql)
                    If oRs.RecordCount > 0 Then
                        oRs.MoveFirst()
                        oMakCombo.Select(oRs.Fields.Item("U_makedesc").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                        BoolMake = True
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
    ''' This is used to Load the Years from 1900 to the Current Year in the Combo
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub YearMake()
        Dim PrevYr, CurrYr As Double
        PrevYr = 1900
        CurrYr = Date.Today.Year

        Try
            If oYrMakCombo.ValidValues.Count > 0 Then
                For i As Int16 = oYrMakCombo.ValidValues.Count - 1 To 0 Step -1
                    oYrMakCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            For i As Int16 = PrevYr To CurrYr
                oYrMakCombo.ValidValues.Add(i, "")
            Next
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try

    End Sub
    ''' <summary>
    ''' This is used to Load the Group Under in the Combo based on the Condition
    ''' </summary>
    ''' <param name="oCombo"></param>
    ''' <remarks></remarks>
    Private Sub UnderGrpCombo(ByVal oCombo As SAPbouiCOM.ComboBox)
        Dim rs As SAPbobsCOM.Recordset
        If oGrpCheck.Checked = True Then
            Try
                rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rs.DoQuery("select U_wcname from [@PSSIT_PMWCHDR]")
                rs.MoveFirst()
                If oCombo.ValidValues.Count > 0 Then
                    For i As Int16 = oCombo.ValidValues.Count - 1 To 0 Step -1
                        oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                    Next
                End If
                For i As Int16 = 0 To rs.RecordCount - 1
                    oCombo.ValidValues.Add(rs.Fields.Item(0).Value, "")
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
    ''' <summary>
    ''' This is used to Load the Measurement Unit in the Combo
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub MeasUnitCombo()
        Dim rs As SAPbobsCOM.Recordset
        Try
            rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rs.DoQuery("select Code,U_uomdesc from [@PSSIT_PMWCUOM] where code is not null and U_uomdesc is not Null")
            rs.MoveFirst()
            If oMsrUntCombo.ValidValues.Count > 0 Then
                For i As Int16 = oMsrUntCombo.ValidValues.Count - 1 To 0 Step -1
                    oMsrUntCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            For i As Int16 = 0 To rs.RecordCount - 1
                oMsrUntCombo.ValidValues.Add(rs.Fields.Item(1).Value, rs.Fields.Item(0).Value)
                rs.MoveNext()
            Next
            oMsrUntCombo.ValidValues.Add("Define New", "")
        Catch ex As Exception
            Throw ex
        Finally
            rs = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' Define New
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub MeasUnitDFN()
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim StrSql As String
        Try
            If BoolUOM = False Then
                If Not oMsrUntCombo Is Nothing Then
                    MeasUnitCombo()
                    StrSql = "select * from [@PSSIT_PMWCUOM] where DocEntry=(Select IsNull(Max(DocEntry),0) as Code from [@PSSIT_PMWCUOM])"
                    oRs.DoQuery(StrSql)
                    If oRs.RecordCount > 0 Then
                        oRs.MoveFirst()
                        oMsrUntCombo.Select(oRs.Fields.Item("U_uomdesc").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                        BoolUOM = True
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
    ''' This is used to Set the values in the Specification matrix while Adding the empty row
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SpecSetValue()
        Try
            oSpecDB.SetValue("U_speccode", oSpecDB.Offset, "")
            oSpecDB.SetValue("U_specval", oSpecDB.Offset, "")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This is used to delete the empty rows in the Specification Matrix
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SpecDeleteEmptyRow()
        Dim oSpecIDEdit, oSpecValEdit As SAPbouiCOM.EditText
        Dim IntICount As Integer
        Try
            For IntICount = oSMatrix.RowCount To 1 Step -1
                oSpecIDEdit = oSpecIdCol.Cells.Item(IntICount).Specific
                oSpecValEdit = oSpecValCol.Cells.Item(IntICount).Specific
                If oSpecIDEdit.Value.Length = 0 And oSpecValEdit.Value.Length = 0 And oSMatrix.RowCount > 1 Then
                    oSMatrix.DeleteRow(IntICount)
                    oSMatrix.FlushToDataSource()
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
    Private Sub ParamSetValue()
        Try
            oParamDB.SetValue("U_paracode", oParamDB.Offset, "")
            oParamDB.SetValue("U_paradesc", oParamDB.Offset, "")
            oParamDB.SetValue("U_paraval", oParamDB.Offset, "")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This is used to delete the empty rows in the Production Parameter Matrix
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ParamDeleteEmptyRow()
        Dim oParamID As SAPbouiCOM.EditText
        Dim IntICount As Integer
        Try
            For IntICount = oPMatrix.RowCount To 1 Step -1
                oParamID = oParIdCol.Cells.Item(IntICount).Specific
                If oParamID.Value.Length = 0 And oPMatrix.RowCount > 1 Then
                    oPMatrix.DeleteRow(IntICount)
                    oPMatrix.FlushToDataSource()
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This is used to Set the values in the Critical Item matrix while Adding the empty row
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CriticalSetValue()
        Try
            oCriticalDB.SetValue("U_ItemCode", oCriticalDB.Offset, "")
            oCriticalDB.SetValue("U_itemdesc", oCriticalDB.Offset, "")
            oCriticalDB.SetValue("U_insdate", oCriticalDB.Offset, "")
            oCriticalDB.SetValue("U_lifeday", oCriticalDB.Offset, "")
            oCriticalDB.SetValue("U_units", oCriticalDB.Offset, "Days")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This is used to Set the values in the Shift matrix while Adding the empty row
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ShiftSetValue()
        Try
            oSftDB.SetValue("U_SCode", oSftDB.Offset, "")
            oSftDB.SetValue("U_Sdescr", oSftDB.Offset, "")
            oSftDB.SetValue("U_Duratmin", oSftDB.Offset, "")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This is used to delete the empty rows in the Critical Item Matrix
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CriticalDeleteEmptyRow()
        Dim oItemCodeEdit, oItemDescEdit As SAPbouiCOM.EditText
        Dim IntICount As Integer
        Try
            For IntICount = oCMatrix.RowCount To 1 Step -1
                oItemCodeEdit = oItemIdCol.Cells.Item(IntICount).Specific
                oItemDescEdit = oItemDescCol.Cells.Item(IntICount).Specific
                If oItemCodeEdit.Value.Length = 0 And oItemDescEdit.Value.Length = 0 And oCMatrix.RowCount > 1 Then
                    oCMatrix.DeleteRow(IntICount)
                    oCMatrix.FlushToDataSource()
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This is used to delete the empty rows in the Shift Matrix
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ShiftDeleteEmptyRow()
        Dim oShiftCodeEdit, oShiftDescEdit As SAPbouiCOM.EditText
        Dim IntICount As Integer
        Try
            For IntICount = oSFMatrix.RowCount To 1 Step -1
                oShiftCodeEdit = oSftCodeCol.Cells.Item(IntICount).Specific
                oShiftDescEdit = oSftNameCol.Cells.Item(IntICount).Specific
                If oShiftCodeEdit.Value.Length = 0 And oShiftDescEdit.Value.Length = 0 And oSFMatrix.RowCount > 1 Then
                    oSFMatrix.DeleteRow(IntICount)
                    oSFMatrix.FlushToDataSource()
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub Form_Resize()
        Try
            If BoolResize = False Then
                oForm.Freeze(True)
                oForm.Items.Item("rectpi").Height = 282
                oForm.Items.Item("rectpi").Top = 192
                oForm.Items.Item("rectpi").Width = oForm.Width - 20
                oForm.Items.Item("rectsp").Height = 282
                oForm.Items.Item("rectsp").Top = 192
                oForm.Items.Item("rectsp").Width = oForm.Width - 20
                oForm.Items.Item("rectpp").Height = 282
                oForm.Items.Item("rectpp").Top = 192
                oForm.Items.Item("rectpp").Width = oForm.Width - 20
                oForm.Items.Item("rectci").Height = 282
                oForm.Items.Item("rectci").Top = 192
                oForm.Items.Item("rectci").Width = oForm.Width - 20
                oForm.Items.Item("rectcd").Height = 282
                oForm.Items.Item("rectcd").Top = 192
                oForm.Items.Item("rectcd").Width = oForm.Width - 20
                oForm.Items.Item("rectacc").Height = 282
                oForm.Items.Item("rectacc").Top = 192
                oForm.Items.Item("rectacc").Width = oForm.Width - 20
                oForm.Items.Item("rectsft").Height = 282
                oForm.Items.Item("rectsft").Top = 192
                oForm.Items.Item("rectsft").Width = oForm.Width - 20

                oForm.Items.Item("folpi").Top = 173
                oForm.Items.Item("folspec").Top = 173
                oForm.Items.Item("folpp").Top = 173
                oForm.Items.Item("folci").Top = 173
                oForm.Items.Item("folcd").Top = 173
                oForm.Items.Item("folacc").Top = 173
                oForm.Items.Item("folsft").Top = 173

                oForm.Items.Item("lblbp").Top = 197
                oForm.Items.Item("txtbp").Top = 197
                ' oForm.Items.Item("lnkbp").Top = 197
                oForm.Items.Item("btnbp").Top = 197

                oForm.Items.Item("lblpono").Top = 212
                oForm.Items.Item("txtpono").Top = 212
                oForm.Items.Item("lblpodt").Top = 212
                oForm.Items.Item("txtpodt").Top = 212
                oForm.Items.Item("lblwarrexp").Top = 227
                oForm.Items.Item("txtwarrexp").Top = 227

                oForm.Items.Item("lbloprcost").Top = 197
                oForm.Items.Item("txtoprcost").Top = 197
                oForm.Items.Item("lblpwrcost").Top = 212
                oForm.Items.Item("txtpwrcost").Top = 212
                oForm.Items.Item("lblsetup").Top = 227
                oForm.Items.Item("txtsetup").Top = 227
                oForm.Items.Item("lblcost1").Top = 242
                oForm.Items.Item("txtcost1").Top = 242
                oForm.Items.Item("lblcost2").Top = 257
                oForm.Items.Item("txtcost2").Top = 257

                oForm.Items.Item("lblacccode").Top = 197
                oForm.Items.Item("lnkraccod").Top = 198
                oForm.Items.Item("txtacccode").Top = 197
                oForm.Items.Item("btnacct").Top = 197
                oForm.Items.Item("lblaccdesc").Top = 212
                oForm.Items.Item("txtaccdesc").Top = 212
                oForm.Items.Item("lblsetupac").Top = 227
                oForm.Items.Item("lnksaccod").Top = 228
                oForm.Items.Item("txtsetupac").Top = 227
                oForm.Items.Item("btnsacct").Top = 227
                oForm.Items.Item("lblstacnam").Top = 242
                oForm.Items.Item("txtstaccod").Top = 242



                oForm.Items.Item("matspec").Top = 197
                oForm.Items.Item("matspec").Height = 272
                oForm.Items.Item("matspec").Width = oForm.Width - 30
                oSColumns.Item("specid").Width = 135
                oSColumns.Item("specval").Width = 135

                oForm.Items.Item("matpp").Top = 197
                oForm.Items.Item("matpp").Height = 272
                oForm.Items.Item("matpp").Width = oForm.Width - 30
                oPColumns.Item("paramid").Width = 135
                oPColumns.Item("paramname").Width = 200
                oPColumns.Item("paramval").Width = 135

                oForm.Items.Item("matsft").Top = 197
                oForm.Items.Item("matsft").Height = 272
                oForm.Items.Item("matsft").Width = oForm.Width - 30
                oSFcolumns.Item("sftcode").Width = 135
                oSFcolumns.Item("shiftname").Width = 200
                oSFcolumns.Item("durmins").Width = 200

                oForm.Items.Item("matci").Top = 197
                oForm.Items.Item("matci").Height = 272
                oForm.Items.Item("matci").Width = oForm.Width - 30
                oCColumns.Item("itemid").Width = 135
                oCColumns.Item("itemdesc").Width = 200
                oCColumns.Item("insdate").Width = 100
                oCColumns.Item("lifespan").Width = 100
                oCColumns.Item("units").Width = 100

                oForm.Freeze(False)
                oForm.Update()
                BoolResize = True
            ElseIf BoolResize = True Then
                oForm.Freeze(True)
                oForm.Items.Item("rectpi").Height = 135
                oForm.Items.Item("rectpi").Left = 5
                oForm.Items.Item("rectpi").Top = 192
                oForm.Items.Item("rectpi").Width = 590

                oForm.Items.Item("rectsp").Height = 135
                oForm.Items.Item("rectsp").Left = 5
                oForm.Items.Item("rectsp").Top = 192
                oForm.Items.Item("rectsp").Width = 590

                oForm.Items.Item("rectpp").Height = 135
                oForm.Items.Item("rectpp").Left = 5
                oForm.Items.Item("rectpp").Top = 192
                oForm.Items.Item("rectpp").Width = 590

                oForm.Items.Item("rectacc").Height = 282
                oForm.Items.Item("rectacc").Top = 192
                oForm.Items.Item("rectacc").Left = 5
                oForm.Items.Item("rectacc").Width = 590


                oForm.Items.Item("rectci").Height = 135
                oForm.Items.Item("rectci").Left = 5
                oForm.Items.Item("rectci").Top = 192
                oForm.Items.Item("rectci").Width = 590

                oForm.Items.Item("rectcd").Height = 135
                oForm.Items.Item("rectcd").Left = 5
                oForm.Items.Item("rectcd").Top = 192
                oForm.Items.Item("rectcd").Width = 590

                oForm.Items.Item("rectsft").Height = 135
                oForm.Items.Item("rectsft").Left = 5
                oForm.Items.Item("rectsft").Top = 192
                oForm.Items.Item("rectsft").Width = 590

                oForm.Items.Item("lblbp").Top = 197
                oForm.Items.Item("txtbp").Top = 197
                ' oForm.Items.Item("lnkbp").Top = 197
                oForm.Items.Item("btnbp").Top = 197

                oForm.Items.Item("lblpono").Top = 212
                oForm.Items.Item("txtpono").Top = 212
                oForm.Items.Item("lblpodt").Top = 212
                oForm.Items.Item("txtpodt").Top = 212
                oForm.Items.Item("lblwarrexp").Top = 227
                oForm.Items.Item("txtwarrexp").Top = 227

                oForm.Items.Item("lbloprcost").Top = 197
                oForm.Items.Item("txtoprcost").Top = 197
                oForm.Items.Item("lblpwrcost").Top = 212
                oForm.Items.Item("txtpwrcost").Top = 212
                oForm.Items.Item("lblsetup").Top = 227
                oForm.Items.Item("txtsetup").Top = 227
                oForm.Items.Item("lblcost1").Top = 242
                oForm.Items.Item("txtcost1").Top = 242
                oForm.Items.Item("lblcost2").Top = 257
                oForm.Items.Item("txtcost2").Top = 257

                oForm.Items.Item("lblacccode").Top = 197
                oForm.Items.Item("txtacccode").Top = 197
                oForm.Items.Item("lblaccdesc").Top = 212
                oForm.Items.Item("txtaccdesc").Top = 212
                oForm.Items.Item("lblsetupac").Top = 227
                oForm.Items.Item("txtsetupac").Top = 227
                oForm.Items.Item("lblstacnam").Top = 242
                oForm.Items.Item("txtstaccod").Top = 242

                oForm.Items.Item("matspec").Top = 197
                oForm.Items.Item("matspec").Height = 125
                oForm.Items.Item("matspec").Width = 580
                oSColumns.Item("specid").Width = 100
                oSColumns.Item("specval").Width = 100

                oForm.Items.Item("matpp").Top = 197
                oForm.Items.Item("matpp").Height = 125
                oForm.Items.Item("matpp").Width = 580
                oPColumns.Item("paramid").Width = 100
                oPColumns.Item("paramname").Width = 135
                oPColumns.Item("paramval").Width = 100

                oForm.Items.Item("matsft").Top = 197
                oForm.Items.Item("matsft").Height = 125
                oForm.Items.Item("matsft").Width = 580
                oSFcolumns.Item("sftcode").Width = 100
                oSFcolumns.Item("shiftname").Width = 135
                oSFcolumns.Item("durmins").Width = 135


                oForm.Items.Item("matci").Top = 197
                oForm.Items.Item("matci").Height = 125
                oForm.Items.Item("matci").Width = 580
                oCColumns.Item("itemid").Width = 100
                oCColumns.Item("itemdesc").Width = 135
                oCColumns.Item("insdate").Width = 75
                oCColumns.Item("lifespan").Width = 75
                oCColumns.Item("units").Width = 75

                BoolResize = False
                oForm.Freeze(False)
                oForm.Update()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub SBO_Application_FormDataEvent1(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent

    End Sub
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Dim ChooseFromListEvent As SAPbouiCOM.ChooseFromListEvent
        Try
            If (FormUID = "FM") Then
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE) Then
                    SBO_Application = Nothing
                    GC.Collect()
                End If

                If pVal.Before_Action = True And pVal.ItemUID = "matspec" And pVal.ColUID = "specid" Or pVal.ColUID = "specval" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN Then
                    If pVal.CharPressed <> 9 And pVal.CharPressed <> 36 And pVal.CharPressed <> 8 And pVal.CharPressed <> 0 Then
                        If pVal.CharPressed < 48 Or pVal.CharPressed > 57 Then
                            BubbleEvent = False
                            Exit Sub
                        End If

                    End If
                End If

                If pVal.Before_Action = True And pVal.ItemUID = "matpp" And pVal.ColUID = "paramid" Or pVal.ColUID = "paramval" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN Then
                    If pVal.CharPressed <> 9 And pVal.CharPressed <> 36 And pVal.CharPressed <> 8 And pVal.CharPressed <> 0 Then
                        If pVal.CharPressed < 48 Or pVal.CharPressed > 57 Then
                            BubbleEvent = False
                            Exit Sub
                        End If

                    End If
                End If
                '**********ChooseFromList Event is called using the raiseevent*********
                If (FormUID = "FM") And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST And pVal.BeforeAction = False Then
                    ChooseFromListEvent = pVal
                    RaiseEvent ChooseFromList(pVal.ItemUID, pVal.ColUID, pVal.Row, ChooseFromListEvent.ChooseFromListUID, ChooseFromListEvent.SelectedObjects)
                End If
                'Added by Manimaran------s
                If pVal.ItemUID = "txtbp" And pVal.BeforeAction = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then
                    CreateNewConditions(oPOCFL, "DocStatus", SAPbouiCOM.BoConditionOperation.co_EQUAL, "O", 0, 0, SAPbouiCOM.BoConditionRelationship.cr_AND)
                    CreateNewConditions(oPOCFL, "CardCode", SAPbouiCOM.BoConditionOperation.co_EQUAL, oForm.Items.Item("txtbp").Specific.string)
                End If
                'Added by Manimaran------ePOLst
                '********* Setting The PaneLevels for the Folders *********
                If (pVal.ItemUID = "folpi") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And pVal.BeforeAction = False Then
                    oForm.PaneLevel = 1
                End If
                If (pVal.ItemUID = "folspec") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And pVal.BeforeAction = False Then
                    oForm.PaneLevel = 2
                    If Not oSMatrix Is Nothing Then
                        fSettings.MatrixUID = "matspec"
                    End If
                End If
                If (pVal.ItemUID = "folpp") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And pVal.BeforeAction = False Then
                    oForm.PaneLevel = 3
                    If Not oPMatrix Is Nothing Then
                        fSettings.MatrixUID = "matpp"
                    End If
                End If
                If (pVal.ItemUID = "folci") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And pVal.BeforeAction = False Then

                    oForm.PaneLevel = 4
                    If Not oCMatrix Is Nothing Then
                        fSettings.MatrixUID = "matci"
                    End If
                End If
                If (pVal.ItemUID = "folcd") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And pVal.BeforeAction = False Then
                    oForm.PaneLevel = 5
                End If
                If (pVal.ItemUID = "folacc") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And pVal.BeforeAction = False Then
                    oForm.PaneLevel = 6
                End If
                If (pVal.ItemUID = "folsft") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And pVal.BeforeAction = False Then
                    oForm.PaneLevel = 7
                    If Not oSFMatrix Is Nothing Then
                        fSettings.MatrixUID = "matsft"
                    End If
                End If

                '****************Matrix Link Pressed**********************************
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED) And pVal.BeforeAction = False Then
                    If (pVal.ItemUID = "matsft") And pVal.BeforeAction = False Then
                        Dim oShift As String
                        Dim oShiftEdit As SAPbouiCOM.EditText
                        oShiftEdit = oSftCodeCol.Cells.Item(pVal.Row).Specific
                        oShift = oShiftEdit.Value
                        ShiftClass = New Shift(SBO_Application, oCompany, oShift, "Machine")
                    End If
                End If

                '***** Reloads the Combo's if Define New is selected and data added in the Forms *****
                If (pVal.FormTypeEx = "FM") And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE) And pVal.BeforeAction = False Then
                    ModelDFN()
                    MakeDFN()
                    MeasUnitDFN()
                End If

                '******** Model Combo Select *********
                If (pVal.ItemUID = "cmbmodel") And (pVal.BeforeAction = False) And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT) Then
                    oParentDB.SetValue("U_modecode", oParentDB.Offset, oModlCombo.Selected.Description)
                    '**** Model Combo Define New Selection *****
                    If oModlCombo.Selected.Value = "Define New" Then
                        LoadDefaultForm("PSSIT_MOD")
                        BubbleEvent = False
                        oParentDB.SetValue("U_modedesc", oParentDB.Offset, "")
                        BoolModel = False
                    End If
                End If
                '********** Make Combo Select **********
                If (pVal.ItemUID = "cmbmake") And (pVal.BeforeAction = False) And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT) Then
                    oParentDB.SetValue("U_makecode", oParentDB.Offset, oMakCombo.Selected.Description)
                    '***** Make Combo Define New Selection *****
                    If oMakCombo.Selected.Value = "Define New" Then
                        LoadDefaultForm("PSSIT_MAK")
                        BubbleEvent = False
                        oParentDB.SetValue("U_makedesc", oParentDB.Offset, "")
                        BoolMake = False
                    End If
                End If

                '******** Measurement Unit Combo Select ********
                If (pVal.ItemUID = "cmbmsrunit") And (pVal.BeforeAction = False) And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT) Then
                    oParentDB.SetValue("U_uomcode", oParentDB.Offset, oMsrUntCombo.Selected.Description)
                    '***** Measurement Unit Combo Define New Selection *****
                    If oMsrUntCombo.Selected.Value = "Define New" Then
                        LoadDefaultForm("PSSIT_OUOM")
                        BubbleEvent = False
                        oParentDB.SetValue("U_uomdesc", oParentDB.Offset, "")
                        BoolUOM = False
                    End If
                End If
                '******* Add Next Row If the Specification Id is given **********
                If (pVal.ItemUID = "matspec") And pVal.CharPressed = Keys.Tab And pVal.BeforeAction = False Then
                    If Len(oSColumns.Item("specid").Cells.Item(oSMatrix.RowCount).Specific.value) > 0 And pVal.Row = oSMatrix.RowCount Then
                        oSpecDB.InsertRecord(oSpecDB.Size)
                        oSpecDB.Offset = oSpecDB.Size - 1
                        SpecSetValue()
                        oSMatrix.AddRow(1, oSMatrix.RowCount)
                    End If
                End If
                '******* Add Next Row If the Machine Parameter Id is given **********
                If (pVal.ItemUID = "matpp") And pVal.CharPressed = Keys.Tab And pVal.BeforeAction = False Then
                    If Len(oPColumns.Item("paramid").Cells.Item(oPMatrix.RowCount).Specific.value) > 0 And pVal.Row = oPMatrix.RowCount Then
                        oParamDB.InsertRecord(oParamDB.Size)
                        oParamDB.Offset = oParamDB.Size - 1
                        ParamSetValue()
                        oPMatrix.AddRow(1, oPMatrix.RowCount)
                    End If
                End If
                ' *************Specification Id Validation **********
                If (pVal.ItemUID = "matspec") And (pVal.ColUID = "specid") And (pVal.BeforeAction = False) And ((pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Or (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE)) And pVal.CharPressed = Keys.Tab Then
                    Dim RS As SAPbobsCOM.Recordset
                    Dim oSpecID As SAPbouiCOM.EditText
                    Try
                        oSpecID = oSpecIdCol.Cells.Item(pVal.Row).Specific
                        RS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        RS.DoQuery("select distinct a.U_speccode from [@PSSIT_PMWCSPEC] a,[@PSSIT_PMWCHDR] b                        where a.code = b.code And b.code = '" & oCodeTxt.Value & "' and a.U_speccode= '" & oSpecID.Value & "'")
                        If RS.RecordCount > 0 Then
                            SBO_Application.SetStatusBarMessage("Specification ID" & oSpecID.Value & "Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                        End If
                        '********** Row Validation ***********
                        Dim IntICount As Integer
                        Dim oCurrentRow As Integer
                        Dim oSpecIDEdit As SAPbouiCOM.EditText
                        Try
                            oCurrentRow = pVal.Row
                            oSpecIDEdit = oSpecIdCol.Cells.Item(oCurrentRow).Specific
                            For IntICount = 1 To oSMatrix.RowCount - 1
                                If IntICount <> oCurrentRow Then
                                    If Len(oSpecIdCol.Cells.Item(IntICount).Specific.value) > 0 Then
                                        If UCase(oSpecIDEdit.Value) = UCase(oSpecIdCol.Cells.Item(IntICount).Specific.value) Then
                                            SBO_Application.StatusBar.SetText("Specification ID Value " & oSpecIDEdit.Value & " Already exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                        End If
                                    End If
                                End If
                            Next
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End Try
                    Catch ex As Exception
                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Finally
                        RS = Nothing
                        GC.Collect()
                    End Try
                End If

                ' *************Machine Parameter Id Validation **********
                If (pVal.ItemUID = "matpp") And (pVal.ColUID = "paramid") And (pVal.BeforeAction = False) And ((pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Or (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE)) And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) Then
                    Dim RS As SAPbobsCOM.Recordset
                    Dim oParamID As SAPbouiCOM.EditText
                    Try
                        oParamID = oParIdCol.Cells.Item(pVal.Row).Specific
                        RS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        RS.DoQuery("select distinct a.U_paracode from [@PSSIT_PMWCPARA] a,[@PSSIT_PMWCHDR] b where a.code = b.code And b.code = '" & oCodeTxt.Value & "' and a.U_paracode= '" & oParamID.Value & "'")
                        If RS.RecordCount > 0 Then
                            SBO_Application.SetStatusBarMessage("Parameter ID" & oParamID.Value & "Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                        End If
                        '********** Row Validation ***********
                        Dim IntICount As Integer
                        Dim oCurrentRow As Integer
                        Dim oParamIDEdit As SAPbouiCOM.EditText
                        Try
                            oCurrentRow = pVal.Row
                            oParamIDEdit = oParIdCol.Cells.Item(oCurrentRow).Specific
                            For IntICount = 1 To oPMatrix.RowCount - 1
                                If IntICount <> oCurrentRow Then
                                    If Len(oParIdCol.Cells.Item(IntICount).Specific.value) > 0 Then
                                        If UCase(oParamIDEdit.Value) = UCase(oParIdCol.Cells.Item(IntICount).Specific.value) Then
                                            SBO_Application.StatusBar.SetText("Parameter ID Value " & oParamIDEdit.Value & " Already exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                        End If
                                    End If
                                End If
                            Next
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End Try
                    Catch ex As Exception
                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Finally
                        RS = Nothing
                        GC.Collect()
                    End Try
                End If

                '************** Add Row in Matrix *************
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK) And pVal.BeforeAction = False Then
                    If (pVal.ItemUID = "matspec") And pVal.ColUID = "#" Then
                        SpecUID = pVal.ItemUID
                        CritUID = ""
                        ParamUID = ""
                        ShiftUID = ""
                    End If

                    If (pVal.ItemUID = "matci") And pVal.ColUID = "#" Then
                        CritUID = pVal.ItemUID
                        SpecUID = ""
                        ParamUID = ""
                        ShiftUID = ""
                    End If
                    If (pVal.ItemUID = "matpp") And pVal.ColUID = "#" Then
                        ParamUID = pVal.ItemUID
                        CritUID = ""
                        SpecUID = ""
                        ShiftUID = ""
                    End If
                    If (pVal.ItemUID = "matsft") And pVal.ColUID = "#" Then
                        ShiftUID = pVal.ItemUID
                        CritUID = ""
                        SpecUID = ""
                        ParamUID = ""
                    End If
                End If

                '******* Add Next Row If the Parameter Id is Selected **********
                If (pVal.ItemUID = "matpp") And (pVal.ColUID = "paramid") And (pVal.BeforeAction = False) And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT) Then
                    Dim oPCombo As SAPbouiCOM.ComboBox
                    Dim oPEdit As SAPbouiCOM.EditText
                    Dim oParId, oParName As String
                    Dim CurrentRow As Integer
                    oPCombo = oParIdCol.Cells.Item(pVal.Row).Specific
                    oPEdit = oParNameCol.Cells.Item(pVal.Row).Specific
                    oParId = oPCombo.Selected.Value
                    oPEdit.Value = oPCombo.Selected.Description
                    oParName = oPEdit.Value

                    CurrentRow = pVal.Row
                    If CurrentRow = oPMatrix.VisualRowCount Then
                        oParamDB.Offset = oParamDB.Size - 1
                        ParamSetValue()
                        oPMatrix.SetLineData(CurrentRow)
                        oPMatrix.FlushToDataSource()
                    End If
                    oParamDB.SetValue("U_paracode", oParamDB.Offset, oParId)
                    oParamDB.SetValue("U_paradesc", oParamDB.Offset, oParName)
                    oParamDB.SetValue("U_paraval", oParamDB.Offset, "")

                    oPMatrix.SetLineData(CurrentRow)
                    oPMatrix.FlushToDataSource()
                    If Len(oParId) > 0 Then
                        oParamDB.InsertRecord(oParamDB.Size)
                        oParamDB.Offset = oParamDB.Size - 1
                        ParamSetValue()
                        oPMatrix.AddRow(1, oPMatrix.RowCount)
                    End If
                End If
                '*********************Link Button Press************************
                If pVal.ItemUID = "lnkwccod" And pVal.BeforeAction = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    Dim oWCCode As String
                    oWCCode = oDeptNoTxt.Value
                    WorkCentreClass = New WorkCentre(SBO_Application, oCompany, oWCCode, "MachineMaster")
                End If
                If pVal.ItemUID = "lnkmgcod" And pVal.BeforeAction = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    Dim oMGCode As String
                    oMGCode = oMachineGrpCodeTxt.Value
                    MachineGroupsClass = New MachineGroups(SBO_Application, oCompany, oMGCode, "MachineMaster")
                End If

                ''************ Add Mode Item Press **********
                If (FormUID = "FM") And pVal.ItemUID = "1" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If (pVal.BeforeAction = True) Then
                        If (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            '********* Delete Empty Rows In the matrix **********
                            Try
                                SpecDeleteEmptyRow()       'Delete Empty Rows of Specification 
                                ParamDeleteEmptyRow()      'Delete Empty Rows of Production Parameters 
                                CriticalDeleteEmptyRow()   'Delete Empty Rows of Crtical Items
                                ShiftDeleteEmptyRow()
                                AccKeyCheck()
                                CommonValidation()
                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    MachValidation()
                                End If
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                            End Try
                        End If
                    End If

                    If (pVal.BeforeAction = False) Then
                        Try
                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                oForm.Items.Item("txtmcno").Enabled = False
                                oForm.Items.Item("txtmcname").Enabled = True
                            End If
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                DeleteEmptyRows()
                                RefreshForm()
                                oCodeTxt.Value = GenerateSerialNo("PSSIT_PMWCHDR")
                                oParentDB.SetValue("Code", oParentDB.Offset, GenerateSerialNo("PSSIT_PMWCHDR"))
                                StatusCombo(oStsCombo)
                                AddRowMatrix()
                                oActiveCheck.Checked = True
                                oInstlDtTxt.String = "t" 'System.DateTime.Today.Date.ToString("dd/MM/yyyy")
                                SBO_Application.SendKeys("{TAB}")
                                oPODtTxt.String = "t" 'System.DateTime.Today.Date.ToString("dd/MM/yyyy")
                                SBO_Application.SendKeys("{TAB}")
                                oWrtyExpTxt.String = "t" 'System.DateTime.Today.Date.ToString("dd/MM/yyyy")
                                SBO_Application.SendKeys("{TAB}")
                                oYrMakCombo.Select(Date.Today, SAPbouiCOM.BoSearchKey.psk_ByValue)
                                oWCNoTxt.Active = True
                            End If
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        End Try

                    End If
                End If
            End If
            '**********Realligning the values when the form is resized***********
            If pVal.FormUID = "FM" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_RESIZE And pVal.BeforeAction = False And pVal.InnerEvent = False Then
                'Form_Resize()
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Dim FormID As String
        Try
            FormID = SBO_Application.Forms.ActiveForm.UniqueID
            If pVal.MenuUID = "1281" And FormID = "FM" Then
                If pVal.BeforeAction = True Then
                    oForm.Items.Item("txtmcno").Enabled = True
                    oForm.Items.Item("txtmcname").Enabled = True
                    oWCNoTxt.Active = True
                End If
                If pVal.BeforeAction = False Then
                    oWCNoTxt.Active = True
                End If
            End If
            '*****************************Navigation*******************************
            If (pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") And FormID = "FM" And pVal.BeforeAction = False Then
                Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Try
                    oRs.DoQuery("Select * from [@PSSIT_PMWCHDR]")
                    If oRs.RecordCount > 0 Then
                        oForm.Items.Item("txtmcno").Enabled = False
                        oForm.Items.Item("txtmcname").Enabled = True
                    Else
                        oForm.Items.Item("txtmcno").Enabled = True
                        oForm.Items.Item("txtmcname").Enabled = True
                        'oWCNoTxt.Active = True
                    End If
                Catch ex As Exception
                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Finally
                    oRs = Nothing
                    GC.Collect()
                End Try
            End If
            '*************Add Mode******************
            If pVal.BeforeAction = False And FormID = "FM" Then
                If pVal.MenuUID = "1282" Then
                    oCodeTxt.Value = GenerateSerialNo("PSSIT_PMWCHDR")
                    oParentDB.SetValue("Code", oParentDB.Offset, GenerateSerialNo("PSSIT_PMWCHDR"))
                    RefreshForm()
                    oActiveCheck.Checked = True
                    oForm.Items.Item("txtmcno").Enabled = True
                    oForm.Items.Item("txtmcname").Enabled = True
                    StatusCombo(oStsCombo)
                    AddRowMatrix()
                    oInstlDtTxt.String = "t" 'System.DateTime.Today.Date.ToString("dd/MM/yyyy")
                    SBO_Application.SendKeys("{TAB}")
                    oPODtTxt.String = "t" 'System.DateTime.Today.Date.ToString("dd/MM/yyyy")
                    SBO_Application.SendKeys("{TAB}")
                    oWrtyExpTxt.String = "T" 'System.DateTime.Today.Date.ToString("dd/MM/yyyy")
                    SBO_Application.SendKeys("{TAB}")
                    oYrMakCombo.Select(Date.Today, SAPbouiCOM.BoSearchKey.psk_ByValue)
                    oWCNoTxt.Active = True
                End If
            End If

            If pVal.MenuUID = "1283" And FormID = "FM" Then
                If pVal.BeforeAction = True Then
                    Dim oStrSql As String
                    Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Try
                        oStrSql = "Select (sum(a.cnt) + Sum (b.cnt)) as ReferredCount " _
                                & "from (Select count(*) as cnt from [@PSSIT_RTE1]  Where U_Wcno = '" & oWCNoTxt.Value & "') as a, " _
                                & "(Select count(*) as cnt from [@PSSIT_PRN1]  Where U_Wcno = '" & oWCNoTxt.Value & "') as b "
                        oRs.DoQuery(oStrSql)
                        If oRs.RecordCount > 0 Then
                            oRs.MoveFirst()
                            If oRs.Fields.Item("ReferredCount").Value > 0 Then
                                SBO_Application.SetStatusBarMessage("Cannot be removed. Transactions are linked to an object, '" & oWCNoTxt.Value & "'", SAPbouiCOM.BoMessageTime.bmt_Short, True)
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
            '*************** Delete Row in Matrix ******************
            If pVal.MenuUID = "1293" And pVal.BeforeAction = True Then
                '*************Specification Matrix**************
                If SpecUID = "matspec" Then
                    Try
                        'oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        oSpecDB = oForm.DataSources.DBDataSources.Item("@PSSIT_PMWCSPEC")

                        oSMatrix.DeleteRow(oSMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                        oSMatrix.FlushToDataSource()

                        If (oSMatrix.RowCount = 0) Then
                            oSpecDB.RemoveRecord(0)
                        End If
                        BubbleEvent = False
                    Catch ex As Exception
                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    End Try
                End If
                '****************Parameter Matrix**********************
                If ParamUID = "matpp" Then
                    Try
                        'oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        oParamDB = oForm.DataSources.DBDataSources.Item("@PSSIT_PMWCPARA")
                        oPMatrix.DeleteRow(oPMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                        oPMatrix.FlushToDataSource()
                        If (oPMatrix.RowCount = 0) Then
                            oParamDB.RemoveRecord(0)
                        End If
                        BubbleEvent = False
                    Catch ex As Exception
                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    End Try
                End If
                '***************Critical Item Matrix***************
                If CritUID = "matci" Then
                    Try
                        'oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        oCriticalDB = oForm.DataSources.DBDataSources.Item("@PSSIT_PMWCITEM")

                        oCMatrix.DeleteRow(oCMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                        oCMatrix.FlushToDataSource()
                        If (oCMatrix.RowCount = 0) Then
                            oCriticalDB.RemoveRecord(0)
                        End If
                        BubbleEvent = False
                    Catch ex As Exception
                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    End Try
                End If
                '***************Shift Matrix***************
                If CritUID = "matsft" Then
                    Try
                        ' oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        oSftDB = oForm.DataSources.DBDataSources.Item("@PSSIT_PMWCSFT")

                        oSFMatrix.DeleteRow(oSFMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                        oSFMatrix.FlushToDataSource()
                        If (oSFMatrix.RowCount = 0) Then
                            oSftDB.RemoveRecord(0)
                        End If
                        BubbleEvent = False
                    Catch ex As Exception
                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    End Try
                End If
            End If
            '*************** Add Row in Matrix ******************
            If pVal.MenuUID = "1292" And pVal.BeforeAction = True Then
                '*************Specification Matrix**************
                If SpecUID = "matspec" Then
                    Try
                        'oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        If oSMatrix.RowCount > 0 Then
                            If oSMatrix.Columns.Item("specid").Cells.Item(oSMatrix.RowCount).Specific.value <> "" And oSMatrix.RowCount > 0 Then
                                oSpecDB.InsertRecord(oSpecDB.Size)
                                oSpecDB.Offset = oSpecDB.Size - 1
                                SpecSetValue()
                                oSMatrix.AddRow(1, oSMatrix.RowCount)
                            End If
                        Else
                            oSpecDB.InsertRecord(oSpecDB.Size)
                            oSpecDB.Offset = oSpecDB.Size - 1
                            SpecSetValue()
                            oSMatrix.AddRow(1, oSMatrix.RowCount)

                        End If
                        
                        'oForm.Update()
                    Catch ex As Exception
                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    End Try
                End If
                '****************Parameter Matrix**********************
                If ParamUID = "matpp" Then
                    Try
                        If oPMatrix.RowCount > 0 Then
                            If oPMatrix.Columns.Item("paramid").Cells.Item(oPMatrix.RowCount).Specific.value <> "" Then
                                oParamDB.InsertRecord(oParamDB.Size)
                                oParamDB.Offset = oParamDB.Size - 1
                                ParamSetValue()
                                oPMatrix.AddRow(1, oPMatrix.RowCount)

                            End If
                        Else
                            oParamDB.InsertRecord(oParamDB.Size)
                            oParamDB.Offset = oParamDB.Size - 1
                            ParamSetValue()
                            oPMatrix.AddRow(1, oPMatrix.RowCount)

                        End If

                        'oForm.Update()
                    Catch ex As Exception
                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    End Try
                End If
                '***************Critical Item Matrix***************
                If CritUID = "matci" Then
                    Try
                        'oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        If oCMatrix.RowCount > 0 Then

                      
                        If oCMatrix.Columns.Item("itemid").Cells.Item(oCMatrix.RowCount).Specific.value <> "" Then

                            oCriticalDB.InsertRecord(oCriticalDB.Size)
                            oCriticalDB.Offset = oCriticalDB.Size - 1
                            CriticalSetValue()
                            oCMatrix.AddRow(1, oCMatrix.RowCount)
                            SetInstalledDate(oCMatrix.RowCount)
                            End If
                        Else
                            oCriticalDB.InsertRecord(oCriticalDB.Size)
                            oCriticalDB.Offset = oCriticalDB.Size - 1
                            CriticalSetValue()
                            oCMatrix.AddRow(1, oCMatrix.RowCount)
                            SetInstalledDate(oCMatrix.RowCount)
                        End If
                        'oForm.Update()
                    Catch ex As Exception
                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    End Try
                End If
                '***************Shift Matrix***************
                If ShiftUID = "matsft" Then
                    Try
                        'oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        If oSFMatrix.RowCount > 0 Then


                            If oSFMatrix.Columns.Item("sftcode").Cells.Item(oSFMatrix.RowCount).Specific.value <> "" Then
                                oSftDB.InsertRecord(oSftDB.Size)
                                oSftDB.Offset = oSftDB.Size - 1
                                ShiftSetValue()
                                oSFMatrix.AddRow(1, oSFMatrix.RowCount)
                            End If
                        Else
                            oSftDB.InsertRecord(oSftDB.Size)
                            oSftDB.Offset = oSftDB.Size - 1
                            ShiftSetValue()
                            oSFMatrix.AddRow(1, oSFMatrix.RowCount)
                        End If
                        'oForm.Update()
                    Catch ex As Exception
                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    End Try
                End If
            End If
            'End If

        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    ''' <summary>
    ''' This is used to refresh the form
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub RefreshForm()
        Try
            Dim f As SAPbouiCOM.Form
            f = SBO_Application.Forms.Item("FM")
            f.Refresh()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub DeleteEmptyRows()
        Dim oRS, oRS1, oRS2, oRS3 As SAPbobsCOM.Recordset
        Try
            oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS3 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery("Delete from [@PSSIT_PMWCSPEC] where U_speccode is null")
            oRS1.DoQuery("Delete from [@PSSIT_PMWCPARA] where U_paracode is null")
            oRS2.DoQuery("Delete from [@PSSIT_PMWCITEM] where U_itemcode is null")
            oRS2.DoQuery("Delete from [@PSSIT_PMWCSFT] where U_SCode is null")
        Catch ex As Exception
            Throw ex
        Finally
            oRS = Nothing
            oRS1 = Nothing
            oRS2 = Nothing
            oRS3 = Nothing
            GC.Collect()
        End Try
    End Sub
    Private Sub AddRowMatrix()
        '******* Adding Empty Row to the Specification Matrix *******

        oSpecDB.InsertRecord(oSpecDB.Size)
        oSpecDB.Offset = oSpecDB.Size - 1
        SpecSetValue()
        oSMatrix.AddRow(1, oSMatrix.RowCount)

        '***** Adding Empty Row to the Production Parameter Matrix *******

        oParamDB.InsertRecord(oParamDB.Size)
        oParamDB.Offset = oParamDB.Size - 1
        ParamSetValue()
        oPMatrix.AddRow(1, oPMatrix.RowCount)

        '******** Adding Empty Row to the Critical Items Matrix *********

        oCriticalDB.InsertRecord(oCriticalDB.Size)
        oCriticalDB.Offset = oCriticalDB.Size - 1
        CriticalSetValue()
        oCMatrix.AddRow(1, oCMatrix.RowCount)
        SetInstalledDate(oCMatrix.RowCount)

        '******** Adding Empty Row to the Shift Detail Matrix *********

        oSftDB.InsertRecord(oSftDB.Size)
        oSftDB.Offset = oSftDB.Size - 1
        ShiftSetValue()
        oSFMatrix.AddRow(1, oSFMatrix.RowCount)
    End Sub
    Function String2Date(ByVal S As String, ByVal Fmt As String) As Object
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
                    oACFldr.Select()
                    oRAcctCodeTxt.Active = True
                    Throw New Exception("Select Account from the List")
                End If
                If oSAcctCodeTxt.Value.Length = 0 Then
                    oACFldr.Select()
                    oSAcctCodeTxt.Active = True
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
    ''' Validation
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CommonValidation()
        '**************** Mandatory **************
        If oWcNamTxt.Value = "" Or oWcNamTxt.Value = Nothing Then
            oWcNamTxt.Active = True
            Throw New Exception("Enter Machine Name")
        End If
        If oWCNoTxt.Value = "" Or oWCNoTxt.Value = Nothing Then
            oWCNoTxt.Active = True
            Throw New Exception("Enter Machine No")
        End If
        If oMachineGrpCodeTxt.Value = "" Or oMachineGrpCodeTxt.Value = Nothing Then
            oMachineGrpCodeTxt.Active = True
            Throw New Exception("Enter Machine Group")
        End If
        If oDeptNoTxt.Value = "" Or oDeptNoTxt.Value = Nothing Then
            oDeptCodTxt.Active = True
            Throw New Exception("Enter WorkCenter/Department")
        End If
        'If CDate(oInstlDtTxt.String) = "" Then

        'Else
        '    If Format(oInstlDtTxt, "MM/DD/YY") Then
        '        If DateDiff(DateInterval.Day, CDate(oInstlDtTxt.String), CDate(SBO_Application.Company.ServerDate)) < 0 Then
        '            oInstlDtTxt.Value = ""
        '            Throw New Exception(" Enter the Correct Installation Date")

        '        End If
        '    Else
        '        If DateDiff(DateInterval.Day, CDate(Format(oInstlDtTxt.String, "MM/DD/YY")), CDate(SBO_Application.Company.ServerDate)) < 0 Then
        '            oInstlDtTxt.Value = ""
        '            Throw New Exception(" Enter the Correct Installation Date")

        '        End If

        '    End If

        'End If
        'If oPODtTxt.String >= System.DateTime.Today.Date.ToString("DD/MM/YYYY") Then
        'If oPODtTxt.String <> "" Then
        '    If DateDiff(DateInterval.Day, CDate(oPODtTxt.String), CDate(SBO_Application.Company.ServerDate)) < 0 Then
        '        oPODtTxt.Value = ""
        '        Throw New Exception("PO date should not be greater than the current date")
        '    End If
        'End If
        If oOprCostTxt.Value <> "" Then

            If oOprCostTxt.Value < 0 Then
                oOprCostTxt.Value = ""
                Throw New Exception("Operational Cost can't be negative value")
            End If
        End If

        If oPwrCostTxt.Value <> "" Then
            If oPwrCostTxt.Value < 0 Then
                oPwrCostTxt.Value = ""
                Throw New Exception("Power Cost can't be negative value")
            End If
        End If
        If oSetupCost.Value <> "" Then
            If oSetupCost.Value < 0 Then
                oSetupCost.Value = ""
                Throw New Exception("Setup Cost can't be negative value")
            End If
        End If

        If oCost1Txt.Value <> "" Then
            If oCost1Txt.Value < 0 Then
                oCost1Txt.Value = ""
                Throw New Exception("Cost can't be negative value")
            End If
        End If

        If oCost2Txt.Value <> "" Then
            If oCost2Txt.Value < 0 Then
                oCost2Txt.Value = ""
                Throw New Exception("Cost can't be negative value")
            End If
        End If

        If oSFMatrix.RowCount = 0 Then
            oSFTFldr.Select()
            Throw New Exception("Enter Shift Details")
        End If
        'Added by Manimaran------s
        If oForm.Items.Item("txtacccode").Specific.value.length = 0 Then
            Throw New Exception("Account Code should not be empty")
        End If
        If oForm.Items.Item("txtsetupac").Specific.value.length = 0 Then
            Throw New Exception("Account Code should not be empty")
        End If
        'Added by Manimaran------e
    End Sub
    ''' <summary>
    ''' Machine Validation
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub MachValidation()
        Dim oRs, oRs1 As SAPbobsCOM.Recordset
        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            ' *************Work Center No Validation **********
            oRs.DoQuery("select U_Wcno from [@PSSIT_PMWCHDR]  where U_Wcno= '" & oWCNoTxt.Value & "' ")
            If oRs.RecordCount > 0 Then
                oWCNoTxt.Active = True
                Throw New Exception("Machine No Already Exists")
            End If

            ' *************Work Center Name Validation **********
            'oRs1.DoQuery("select U_wcname from [@PSSIT_PMWCHDR]  where U_wcname= '" & oWcNamTxt.Value & "' ")
            'If oRs1.RecordCount > 0 Then
            '    oWcNamTxt.Active = True
            '    Throw New Exception("Machine Name Already Exists")
            'End If
       
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            oRs1 = Nothing
            GC.Collect()
        End Try
    End Sub
End Class
