'''' <remarks> This class is used for entering the production transaction details.</remarks>

Public Class ProductionEntry
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
    Private UFolderDs, UConAcCode, UConAcName, URewQty, UAccMacCost, UAccToolCost, UAccLabCost, UPODocEnt, UTotFCost As SAPbouiCOM.UserDataSource
    Private oParentDB, oMachinesDB, oScrapDB, oLabourDB, oToolsDB, oFixedCostDB, oRewrkDB As SAPbouiCOM.DBDataSource
    Private PSSIT_OPEY, PSSIT_PEY1, PSSIT_PEY2, PSSIT_PEY3, PSSIT_PEY4, PSSIT_PEY5, PSSIT_PEY6 As SAPbobsCOM.UserTable
    '**************************ChooseFromList************************************
    Private oChMacList, oChPOList, oChShiftList, oChLabList As SAPbouiCOM.ChooseFromList
    Private oChPOBtnList, oChShiftBtnList As SAPbouiCOM.ChooseFromList
    Private Event ChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable)
    '**************************Items - EditText************************************
    Private oPOSeriesTxt, oPODocEntryTxt, oPONoTxt, oWORNoTxt, oPODateTxt, oShiftCodeTxt, oPENoTxt, oDocDateTxt, oShiftNameTxt, oSftFromTimeTxt, oSftToTimeTxt As SAPbouiCOM.EditText
    Private oItemCodeTxt, oItemNameTxt, oGLAccTxt, oWhsCodeTxt, oWhsNameTxt, oPlanQtyTxt, oCompQtyTxt, oRejQtyTxt, oRewQtyTxt, oScrapQtyTxt As SAPbouiCOM.EditText
    Private oRteIDTxt, oOprCodeTxt, oOprLineIDTxt, oProdQtyTxt, oPassedQtyTxt, oOprRewQtyTxt, oOprScrQtyTxt, oOprRewActCodeTxt, oOprRewActNameTxt, oOprScrActCodeTxt, oOprScrActNameTxt, oStpRea As SAPbouiCOM.EditText
    Private oInfo1Txt, oInfo2Txt, oTotMacCostTxt, oTotToolCostTxt, oTotLabCostTxt, oConAcCodeTxt, oConAcNameTxt, oJVNoTxt, oRemarksTxt, oActRewQtyTxt As SAPbouiCOM.EditText
    Private oAccProdQtyTxt, oAccPassQtyTxt, oAccRewQtyTxt, oAccScrapQtyTxt As SAPbouiCOM.EditText
    Private oAccMacCstTxt, oAccLabCstTxt, oAccToolCstTxt, oTotFCostTxt As SAPbouiCOM.EditText
    '**************************Link Button************************************
    Private oPOLink, oShiftLink, oItemLink, oWarehouselink, oJvLink As SAPbouiCOM.LinkedButton
    '**************************Items - CheckBox************************************
    Private oRewCheck, oAccKeyCheck, oClosedCheck As SAPbouiCOM.CheckBox
    '**************************Items - ComboBox************************************
    Private oPESeriesCombo, oOprCombo, oOprRewRsnCombo, oOprScrRsnCombo, cmbintime As SAPbouiCOM.ComboBox
    Private oCombo, oCombo1, oComb, oCombo2 As SAPbouiCOM.ComboBox
    '**************************Items - Button************************************
    Private oPOBtn, oShiftBtn As SAPbouiCOM.Button
    '**************************Items - Matrix************************************
    Private oMacMatrix, oLabMatrix, oToolsMatrix, oFCMatrix, oReWrkMatrix, oScrpMatrix As SAPbouiCOM.Matrix
    Private oMacColumns, oToolsColumns, oLabColumns, oFCColumns, oReWrkCol, oScrpCol As SAPbouiCOM.Columns

    '**************************Items - Machine Column************************************
    Private oMDocEntryCol, oMacCodeCol, oMacNameCol, oMWCCodeCol, oMWCNameCol, oMTypeCol, oMFromTimeCol, oMstopCol As SAPbouiCOM.Column
    Private oMToTimeCol, oMRunTimeCol, oMQtyCol, oMAccKeyCol, oMAccCodeCol, oMAccNameCol As SAPbouiCOM.Column
    Private oMConAccCodeCol, oMConAccNameCol, oMOprCstCol, oMPowCstCol, oMOthCst1Col As SAPbouiCOM.Column
    Private oMOthCst2Col, oMRunOprCstCol, oMRunPowCstCol, oMRunOthCst1Col, oMRunOthCst2Col, oMTotRunCostCol As SAPbouiCOM.Column
    Private oMInfo1Col, oMInfo2Col As SAPbouiCOM.Column
    '**************************Items - FixedCost Column************************************
    Private oFCodeCol, oFMacCodeCol, oFWrkCentreCodeCol, oFFixedCostCol, oFUnitCostCol As SAPbouiCOM.Column
    Private oFAbsMthdCol, oFActCodeCol, oFActNameCol, oFTotCostCol, oFPOSerCol, oFPONoCol, oFPENoCol As SAPbouiCOM.Column
    '**************************Items - Tools Column************************************
    Private oTCodeCol, oTPENoCol, oTMLineIDCol, oTMDocEntryCol, oTMacNoCol, oToolCodeCol As SAPbouiCOM.Column
    Private oToolDescCol, oTQtyCol, oTAccKeyCol, oTAccCodeCol, oTAccNameCol, oTConAccCodeCol As SAPbouiCOM.Column
    Private oTConAccNameCol, oToolCstCol, oTotToolCstCol, oTInfo1Col, oTInfo2Col As SAPbouiCOM.Column
    '************************** Scrap column
    Private oScrpMCodeCol, oScrpLCodeCol As SAPbouiCOM.Column
    '**************************Items - Labour Column************************************
    Private oLCodeCol, oLPENoCol, oLMLineIDCol, oLMDocEntryCol, oLMacNoCol, oLabCodeCol, oLabNameCol As SAPbouiCOM.Column
    Private oLSkGroupCodeCol, oLSkGroupCodeCol1 As SAPbouiCOM.Column
    Private oLSkGroupNameCol, oLReqNosCol, oLabKeyCol, oLParCol, oLFromTimeCol, oLToTimeCol, oLWrkTimeCol, oLotTimeCol, oNOPCol As SAPbouiCOM.Column
    Private oLQtyCol, oLAccKeyCol, oLAccCodeCol, oLAccNameCol, oLConAccCodeCol, oLConAccNameCol As SAPbouiCOM.Column
    Private oLabRateCol, oLTotRunCstCol, oLInfo1Col, oLInfo2Col As SAPbouiCOM.Column
    Private oRwrkQty, oRwrkRea, oScrpQty, oScrpRea As SAPbouiCOM.Column
    Private matcol10, matcol11, matcol12, matcol13, matcol14 As SAPbouiCOM.Column
    '**************************Folder************************************
    Private oToolsFldr, oLabFldr, oMacFldr As SAPbouiCOM.Folder
    '**************************Variables************************************
    Private oToolsSerialNo, oLabSerialNo, oFCSerialNo As Integer
    Private oToolsUID, oLabourUID, oMachineUId, oRwrkUID, oScrpUID As String
    Private oBoolResize As Boolean
    Private BoolRewDefine = True, BoolScpDefine As Boolean = True
    Private oFormName As String
    Private oMBoolFromTime = True, oMBoolToTime = True, oLBoolFromTime = True, oLBoolToTime As Boolean = True
    Private iJournal As Integer = 0
    Private WithEvents oShiftClass As Shift
    Private WithEvents oToolsClass As Tools
    Private WithEvents oLabourClass As Labour
    Private oDataTable As DataTable
    Private DataRow As DataRow
    Private TotRwrkQty, TotScrpQty As Double
    Private oShiftFromTime, oShiftToTime As String
    Dim oToTime As Integer
    Dim ofrTime As Integer
    Private Rs As SAPbobsCOM.Recordset
    Private sQry As String
    Private Typofitm As Integer
    Private oCmbRew As SAPbouiCOM.ComboBox
    Private oCmbScrp As SAPbouiCOM.ComboBox
    Private dblJournalDebit, dblJournalCredit As Double
#End Region
    ''' <summary>
    ''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmProductionEntry.srf") method is called to load the Labour form.
    ''' Drawform() method is called to Initialize the form,Datasources and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company, Optional ByVal aRouteCode As String = Nothing, Optional ByVal aFormName As String = Nothing)
        MyBase.New(aSBO_Application, aCompany)
        SBO_Application = aSBO_Application
        oCompany = aCompany
        LoadFromXML("FrmProductionEntry.srf")
        DrawForm()
        oForm.DataBrowser.BrowseBy = "txtpeyno"
        EnableMenu()
        SetItemEnabled()
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
            oParentDB = oForm.DataSources.DBDataSources.Item("@PSSIT_OPEY")
            oMachinesDB = oForm.DataSources.DBDataSources.Item("@PSSIT_PEY1")
            oLabourDB = oForm.DataSources.DBDataSources.Add("@PSSIT_PEY2")
            oToolsDB = oForm.DataSources.DBDataSources.Add("@PSSIT_PEY3")
            oFixedCostDB = oForm.DataSources.DBDataSources.Add("@PSSIT_PEY4")
            oRewrkDB = oForm.DataSources.DBDataSources.Add("@PSSIT_PEY5")
            oScrapDB = oForm.DataSources.DBDataSources.Add("@PSSIT_PEY6")
            Initialize()
            AddUserDataSources()
            InitializeFormComponent()
            LoadLookups()
            ConfigureMachineMatrix()
            ConfigureFCMatrix()
            ConfigureToolsMatrix()
            ConfigureLabourMatrix()
            'Added by Manimaran-------s
            ConfigureRWrkMatrix()
            ConfigureScrpMatrix()
            oMacFldr = oForm.Items.Item("101").Specific
            oMacFldr.Select()
            'Added by Manimaran-------e
            oForm.Freeze(False)
            oForm.Update()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    'Added by Manimaran-------s
    Private Sub ConfigureRWrkMatrix()
        Try
            oReWrkMatrix = oForm.Items.Item("104").Specific
            oReWrkMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oReWrkCol = oReWrkMatrix.Columns

            oRwrkQty = oReWrkCol.Item("V_1")
            oRwrkQty.Editable = True
            oRwrkQty.DataBind.SetBound(True, "@PSSIT_PEY5", "U_Rewrkqty")

            oRwrkRea = oReWrkCol.Item("V_0")
            oRwrkRea.Editable = True
            oRwrkRea.DataBind.SetBound(True, "@PSSIT_PEY5", "U_Rerewk")
        Catch ex As Exception
        End Try
    End Sub
    Private Sub ConfigureScrpMatrix()
        Try
            oScrpMatrix = oForm.Items.Item("105").Specific
            oScrpMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oScrpCol = oScrpMatrix.Columns

            oScrpQty = oScrpCol.Item("V_1")
            oScrpQty.Editable = True
            oScrpQty.DataBind.SetBound(True, "@PSSIT_PEY6", "U_scrapqty")

            oScrpRea = oScrpCol.Item("V_0")
            oScrpRea.Editable = True
            oScrpRea.DataBind.SetBound(True, "@PSSIT_PEY6", "U_Rescrp")

            ' added by kabilahan  -b
            oScrpMCodeCol = oScrpCol.Item("V_2")
            oScrpMCodeCol.Editable = True
            oScrpMCodeCol.DataBind.SetBound(True, "@PSSIT_PEY6", "U_MacCode")
            oScrpMCodeCol.ChooseFromListUID = "ScrpMacLst"
            oScrpMCodeCol.ChooseFromListAlias = "Code"

            oScrpLCodeCol = oScrpCol.Item("V_3")
            oScrpLCodeCol.Editable = True
            oScrpLCodeCol.DataBind.SetBound(True, "@PSSIT_PEY6", "U_LabCode")
            oScrpLCodeCol.ChooseFromListUID = "LabLst"
            oScrpLCodeCol.ChooseFromListAlias = "code"

            ' added by kabilahan - e
        Catch ex As Exception
        End Try
    End Sub
    'Added by Manimaran-------e
    ''' <summary>
    ''' Enabling the AddRow and DeleteRow Menu.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub EnableMenu()
        Try
            'oForm.EnableMenu("1292", True)
            'oForm.EnableMenu("1293", True)
            oForm.EnableMenu("1288", True)
            oForm.EnableMenu("1289", True)
            oForm.EnableMenu("1290", True)
            oForm.EnableMenu("1291", True)
            oForm.EnableMenu("1292", True)
            oForm.EnableMenu("1293", True)
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
            PSSIT_OPEY = UserTables.Item("PSSIT_OPEY")
            PSSIT_PEY1 = UserTables.Item("PSSIT_PEY1")
            PSSIT_PEY2 = UserTables.Item("PSSIT_PEY2")
            PSSIT_PEY3 = UserTables.Item("PSSIT_PEY3")
            PSSIT_PEY4 = UserTables.Item("PSSIT_PEY4")
            PSSIT_PEY5 = UserTables.Item("PSSIT_PEY5")
            PSSIT_PEY6 = UserTables.Item("PSSIT_PEY6")
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
            UFolderDs = oForm.DataSources.UserDataSources.Add("UFol", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            UConAcCode = oForm.DataSources.UserDataSources.Add("UCACCde", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 15)
            UConAcName = oForm.DataSources.UserDataSources.Add("UCACNme", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            URewQty = oForm.DataSources.UserDataSources.Add("UActRew", SAPbouiCOM.BoDataType.dt_QUANTITY, 19.6)
            UAccMacCost = oForm.DataSources.UserDataSources.Add("UAccMac", SAPbouiCOM.BoDataType.dt_PRICE, 19.6)
            UAccLabCost = oForm.DataSources.UserDataSources.Add("UActLab", SAPbouiCOM.BoDataType.dt_PRICE, 19.6)
            UAccToolCost = oForm.DataSources.UserDataSources.Add("UActTool", SAPbouiCOM.BoDataType.dt_PRICE, 19.6)
            UPODocEnt = oForm.DataSources.UserDataSources.Add("UPODocEnt", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 10)
            UTotFCost = oForm.DataSources.UserDataSources.Add("UTotFCst", SAPbouiCOM.BoDataType.dt_PRICE, 19.6)
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
                    'oToolsFldr.Select()
                ElseIf IntICount = 2 Then
                    oLabFldr = oForm.Items.Item("follabour").Specific
                    oForm.Items.Item("follabour").AffectsFormMode = False
                    oLabFldr.DataBind.SetBound(True, "", "UFol")
                    oLabFldr.GroupWith("foltools")
                End If
            Next


            '**************************Header Data******************************************
            oPOSeriesTxt = oForm.Items.Item("txtseris").Specific
            oForm.Items.Item("txtseris").Enabled = False
            oPOSeriesTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Pnordser")

            oPONoTxt = oForm.Items.Item("txtprdno").Specific
            oForm.Items.Item("txtprdno").Enabled = True
            'oForm.Items.Item("txtprdno").LinkTo = "lnkpo"
            oPONoTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Pnordno")

            oPODocEntryTxt = oForm.Items.Item("txtpodcent").Specific
            oForm.Items.Item("txtpodcent").Enabled = True
            oForm.Items.Item("txtpodcent").Visible = True
            oForm.Items.Item("txtpodcent").LinkTo = "lnkpo"
            oPODocEntryTxt.DataBind.SetBound(True, "", "UPODocEnt")
            oForm.Items.Add("lnkpo", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnkpo").Visible = True
            oForm.Items.Item("lnkpo").LinkTo = "txtpodcent"
            oForm.Items.Item("lnkpo").Height = 12
            oForm.Items.Item("lnkpo").Width = 9
            oForm.Items.Item("lnkpo").Top = 6
            oForm.Items.Item("lnkpo").Left = 116
            oForm.Items.Item("lnkpo").Description = "Link to" & vbNewLine & "Production Order"
            oPOLink = oForm.Items.Item("lnkpo").Specific
            oPOLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_ProductionOrder


            oWORNoTxt = oForm.Items.Item("txtworno").Specific
            oForm.Items.Item("txtworno").Enabled = False
            oForm.Items.Item("lblworno").Visible = False
            oForm.Items.Item("txtworno").Visible = False
            oWORNoTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_WORNo")

            oPOBtn = oForm.Items.Item("btnpo").Specific
            oForm.Items.Item("btnpo").Description = "Choose from List" & vbNewLine & "Production Order List View"
            oPOBtn.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            oPOBtn.Image = sPath & "\Resources\CFL.bmp"

            oPODateTxt = oForm.Items.Item("txtpordt").Specific
            oForm.Items.Item("txtpordt").Enabled = False
            oPODateTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Pordt")

            oPESeriesCombo = oForm.Items.Item("cmbseris1").Specific
            oForm.Items.Item("cmbseris1").DisplayDesc = True
            oPESeriesCombo.DataBind.SetBound(True, "@PSSIT_OPEY", "Series")
            oPESeriesCombo.ValidValues.LoadSeries(oForm.BusinessObject.Type, SAPbouiCOM.BoSeriesMode.sf_Add)
            ' oPESeriesCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            Dim pa As Integer
            For pa = oPESeriesCombo.ValidValues.Count - 1 To 0 Step -1
                oPESeriesCombo.Select(pa, SAPbouiCOM.BoSearchKey.psk_Index)
            Next

            oPENoTxt = oForm.Items.Item("txtpeyno").Specific
            oForm.Items.Item("txtpeyno").Enabled = False
            oPENoTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "DocNum")
            With oForm.DataSources.DBDataSources.Item("@PSSIT_OPEY")
                .SetValue("DocNum", .Offset, oForm.BusinessObject.GetNextSerialNumber(Trim(.GetValue("Series", .Offset))).ToString)
            End With

            oDocDateTxt = oForm.Items.Item("txtdocdt").Specific
            oForm.Items.Item("txtdocdt").Enabled = True
            oDocDateTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Docdt")
            oDocDateTxt.String = "t" 'System.DateTime.Today.Date.ToString("dd/MM/yyyy")
            SBO_Application.SendKeys("{TAB}")

            oShiftCodeTxt = oForm.Items.Item("txtscode").Specific
            oForm.Items.Item("txtscode").Enabled = False
            oForm.Items.Item("txtscode").LinkTo = "lnksft"
            oShiftCodeTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Scode")
            oForm.Items.Add("lnksft", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnksft").Visible = True
            oForm.Items.Item("lnksft").LinkTo = "txtscode"
            oForm.Items.Item("lnksft").Height = 12
            oForm.Items.Item("lnksft").Width = 9
            oForm.Items.Item("lnksft").Top = 36
            oForm.Items.Item("lnksft").Left = 116
            oForm.Items.Item("lnksft").Description = "Link to" & vbNewLine & "Shift"
            oShiftLink = oForm.Items.Item("lnksft").Specific

            oShiftBtn = oForm.Items.Item("btnscode").Specific
            oForm.Items.Item("btnscode").Enabled = False
            oForm.Items.Item("btnscode").Description = "Choose from List" & vbNewLine & "Shift List View"
            oShiftBtn.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            oShiftBtn.Image = sPath & "\Resources\CFL.bmp"

            oShiftNameTxt = oForm.Items.Item("txtsdesc").Specific
            oForm.Items.Item("txtsdesc").Enabled = False
            oShiftNameTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Sdesc")

            oSftFromTimeTxt = oForm.Items.Item("txtsfrtime").Specific
            oForm.Items.Item("txtsfrtime").Enabled = False
            'oForm.Items.Item("txtsfrtime").Visible = False

            oSftToTimeTxt = oForm.Items.Item("txtstotime").Specific
            oForm.Items.Item("txtstotime").Enabled = False
            'oForm.Items.Item("txtstotime").Visible = False

            oItemCodeTxt = oForm.Items.Item("txtitmcd").Specific
            oForm.Items.Item("txtitmcd").Enabled = False
            oForm.Items.Item("txtitmcd").LinkTo = "lnkitem"
            oItemCodeTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Itemcode")
            oForm.Items.Add("lnkitem", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnkitem").Visible = True
            oForm.Items.Item("lnkitem").LinkTo = "txtitmcd"
            oForm.Items.Item("lnkitem").Height = 12
            oForm.Items.Item("lnkitem").Width = 9
            oForm.Items.Item("lnkitem").Top = 71
            oForm.Items.Item("lnkitem").Left = 116
            oForm.Items.Item("lnkitem").Description = "Link to" & vbNewLine & "Items"
            oItemLink = oForm.Items.Item("lnkitem").Specific
            oItemLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Items

            oGLAccTxt = oForm.Items.Item("txtglmthd").Specific
            oForm.Items.Item("txtglmthd").Enabled = False
            oGLAccTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_GLMethod")

            oItemNameTxt = oForm.Items.Item("txtitnam").Specific
            oForm.Items.Item("txtitnam").Enabled = False
            oItemNameTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_ItemName")

            oWhsCodeTxt = oForm.Items.Item("txtwhcod").Specific
            oForm.Items.Item("txtwhcod").Enabled = False
            oForm.Items.Item("txtwhcod").LinkTo = "lnkwhse"
            oWhsCodeTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Whscode")
            oForm.Items.Add("lnkwhse", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnkwhse").Visible = True
            oForm.Items.Item("lnkwhse").LinkTo = "txtwhcod"
            oForm.Items.Item("lnkwhse").Height = 12
            oForm.Items.Item("lnkwhse").Width = 9
            oForm.Items.Item("lnkwhse").Top = 101
            oForm.Items.Item("lnkwhse").Left = 116
            oForm.Items.Item("lnkwhse").Description = "Link to" & vbNewLine & "Items"
            oWarehouselink = oForm.Items.Item("lnkwhse").Specific
            oWarehouselink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Warehouses

            oWhsNameTxt = oForm.Items.Item("txtwhnam").Specific
            oForm.Items.Item("txtwhnam").Enabled = False
            oWhsNameTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Whsname")

            oPlanQtyTxt = oForm.Items.Item("txtplqty").Specific
            oForm.Items.Item("txtplqty").Enabled = False
            oPlanQtyTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Planqty")

            oCompQtyTxt = oForm.Items.Item("txtcmqty").Specific
            oForm.Items.Item("txtcmqty").Enabled = False
            oCompQtyTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Comqty")

            oRejQtyTxt = oForm.Items.Item("txtrjqty").Specific
            oForm.Items.Item("txtrjqty").Enabled = False
            oRejQtyTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_RejQty")

            'oRewQtyTxt = oForm.Items.Item("txtrwqty").Specific
            'oForm.Items.Item("txtrwqty").Enabled = False
            'oRewQtyTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Rewqty")

            'oScrapQtyTxt = oForm.Items.Item("txtspqty").Specific
            'oForm.Items.Item("txtspqty").Enabled = False
            'oScrapQtyTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Scpqty")

            oRewQtyTxt = oForm.Items.Item("107").Specific
            oForm.Items.Item("txtrwqty").Enabled = False
            oRewQtyTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Rewqty")

            oScrapQtyTxt = oForm.Items.Item("109").Specific
            oForm.Items.Item("txtspqty").Enabled = False
            oScrapQtyTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Scpqty")

            oOprCombo = oForm.Items.Item("cmbopcd").Specific
            oForm.Items.Item("cmbopcd").Enabled = False
            oOprCombo.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Oplnid")
            oForm.Items.Item("cmbopcd").DisplayDesc = True

            oOprCodeTxt = oForm.Items.Item("txtopcd").Specific
            oForm.Items.Item("txtopcd").Enabled = False
            oForm.Items.Item("txtopcd").Visible = False
            oOprCodeTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Oprcode")

            oRteIDTxt = oForm.Items.Item("txtrteid").Specific
            oForm.Items.Item("txtrteid").Enabled = False
            oForm.Items.Item("lblrteid").Visible = False
            oForm.Items.Item("txtrteid").Visible = False
            oRteIDTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_RteId")

            oOprLineIDTxt = oForm.Items.Item("txtopln").Specific
            oForm.Items.Item("txtopln").Enabled = False
            oForm.Items.Item("txtopln").Visible = False
            oOprLineIDTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Oprname")

            oRewCheck = oForm.Items.Item("chkrew").Specific
            'oForm.Items.Item("chkrew").Enabled = True
            oForm.Items.Item("chkrew").Visible = False
            oRewCheck.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Rework")
            oRewCheck.Checked = False

            oAccKeyCheck = oForm.Items.Item("chkackey").Specific
            oForm.Items.Item("chkackey").Enabled = True
            oForm.Items.Item("chkackey").Visible = False
            oAccKeyCheck.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Acckey")
            AccKeyCheck()

            oProdQtyTxt = oForm.Items.Item("txtpdqty").Specific
            oForm.Items.Item("txtpdqty").Enabled = True
            oProdQtyTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_ProdQty")

            oPassedQtyTxt = oForm.Items.Item("txtpsqty").Specific
            oForm.Items.Item("txtpsqty").Enabled = True
            oPassedQtyTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Passqty")
            'Commented by Manimaran-------s
            'oOprRewQtyTxt = oForm.Items.Item("txtoprwqty").Specific
            'oForm.Items.Item("txtoprwqty").Enabled = True
            'oOprRewQtyTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Rewrkqty")

            'oOprRewRsnCombo = oForm.Items.Item("cmboprwres").Specific
            'oForm.Items.Item("cmboprwres").Enabled = True
            'oForm.Items.Item("cmboprwres").DisplayDesc = True
            'oOprRewRsnCombo.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Rerewk")
            'LoadReasonCombo(oOprRewRsnCombo)

            'oOprScrQtyTxt = oForm.Items.Item("txtopscqty").Specific
            'oForm.Items.Item("txtopscqty").Enabled = True
            'oOprScrQtyTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_scrapqty")

            'oOprScrRsnCombo = oForm.Items.Item("cmbopscres").Specific
            'oForm.Items.Item("cmbopscres").Enabled = True
            'oForm.Items.Item("cmbopscres").DisplayDesc = True
            'oOprScrRsnCombo.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Rescrp")
            'LoadReasonCombo(oOprScrRsnCombo)
            'Commented by Manimaran-------e
            oOprRewActCodeTxt = oForm.Items.Item("txtracod").Specific
            oForm.Items.Item("txtracod").Enabled = False
            oForm.Items.Item("lblracod").Visible = False
            oForm.Items.Item("txtracod").Visible = False
            oOprRewActCodeTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Raccode")

            oOprRewActNameTxt = oForm.Items.Item("txtracnm").Specific
            oForm.Items.Item("txtracnm").Enabled = False
            oForm.Items.Item("txtracnm").Visible = False
            oOprRewActNameTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Racname")

            oOprScrActCodeTxt = oForm.Items.Item("txtsacod").Specific
            oForm.Items.Item("txtsacod").Enabled = False
            oForm.Items.Item("lblsacod").Visible = False
            oForm.Items.Item("txtsacod").Visible = False
            oOprScrActCodeTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Saccode")

            oOprScrActNameTxt = oForm.Items.Item("txtsacnm").Specific
            oForm.Items.Item("txtsacnm").Enabled = False
            oForm.Items.Item("txtsacnm").Visible = False
            oOprScrActNameTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Sacname")

            oInfo1Txt = oForm.Items.Item("txtadnl1").Specific
            oForm.Items.Item("txtadnl1").Enabled = True
            oInfo1Txt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Adnl1")

            oInfo2Txt = oForm.Items.Item("txtadnl2").Specific
            oForm.Items.Item("txtadnl2").Enabled = True
            oInfo2Txt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Adnl2")

            oTotMacCostTxt = oForm.Items.Item("txtmcst").Specific
            oForm.Items.Item("txtmcst").Enabled = False
            oForm.Items.Item("lbltmcst").Visible = False
            oForm.Items.Item("txtmcst").Visible = False
            oTotMacCostTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Totmcst")

            oTotToolCostTxt = oForm.Items.Item("txtttcst").Specific
            oForm.Items.Item("txtttcst").Enabled = False
            oForm.Items.Item("lblttcst").Visible = False
            oForm.Items.Item("txtttcst").Visible = False
            oTotToolCostTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Tottcst")

            oTotLabCostTxt = oForm.Items.Item("txttlcod").Specific
            oForm.Items.Item("txttlcod").Enabled = False
            oForm.Items.Item("lbltlcod").Visible = False
            oForm.Items.Item("txttlcod").Visible = False
            oTotLabCostTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Totlcst")

            oConAcCodeTxt = oForm.Items.Item("txtconaccd").Specific
            oForm.Items.Item("txtconaccd").Enabled = False
            oForm.Items.Item("txtconaccd").Visible = False
            oConAcCodeTxt.DataBind.SetBound(True, "", "UCACCde")

            oConAcNameTxt = oForm.Items.Item("txtconacnm").Specific
            oForm.Items.Item("txtconacnm").Enabled = False
            oForm.Items.Item("lblconact").Visible = False
            oForm.Items.Item("txtconacnm").Visible = False
            oConAcNameTxt.DataBind.SetBound(True, "", "UCACNme")

            oJVNoTxt = oForm.Items.Item("txtjvno").Specific
            oForm.Items.Item("txtjvno").LinkTo = "lnkjvno"
            oForm.Items.Item("txtjvno").Enabled = False
            oJVNoTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Jvno")
            oJVNoTxt.Value = "0"
            oForm.Items.Add("lnkjvno", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oForm.Items.Item("lnkjvno").Visible = True
            oForm.Items.Item("lnkjvno").LinkTo = "txtjvno"
            oForm.Items.Item("lnkjvno").Height = 12
            oForm.Items.Item("lnkjvno").Width = 9
            oForm.Items.Item("lnkjvno").Top = 360
            oForm.Items.Item("lnkjvno").Left = 118
            oForm.Items.Item("lnkjvno").Description = "Link to" & vbNewLine & "Journal Posting"
            oJvLink = oForm.Items.Item("lnkjvno").Specific
            oJvLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_JournalPosting


            oRemarksTxt = oForm.Items.Item("txtremar").Specific
            oForm.Items.Item("txtremar").Enabled = True
            oRemarksTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Remarks")

            oActRewQtyTxt = oForm.Items.Item("txtRewqty").Specific
            oForm.Items.Item("txtRewqty").Enabled = False
            oForm.Items.Item("txtRewqty").Visible = False
            oActRewQtyTxt.DataBind.SetBound(True, "", "UActRew")

            oClosedCheck = oForm.Items.Item("chkclosed").Specific
            oForm.Items.Item("chkclosed").Enabled = True
            oForm.Items.Item("chkclosed").Visible = False
            oClosedCheck.DataBind.SetBound(True, "@PSSIT_OPEY", "U_Closekey")

            oAccProdQtyTxt = oForm.Items.Item("txtaccprod").Specific
            oForm.Items.Item("txtaccprod").Enabled = False
            oForm.Items.Item("txtaccprod").Visible = False
            oAccProdQtyTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_AccProdQty")

            oAccPassQtyTxt = oForm.Items.Item("txtaccpass").Specific
            oForm.Items.Item("txtaccpass").Enabled = False
            oForm.Items.Item("txtaccpass").Visible = False
            oAccPassQtyTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_AccPassQty")

            oAccRewQtyTxt = oForm.Items.Item("txtaccrew").Specific
            oForm.Items.Item("txtaccrew").Enabled = False
            oForm.Items.Item("txtaccrew").Visible = False
            oAccRewQtyTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_AccRewQty")

            oAccScrapQtyTxt = oForm.Items.Item("txtaccscrp").Specific
            oForm.Items.Item("txtaccscrp").Enabled = False
            oForm.Items.Item("txtaccscrp").Visible = False
            oAccScrapQtyTxt.DataBind.SetBound(True, "@PSSIT_OPEY", "U_AccScrapQty")

            oAccMacCstTxt = oForm.Items.Item("txtacmccst").Specific
            oForm.Items.Item("txtacmccst").Enabled = False
            oForm.Items.Item("txtacmccst").Visible = False
            oAccMacCstTxt.DataBind.SetBound(True, "", "UAccMac")

            oAccLabCstTxt = oForm.Items.Item("txtaclbcst").Specific
            oForm.Items.Item("txtaclbcst").Enabled = False
            oForm.Items.Item("txtaclbcst").Visible = False
            oAccLabCstTxt.DataBind.SetBound(True, "", "UActLab")

            oAccToolCstTxt = oForm.Items.Item("txtactlcst").Specific
            oForm.Items.Item("txtactlcst").Enabled = False
            oForm.Items.Item("txtactlcst").Visible = False
            oAccToolCstTxt.DataBind.SetBound(True, "", "UActTool")

            oTotFCostTxt = oForm.Items.Item("txttotfcst").Specific
            oForm.Items.Item("txttotfcst").Enabled = False
            oForm.Items.Item("txttotfcst").Visible = False
            oTotFCostTxt.DataBind.SetBound(True, "", "UTotFCst")

            'added by Manimaran------s

            cmbintime = oForm.Items.Item("111").Specific
            oForm.Items.Item("111").Enabled = True
            cmbintime.DataBind.SetBound(True, "@PSSIT_OPEY", "U_InTime")

            oStpRea = oForm.Items.Item("113").Specific
            oForm.Items.Item("113").Enabled = True
            oStpRea.DataBind.SetBound(True, "@PSSIT_OPEY", "U_StpRea")
            'added by Manimaran------e
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
            oChPOBtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "202", "POBtnLst"))
            SetPOCFLConditions()
            oPOBtn.ChooseFromListUID = "POBtnLst"


            oChPOList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "202", "POTxtLst"))
            SetPOCFLConditions()
            oPONoTxt.ChooseFromListUID = "POTxtLst"
            oPONoTxt.ChooseFromListAlias = "DocNum"

            oChShiftBtnList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_SFT", "SftBtnLst"))
            CreateNewConditions(oChShiftBtnList, "U_Active", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            oShiftBtn.ChooseFromListUID = "SftBtnLst"

            oChShiftList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_SFT", "SftLst"))
            CreateNewConditions(oChShiftList, "U_Active", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
            oShiftCodeTxt.ChooseFromListUID = "SftLst"
            oShiftCodeTxt.ChooseFromListAlias = "Code"

            oChMacList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_WCHDR", "ScrpMacLst"))
            CreateNewConditions(oChMacList, "U_Active", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")

            oChLabList = CreateNewChoosefromList(CreateNewChooseFromListParams(False, "PSSIT_LBR", "LabLst"))
            CreateNewConditions(oChLabList, "U_Active", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Y")
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
                If (oCFL.UniqueID.Equals("POBtnLst") Or oCFL.UniqueID.Equals("POTxtLst")) Then
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
    ''' Configuring the Machine Matrix items/controls in the form(.srf) by bounding to the object and to the DBDatasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ConfigureMachineMatrix()
        Try
            oMacMatrix = oForm.Items.Item("matmac").Specific
            oMacMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oMacColumns = oMacMatrix.Columns

            'oMacRowNumCol = oColumns.Item("#")
            'oMacRowNumCol.Editable = False

            oMacCodeCol = oMacColumns.Item("colwcno")
            oMacCodeCol.Editable = True
            oMacCodeCol.DataBind.SetBound(True, "@PSSIT_PEY1", "U_wcno")

            oMacNameCol = oMacColumns.Item("colwcnam")
            oMacNameCol.Editable = False
            oMacNameCol.DataBind.SetBound(True, "@PSSIT_PEY1", "U_wcname")

            oMWCCodeCol = oMacColumns.Item("colwrkno")
            oMWCCodeCol.Editable = False
            oMWCCodeCol.Visible = False
            oMWCCodeCol.DataBind.SetBound(True, "@PSSIT_PEY1", "U_Wrkno")

            oMWCNameCol = oMacColumns.Item("colwrknm")
            oMWCNameCol.Editable = False
            oMWCNameCol.Visible = False
            oMWCNameCol.DataBind.SetBound(True, "@PSSIT_PEY1", "U_Wrkname")

            oMTypeCol = oMacColumns.Item("coltype")
            oMTypeCol.Editable = True
            LoadMachineTypeCombo()
            oMTypeCol.DisplayDesc = False
            oMTypeCol.DataBind.SetBound(True, "@PSSIT_PEY1", "U_type")

            'Added by Manimaran----s
            oMstopCol = oMacColumns.Item("colstoptim")
            oMstopCol.Editable = True
            oMstopCol.DataBind.SetBound(True, "@PSSIT_PEY1", "U_sotime")
            'Added by Manimaran----e

            oMFromTimeCol = oMacColumns.Item("colfrtim")
            oMFromTimeCol.Editable = True
            oMFromTimeCol.DataBind.SetBound(True, "@PSSIT_PEY1", "U_Frtime")

            oMToTimeCol = oMacColumns.Item("coltotim")
            oMToTimeCol.Editable = True
            oMToTimeCol.DataBind.SetBound(True, "@PSSIT_PEY1", "U_Totime")

            oMRunTimeCol = oMacColumns.Item("colrntim")
            oMRunTimeCol.Editable = False
            oMRunTimeCol.DataBind.SetBound(True, "@PSSIT_PEY1", "U_Rntime")

            oMQtyCol = oMacColumns.Item("colqty")
            oMQtyCol.Visible = False
            oMQtyCol.Editable = True
            oMQtyCol.DataBind.SetBound(True, "@PSSIT_PEY1", "U_Qty")

            oMAccKeyCol = oMacColumns.Item("colackey")
            oMAccKeyCol.Editable = False
            oMAccKeyCol.Visible = False
            oMAccKeyCol.DataBind.SetBound(True, "@PSSIT_PEY1", "U_Acckey")

            oMAccCodeCol = oMacColumns.Item("colacode")
            oMAccCodeCol.Editable = False
            oMAccCodeCol.Visible = False
            oMAccCodeCol.DataBind.SetBound(True, "@PSSIT_PEY1", "U_Accode")

            oMAccNameCol = oMacColumns.Item("colacnam")
            oMAccNameCol.Editable = False
            oMAccNameCol.Visible = False
            oMAccNameCol.DataBind.SetBound(True, "@PSSIT_PEY1", "U_Acname")

            oMConAccCodeCol = oMacColumns.Item("colcaccd")
            oMConAccCodeCol.Editable = False
            oMConAccCodeCol.Visible = False
            oMConAccCodeCol.DataBind.SetBound(True, "@PSSIT_PEY1", "U_CAccode")

            oMConAccNameCol = oMacColumns.Item("colcacnm")
            oMConAccNameCol.Editable = False
            'oMConAccNameCol.Visible = False
            oMConAccNameCol.DataBind.SetBound(True, "@PSSIT_PEY1", "U_CAcname")

            oMOprCstCol = oMacColumns.Item("colmocph")
            oMOprCstCol.Editable = False
            ' oMOprCstCol.Visible = False
            oMOprCstCol.DataBind.SetBound(True, "@PSSIT_PEY1", "U_Mopcph")

            oMPowCstCol = oMacColumns.Item("colmpcph")
            oMPowCstCol.Editable = False
            '  oMPowCstCol.Visible = False
            oMPowCstCol.DataBind.SetBound(True, "@PSSIT_PEY1", "U_Mprcph")

            oMOthCst1Col = oMacColumns.Item("colmoch1")
            oMOthCst1Col.Editable = False
            '  oMOthCst1Col.Visible = False
            oMOthCst1Col.DataBind.SetBound(True, "@PSSIT_PEY1", "U_Mohcph1")

            oMOthCst2Col = oMacColumns.Item("colmoch2")
            oMOthCst2Col.Editable = False
            ' oMOthCst2Col.Visible = False
            oMOthCst2Col.DataBind.SetBound(True, "@PSSIT_PEY1", "U_Mohcph2")

            oMRunOprCstCol = oMacColumns.Item("colrmcph")
            oMRunOprCstCol.Editable = False
            '  oMRunOprCstCol.Visible = False
            oMRunOprCstCol.DataBind.SetBound(True, "@PSSIT_PEY1", "U_RMopcph")

            oMRunPowCstCol = oMacColumns.Item("colrpcph")
            oMRunPowCstCol.Editable = False
            ' oMRunPowCstCol.Visible = False
            oMRunPowCstCol.DataBind.SetBound(True, "@PSSIT_PEY1", "U_RMprcph")

            oMRunOthCst1Col = oMacColumns.Item("colrmch1")
            oMRunOthCst1Col.Editable = False
            ' oMRunOthCst1Col.Visible = False
            oMRunOthCst1Col.DataBind.SetBound(True, "@PSSIT_PEY1", "U_RMohcph1")

            oMRunOthCst2Col = oMacColumns.Item("colrmch2")
            oMRunOthCst2Col.Editable = False
            '  oMRunOthCst2Col.Visible = False
            oMRunOthCst2Col.DataBind.SetBound(True, "@PSSIT_PEY1", "U_RMohcph2")

            oMTotRunCostCol = oMacColumns.Item("coltotct")
            oMTotRunCostCol.Editable = False
            ' oMTotRunCostCol.Visible = False
            oMTotRunCostCol.DataBind.SetBound(True, "@PSSIT_PEY1", "U_Totcost")

            oMInfo1Col = oMacColumns.Item("coladnl1")
            oMInfo1Col.Editable = False
            oMInfo1Col.Visible = False
            oMInfo1Col.Visible = True
            oMInfo1Col.DataBind.SetBound(True, "@PSSIT_PEY1", "U_Adnl1")

            oMInfo2Col = oMacColumns.Item("coladnl2")
            oMInfo2Col.Editable = False
            oMInfo2Col.Visible = False
            oMInfo2Col.Visible = True
            oMInfo2Col.DataBind.SetBound(True, "@PSSIT_PEY1", "U_Adnl2")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function getdatetime(ByVal dateString As String) As Date
        Dim objBridge As SAPbobsCOM.SBObob
        objBridge = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Return objBridge.Format_StringToDate(dateString).Fields.Item(0).Value

    End Function
#Region "Validate Machine entry for same shift"
    Private Function ValidateMachineTime(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Dim strPONo, strPEDate, strShiftId, strOperID, strOperName, strmachine, strFromTime, strToTime As String
        Dim strSQL, strdocEntry As String
        Dim oTempRs As SAPbobsCOM.Recordset
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim dtDocdate As Date
        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
            Return True
        End If
        strPONo = oPODocEntryTxt.String
        strPEDate = oDocDateTxt.String
        dtDocdate = getdatetime(strPEDate)
        strShiftId = oShiftCodeTxt.String
        strOperID = oOprCodeTxt.String
        strOperName = oOprCombo.Selected.Description
        strSQL = "SELECT T0.[DocEntry],T0.[U_Pnordno], T0.[U_Scode], T0.[U_Pordt], T0.[U_Oprcode], T0.[U_Oprname] FROM [dbo].[@PSSIT_OPEY]  T0"
        strSQL = strSQL & " Where convert(nvarchar(10),U_Docdt,103)='" & dtDocdate.ToString("dd/MM/yyyy") & "' and  U_PnordNo=" & strPONo & " and U_Scode='" & strShiftId & " ' and U_OprCode='" & strOperID & "'"
        oTempRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRs.DoQuery(strSQL)
        If oTempRs.RecordCount > 0 Then
            strdocEntry = oTempRs.Fields.Item(0).Value
        Else
            strdocEntry = ""
            Return True
        End If
        Dim strMacFrm, strMacTo As String
        If strdocEntry <> "" Then
            strSQL = "SELECT T0.[U_wcno], T0.[U_wcname], T0.[U_Frtime], T0.[U_Totime], T0.[U_Rntime] FROM [dbo].[@PSSIT_PEY1]  T0 "
            strSQL = strSQL & " WHERE T0.[DocEntry] =" & strdocEntry
            oMatrix = oForm.Items.Item("matmac").Specific
            For intRow As Integer = 1 To oMatrix.RowCount
                strmachine = oMacCodeCol.Cells.Item(intRow).Specific.value
                strMacFrm = oMatrix.Columns.Item("colfrtim").Cells.Item(intRow).Specific.value
                strMacTo = oMatrix.Columns.Item("coltotim").Cells.Item(intRow).Specific.value
                strSQL = strSQL & " and U_wcno='" & strmachine & "'"
                oTempRs.DoQuery(strSQL)
                If oTempRs.RecordCount > 0 Then
                    strFromTime = oTempRs.Fields.Item(2).Value
                    strToTime = oTempRs.Fields.Item(3).Value
                    If CInt(strMacFrm) >= CInt(strFromTime) And CInt(strMacFrm) <= CInt(strToTime) Then
                        Return False
                    ElseIf CInt(strMacTo) >= CInt(strFromTime) And CInt(strMacTo) <= CInt(strToTime) Then
                        Return False
                    End If

                End If
            Next

        End If





        Return True
    End Function
#End Region

#Region "Validat the Operation Sequence"
    Private Function ValiateOperationSequence(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim strSQL1, strSQl2, strOprId, strPONo, strPOSeries, strparentid As String
        Dim oTemprs, oTemprs1 As SAPbobsCOM.Recordset
        strOprId = oOprCombo.Selected.Value
        strPONo = aForm.Items.Item("txtprdno").Specific.value
        oTemprs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemprs.DoQuery("SELECT U_OprCode,U_Parid,U_Seqnce,U_Pordser,U_POrdNo FROM [dbo].[@PSSIT_WOR2]  where U_Rework<>'Y' and U_PordNo=" & strPONo & " and U_Seqnce='" & strOprId & "'")
        If oTemprs.RecordCount > 0 Then
            strparentid = oTemprs.Fields.Item(1).Value
        Else
            strparentid = ""
        End If
        If strparentid = "0" Then
            Return True
        End If
        If strparentid <> "" Then
            oTemprs.DoQuery("SELECT max(U_Parid) FROM [dbo].[@PSSIT_WOR2]  where U_Rework<>'Y' and  U_Parid <" & strparentid & " and U_PordNo=" & strPONo) '& " and U_Seqnce=" & strOprId)
            If oTemprs.RecordCount > 0 Then
                strparentid = oTemprs.Fields.Item(0).Value
            Else
                strparentid = ""
            End If
        End If

        If strparentid <> "" Then
            oTemprs.DoQuery("SELECT U_OprCode,U_Parid FROM [dbo].[@PSSIT_WOR2]  where  U_Rework<>'Y' and  U_Parid=" & strparentid & " and U_PordNo=" & strPONo) ' & " and U_Seqnce=" & strOprId)
            If oTemprs.RecordCount > 0 Then
                strparentid = oTemprs.Fields.Item(0).Value
            Else
                strparentid = ""
            End If
        End If

        If strparentid <> "" Then
            strSQL1 = "SELECT T0.[DocEntry], T0.[U_Pnordno], T0.[U_Pnordser], T0.[U_Oprcode], T0.[U_Oprname], T0.[U_ProdQty], T0.[U_Passqty] FROM [dbo].[@PSSIT_OPEY]  T0"
            strSQL1 = strSQL1 & " where T0.[U_PnordNo]=" & strPONo & " and U_OprCode='" & strparentid & "'"
            oTemprs.DoQuery(strSQL1)
            If oTemprs.RecordCount > 0 Then
                Return True
            Else
                Return False
            End If
        Else
            Return True
        End If
        Return True

    End Function
#End Region


    Private Sub ConfigureFCMatrix()
        Try
            oFCMatrix = oForm.Items.Item("matfxdcst").Specific
            oFCMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oFCColumns = oFCMatrix.Columns

            'oFCRowNumCol = oColumns.Item("#")
            'oFCRowNumCol.Editable = False 

            oFCodeCol = oFCColumns.Item("colcode")
            oFCodeCol.Editable = False
            oFCodeCol.DataBind.SetBound(True, "@PSSIT_PEY4", "Code")

            oFPOSerCol = oFCColumns.Item("colposer")
            oFPOSerCol.Editable = False
            oFPOSerCol.DataBind.SetBound(True, "@PSSIT_PEY4", "U_Pordser")

            oFPONoCol = oFCColumns.Item("colpono")
            oFPONoCol.Editable = False
            oFPONoCol.DataBind.SetBound(True, "@PSSIT_PEY4", "U_Pordno")

            oFPENoCol = oFCColumns.Item("colpeno")
            oFPENoCol.Editable = False
            oFPENoCol.DataBind.SetBound(True, "@PSSIT_PEY4", "U_Prdentno")

            oFMacCodeCol = oFCColumns.Item("colwcno")
            oFMacCodeCol.Editable = False
            oFMacCodeCol.DataBind.SetBound(True, "@PSSIT_PEY4", "U_wcno")

            oFWrkCentreCodeCol = oFCColumns.Item("colwrkno")
            oFWrkCentreCodeCol.Editable = False
            oFWrkCentreCodeCol.DataBind.SetBound(True, "@PSSIT_PEY4", "U_wrkno")

            oFFixedCostCol = oFCColumns.Item("colfxdcst")
            oFFixedCostCol.Editable = False
            oFFixedCostCol.DataBind.SetBound(True, "@PSSIT_PEY4", "U_Fcost")

            oFUnitCostCol = oFCColumns.Item("colunitcst")
            oFUnitCostCol.Editable = False
            oFUnitCostCol.DataBind.SetBound(True, "@PSSIT_PEY4", "U_UnitCost")

            oFAbsMthdCol = oFCColumns.Item("colabsmthd")
            oFAbsMthdCol.Editable = False
            oFAbsMthdCol.DataBind.SetBound(True, "@PSSIT_PEY4", "U_Absmthd")

            oFActCodeCol = oFCColumns.Item("colaccode")
            oFActCodeCol.Editable = False
            oFActCodeCol.DataBind.SetBound(True, "@PSSIT_PEY4", "U_Accode")

            oFActNameCol = oFCColumns.Item("colacname")
            oFActNameCol.Editable = False
            oFActNameCol.DataBind.SetBound(True, "@PSSIT_PEY4", "U_Acname")

            oFTotCostCol = oFCColumns.Item("coltotfc")
            oFTotCostCol.Editable = False
            oFTotCostCol.DataBind.SetBound(True, "@PSSIT_PEY4", "U_Totfcst")

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
            oTCodeCol.DataBind.SetBound(True, "@PSSIT_PEY3", "Code")

            oTPENoCol = oToolsColumns.Item("colpeyno")
            oTPENoCol.Editable = False
            oTPENoCol.Visible = False
            oTPENoCol.DataBind.SetBound(True, "@PSSIT_PEY3", "U_Prdentno")

            oTMLineIDCol = oToolsColumns.Item("colmlid")
            oTMLineIDCol.Editable = False
            oTMLineIDCol.Visible = False
            oTMLineIDCol.DataBind.SetBound(True, "@PSSIT_PEY3", "U_Maclid")

            oTMDocEntryCol = oToolsColumns.Item("colmdey")
            oTMDocEntryCol.Editable = False
            oTMDocEntryCol.Visible = False
            oTMDocEntryCol.DataBind.SetBound(True, "@PSSIT_PEY3", "U_Madcey")

            oTMacNoCol = oToolsColumns.Item("colwcno")
            oTMacNoCol.Editable = False
            oTMacNoCol.DataBind.SetBound(True, "@PSSIT_PEY3", "U_wcno")

            oToolCodeCol = oToolsColumns.Item("coltlcode")
            oToolCodeCol.Editable = False
            oToolCodeCol.DataBind.SetBound(True, "@PSSIT_PEY3", "U_Toolcode")

            oToolDescCol = oToolsColumns.Item("coltlnam")
            oToolDescCol.Editable = False
            oToolDescCol.DataBind.SetBound(True, "@PSSIT_PEY3", "U_TLname")

            oTQtyCol = oToolsColumns.Item("colqty")
            oTQtyCol.Editable = True
            oTQtyCol.DataBind.SetBound(True, "@PSSIT_PEY3", "U_Qty")

            oTAccKeyCol = oToolsColumns.Item("colackey")
            oTAccKeyCol.Editable = False
            oTAccKeyCol.Visible = False
            oTAccKeyCol.DataBind.SetBound(True, "@PSSIT_PEY3", "U_Acckey")

            oTAccCodeCol = oToolsColumns.Item("colaccod")
            oTAccCodeCol.Editable = False
            oTAccCodeCol.Visible = False
            oTAccCodeCol.DataBind.SetBound(True, "@PSSIT_PEY3", "U_Accode")

            oTAccNameCol = oToolsColumns.Item("colacnam")
            oTAccNameCol.Editable = False
            oTAccNameCol.Visible = False
            oTAccNameCol.DataBind.SetBound(True, "@PSSIT_PEY3", "U_Acname")

            oTConAccCodeCol = oToolsColumns.Item("colcacod")
            oTConAccCodeCol.Editable = False
            oTConAccCodeCol.Visible = False
            oTConAccCodeCol.DataBind.SetBound(True, "@PSSIT_PEY3", "U_CAccode")

            oTConAccNameCol = oToolsColumns.Item("colcanam")
            oTConAccNameCol.Editable = False
            oTConAccNameCol.Visible = False
            oTConAccNameCol.DataBind.SetBound(True, "@PSSIT_PEY3", "U_CAcname")

            oToolCstCol = oToolsColumns.Item("coltlcst")
            oToolCstCol.Editable = False
            oToolCstCol.Visible = False
            oToolCstCol.DataBind.SetBound(True, "@PSSIT_PEY3", "U_Tlctppie")

            oTotToolCstCol = oToolsColumns.Item("colttcst")
            oTotToolCstCol.Editable = False
            oTotToolCstCol.Visible = False
            oTotToolCstCol.DataBind.SetBound(True, "@PSSIT_PEY3", "U_Totcost")

            oTInfo1Col = oToolsColumns.Item("coladnl1")
            oTInfo1Col.Editable = False
            oTInfo1Col.Visible = False
            oTInfo1Col.Visible = True
            oTInfo1Col.DataBind.SetBound(True, "@PSSIT_PEY3", "U_Adnl1")

            oTInfo2Col = oToolsColumns.Item("coladnl2")
            oTInfo2Col.Editable = False
            oTInfo2Col.Visible = False
            oTInfo2Col.Visible = True
            oTInfo2Col.DataBind.SetBound(True, "@PSSIT_PEY3", "U_Adnl2")
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
            oLabMatrix = oForm.Items.Item("matlab").Specific
            oLabMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oLabColumns = oLabMatrix.Columns

            'oLabRowNumCol = oColumns.Item("#")
            'oLabRowNumCol.Editable = False

            oLCodeCol = oLabColumns.Item("colcode")
            oLCodeCol.Editable = False
            oLCodeCol.Visible = False
            oLCodeCol.DataBind.SetBound(True, "@PSSIT_PEY2", "Code")

            oLPENoCol = oLabColumns.Item("colpeyno")
            oLPENoCol.Editable = False
            oLPENoCol.Visible = False
            oLPENoCol.DataBind.SetBound(True, "@PSSIT_PEY2", "U_Prdentno")

            oLMLineIDCol = oLabColumns.Item("colmlid")
            oLMLineIDCol.Editable = False
            oLMLineIDCol.Visible = False
            oLMLineIDCol.DataBind.SetBound(True, "@PSSIT_PEY2", "U_Maclid")

            oLMDocEntryCol = oLabColumns.Item("colmdery")
            oLMDocEntryCol.Editable = False
            oLMDocEntryCol.Visible = False
            oLMDocEntryCol.DataBind.SetBound(True, "@PSSIT_PEY2", "U_Madcey")

            oLMacNoCol = oLabColumns.Item("colwcno")
            oLMacNoCol.Editable = False
            oLMacNoCol.DataBind.SetBound(True, "@PSSIT_PEY2", "U_wcno")

            oLabCodeCol = oLabColumns.Item("collrcod")
            oLabCodeCol.Editable = True
            oLabCodeCol.DataBind.SetBound(True, "@PSSIT_PEY2", "U_Lrcode")
            oLabCodeCol.ChooseFromListUID = "LabLst"
            oLabCodeCol.ChooseFromListAlias = "Code"
            'Added by Manimaran------s
            oLabCodeCol.Visible = False
            'Added by Manimaran------e

            oLSkGroupCodeCol = oLabColumns.Item("collgcod")
            oLSkGroupCodeCol.Editable = True
            oLSkGroupCodeCol.Visible = True
            oLSkGroupCodeCol.DataBind.SetBound(True, "@PSSIT_PEY2", "U_LGCode")


            oLSkGroupCodeCol1 = oLabColumns.Item("colname")
            oLSkGroupCodeCol1.Editable = True
            oLSkGroupCodeCol1.Visible = True
            oLSkGroupCodeCol1.DataBind.SetBound(True, "@PSSIT_PEY2", "U_Labname")

            oLSkGroupNameCol = oLabColumns.Item("collgnam")
            oLSkGroupNameCol.Editable = False
            oLSkGroupNameCol.Visible = False
            oLSkGroupNameCol.DataBind.SetBound(True, "@PSSIT_PEY2", "U_LGname")

            oLReqNosCol = oLabColumns.Item("colreqno")
            oLReqNosCol.Editable = False
            oLReqNosCol.Visible = False
            oLReqNosCol.DataBind.SetBound(True, "@PSSIT_PEY2", "U_Reqno")

            oLabKeyCol = oLabColumns.Item("collbkey")
            oLabKeyCol.Editable = False
            oLabKeyCol.Visible = False
            oLabKeyCol.DataBind.SetBound(True, "@PSSIT_PEY2", "U_Labkey")

            oLParCol = oLabColumns.Item("colparll")
            oLParCol.Editable = True
            oLParCol.DataBind.SetBound(True, "@PSSIT_PEY2", "U_Parallel")

            oLFromTimeCol = oLabColumns.Item("colfrtim")
            oLFromTimeCol.Editable = True
            oLFromTimeCol.DataBind.SetBound(True, "@PSSIT_PEY2", "U_Frtime")

            oLToTimeCol = oLabColumns.Item("coltotim")
            oLToTimeCol.Editable = True
            oLToTimeCol.DataBind.SetBound(True, "@PSSIT_PEY2", "U_Totime")

            'Added by Manimaran------s
            oLotTimeCol = oLabColumns.Item("colottime")
            oLotTimeCol.Editable = True
            oLotTimeCol.DataBind.SetBound(True, "@PSSIT_PEY2", "U_OTtime")

            oNOPCol = oLabColumns.Item("colnop")
            oNOPCol.Editable = True
            oNOPCol.DataBind.SetBound(True, "@PSSIT_PEY2", "U_Nop")

            'Added by manimaran------e

            oLWrkTimeCol = oLabColumns.Item("colwktim")
            oLWrkTimeCol.Editable = False
            oLWrkTimeCol.DataBind.SetBound(True, "@PSSIT_PEY2", "U_Wrktime")

            oLQtyCol = oLabColumns.Item("colqty")
            oLQtyCol.Editable = True
            oLQtyCol.DataBind.SetBound(True, "@PSSIT_PEY2", "U_Qty")

            oLAccKeyCol = oLabColumns.Item("colackey")
            oLAccKeyCol.Editable = False
            oLAccKeyCol.Visible = False
            oLAccKeyCol.DataBind.SetBound(True, "@PSSIT_PEY2", "U_Acckey")

            oLAccCodeCol = oLabColumns.Item("colacode")
            oLAccCodeCol.Editable = False
            'oLAccCodeCol.Visible = False
            oLAccCodeCol.DataBind.SetBound(True, "@PSSIT_PEY2", "U_Accode")

            oLAccNameCol = oLabColumns.Item("colacnam")
            oLAccNameCol.Editable = False
            'oLAccNameCol.Visible = False
            oLAccNameCol.DataBind.SetBound(True, "@PSSIT_PEY2", "U_Acname")

            oLConAccCodeCol = oLabColumns.Item("colcacod")
            oLConAccCodeCol.Editable = False
            'oLConAccCodeCol.Visible = False
            oLConAccCodeCol.DataBind.SetBound(True, "@PSSIT_PEY2", "U_CAccode")

            oLConAccNameCol = oLabColumns.Item("colcanam")
            oLConAccNameCol.Editable = False
            'oLConAccNameCol.Visible = False
            oLConAccNameCol.DataBind.SetBound(True, "@PSSIT_PEY2", "U_CAcname")

            oLabRateCol = oLabColumns.Item("collhrph")
            oLabRateCol.Editable = False
            oLabRateCol.Visible = False
            oLabRateCol.DataBind.SetBound(True, "@PSSIT_PEY2", "U_Lrtph")

            oLTotRunCstCol = oLabColumns.Item("coltotct")
            oLTotRunCstCol.Editable = False
            oLTotRunCstCol.Visible = False
            oLTotRunCstCol.DataBind.SetBound(True, "@PSSIT_PEY2", "U_Totcost")

            oLInfo1Col = oLabColumns.Item("coladnl1")
            oLInfo1Col.Editable = True
            oLInfo1Col.Visible = False
            oLInfo1Col.DataBind.SetBound(True, "@PSSIT_PEY2", "U_Adnl1")

            oLInfo2Col = oLabColumns.Item("coladnl2")
            oLInfo2Col.Editable = True
            oLInfo2Col.Visible = False
            oLInfo2Col.DataBind.SetBound(True, "@PSSIT_PEY2", "U_Adnl2")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent
        If BusinessObjectInfo.FormUID = "FPE" And BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
            If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD Then
                LoadProdOrderDocEntryNo(oPONoTxt.Value)
                LoadToolsDataFromDB()
                LoadLabourDataFromDB()
                SetItemEnabled()

            End If
        End If
    End Sub
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Dim ChooseFromListEvent As SAPbouiCOM.ChooseFromListEvent
        Try
            If pVal.FormUID = "FPE" Then
                '*****************************ChooseFromList Event is called using the raiseevent*******************************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST And pVal.BeforeAction = False Then
                    ChooseFromListEvent = pVal
                    RaiseEvent ChooseFromList(pVal.ItemUID, pVal.ColUID, pVal.Row, ChooseFromListEvent.ChooseFromListUID, ChooseFromListEvent.SelectedObjects)
                End If



                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED Then
                    If pVal.BeforeAction = False Then
                        If pVal.ItemUID = "mattool" And pVal.ColUID = "coltlcode" And pVal.Row > 0 Then
                            Dim oToolCodeEdit As SAPbouiCOM.EditText
                            Dim oCurrentRow As Integer
                            Try
                                oCurrentRow = pVal.Row
                                oToolsMatrix.GetLineData(pVal.Row)
                                oToolCodeEdit = oToolCodeCol.Cells.Item(oCurrentRow).Specific
                                oToolsClass = New Tools(SBO_Application, oCompany, oToolCodeEdit.Value, "Production Entry")
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End Try
                        End If
                        If pVal.ItemUID = "matlab" And pVal.ColUID = "collrcod" And pVal.Row > 0 Then
                            Dim oLabCodeEdit As SAPbouiCOM.EditText
                            Dim oCurrentRow As Integer
                            Try
                                oCurrentRow = pVal.Row
                                oToolsMatrix.GetLineData(pVal.Row)
                                oLabCodeEdit = oLabCodeCol.Cells.Item(oCurrentRow).Specific
                                oLabourClass = New Labour(SBO_Application, oCompany, oLabCodeEdit.Value, "Production Entry")
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End Try
                        End If
                    End If
                End If
                'Added by manimaran--------s
                If pVal.ItemUID = "matlab" And pVal.ColUID = "collgcod" And pVal.Before_Action = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT Then
                    Dim oLSkGroupCod As SAPbouiCOM.ComboBox
                    oLSkGroupCod = oLSkGroupCodeCol.Cells.Item(oLabMatrix.RowCount).Specific
                    Rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    sQry = "select a.U_LGname ,a.U_Labrate ,a.U_Accode ,a.U_Acname,B.U_Empnam  from [@PSSIT_OLGP] a INNER JOIN [@PSSIT_olbr] B ON A.Code = B.U_LGCode where a.Code = '" & oLSkGroupCod.Selected.Value & "'"
                    Rs.DoQuery(sQry)
                    If Rs.RecordCount > 0 Then
                        oLabMatrix.GetLineData(oLabMatrix.RowCount)
                        oLabourDB.SetValue("U_LGname", oLabourDB.Offset, Rs.Fields.Item("U_LGname").Value)
                        oLabourDB.SetValue("U_Lrtph", oLabourDB.Offset, Rs.Fields.Item("U_Labrate").Value)
                        oLabourDB.SetValue("U_Accode", oLabourDB.Offset, Rs.Fields.Item("U_Accode").Value)
                        oLabourDB.SetValue("U_Acname", oLabourDB.Offset, Rs.Fields.Item("U_Acname").Value)
                        oLabMatrix.SetLineData(oLabMatrix.RowCount)
                    End If
                    Rs = Nothing
                    '************
                    Try
                        oLabMatrix = oForm.Items.Item("matlab").Specific
                        oLabColumns = oLabMatrix.Columns

                        matcol11 = oLabColumns.Item("colname")
                        Dim oCombo5 As SAPbouiCOM.Column = matcol11
                        Rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Rs.DoQuery("select B.U_Empnam  from [@PSSIT_OLGP] a INNER JOIN [@PSSIT_olbr] B ON A.Code = B.U_LGCode where a.Code = '" & oLSkGroupCod.Selected.Value & "'")
                        'Rs.MoveNext()
                        If oCombo5.ValidValues.Count > 0 Then
                            For i As Int16 = oCombo5.ValidValues.Count - 1 To 0 Step -1
                                oCombo5.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                            Next
                        End If
                        For i As Int16 = 0 To Rs.RecordCount - 1
                            oCombo5.ValidValues.Add(Rs.Fields.Item(0).Value, Rs.Fields.Item(0).Value)
                            Rs.MoveNext()
                        Next
                        Rs = Nothing
                    Catch ex As Exception
                    End Try
                    
                    '************

                    'Rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    'sQry = "select B.U_Empnam  from [@PSSIT_OLGP] a INNER JOIN [@PSSIT_olbr] B ON A.Code = B.U_LGCode where a.Code = '" & oLSkGroupCod.Selected.Value & "'"
                    'Rs.DoQuery(sQry)
                    'If Rs.RecordCount > 0 Then
                    '    While Not Rs.EoF
                    '        ' oLabMatrix.GetLineData(oLabMatrix.RowCount)
                    '        oLabourDB.SetValue("U_labname", oLabourDB.Offset, Rs.Fields.Item("U_Empnam").Value)
                    '        'oLabMatrix.SetLineData(oLabMatrix.RowCount)
                    '        Rs.MoveNext()
                    '    End While
                    'End If
                    'Rs = Nothing
                End If
                'Added by manimaran--------e
                '******************************** time validation [kabilahan]*************************
                If (pVal.ColUID = "colfrtim" Or pVal.ColUID = "coltotim") And pVal.ItemUID = "matmac" And pVal.CharPressed = Keys.Tab And pVal.BeforeAction = False Then
                    Dim StpInMins As String
                    Try
                        Dim oFromTime, oToTime As DateTime
                        Dim oIntDt, oStDt As Date
                        Dim oIntTime, oStTime As DateTime
                        Dim oStpTime As SAPbouiCOM.EditText
                        Dim dtDate As Date
                        dtDate = getdatetime(oDocDateTxt.String)
                        oIntDt = dtDate ' Convert.ToDateTime(Date.Parse(oDocDateTxt.String)) ' String2Date(oIntDtTxt.String, "DD/MM/YY")
                        'If shiftTimeValidation(pVal) = True Then
                        If pVal.ColUID = "coltotim" = True Then
                            If validateTime(oIntDt, pVal) = True Then
                                oIntTime = CDate(oMacMatrix.Columns.Item("colfrtim").Cells.Item(oMacMatrix.RowCount - 1).Specific.string)
                                oFromTime = New Date(oIntDt.Year, oIntDt.Month, oIntDt.Day, oIntTime.Hour, oIntTime.Minute, oIntTime.Second)
                                oStDt = Convert.ToDateTime(Date.Parse(oDocDateTxt.String)) 'String2Date(oCmpltDtTxt.String, "DD/MM/YY") 
                                oStTime = CDate(oMacMatrix.Columns.Item("coltotim").Cells.Item(oMacMatrix.RowCount - 1).Specific.string)
                                oToTime = New Date(oStDt.Year, oStDt.Month, oStDt.Day, oStTime.Hour, oStTime.Minute, oStTime.Second)
                                StpInMins = DateDiff(DateInterval.Minute, oFromTime, oToTime)
                                oStpTime = oMacMatrix.Columns.Item("colrntim").Cells.Item(oMacMatrix.RowCount - 1).Specific

                                Try
                                    oStpTime.String = StpInMins
                                Catch ex As Exception
                                End Try
                            Else

                                Exit Sub
                            End If

                        End If
                        'End If
                    Catch ex As Exception
                        Throw ex
                    End Try
                End If

                '*************************************************************************************


                '**********************Adding the child data to the database table********************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.BeforeAction = True Then
                        Try
                            'Added by Manimaran-------s
                            oForm = SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                            If (oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.Before_Action = True Then
                                oForm.Items.Item("104").Enabled = False
                                oForm.Items.Item("105").Enabled = False
                            End If
                            'Added by Manimaran-------e
                            '**********************Adding the child data to the database table********************
                            If pVal.ItemUID = "1" Then
                                AccKeyCheck()
                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    oForm.Freeze(True)
                                    If oMacMatrix.RowCount > 0 Then
                                        oMacMatrix.SelectRow(1, True, False)
                                    End If
                                    LoadFCDataFromDB()
                                    LoadToolsDataFromDB()
                                    LoadLabourDataFromDB()
                                    If oPONoTxt.Value.Length > 0 Then

                                        '   LoadProdOrderDocEntryNo(oPONoTxt.Value)
                                    End If
                                    'If checkProductionOrderstatus(oPONoTxt.Value) = True Then
                                    '    disable()
                                    'Else
                                    '    SetItemEnabled()
                                    'End If
                                    SetItemEnabled()
                                    oForm.Freeze(False)
                                End If
                                Dim oTTransaction, oLTransaction, oJTransaction As Boolean
                                Dim IntICount As Integer
                                Dim oMacAcCode, oMacAcName, oMacConAcCode, oMacTotRunCost As SAPbouiCOM.EditText
                                Dim oCAcCode As String = ""
                                Try
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then ' Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        '************************Deleting Empty Row in MachineMatrix,ToolsMatrix,Labour Matrix*******************
                                        MachinesDeleteEmptyRow()
                                        ToolsDeleteEmptyRow()
                                        LabourDeleteEmptyRow()
                                        'Added by Manimaran-------s
                                        If oReWrkMatrix.RowCount > 1 Then
                                            ReworkDeleteEmptyRow()
                                        End If
                                        If oScrpMatrix.RowCount > 1 Then
                                            ScrapDeleteEmptyRow()
                                        End If
                                        'Added by Manimaran-------e
                                        Validation()
                                        If ValiateOperationSequence(oForm) = False Then
                                            SBO_Application.SetStatusBarMessage("The previous operation sequence is not entered", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If

                                        If ValidateMachineTime(oForm) = False Then
                                            SBO_Application.SetStatusBarMessage("Machine details already entered for the same date and shift", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                        Try
                                            If Not oCompany.InTransaction Then
                                                oCompany.StartTransaction()
                                            End If
                                            '****************************Journal DataTable**********************************
                                            DataTableFieldCreation()
                                            dblJournalCredit = 0
                                            dblJournalDebit = 0
                                            For IntICount = 1 To oMacMatrix.RowCount
                                                oMacAcCode = oMAccCodeCol.Cells.Item(IntICount).Specific
                                                oMacAcName = oMAccNameCol.Cells.Item(IntICount).Specific
                                                oMacConAcCode = oMConAccCodeCol.Cells.Item(IntICount).Specific
                                                oMacTotRunCost = oMTotRunCostCol.Cells.Item(IntICount).Specific
                                                oMacMatrix.GetLineData(IntICount)
                                                DataRow = oDataTable.NewRow()
                                                DataRow.Item("AcCode") = CStr(oMacAcCode.Value)
                                                DataRow.Item("CAcCode") = CStr(UConAcCode.Value)
                                                DataRow.Item("CAcName") = CStr(UConAcName.Value)
                                                DataRow.Item("Debit") = 0
                                                DataRow.Item("Credit") = CDbl(oMacTotRunCost.Value)
                                                DataRow.Item("ShortName") = CStr(oMacAcCode.Value)
                                                dblJournalCredit = dblJournalCredit + CDbl(oMacTotRunCost.Value)
                                                dblJournalDebit = dblJournalDebit + 0

                                                oCAcCode = CStr(oMacAcCode.Value)
                                                oDataTable.Rows.Add(DataRow)
                                            Next
                                            AddFixedCostDatatoDB()
                                            '****************************Tool Details********************************
                                            If oToolsMatrix.RowCount > 0 Then
                                                oTTransaction = True
                                                AddToolDatatoDB()
                                            Else
                                                oTTransaction = True
                                            End If
                                            '****************************Labour Details********************************
                                            If oLabMatrix.RowCount > 0 Then
                                                oLTransaction = True
                                                AddLabourDatatoDB()
                                            Else
                                                oLTransaction = True
                                            End If
                                            '****************************Journal Posting******************************
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then ' Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                                'Commented by Manimaran----s  Sicame Client no need this 
                                                'If oAccKeyCheck.Checked = True Then
                                                '    oJTransaction = JournalPosting()
                                                'ElseIf oAccKeyCheck.Checked = False Then
                                                '    oJTransaction = True
                                                'End If
                                                oJTransaction = True
                                                'Commented by Manimaran-----e
                                            End If
                                            '*************************************************************************************
                                            If oTTransaction = True And oLTransaction = True And oJTransaction = True Then
                                                '************************Updating [@PSSIT_WOR2] based on the Production Order No*******************
                                                UpdateProductionOrder()
                                                '**************************************************************************************************

                                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)

                                            End If
                                        Catch ex As Exception
                                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            BubbleEvent = False
                                        Finally
                                            If oTTransaction = False And oLTransaction = False And oJTransaction = False Then
                                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                BubbleEvent = False
                                            ElseIf oTTransaction = True And oLTransaction = True And oJTransaction = False Then
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
                        Finally
                            iJournal = 0
                            GC.Collect()
                        End Try
                    End If
                    If pVal.BeforeAction = False Then
                        Try
                            If pVal.ItemUID = "lnksft" Then
                                oShiftClass = New Shift(SBO_Application, oCompany, oShiftCodeTxt.Value, "Production Entry")
                            End If
                            If pVal.ItemUID = "1" Then
                                '**********************Setting the Item Enabled********************
                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oForm.Freeze(True)
                                    SetItemEnabled()
                                    If oMacMatrix.RowCount > 0 Then
                                        oMacMatrix.SelectRow(1, True, False)
                                    End If
                                    If oMacMatrix.RowCount > 0 Then
                                        oMacMatrix.SelectRow(1, True, False)
                                    End If
                                    oForm.Freeze(False)
                                End If
                                '**********************Refreshing the form to initiate default values********************
                                If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    oForm.Refresh()
                                    oForm.Freeze(True)
                                    SetItemEnabled()
                                    AccKeyCheck()
                                    UPODocEnt.Value = ""
                                    UTotFCost.Value = 0
                                    oPONoTxt.Active = True
                                    oPESeriesCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                    With oForm.DataSources.DBDataSources.Item("@PSSIT_OPEY")
                                        .SetValue("DocNum", .Offset, oForm.BusinessObject.GetNextSerialNumber(Trim(.GetValue("Series", .Offset))).ToString)
                                    End With
                                    oDocDateTxt.String = "t" 'System.DateTime.Today.Date.ToString("dd/MM/yyyy")
                                    SBO_Application.SendKeys("{TAB}")
                                    oForm.Freeze(False)
                                End If
                            End If
                            '**********Setting the Form panelevel as per the Folder selected********************
                            oForm.Freeze(True)
                            If pVal.ItemUID = "foltools" Then
                                oForm.PaneLevel = 1
                            End If
                            If pVal.ItemUID = "follabour" Then
                                oForm.PaneLevel = 2
                            End If
                            'Added by Manimaran-------s
                            If pVal.ItemUID = "102" Then
                                'If oReWrkMatrix.RowCount <= 0 Then
                                '    oReWrkMatrix.AddRow(1, oReWrkMatrix.RowCount)
                                'End If

                                oForm.PaneLevel = 4
                            End If
                            If pVal.ItemUID = "103" Then
                                'If oScrpMatrix.RowCount <= 0 Then
                                '    oScrpMatrix.AddRow(1, oScrpMatrix.RowCount)
                                'End If

                                oForm.PaneLevel = 5
                            End If
                            If pVal.ItemUID = "101" Then
                                oForm.PaneLevel = 3
                            End If
                            'Added by Manimaran-------e
                            oForm.Update()
                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                            End If
                            oForm.Freeze(False)
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        End Try
                    End If
                End If
                '**********************Loading Default form for the Rework and Scrap Reason*******************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT Then
                    If pVal.BeforeAction = True Then
                        If Not oShiftCodeTxt Is Nothing Then
                            Try
                                If oParentDB.GetValue("U_Pnordno", oParentDB.Offset).Trim().Length > 0 Then
                                    If oShiftCodeTxt.Value.Length = 0 Then
                                        SBO_Application.SetStatusBarMessage("Enter Shift Code and Continue...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        BubbleEvent = False
                                    End If
                                End If
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                            End Try
                        End If
                    End If
                    If pVal.BeforeAction = False Then
                        'Commented and modified by Manimaran------s
                        'If pVal.ItemUID = "cmboprwres" Then
                        '    Try
                        '        If oOprRewRsnCombo.Selected.Value = "Define New" Then
                        '            LoadDefaultForm("PSSIT_RES")
                        '            BubbleEvent = False
                        '            oParentDB.SetValue("U_Rerewk", oParentDB.Offset, "")
                        '            BoolRewDefine = False
                        '        End If
                        '    Catch ex As Exception
                        '        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        '    End Try
                        'End If
                        'If pVal.ItemUID = "cmbopscres" Then
                        '    Try
                        '        If oOprScrRsnCombo.Selected.Value = "Define New" Then
                        '            LoadDefaultForm("PSSIT_RES")
                        '            BubbleEvent = False
                        '            oParentDB.SetValue("U_Rescrp", oParentDB.Offset, "")
                        '            BoolScpDefine = False
                        '        End If
                        '    Catch ex As Exception
                        '        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        '    End Try
                        'End If


                        If pVal.ItemUID = "104" And pVal.ColUID = "V_0" Then
                            Try

                                oCmbRew = oRwrkRea.Cells.Item(pVal.Row).Specific

                                If oCmbRew.Selected.Value = "Define New" Then
                                    LoadDefaultForm("PSSIT_RES")
                                    BubbleEvent = False
                                    oRewrkDB.SetValue("U_Rerewk", oRewrkDB.Offset, "")
                                    BoolRewDefine = False
                                End If
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End Try
                        End If

                        If pVal.ItemUID = "105" And pVal.ColUID = "V_0" Then
                            Try

                                oCmbScrp = oScrpRea.Cells.Item(oScrpMatrix.RowCount).Specific
                                If oCmbScrp.Selected.Value = "Define New" Then

                                    LoadDefaultForm("PSSIT_RES")
                                    'LoadDefaultForm("Sampling Level Master")
                                    BubbleEvent = False
                                    oScrapDB.SetValue("U_Rescrp", oScrapDB.Offset, "")
                                    BoolScpDefine = False
                                End If
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End Try
                        End If
                        'Commented and modified by Manimaran------s
                        '**********************Loading the machine details based on the operation********************
                        If pVal.ItemUID = "cmbopcd" Then

                            Dim oStrSql As String
                            Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'Added by Manimaran-----s
                            oRewCheck.Checked = False
                            If ValiateOperationSequence(oForm) = False Then
                                SBO_Application.SetStatusBarMessage("The previous operation sequence is not entered", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                Exit Sub
                            End If

                            'oStrSql = "select isnull(sum(U_Passqty),0) pqty ,isnull(sum(U_RewrkQty),0) rqty,isnull(sum(U_ScrapQty),0) sqty from [@PSSIT_WOR2] where U_Pordno = " & oPONoTxt.Value
                            'oRs.DoQuery(oStrSql)
                            'If oRs.RecordCount > 0 Then
                            '    oForm.Items.Item("txtcmqty").Specific.string = CStr(oRs.Fields.Item("pqty").Value)
                            '    oForm.Items.Item("txtrwqty").Specific.value = CStr(oRs.Fields.Item("rqty").Value)
                            '    oForm.Items.Item("txtspqty").Specific.value = CStr(oRs.Fields.Item("sqty").Value)
                            'End If

                            'oRs = Nothing
                            'oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oStrSql = "select isnull(sum(U_Passqty),0) pqty ,isnull(sum(U_RewrkQty),0) rqty,isnull(sum(U_ScrapQty),0) sqty from [@PSSIT_WOR2] where U_Pordno = " & oPONoTxt.Value & " and U_Oprname = '" & oOprCombo.Selected.Description & "'"
                            oRs.DoQuery(oStrSql)
                            If oRs.RecordCount > 0 Then
                                oForm.Items.Item("txtopqty").Specific.string = CStr(oRs.Fields.Item("pqty").Value)
                                oForm.Items.Item("107").Specific.string = CStr(oRs.Fields.Item("rqty").Value)
                                oForm.Items.Item("109").Specific.string = CStr(oRs.Fields.Item("sqty").Value)
                            End If

                            oRs = Nothing
                            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'Added by manimaran-----e
                            Try
                                If Not oOprCombo Is Nothing Then
                                    If oOprCombo.Selected.Value.Length > 0 Then
                                        oStrSql = "Select T0.U_Baslino,T0.U_OprCode,T0.U_OprName,T0.U_RteId,U_ProdQty, " _
                                        & "U_PassQty,U_RewrkQty,U_ScrapQty,U_Rework,U_Mccst,U_Toolcst,U_Lbrcst from [@PSSIT_WOR2] T0 " _
                                        & "Inner Join OWOR T1 On T1.DocNum  = T0.U_Pordno and T1.Series  = T0.U_Pordser " _
                                        & "Where T0.U_OprName = '" & oOprCombo.Selected.Description & "' and T0.U_Baslino = " & oOprCombo.Selected.Value & " and T0.U_Pordno = " & oPONoTxt.Value
                                        oRs.DoQuery(oStrSql)
                                        oParentDB.SetValue("U_Oprcode", oParentDB.Offset, oRs.Fields.Item("U_OprCode").Value)
                                        oParentDB.SetValue("U_Oprname", oParentDB.Offset, oOprCombo.Selected.Description)
                                        oParentDB.SetValue("U_RteID", oParentDB.Offset, oRs.Fields.Item("U_RteID").Value)
                                        oParentDB.SetValue("U_AccProdQty", oParentDB.Offset, oRs.Fields.Item("U_ProdQty").Value)
                                        oParentDB.SetValue("U_AccPassQty", oParentDB.Offset, oRs.Fields.Item("U_PassQty").Value)
                                        oParentDB.SetValue("U_AccRewQty", oParentDB.Offset, oRs.Fields.Item("U_RewrkQty").Value)
                                        oParentDB.SetValue("U_AccScrapQty", oParentDB.Offset, oRs.Fields.Item("U_ScrapQty").Value)
                                        UAccMacCost.Value = oRs.Fields.Item("U_Mccst").Value
                                        UAccToolCost.Value = oRs.Fields.Item("U_Toolcst").Value
                                        UAccLabCost.Value = oRs.Fields.Item("U_Lbrcst").Value
                                        '****************Enabling the Operation Qty Group******************
                                        oForm.Items.Item("txtpdqty").Enabled = True
                                        oForm.Items.Item("txtpsqty").Enabled = True
                                        'Commented by Manimaran------s
                                        'oForm.Items.Item("txtoprwqty").Enabled = True
                                        'oForm.Items.Item("cmboprwres").Enabled = True
                                        'oForm.Items.Item("txtopscqty").Enabled = True
                                        'oForm.Items.Item("cmbopscres").Enabled = True
                                        'Commented by Manimaran------e
                                        ' LoadCombo(matlab)
                                        LoadMachineCombo()
                                        oMachinesDB.InsertRecord(oMachinesDB.Size)
                                        oMachinesDB.Offset = oMachinesDB.Size - 1
                                        oMacMatrix.Clear()
                                        oMacMatrix.FlushToDataSource()
                                        oMacMatrix.AddRow(1, oMacMatrix.RowCount)
                                        ''  oReWrkMatrix.AddRow(1, oReWrkMatrix.RowCount)
                                        ' loadReaCombo(oRwrkUID)
                                        oRs.DoQuery("Select U_ProdQty,U_PassQty,U_RewrkQty,U_ScrapQty,IsNull(U_PenRewQty,0) U_PenRewQty from [@PSSIT_WOR2] T0 " _
                                         & "Where T0.U_OprName = '" & oOprCombo.Selected.Description & "' and T0.U_Baslino = " & oOprCombo.Selected.Value & " and T0.U_Pordno = " & oPONoTxt.Value)
                                        If oRs.RecordCount > 0 Then
                                            oRs.MoveFirst()
                                            URewQty.Value = oRs.Fields.Item("U_PenRewQty").Value
                                        Else
                                            URewQty.Value = "0"
                                        End If
                                    End If
                                End If
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Finally
                                oStrSql = Nothing
                                oRs = Nothing
                                GC.Collect()
                            End Try
                        End If
                        If pVal.ItemUID = "matmac" And pVal.Row > 0 Then
                            '**********************Setting the machine details********************

                            If pVal.ColUID = "colwcno" Then
                                Dim oMacCodeCombo, oMacTypeCombo, oMstopCombo As SAPbouiCOM.ComboBox
                                Dim oWCCodeEdit As SAPbouiCOM.EditText
                                Dim oStrSql As String
                                Dim oCurrentRow As Integer
                                Dim p As Integer
                                Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                If CDbl(oForm.Items.Item("txtpdqty").Specific.string) = 0 Then
                                    Throw New Exception("Enter Produced Qty......")
                                End If
                                Try
                                    oCurrentRow = pVal.Row
                                    oMacMatrix.GetLineData(pVal.Row)
                                    oMacCodeCombo = oMacCodeCol.Cells.Item(oCurrentRow).Specific
                                    oMacTypeCombo = oMTypeCol.Cells.Item(oCurrentRow).Specific
                                    If oMacCodeCombo.Selected.Value.Length > 0 Then
                                        If pVal.Row = oMacMatrix.VisualRowCount Then
                                            oMachinesDB.Offset = oMachinesDB.Size - 1
                                            SetMachinesValue()
                                            SetMachinesDefaultValue()
                                            oMacMatrix.SetLineData(pVal.Row)
                                            oMacMatrix.FlushToDataSource()
                                        End If
                                        oMachinesDB.SetValue("U_wcname", oMachinesDB.Offset, oMacCodeCombo.Selected.Description)
                                        oStrSql = "Select T0.U_deptcode,T0.U_deptdesc,T0.U_OperCost,T0.U_Powecost,T0.U_Cost1, " _
                                        & "T0.U_Cost2 from [@PSSIT_PMWCHDR] T0 Where T0.U_wcno = '" & oMacCodeCombo.Selected.Value & "'"
                                        oRs.DoQuery(oStrSql)
                                        If oRs.RecordCount > 0 Then
                                            oMachinesDB.SetValue("U_Wrkno", oMachinesDB.Offset, oRs.Fields.Item("U_deptcode").Value)
                                            oMachinesDB.SetValue("U_Wrkname", oMachinesDB.Offset, oRs.Fields.Item("U_deptdesc").Value)
                                            oMachinesDB.SetValue("U_Mprcph", oMachinesDB.Offset, oRs.Fields.Item("U_Powecost").Value)
                                            oMachinesDB.SetValue("U_Mohcph1", oMachinesDB.Offset, oRs.Fields.Item("U_Cost1").Value)
                                            oMachinesDB.SetValue("U_Mohcph2", oMachinesDB.Offset, oRs.Fields.Item("U_Cost2").Value)
                                        End If
                                        SetMachinesDefaultValue()
                                        oMacMatrix.SetLineData(pVal.Row)
                                        '****************************Running Machine Operation Cost**************************************
                                        oMacMatrix.GetLineData(pVal.Row)
                                        oMachinesDB.SetValue("U_RMopcph", oMachinesDB.Offset, RunMachineOprCostCalculation(pVal.Row))
                                        oMacMatrix.SetLineData(pVal.Row)
                                        '****************************Running Machine Power Cost******************************************
                                        oMacMatrix.GetLineData(pVal.Row)
                                        oMachinesDB.SetValue("U_RMprcph", oMachinesDB.Offset, RunMachinePowCostCalculation(pVal.Row))
                                        oMacMatrix.SetLineData(pVal.Row)
                                        '****************************Running Machine Other Cost1*****************************************
                                        oMacMatrix.GetLineData(pVal.Row)
                                        oMachinesDB.SetValue("U_RMohcph1", oMachinesDB.Offset, RunMachineOtherCost1Calculation(pVal.Row))
                                        oMacMatrix.SetLineData(pVal.Row)
                                        '****************************Running Machine Other Cost2*****************************************
                                        oMacMatrix.GetLineData(pVal.Row)
                                        oMachinesDB.SetValue("U_RMohcph2", oMachinesDB.Offset, RunMachineOtherCost2Calculation(pVal.Row))
                                        oMacMatrix.SetLineData(pVal.Row)
                                        '****************************Running Machine Total Cost******************************************
                                        oMacMatrix.GetLineData(pVal.Row)
                                        oMachinesDB.SetValue("U_Totcost", oMachinesDB.Offset, RunMachineTotalCost(pVal.Row))
                                        oMacMatrix.SetLineData(pVal.Row)
                                        '************************************************************************************************
                                        oMstopCombo = oMstopCol.Cells.Item(oCurrentRow).Specific
                                        If oMacTypeCombo.Selected.Value = "Operation Time" Then
                                            oMacMatrix.GetLineData(pVal.Row)
                                            oStrSql = "Select T0.*,T1.AcctCode from [@PSSIT_PMWCHDR] T0 " _
                                            & "Inner Join OACT T1 On T1.FormatCode = T0.U_ActAccode Where U_wcno = '" & oMachinesDB.GetValue("U_wcno", oMachinesDB.Offset).Trim() & "'"
                                            oRs.DoQuery(oStrSql)
                                            If oRs.RecordCount > 0 Then
                                                'oMachinesDB.SetValue("U_Accode", oMachinesDB.Offset, oRs.Fields.Item("U_Accode").Value)
                                                oMachinesDB.SetValue("U_Accode", oMachinesDB.Offset, oRs.Fields.Item("AcctCode").Value)
                                                oMachinesDB.SetValue("U_Acname", oMachinesDB.Offset, oRs.Fields.Item("U_Acname").Value)
                                                oMachinesDB.SetValue("U_Mopcph", oMachinesDB.Offset, oRs.Fields.Item("U_OperCost").Value)
                                                'Added by Manimaran--------s
                                                oMachinesDB.SetValue("U_Frtime", oMachinesDB.Offset, oShiftFromTime)
                                                oMachinesDB.SetValue("U_Totime", oMachinesDB.Offset, oShiftToTime)

                                                oMachinesDB.SetValue("U_Qty", oMachinesDB.Offset, oForm.Items.Item("txtpdqty").Specific.string)

                                                'Added by Manimaran--------e
                                            End If
                                            'display operation time added by kabilahan --b


                                            oStrSql = "select distinct U_opertime  from [@PSSIT_RTE1]  r join [@pssit_Oprn] o "
                                            oStrSql += " on r.U_OprCode = o.code where r.U_wcno='" & oMachinesDB.GetValue("U_wcno", oMachinesDB.Offset).Trim() & "' and o.U_Oprname ='" & oOprCombo.Selected.Description & "' group by U_opertime"
                                            oRs.DoQuery(oStrSql)
                                            If oRs.RecordCount > 0 Then
                                                'Modified by Manimaran-------s
                                                For p = oMstopCombo.ValidValues.Count - 1 To 0 Step -1
                                                    oMstopCombo.ValidValues.Remove(p, SAPbouiCOM.BoSearchKey.psk_Index)
                                                Next
                                                'oMachinesDB.SetValue("U_sotime", oMachinesDB.Offset, oRs.Fields.Item("U_opertime").Value)
                                                While Not oRs.EoF
                                                    oMstopCombo.ValidValues.Add(oRs.Fields.Item("U_opertime").Value, oRs.Fields.Item("U_opertime").Value)
                                                    oRs.MoveNext()
                                                End While
                                                'Modified by Manimaran-------e
                                            End If
                                            'added by kabilahan --E
                                        ElseIf oMacTypeCombo.Selected.Value = "Setup Time" Then
                                            oMacMatrix.GetLineData(pVal.Row)
                                            oStrSql = "Select T0.*,T1.AcctCode from [@PSSIT_PMWCHDR] T0 " _
                                            & "Inner Join OACT T1 On T1.FormatCode = T0.U_SActAccode Where U_wcno = '" & oMachinesDB.GetValue("U_wcno", oMachinesDB.Offset).Trim() & "'"
                                            oRs.DoQuery(oStrSql)
                                            If oRs.RecordCount > 0 Then
                                                'oMachinesDB.SetValue("U_Accode", oMachinesDB.Offset, oRs.Fields.Item("U_SAccode").Value)
                                                oMachinesDB.SetValue("U_Accode", oMachinesDB.Offset, oRs.Fields.Item("AcctCode").Value)
                                                oMachinesDB.SetValue("U_Acname", oMachinesDB.Offset, oRs.Fields.Item("U_SAcname").Value)
                                                oMachinesDB.SetValue("U_Mopcph", oMachinesDB.Offset, oRs.Fields.Item("U_SetupCost").Value)
                                                'Added by Manimaran--------s
                                                oMachinesDB.SetValue("U_Frtime", oMachinesDB.Offset, oShiftFromTime)
                                                oMachinesDB.SetValue("U_Totime", oMachinesDB.Offset, oShiftToTime)
                                                oMachinesDB.SetValue("U_Qty", oMachinesDB.Offset, oForm.Items.Item("txtpdqty").Specific.string)
                                                'Added by Manimaran--------e
                                            End If
                                            'display setup time added by kabilahan --b

                                            oStrSql = "select U_Setime  from [@PSSIT_RTE1]  r join [@pssit_Oprn] o "
                                            oStrSql += " on r.U_OprCode = o.code where r.U_wcno='" & oMachinesDB.GetValue("U_wcno", oMachinesDB.Offset).Trim() & "' and o.U_Oprname ='" & oOprCombo.Selected.Description & "' group by U_Setime"
                                            oRs.DoQuery(oStrSql)
                                            If oRs.RecordCount > 0 Then
                                                'Modified by Manimaran-------s
                                                For p = oMstopCombo.ValidValues.Count - 1 To 0 Step -1
                                                    oMstopCombo.ValidValues.Remove(p, SAPbouiCOM.BoSearchKey.psk_Index)
                                                Next
                                                'oMachinesDB.SetValue("U_sotime", oMachinesDB.Offset, oRs.Fields.Item("U_Setime").Value)
                                                While Not oRs.EoF
                                                    oMstopCombo.ValidValues.Add(oRs.Fields.Item("U_Setime").Value, oRs.Fields.Item("U_Setime").Value)
                                                    oRs.MoveNext()
                                                End While
                                                'Modified by Manimaran-------e
                                            End If
                                            'added by kabilahan --E
                                        End If
                                        oMacMatrix.SetLineData(pVal.Row)
                                        '************************************************************************************************
                                        oMacMatrix.FlushToDataSource()
                                        '***************************Loading Tools Data***************************************************
                                        oMacMatrix.GetLineData(pVal.Row)
                                        oWCCodeEdit = oMWCCodeCol.Cells.Item(oCurrentRow).Specific
                                        LoadFixedCostData(pVal.Row, oMacCodeCombo.Selected.Value, oWCCodeEdit.Value)
                                        If oMacTypeCombo.Selected.Value = "Operation Time" Then
                                            LoadToolsData(pVal.Row, oMacCodeCombo.Selected.Value, pVal.Row)
                                        End If
                                        'commented by Manimaran------s
                                        'Reason---- sicame client needs to fill the skill group not labour
                                        'Added by Manimaran-----s
                                        'LoadLab(pVal.Row, oMacCodeCombo.Selected.Value, pVal.Row)
                                        'Added by Manimaran-----e
                                        'commented by Manimaran------e

                                        oMToTimeCol.Cells.Item(oMacMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        SBO_Application.SendKeys("{TAB}")
                                        ' oMAccCodeCol.Cells.Item(oMacMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        If oMacCodeCombo.Selected.Value.Length > 0 Then
                                            oMachinesDB.InsertRecord(oMachinesDB.Size)
                                            oMachinesDB.Offset = oMachinesDB.Size - 1
                                            SetMachinesValue()
                                            SetMachinesDefaultValue()
                                            oMacMatrix.AddRow(1, oMacMatrix.RowCount)
                                        End If
                                    End If
                                Catch ex As Exception
                                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                Finally
                                    oCurrentRow = Nothing
                                    oStrSql = Nothing
                                    oRs = Nothing
                                    GC.Collect()
                                End Try
                            End If
                            'Added by Manimaran- -------s
                            If pVal.ColUID = "coltype" Then
                                Dim oMacTypeCombo, oMstopCombo As SAPbouiCOM.ComboBox
                                Dim oStrSql As String
                                Dim oCurrentRow As Integer
                                oCurrentRow = pVal.Row
                                Dim p As Integer
                                oMacTypeCombo = oMTypeCol.Cells.Item(oCurrentRow).Specific
                                Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oMstopCombo = oMstopCol.Cells.Item(oCurrentRow).Specific
                                If oMacTypeCombo.Selected.Value = "Operation Time" Then
                                    oMacMatrix.GetLineData(pVal.Row)
                                    oStrSql = "Select T0.*,T1.AcctCode from [@PSSIT_PMWCHDR] T0 " _
                                    & "Inner Join OACT T1 On T1.FormatCode = T0.U_ActAccode Where U_wcno = '" & oMachinesDB.GetValue("U_wcno", oMachinesDB.Offset).Trim() & "'"
                                    oRs.DoQuery(oStrSql)
                                    If oRs.RecordCount > 0 Then

                                        oMachinesDB.SetValue("U_Accode", oMachinesDB.Offset, oRs.Fields.Item("AcctCode").Value)
                                        oMachinesDB.SetValue("U_Acname", oMachinesDB.Offset, oRs.Fields.Item("U_Acname").Value)
                                        oMachinesDB.SetValue("U_Mopcph", oMachinesDB.Offset, oRs.Fields.Item("U_OperCost").Value)

                                        oMachinesDB.SetValue("U_Frtime", oMachinesDB.Offset, oShiftFromTime)
                                        oMachinesDB.SetValue("U_Totime", oMachinesDB.Offset, oShiftToTime)

                                    End If



                                    oStrSql = "select distinct U_opertime  from [@PSSIT_RTE1]  r join [@pssit_Oprn] o "
                                    oStrSql += " on r.U_OprCode = o.code where r.U_wcno='" & oMachinesDB.GetValue("U_wcno", oMachinesDB.Offset).Trim() & "' and o.U_Oprname ='" & oOprCombo.Selected.Description & "' group by U_opertime"
                                    oRs.DoQuery(oStrSql)
                                    If oRs.RecordCount > 0 Then
                                        For p = oMstopCombo.ValidValues.Count - 1 To 0 Step -1
                                            oMstopCombo.ValidValues.Remove(p, SAPbouiCOM.BoSearchKey.psk_Index)
                                        Next
                                        While Not oRs.EoF
                                            oMstopCombo.ValidValues.Add(oRs.Fields.Item("U_opertime").Value, oRs.Fields.Item("U_opertime").Value)
                                            oRs.MoveNext()
                                        End While

                                    End If

                                ElseIf oMacTypeCombo.Selected.Value = "Setup Time" Then
                                    oMacMatrix.GetLineData(pVal.Row)
                                    oStrSql = "Select T0.*,T1.AcctCode from [@PSSIT_PMWCHDR] T0 " _
                                    & "Inner Join OACT T1 On T1.FormatCode = T0.U_SActAccode Where U_wcno = '" & oMachinesDB.GetValue("U_wcno", oMachinesDB.Offset).Trim() & "'"
                                    oRs.DoQuery(oStrSql)
                                    If oRs.RecordCount > 0 Then

                                        oMachinesDB.SetValue("U_Accode", oMachinesDB.Offset, oRs.Fields.Item("AcctCode").Value)
                                        oMachinesDB.SetValue("U_Acname", oMachinesDB.Offset, oRs.Fields.Item("U_SAcname").Value)
                                        oMachinesDB.SetValue("U_Mopcph", oMachinesDB.Offset, oRs.Fields.Item("U_SetupCost").Value)

                                        oMachinesDB.SetValue("U_Frtime", oMachinesDB.Offset, oShiftFromTime)
                                        oMachinesDB.SetValue("U_Totime", oMachinesDB.Offset, oShiftToTime)

                                    End If


                                    oStrSql = "select U_Setime  from [@PSSIT_RTE1]  r join [@pssit_Oprn] o "
                                    oStrSql += " on r.U_OprCode = o.code where r.U_wcno='" & oMachinesDB.GetValue("U_wcno", oMachinesDB.Offset).Trim() & "' and o.U_Oprname ='" & oOprCombo.Selected.Description & "' group by U_Setime"
                                    oRs.DoQuery(oStrSql)
                                    If oRs.RecordCount > 0 Then
                                        For p = oMstopCombo.ValidValues.Count - 1 To 0 Step -1
                                            oMstopCombo.ValidValues.Remove(p, SAPbouiCOM.BoSearchKey.psk_Index)
                                        Next
                                        While Not oRs.EoF
                                            oMstopCombo.ValidValues.Add(oRs.Fields.Item("U_Setime").Value, oRs.Fields.Item("U_Setime").Value)
                                            oRs.MoveNext()
                                        End While
                                    End If

                                End If
                            End If

                            'Added by Manimaran---------e
                            '*********Setting the value for the account as per the operation type selected********************
                            If pVal.ColUID = "coltype" Then
                                Dim oMacCodeCombo, oMacTypeCombo As SAPbouiCOM.ComboBox
                                Dim oMacLineId As String
                                Dim oStrSql As String
                                Dim oCurrentRow, IntICount As Integer
                                Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Try
                                    oCurrentRow = pVal.Row
                                    oMacCodeCombo = oMacCodeCol.Cells.Item(oCurrentRow).Specific
                                    oMacTypeCombo = oMTypeCol.Cells.Item(oCurrentRow).Specific
                                    oMacLineId = oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder)
                                    oMacMatrix.GetLineData(pVal.Row)
                                    If oMacTypeCombo.Selected.Value.Length > 0 Then
                                        If oMacTypeCombo.Selected.Value = "Operation Time" Then
                                            oMacMatrix.GetLineData(pVal.Row)
                                            oStrSql = "Select T0.*,T1.AcctCode from [@PSSIT_PMWCHDR] T0 " _
                                            & "Inner Join OACT T1 On T1.FormatCode = T0.U_ActAccode Where U_wcno = '" & oMachinesDB.GetValue("U_wcno", oMachinesDB.Offset).Trim() & "'"
                                            oRs.DoQuery(oStrSql)
                                            If oRs.RecordCount > 0 Then
                                                'oMachinesDB.SetValue("U_Accode", oMachinesDB.Offset, oRs.Fields.Item("U_Accode").Value)
                                                oMachinesDB.SetValue("U_Accode", oMachinesDB.Offset, oRs.Fields.Item("AcctCode").Value)
                                                oMachinesDB.SetValue("U_Acname", oMachinesDB.Offset, oRs.Fields.Item("U_Acname").Value)
                                                oMachinesDB.SetValue("U_Mopcph", oMachinesDB.Offset, oRs.Fields.Item("U_OperCost").Value)
                                            End If

                                        ElseIf oMacTypeCombo.Selected.Value = "Setup Time" Then
                                            oMacMatrix.GetLineData(pVal.Row)
                                            For IntICount = oToolsMatrix.RowCount To 1 Step -1
                                                If oMacCodeCombo.Selected.Value = oTMacNoCol.Cells.Item(IntICount).Specific.value And oMacLineId = oTMLineIDCol.Cells.Item(IntICount).Specific.Value Then
                                                    oToolsMatrix.DeleteRow(IntICount)
                                                    oToolsMatrix.FlushToDataSource()
                                                    BubbleEvent = False
                                                End If
                                            Next
                                            oStrSql = "Select T0.*,T1.AcctCode from [@PSSIT_PMWCHDR] T0 " _
                                            & "Inner Join OACT T1 On T1.FormatCode = T0.U_SActAccode Where U_wcno = '" & oMachinesDB.GetValue("U_wcno", oMachinesDB.Offset).Trim() & "'"
                                            oRs.DoQuery(oStrSql)
                                            If oRs.RecordCount > 0 Then
                                                'oMachinesDB.SetValue("U_Accode", oMachinesDB.Offset, oRs.Fields.Item("U_SAccode").Value)
                                                oMachinesDB.SetValue("U_Accode", oMachinesDB.Offset, oRs.Fields.Item("AcctCode").Value)
                                                oMachinesDB.SetValue("U_Acname", oMachinesDB.Offset, oRs.Fields.Item("U_SAcname").Value)
                                                oMachinesDB.SetValue("U_Mopcph", oMachinesDB.Offset, oRs.Fields.Item("U_SetupCost").Value)
                                            End If

                                        End If
                                        oMacMatrix.SetLineData(pVal.Row)
                                    End If
                                Catch ex As Exception
                                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                Finally
                                    oCurrentRow = Nothing
                                    oStrSql = Nothing
                                    oRs = Nothing
                                    GC.Collect()
                                End Try
                            End If
                        End If
                    End If
                End If
                '********************* Reloads the Combo's if Define New is selected and data added in the Forms************
                If (pVal.FormTypeEx = "FPE") And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE) And pVal.BeforeAction = False Then
                    Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Try
                        If BoolRewDefine = False Then
                            If Not oCmbRew Is Nothing Then
                                'Modified by Manimaran-----s
                                'LoadReasonCombo(oOprRewRsnCombo)
                                loadReaCombo("104")
                                'Modified by Manimaran-----e
                                oRs.DoQuery("Select * from [@PSSIT_ORES] Where DocEntry = (Select IsNull(Max(DocEntry),0) as DocEntry from [@PSSIT_ORES])")
                                If oRs.RecordCount > 0 Then
                                    oRs.MoveFirst()
                                    oCmbRew.Select(oRs.Fields.Item("Code").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                                    BoolRewDefine = True
                                End If
                            End If
                        End If
                        If BoolScpDefine = False Then
                            If Not oCmbScrp Is Nothing Then
                                'Modified by Manimaran-----s
                                'LoadReasonCombo(oOprScrRsnCombo)
                                loadReaCombo("105")
                                'Modified by Manimaran-----e
                                oRs.DoQuery("Select * from [@PSSIT_ORES] Where DocEntry = (Select IsNull(Max(DocEntry),0) as DocEntry from [@PSSIT_ORES])")
                                If oRs.RecordCount > 0 Then
                                    oRs.MoveFirst()
                                    oCmbScrp.Select(oRs.Fields.Item("Code").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                                    BoolScpDefine = True
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    End Try
                End If
                '**********************Selecting a row in the Machine matrix********************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                    If pVal.BeforeAction = False Then
                        If pVal.ItemUID = "matmac" And pVal.Row > 0 Then
                            Try
                                If (pVal.ColUID = "#" Or pVal.ColUID = "colwcno" Or pVal.ColUID = "colwcnam" Or pVal.ColUID = "coltype" Or pVal.ColUID = "colfrtim" Or pVal.ColUID = "coltotim" Or pVal.ColUID = "colrntim" Or pVal.ColUID = "colqty" Or pVal.ColUID = "coladnl1" Or pVal.ColUID = "coladnl2") Then
                                    oMacMatrix.SelectRow(pVal.Row, True, False)
                                End If
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End Try
                        End If
                        '**********************Getting the ItemUID from the selected matrix********************
                        If (pVal.ItemUID = "matmac") And pVal.ColUID = "#" Then
                            Try
                                oMachineUId = pVal.ItemUID
                                oToolsUID = ""
                                oLabourUID = ""
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End Try
                        End If
                        If (pVal.ItemUID = "mattool") And pVal.ColUID = "#" Then
                            Try
                                oMachineUId = ""
                                oToolsUID = pVal.ItemUID
                                oLabourUID = ""
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End Try
                        End If
                        If (pVal.ItemUID = "matlab") And pVal.ColUID = "#" Then
                            Try
                                oLabourUID = pVal.ItemUID
                                oToolsUID = ""
                                oMachineUId = ""
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End Try
                        End If
                        'Added by Manimaran-------s
                        If (pVal.ItemUID = "104") Then 'And pVal.ColUID = "V_-1" Then
                            Try
                                oRwrkUID = pVal.ItemUID
                                oScrpUID = ""
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End Try
                        End If
                        If (pVal.ItemUID = "105") Then 'And pVal.ColUID = "V_-1" Then
                            Try
                                oScrpUID = pVal.ItemUID
                                oRwrkUID = ""
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End Try
                        End If
                        'Added by Manimaran-------e
                    End If
                End If
                '*****************************Validating the values in the items*******************************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                    If pVal.BeforeAction = True Then
                        If pVal.ItemUID = "txtpsqty" Or pVal.ItemUID = "txtoprwqty" Or pVal.ItemUID = "txtopscqty" Then
                            Try
                                If oProdQtyTxt.Value > 0 Then
                                    If ProducedQtyCalculation() > CDbl(oProdQtyTxt.Value) Then
                                        SBO_Application.SetStatusBarMessage("Sum Of Passed Qty,Rework Qty,Scrap Qty should be equal to Produced Qty", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        'BubbleEvent = False
                                    End If
                                End If
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                            End Try
                        End If
                        'Added by Manimaran----------s
                        'Try
                        '    If oForm.Items.Item("txtprdno").Specific.string <> "" And pVal.ItemUID <> "txtscode" Then
                        '        Dim sqry As String
                        '        Dim TPassQty As Double
                        '        Dim rs As SAPbobsCOM.Recordset
                        '        rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '        'sqry = "select isnull(sum(t0.u_prodqty),0) from [@PSSIT_OPEY] t0"
                        '        'sqry = sqry + " where t0.u_pnordno = '" & oForm.Items.Item("txtprdno").Specific.string & "'"
                        '        sqry = "select (isnull(sum(U_Passqty),0)+ isnull(SUM(U_Rewrkqty),0) + isnull(sum(U_scrapqty ),0)) pqty  from [@PSSIT_WOR2] where U_Pordno = '" & oForm.Items.Item("txtprdno").Specific.string & "' and U_Oprname = '" & oOprCombo.Selected.Description & "'"
                        '        rs.DoQuery(sqry)
                        '        If rs.RecordCount > 0 Then
                        '            TPassQty = CDbl(rs.Fields.Item(0).Value)
                        '        End If
                        '        If pVal.ItemUID = "txtpdqty" Then
                        '            If CDbl(oForm.Items.Item("txtplqty").Specific.value) < TPassQty + CDbl(oForm.Items.Item("txtpdqty").Specific.value) Then
                        '                SBO_Application.SetStatusBarMessage("Operation quantity should be less or equal to the Planned quantity", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        '                BubbleEvent = False
                        '            End If
                        '        End If
                        '    End If
                        'Catch ex As Exception

                        'End Try

                        'Added by Manimaran----------e
                        '*****************************Validating the from time in the machine matrix*******************************
                        If pVal.ItemUID = "matmac" Then
                            If pVal.ColUID = "colfrtim" And pVal.Row > 0 Then
                                Dim oCurrentRow As Integer
                                Dim oFromTime As Integer
                                Try
                                    oCurrentRow = pVal.Row
                                    oMacMatrix.GetLineData(pVal.Row)
                                    If CInt(oSftFromTimeTxt.Value.ToString.Replace(":", "")) < CInt(oSftToTimeTxt.Value.Replace(":", "")) Then
                                        If oMachinesDB.GetValue("U_Frtime", oMachinesDB.Offset).Trim().Length > 0 Then
                                            oFromTime = CInt(oMachinesDB.GetValue("U_Frtime", oMachinesDB.Offset).Trim())
                                            If oFromTime < CInt(oSftFromTimeTxt.Value.Replace(":", "")) Then
                                                SBO_Application.SetStatusBarMessage("From Time should not be less than Shift From Time : " & TimeFormat(CStr(oSftFromTimeTxt.Value)), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                oMBoolFromTime = False
                                                BubbleEvent = False
                                            ElseIf oFromTime > CInt(oSftToTimeTxt.Value.Replace(":", "")) Then
                                                SBO_Application.SetStatusBarMessage("To Time should not be greater than Shift To Time " & TimeFormat(CStr(oSftToTimeTxt.Value)), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                oMBoolFromTime = False
                                                BubbleEvent = False
                                            Else
                                                oMBoolFromTime = True
                                                BubbleEvent = True
                                            End If
                                        End If
                                    Else
                                        If oMachinesDB.GetValue("U_Frtime", oMachinesDB.Offset).Trim().Length > 0 Then
                                            oFromTime = CInt(oMachinesDB.GetValue("U_Frtime", oMachinesDB.Offset).Trim())
                                            If CInt(oSftToTimeTxt.Value.Replace(":", "")) <= oFromTime Then
                                                If oFromTime < CInt(oSftFromTimeTxt.Value.Replace(":", "")) Then
                                                    SBO_Application.SetStatusBarMessage("From Time should not be less than Shift From Time : " & TimeFormat(CStr(oSftFromTimeTxt.Value)), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                    oMBoolFromTime = False
                                                    BubbleEvent = False
                                                ElseIf oFromTime < CInt(oSftToTimeTxt.Value.Replace(":", "")) Then
                                                    SBO_Application.SetStatusBarMessage("To Time should not be greater than Shift To Time " & TimeFormat(CStr(oSftToTimeTxt.Value)), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                    oMBoolFromTime = False
                                                    BubbleEvent = False
                                                Else
                                                    oMBoolFromTime = True
                                                    BubbleEvent = True
                                                End If
                                            End If
                                        End If

                                    End If

                                Catch ex As Exception
                                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    BubbleEvent = False
                                End Try
                            End If
                            '*****************************Validating the to time in the machine matrix*******************************
                            If pVal.ColUID = "coltotim" And pVal.Row > 0 Then
                                Dim oCurrentRow As Integer
                                Dim oToTime As Integer
                                Dim ofrTime As Integer
                                Try
                                    oCurrentRow = pVal.Row
                                    oMacMatrix.GetLineData(pVal.Row)
                                    If CInt(oSftFromTimeTxt.Value.Replace(":", "")) < CInt(oSftToTimeTxt.Value.Replace(":", "")) Then
                                        If oMachinesDB.GetValue("U_Totime", oMachinesDB.Offset).Trim().Length > 0 Then
                                            oToTime = CInt(oMachinesDB.GetValue("U_Totime", oMachinesDB.Offset).Trim())
                                            If oToTime < CInt(oSftFromTimeTxt.Value.Replace(":", "")) Then
                                                SBO_Application.SetStatusBarMessage("To Time should be greater than Shift From Time : " & TimeFormat(CStr(oSftFromTimeTxt.Value)), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                oMBoolToTime = False
                                                BubbleEvent = False
                                            ElseIf oToTime > CInt(oSftToTimeTxt.Value.Replace(":", "")) Then
                                                SBO_Application.SetStatusBarMessage("To Time should be less than Shift To Time :" & TimeFormat(CStr(oSftToTimeTxt.Value)), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                oMBoolToTime = False
                                                BubbleEvent = False
                                            Else
                                                oMBoolToTime = True
                                                BubbleEvent = True
                                            End If
                                        End If
                                    Else
                                        If oMachinesDB.GetValue("U_Totime", oMachinesDB.Offset).Trim().Length > 0 Then
                                            oToTime = CInt(oMachinesDB.GetValue("U_Totime", oMachinesDB.Offset).Trim())
                                            ofrTime = CInt(oMachinesDB.GetValue("U_Frtime", oMachinesDB.Offset).Trim())

                                            If ofrTime > oToTime Then
                                                If oToTime > CInt(oSftFromTimeTxt.Value.Replace(":", "")) Then
                                                    SBO_Application.SetStatusBarMessage("To Time should be greater than Shift From Time : " & TimeFormat(CStr(oSftFromTimeTxt.Value)), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                    oMBoolToTime = False
                                                    BubbleEvent = False
                                                ElseIf oToTime > CInt(oSftToTimeTxt.Value.Replace(":", "")) Then
                                                    SBO_Application.SetStatusBarMessage("To Time should be less than Shift To Time :" & TimeFormat(CStr(oSftToTimeTxt.Value)), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                    oMBoolToTime = False
                                                    BubbleEvent = False
                                                Else
                                                    oMBoolToTime = True
                                                    BubbleEvent = True
                                                End If
                                            Else
                                                If ofrTime < oToTime And CInt(oSftToTimeTxt.Value.Replace(":", "")) < oToTime And CInt(oSftFromTimeTxt.Value.Replace(":", "")) > ofrTime Then
                                                    SBO_Application.SetStatusBarMessage("To Time should be lesser than Shift to Time : " & TimeFormat(CStr(oSftToTimeTxt.Value)), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                    oMBoolToTime = False
                                                    BubbleEvent = False
                                                End If
                                            End If
                                        End If
                                    End If
                                Catch ex As Exception
                                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    BubbleEvent = False
                                End Try
                            End If
                            'Commented by Manimaran--------s
                            '*****************************Validating the Qty in the machine matrix*******************************
                            'If pVal.ColUID = "colqty" And pVal.Row > 0 Then
                            '    Dim oCurrentRow As Integer
                            '    Dim oMQtyEdit As SAPbouiCOM.EditText
                            '    Try
                            '        oCurrentRow = pVal.Row
                            '        oMacMatrix.GetLineData(pVal.Row)
                            '        oMQtyEdit = oMQtyCol.Cells.Item(oCurrentRow).Specific
                            '        If oMQtyEdit.Value.Length = 0 Then
                            '            oMQtyEdit.Value = "0.00"
                            '        End If
                            '        If CDbl(oMQtyEdit.Value) > 0 Then
                            '            If MacQtyCalculation() > CDbl(oProdQtyTxt.Value) Then
                            '                SBO_Application.SetStatusBarMessage("Sum Of Machine Qty should be less than or equal to Produced Qty", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            '                BubbleEvent = False
                            '            End If
                            '        End If
                            '    Catch ex As Exception
                            '        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            '        BubbleEvent = False
                            '    End Try
                            'End If
                            'Commented by Manimaran--------e
                        End If
                    End If
                    If pVal.ItemUID = "matlab" Then
                        '*****************************Validating the from time in the Labour matrix*******************************
                        If pVal.ColUID = "colfrtim" And pVal.Row > 0 Then
                            Dim oCurrentRow As Integer
                            Dim oLFromTime As Integer
                            Try
                                oCurrentRow = pVal.Row
                                oLabMatrix.GetLineData(pVal.Row)
                                If CInt(oSftFromTimeTxt.Value.Replace(":", "")) < CInt(oSftToTimeTxt.Value.Replace(":", "")) Then
                                    If oLabourDB.GetValue("U_Frtime", oLabourDB.Offset).Trim().Length > 0 Then
                                        oLFromTime = CInt(oLabourDB.GetValue("U_Frtime", oLabourDB.Offset).Trim())
                                        If oLFromTime < CInt(oSftFromTimeTxt.Value.Replace(":", "")) Then
                                            SBO_Application.SetStatusBarMessage("From Time should not be less than Shift From Time : " & TimeFormat(CStr(oSftFromTimeTxt.Value)), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            oLBoolFromTime = False
                                            BubbleEvent = False
                                        ElseIf oLFromTime > CInt(oSftToTimeTxt.Value.Replace(":", "")) Then
                                            SBO_Application.SetStatusBarMessage("From Time should not be greater than Shift To Time : " & TimeFormat(CStr(oSftToTimeTxt.Value)), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            oLBoolFromTime = False
                                            BubbleEvent = False
                                        Else
                                            oLBoolFromTime = True
                                            BubbleEvent = True
                                        End If
                                    End If
                                Else
                                    If oLabourDB.GetValue("U_Frtime", oLabourDB.Offset).Trim().Length > 0 Then
                                        oLFromTime = CInt(oLabourDB.GetValue("U_Frtime", oLabourDB.Offset).Trim())
                                        If CInt(oSftToTimeTxt.Value.Replace(":", "")) <= oLFromTime Then
                                            If oLFromTime < CInt(oSftFromTimeTxt.Value.Replace(":", "")) Then
                                                SBO_Application.SetStatusBarMessage("From Time should not be less than Shift From Time : " & TimeFormat(CStr(oSftFromTimeTxt.Value)), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                oLBoolFromTime = False
                                                BubbleEvent = False
                                            ElseIf oLFromTime < CInt(oSftToTimeTxt.Value.Replace(":", "")) Then
                                                SBO_Application.SetStatusBarMessage("From Time should not be greater than Shift To Time : " & TimeFormat(CStr(oSftToTimeTxt.Value)), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                oLBoolFromTime = False
                                                BubbleEvent = False
                                            Else
                                                oLBoolFromTime = True
                                                BubbleEvent = True
                                            End If
                                        End If
                                    End If
                                End If
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                            End Try
                        End If
                        '*****************************Validating the to time in the Labour matrix*******************************
                        If pVal.ColUID = "coltotim" And pVal.Row > 0 Then
                            Dim oCurrentRow As Integer
                            Dim oToTime As Integer
                            Dim oLFrTime As Integer
                            Try
                                oCurrentRow = pVal.Row
                                oLabMatrix.GetLineData(pVal.Row)
                                If CInt(oSftFromTimeTxt.Value.Replace(":", "")) < CInt(oSftToTimeTxt.Value.Replace(":", "")) Then
                                    If oLabourDB.GetValue("U_Totime", oLabourDB.Offset).Trim().Length > 0 Then
                                        oToTime = CInt(oLabourDB.GetValue("U_Totime", oLabourDB.Offset).Trim())
                                        If oToTime < CInt(oSftFromTimeTxt.Value.Replace(":", "")) Then
                                            SBO_Application.SetStatusBarMessage("To Time should be greater than Shift From Time : " & TimeFormat(CStr(oSftFromTimeTxt.Value)), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            oLBoolToTime = False
                                            BubbleEvent = False
                                        ElseIf oToTime > CInt(oSftToTimeTxt.Value.Replace(":", "")) Then
                                            SBO_Application.SetStatusBarMessage("To Time should be less than Shift To Time : " & TimeFormat(CStr(oSftToTimeTxt.Value)), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            oLBoolToTime = False
                                            BubbleEvent = False
                                        Else
                                            oLBoolToTime = True
                                            BubbleEvent = True
                                        End If
                                    End If
                                Else
                                    If oLabourDB.GetValue("U_Totime", oLabourDB.Offset).Trim().Length > 0 Then
                                        oToTime = CInt(oLabourDB.GetValue("U_Totime", oLabourDB.Offset).Trim())
                                        oLFrTime = CInt(oLabourDB.GetValue("U_Frtime", oLabourDB.Offset).Trim())

                                        If oLFrTime > oToTime Then
                                            If oToTime > CInt(oSftFromTimeTxt.Value.Replace(":", "")) Then
                                                SBO_Application.SetStatusBarMessage("To Time should be greater than Shift From Time : " & TimeFormat(CStr(oSftFromTimeTxt.Value)), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                oLBoolToTime = False
                                                BubbleEvent = False
                                            ElseIf oToTime > CInt(oSftToTimeTxt.Value.Replace(":", "")) Then
                                                SBO_Application.SetStatusBarMessage("To Time should be less than Shift To Time : " & TimeFormat(CStr(oSftToTimeTxt.Value)), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                oLBoolToTime = False
                                                BubbleEvent = False
                                            Else
                                                oLBoolToTime = True
                                                BubbleEvent = True
                                            End If
                                        Else
                                            If oLFrTime < oToTime And CInt(oSftToTimeTxt.Value.Replace(":", "")) < oToTime And CInt(oSftFromTimeTxt.Value.Replace(":", "")) > oLFrTime Then
                                                SBO_Application.SetStatusBarMessage("To Time should be lesser than Shift to Time : " & TimeFormat(CStr(oSftToTimeTxt.Value)), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                oLBoolToTime = False
                                                BubbleEvent = False
                                            End If
                                        End If

                                    End If
                                End If
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                            End Try
                        End If
                        '*****************************Validating the Qty in the Labour Matrix*******************************
                        If pVal.ColUID = "colqty" And pVal.Row > 0 Then
                            Dim oCurrentRow As Integer
                            Dim oLQtyEdit As SAPbouiCOM.EditText
                            Try
                                oCurrentRow = pVal.Row
                                oLQtyEdit = oLQtyCol.Cells.Item(oCurrentRow).Specific
                                oLabMatrix.GetLineData(pVal.Row)
                                If oLQtyEdit.Value.Length = 0 Then
                                    oLQtyEdit.Value = "0.00"
                                End If
                                If CDbl(oLQtyEdit.Value) > 0 Then
                                    If LabQtyCalculation() > CDbl(oProdQtyTxt.Value) Then
                                        SBO_Application.SetStatusBarMessage("Sum Of Labour Qty should be less than or equal to Produced Qty", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        BubbleEvent = False
                                    End If
                                End If
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                            End Try
                        End If

                    End If
                    '*****************************Validating the Qty in the Tools matrix*******************************
                    If pVal.ItemUID = "mattool" Then
                        If pVal.ColUID = "colqty" And pVal.Row > 0 Then
                            Dim oCurrentRow As Integer
                            Dim oTQtyEdit, oToolCostPerPiece As SAPbouiCOM.EditText
                            Try
                                oCurrentRow = pVal.Row
                                oTQtyEdit = oTQtyCol.Cells.Item(oCurrentRow).Specific
                                oToolsMatrix.GetLineData(pVal.Row)
                                If oTQtyEdit.Value.Length = 0 Then
                                    oTQtyEdit.Value = "0.00"
                                End If
                                If CDbl(oTQtyEdit.Value) > 0 Then
                                    If ToolsQtyCalculation() > CDbl(oProdQtyTxt.Value) Then
                                        SBO_Application.SetStatusBarMessage("Sum Of Tools Qty should be less than or equal to Produced Qty", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        BubbleEvent = False
                                    Else
                                        'oQty = oTQtyCol.Cells.Item(oCurrentRow).Specific
                                        oToolCostPerPiece = oToolCstCol.Cells.Item(oCurrentRow).Specific
                                        oToolsMatrix.GetLineData(pVal.Row)
                                        oToolsDB.SetValue("U_Totcost", oToolsDB.Offset, RunToolTotalCost(oCurrentRow, CDbl(oToolCostPerPiece.Value), CDbl(oTQtyEdit.Value)))
                                        oToolsMatrix.SetLineData(pVal.Row)
                                    End If
                                    oParentDB.SetValue("U_Tottcst", oParentDB.Offset, TotalToolCost())
                                End If
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                            End Try
                        End If
                    End If
                End If
                '****************************Calculating the Run Time in machine matrix*****************************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then
                    If pVal.BeforeAction = False Then
                        If pVal.ItemUID = "matmac" Then
                            '*****************************Calculating the Run Time in machine matrix*******************************
                            If (pVal.ColUID = "colfrtim" And oMBoolFromTime = True) Or (pVal.ColUID = "coltotim" And oMBoolToTime = True) And pVal.Row > 0 Then
                                Dim oDuration As String
                                Dim oFromTimeEdit, oToTimeEdit, oMOprCost As SAPbouiCOM.EditText
                                Dim oCurrentRow As Integer
                                Try
                                    oCurrentRow = pVal.Row
                                    oMacMatrix.GetLineData(pVal.Row)
                                    oFromTimeEdit = oMFromTimeCol.Cells.Item(oCurrentRow).Specific
                                    oToTimeEdit = oMToTimeCol.Cells.Item(oCurrentRow).Specific
                                    oMOprCost = oMOprCstCol.Cells.Item(oCurrentRow).Specific
                                    If oFromTimeEdit.Value.Length > 0 And oToTimeEdit.Value.Length > 0 Then
                                        oDuration = DurationMinsCalculation(oFromTimeEdit.String, oToTimeEdit.String)
                                        'Added by Manimaran-----s
                                        If cmbintime.Selected.Value = "Y" Then
                                            oDuration = oDuration - shiftInterval(oForm.Items.Item("txtscode").Specific.string)
                                            oMachinesDB.SetValue("U_Rntime", oMachinesDB.Offset, oDuration)
                                        Else
                                            oMachinesDB.SetValue("U_Rntime", oMachinesDB.Offset, oDuration)
                                        End If
                                        'Added by Manimaran-----e
                                    End If
                                    oMacMatrix.SetLineData(pVal.Row)
                                    '****************************Running Machine Operation Cost**************************************
                                    oMacMatrix.GetLineData(pVal.Row)
                                    oMachinesDB.SetValue("U_RMopcph", oMachinesDB.Offset, RunMachineOprCostCalculation(pVal.Row))
                                    oMacMatrix.SetLineData(pVal.Row)
                                    '****************************Running Machine Power Cost**************************************
                                    oMacMatrix.GetLineData(pVal.Row)
                                    oMachinesDB.SetValue("U_RMprcph", oMachinesDB.Offset, RunMachinePowCostCalculation(pVal.Row))
                                    oMacMatrix.SetLineData(pVal.Row)
                                    '****************************Running Machine Other Cost1**************************************
                                    oMacMatrix.GetLineData(pVal.Row)
                                    oMachinesDB.SetValue("U_RMohcph1", oMachinesDB.Offset, RunMachineOtherCost1Calculation(pVal.Row))
                                    oMacMatrix.SetLineData(pVal.Row)
                                    '****************************Running Machine Other Cost2**************************************
                                    oMacMatrix.GetLineData(pVal.Row)
                                    oMachinesDB.SetValue("U_RMohcph2", oMachinesDB.Offset, RunMachineOtherCost2Calculation(pVal.Row))
                                    oMacMatrix.SetLineData(pVal.Row)
                                    '****************************Running Machine Total Cost**************************************
                                    oMacMatrix.GetLineData(pVal.Row)
                                    oMachinesDB.SetValue("U_Totcost", oMachinesDB.Offset, RunMachineTotalCost(pVal.Row))
                                    oMacMatrix.SetLineData(pVal.Row)
                                    oParentDB.SetValue("U_Totmcst", oParentDB.Offset, TotalMachineCost())
                                    oMBoolFromTime = True
                                    oMBoolToTime = True
                                Catch ex As Exception
                                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    BubbleEvent = False
                                End Try
                            End If
                        End If
                        '****************************Calculating Total Running Labour Cost**************************************
                        If pVal.ItemUID = "matlab" Then
                            If pVal.ColUID = "colnop" Or pVal.ColUID = "colottime" Or (pVal.ColUID = "colfrtim" And oLBoolFromTime = True) Or (pVal.ColUID = "coltotim" And oLBoolToTime = True) And pVal.Row > 0 Then
                                'Modified by Manimaran------s
                                Dim oLDuration As String
                                Dim oLFromTimeEdit, oLToTimeEdit, ottimeedit, onopedit As SAPbouiCOM.EditText
                                Dim oCurrentRow As Integer
                                Try
                                    oCurrentRow = pVal.Row
                                    oLabMatrix.GetLineData(pVal.Row)
                                    oLFromTimeEdit = oLFromTimeCol.Cells.Item(oCurrentRow).Specific
                                    oLToTimeEdit = oLToTimeCol.Cells.Item(oCurrentRow).Specific
                                    ottimeedit = oLabMatrix.Columns.Item("colottime").Cells.Item(oCurrentRow).Specific
                                    onopedit = oLabMatrix.Columns.Item("colnop").Cells.Item(oCurrentRow).Specific
                                    If ottimeedit.Value.Length = 0 Then
                                        ottimeedit.Value = 0
                                    End If
                                    If oLFromTimeEdit.Value.Length > 0 And oLToTimeEdit.Value.Length > 0 Then
                                        oLDuration = DurationMinsCalculation(oLFromTimeEdit.String, oLToTimeEdit.String)
                                        'Added by Manimaran-----s
                                        If cmbintime.Selected.Value = "Y" Then
                                            oLDuration = oLDuration - shiftInterval(oForm.Items.Item("txtscode").Specific.string)
                                            oLabourDB.SetValue("U_Wrktime", oLabourDB.Offset, CStr(CDbl(oLDuration) + CDbl(ottimeedit.Value)))
                                        Else
                                            oLabourDB.SetValue("U_Wrktime", oLabourDB.Offset, CStr(CDbl(oLDuration) + CDbl(ottimeedit.Value)))
                                        End If
                                        'Added by Manimaran-----e

                                    End If
                                    'Modified by Manimaran------e
                                    oLabMatrix.SetLineData(pVal.Row)
                                    '****************************Total Running Labour Cost**************************************
                                    oLabMatrix.GetLineData(pVal.Row)
                                    'Modified by Manimaran-----s
                                    If onopedit.Value = "" Then
                                        SBO_Application.SetStatusBarMessage("Enter No. of Persons....", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        Exit Sub
                                    End If
                                    'oLabourDB.SetValue("U_Totcost", oLabourDB.Offset, RunLabourTotalCost(pVal.Row))
                                    oLabourDB.SetValue("U_Totcost", oLabourDB.Offset, RunLabourTotalCost(pVal.Row) * CDbl(onopedit.Value))
                                    oLabMatrix.SetLineData(pVal.Row)
                                    'oParentDB.SetValue("U_Totlcst", oParentDB.Offset, TotalLabourCost())
                                    oParentDB.SetValue("U_Totlcst", oParentDB.Offset, TotalLabourCost() * CDbl(onopedit.Value))
                                    'Modified by Manimaran-----e
                                    oLBoolFromTime = True
                                    oLBoolToTime = True
                                Catch ex As Exception
                                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    BubbleEvent = False
                                End Try
                            End If
                        End If
                        'Added by Manimaran------s
                        If pVal.ItemUID = "txtpdqty" Then
                            Dim sqry As String
                            Dim TPassQty As Double
                            Dim rs As SAPbobsCOM.Recordset
                            rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'sqry = "select isnull(sum(t0.u_prodqty),0) from [@PSSIT_OPEY] t0"
                            'sqry = sqry + " where t0.u_pnordno = '" & oForm.Items.Item("txtprdno").Specific.string & "'"
                            sqry = "select (isnull(sum(U_Passqty),0)+ isnull(SUM(U_Rewrkqty),0) + isnull(sum(U_scrapqty ),0)) pqty  from [@PSSIT_WOR2] where U_Pordno = '" & oForm.Items.Item("txtprdno").Specific.string & "' and U_Oprname = '" & oOprCombo.Selected.Description & "'"
                            rs.DoQuery(sqry)
                            If rs.RecordCount > 0 Then
                                TPassQty = CDbl(rs.Fields.Item(0).Value)
                            End If
                            If pVal.ItemUID = "txtpdqty" Then
                                If CDbl(oForm.Items.Item("txtplqty").Specific.value) < TPassQty + CDbl(oForm.Items.Item("txtpdqty").Specific.value) Then
                                    'Throw New Exception("Produced quantity should be less or equal to the operation quantity")
                                End If
                            End If
                        End If
                        'Added by Manimaran------e
                    End If
                End If
                '****************************Form_Resize() method is called**************************************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_RESIZE And pVal.BeforeAction = False Then
                    Try
                        Form_Resize()
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
    'Added by Manimaran------s
    Private Function shiftInterval(ByVal scode As String) As Double
        Dim sqry As String
        Dim rs As SAPbobsCOM.Recordset
        rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        sqry = "select U_Sbreak from [@PSSIT_OSFT] where code = '" & scode & "' "
        rs.DoQuery(sqry)
        If rs.RecordCount > 0 Then
            Dim inttime As DateTime
            Dim oDuration As String
            Dim intim As String = rs.Fields.Item(0).Value.ToString
            If intim.Length > 2 Then
                intim = Strings.Left(intim, 1) + ":" + Strings.Right(intim, 2)
            Else
                intim = "0" + ":" + Strings.Right(intim, 2)
            End If

            Try
                inttime = Convert.ToDateTime(Date.Parse(intim))
                ''Dim runLength As System.TimeSpan = inttime.
                ''Dim secs As Integer = runLength.Seconds
                ''Dim minutes As Integer = runLength.Minutes
                ''Dim hours As Integer = runLength.Hours
                oDuration = inttime.Hour * 60 + inttime.Minute   'runLength.Hours * 60 + runLength.Minutes

            Catch ex As Exception
                Throw ex
            End Try
            Return oDuration
        End If
    End Function

#Region "Check Production order status"
    Private Function checkProductionOrderstatus(ByVal strDocEntry As String) As Boolean
        Dim oTempRs As SAPbobsCOM.Recordset
        oTempRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRs.DoQuery("Select * from OWOR where docnum=" & strDocEntry)
        If oTempRs.RecordCount > 0 Then
            If oTempRs.Fields.Item("Status").Value = "L" Then
                Return False
            Else
                Return True
            End If
        End If
    End Function
#End Region
    'Private Sub updateQuantities()
    '    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
    '        Dim sqry As String
    '        Dim rs As SAPbobsCOM.Recordset
    '        rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        sqry = "select isnull(sum(t0.u_passqty),0),isnull(sum(t1.U_Rewrkqty),0),isnull(sum(t2.U_scrapqty),0) from [@PSSIT_OPEY] t0"
    '        sqry = sqry + " inner join [@PSSIT_PEY5] t1 on t0.docentry = t1.docentry"
    '        sqry = sqry + " inner join [@PSSIT_PEY6] t2 on t0.docentry = t2.docentry"
    '        sqry = sqry + " where t0.u_pnordno = " & oForm.Items.Item("txtprdno").Specific.string & ""
    '        rs.DoQuery(sqry)
    '        If rs.RecordCount > 0 Then
    '            oForm.Items.Item("txtcmqty").Specific.value = CDbl(oForm.Items.Item("txtpsqty").Specific.value) + CDbl(rs.Fields.Item(0).Value)
    '            oForm.Items.Item("txtrwqty").Specific.value = TotRwrkQty + CDbl(rs.Fields.Item(1).Value)
    '            oForm.Items.Item("txtspqty").Specific.value = TotScrpQty + CDbl(rs.Fields.Item(2).Value)
    '        Else
    '            oForm.Items.Item("txtcmqty").Specific.value = oForm.Items.Item("txtpsqty").Specific.value
    '            oForm.Items.Item("txtrwqty").Specific.value = TotRwrkQty
    '            oForm.Items.Item("txtspqty").Specific.value = TotScrpQty
    '        End If
    '    End If
    'End Sub
    'Added by Manimaran------e
    ''' <summary>
    ''' SetItemEnabled() method is called to set the item enabled as per the form mode.
    ''' setting the focus to the ProductionEntryNo EditText.
    ''' </summary>
    ''' <param name="pVal"></param>
    ''' <param name="BubbleEvent"></param>
    ''' <remarks></remarks>
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Dim FormID As String
        Dim oMacCodeCombo As SAPbouiCOM.ComboBox
        Dim IntIcount, IntJCount, IntKCount As Integer
        Dim oToolsDelCode, oLabDelCode, oFCDelCode As SAPbouiCOM.EditText
        Try
            FormID = SBO_Application.Forms.ActiveForm.UniqueID
            '*****************************Setting Item Enabled in FIND Mode*******************************
            If pVal.MenuUID = "1281" And FormID = "FPE" Then
                If pVal.BeforeAction = False Then
                    SetItemEnabled()
                    oPENoTxt.Active = True
                End If
            End If
            '*****************************Initiating Default Values in ADD Mode*******************************
            If pVal.MenuUID = "1282" And pVal.BeforeAction = False And FormID = "FPE" Then
                SetItemEnabled()
                AccKeyCheck()
                oForm.Freeze(True)
                UPODocEnt.Value = ""
                UTotFCost.Value = 0
                oPONoTxt.Active = True
                oJVNoTxt.Value = 0
                oPODateTxt.String = "t" ' System.DateTime.Today.Date.ToString("dd/MM/yyyy")
                SBO_Application.SendKeys("{TAB}")
                oPESeriesCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                With oForm.DataSources.DBDataSources.Item("@PSSIT_OPEY")
                    .SetValue("DocNum", .Offset, oForm.BusinessObject.GetNextSerialNumber(Trim(.GetValue("Series", .Offset))).ToString)
                End With
                oDocDateTxt.String = "t" 'System.DateTime.Today.Date.ToString("dd/MM/yyyy")
                SBO_Application.SendKeys("{TAB}")
                oParentDB.SetValue("U_Rework", oParentDB.Offset, "N")
                oForm.Freeze(False)
            End If
            '************Adding a row to the Matrix - While ADD ROW is clicked*******************************
            If pVal.MenuUID = "1292" And pVal.BeforeAction = True And FormID = "FPE" Then
                If oMachineUId = "matmac" Then
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                    If oMacMatrix.RowCount = 0 Then
                        If oOprCombo.Selected.Description <> "" Then
                            If oOprCombo.Selected.Description.Length > 0 Then
                                oMachinesDB.InsertRecord(oMachinesDB.Size)
                                oMachinesDB.Offset = oMachinesDB.Size - 1
                                oMacMatrix.Clear()
                                SetMachinesValue()
                                SetMachinesDefaultValue()
                                oMacMatrix.FlushToDataSource()
                                oMacMatrix.AddRow(1, oMacMatrix.RowCount)
                            End If
                        End If
                    ElseIf oMacMatrix.RowCount > 0 Then
                        If oOprCombo.Selected.Description <> "" Then
                            If oOprCombo.Selected.Description.Length > 0 Then
                                oMachinesDB.Offset = oMachinesDB.Size - 1
                                SetMachinesValue()
                                SetMachinesDefaultValue()
                                oMacMatrix.AddRow(1, oMacMatrix.RowCount)
                            End If
                        End If
                    End If
                End If
                '*****************************Adding Rows to Tools Matrix*******************************
                'If oToolsUID = "mattool" Then
                '    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                '        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                '    End If
                '    If oToolsMatrix.RowCount = 0 Then
                '        oToolsDB.InsertRecord(oToolsDB.Size)
                '        oMacMatrix.GetLineData(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                '        oMacCodeCombo = oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific
                '        oToolCodeCombo = oToolCodeCol.Cells.Item(oToolsMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific
                '        AddToolsRow(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder), _
                '        oMacCodeCombo.Selected.Value, oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                '        oToolsDB.SetValue("Code", oToolsDB.Offset, oToolsSerialNo)
                '        oToolsMatrix.SetLineData(oToolsMatrix.RowCount)
                '    ElseIf oToolsMatrix.RowCount > 0 Then
                '        'oToolCodeCombo = oToolCodeCol.Cells.Item(oToolsMatrix.RowCount).Specific
                '        If Len(oToolsDB.GetValue("U_Toolcode", oToolsDB.Offset).Trim()) <= 0 Then
                '            SBO_Application.SetStatusBarMessage("Tool Details should not be empty", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                '            BubbleEvent = False
                '        End If
                '        If Len(oToolsDB.GetValue("U_Toolcode", oToolsDB.Offset).Trim()) > 0 Then
                '            oMacMatrix.GetLineData(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                '            oMacCodeCombo = oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific
                '            oToolCodeCombo = oToolCodeCol.Cells.Item(oToolsMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific
                '            AddToolsRow(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder), _
                '            oMacCodeCombo.Selected.Value, oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                '            oToolsDB.SetValue("Code", oToolsDB.Offset, oToolsSerialNo)
                '            oToolsMatrix.SetLineData(oToolsMatrix.RowCount)
                '        End If
                '    End If
                'End If
                'Added by Manimaran-----s
                If oRwrkUID = "104" Then
                    If oReWrkMatrix.VisualRowCount = 0 Then
                        oReWrkMatrix.AddRow()
                        loadReaCombo(oRwrkUID)
                    Else
                        If CDbl(oReWrkMatrix.Columns.Item("V_1").Cells.Item(oReWrkMatrix.VisualRowCount).Specific.string) <> 0 Then
                            oForm.DataSources.DBDataSources.Item("@PSSIT_PEY5").Clear()
                            oReWrkMatrix.AddRow()
                            loadReaCombo(oRwrkUID)
                        End If
                    End If
                    oScrpUID = ""
                End If

                If oScrpUID = "105" Then
                    If oScrpMatrix.VisualRowCount = 0 Then
                        oScrpMatrix.AddRow()
                        loadReaCombo(oScrpUID)
                    Else
                        If oScrpMatrix.Columns.Item("V_1").Cells.Item(oScrpMatrix.VisualRowCount).Specific.string <> "" Then
                            oForm.DataSources.DBDataSources.Item("@PSSIT_PEY6").Clear()
                            oScrpMatrix.AddRow()
                            loadReaCombo(oScrpUID)
                        End If
                    End If
                    oRwrkUID = ""
                End If

                'Added by Manimaran-----e

                '*****************************Adding Rows to Labour Matrix*******************************
                If oLabourUID = "matlab" Then
                    Dim sql As String
                    Dim rs As SAPbobsCOM.Recordset
                    Dim k As Integer
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                    If oLabMatrix.RowCount = 0 Then
                        oLabourDB.InsertRecord(oLabourDB.Size)
                        If oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder) > 0 Then
                            oMacMatrix.GetLineData(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                            oMacCodeCombo = oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific
                            AddLabourRow(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder), _
                            oMacCodeCombo.Selected.Value, oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                            oLabourDB.SetValue("Code", oLabourDB.Offset, oLabSerialNo)
                            'Added by Manimaran------s
                            rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'sql = "select a.U_Skilgrp  from [@PSSIT_RTE2]  a"
                            'sql = sql + " inner join [@PSSIT_RTE1] b on a.U_wcno = b.U_wcno "
                            'sql = sql + " inner join [@PSSIT_RTE4] c on b.U_OprCode = c.U_Oprcode "
                            'sql = sql + " where a.U_wcno = '" & oMacCodeCombo.Selected.Value & "' and c.U_Oprname = '" & oOprCombo.Selected.Description & "'"
                            'sql = sql + " group by a.U_Skilgrp "
                            sql = "select a.Code  from [@PSSIT_OLGP]  a where a.U_Active = 'Y'"
                            rs.DoQuery(sql)
                            If rs.RecordCount > 0 Then
                                For k = oLSkGroupCodeCol.ValidValues.Count - 1 To 0 Step -1
                                    oLSkGroupCodeCol.ValidValues.Remove(k, SAPbouiCOM.BoSearchKey.psk_Index)
                                Next
                                While Not rs.EoF
                                    oLSkGroupCodeCol.ValidValues.Add(rs.Fields.Item(0).Value, rs.Fields.Item(0).Value)
                                    rs.MoveNext()
                                End While
                            End If
                            oLabourDB.SetValue("U_Frtime", oLabourDB.Offset, oShiftFromTime)
                            oLabourDB.SetValue("U_Totime", oLabourDB.Offset, oShiftToTime)
                            'added by Manimaran------e
                            oLabMatrix.SetLineData(oLabMatrix.RowCount)
                        Else
                            SBO_Application.SetStatusBarMessage("Select the machine for which the labour to be added", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                    ElseIf oLabMatrix.RowCount > 0 Then
                        If oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder) > 0 Then
                            'If Len(oLabCodeCol.Cells.Item(oLabMatrix.RowCount).Specific.value) <= 0 Then
                            If CDbl(oLabMatrix.Columns.Item("colwktim").Cells.Item(oLabMatrix.RowCount).Specific.string) = 0 Then
                                SBO_Application.SetStatusBarMessage("Labour Details should not be empty", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                            End If
                            'If Len(oLabCodeCol.Cells.Item(oLabMatrix.RowCount).Specific.value) > 0 Then
                            If CDbl(oLabMatrix.Columns.Item("colwktim").Cells.Item(oLabMatrix.RowCount).Specific.string) > 0 Then
                                oMacMatrix.GetLineData(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                                oMacCodeCombo = oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific
                                If oMachinesDB.GetValue("U_wcno", oMachinesDB.Offset).Trim.Length > 0 Then
                                    AddLabourRow(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder), _
                                    oMacCodeCombo.Selected.Value, oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                                    oLabourDB.SetValue("Code", oLabourDB.Offset, oLabSerialNo)
                                    'Added by Manimaran------s                                   
                                    oLabourDB.SetValue("U_Frtime", oLabourDB.Offset, oShiftFromTime)
                                    oLabourDB.SetValue("U_Totime", oLabourDB.Offset, oShiftToTime)
                                    'added by Manimaran------e
                                    oLabMatrix.SetLineData(oLabMatrix.RowCount)
                                End If
                            End If
                        Else
                            SBO_Application.SetStatusBarMessage("Select the machine for which the labour to be added", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                    End If
                    rs = Nothing
                End If

            End If
            '*****************************LoadFCDataFromDB(),LoadToolsDataFromDB(),LoadLabourDataFromDB() is called.*******************************
            If (pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") And pVal.BeforeAction = False And FormID = "FPE" Then
                Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Try
                    oRs.DoQuery("Select * from [@PSSIT_OPEY]")
                    If oRs.RecordCount > 0 Then
                        oForm.Freeze(True)
                        oRs1.DoQuery("Select * from [@PSSIT_OPEY] where DocNum = " & oPENoTxt.Value)
                        If oOprCombo.ValidValues.Count > 0 Then
                            For IntIcount = oOprCombo.ValidValues.Count - 1 To 0 Step -1
                                oOprCombo.ValidValues.Remove(IntIcount, SAPbouiCOM.BoSearchKey.psk_Index)
                            Next
                        End If
                        If oRs1.RecordCount > 0 Then
                            oRs1.MoveFirst()
                            While oRs1.EoF = False
                                oOprCombo.ValidValues.Add(oRs1.Fields.Item("U_Oplnid").Value, oRs1.Fields.Item("U_OprName").Value)
                                oRs1.MoveNext()
                            End While
                        End If
                        oForm.Items.Item("cmbopcd").DisplayDesc = True
                        oOprCombo.Select(oOprLineIDTxt.Value, SAPbouiCOM.BoSearchKey.psk_ByDescription)
                        LoadFCDataFromDB()
                        LoadMachineDataFromDB()
                        LoadToolsDataFromDB()
                        LoadLabourDataFromDB()
                        If oPONoTxt.Value.Length > 0 Then
                            LoadProdOrderDocEntryNo(oPONoTxt.Value)
                        End If
                        GLMethod(oGLAccTxt.Value)
                        'SetItemEnabled()
                        If oMacMatrix.RowCount > 0 Then
                            oMacMatrix.SelectRow(1, True, False)
                        End If
                        If oMacMatrix.RowCount > 0 Then
                            oMacMatrix.SelectRow(1, True, False)
                        End If
                        'Dim intFrmTime, intToTime As Integer
                        'intFrmTime = CInt(oSftFromTimeTxt.Value)
                        'intToTime = CInt(oSftToTimeTxt.Value)

                        'oSftFromTimeTxt.Value = Format(intFrmTime, "00:00") 'oShiftFromTime.Format("00:00")
                        'oSftToTimeTxt.Value = Format(intToTime, "00:00") 'oShiftToTime
                        Dim ostrsql As String
                        '   oStrSql = "select isnull(sum(U_Passqty),0) pqty ,isnull(sum(U_RewrkQty),0) rqty,isnull(sum(U_ScrapQty),0) sqty from [@PSSIT_WOR2] where U_Pordno = " & oPONoTxt.Value & " and U_Oprname = '" & oOprCombo.Selected.Description & "'"
                        ostrsql = "select isnull(sum(U_Passqty),0) pqty ,isnull(sum(U_RewrkQty),0) rqty,isnull(sum(U_ScrapQty),0) sqty from [@PSSIT_WOR2] where U_Pordno = " & oPONoTxt.Value
                        oRs.DoQuery(ostrsql)
                        If oRs.RecordCount > 0 Then
                            'oForm.Items.Item("txtopqty").Specific.string = CStr(oRs.Fields.Item("pqty").Value)
                            oForm.Items.Item("txtcmqty").Specific.string = CStr(oRs.Fields.Item(0).Value)
                            'oForm.Items.Item("107").Specific.string = CStr(oRs.Fields.Item("rqty").Value)
                            'oForm.Items.Item("109").Specific.string = CStr(oRs.Fields.Item("sqty").Value)
                            oForm.Items.Item("txtrwqty").Specific.string = CStr(oRs.Fields.Item("rqty").Value)
                            oForm.Items.Item("txtspqty").Specific.string = CStr(oRs.Fields.Item("sqty").Value)

                        End If
                        If checkProductionOrderstatus(oPONoTxt.Value) = False Then
                            disable()
                        Else
                            SetItemEnabled()
                            disable()
                        End If
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                        oForm.Freeze(False)
                    Else
                        oPONoTxt.Active = True
                    End If
                Catch ex As Exception
                    SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End Try
            End If
            '********************************Deleting the selected row from the matrix ***************************
            If pVal.MenuUID = "1293" And pVal.BeforeAction = True And FormID = "FPE" Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End If
                '********************************Deleting the selected row in the machine matrix***************************
                If oMachineUId = "matmac" Then
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If oMacMatrix.RowCount > 0 Then
                            oMacCodeCombo = oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific
                            Dim oMachineCode As String = oMachinesDB.GetValue("U_wcno", oMachinesDB.Offset).Trim()
                            For IntIcount = oToolsMatrix.RowCount To 1 Step -1
                                oToolsMatrix.GetLineData(IntIcount)
                                If oMachineCode = oTMacNoCol.Cells.Item(IntIcount).Specific.Value Then
                                    oToolsDelCode = oTCodeCol.Cells.Item(IntIcount).Specific
                                    If PSSIT_PEY3.GetByKey(oToolsDelCode.Value) = True Then
                                        Dim I As Integer = PSSIT_PEY3.Remove()
                                        oToolsMatrix.DeleteRow(IntIcount)
                                        oToolsMatrix.FlushToDataSource()
                                    ElseIf PSSIT_PEY3.GetByKey(oToolsDelCode.Value) = False Then
                                        oToolsMatrix.DeleteRow(IntIcount)
                                        oToolsMatrix.FlushToDataSource()
                                        oToolsMatrix.LoadFromDataSource()
                                    End If
                                End If
                            Next
                            For IntKCount = oFCMatrix.RowCount To 1 Step -1
                                oFCMatrix.GetLineData(IntKCount)
                                If oMachineCode = oFMacCodeCol.Cells.Item(IntKCount).Specific.Value Then
                                    oFCDelCode = oFCodeCol.Cells.Item(IntKCount).Specific
                                    If PSSIT_PEY4.GetByKey(oFCDelCode.Value) = True Then
                                        Dim I As Integer = PSSIT_PEY4.Remove()
                                        oFCMatrix.DeleteRow(IntJCount)
                                        oFCMatrix.FlushToDataSource()
                                    ElseIf PSSIT_PEY4.GetByKey(oFCDelCode.Value) = False Then
                                        oFCMatrix.DeleteRow(IntJCount)
                                        oFCMatrix.FlushToDataSource()
                                        oFCMatrix.LoadFromDataSource()
                                    End If
                                End If
                            Next
                            '********************************Deleting the selected row in the labour matrix***************************
                            For IntJCount = oLabMatrix.RowCount To 1 Step -1
                                oLabMatrix.GetLineData(IntJCount)
                                If oMachineCode = oLMacNoCol.Cells.Item(IntJCount).Specific.Value Then
                                    oLabDelCode = oLabCodeCol.Cells.Item(IntJCount).Specific
                                    If PSSIT_PEY2.GetByKey(oLabDelCode.Value) = True Then
                                        Dim I As Integer = PSSIT_PEY2.Remove()
                                        oLabMatrix.DeleteRow(IntJCount)
                                        oLabMatrix.FlushToDataSource()
                                    ElseIf PSSIT_PEY2.GetByKey(oLabDelCode.Value) = False Then
                                        oLabMatrix.DeleteRow(IntJCount)
                                        oLabMatrix.FlushToDataSource()
                                        oLabMatrix.LoadFromDataSource()
                                    End If
                                End If
                            Next
                            oMacMatrix.DeleteRow(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                            oMacMatrix.FlushToDataSource()
                        End If
                    End If
                End If
                '********************************Deleting the selected row in the tools matrix***************************
                If oToolsUID = "mattool" Then
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        oToolsDelCode = oTCodeCol.Cells.Item(oToolsMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific
                        If PSSIT_PEY3.GetByKey(oToolsDelCode.Value) = True Then
                            Dim I As Integer = PSSIT_PEY3.Remove()
                            oToolsMatrix.DeleteRow(oToolsMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                            oToolsMatrix.FlushToDataSource()
                        Else
                            oToolsMatrix.DeleteRow(oToolsMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                            oToolsMatrix.FlushToDataSource()
                        End If
                    End If
                End If
                '********************************Deleting the selected row in the labour matrix***************************
                If oLabourUID = "matlab" Then
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        oLabDelCode = oLCodeCol.Cells.Item(oLabMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific
                        If PSSIT_PEY2.GetByKey(oLabDelCode.Value) = True Then
                            Dim I As Integer = PSSIT_PEY2.Remove()
                            oLabMatrix.DeleteRow(oLabMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                            oLabMatrix.FlushToDataSource()
                        Else
                            oLabMatrix.DeleteRow(oLabMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                            oLabMatrix.FlushToDataSource()
                        End If
                    End If
                End If
                BubbleEvent = False
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    Private Function ChkSUser() As Boolean
        Rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        sQry = "select superuser from ousr where u_name = '" & oCompany.UserName & "'"
        Rs.DoQuery(sQry)
        If Rs.RecordCount > 0 Then
            If (Rs.Fields.Item("superuser").Value = "N") Then
                Return False
            Else
                Return True
            End If

        End If
    End Function
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
    Private Sub ProductionEntry_ChooseFromList(ByVal ControlName As String, ByVal ColumnUID As String, ByVal CurrentRow As String, ByVal ChoosefromListUID As String, ByVal ChooseFromListSelectedObjects As SAPbouiCOM.DataTable) Handles Me.ChooseFromList
        Dim oDataTable As SAPbouiCOM.DataTable = ChooseFromListSelectedObjects
        Dim oPONo, oPOSeries, oItemCode, oGLMethod, oItemName, oWhsCode, oWhsName, oPlannedQty, oCompQty, oRejectedQty, _
        oShiftCode, oShiftDesc, oLabourCode, oSkGroupCode, oSkGroupName, oLLabRate, oLAcCode, oLAcName, oLActAcCode As String
        Dim oLReqNo As String = ""
        Dim oPODate As Object
        Dim oRs As SAPbobsCOM.Recordset
        Dim oMacLineID, oMacDocEntry As Integer, oMacCode, oLabCode As String
        Dim oMacCodeCombo As SAPbouiCOM.ComboBox
        Dim oStrSql As String
        Dim oAccCode As String = ""
        Try
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                '*********************************Production Order**********************************
                If (ControlName = "btnpo" Or ControlName = "txtprdno") And (ChoosefromListUID = "POBtnLst" Or ChoosefromListUID = "POTxtLst") Then
                    If Not oDataTable Is Nothing Then
                        oPONo = oDataTable.GetValue("DocNum", 0)
                        oPOSeries = oDataTable.GetValue("Series", 0)
                        oPODate = CDate(oDataTable.GetValue("PostDate", 0)).ToString("yyyyMMdd")
                        oItemCode = oDataTable.GetValue("ItemCode", 0)
                        oRs.DoQuery("Select ItemName,GLMethod From OITM Where ItemCode = '" & oItemCode & "'")
                        oGLMethod = oRs.Fields.Item("GLMethod").Value
                        oItemName = oRs.Fields.Item("ItemName").Value
                        oWhsCode = oDataTable.GetValue("Warehouse", 0)
                        oRs.DoQuery("Select WhsName from OWHS Where WhsCode = '" & oWhsCode & "'")
                        oWhsName = oRs.Fields.Item("WhsName").Value
                        oPlannedQty = oDataTable.GetValue("PlannedQty", 0)
                        'Modified by Manimaran-----s
                        'oCompQty = oDataTable.GetValue("CmpltQty", 0)
                        oRs = Nothing
                        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oStrSql = "select top 1 isnull(sum(U_Passqty),0) pqty ,isnull(sum(U_RewrkQty),0) rqty,isnull(sum(U_ScrapQty),0) sqty,U_Oprname from [@PSSIT_WOR2] where U_Pordno = " & oPONo & " group by U_Oprname order by pqty asc"
                        oRs.DoQuery(oStrSql)
                        If oRs.RecordCount > 0 Then
                            oCompQty = CStr(oRs.Fields.Item("pqty").Value)
                        End If
                        oRs = Nothing
                        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oStrSql = "select isnull(sum(U_RewrkQty),0) rqty,isnull(sum(U_ScrapQty),0) sqty from [@PSSIT_WOR2] where U_Pordno = " & oPONo
                        oRs.DoQuery(oStrSql)
                        If oRs.RecordCount > 0 Then
                            oForm.Items.Item("txtrwqty").Specific.value = CStr(oRs.Fields.Item("rqty").Value)
                            oForm.Items.Item("txtspqty").Specific.value = CStr(oRs.Fields.Item("sqty").Value)
                        End If
                        'Modified by Manimaran-----e
                        oRejectedQty = oDataTable.GetValue("RjctQty", 0)
                        LoadProdOrderDocEntryNo(oPONo)
                        oParentDB.Offset = oParentDB.Size - 1
                        oParentDB.SetValue("U_Pnordno", oParentDB.Offset, oPONo)
                        oParentDB.SetValue("U_Pnordser", oParentDB.Offset, oPOSeries)
                        oParentDB.SetValue("U_Pordt", oParentDB.Offset, oPODate)
                        oParentDB.SetValue("U_Itemcode", oParentDB.Offset, oItemCode)
                        oParentDB.SetValue("U_GLMethod", oParentDB.Offset, oGLMethod)
                        oParentDB.SetValue("U_ItemName", oParentDB.Offset, oItemName)
                        oParentDB.SetValue("U_WhsCode", oParentDB.Offset, oWhsCode)
                        oParentDB.SetValue("U_Whsname", oParentDB.Offset, oWhsName)
                        oParentDB.SetValue("U_Planqty", oParentDB.Offset, oPlannedQty)
                        oParentDB.SetValue("U_Comqty", oParentDB.Offset, oCompQty)
                        oParentDB.SetValue("U_RejQty", oParentDB.Offset, oRejectedQty)
                        oParentDB.SetValue("U_WORNo", oParentDB.Offset, oPONo)
                        GLMethod(oGLMethod)
                        If oParentDB.GetValue("U_Pnordno", oParentDB.Offset).Trim.Length > 0 Then
                            oForm.Items.Item("txtscode").Enabled = True
                            oForm.Items.Item("btnscode").Enabled = True
                            oForm.Items.Item("cmbopcd").Enabled = True
                            LoadOperationCombo()
                        End If
                    End If
                End If
                '******************************************Shift Details***************************************************
                If (ControlName = "btnscode" Or ControlName = "txtscode") And (ChoosefromListUID = "SftBtnLst" Or ChoosefromListUID = "SftLst") Then
                    If Not oDataTable Is Nothing Then
                        oShiftCode = oDataTable.GetValue("Code", 0)
                        oShiftDesc = oDataTable.GetValue("U_Sdescr", 0)
                        oShiftFromTime = oDataTable.GetValue("U_Sftime", 0)
                        oShiftToTime = oDataTable.GetValue("U_Sttime", 0)
                        Dim intFrmTime, intToTime As Integer
                        intFrmTime = CInt(oShiftFromTime)
                        intToTime = CInt(oShiftToTime)

                        oParentDB.Offset = oParentDB.Size - 1
                        oParentDB.SetValue("U_Scode", oParentDB.Offset, oShiftCode)
                        oParentDB.SetValue("U_Sdesc", oParentDB.Offset, oShiftDesc)
                        oSftFromTimeTxt.Value = Format(intFrmTime, "00:00") 'oShiftFromTime.Format("00:00")
                        oSftToTimeTxt.Value = Format(intToTime, "00:00") 'oShiftToTime
                    End If
                End If
                '******************************************Labour Details***********************************************
                If ControlName = "matlab" And ChoosefromListUID = "LabLst" Then
                    oMacMatrix.GetLineData(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                    oMacCodeCombo = oMacCodeCol.Cells.Item(oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific
                    If Not oDataTable Is Nothing Then
                        oLabourCode = oDataTable.GetValue("Code", 0)
                        oSkGroupCode = oDataTable.GetValue("U_LGCode", 0)
                        oSkGroupName = oDataTable.GetValue("U_LGname", 0)
                        oLLabRate = oDataTable.GetValue("U_Labrate", 0)
                        oLAcCode = oDataTable.GetValue("U_Accode", 0)
                        oLActAcCode = oDataTable.GetValue("U_ActAcCode", 0)
                        oRs.DoQuery("Select * from OACT Where FormatCode = '" & oLActAcCode & "'")
                        If oRs.RecordCount > 0 Then
                            oRs.MoveFirst()
                            oAccCode = oRs.Fields.Item("AcctCode").Value
                        End If
                        oLAcName = oDataTable.GetValue("U_Acname", 0)
                        oStrSql = "Select T0.* from [@PSSIT_RTE2] T0 where U_Skilgrp = '" & oSkGroupCode & "' and U_wcno = '" & oMacCodeCombo.Selected.Value & "' and U_OprCode = '" & oOprCodeTxt.Value & "' and U_Rteid = '" & oRteIDTxt.Value & "'"
                        oRs.DoQuery(oStrSql)
                        If oRs.RecordCount > 0 Then
                            oRs.MoveFirst()
                            oLReqNo = oRs.Fields.Item("U_Reqno").Value
                        End If
                        oMacLineID = oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)
                        oMacDocEntry = oMacMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)
                        oMacCode = oMacCodeCombo.Selected.Value
                        If CurrentRow = oLabMatrix.VisualRowCount Then
                            oLabourDB.Offset = oLabourDB.Size - 1
                            SetLabourDefaultValue(oMacLineID, oMacCode, oMacDocEntry)
                            oLabMatrix.SetLineData(CurrentRow)
                            oLabMatrix.FlushToDataSource()
                        End If
                        oLabMatrix.GetLineData(CurrentRow)
                        oLabourDB.SetValue("U_Lrcode", oLabourDB.Offset, oLabourCode)
                        oLabourDB.SetValue("U_LGCode", oLabourDB.Offset, oSkGroupCode)
                        oLabourDB.SetValue("U_LGname", oLabourDB.Offset, oSkGroupName)
                        oLabourDB.SetValue("U_Reqno", oLabourDB.Offset, oLReqNo)
                        oLabourDB.SetValue("U_wcno", oLabourDB.Offset, oMacCode)
                        oLabourDB.SetValue("U_Lrtph", oLabourDB.Offset, oLLabRate)
                        'oLabourDB.SetValue("U_Accode", oLabourDB.Offset, oLAcCode)
                        oLabourDB.SetValue("U_Accode", oLabourDB.Offset, oAccCode)
                        oLabourDB.SetValue("U_Acname", oLabourDB.Offset, oLAcName)
                        oLabMatrix.SetLineData(CurrentRow)
                        oLabMatrix.FlushToDataSource()
                    End If
                End If

                'Load machine CFL ---------- added by kabilahan -b 
                If ControlName = "105" And ChoosefromListUID = "ScrpMacLst" Then
                    If Not oDataTable Is Nothing Then
                        oScrpMatrix.GetLineData(oScrpMatrix.RowCount)
                        oMacCode = oDataTable.GetValue("U_wcno", 0)
                        oScrapDB.SetValue("U_MacCode", oScrapDB.Offset, oMacCode)
                        oScrpMatrix.SetLineData(oScrpMatrix.RowCount)
                    End If
                End If

                'load labour CFL -- added by kabilahan - b
                If ControlName = "105" And ChoosefromListUID = "LabLst" Then
                    If Not oDataTable Is Nothing Then
                        oScrpMatrix.GetLineData(oScrpMatrix.RowCount)
                        oLabCode = oDataTable.GetValue("Code", 0)
                        oScrapDB.SetValue("U_LabCode", oScrapDB.Offset, oLabCode)
                        oScrpMatrix.SetLineData(oScrpMatrix.RowCount)
                    End If
                End If
                'load labour CFG -- added by kabilahan - E
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' Loading the Operation Details in the Operation Combo.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadOperationCombo()
        Dim oStrSql As String
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim IntICount As Integer
        Try
            'oStrSql = "Select T0.U_Baslino,T0.U_OprName from [@PSSIT_WOR2] T0 " _
            '          & "Inner Join OWOR T1 On T1.DocNum  = T0.U_Pordno and T1.Series  = T0.U_Pordser " _
            '          & "Where T0.U_Pordno = " & oPONoTxt.Value & " and T0.U_CloseKey = 'N'"

            'oStrSql = "Select T0.U_Baslino,T0.U_OprCode,T0.U_OprName " _
            '& "From [@PSSIT_WOR2] T0 " _
            '& "Inner Join OWOR T1 On T1.DocNum  = T0.U_Pordno and T1.Series  = T0.U_Pordser " _
            '& "Left Join " _
            '& "(Select T2.U_OprCode, Sum(Isnull(T2.U_PassQty,0)) as PassedQty,Sum(Isnull(T2.U_ScrapQty,0)) as ScrapQty " _
            '& "From [@PSSIT_WOR2] T2 " _
            '& "Inner Join OWOR T3 On T3.DocNum  = T2.U_Pordno and T3.Series  = T2.U_Pordser " _
            '& "Group by T2.U_OprCode " _
            '& ")Tbl On Tbl.U_OprCode = T0.U_OprCode " _
            '& "Where T0.U_Pordno = " & oPONoTxt.Value _
            '& "and (T1.PlannedQty - (Isnull(Tbl.PassedQty,0) + Isnull(Tbl.ScrapQty,0))) > 0 "
            oStrSql = "Select distinct T0.U_POrdno,T0.U_Baslino,T0.U_Parid,T0.U_OprCode,T0.U_OprName,T1.PlannedQty,T0.U_ProdQty, " _
            & "T0.U_RewrkQty,T0.U_ScrapQty,IsNull(T1.PlannedQty + Tbl.AccReworkQty - Tbl1.ProducedQty, 0) as TotalQty " _
            & "From [@PSSIT_WOR2] T0 " _
            & "Inner Join OWOR T1 On T1.DocNum  = T0.U_Pordno and T1.Series  = T0.U_Pordser " _
            & "Inner Join [@PSSIT_OPRN] T2 On T2.Code = T0.U_OprCode " _
            & "Inner Join " _
            & "(Select T0.U_POrdno,T0.U_OprCode,Sum(IsNull(T0.U_RewrkQty,0)) as AccReworkQty From [@PSSIT_WOR2] T0 " _
            & "Inner Join OWOR T1 On T1.DocNum  = T0.U_Pordno and T1.Series  = T0.U_Pordser " _
            & "Group by T0.U_OprCode,T0.U_POrdno)Tbl On Tbl.U_POrdno = T0.U_POrdno and Tbl.U_OprCode = T0.U_OprCode " _
            & "Inner Join " _
            & "(Select T1.U_POrdno,T1.U_OprCode,Sum(IsNull(T1.U_ProdQty,0)) as ProducedQty From [@PSSIT_WOR2] T1 " _
            & "Group By T1.U_OprCode,T1.U_POrdno)Tbl1 On Tbl1.U_POrdno = T0.U_POrdno and Tbl1.U_OprCode = T0.U_OprCode " _
            & "Where  T0.U_POrdno = " & oPONoTxt.Value & " And (T1.PlannedQty + Tbl.AccReworkQty - Tbl1.ProducedQty) > 0 " _
            '& "and T2.U_OprType <> 'SubContract' "
            If oOprCombo.ValidValues.Count > 0 Then
                For IntICount = oOprCombo.ValidValues.Count - 1 To 0 Step -1
                    oOprCombo.ValidValues.Remove(IntICount, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            oRs.DoQuery(oStrSql)
            If oRs.RecordCount > 0 Then
                oRs.MoveFirst()
                While oRs.EoF = False
                    oOprCombo.ValidValues.Add(oRs.Fields.Item("U_Baslino").Value, oRs.Fields.Item("U_OprName").Value)
                    ' oOprCombo.ValidValues.Add(oRs.Fields.Item("U_Parid").Value, oRs.Fields.Item("U_OprName").Value)
                    oRs.MoveNext()
                End While
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oStrSql = Nothing
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub

    ''' <summary>
    ''' loading reason details in the reason combo.
    ''' </summary>
    ''' <param name="aComboBox"></param>
    ''' <remarks></remarks>
    Private Sub LoadReasonCombo(ByVal aComboBox As SAPbouiCOM.ComboBox)
        Dim oStrSql As String, IntICount As Integer
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            oStrSql = "Select * from [@PSSIT_ORES]"
            If aComboBox.ValidValues.Count > 0 Then
                For IntICount = aComboBox.ValidValues.Count - 1 To 0 Step -1
                    aComboBox.ValidValues.Remove(IntICount, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            oRs.DoQuery(oStrSql)
            If oRs.RecordCount > 0 Then
                oRs.MoveFirst()
                While oRs.EoF = False
                    aComboBox.ValidValues.Add(oRs.Fields.Item("Code").Value, oRs.Fields.Item("Name").Value)
                    oRs.MoveNext()
                End While
            End If
            aComboBox.ValidValues.Add("Define New", "")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    'Added by Manimaran--------s
    Private Sub loadReaCombo(ByVal uid As String)
        Dim oStrSql As String, IntICount As Integer
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            oStrSql = "Select * from [@PSSIT_ORES]"
            If uid = "104" Then
                If oRwrkRea.ValidValues.Count > 0 Then
                    For IntICount = oRwrkRea.ValidValues.Count - 1 To 0 Step -1
                        oRwrkRea.ValidValues.Remove(IntICount, SAPbouiCOM.BoSearchKey.psk_Index)
                    Next
                End If
                oRs.DoQuery(oStrSql)
                If oRs.RecordCount > 0 Then
                    oRs.MoveFirst()
                    While oRs.EoF = False
                        oRwrkRea.ValidValues.Add(oRs.Fields.Item("Code").Value, oRs.Fields.Item("Name").Value)
                        oRs.MoveNext()
                    End While
                End If
                oRwrkRea.ValidValues.Add("Define New", "")
            Else
                If oScrpRea.ValidValues.Count > 0 Then
                    For IntICount = oScrpRea.ValidValues.Count - 1 To 0 Step -1
                        oScrpRea.ValidValues.Remove(IntICount, SAPbouiCOM.BoSearchKey.psk_Index)
                    Next
                End If
                oRs.DoQuery(oStrSql)
                If oRs.RecordCount > 0 Then
                    oRs.MoveFirst()
                    While oRs.EoF = False
                        oScrpRea.ValidValues.Add(oRs.Fields.Item("Code").Value, oRs.Fields.Item("Name").Value)
                        oRs.MoveNext()
                    End While
                End If
                oScrpRea.ValidValues.Add("Define New", "")

            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub LoadCombo(ByVal oCombo As SAPbouiCOM.Column)
        Dim StrSql As String, IntICount As Integer
        Dim Rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            StrSql = "select U_Empnam,U_LGCode from [@PSSIT_olbr] "

            'If oLabNameCol.ValidValues.Count > 0 Then
            '    For IntICount = oLabNameCol.ValidValues.Count - 1 To 0 Step -1
            '        oLabNameCol.ValidValues.Remove(IntICount, SAPbouiCOM.BoSearchKey.psk_Index)
            '    Next
            'End If

            Rs.DoQuery(StrSql)
            If Rs.RecordCount > 0 Then
                While Not Rs.EoF
                    oCombo.ValidValues.Add(Rs.Fields.Item("0").Value, Rs.Fields.Item("0").Value)
                    Rs.MoveNext()
                End While
            End If
        Catch ex As Exception
            Throw ex
        Finally
            Rs = Nothing
            GC.Collect()
        End Try
    End Sub
    'Added by Manimaran--------e
    ''' <summary>
    ''' Loading machine details in the machine combo.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadMachineCombo()
        Dim oStrSql As String
        Dim IntICount As Integer
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            oStrSql = "Select T0.U_WcNo,T0.U_WCName from [@PSSIT_RTE1] T0 " _
            & "Inner Join [@PSSIT_RTE4] T1 On T1.Code = T0.U_Rteid " _
            & "Where T0.U_OprCode = '" & oOprCodeTxt.Value & "' Group by T0.U_WcNo,T0.U_WCName"
            If oMacCodeCol.ValidValues.Count > 0 Then
                For IntICount = oMacCodeCol.ValidValues.Count - 1 To 0 Step -1
                    oMacCodeCol.ValidValues.Remove(IntICount, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            oRs.DoQuery(oStrSql)
            If oRs.RecordCount > 0 Then
                oRs.MoveFirst()
                While oRs.EoF = False
                    oMacCodeCol.ValidValues.Add(oRs.Fields.Item("U_WcNo").Value, oRs.Fields.Item("U_WcName").Value)
                    oRs.MoveNext()
                End While
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' Loading Machine Time Details in the machine type combo.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadMachineTypeCombo()
        Dim IntICount As Integer
        Try
            If oMTypeCol.ValidValues.Count > 0 Then
                For IntICount = oMTypeCol.ValidValues.Count - 1 To 0 Step -1
                    oMTypeCol.ValidValues.Remove(IntICount, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            oMTypeCol.ValidValues.Add("Operation Time", "")
            oMTypeCol.ValidValues.Add("Setup Time", "")
            oMTypeCol.ValidValues.Add("Stoppage Time", "")
            oMTypeCol.ValidValues.Add("Rework Time", "")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Loading machine details in the machine combo.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadToolsCombo(ByVal aMacCode As String)
        Dim oStrSql As String
        Dim IntICount As Integer
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            oStrSql = "Select Distinct T0.*,T2.U_AcCode,T2.U_AcName from [@PSSIT_RTE3] T0 " _
            & "Inner Join [@PSSIT_RTE1] T1 On T1.U_wcno = T0.U_wcno " _
            & "Inner Join [@PSSIT_OTLS] T2 On T2.Code = T0.U_ToolCode " _
            & "Where T0.U_wcNo = '" & aMacCode & "' and T0.U_OprCode = '" & oOprCodeTxt.Value & "' and T0.U_RteID = '" & oRteIDTxt.Value & "'"
            oRs.DoQuery(oStrSql)
            If oToolCodeCol.ValidValues.Count > 0 Then
                For IntICount = oToolCodeCol.ValidValues.Count - 1 To 0 Step -1
                    oToolCodeCol.ValidValues.Remove(IntICount, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If
            oRs.DoQuery(oStrSql)
            If oRs.RecordCount > 0 Then
                oRs.MoveFirst()
                While oRs.EoF = False
                    oToolCodeCol.ValidValues.Add(oRs.Fields.Item("U_Toolcode").Value, oRs.Fields.Item("U_TLname").Value)
                    oRs.MoveNext()
                End While
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' Setting default value to the newly inserting column in the datasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetMachinesValue()
        Try
            oMachinesDB.SetValue("U_wcname", oMachinesDB.Offset, "")
            oMachinesDB.SetValue("U_Wrkno", oMachinesDB.Offset, "")
            oMachinesDB.SetValue("U_Wrkname", oMachinesDB.Offset, "")
            oMachinesDB.SetValue("U_Mopcph", oMachinesDB.Offset, "0.00")
            oMachinesDB.SetValue("U_Mprcph", oMachinesDB.Offset, "0.00")
            oMachinesDB.SetValue("U_Mohcph1", oMachinesDB.Offset, "0.00")
            oMachinesDB.SetValue("U_Mohcph2", oMachinesDB.Offset, "0.00")
            oMachinesDB.SetValue("U_RMopcph", oMachinesDB.Offset, "0.00")
            oMachinesDB.SetValue("U_RMprcph", oMachinesDB.Offset, "0.00")
            oMachinesDB.SetValue("U_RMohcph1", oMachinesDB.Offset, "0.00")
            oMachinesDB.SetValue("U_RMohcph2", oMachinesDB.Offset, "0.00")
            oMachinesDB.SetValue("U_Totcost", oMachinesDB.Offset, "0.00")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Setting default value to the newly inserting column in the datasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetMachinesDefaultValue()
        Try
            oMachinesDB.SetValue("U_type", oMachinesDB.Offset, "Operation Time")
            oMachinesDB.SetValue("U_Frtime", oMachinesDB.Offset, "") 'FormatDateTime(Now(), DateFormat.ShortTime))
            oMachinesDB.SetValue("U_Totime", oMachinesDB.Offset, "") 'FormatDateTime(Now(), DateFormat.ShortTime))
            oMachinesDB.SetValue("U_Rntime", oMachinesDB.Offset, 0)
            oMachinesDB.SetValue("U_Qty", oMachinesDB.Offset, "0.00")
            If oAccKeyCheck.Checked = True Then
                oMachinesDB.SetValue("U_Acckey", oMachinesDB.Offset, "Y")
            ElseIf oAccKeyCheck.Checked = False Then
                oMachinesDB.SetValue("U_Acckey", oMachinesDB.Offset, "N")
            End If
            oMachinesDB.SetValue("U_Accode", oMachinesDB.Offset, "")
            oMachinesDB.SetValue("U_Acname", oMachinesDB.Offset, "")
            oMachinesDB.SetValue("U_CAccode", oMachinesDB.Offset, UConAcCode.Value)
            oMachinesDB.SetValue("U_CAcname", oMachinesDB.Offset, UConAcName.Value)
            oMachinesDB.SetValue("U_Mopcph", oMachinesDB.Offset, "0.00")
            oMachinesDB.SetValue("U_Adnl1", oMachinesDB.Offset, "")
            oMachinesDB.SetValue("U_Adnl2", oMachinesDB.Offset, "")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    'Added by Manimaran--------s
    Private Sub LoadLab(ByVal aCurrentRow As Integer, ByVal aMacCode As String, ByVal aMacDocEntry As Integer)
        Dim oStrSql, TNo As String
        Dim IntICount, i As Integer
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oStrSql = "Select T0.*,T3.U_Accode ,T3.U_Acname,T4.U_LGCode ,t4.U_LGname ,t4.U_Labrate   from [@PSSIT_PRN2] T0 "
            oStrSql = oStrSql + " Inner Join [@PSSIT_OPRN] T2 On T2.Code = T0.Code "
            oStrSql = oStrSql + " inner join [@PSSIT_OLBR] T3 on T0.U_Skilgrp = T3.Code "
            oStrSql = oStrSql + " inner join [@PSSIT_OLBR] T4 on T4.Code = T0.U_Skilgrp "
            oStrSql = oStrSql + " Where T2.U_oprname = '" & oOprCombo.Selected.Description & "' "
            oRs.DoQuery(oStrSql)

            If oRs.RecordCount > 0 Then

                oLabMatrix.AddRow(1, oLabMatrix.RowCount)
                If oLabMatrix.RowCount = 1 Then
                    oLabSerialNo = GenerateSerialNo("PSSIT_PEY2")
                ElseIf oLabMatrix.RowCount > 1 Then
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        oLabSerialNo = oLabSerialNo + 1
                    ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        oLabSerialNo = GenerateSerialNo("PSSIT_PEY2")
                    End If
                End If

                oRs.MoveFirst()
                For IntICount = 0 To oRs.RecordCount - 1
                    If aCurrentRow = oLabMatrix.VisualRowCount Then
                        oLabourDB.Offset = oLabourDB.Size - 1
                        SetLabourDefaultValue(aCurrentRow, aMacCode, aMacDocEntry)
                        oLabMatrix.SetLineData(aCurrentRow)
                        oLabMatrix.FlushToDataSource()
                    End If
                    oLabourDB.SetValue("Code", oLabourDB.Offset, oLabSerialNo)
                    oLabourDB.SetValue("U_Prdentno", oLabourDB.Offset, oParentDB.GetValue("DocNum", oParentDB.Offset).Trim())
                    oLabourDB.SetValue("U_Maclid", oLabourDB.Offset, aCurrentRow)
                    oLabourDB.SetValue("U_Madcey", oLabourDB.Offset, aMacDocEntry)
                    oLabourDB.SetValue("U_wcno", oLabourDB.Offset, aMacCode)
                    oLabourDB.SetValue("U_Lrcode", oLabourDB.Offset, oRs.Fields.Item("U_Skilgrp").Value)
                    oLabourDB.SetValue("U_Accode", oLabourDB.Offset, oRs.Fields.Item("U_Accode").Value)
                    oLabourDB.SetValue("U_Acname", oLabourDB.Offset, oRs.Fields.Item("U_Acname").Value)
                    oLabourDB.SetValue("U_Frtime", oLabourDB.Offset, oShiftFromTime)
                    oLabourDB.SetValue("U_Totime", oLabourDB.Offset, oShiftToTime)
                    oLabourDB.SetValue("U_Qty", oLabourDB.Offset, oForm.Items.Item("txtpdqty").Specific.string)
                    oLabourDB.SetValue("U_LGCode", oLabourDB.Offset, oRs.Fields.Item("U_LGCode").Value)
                    oLabourDB.SetValue("U_LGname", oLabourDB.Offset, oRs.Fields.Item("U_LGname").Value)
                    oLabourDB.SetValue("U_Lrtph", oLabourDB.Offset, oRs.Fields.Item("U_Labrate").Value)

                    oLabMatrix.SetLineData(oLabMatrix.RowCount)
                    If IntICount <> oRs.RecordCount - 1 Then
                        oLabourDB.InsertRecord(oLabourDB.Size)
                        oLabourDB.Offset = oLabourDB.Size - 1
                        SetToolsDefaultValue(aMacCode, aCurrentRow, aMacDocEntry)
                        oLabMatrix.AddRow(1, oToolsMatrix.RowCount)
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
    End Sub
    'Added by Manimaran--------e
    ''' <summary>
    ''' Loading Tools Details based on the machines from the Routing Operations.
    ''' </summary>
    ''' <param name="aCurrentRow"></param>
    ''' <param name="aMacCode"></param>
    ''' <param name="aMacDocEntry"></param>
    ''' <remarks></remarks>
    Private Sub LoadToolsData(ByVal aCurrentRow As Integer, ByVal aMacCode As String, ByVal aMacDocEntry As Integer)
        Dim oStrSql, TNo As String
        Dim IntICount, i As Integer
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oToolsMatrix.RowCount > 0 Then
                For i = 1 To oToolsMatrix.RowCount
                    If TNo = "" Then
                        TNo = "'" & "" & oToolsMatrix.Columns.Item("coltlcode").Cells.Item(i).Specific.string & "" & "'"
                    Else
                        TNo = TNo + "," + "'" & "" & oToolsMatrix.Columns.Item("coltlcode").Cells.Item(i).Specific.string & "" & "'"
                    End If
                Next
            End If
            If TNo = "" Then
                oStrSql = "Select Distinct T0.*,T2.U_AcCode,T2.U_AcName,T2.U_Cpno,T3.AcctCode,T2.U_TypOfItm,T0.U_Adnl1,T0.U_Adnl2 from [@PSSIT_RTE3] T0 " _
                & "Inner Join [@PSSIT_RTE1] T1 On T1.U_wcno = T0.U_wcno " _
                & "Inner Join [@PSSIT_OTLS] T2 On T2.Code = T0.U_ToolCode " _
                & "Inner Join OACT T3 On T3.FormatCode = T2.U_ActAcCode " _
                & "Where T0.U_wcNo = '" & aMacCode & "' and T0.U_OprCode = '" & oOprCodeTxt.Value & "' and T0.U_RteID = '" & oRteIDTxt.Value & "'"
            Else
                oStrSql = "Select Distinct T0.*,T2.U_AcCode,T2.U_AcName,T2.U_Cpno,T3.AcctCode,T2.U_TypOfItm,T0.U_Adnl1,T0.U_Adnl2 from [@PSSIT_RTE3] T0 " _
                & "Inner Join [@PSSIT_RTE1] T1 On T1.U_wcno = T0.U_wcno " _
                & "Inner Join [@PSSIT_OTLS] T2 On T2.Code = T0.U_ToolCode " _
                & "Inner Join OACT T3 On T3.FormatCode = T2.U_ActAcCode " _
                & "Where T0.U_wcNo = '" & aMacCode & "' and T0.U_OprCode = '" & oOprCodeTxt.Value & "' and T0.U_RteID = '" & oRteIDTxt.Value & "' and U_Toolcode not in (" & TNo & ")"
            End If
            oRs.DoQuery(oStrSql)

            If oRs.RecordCount > 0 Then

                oToolsMatrix.AddRow(1, oToolsMatrix.RowCount)
                If oToolsMatrix.RowCount = 1 Then
                    oToolsSerialNo = GenerateSerialNo("PSSIT_PEY3")
                ElseIf oToolsMatrix.RowCount > 1 Then
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        oToolsSerialNo = oToolsSerialNo + 1
                    ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        oToolsSerialNo = GenerateSerialNo("PSSIT_PEY3")
                    End If
                End If

                oRs.MoveFirst()
                For IntICount = 0 To oRs.RecordCount - 1
                    If aCurrentRow = oToolsMatrix.VisualRowCount Then
                        oToolsDB.Offset = oToolsDB.Size - 1
                        SetToolsDefaultValue(aMacCode, aCurrentRow, aMacDocEntry)
                        oToolsMatrix.SetLineData(aCurrentRow)
                        oToolsMatrix.FlushToDataSource()
                    End If
                    oToolsDB.SetValue("Code", oToolsDB.Offset, oToolsSerialNo)
                    oToolsDB.SetValue("U_Prdentno", oToolsDB.Offset, oParentDB.GetValue("DocNum", oParentDB.Offset).Trim())
                    oToolsDB.SetValue("U_Maclid", oToolsDB.Offset, aCurrentRow)
                    oToolsDB.SetValue("U_Madcey", oToolsDB.Offset, aMacDocEntry)
                    oToolsDB.SetValue("U_wcno", oToolsDB.Offset, aMacCode)
                    oToolsDB.SetValue("U_Toolcode", oToolsDB.Offset, oRs.Fields.Item("U_Toolcode").Value)
                    oToolsDB.SetValue("U_TLname", oToolsDB.Offset, oRs.Fields.Item("U_TLname").Value)

                    'Modified by Manimaran------s
                    'oToolsDB.SetValue("U_Qty", oToolsDB.Offset, "0.00")
                    oToolsDB.SetValue("U_Qty", oToolsDB.Offset, oForm.Items.Item("txtpdqty").Specific.string)
                    'Modified by Manimaran------e
                    If oAccKeyCheck.Checked = True Then
                        oToolsDB.SetValue("U_Acckey", oToolsDB.Offset, "Y")
                    ElseIf oAccKeyCheck.Checked = False Then
                        oToolsDB.SetValue("U_Acckey", oToolsDB.Offset, "N")
                    End If
                    'oToolsDB.SetValue("U_Accode", oToolsDB.Offset, oRs.Fields.Item("U_AcCode").Value)
                    oToolsDB.SetValue("U_Accode", oToolsDB.Offset, oRs.Fields.Item("AcctCode").Value)
                    oToolsDB.SetValue("U_Acname", oToolsDB.Offset, oRs.Fields.Item("U_AcName").Value)
                    oToolsDB.SetValue("U_CAccode", oToolsDB.Offset, UConAcCode.Value)
                    oToolsDB.SetValue("U_CAcname", oToolsDB.Offset, UConAcName.Value)
                    oToolsDB.SetValue("U_Tlctppie", oToolsDB.Offset, oRs.Fields.Item("U_Cpno").Value)
                    oToolsDB.SetValue("U_Totcost", oToolsDB.Offset, RunToolTotalCost(oToolsMatrix.RowCount, oRs.Fields.Item("U_Cpno").Value, CDbl(oToolsDB.GetValue("U_Qty", oToolsDB.Offset).Trim())))
                    'oToolsDB.SetValue("U_Adnl1", oToolsDB.Offset, "")
                    'oToolsDB.SetValue("U_Adnl2", oToolsDB.Offset, "")
                    oToolsDB.SetValue("U_Adnl1", oToolsDB.Offset, oRs.Fields.Item("U_Adnl1").Value)
                    oToolsDB.SetValue("U_Adnl2", oToolsDB.Offset, oRs.Fields.Item("U_Adnl1").Value)


                    oToolsMatrix.SetLineData(oToolsMatrix.RowCount)
                    If IntICount <> oRs.RecordCount - 1 Then
                        oToolsDB.InsertRecord(oToolsDB.Size)
                        oToolsDB.Offset = oToolsDB.Size - 1
                        SetToolsDefaultValue(aMacCode, aCurrentRow, aMacDocEntry)
                        oToolsMatrix.AddRow(1, oToolsMatrix.RowCount)
                        oToolsSerialNo = oToolsSerialNo + 1
                    End If
                    'Added by Manimaran------s
                    Typofitm = oRs.Fields.Item("U_TypOfItm").Value
                    'Added by Manimaran------e
                    oRs.MoveNext()
                Next
            End If
            oParentDB.SetValue("U_Tottcst", oParentDB.Offset, TotalToolCost())
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        Finally
            oRs = Nothing
            oStrSql = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' Loading Tools Details based on the machines from the Routing Operations.
    ''' </summary>
    ''' <param name="aCurrentRow"></param>
    ''' <param name="aMacCode"></param>
    ''' <param name="aWCCode"></param>
    ''' <remarks></remarks>
    Private Sub LoadFixedCostData(ByVal aCurrentRow As Integer, ByVal aMacCode As String, ByVal aWCCode As String)
        Dim oStrSql As String
        Dim IntICount As Integer
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oStrSql = "Select T0.Code,T0.U_WCname,T0.U_WCType,T1.U_FCost,T1.U_UnitCost,T1.U_Absmthd,T1.U_Accode,T1.U_Acname,T1.U_ActAcCode,T2.AcctCode " _
            & "from [@PSSIT_WCR1] T1 " _
            & "Inner Join [@PSSIT_OWCR] T0 On T1.Code = T0.Code " _
            & "Inner Join OACT T2 On T2.FormatCode = T1.U_ActAcCode " _
            & "Where T0.Code = '" & aWCCode & "'"
            oRs.DoQuery(oStrSql)
            If oRs.RecordCount > 0 Then
                oFCMatrix.AddRow(1, oFCMatrix.RowCount)
                If oFCMatrix.RowCount = 1 Then
                    oFCSerialNo = GenerateSerialNo("PSSIT_PEY4")
                ElseIf oFCMatrix.RowCount > 1 Then
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        oFCSerialNo = oFCSerialNo + 1
                    ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        oFCSerialNo = GenerateSerialNo("PSSIT_PEY4")
                    End If
                End If
                oRs.MoveFirst()
                For IntICount = 0 To oRs.RecordCount - 1
                    If aCurrentRow = oFCMatrix.VisualRowCount Then
                        oFixedCostDB.Offset = oFixedCostDB.Size - 1
                        SetFCDefaultValue(aMacCode, aCurrentRow, aWCCode)
                        oFCMatrix.SetLineData(aCurrentRow)
                        oFCMatrix.FlushToDataSource()
                    End If
                    oFixedCostDB.SetValue("Code", oFixedCostDB.Offset, oFCSerialNo)
                    oFixedCostDB.SetValue("U_Pordser", oFixedCostDB.Offset, oPOSeriesTxt.Value)
                    oFixedCostDB.SetValue("U_Pordno", oFixedCostDB.Offset, oPONoTxt.Value)
                    oFixedCostDB.SetValue("U_Prdentno", oFixedCostDB.Offset, oPENoTxt.Value)
                    oFixedCostDB.SetValue("U_wcno", oFixedCostDB.Offset, aMacCode)
                    oFixedCostDB.SetValue("U_Wrkno", oFixedCostDB.Offset, aWCCode)
                    oFixedCostDB.SetValue("U_FCost", oFixedCostDB.Offset, oRs.Fields.Item("U_Fcost").Value)
                    oFixedCostDB.SetValue("U_UnitCost", oFixedCostDB.Offset, oRs.Fields.Item("U_UnitCost").Value)
                    oFixedCostDB.SetValue("U_Absmthd", oFixedCostDB.Offset, oRs.Fields.Item("U_Absmthd").Value)
                    oFixedCostDB.SetValue("U_Accode", oFixedCostDB.Offset, oRs.Fields.Item("AcctCode").Value)
                    oFixedCostDB.SetValue("U_Acname", oFixedCostDB.Offset, oRs.Fields.Item("U_Acname").Value)
                    oFixedCostDB.SetValue("U_Totfcst", oFixedCostDB.Offset, "0.00")
                    oFCMatrix.SetLineData(oFCMatrix.RowCount)
                    oFCMatrix.GetLineData(oFCMatrix.RowCount)
                    oFixedCostDB.SetValue("U_Totfcst", oFixedCostDB.Offset, (CDbl(oFixedCostDB.GetValue("U_UnitCost", oFixedCostDB.Offset).Trim()) * CDbl(oProdQtyTxt.Value)))
                    oFCMatrix.SetLineData(oFCMatrix.RowCount)
                    oFCMatrix.GetLineData(oFCMatrix.RowCount)
                    UTotFCost.Value = CDbl(UTotFCost.Value) + CDbl(oFixedCostDB.GetValue("U_Totfcst", oFixedCostDB.Offset).Trim())
                    If IntICount <> oRs.RecordCount - 1 Then
                        oFixedCostDB.InsertRecord(oFixedCostDB.Size)
                        oFixedCostDB.Offset = oFixedCostDB.Size - 1
                        SetFCDefaultValue(aMacCode, aCurrentRow, aMacCode)
                        oFCMatrix.AddRow(1, oFCMatrix.RowCount)
                        oFCSerialNo = oFCSerialNo + 1
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
    End Sub
    ''' <summary>
    ''' This method is used to set default value to the newly inserting column in the datasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetToolsDefaultValue(ByVal aMacCode As String, ByVal aMacLineId As Integer, ByVal aMacDocEntry As Integer)
        Try
            oToolsDB.SetValue("U_Prdentno", oToolsDB.Offset, oParentDB.GetValue("DocNum", oParentDB.Offset).Trim())
            oToolsDB.SetValue("U_Maclid", oToolsDB.Offset, aMacLineId)
            oToolsDB.SetValue("U_Madcey", oToolsDB.Offset, aMacDocEntry)
            oToolsDB.SetValue("U_wcno", oToolsDB.Offset, aMacCode)
            oToolsDB.SetValue("U_Toolcode", oToolsDB.Offset, "")
            oToolsDB.SetValue("U_TLname", oToolsDB.Offset, "")
            oToolsDB.SetValue("U_Qty", oToolsDB.Offset, "0.00")
            If oAccKeyCheck.Checked = True Then
                oToolsDB.SetValue("U_Acckey", oToolsDB.Offset, "Y")
            ElseIf oAccKeyCheck.Checked = False Then
                oToolsDB.SetValue("U_Acckey", oToolsDB.Offset, "N")
            End If
            oToolsDB.SetValue("U_Accode", oToolsDB.Offset, "")
            oToolsDB.SetValue("U_Acname", oToolsDB.Offset, "")
            oToolsDB.SetValue("U_CAccode", oToolsDB.Offset, UConAcCode.Value)
            oToolsDB.SetValue("U_CAcname", oToolsDB.Offset, UConAcName.Value)
            oToolsDB.SetValue("U_Tlctppie", oToolsDB.Offset, "0.00")
            oToolsDB.SetValue("U_Totcost", oToolsDB.Offset, "0.00")
            oToolsDB.SetValue("U_Adnl1", oToolsDB.Offset, "")
            oToolsDB.SetValue("U_Adnl2", oToolsDB.Offset, "")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This method is used to set default value to the newly inserting column in the datasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetFCDefaultValue(ByVal aMacCode As String, ByVal aMacLineId As Integer, ByVal aWcCode As String)
        Try
            oFixedCostDB.SetValue("U_PordSer", oFixedCostDB.Offset, oPOSeriesTxt.Value)
            oFixedCostDB.SetValue("U_Pordno", oFixedCostDB.Offset, oPONoTxt.Value)
            oFixedCostDB.SetValue("U_Prdentno", oFixedCostDB.Offset, oPENoTxt.Value)
            oFixedCostDB.SetValue("U_wcno", oFixedCostDB.Offset, aMacCode)
            oFixedCostDB.SetValue("U_Wrkno", oFixedCostDB.Offset, aWcCode)
            oFixedCostDB.SetValue("U_Fcost", oFixedCostDB.Offset, "")
            oFixedCostDB.SetValue("U_UnitCost", oFixedCostDB.Offset, "0.00")
            oFixedCostDB.SetValue("U_Absmthd", oFixedCostDB.Offset, "")
            oFixedCostDB.SetValue("U_Accode", oFixedCostDB.Offset, "")
            oFixedCostDB.SetValue("U_Acname", oFixedCostDB.Offset, "")
            oFixedCostDB.SetValue("U_Totfcst", oFixedCostDB.Offset, "0.00")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' Setting default value to the newly inserting column in the datasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetLabourDefaultValue(ByVal aMacLineId As Integer, ByVal aMacCode As String, ByVal aMacDocEntry As Integer)
        Dim oStrSql As String
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            oLabourDB.SetValue("U_Prdentno", oLabourDB.Offset, oParentDB.GetValue("DocNum", oParentDB.Offset).Trim())
            oLabourDB.SetValue("U_Maclid", oLabourDB.Offset, aMacLineId)
            oLabourDB.SetValue("U_Madcey", oLabourDB.Offset, aMacDocEntry)
            oLabourDB.SetValue("U_wcno", oLabourDB.Offset, aMacCode)
            oLabourDB.SetValue("U_Lrcode", oLabourDB.Offset, "")
            oLabourDB.SetValue("U_LGCode", oLabourDB.Offset, "")
            oLabourDB.SetValue("U_LGname", oLabourDB.Offset, "")
            oLabourDB.SetValue("U_Reqno", oLabourDB.Offset, "0")
            'Added by Manimaran-------s
            oLabourDB.SetValue("U_Nop", oLabourDB.Offset, "0")
            oLabourDB.SetValue("U_OTtime", oLabourDB.Offset, "0")
            'Added by Manimaran-------e
            oStrSql = "Select * from [@PSSIT_OCON]"
            oRs.DoQuery(oStrSql)
            If oRs.RecordCount > 0 Then
                oRs.MoveFirst()
                If oRs.Fields.Item("U_AccKey").Value = "Y" Or oRs.Fields.Item("U_AccKey").Value = "y" Then
                    oLabourDB.SetValue("U_Labkey", oLabourDB.Offset, "Y")
                ElseIf oRs.Fields.Item("U_AccKey").Value = "N" Or oRs.Fields.Item("U_AccKey").Value = "n" Then
                    oLabourDB.SetValue("U_Labkey", oLabourDB.Offset, "N")
                End If
            End If
            oLabourDB.SetValue("U_Parallel", oLabourDB.Offset, "")
            oLabourDB.SetValue("U_Frtime", oLabourDB.Offset, "") 'FormatDateTime(Now(), DateFormat.ShortTime))
            oLabourDB.SetValue("U_Totime", oLabourDB.Offset, "") 'FormatDateTime(Now(), DateFormat.ShortTime))
            oLabourDB.SetValue("U_Wrktime", oLabourDB.Offset, "0")
            'Added by Manimaran---------s
            Dim i As Integer
            Dim Tqty As Double
            If oLabMatrix.RowCount = 0 Then
                oLabourDB.SetValue("U_Qty", oLabourDB.Offset, "0.00")
            Else
                For i = 1 To oLabMatrix.RowCount
                    If Tqty = 0 Then
                        Tqty = CDbl(oLabourDB.GetValue("U_Qty", oLabourDB.Offset))
                    Else
                        Tqty = Tqty + CDbl(oLabourDB.GetValue("U_Qty", oLabourDB.Offset))
                    End If
                Next
                oLabourDB.SetValue("U_Qty", oLabourDB.Offset, CStr(CDbl(oForm.Items.Item("txtpdqty").Specific.string) - Tqty))
            End If
            'Added by Manimaran---------e
            If oAccKeyCheck.Checked = True Then
                oLabourDB.SetValue("U_Acckey", oLabourDB.Offset, "Y")
            ElseIf oAccKeyCheck.Checked = False Then
                oLabourDB.SetValue("U_Acckey", oLabourDB.Offset, "N")
            End If
            oLabourDB.SetValue("U_Accode", oLabourDB.Offset, "")
            oLabourDB.SetValue("U_Acname", oLabourDB.Offset, "")
            oLabourDB.SetValue("U_CAccode", oLabourDB.Offset, UConAcCode.Value)
            oLabourDB.SetValue("U_CAcname", oLabourDB.Offset, UConAcName.Value)
            oLabourDB.SetValue("U_Lrtph", oLabourDB.Offset, "0.00")
            oLabourDB.SetValue("U_Totcost", oLabourDB.Offset, "0.00")
            oLabourDB.SetValue("U_Adnl1", oLabourDB.Offset, "")
            oLabourDB.SetValue("U_Adnl2", oLabourDB.Offset, "")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This method is used to add a row in the ToolsDB.
    ''' </summary>
    ''' <param name="aMacCode"></param>
    ''' <remarks></remarks>
    Private Sub AddToolsRow(ByVal aMacLineID As Integer, ByVal aMacCode As String, ByVal aMacDocEntry As Integer)
        Try
            oToolsDB.Offset = oToolsDB.Size - 1
            SetToolsDefaultValue(aMacCode, aMacLineID, aMacDocEntry)
            oToolsMatrix.AddRow(1, oToolsMatrix.RowCount)
            If oToolsMatrix.RowCount = 1 Then
                oToolsSerialNo = GenerateSerialNo("PSSIT_PEY3")
            ElseIf oToolsMatrix.RowCount > 1 Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    oToolsSerialNo = oToolsSerialNo + 1
                ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    oToolsSerialNo = GenerateSerialNo("PSSIT_PEY3")
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This method is used to add a row in the ToolsDB.
    ''' </summary>
    ''' <param name="aMacCode"></param>
    ''' <remarks></remarks>
    Private Sub AddLabourRow(ByVal aMacLineID As Integer, ByVal aMacCode As String, ByVal aMacDocEntry As Integer)
        Try
            oLabourDB.Offset = oLabourDB.Size - 1
            SetLabourDefaultValue(aMacLineID, aMacCode, aMacDocEntry)
            oLabMatrix.AddRow(1, oLabMatrix.RowCount)
            If oLabMatrix.RowCount = 1 Then
                oLabSerialNo = GenerateSerialNo("PSSIT_PEY2")
            ElseIf oLabMatrix.RowCount > 1 Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    oLabSerialNo = oLabSerialNo + 1
                ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    oLabSerialNo = GenerateSerialNo("PSSIT_PEY2")
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Loading the tools from the database as per the conditions.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadToolsDataFromDB()
        Dim oConditions As SAPbouiCOM.Conditions
        Dim oCondition As SAPbouiCOM.Condition
        Try
            oConditions = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions)
            oCondition = oConditions.Add
            oCondition.Alias = "U_Prdentno"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = oPENoTxt.Value
            oToolsDB.Query(oConditions)
            oToolsMatrix.LoadFromDataSource()
            oToolsMatrix.FlushToDataSource()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Loading the Labour from the database as per the conditions.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadLabourDataFromDB()
        Dim oConditions As SAPbouiCOM.Conditions
        Dim oCondition As SAPbouiCOM.Condition
        Try
            oConditions = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions)
            oCondition = oConditions.Add
            oCondition.Alias = "U_Prdentno"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = oPENoTxt.Value
            oLabourDB.Query(oConditions)
            oLabMatrix.LoadFromDataSource()
            oLabMatrix.FlushToDataSource()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Loading the Labour from the database as per the conditions.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadFCDataFromDB()
        Dim oConditions As SAPbouiCOM.Conditions
        Dim oCondition As SAPbouiCOM.Condition
        Try
            oConditions = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions)
            oCondition = oConditions.Add
            oCondition.Alias = "U_Prdentno"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = oPENoTxt.Value
            oFixedCostDB.Query(oConditions)
            oFCMatrix.LoadFromDataSource()
            oFCMatrix.FlushToDataSource()
            CalculateTotalFixedCost()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub LoadMachineDataFromDB()
        Dim oConditions As SAPbouiCOM.Conditions
        Dim oCondition As SAPbouiCOM.Condition
        Try
            oConditions = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions)
            oCondition = oConditions.Add
            'oCondition.Alias = "U_Prdentno"
            oCondition.Alias = "DocEntry"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = oPENoTxt.Value
            oMachinesDB.Query(oConditions)
            oMacMatrix.LoadFromDataSource()
            oMacMatrix.FlushToDataSource()
            CalculateTotalFixedCost()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub CalculateTotalFixedCost()
        Dim oTotFCostEdit As SAPbouiCOM.EditText
        Dim oTotFCost As Double = 0
        Try
            For IntICount As Integer = 1 To oFCMatrix.RowCount
                oTotFCostEdit = oFTotCostCol.Cells.Item(IntICount).Specific
                oFCMatrix.GetLineData(IntICount)
                oTotFCost = oTotFCost + CDbl(oTotFCostEdit.Value)
            Next
            UTotFCost.Value = oTotFCost
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    'Added by Manimaran-------s
    Private Sub ReworkDeleteEmptyRow()
        Dim oRwkqty As SAPbouiCOM.EditText
        Dim IntICount, RowCnt As Integer
        RowCnt = oReWrkMatrix.VisualRowCount
        Try
            For IntICount = 1 To RowCnt
                oReWrkMatrix.GetLineData(IntICount)
                oRwkqty = oRwrkQty.Cells.Item(IntICount).Specific
                If oRwkqty.Value = 0 Then
                    oReWrkMatrix.DeleteRow(IntICount)
                    oReWrkMatrix.FlushToDataSource()
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub ScrapDeleteEmptyRow()
        Dim oScpqty As SAPbouiCOM.EditText
        Dim IntICount, RowCnt As Integer
        RowCnt = oScrpMatrix.VisualRowCount
        Try
            For IntICount = 1 To RowCnt
                oScrpMatrix.GetLineData(IntICount)
                oScpqty = oScrpQty.Cells.Item(IntICount).Specific
                If oScpqty.Value = 0 Then
                    oScrpMatrix.DeleteRow(IntICount)
                    oScrpMatrix.FlushToDataSource()
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    'Added by Manimaran-------e
    ''' <summary>
    ''' This method is used to delete the empty rows in the Machine Matrix.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub MachinesDeleteEmptyRow()
        Dim oMacNameEdit As SAPbouiCOM.EditText
        ' Dim oMacCodeCombo As SAPbouiCOM.ComboBox
        Dim IntICount As Integer
        Try
            For IntICount = oMacMatrix.RowCount To 1 Step -1
                oMacMatrix.GetLineData(IntICount)
                'oMacCodeCombo = oMacCodeCol.Cells.Item(IntICount).Specific
                oMacNameEdit = oMacNameCol.Cells.Item(IntICount).Specific
                If oMacNameEdit.Value.Length = 0 Then
                    oMacMatrix.DeleteRow(IntICount)
                    oMacMatrix.FlushToDataSource()
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
        Dim oToolNameEdit As SAPbouiCOM.EditText
        Dim IntICount As Integer
        Try
            For IntICount = 1 To oToolsMatrix.VisualRowCount
                oToolsMatrix.GetLineData(IntICount)
                'oToolCodeEdit = oToolCodeCol.Cells.Item(IntICount).Specific
                oToolNameEdit = oToolDescCol.Cells.Item(IntICount).Specific
                If oToolNameEdit.Value.Length = 0 Then
                    oToolsMatrix.DeleteRow(IntICount)
                    oToolsMatrix.FlushToDataSource()
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
        Dim oSKCodeEdit As SAPbouiCOM.ComboBox
        Dim IntICount As Integer
        Try
            'For IntICount = 1 To oLabMatrix.VisualRowCount
            '    oLabMatrix.GetLineData(IntICount)
            '    oLabCodeEdit = oLabCodeCol.Cells.Item(IntICount).Specific
            '    If oLabCodeEdit.Value.Length = 0 Then
            '        oLabMatrix.DeleteRow(IntICount)
            '        oLabMatrix.FlushToDataSource()
            '    End If
            'Next
            'Added by Manimaran---------s
            For IntICount = 1 To oLabMatrix.VisualRowCount
                oLabMatrix.GetLineData(IntICount)
                oSKCodeEdit = oLabMatrix.Columns.Item("collgcod").Cells.Item(IntICount).Specific
                If oSKCodeEdit.Value.Length = 0 Then
                    oLabMatrix.DeleteRow(IntICount)
                    oLabMatrix.FlushToDataSource()
                End If
            Next
            'Added by Manimaran---------e
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    'Calculating the Produced Qty.
    'Modified by Manimaran--------s
    Private Function ProducedQtyCalculation() As Double
        Dim oTotProducedQty As Double
        Dim k As Integer
        TotRwrkQty = 0
        TotScrpQty = 0
        Try
            'oTotProducedQty = CDbl(oParentDB.GetValue("U_Passqty", oParentDB.Offset).Trim()) + CDbl(oParentDB.GetValue("U_Rewrkqty", oParentDB.Offset).Trim()) + CDbl(oParentDB.GetValue("U_scrapqty", oParentDB.Offset).Trim())
            For k = 1 To oReWrkMatrix.RowCount
                TotRwrkQty = TotRwrkQty + CDbl(oReWrkMatrix.Columns.Item("V_1").Cells.Item(k).Specific.string)
            Next

            For k = 1 To oScrpMatrix.RowCount
                TotScrpQty = TotScrpQty + CDbl(oScrpMatrix.Columns.Item("V_1").Cells.Item(k).Specific.string)
            Next

            oTotProducedQty = CDbl(oForm.Items.Item("txtpsqty").Specific.value) + TotRwrkQty + TotScrpQty
        Catch ex As Exception
            Throw ex
        End Try
        Return oTotProducedQty
    End Function
    'Modified by Manimaran--------e
    ''' <summary>
    ''' Calculating the Produced Qty.
    ''' </summary>
    ''' <remarks></remarks>
    Private Function MacQtyCalculation() As Double
        Dim oTotQty As Double
        Dim IntICount As Integer
        Dim oQtyEdit As SAPbouiCOM.EditText
        Try
            For IntICount = 1 To oMacMatrix.RowCount
                oQtyEdit = oMQtyCol.Cells.Item(IntICount).Specific
                oMacMatrix.GetLineData(IntICount)
                oTotQty = oTotQty + CDbl(oQtyEdit.Value)
            Next
        Catch ex As Exception
            Throw ex
        End Try
        Return oTotQty
    End Function
    ''' <summary>
    ''' Calculating the Produced Qty.
    ''' </summary>
    ''' <remarks></remarks>
    Private Function LabQtyCalculation() As Double
        Dim oTotQty, oParallelQty As Double
        Dim IntICount As Integer
        Dim oQtyEdit As SAPbouiCOM.EditText
        Dim oParallelCheck As SAPbouiCOM.CheckBox
        Dim oNoOfParallelQty As Integer
        Try
            For IntICount = 1 To oLabMatrix.RowCount
                oQtyEdit = oLQtyCol.Cells.Item(IntICount).Specific
                oParallelCheck = oLParCol.Cells.Item(IntICount).Specific
                oLabMatrix.GetLineData(IntICount)
                If oParallelCheck.Checked = True Then
                    oNoOfParallelQty = oNoOfParallelQty + 1
                    oParallelQty = oParallelQty + CDbl(oQtyEdit.Value)
                    oTotQty = oParallelQty / oNoOfParallelQty
                End If
            Next
            For IntICount = 1 To oLabMatrix.RowCount
                oQtyEdit = oLQtyCol.Cells.Item(IntICount).Specific
                oParallelCheck = oLParCol.Cells.Item(IntICount).Specific
                oLabMatrix.GetLineData(IntICount)
                If oParallelCheck.Checked = False Then
                    oTotQty = oTotQty + CDbl(oQtyEdit.Value)
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
        Return oTotQty
    End Function
    ''' <summary>
    ''' Calculating the Produced Qty.
    ''' </summary>
    ''' <remarks></remarks>
    Private Function ToolsQtyCalculation() As Double
        Dim oTotQty As Double
        Dim IntICount As Integer
        Dim oQtyEdit As SAPbouiCOM.EditText
        Try
            For IntICount = 1 To oToolsMatrix.RowCount
                oQtyEdit = oTQtyCol.Cells.Item(IntICount).Specific
                oToolsMatrix.GetLineData(IntICount)
                oTotQty = oTotQty + CDbl(oQtyEdit.Value)
            Next
        Catch ex As Exception
            Throw ex
        End Try
        Return oTotQty
    End Function
    ''' <summary>
    ''' This function is for calculating the duration in Mins.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function DurationMinsCalculation(ByVal oMatFromTime As String, ByVal oMatToTime As String) As String
        Dim oFromTime, oToTime As DateTime
        Dim oDuration As String
        Try
            oFromTime = Convert.ToDateTime(Date.Parse(oMatFromTime))
            oToTime = Convert.ToDateTime(Date.Parse(oMatToTime))
            Dim runLength As System.TimeSpan = oToTime.Subtract(oFromTime.ToShortTimeString)
            Dim secs As Integer = runLength.Seconds
            Dim minutes As Integer = runLength.Minutes
            Dim hours As Integer = runLength.Hours
            oDuration = runLength.Hours * 60 + runLength.Minutes
            'oDuration = runLength.Hours.ToString("00") + ":" + runLength.Minutes.ToString("00")
        Catch ex As Exception
            Throw ex
        End Try
        Return oDuration
    End Function
    ''' <summary>
    ''' Enabling th Items in th form as per the form mode.
    ''' </summary>
    ''' <remarks></remarks>
    ''' 
#Region "Disable all fields"
    Private Sub disable()
        Try


            oForm.Items.Item("txtprdno").Enabled = False
            oForm.Items.Item("cmbseris1").Enabled = False
            oForm.Items.Item("txtsdesc").Enabled = False
            oForm.Items.Item("txtitmcd").Enabled = False
            oForm.Items.Item("txtglmthd").Enabled = False
            oForm.Items.Item("txtitnam").Enabled = False
            oForm.Items.Item("txtwhcod").Enabled = False
            oForm.Items.Item("txtwhnam").Enabled = False
            oForm.Items.Item("txtplqty").Enabled = False
            oForm.Items.Item("txtcmqty").Enabled = False
            oForm.Items.Item("txtrjqty").Enabled = False
            oForm.Items.Item("txtrwqty").Enabled = False
            oForm.Items.Item("txtspqty").Enabled = False
            oForm.Items.Item("txtdocdt").Enabled = False
            oForm.Items.Item("txtpdqty").Enabled = False
            oForm.Items.Item("txtpsqty").Enabled = False
            'Commented by Manimaran-----s
            'oForm.Items.Item("txtoprwqty").Enabled = False
            'oForm.Items.Item("cmboprwres").Enabled = False
            'oForm.Items.Item("txtopscqty").Enabled = False
            'oForm.Items.Item("cmbopscres").Enabled = False
            'Commented by Manimaran-----e
            oForm.Items.Item("txtadnl1").Enabled = False
            oForm.Items.Item("txtadnl2").Enabled = False
            oForm.Items.Item("txtremar").Enabled = False
            oForm.Items.Item("matmac").Enabled = False
            oForm.Items.Item("matlab").Enabled = False
            oForm.Items.Item("mattool").Enabled = False
            oForm.Items.Item("104").Enabled = False 'Added by Manimaran
            oForm.Items.Item("105").Enabled = False 'Added by Manimaran
            oForm.Items.Item("chkackey").Enabled = False
            'oForm.Items.Item("chkrew").Enabled = False
        Catch ex As Exception
            oForm.Freeze(False)
        End Try
    End Sub
#End Region
    Private Sub SetItemEnabled()
        Try
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                oForm.Items.Item("txtprdno").Enabled = True
                oForm.Items.Item("cmbseris1").Enabled = True
                oForm.Items.Item("txtsdesc").Enabled = False
                oForm.Items.Item("txtitmcd").Enabled = False
                oForm.Items.Item("txtglmthd").Enabled = False
                oForm.Items.Item("txtitnam").Enabled = False
                oForm.Items.Item("txtwhcod").Enabled = False
                oForm.Items.Item("txtwhnam").Enabled = False
                oForm.Items.Item("txtplqty").Enabled = False
                oForm.Items.Item("txtcmqty").Enabled = False
                oForm.Items.Item("txtrjqty").Enabled = False
                oForm.Items.Item("txtrwqty").Enabled = False
                oForm.Items.Item("txtspqty").Enabled = False
                oForm.Items.Item("txtdocdt").Enabled = True
                oForm.Items.Item("txtpdqty").Enabled = False
                oForm.Items.Item("txtpsqty").Enabled = False
                'Commented by Manimaran-----s
                'oForm.Items.Item("txtoprwqty").Enabled = False
                'oForm.Items.Item("cmboprwres").Enabled = False
                'oForm.Items.Item("txtopscqty").Enabled = False
                'oForm.Items.Item("cmbopscres").Enabled = False
                'Commented by Manimaran-----e
                oForm.Items.Item("txtadnl1").Enabled = True
                oForm.Items.Item("txtadnl2").Enabled = True
                oForm.Items.Item("txtremar").Enabled = True
                oForm.Items.Item("matmac").Enabled = True
                oForm.Items.Item("matlab").Enabled = True
                oForm.Items.Item("mattool").Enabled = True
                oForm.Items.Item("104").Enabled = True 'Added by Manimaran
                oForm.Items.Item("105").Enabled = True 'Added by Manimaran
                oForm.Items.Item("chkackey").Enabled = False
                'oForm.Items.Item("chkrew").Enabled = True
            ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                oForm.Items.Item("txtseris").Enabled = True
                oForm.Items.Item("txtprdno").Enabled = True
                oForm.Items.Item("txtpordt").Enabled = True
                oForm.Items.Item("txtscode").Enabled = True
                oForm.Items.Item("txtitmcd").Enabled = True
                oForm.Items.Item("txtglmthd").Enabled = True
                oForm.Items.Item("txtitnam").Enabled = True
                oForm.Items.Item("txtwhcod").Enabled = True
                oForm.Items.Item("txtwhnam").Enabled = True
                oForm.Items.Item("txtplqty").Enabled = True
                oForm.Items.Item("txtcmqty").Enabled = True
                oForm.Items.Item("txtrjqty").Enabled = True
                oForm.Items.Item("txtrwqty").Enabled = True
                oForm.Items.Item("txtspqty").Enabled = True
                oForm.Items.Item("cmbseris1").Enabled = True
                oForm.Items.Item("txtpeyno").Enabled = True
                oForm.Items.Item("txtdocdt").Enabled = True
                oForm.Items.Item("txtsdesc").Enabled = True
                oForm.Items.Item("cmbopcd").Enabled = True
                'oForm.Items.Item("chkrew").Enabled = True
                oForm.Items.Item("txtpdqty").Enabled = True
                oForm.Items.Item("txtpsqty").Enabled = True
                'oForm.Items.Item("txtoprwqty").Enabled = True
                'oForm.Items.Item("cmboprwres").Enabled = True
                'oForm.Items.Item("txtopscqty").Enabled = True
                'oForm.Items.Item("cmbopscres").Enabled = True
                oForm.Items.Item("txtadnl1").Enabled = True
                oForm.Items.Item("txtadnl2").Enabled = True
                oForm.Items.Item("txtremar").Enabled = True
                oForm.Items.Item("txtjvno").Enabled = True
                oForm.Items.Item("matmac").Enabled = True
                oForm.Items.Item("matlab").Enabled = True
                oForm.Items.Item("mattool").Enabled = True
                oForm.Items.Item("104").Enabled = True 'Added by Manimaran
                oForm.Items.Item("105").Enabled = True 'Added by Manimaran
                oForm.Items.Item("chkackey").Enabled = True
                'Modified by Manimaran-----s
            ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE And ChkSUser() = False Then
                oForm.Items.Item("txtseris").Enabled = False
                oForm.Items.Item("cmbseris1").Enabled = False
                oForm.Items.Item("txtpordt").Enabled = False
                oForm.Items.Item("txtscode").Enabled = False
                oForm.Items.Item("txtsdesc").Enabled = False
                oForm.Items.Item("txtitmcd").Enabled = False
                oForm.Items.Item("txtglmthd").Enabled = False
                oForm.Items.Item("txtitnam").Enabled = False
                oForm.Items.Item("txtwhcod").Enabled = False
                oForm.Items.Item("txtwhnam").Enabled = False
                oForm.Items.Item("txtplqty").Enabled = False
                oForm.Items.Item("txtcmqty").Enabled = False
                oForm.Items.Item("txtrjqty").Enabled = False
                oForm.Items.Item("txtrwqty").Enabled = False
                oForm.Items.Item("txtspqty").Enabled = False
                oForm.Items.Item("cmbseris1").Enabled = False
                oForm.Items.Item("txtpeyno").Enabled = False
                oForm.Items.Item("txtdocdt").Enabled = False
                'oForm.Items.Item("cmbopcd").Enabled = False
                'oForm.Items.Item("chkrew").Enabled = False
                oForm.Items.Item("chkackey").Enabled = False
                oForm.Items.Item("txtpdqty").Enabled = False
                oForm.Items.Item("txtpsqty").Enabled = False
                'Commented by Manimaran-----s
                'oForm.Items.Item("txtoprwqty").Enabled = False
                'oForm.Items.Item("cmboprwres").Enabled = False
                'oForm.Items.Item("txtopscqty").Enabled = False
                'oForm.Items.Item("cmbopscres").Enabled = False
                'Commented by Manimaran-----e
                oForm.Items.Item("matmac").Enabled = False
                oForm.Items.Item("matlab").Enabled = False
                oForm.Items.Item("mattool").Enabled = False
                'Added by Manimaran----s
                oForm.Items.Item("104").Enabled = False
                oForm.Items.Item("105").Enabled = False
                oForm.Items.Item("txtprdno").Enabled = False
                oForm.Items.Item("cmbseris1").Enabled = False
                oForm.Items.Item("txtdocdt").Enabled = False
                oForm.Items.Item("txtscode").Enabled = False
                'Added by Manimaran----e
                oForm.Items.Item("txtadnl1").Enabled = False
                oForm.Items.Item("txtadnl2").Enabled = False
                oForm.Items.Item("txtjvno").Enabled = False
                oForm.Items.Item("chkackey").Enabled = False

            ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE And ChkSUser() = True Then
                disable()
                oForm.Items.Item("txtprdno").Enabled = True
                oForm.Items.Item("cmbseris1").Enabled = True
                oForm.Items.Item("txtsdesc").Enabled = False
                oForm.Items.Item("txtitmcd").Enabled = False
                oForm.Items.Item("txtglmthd").Enabled = False
                oForm.Items.Item("txtitnam").Enabled = False
                oForm.Items.Item("txtwhcod").Enabled = False
                oForm.Items.Item("txtwhnam").Enabled = False
                oForm.Items.Item("txtplqty").Enabled = False
                oForm.Items.Item("txtcmqty").Enabled = False
                oForm.Items.Item("txtrjqty").Enabled = False
                oForm.Items.Item("txtrwqty").Enabled = False
                oForm.Items.Item("txtspqty").Enabled = False
                oForm.Items.Item("txtdocdt").Enabled = True
                oForm.Items.Item("txtpdqty").Enabled = False
                oForm.Items.Item("txtpsqty").Enabled = False
                oForm.Items.Item("txtadnl1").Enabled = True
                oForm.Items.Item("txtadnl2").Enabled = True
                oForm.Items.Item("txtremar").Enabled = True
                oForm.Items.Item("matmac").Enabled = True
                oForm.Items.Item("matlab").Enabled = True
                oForm.Items.Item("mattool").Enabled = True
                'Added by Manimaran-------s
                oForm.Items.Item("txtscode").Enabled = False
                oForm.Items.Item("104").Enabled = True
                oForm.Items.Item("105").Enabled = True
                oForm.Items.Item("txtprdno").Enabled = False
                oForm.Items.Item("cmbseris1").Enabled = False
                oForm.Items.Item("txtdocdt").Enabled = False
                oForm.Items.Item("txtpdqty").Enabled = False
                oForm.Items.Item("txtpsqty").Enabled = False
                Rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                sQry = "select * from [@PSSIT_OSFT] where code = '" & oForm.Items.Item("txtscode").Specific.string & "'"
                Rs.DoQuery(sQry)
                If Rs.RecordCount > 0 Then
                    oSftFromTimeTxt.Value = Rs.Fields.Item("U_Sftime").Value
                    oSftToTimeTxt.Value = Rs.Fields.Item("U_Sttime").Value

                    Dim intFrmTime, intToTime As Integer
                    intFrmTime = CInt(oSftFromTimeTxt.Value)
                    intToTime = CInt(oSftToTimeTxt.Value)

                    oSftFromTimeTxt.Value = Format(intFrmTime, "00:00") 'oShiftFromTime.Format("00:00")
                    oSftToTimeTxt.Value = Format(intToTime, "00:00") 'oShiftToTime


                End If
                'Added by Manimaran-------e
                oForm.Items.Item("chkackey").Enabled = False
                'oForm.Items.Item("chkrew").Enabled = True
            End If
            'Modified by Manimaran-----s
        Catch ex As Exception
            'Throw ex
            oForm.Freeze(False)
        End Try
    End Sub
    ''' <summary>
    ''' Resizing the form 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Form_Resize()
        Try
            'If oBoolResize = False Then
            '    oForm.Freeze(True)

            '    oForm.Items.Item("rectmac").Height = 110
            '    oForm.Items.Item("rectmac").Left = 5
            '    oForm.Items.Item("rectmac").Width = oForm.Width - 20
            '    oForm.Items.Item("rectmac").Top = oForm.Items.Item("lblmac").Top + oForm.Items.Item("lblmac").Height + 5

            '    oForm.Items.Item("matmac").Height = 100
            '    oForm.Items.Item("matmac").Left = oForm.Items.Item("rectmac").Left + 5
            '    oForm.Items.Item("matmac").Top = oForm.Items.Item("rectmac").Top + 5
            '    oForm.Items.Item("matmac").Width = oForm.Items.Item("rectmac").Width - 10

            '    oForm.Items.Item("foltools").Top = oForm.Items.Item("rectmac").Top + oForm.Items.Item("rectmac").Height + 5
            '    oForm.Items.Item("follabour").Top = oForm.Items.Item("rectmac").Top + oForm.Items.Item("rectmac").Height + 5

            '    oForm.Items.Item("recttool").Left = 5
            '    oForm.Items.Item("recttool").Height = 100
            '    oForm.Items.Item("recttool").Top = oForm.Items.Item("foltools").Top + oForm.Items.Item("foltools").Height
            '    oForm.Items.Item("recttool").Width = oForm.Width - 20

            '    oForm.Items.Item("rectlab").Left = 5
            '    oForm.Items.Item("rectlab").Height = 100
            '    oForm.Items.Item("rectlab").Top = oForm.Items.Item("follabour").Top + oForm.Items.Item("follabour").Height
            '    oForm.Items.Item("rectlab").Width = oForm.Width - 20

            '    oForm.Items.Item("mattool").Left = oForm.Items.Item("recttool").Left + 5
            '    oForm.Items.Item("mattool").Height = oForm.Items.Item("recttool").Height - 10
            '    oForm.Items.Item("mattool").Top = oForm.Items.Item("recttool").Top + 5
            '    oForm.Items.Item("mattool").Width = oForm.Items.Item("recttool").Width - 10

            '    oForm.Items.Item("matlab").Left = oForm.Items.Item("rectlab").Left + 5
            '    oForm.Items.Item("matlab").Height = oForm.Items.Item("rectlab").Height - 10
            '    oForm.Items.Item("matlab").Top = oForm.Items.Item("rectlab").Top + 5
            '    oForm.Items.Item("matlab").Width = oForm.Items.Item("rectlab").Width - 10
            '    oForm.Freeze(False)
            '    oForm.Update()
            '    oBoolResize = True
            'ElseIf oBoolResize = True Then
            '    oForm.Freeze(True)
            '    oForm.Items.Item("rectmac").Height = 105
            '    oForm.Items.Item("rectmac").Left = 5
            '    oForm.Items.Item("rectmac").Width = 590
            '    oForm.Items.Item("rectmac").Top = 217

            '    oForm.Items.Item("matmac").Height = 95
            '    oForm.Items.Item("matmac").Left = 10
            '    oForm.Items.Item("matmac").Top = 222
            '    oForm.Items.Item("matmac").Width = 580

            '    oForm.Items.Item("foltools").Top = 324
            '    oForm.Items.Item("foltools").Left = 5

            '    oForm.Items.Item("follabour").Top = 324
            '    oForm.Items.Item("follabour").Left = 84

            '    oForm.Items.Item("recttool").Left = 5
            '    oForm.Items.Item("recttool").Height = 95
            '    oForm.Items.Item("recttool").Top = 343
            '    oForm.Items.Item("recttool").Width = 590

            '    oForm.Items.Item("rectlab").Left = 5
            '    oForm.Items.Item("rectlab").Height = 95
            '    oForm.Items.Item("rectlab").Top = 343
            '    oForm.Items.Item("rectlab").Width = 590

            '    oForm.Items.Item("mattool").Left = 10
            '    oForm.Items.Item("mattool").Height = 85
            '    oForm.Items.Item("mattool").Top = 348
            '    oForm.Items.Item("mattool").Width = 580

            '    oForm.Items.Item("matlab").Left = 10
            '    oForm.Items.Item("matlab").Height = 85
            '    oForm.Items.Item("matlab").Top = 348
            '    oForm.Items.Item("matlab").Width = 580
            '    oBoolResize = False
            '    oForm.Freeze(False)
            '    oForm.Update()
            'End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Calculating the running machine operation cost.
    ''' </summary>
    ''' <param name="aCurrentRow"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RunMachineOprCostCalculation(ByVal aCurrentRow As Integer) As Double
        Dim oRunOprCost As Double
        Dim oMacOprCost, oRunTime As SAPbouiCOM.EditText
        Try
            oMacMatrix.GetLineData(aCurrentRow)
            oMacOprCost = oMOprCstCol.Cells.Item(aCurrentRow).Specific
            oRunTime = oMRunTimeCol.Cells.Item(aCurrentRow).Specific
            oRunOprCost = CDbl(CDbl(oMacOprCost.Value) / 60) * CInt(oRunTime.Value)
        Catch ex As Exception
            Throw ex
        End Try
        If Typofitm = 0 Then
            Return oRunOprCost
        Else
            Return oRunOprCost / Typofitm
        End If

    End Function
    ''' <summary>
    ''' Calculating the running machine power cost.
    ''' </summary>
    ''' <param name="aCurrentRow"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RunMachinePowCostCalculation(ByVal aCurrentRow As Integer) As Double
        Dim oRunPowerCost As Double
        Dim oMacPowCost, oRunTime As SAPbouiCOM.EditText
        Try
            oMacMatrix.GetLineData(aCurrentRow)
            oMacPowCost = oMPowCstCol.Cells.Item(aCurrentRow).Specific
            oRunTime = oMRunTimeCol.Cells.Item(aCurrentRow).Specific
            oRunPowerCost = CDbl(CDbl(oMacPowCost.Value) / 60) * CInt(oRunTime.Value)
        Catch ex As Exception
            Throw ex
        End Try
        Return oRunPowerCost
    End Function
    ''' <summary>
    ''' Calculating the running machine other cost.
    ''' </summary>
    ''' <param name="aCurrentRow"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RunMachineOtherCost1Calculation(ByVal aCurrentRow As Integer) As Double
        Dim oRunOtherCost1 As Double
        Dim oMacOtherCost1, oRunTime As SAPbouiCOM.EditText
        Try
            oMacMatrix.GetLineData(aCurrentRow)
            oMacOtherCost1 = oMOthCst1Col.Cells.Item(aCurrentRow).Specific
            oRunTime = oMRunTimeCol.Cells.Item(aCurrentRow).Specific
            oRunOtherCost1 = CDbl(CDbl(oMacOtherCost1.Value) / 60) * CInt(oRunTime.Value)
        Catch ex As Exception
            Throw ex
        End Try
        Return oRunOtherCost1
    End Function
    ''' <summary>
    ''' Calculating the running machine other cost.
    ''' </summary>
    ''' <param name="aCurrentRow"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RunMachineOtherCost2Calculation(ByVal aCurrentRow As Integer) As Double
        Dim oRunOtherCost2 As Double
        Dim oMacOtherCost2, oRunTime As SAPbouiCOM.EditText
        Try
            oMacMatrix.GetLineData(aCurrentRow)
            oMacOtherCost2 = oMOthCst2Col.Cells.Item(aCurrentRow).Specific
            oRunTime = oMRunTimeCol.Cells.Item(aCurrentRow).Specific
            oRunOtherCost2 = CDbl(CDbl(oMacOtherCost2.Value) / 60) * CInt(oRunTime.Value)
        Catch ex As Exception
            Throw ex
        End Try
        Return oRunOtherCost2
    End Function
    ''' <summary>
    ''' Calculating the running machine total cost.
    ''' </summary>
    ''' <param name="aCurrentRow"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RunMachineTotalCost(ByVal aCurrentRow As Integer) As Double
        Dim oRunTotMacCost As Double
        Dim oRunOprCost, oRunPowerCost, oRunOtherCost1, oRunOtherCost2 As SAPbouiCOM.EditText
        Try
            oRunOprCost = oMRunOprCstCol.Cells.Item(aCurrentRow).Specific
            oRunPowerCost = oMRunPowCstCol.Cells.Item(aCurrentRow).Specific
            oRunOtherCost1 = oMRunOthCst1Col.Cells.Item(aCurrentRow).Specific
            oRunOtherCost2 = oMRunOthCst2Col.Cells.Item(aCurrentRow).Specific
            oRunTotMacCost = CDbl(oRunOprCost.Value) + CDbl(oRunPowerCost.Value) + CDbl(oRunOtherCost1.Value) + CDbl(oRunOtherCost2.Value)
        Catch ex As Exception
            Throw ex
        End Try
        Return oRunTotMacCost
    End Function
    ''' <summary>
    ''' Calculating the running labour cost.
    ''' </summary>
    ''' <param name="aCurrentRow"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RunLabourTotalCost(ByVal aCurrentRow As Integer) As Double
        Dim oRunTotLabCost As Double
        Dim oLabRatePerHour, oLWrkTime As SAPbouiCOM.EditText
        Try
            oLabRatePerHour = oLabRateCol.Cells.Item(aCurrentRow).Specific
            oLWrkTime = oLWrkTimeCol.Cells.Item(aCurrentRow).Specific
            oRunTotLabCost = CDbl(CDbl(oLabRatePerHour.Value) / 60) * CDbl(oLWrkTime.Value)
        Catch ex As Exception
            Throw ex
        End Try
        Return oRunTotLabCost
    End Function
    ''' <summary>
    ''' Calculating the running tool cost.
    ''' </summary>
    ''' <param name="aCurrentRow"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RunToolTotalCost(ByVal aCurrentRow As Integer, ByVal aToolCostPerPiece As Double, ByVal aQty As Double) As Double
        Dim oRunTotToolCost As Double
        Try
            oRunTotToolCost = aToolCostPerPiece * aQty
        Catch ex As Exception
            Throw ex
        End Try
        Return oRunTotToolCost
    End Function
    ''' <summary>
    ''' Calculating the running machine total cost.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function TotalMachineCost() As Double
        Dim oTotMacCost As Double
        Dim IntICount As Integer
        Dim oRunMachineTotalCost As SAPbouiCOM.EditText
        Try
            For IntICount = 1 To oMacMatrix.RowCount
                oRunMachineTotalCost = oMTotRunCostCol.Cells.Item(IntICount).Specific
                oMacMatrix.GetLineData(IntICount)
                oTotMacCost = oTotMacCost + CDbl(oRunMachineTotalCost.Value)
            Next
            oParentDB.SetValue("U_Totmcst", oParentDB.Offset, oTotMacCost)
        Catch ex As Exception
            Throw ex
        End Try
        Return oTotMacCost
    End Function
    ''' <summary>
    ''' Calculating the total Labour cost.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function TotalLabourCost() As Double
        Dim oTotLabCost As Double
        Dim IntICount As Integer
        Dim oRunLabourTotalCost As SAPbouiCOM.EditText
        Try
            For IntICount = 1 To oLabMatrix.RowCount
                oRunLabourTotalCost = oLTotRunCstCol.Cells.Item(IntICount).Specific
                oLabMatrix.GetLineData(IntICount)
                oTotLabCost = oTotLabCost + CDbl(oRunLabourTotalCost.Value)
            Next
        Catch ex As Exception
            Throw ex
        End Try
        Return oTotLabCost
    End Function
    ''' <summary>
    ''' Calculating the total tool cost.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function TotalToolCost() As Double
        Dim oTotToolCost As Double
        Dim IntICount As Integer
        Dim oRunToolTotalCost As SAPbouiCOM.EditText
        Try
            For IntICount = 1 To oToolsMatrix.RowCount
                oRunToolTotalCost = oTotToolCstCol.Cells.Item(IntICount).Specific
                oToolsMatrix.GetLineData(IntICount)
                oTotToolCost = oTotToolCost + CDbl(oRunToolTotalCost.Value)
            Next
        Catch ex As Exception
            Throw ex
        End Try
        Return oTotToolCost
    End Function
    ''' <summary>
    ''' Fetching the GLMethod from the corresponding tables based on the GLMethod in ItemMaster Form.
    ''' </summary>
    ''' <param name="aGLMethod"></param>
    ''' <remarks></remarks>
    Private Sub GLMethod(ByVal aGLMethod As String)
        Dim oStrSql As String
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            If aGLMethod = "W" Then           '*********Warehouse**********
                oStrSql = "Select T0.WIPAcct,T1.FormatCode,T1.AcctName from OWHS T0 " _
                & "Inner Join OACT T1 On T1.AcctCode = T0.WIPAcct  Where WhsCode = '" & oWhsCodeTxt.Value & "'"
                oRs.DoQuery(oStrSql)
                If oRs.RecordCount > 0 Then
                    oRs.MoveFirst()
                    UConAcCode.Value = oRs.Fields.Item("WipAcct").Value
                    'UConAcCode.Value = oRs.Fields.Item("FormatCode").Value
                    UConAcName.Value = oRs.Fields.Item("AcctName").Value
                End If
            ElseIf aGLMethod = "C" Then       '*********Item Group*********
                oStrSql = "Select T0.WIPAcct,T1.FormatCode,T1.AcctName from OWHS T0 " _
                & "Inner Join OACT T1 On T1.AcctCode = T0.WIPAcct Where WhsCode = '" & oWhsCodeTxt.Value & "'"
                oRs.DoQuery(oStrSql)
                If oRs.RecordCount > 0 Then
                    oRs.MoveFirst()
                    UConAcCode.Value = oRs.Fields.Item("WipAcct").Value
                    'UConAcCode.Value = oRs.Fields.Item("FormatCode").Value
                    UConAcName.Value = oRs.Fields.Item("AcctName").Value
                End If
            ElseIf aGLMethod = "L" Then       '*********Item Level*********
                oStrSql = "Select T0.WIPAcct,T1.FormatCode,T1.AcctName from OITW T0 " _
                & "Inner Join OACT T1 On T1.AcctCode = T0.WIPAcct " _
                & "Where ItemCode = '" & oItemCodeTxt.Value & "' and WhsCode = '" & oWhsCodeTxt.Value & "'"
                oRs.DoQuery(oStrSql)
                If oRs.RecordCount > 0 Then
                    oRs.MoveFirst()
                    UConAcCode.Value = oRs.Fields.Item("WipAcct").Value
                    'UConAcCode.Value = oRs.Fields.Item("FormatCode").Value
                    UConAcName.Value = oRs.Fields.Item("AcctName").Value
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
    ''' Validating the values in the cell.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Validation()
        Dim IntICount As Integer
        Dim oLabKey As SAPbouiCOM.CheckBox
        Dim nop, oMacName As SAPbouiCOM.EditText
        Dim oSKCode As SAPbouiCOM.ComboBox
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oStrSql As String

        Try
            '****** date validation************
            Dim dtdat As Date
            dtdat = getdatetime(oForm.Items.Item("txtdocdt").Specific.string)

            'MsgBox(dtdat)

            If dtdat > Now() Then
                Throw New Exception("Date Should not be greater than Current date.....")
            End If

            '**************************

            If oPONoTxt.Value.Length = 0 Then
                oPONoTxt.Active = True
                Throw New Exception("Select Production Order No from the List")
            End If
            If oShiftCodeTxt.Value.Length = 0 Then
                oShiftCodeTxt.Active = True
                Throw New Exception("Select Shift Code from the List")
            End If
            If oParentDB.GetValue("U_oplnid", oParentDB.Offset).Trim().Length = 0 Then
                oOprCombo.Active = True
                Throw New Exception("Operation is Mandatory")
            End If
            If oProdQtyTxt.Value = 0 Then
                oProdQtyTxt.Active = True
                Throw New Exception("Produced Qty should be entered")
            End If
            If oMacMatrix.RowCount = 1 Then
                oMacName = oMacNameCol.Cells.Item(oMacMatrix.RowCount).Specific
                oMacMatrix.GetLineData(oMacMatrix.RowCount)
                If oMacName.Value.Length = 0 Then
                    Throw New Exception("Atleast a Machine should be added")
                End If
            End If

            'Commented by Manimaran-------s
            'If oMacMatrix.RowCount > 0 Then
            '    If MacQtyCalculation() <> CDbl(oProdQtyTxt.Value) Then
            '        Throw New Exception("Sum Of Machine Qty should be equal to the Produced Qty")
            '    End If
            'End If
            'Commented by Manimaran-------e
            If oToolsMatrix.RowCount > 0 Then
                If ToolsQtyCalculation() <> CDbl(oProdQtyTxt.Value) Then
                    Throw New Exception("Sum Of Tools Qty should be equal to the Produced Qty")
                End If
            End If

            If oLabMatrix.RowCount > 0 Then
                If LabQtyCalculation() <> CDbl(oProdQtyTxt.Value) Then
                    Throw New Exception("Sum Of Labour Qty should be equal to the Produced Qty")
                End If
            End If
            oStrSql = "Select * from [@PSSIT_OCON]"
            oRs.DoQuery(oStrSql)
            If oRs.RecordCount > 0 Then
                oRs.MoveFirst()
                If oRs.Fields.Item("U_Labman").Value = "Y" Then
                    If oLabMatrix.RowCount = 0 Then
                        Throw New Exception("Labour Details are Mandatory")
                    End If
                End If
            End If
            For IntICount = 1 To oLabMatrix.RowCount
                oLabKey = oLabKeyCol.Cells.Item(IntICount).Specific
                'oLabCode = oLabCodeCol.Cells.Item(IntICount).Specific
                nop = oLabMatrix.Columns.Item("colnop").Cells.Item(IntICount).Specific
                oSKCode = oLabMatrix.Columns.Item("collgcod").Cells.Item(IntICount).Specific
                oLabMatrix.GetLineData(IntICount)
                If oLabKey.Checked = True Then
                    If oSKCode.Value.Length = 0 Then
                        Throw New Exception("Labour Details are Mandatory")
                    End If
                End If
                If nop.Value = "" Then
                    Throw New Exception("Enter No. of Persons.....")
                End If
            Next
            'Added by Manimaran--------s
            For IntICount = 1 To oMacMatrix.RowCount
                If CDbl(oMacMatrix.Columns.Item("colrntim").Cells.Item(IntICount).Specific.value) = 0 Then
                    Throw New Exception("Run Time should be greater than 0 in the Machine Matrix")
                End If
            Next

            For IntICount = 1 To oLabMatrix.RowCount
                If CDbl(oLabMatrix.Columns.Item("colwktim").Cells.Item(IntICount).Specific.value) = 0 Then
                    Throw New Exception("Worked Time should be greater than 0 in the Labour Matrix")
                End If
            Next

            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                Dim sqry As String
                Dim TPassQty As Double
                Dim rs As SAPbobsCOM.Recordset
                rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'sqry = "select isnull(sum(t0.u_prodqty),0) from [@PSSIT_OPEY] t0"
                'sqry = sqry + " where t0.u_pnordno = " & oForm.Items.Item("txtprdno").Specific.string & ""
                sqry = "select (isnull(sum(U_Passqty),0)+ isnull(SUM(U_Rewrkqty),0) + isnull(sum(U_scrapqty ),0)) pqty  from [@PSSIT_WOR2] where U_Pordno = '" & oForm.Items.Item("txtprdno").Specific.string & "' and U_Oprname = '" & oOprCombo.Selected.Description & "'"
                rs.DoQuery(sqry)
                If rs.RecordCount > 0 Then
                    TPassQty = CDbl(rs.Fields.Item(0).Value)
                End If
                If CDbl(oForm.Items.Item("txtplqty").Specific.value) < TPassQty + CDbl(oForm.Items.Item("txtpdqty").Specific.value) Then
                    'Throw New Exception("Produced quantity should be less or equal to the operation quantity")
                End If
            End If


            'Added by senthil to display the Rework and scrap quantities in the header.

            Dim dblQty, dblRwqty, dblscqty As Double
            dblRwqty = 0
            dblscqty = 0
            For intRow As Integer = 1 To oReWrkMatrix.VisualRowCount
                dblRwqty = dblRwqty + CDbl(oReWrkMatrix.Columns.Item("V_1").Cells.Item(intRow).Specific.value)
            Next
            For intRow As Integer = 1 To oScrpMatrix.VisualRowCount
                dblscqty = dblscqty + CDbl(oScrpMatrix.Columns.Item("V_1").Cells.Item(intRow).Specific.value)
            Next
            'oForm.Items.Item("txtrwqty").Specific.value = dblRwqty
            'oForm.Items.Item("txtspqty").Specific.value = dblscqty

            oForm.Items.Item("107").Specific.value = dblRwqty
            oForm.Items.Item("109").Specific.value = dblscqty

            If ProducedQtyCalculation() <> CDbl(oProdQtyTxt.Value) Then
                Throw New Exception("Sum Of Passed,Rework and Scrap should be equal to the Produced Qty")
            End If
            'Added by Manimaran----------e
            'If CDbl(URewQty.Value) > 0 Then
            '    If oRewCheck.Checked = False Then
            '        oRewCheck.Checked = True
            '    End If
            '    'If CDbl(oProdQtyTxt.Value) > CDbl(URewQty.Value) Then
            '    '    Throw New Exception("Produced Qty should be less than the Actual Rework Qty : " & CDbl(URewQty.Value))
            '    '    'ElseIf CDbl(oProdQtyTxt.Value) < CDbl(URewQty.Value) Then
            '    '    '    oClosedCheck.Checked = False
            '    '    'ElseIf CDbl(oProdQtyTxt.Value) = CDbl(URewQty.Value) Then
            '    '    '    oClosedCheck.Checked = True
            '    'End If
            'End If
            'If CDbl(oOprRewQtyTxt.Value) > 0 Then
            '    oClosedCheck.Checked = False
            'Else
            '    oClosedCheck.Checked = True
            'End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#Region "Journal Posting"
    Private Function JournalPosting() As Boolean
        Dim ErrCode As Long
        Dim ErrMsg As String = ""
        Dim IntICount As Integer
        Dim Journal As SAPbobsCOM.JournalEntries
        Dim dblDebit As Double
        Dim oCAcCode As String = ""
        Dim D1 As Date
        dblDebit = 0
        'D1 = String2Date(oDocDateTxt.String, "MM/DD/YYYY")
        Try
            DataRow = oDataTable.NewRow()
            DataRow.Item("AcCode") = CStr(UConAcCode.Value)
            DataRow.Item("AcName") = CStr(UConAcName.Value)
            DataRow.Item("CAcCode") = oCAcCode
            'DataRow.Item("Debit") = CDbl(oTotMacCostTxt.Value) + CDbl(oTotLabCostTxt.Value) + CDbl(oTotToolCostTxt.Value) + CDbl(UTotFCost.Value)
            dblDebit = CDbl(oTotMacCostTxt.Value) + CDbl(oTotLabCostTxt.Value) + CDbl(oTotToolCostTxt.Value) + CDbl(UTotFCost.Value)
            If dblJournalCredit <> dblDebit Then
                dblDebit = dblJournalCredit
            End If
            DataRow.Item("Debit") = dblDebit
            DataRow.Item("Credit") = 0
            DataRow.Item("ShortName") = oCAcCode
            oDataTable.Rows.Add(DataRow)

            '*****************Journal Creation *************************************
            Journal = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
            D1 = CDate(Mid(oDocDateTxt.Value, 1, 4) & "-" & Mid(oDocDateTxt.Value, 5, 2) & "-" & Mid(oDocDateTxt.Value, 7, 2))
            Journal.TaxDate = Format(D1, "yyyy-MM-dd")
            Journal.Reference = oParentDB.GetValue("DocNum", oParentDB.Offset).Trim()
            Journal.Reference2 = oParentDB.GetValue("U_Pnordno", oParentDB.Offset).Trim()
            Journal.Memo = "WIP Journal (Internal)"
            For IntICount = 0 To oDataTable.Rows.Count - 1
                Journal.Lines.AccountCode = oDataTable.Rows(IntICount)("AcCode")
                Journal.Lines.ContraAccount = oDataTable.Rows(IntICount)("CAcCode")
                Journal.Lines.Debit = oDataTable.Rows(IntICount)("Debit")
                Journal.Lines.Credit = oDataTable.Rows(IntICount)("Credit")
                Journal.Lines.TaxDate = Format(D1, "yyyy-MM-dd")
                Journal.Lines.DueDate = Format(D1, "yyyy-MM-dd")
                Journal.Lines.ReferenceDate1 = Now
                Journal.Lines.ShortName = oDataTable.Rows(IntICount)("AcCode")
                If IntICount <> oDataTable.Rows.Count - 1 Then
                    Call Journal.Lines.Add()
                    Call Journal.Lines.SetCurrentLine(IntICount + 1)
                End If
            Next
            iJournal = Journal.Add()
            If iJournal <> 0 Then
                oCompany.GetLastError(ErrCode, ErrMsg)
                MsgBox(ErrCode & " " & ErrMsg)
                Return False
            End If
            LoadJvno()
            If iJournal = 0 Then
                Return True
            End If
        Catch ex As Exception
            MsgBox("Exception : " & ex.Message)
            'Return False
        End Try
    End Function

    Private Sub LoadJvno()
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            oRs.DoQuery("Select * from OJDT Where Ref1 = '" & oPENoTxt.Value & "' and Ref2 = '" & oPONoTxt.Value & "'")
            If oRs.RecordCount > 0 Then
                oRs.MoveFirst()
                oJVNoTxt.Value = oRs.Fields.Item("TransId").Value
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
#End Region
    ''' <summary>
    ''' Updating the Produced Qty, Passed Qty, Rework Qty, Scrap Qty, Total Machine Cost,
    ''' Total Labour Cost, Total Tool Cost in the Production Order Route Details Table.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub UpdateProductionOrder()
        Dim oStrSql As String
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try

            'If oRewCheck.Checked = False Then
            '    'Modified by Manimaran------s
            '    'oStrSql = "Update [@PSSIT_WOR2] Set  " _
            '    '& "U_ProdQty = " & (CDbl(oParentDB.GetValue("U_ProdQty", oParentDB.Offset).Trim()) + CDbl(oParentDB.GetValue("U_AccProdQty", oParentDB.Offset).Trim())) _
            '    '    & " , U_PassQty = " & (CDbl(oParentDB.GetValue("U_Passqty", oParentDB.Offset).Trim()) + CDbl(oParentDB.GetValue("U_AccPassqty", oParentDB.Offset).Trim())) _
            '    '    & " , U_RewrkQty = " & (CDbl(oParentDB.GetValue("U_Rewrkqty", oParentDB.Offset).Trim()) + CDbl(oParentDB.GetValue("U_AccRewqty", oParentDB.Offset).Trim())) _
            '    '    & " , U_PenRewQty = " & (CDbl(oParentDB.GetValue("U_Rewrkqty", oParentDB.Offset).Trim()) + CDbl(oParentDB.GetValue("U_AccRewqty", oParentDB.Offset).Trim())) _
            '    '    & " , U_ScrapQty = " & (CDbl(oParentDB.GetValue("U_scrapqty", oParentDB.Offset).Trim()) + CDbl(oParentDB.GetValue("U_Accscrapqty", oParentDB.Offset).Trim())) _
            '    '    & " , U_LbrCst = " & (CDbl(oParentDB.GetValue("U_Totlcst", oParentDB.Offset).Trim()) + CDbl(UAccLabCost.Value)) _
            '    '    & " , U_Mccst = " & (CDbl(oParentDB.GetValue("U_Totmcst", oParentDB.Offset).Trim()) + CDbl(UAccMacCost.Value)) _
            '    '    & " , U_Toolcst = " & (CDbl(oParentDB.GetValue("U_Tottcst", oParentDB.Offset).Trim()) + CDbl(UAccToolCost.Value)) _
            '    '    & " , U_CloseKey = '" & oParentDB.GetValue("U_CloseKey", oParentDB.Offset).Trim() _
            '    '    & "' Where U_POrdno = " & oPONoTxt.Value & " and U_OprCode = '" & oOprCodeTxt.Value _
            '    '    & "' and U_Rteid = '" & oRteIDTxt.Value & "' and U_Baslino = " & oOprCombo.Selected.Value
            '    oStrSql = "Update [@PSSIT_WOR2] Set  " _
            '   & "U_ProdQty = " & (CDbl(oParentDB.GetValue("U_ProdQty", oParentDB.Offset).Trim()) + CDbl(oParentDB.GetValue("U_AccProdQty", oParentDB.Offset).Trim())) _
            '       & " , U_PassQty = " & (CDbl(oParentDB.GetValue("U_Passqty", oParentDB.Offset).Trim()) + CDbl(oParentDB.GetValue("U_AccPassqty", oParentDB.Offset).Trim())) _
            '       & " , U_RewrkQty = " & TotRwrkQty + CDbl(oParentDB.GetValue("U_AccRewqty", oParentDB.Offset).Trim()) _
            '       & " , U_PenRewQty = " & TotRwrkQty + CDbl(oParentDB.GetValue("U_AccRewqty", oParentDB.Offset).Trim()) _
            '       & " , U_ScrapQty = " & TotScrpQty + CDbl(oParentDB.GetValue("U_Accscrapqty", oParentDB.Offset).Trim()) _
            '       & " , U_LbrCst = " & (CDbl(oParentDB.GetValue("U_Totlcst", oParentDB.Offset).Trim()) + CDbl(UAccLabCost.Value)) _
            '       & " , U_Mccst = " & (CDbl(oParentDB.GetValue("U_Totmcst", oParentDB.Offset).Trim()) + CDbl(UAccMacCost.Value)) _
            '       & " , U_Toolcst = " & (CDbl(oParentDB.GetValue("U_Tottcst", oParentDB.Offset).Trim()) + CDbl(UAccToolCost.Value)) _
            '       & " , U_CloseKey = '" & oParentDB.GetValue("U_CloseKey", oParentDB.Offset).Trim() _
            '       & "' Where U_POrdno = " & oPONoTxt.Value & " and U_OprCode = '" & oOprCodeTxt.Value _
            '       & "' and U_Rteid = '" & oRteIDTxt.Value & "' and U_Baslino = " & oOprCombo.Selected.Value
            '    'Modified by Manimaran------e
            '    oRs.DoQuery(oStrSql)
            'ElseIf oRewCheck.Checked = True Then
            '    'If CDbl(oProdQtyTxt.Value) = CDbl(URewQty.Value) Then
            '    '    oStrSql = "Update [@PSSIT_WOR2] Set U_PenRewQty = " & (CDbl(oProdQtyTxt.Value) - CDbl(URewQty.Value)) _
            '    '    & " , U_CloseKey = 'Y'" _
            '    '    & " Where U_POrdno = " & oPONoTxt.Value & " and U_OprCode = '" & oOprCodeTxt.Value _
            '    '    & "' and U_Rteid = '" & oRteIDTxt.Value & "' and U_Baslino = " & oOprCombo.Selected.Value
            '    '    oRs.DoQuery(oStrSql)
            '    'ElseIf CDbl(oProdQtyTxt.Value) < CDbl(URewQty.Value) Then
            '    '    oStrSql = "Update [@PSSIT_WOR2] Set U_PenRewQty = " & (CDbl(URewQty.Value) - CDbl(oProdQtyTxt.Value)) _
            '    '    & " , U_CloseKey = 'N'" _
            '    '    & " Where U_POrdno = " & oPONoTxt.Value & " and U_OprCode = '" & oOprCodeTxt.Value _
            '    '    & "' and U_Rteid = '" & oRteIDTxt.Value & "' and U_Baslino = " & oOprCombo.Selected.Value
            '    '    oRs.DoQuery(oStrSql)
            '    'End If
            '    Dim osbSQL As New System.Text.StringBuilder
            '    ' INSERT INTO *****************************
            '    osbSQL.Append("INSERT INTO [@PSSIT_WOR2] (")
            '    osbSQL.Append("[Code]")
            '    osbSQL.Append(",[Name]")
            '    osbSQL.Append(",[U_POrdSer]")
            '    osbSQL.Append(",[U_POrdno]")
            '    osbSQL.Append(",[U_Baslino]")
            '    osbSQL.Append(",[U_Seqnce]")
            '    osbSQL.Append(",[U_Parid]")
            '    osbSQL.Append(",[U_OprCode]")
            '    osbSQL.Append(",[U_OprName]")
            '    osbSQL.Append(",[U_Rteid]")
            '    osbSQL.Append(",[U_Seqbaslino]")
            '    osbSQL.Append(",[U_ProdQty]")
            '    osbSQL.Append(",[U_Passqty]")
            '    osbSQL.Append(",[U_Rewrkqty]")
            '    osbSQL.Append(",[U_Scrapqty]")
            '    osbSQL.Append(",[U_PenRewqty]")
            '    osbSQL.Append(",[U_Lbrcst]")
            '    osbSQL.Append(",[U_Mccst]")
            '    osbSQL.Append(",[U_Toolcst]")
            '    osbSQL.Append(",[U_Subcst]")
            '    osbSQL.Append(",[U_Wodoc]")
            '    osbSQL.Append(",[U_Rework]")
            '    osbSQL.Append(",[U_Adnl1]")
            '    osbSQL.Append(",[U_Adnl2]")
            '    osbSQL.Append(",[U_Adnl3]")
            '    osbSQL.Append(",[U_Adnl4]")
            '    osbSQL.Append(",[U_CloseKey]")
            '    ' VALUES ***********************************    
            '    Dim xLCode As String = GenerateSerialNo("PSSIT_WOR2")
            '    osbSQL.Append(") VALUES (")
            '    osbSQL.Append(xLCode)
            '    osbSQL.Append("," & xLCode)
            '    osbSQL.Append("," & oPOSeriesTxt.Value)
            '    osbSQL.Append("," & oPONoTxt.Value)
            '    Dim oBRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '    oRs.DoQuery("Select IsNull(Max(Convert(Float,U_Baslino)),0) as BaseLino From [@PSSIT_WOR2] where U_Pordno = " & oPONoTxt.Value)
            '    osbSQL.Append("," & CInt(oRs.Fields.Item("BaseLino").Value) + 1)
            '    osbSQL.Append("," & CInt(oRs.Fields.Item("BaseLino").Value) + 1)
            '    osbSQL.Append("," & oOprCombo.Selected.Value)
            '    osbSQL.Append(",'" & oOprCodeTxt.Value)
            '    osbSQL.Append("','" & oOprLineIDTxt.Value)
            '    osbSQL.Append("','" & oRteIDTxt.Value)
            '    osbSQL.Append("',0")
            '    osbSQL.Append("," & CDbl(oParentDB.GetValue("U_ProdQty", oParentDB.Offset).Trim()))
            '    osbSQL.Append("," & CDbl(oParentDB.GetValue("U_PassQty", oParentDB.Offset).Trim()))
            '    'Modified by Manimaran-----s
            '    'osbSQL.Append("," & CDbl(oParentDB.GetValue("U_RewrkQty", oParentDB.Offset).Trim()))
            '    'osbSQL.Append("," & CDbl(oParentDB.GetValue("U_scrapQty", oParentDB.Offset).Trim()))
            '    'osbSQL.Append("," & CDbl(oParentDB.GetValue("U_RewrkQty", oParentDB.Offset).Trim()))
            '    osbSQL.Append("," & TotRwrkQty)
            '    osbSQL.Append("," & TotScrpQty)
            '    osbSQL.Append("," & TotRwrkQty)
            '    'Modified by Manimaran-----e
            '    osbSQL.Append("," & CDbl(oParentDB.GetValue("U_Totlcst", oParentDB.Offset).Trim()))
            '    osbSQL.Append("," & CDbl(oParentDB.GetValue("U_Totmcst", oParentDB.Offset).Trim()))
            '    osbSQL.Append("," & CDbl(oParentDB.GetValue("U_Tottcst", oParentDB.Offset).Trim()))
            '    osbSQL.Append(",0")
            '    osbSQL.Append(",NULL")
            '    osbSQL.Append(",'" & oParentDB.GetValue("U_Rework", oParentDB.Offset).Trim())
            '    osbSQL.Append("',0")
            '    osbSQL.Append(",0")
            '    osbSQL.Append(",0")
            '    osbSQL.Append(",0")
            '    osbSQL.Append(",'" & oParentDB.GetValue("U_CloseKey", oParentDB.Offset).Trim())
            '    osbSQL.Append("')")
            '    ' INSERT PARENT **************************************************************
            '    Dim sLSQL As String = osbSQL.ToString
            '    Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '    oRs1.DoQuery(sLSQL)
            'End If
            oStrSql = "Update [@PSSIT_WOR2] Set  " _
              & "U_ProdQty = " & (CDbl(oParentDB.GetValue("U_ProdQty", oParentDB.Offset).Trim()) + CDbl(oParentDB.GetValue("U_AccProdQty", oParentDB.Offset).Trim())) _
                  & " , U_PassQty = " & (CDbl(oParentDB.GetValue("U_Passqty", oParentDB.Offset).Trim()) + CDbl(oParentDB.GetValue("U_AccPassqty", oParentDB.Offset).Trim())) _
                  & " , U_RewrkQty = " & TotRwrkQty + CDbl(oParentDB.GetValue("U_AccRewqty", oParentDB.Offset).Trim()) _
                  & " , U_PenRewQty = " & TotRwrkQty + CDbl(oParentDB.GetValue("U_AccRewqty", oParentDB.Offset).Trim()) _
                  & " , U_ScrapQty = " & TotScrpQty + CDbl(oParentDB.GetValue("U_Accscrapqty", oParentDB.Offset).Trim()) _
                  & " , U_LbrCst = " & (CDbl(oParentDB.GetValue("U_Totlcst", oParentDB.Offset).Trim()) + CDbl(UAccLabCost.Value)) _
                  & " , U_Mccst = " & (CDbl(oParentDB.GetValue("U_Totmcst", oParentDB.Offset).Trim()) + CDbl(UAccMacCost.Value)) _
                  & " , U_Toolcst = " & (CDbl(oParentDB.GetValue("U_Tottcst", oParentDB.Offset).Trim()) + CDbl(UAccToolCost.Value)) _
                  & " , U_CloseKey = '" & oParentDB.GetValue("U_CloseKey", oParentDB.Offset).Trim() _
                  & "' Where U_POrdno = " & oPONoTxt.Value & " and U_OprCode = '" & oOprCodeTxt.Value _
                  & "' and U_Rteid = '" & oRteIDTxt.Value & "' and U_Baslino = " & oOprCombo.Selected.Value
            'Modified by Manimaran------e
            oRs.DoQuery(oStrSql)
            '******************Updating Cost Header [@PSSIT_WOR3]*********************************
            oStrSql = "Update [@PSSIT_WOR3] Set  " _
            & "U_Totlbrcst = U_Totlbrcst + " & (CDbl(oParentDB.GetValue("U_Totlcst", oParentDB.Offset).Trim())) _
            & " , U_Totmccst = U_Totmccst + " & (CDbl(oParentDB.GetValue("U_Totmcst", oParentDB.Offset).Trim())) _
            & " , U_Tottoolcst = U_Tottoolcst + " & (CDbl(oParentDB.GetValue("U_Tottcst", oParentDB.Offset).Trim())) _
            & " ,  U_Totcst = U_TotCst + " & (CDbl(oParentDB.GetValue("U_Totlcst", oParentDB.Offset).Trim()) + CDbl(oParentDB.GetValue("U_Totmcst", oParentDB.Offset).Trim()) + CDbl(oParentDB.GetValue("U_Tottcst", oParentDB.Offset).Trim()) + CDbl(UTotFCost.Value)) _
            & " Where U_POrdno = " & oPONoTxt.Value & " and U_POrdSer = " & oPOSeriesTxt.Value
            oRs.DoQuery(oStrSql)
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    Private Sub UpdateTools()
        Dim IntICount As Integer
        Dim oStrSql, oUpStrSql As String
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            For IntICount = 1 To oToolsMatrix.RowCount
                oToolsMatrix.GetLineData(IntICount)
                oStrSql = "Select * from [@PSSIT_OTLS] Where Code = '" & oToolCodeCol.Cells.Item(IntICount).Specific.Value & "'"
                oRs.DoQuery(oStrSql)
                If oRs.RecordCount > 0 Then
                    oUpStrSql = "Update [@PSSIT_OTLS] Set U_Cnou = " & (CDbl(oTQtyCol.Cells.Item(IntICount).Specific.Value) + CDbl(oRs.Fields.Item("U_Cnou").Value)) _
                    & " Where Code = '" & oToolCodeCol.Cells.Item(IntICount).Specific.Value & "'"
                    oRs1.DoQuery(oUpStrSql)
                End If
            Next
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    Private Sub UpdateFixedCostDetails()
        Dim oFTotalCostEdit, oFFxdCostEdit, oFPONoEdit, oFPOSerEdit As SAPbouiCOM.EditText
        Dim oStrSql As String
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            '******************Updating Cost Details [@PSSIT_WOR4]*********************************
            For IntICount As Integer = 1 To oFCMatrix.RowCount
                oFPOSerEdit = oFPOSerCol.Cells.Item(IntICount).Specific
                oFPONoEdit = oFPONoCol.Cells.Item(IntICount).Specific
                oFFxdCostEdit = oFFixedCostCol.Cells.Item(IntICount).Specific
                oFTotalCostEdit = oFTotCostCol.Cells.Item(IntICount).Specific
                oFCMatrix.GetLineData(IntICount)
                oStrSql = "Update [@PSSIT_WOR4] Set " _
                & "U_Totfcst = U_TotFcst + " & (CDbl(oFTotalCostEdit.Value)) _
                & " Where U_Pordser = " & oFPOSerEdit.Value & " and U_Pordno = " & oFPONoEdit.Value _
                & " and U_FCost = '" & oFFxdCostEdit.Value & "'"
                oRs.DoQuery(oStrSql)
            Next
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
    Private Sub AccKeyCheck()
        Dim oStrSql As String
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            oStrSql = "Select * from [@PSSIT_OCON]"
            oRs.DoQuery(oStrSql)
            If oRs.RecordCount > 0 Then
                oRs.MoveFirst()
                If oRs.Fields.Item("U_AccKey").Value = "Y" Or oRs.Fields.Item("U_AccKey").Value = "y" Then
                    oAccKeyCheck.Checked = True
                    oParentDB.SetValue("U_AccKey", oParentDB.Offset, "Y")
                ElseIf oRs.Fields.Item("U_AccKey").Value = "N" Or oRs.Fields.Item("U_AccKey").Value = "n" Then
                    oAccKeyCheck.Checked = False
                    oParentDB.SetValue("U_AccKey", oParentDB.Offset, "N")
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    Function String2Date(ByVal S As String, _
                            ByVal Fmt As String) As Object
        If Format(S) = "DD/MM/YY" Then

            Fmt = "MM/DD/YY"
        ElseIf Format(S) = "DD.MM.YY" Then
            Fmt = "MM/DD/YY"
        ElseIf Format(S) = "MM/DD/YY" Then
            Fmt = "DD/MM/YY"
        Else
            Fmt = "MM/DD/YY"

        End If
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

                String2Date = CDate(Mid(S, 4, 3) & Left(S, 3) & Right(S, 2))
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
    Private Function TimeFormat(ByVal oTime As String) As String
        Dim oResTime As String
        Try
            If oTime.Length = 4 Then
                oResTime = String.Concat(oTime.Substring(0, 2), ":", oTime.Substring(2, 2))
            Else
                oResTime = String.Concat(oTime.Substring(0, 1), ":", oTime.Substring(1, 2))
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return oResTime
    End Function
    Private Sub LoadProdOrderDocEntryNo(ByVal aPONo As Integer)
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            oRs.DoQuery("Select DocEntry from OWOR where DocNum = " & aPONo)
            If oRs.RecordCount > 0 Then
                oRs.MoveFirst()
                UPODocEnt.Value = oRs.Fields.Item("DocEntry").Value
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    Private Sub AddToolDatatoDB()
        Dim IntICount, ITools As Integer
        Dim oToolCodeEdit, oTCodeEdit, oTPENoEdit, oTMacLineIDEdit, oTMacDocEntryEdit, oTMacCodeEdit, oToolDescEdit, oTQtyEdit, oTAcCodeEdit, oTAcNameEdit, oTConAcCodeEdit, oTConAcNameEdit, oTToolCostEdit, oTTotRunCostEdit, oTInfo1Edit, oTInfo2Edit As SAPbouiCOM.EditText
        Dim oTAccKey As SAPbouiCOM.CheckBox
        Try
            For IntICount = 1 To oToolsMatrix.VisualRowCount
                oTCodeEdit = oTCodeCol.Cells.Item(IntICount).Specific
                oTPENoEdit = oTPENoCol.Cells.Item(IntICount).Specific
                oTMacLineIDEdit = oTMLineIDCol.Cells.Item(IntICount).Specific
                oTMacDocEntryEdit = oTMDocEntryCol.Cells.Item(IntICount).Specific
                oTMacCodeEdit = oTMacNoCol.Cells.Item(IntICount).Specific
                oToolCodeEdit = oToolCodeCol.Cells.Item(IntICount).Specific
                oToolDescEdit = oToolDescCol.Cells.Item(IntICount).Specific
                oTQtyEdit = oTQtyCol.Cells.Item(IntICount).Specific
                oTAcCodeEdit = oTAccCodeCol.Cells.Item(IntICount).Specific
                oTAccKey = oTAccKeyCol.Cells.Item(IntICount).Specific
                oTAcNameEdit = oTAccNameCol.Cells.Item(IntICount).Specific
                oTConAcCodeEdit = oTConAccCodeCol.Cells.Item(IntICount).Specific
                oTConAcNameEdit = oTConAccNameCol.Cells.Item(IntICount).Specific
                oTToolCostEdit = oToolCstCol.Cells.Item(IntICount).Specific
                oTTotRunCostEdit = oTotToolCstCol.Cells.Item(IntICount).Specific
                oTInfo1Edit = oTInfo1Col.Cells.Item(IntICount).Specific
                oTInfo2Edit = oTInfo2Col.Cells.Item(IntICount).Specific
                oToolsMatrix.GetLineData(IntICount)
                '***************Get Account Code For Journal - Tools****************
                DataRow = oDataTable.NewRow()
                DataRow.Item("AcCode") = CStr(oTAcCodeEdit.Value)
                DataRow.Item("AcName") = CStr(oTAcNameEdit.Value)
                DataRow.Item("CAcCode") = CStr(UConAcCode.Value)
                DataRow.Item("CAcName") = CStr(UConAcName.Value)
                DataRow.Item("Debit") = 0
                DataRow.Item("Credit") = CDbl(oTTotRunCostEdit.Value)
                DataRow.Item("ShortName") = CStr(oTAcCodeEdit.Value)
                dblJournalCredit = dblJournalCredit + CDbl(oTTotRunCostEdit.Value)
                dblJournalDebit = dblJournalDebit + 0
                oDataTable.Rows.Add(DataRow)
                '***********************************************************
                If PSSIT_PEY3.GetByKey(oTCodeEdit.Value) = True Then
                    PSSIT_PEY3.Code = oTCodeEdit.Value
                    PSSIT_PEY3.Name = oTCodeEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_Prdentno").Value = oTPENoEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_Maclid").Value = oTMacLineIDEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_wcno").Value = oTMacCodeEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_Madcey").Value = oTMacDocEntryEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_Toolcode").Value = oToolCodeEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_TLname").Value = oToolDescEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_Qty").Value = oTQtyEdit.Value
                    If oTAccKey.Checked = True Then
                        PSSIT_PEY3.UserFields.Fields.Item("U_Acckey").Value = "Y"
                    Else
                        PSSIT_PEY3.UserFields.Fields.Item("U_Acckey").Value = "N"
                    End If
                    PSSIT_PEY3.UserFields.Fields.Item("U_Accode").Value = oTAcCodeEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_Acname").Value = oTAcNameEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_CAccode").Value = oTConAcCodeEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_CAcname").Value = oTConAcNameEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_Tlctppie").Value = oTToolCostEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_Totcost").Value = oTTotRunCostEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_Adnl1").Value = oTInfo1Edit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_Adnl2").Value = oTInfo2Edit.Value
                    ITools = PSSIT_PEY3.Update()
                Else
                    PSSIT_PEY3.Code = oTCodeEdit.Value
                    PSSIT_PEY3.Name = oTCodeEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_Prdentno").Value = oTPENoEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_Maclid").Value = oTMacLineIDEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_wcno").Value = oTMacCodeEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_Madcey").Value = oTMacDocEntryEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_Toolcode").Value = oToolCodeEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_TLname").Value = oToolDescEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_Qty").Value = oTQtyEdit.Value
                    If oTAccKey.Checked = True Then
                        PSSIT_PEY3.UserFields.Fields.Item("U_Acckey").Value = "Y"
                    Else
                        PSSIT_PEY3.UserFields.Fields.Item("U_Acckey").Value = "N"
                    End If
                    PSSIT_PEY3.UserFields.Fields.Item("U_Accode").Value = oTAcCodeEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_Acname").Value = oTAcNameEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_CAccode").Value = oTConAcCodeEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_CAcname").Value = oTConAcNameEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_Tlctppie").Value = oTToolCostEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_Totcost").Value = oTTotRunCostEdit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_Adnl1").Value = oTInfo1Edit.Value
                    PSSIT_PEY3.UserFields.Fields.Item("U_Adnl2").Value = oTInfo2Edit.Value
                    ITools = PSSIT_PEY3.Add()
                    Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim oRs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRs.DoQuery("Select * from [@PSSIT_OTLS] Where Code = '" & oToolCodeEdit.Value & "'")
                    If oRs.RecordCount > 0 Then
                        oRs1.DoQuery("Update [@PSSIT_OTLS] Set U_Cnou = " & (CDbl(oTQtyEdit.Value) + CDbl(oRs.Fields.Item("U_Cnou").Value)) _
                        & " Where Code = '" & oToolCodeEdit.Value & "'")
                    End If
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub AddLabourDatatoDB()
        Dim IntICount, ILab As Integer
        Dim oLAccKey, oLLabKey, oLParallel As SAPbouiCOM.CheckBox
        Dim oLCodeEdit, oLPENoEdit, oLMacLineIDEdit, oLMacDocEntryEdit, oLMacCodeEdit, oLabCodeEdit, oLSkGroupNameEdit, oLReqNoEdit, oLFromTimeEdit, oLToTimeEdit, oLWrkTimeEdit, oLQtyEdit, oLAcCodeEdit, oLAcNameEdit, oLConAcCodeEdit, oLConAcNameEdit, oLabRateEdit, oLTotCostEdit, oLInfo1Edit, oLInfo2Edit, oNop, oOTtime As SAPbouiCOM.EditText
        Dim oLSkGroupCodeEdit As SAPbouiCOM.ComboBox
        Try
            For IntICount = 1 To oLabMatrix.VisualRowCount
                oLCodeEdit = oLCodeCol.Cells.Item(IntICount).Specific
                oLPENoEdit = oLPENoCol.Cells.Item(IntICount).Specific
                oLMacLineIDEdit = oLMLineIDCol.Cells.Item(IntICount).Specific
                oLMacDocEntryEdit = oLMDocEntryCol.Cells.Item(IntICount).Specific
                oLMacCodeEdit = oLMacNoCol.Cells.Item(IntICount).Specific
                oLabCodeEdit = oLabCodeCol.Cells.Item(IntICount).Specific
                oLSkGroupCodeEdit = oLSkGroupCodeCol.Cells.Item(IntICount).Specific
                oLSkGroupNameEdit = oLSkGroupNameCol.Cells.Item(IntICount).Specific
                oLReqNoEdit = oLReqNosCol.Cells.Item(IntICount).Specific
                oLFromTimeEdit = oLFromTimeCol.Cells.Item(IntICount).Specific
                oLToTimeEdit = oLToTimeCol.Cells.Item(IntICount).Specific
                oLWrkTimeEdit = oLWrkTimeCol.Cells.Item(IntICount).Specific
                oLQtyEdit = oLQtyCol.Cells.Item(IntICount).Specific
                oLAcCodeEdit = oLAccCodeCol.Cells.Item(IntICount).Specific
                oLAcNameEdit = oLAccNameCol.Cells.Item(IntICount).Specific
                oLConAcCodeEdit = oLConAccCodeCol.Cells.Item(IntICount).Specific
                oLConAcNameEdit = oLConAccNameCol.Cells.Item(IntICount).Specific
                oLabRateEdit = oLabRateCol.Cells.Item(IntICount).Specific
                oLTotCostEdit = oLTotRunCstCol.Cells.Item(IntICount).Specific
                oLInfo1Edit = oLInfo1Col.Cells.Item(IntICount).Specific
                oLInfo2Edit = oLInfo2Col.Cells.Item(IntICount).Specific
                oLLabKey = oLabKeyCol.Cells.Item(IntICount).Specific
                oLParallel = oLParCol.Cells.Item(IntICount).Specific
                oLAccKey = oLAccKeyCol.Cells.Item(IntICount).Specific
                'Added by Manimaran-----s
                oOTtime = oLotTimeCol.Cells.Item(IntICount).Specific
                oNop = oNOPCol.Cells.Item(IntICount).Specific
                'Added by Manimaran-----e
                oLabMatrix.GetLineData(IntICount)
                '***************Get Account Code For Journal - Labour****************
                DataRow = oDataTable.NewRow()
                DataRow.Item("AcCode") = CStr(oLAcCodeEdit.Value)
                DataRow.Item("AcName") = CStr(oLAcNameEdit.Value)
                DataRow.Item("CAcCode") = CStr(UConAcCode.Value)
                DataRow.Item("CAcName") = CStr(UConAcName.Value)
                DataRow.Item("Debit") = 0
                DataRow.Item("Credit") = CDbl(oLTotCostEdit.Value)
                DataRow.Item("ShortName") = CStr(oLAcCodeEdit.Value)

                dblJournalCredit = dblJournalCredit + CDbl(oLTotCostEdit.Value)
                dblJournalDebit = dblJournalDebit + 0
                oDataTable.Rows.Add(DataRow)
                '***********************************************************
                If PSSIT_PEY2.GetByKey(oLCodeEdit.Value) = True Then
                    PSSIT_PEY2.Code = oLCodeEdit.Value
                    PSSIT_PEY2.Name = oLCodeEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_Prdentno").Value = oLPENoEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_Maclid").Value = oLMacLineIDEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_Madcey").Value = oLMacDocEntryEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_wcno").Value = oLMacCodeEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_Lrcode").Value = oLabCodeEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_LGCode").Value = oLSkGroupCodeEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_LGname").Value = oLSkGroupNameEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_Reqno").Value = oLReqNoEdit.Value
                    If oLLabKey.Checked = True Then
                        PSSIT_PEY2.UserFields.Fields.Item("U_Labkey").Value = "Y"
                    Else
                        PSSIT_PEY2.UserFields.Fields.Item("U_Labkey").Value = "N"
                    End If
                    If oLParallel.Checked = True Then
                        PSSIT_PEY2.UserFields.Fields.Item("U_Parallel").Value = "Y"
                    Else
                        PSSIT_PEY2.UserFields.Fields.Item("U_Parallel").Value = "N"
                    End If
                    PSSIT_PEY2.UserFields.Fields.Item("U_Frtime").Value = oLFromTimeEdit.String
                    PSSIT_PEY2.UserFields.Fields.Item("U_Totime").Value = oLToTimeEdit.String
                    PSSIT_PEY2.UserFields.Fields.Item("U_Wrktime").Value = oLWrkTimeEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_Qty").Value = oLQtyEdit.Value
                    If oLAccKey.Checked = True Then
                        PSSIT_PEY2.UserFields.Fields.Item("U_Acckey").Value = "Y"
                    Else
                        PSSIT_PEY2.UserFields.Fields.Item("U_Acckey").Value = "N"
                    End If
                    PSSIT_PEY2.UserFields.Fields.Item("U_Accode").Value = oLAcCodeEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_Acname").Value = oLAcNameEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_CAccode").Value = oLConAcCodeEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_CAcname").Value = oLConAcNameEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_Lrtph").Value = oLabRateEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_Totcost").Value = oLTotCostEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_Adnl1").Value = oLInfo1Edit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_Adnl2").Value = oLInfo2Edit.Value
                    'Added by Manimaran----s
                    PSSIT_PEY2.UserFields.Fields.Item("U_OTtime").Value = oOTtime.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_Nop").Value = oNop.Value
                    'Added by Manimaran----e
                    ILab = PSSIT_PEY2.Update()
                Else
                    PSSIT_PEY2.Code = oLCodeEdit.Value
                    PSSIT_PEY2.Name = oLCodeEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_Prdentno").Value = oLPENoEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_Maclid").Value = oLMacLineIDEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_Madcey").Value = oLMacDocEntryEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_wcno").Value = oLMacCodeEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_Lrcode").Value = oLabCodeEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_LGCode").Value = oLSkGroupCodeEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_LGname").Value = oLSkGroupNameEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_Reqno").Value = oLReqNoEdit.Value
                    If oLLabKey.Checked = True Then
                        PSSIT_PEY2.UserFields.Fields.Item("U_Labkey").Value = "Y"
                    Else
                        PSSIT_PEY2.UserFields.Fields.Item("U_Labkey").Value = "N"
                    End If
                    If oLParallel.Checked = True Then
                        PSSIT_PEY2.UserFields.Fields.Item("U_Parallel").Value = "Y"
                    Else
                        PSSIT_PEY2.UserFields.Fields.Item("U_Parallel").Value = "N"
                    End If
                    PSSIT_PEY2.UserFields.Fields.Item("U_Frtime").Value = oLFromTimeEdit.String
                    PSSIT_PEY2.UserFields.Fields.Item("U_Totime").Value = oLToTimeEdit.String
                    PSSIT_PEY2.UserFields.Fields.Item("U_Wrktime").Value = oLWrkTimeEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_Qty").Value = oLQtyEdit.Value
                    If oLAccKey.Checked = True Then
                        PSSIT_PEY2.UserFields.Fields.Item("U_Acckey").Value = "Y"
                    Else
                        PSSIT_PEY2.UserFields.Fields.Item("U_Acckey").Value = "N"
                    End If
                    PSSIT_PEY2.UserFields.Fields.Item("U_Accode").Value = oLAcCodeEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_Acname").Value = oLAcNameEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_CAccode").Value = oLConAcCodeEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_CAcname").Value = oLConAcNameEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_Lrtph").Value = oLabRateEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_Totcost").Value = oLTotCostEdit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_Adnl1").Value = oLInfo1Edit.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_Adnl2").Value = oLInfo2Edit.Value
                    'Added by Manimaran----s
                    PSSIT_PEY2.UserFields.Fields.Item("U_OTtime").Value = oOTtime.Value
                    PSSIT_PEY2.UserFields.Fields.Item("U_Nop").Value = oNop.Value
                    'Added by Manimaran----e
                    ILab = PSSIT_PEY2.Add()
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub AddFixedCostDatatoDB()
        Dim oFCodeEdit, oFMacCodeEdit, oFWrkCentreCodeEdit, oFFixedCostEdit, oFUnitCostEdit As SAPbouiCOM.EditText
        Dim oFAbsMthdEdit, oFActCodeEdit, oFActNameEdit, oFTotCostEdit, oFPOSerEdit, oFPONoEdit, oFPENoEdit As SAPbouiCOM.EditText
        Dim IntICount, IFixCst As Integer
        Dim oStrSql As String
        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            For IntICount = 1 To oFCMatrix.RowCount
                oFCodeEdit = oFCodeCol.Cells.Item(IntICount).Specific
                oFPONoEdit = oFPONoCol.Cells.Item(IntICount).Specific
                oFPENoEdit = oFPENoCol.Cells.Item(IntICount).Specific
                oFPOSerEdit = oFPOSerCol.Cells.Item(IntICount).Specific
                oFCodeEdit = oFCodeCol.Cells.Item(IntICount).Specific
                oFMacCodeEdit = oFMacCodeCol.Cells.Item(IntICount).Specific
                oFWrkCentreCodeEdit = oFWrkCentreCodeCol.Cells.Item(IntICount).Specific
                oFFixedCostEdit = oFFixedCostCol.Cells.Item(IntICount).Specific
                oFUnitCostEdit = oFUnitCostCol.Cells.Item(IntICount).Specific
                oFAbsMthdEdit = oFAbsMthdCol.Cells.Item(IntICount).Specific
                oFActCodeEdit = oFActCodeCol.Cells.Item(IntICount).Specific
                oFActNameEdit = oFActNameCol.Cells.Item(IntICount).Specific
                oFTotCostEdit = oFTotCostCol.Cells.Item(IntICount).Specific
                oFCMatrix.GetLineData(IntICount)
                '***************Get Account Code For Journal - Labour****************
                DataRow = oDataTable.NewRow()
                DataRow.Item("AcCode") = CStr(oFActCodeEdit.Value)
                DataRow.Item("AcName") = CStr(oFActNameEdit.Value)
                DataRow.Item("CAcCode") = CStr(UConAcCode.Value)
                DataRow.Item("CAcName") = CStr(UConAcName.Value)
                DataRow.Item("Debit") = 0
                DataRow.Item("Credit") = CDbl(oFTotCostEdit.Value)
                DataRow.Item("ShortName") = CStr(oFActCodeEdit.Value)

                dblJournalCredit = dblJournalCredit + CDbl(oFTotCostEdit.Value)
                dblJournalDebit = dblJournalDebit + 0
                oDataTable.Rows.Add(DataRow)
                If PSSIT_PEY4.GetByKey(oFCodeEdit.Value) = True Then
                    PSSIT_PEY4.Code = oFCodeEdit.Value
                    PSSIT_PEY4.Name = oFCodeEdit.Value
                    PSSIT_PEY4.UserFields.Fields.Item("U_Pordser").Value = oFPOSerEdit.Value
                    PSSIT_PEY4.UserFields.Fields.Item("U_Pordno").Value = oFPONoEdit.Value
                    PSSIT_PEY4.UserFields.Fields.Item("U_Prdentno").Value = oFPENoEdit.Value
                    PSSIT_PEY4.UserFields.Fields.Item("U_wcno").Value = oFMacCodeEdit.Value
                    PSSIT_PEY4.UserFields.Fields.Item("U_Wrkno").Value = oFWrkCentreCodeEdit.Value
                    PSSIT_PEY4.UserFields.Fields.Item("U_Fcost").Value = oFFixedCostEdit.Value
                    PSSIT_PEY4.UserFields.Fields.Item("U_UnitCost").Value = oFUnitCostEdit.Value
                    PSSIT_PEY4.UserFields.Fields.Item("U_Absmthd").Value = oFAbsMthdEdit.Value
                    PSSIT_PEY4.UserFields.Fields.Item("U_Accode").Value = oFActCodeEdit.Value
                    PSSIT_PEY4.UserFields.Fields.Item("U_Acname").Value = oFActNameEdit.Value
                    PSSIT_PEY4.UserFields.Fields.Item("U_Totfcst").Value = oFTotCostEdit.Value
                    IFixCst = PSSIT_PEY2.Update()
                Else
                    PSSIT_PEY4.Code = oFCodeEdit.Value
                    PSSIT_PEY4.Name = oFCodeEdit.Value
                    PSSIT_PEY4.UserFields.Fields.Item("U_Pordser").Value = oFPOSerEdit.Value
                    PSSIT_PEY4.UserFields.Fields.Item("U_Pordno").Value = oFPONoEdit.Value
                    PSSIT_PEY4.UserFields.Fields.Item("U_Prdentno").Value = oFPENoEdit.Value
                    PSSIT_PEY4.UserFields.Fields.Item("U_wcno").Value = oFMacCodeEdit.Value
                    PSSIT_PEY4.UserFields.Fields.Item("U_Wrkno").Value = oFWrkCentreCodeEdit.Value
                    PSSIT_PEY4.UserFields.Fields.Item("U_Fcost").Value = oFFixedCostEdit.Value
                    PSSIT_PEY4.UserFields.Fields.Item("U_UnitCost").Value = oFUnitCostEdit.Value
                    PSSIT_PEY4.UserFields.Fields.Item("U_Absmthd").Value = oFAbsMthdEdit.Value
                    PSSIT_PEY4.UserFields.Fields.Item("U_Accode").Value = oFActCodeEdit.Value
                    PSSIT_PEY4.UserFields.Fields.Item("U_Acname").Value = oFActNameEdit.Value
                    PSSIT_PEY4.UserFields.Fields.Item("U_Totfcst").Value = oFTotCostEdit.Value
                    IFixCst = PSSIT_PEY4.Add()
                    oStrSql = "Update [@PSSIT_WOR4] Set " _
                   & "U_Totfcst = U_TotFcst + " & (CDbl(oFTotCostEdit.Value)) _
                   & " Where U_Pordser = " & oFPOSerEdit.Value & " and U_Pordno = " & oFPONoEdit.Value _
                   & " and U_FCost = '" & oFFixedCostEdit.Value & "'"
                    oRs.DoQuery(oStrSql)
                End If
            Next
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
    Private Sub DataTableFieldCreation()
        Try
            oDataTable = New DataTable("Journal")
            oDataTable.Columns.Add("DueDate", Type.GetType("System.String"))
            oDataTable.Columns.Add("AcCode", Type.GetType("System.String"))
            oDataTable.Columns.Add("AcName", Type.GetType("System.String"))
            oDataTable.Columns.Add("CAcCode", Type.GetType("System.String"))
            oDataTable.Columns.Add("CAcName", Type.GetType("System.String"))
            oDataTable.Columns.Add("Debit", Type.GetType("System.Double"))
            oDataTable.Columns.Add("Credit", Type.GetType("System.Double"))
            oDataTable.Columns.Add("RefDate1", Type.GetType("System.String"))
            oDataTable.Columns.Add("ShortName", Type.GetType("System.String"))
            oDataTable.Clear()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    'add by kabilahan b
    Private Function validateTime(ByVal dt As Date, ByVal pVal As SAPbouiCOM.ItemEvent) As Boolean
        Dim qry As String
        Dim rs As SAPbobsCOM.Recordset
        rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim dtdat As Date
        dtdat = getdatetime(oForm.Items.Item("txtdocdt").Specific.string)

        'MsgBox(dtdat)

        'If dtdat > Now() Then
        '    SBO_Application.SetStatusBarMessage("Date Should not be greater than Current date.....", SAPbouiCOM.BoMessageTime.bmt_Short, True)
        '    Return False
        'End If

        'If dt.Day <= 12 Then
        '    qry = "select t2.u_frtime,t2.u_totime from [@PSSIT_OPEY] t1"
        '    qry = qry + " inner join [@PSSIT_PEY1] t2 on t1.docentry = t2.docentry"
        '    qry = qry + " where t1.u_scode = '" & oForm.Items.Item("txtSftCode").Specific.string & "' and t2.u_wcno = '" & oForm.Items.Item("txtmcno").Specific.string & "' and  convert(varchar,t1.u_docdt,101) = '" & Left(String.Format(dt, "dd/mm/yyyy"), 6) & dt.Year & "'"
        'Else
        '    qry = "select t2.u_frtime,t2.u_totime from [@PSSIT_OPEY] t1"
        '    qry = qry + " inner join [@PSSIT_PEY1] t2 on t1.docentry = t2.docentry"
        '    qry = qry + " where t1.u_scode = '" & oForm.Items.Item("txtSftCode").Specific.string & "' and t2.u_wcno = '" & oForm.Items.Item("txtmcno").Specific.string & "' and  convert(varchar,t1.u_docdt,103) = '" & Left(String.Format(dt, "dd/mm/yyyy"), 6) & dt.Year & "'"
        'End If

        qry = "set dateformat dmy select t2.u_stTime,t2.u_EndTime from [@PSSIT_PMWCBREAKHDR] t1 "
        qry = qry + " inner join [@PSSIT_PMWCREASONDTL] t2 on t1.docentry = t2.docentry"
        'modified by Manimaran-----s
        'qry = qry + " where t1.u_scode = '" & oForm.Items.Item("txtscode").Specific.string & "' and t1.U_deptcode = '" & oMacMatrix.Columns.Item("colwcno").Cells.Item(pVal.row).Specific.value & "' and t1.U_Docdate = convert(varchar,'" & oForm.Items.Item("txtpordt").Specific.string & "',102)"

        ' qry = qry + " where t1.u_scode = '" & oForm.Items.Item("txtscode").Specific.string & "' and t1.U_deptcode = '" & oMacMatrix.Columns.Item("colwcno").Cells.Item(pVal.Row).Specific.value & "' and t1.U_Docdate = convert(varchar,'" & oForm.Items.Item("txtdocdt").Specific.string & "',102)"
        qry = qry + " where t1.u_scode = '" & oForm.Items.Item("txtscode").Specific.string & "' and t1.U_deptcode = '" & oMacMatrix.Columns.Item("colwcno").Cells.Item(pVal.Row).Specific.value & "' and t1.U_Docdate = '" & dtdat.ToString("dd/MM/yyyy") & "'"
        qry = qry + " and t1.u_penum = '" & oForm.Items.Item("txtprdno").Specific.string & "'"
        'Modified by Manimaran-----e


        rs.DoQuery(qry)
        If rs.RecordCount > 0 Then
            'Modified by Manimaran-----s
            'If rs.Fields.Item(0).Value < rs.Fields.Item(1).Value Then
            '    If (oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.row).Specific.value >= rs.Fields.Item(0).Value And oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.row).Specific.value >= rs.Fields.Item(0).Value) Or (oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.row).Specific.value >= rs.Fields.Item(1).Value And oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.row).Specific.value >= rs.Fields.Item(1).Value) Then
            '        If (rs.Fields.Item(0).Value > oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.row).Specific.value) Or (rs.Fields.Item(1).Value > oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.row).Specific.value) Or (rs.Fields.Item(0).Value > oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.row).Specific.value) Or (rs.Fields.Item(1).Value > oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.row).Specific.value) Then
            '            SBO_Application.SetStatusBarMessage("Overlapping with the Machine run time", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            '            oMacMatrix.Columns.Item("colrntim").Cells.Item(pVal.row).Specific.value = 0
            '            Return False
            '        End If
            '    ElseIf (oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.row).Specific.value <= rs.Fields.Item(0).Value And oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.row).Specific.value <= rs.Fields.Item(0).Value) Or (oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.row).Specific.value <= rs.Fields.Item(1).Value And oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.row).Specific.value <= rs.Fields.Item(1).Value) Then
            '        If (rs.Fields.Item(0).Value < oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.row).Specific.value) Or (rs.Fields.Item(1).Value < oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.row).Specific.value) Or (rs.Fields.Item(0).Value < oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.row).Specific.value) Or (rs.Fields.Item(1).Value < oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.row).Specific.value) Then
            '            SBO_Application.SetStatusBarMessage("Overlapping with the Machine run time", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            '            oMacMatrix.Columns.Item("colrntim").Cells.Item(pVal.row).Specific.value = 0
            '            Return False
            '        End If
            '    End If
            'Else
            '    If (oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.row).Specific.value <= rs.Fields.Item(0).Value And oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.row).Specific.value <= rs.Fields.Item(0).Value) Or (oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.row).Specific.value <= rs.Fields.Item(1).Value And oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.row).Specific.value <= rs.Fields.Item(1).Value) Then
            '        If (rs.Fields.Item(0).Value < oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.row).Specific.value) Or (rs.Fields.Item(1).Value > oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.row).Specific.value) Or (rs.Fields.Item(0).Value < oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.row).Specific.value) Or (rs.Fields.Item(1).Value > oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.row).Specific.value) Then
            '            SBO_Application.SetStatusBarMessage("Overlapping with the Machine run time", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            '            oMacMatrix.Columns.Item("colrntim").Cells.Item(pVal.row).Specific.value = 0
            '            Return False
            '        End If
            '    ElseIf (oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.row).Specific.value >= rs.Fields.Item(0).Value And oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.row).Specific.value >= rs.Fields.Item(0).Value) Or (oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.row).Specific.value >= rs.Fields.Item(1).Value And oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.row).Specific.value >= rs.Fields.Item(1).Value) Then
            '        If (rs.Fields.Item(0).Value > oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.row).Specific.value) Or (rs.Fields.Item(1).Value < oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.row).Specific.value) Or (rs.Fields.Item(0).Value > oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.row).Specific.value) Or (rs.Fields.Item(1).Value < oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.row).Specific.value) Then
            '            SBO_Application.SetStatusBarMessage("Overlapping with the Machine run time", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            '            oMacMatrix.Columns.Item("colrntim").Cells.Item(pVal.row).Specific.value = 0
            '            Return False
            '        End If
            '    End If
            'End If
            While Not rs.EoF
                If rs.Fields.Item(0).Value < rs.Fields.Item(1).Value Then
                    If (oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value >= rs.Fields.Item(0).Value Or oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value >= rs.Fields.Item(0).Value) Or (oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value >= rs.Fields.Item(1).Value Or oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value >= rs.Fields.Item(1).Value) Then
                        If (rs.Fields.Item(0).Value > oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value) Or (rs.Fields.Item(1).Value > oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value) Or (rs.Fields.Item(0).Value > oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value) Or (rs.Fields.Item(1).Value > oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value) Then
                            SBO_Application.SetStatusBarMessage("Overlapping with the stoppage time", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            oMacMatrix.Columns.Item("colrntim").Cells.Item(pVal.Row).Specific.value = 0
                            Return False
                        End If
                    ElseIf (oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value <= rs.Fields.Item(0).Value Or oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value <= rs.Fields.Item(0).Value) Or (oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value <= rs.Fields.Item(1).Value Or oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value <= rs.Fields.Item(1).Value) Then
                        If (rs.Fields.Item(0).Value < oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value) Or (rs.Fields.Item(1).Value < oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value) Or (rs.Fields.Item(0).Value < oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value) Or (rs.Fields.Item(1).Value < oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value) Then
                            SBO_Application.SetStatusBarMessage("Overlapping with the stoppage time", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            oMacMatrix.Columns.Item("colrntim").Cells.Item(pVal.Row).Specific.value = 0
                            Return False
                        End If
                    End If
                Else
                    If (oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value <= rs.Fields.Item(0).Value Or oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value <= rs.Fields.Item(0).Value) Or (oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value <= rs.Fields.Item(1).Value Or oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value <= rs.Fields.Item(1).Value) Then
                        If (rs.Fields.Item(0).Value < oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value) Or (rs.Fields.Item(1).Value > oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value) Or (rs.Fields.Item(0).Value < oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value) Or (rs.Fields.Item(1).Value > oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value) Then
                            SBO_Application.SetStatusBarMessage("Overlapping with the stoppage time", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            oMacMatrix.Columns.Item("colrntim").Cells.Item(pVal.Row).Specific.value = 0
                            Return False
                        End If
                    ElseIf (oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value >= rs.Fields.Item(0).Value Or oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value >= rs.Fields.Item(0).Value) Or (oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value >= rs.Fields.Item(1).Value Or oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value >= rs.Fields.Item(1).Value) Then
                        If (rs.Fields.Item(0).Value > oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value) Or (rs.Fields.Item(1).Value < oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value) Or (rs.Fields.Item(0).Value > oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value) Or (rs.Fields.Item(1).Value < oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value) Then
                            SBO_Application.SetStatusBarMessage("Overlapping with the stoppage time", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            oMacMatrix.Columns.Item("colrntim").Cells.Item(pVal.Row).Specific.value = 0
                            Return False
                        End If
                    End If
                End If
                rs.MoveNext()
            End While
            rs = Nothing
            rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            qry = ""
            qry = "set dateformat dmy select t2.u_frtime,t2.u_totime from [@PSSIT_OPEY] t1"
            qry = qry + " inner join [@PSSIT_PEY1] t2 on t1.docentry = t2.docentry"
            qry = qry + " where t1.u_scode = '" & oForm.Items.Item("txtscode").Specific.string & "' and t2.u_wcno = '" & oMacMatrix.Columns.Item("colwcno").Cells.Item(pVal.Row).Specific.value & "' and t1.U_Docdt = '" & dtdat.ToString("dd/MM/yyyy") & "'"
            qry = qry + " and t1.u_pnordno = '" & oForm.Items.Item("txtprdno").Specific.string & "'"
            rs.DoQuery(qry)
            While Not rs.EoF
                If rs.Fields.Item(0).Value < rs.Fields.Item(1).Value Then
                    If (oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value >= rs.Fields.Item(0).Value Or oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value >= rs.Fields.Item(0).Value) Or (oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value >= rs.Fields.Item(1).Value Or oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value >= rs.Fields.Item(1).Value) Then
                        If (rs.Fields.Item(0).Value > oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value) Or (rs.Fields.Item(1).Value > oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value) Or (rs.Fields.Item(0).Value > oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value) Or (rs.Fields.Item(1).Value > oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value) Then
                            SBO_Application.SetStatusBarMessage("Overlapping with the previous machine time", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            oMacMatrix.Columns.Item("colrntim").Cells.Item(pVal.Row).Specific.value = 0
                            Return False
                        End If
                    ElseIf (oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value <= rs.Fields.Item(0).Value Or oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value <= rs.Fields.Item(0).Value) Or (oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value <= rs.Fields.Item(1).Value Or oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value <= rs.Fields.Item(1).Value) Then
                        If (rs.Fields.Item(0).Value < oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value) Or (rs.Fields.Item(1).Value < oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value) Or (rs.Fields.Item(0).Value < oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value) Or (rs.Fields.Item(1).Value < oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value) Then
                            SBO_Application.SetStatusBarMessage("Overlapping with the previous machine time", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            oMacMatrix.Columns.Item("colrntim").Cells.Item(pVal.Row).Specific.value = 0
                            Return False
                        End If
                    End If
                Else
                    If (oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value <= rs.Fields.Item(0).Value Or oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value <= rs.Fields.Item(0).Value) Or (oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value <= rs.Fields.Item(1).Value Or oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value <= rs.Fields.Item(1).Value) Then
                        If (rs.Fields.Item(0).Value < oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value) Or (rs.Fields.Item(1).Value > oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value) Or (rs.Fields.Item(0).Value < oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value) Or (rs.Fields.Item(1).Value > oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value) Then
                            SBO_Application.SetStatusBarMessage("Overlapping with the previous machine time", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            oMacMatrix.Columns.Item("colrntim").Cells.Item(pVal.Row).Specific.value = 0
                            Return False
                        End If
                    ElseIf (oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value >= rs.Fields.Item(0).Value Or oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value >= rs.Fields.Item(0).Value) Or (oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value >= rs.Fields.Item(1).Value Or oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value >= rs.Fields.Item(1).Value) Then
                        If (rs.Fields.Item(0).Value > oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value) Or (rs.Fields.Item(1).Value < oMacMatrix.Columns.Item("colfrtim").Cells.Item(pVal.Row).Specific.value) Or (rs.Fields.Item(0).Value > oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value) Or (rs.Fields.Item(1).Value < oMacMatrix.Columns.Item("coltotim").Cells.Item(pVal.Row).Specific.value) Then
                            SBO_Application.SetStatusBarMessage("Overlapping with the previous machine time", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            oMacMatrix.Columns.Item("colrntim").Cells.Item(pVal.Row).Specific.value = 0
                            Return False
                        End If
                    End If
                End If
                rs.MoveNext()
            End While
            'modified by Manimaran-----e
        Else
            'SBO_Application.SetStatusBarMessage("Machine Master Record not found", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return False
        End If
        Return True
    End Function
    Private Function shiftTimeValidation(ByRef pVal As SAPbouiCOM.ItemEvent) As Boolean
        Dim QToTime As Integer
        Dim QFrtime As Integer
        If oMacMatrix.Columns.Item("colfrtim").Cells.Item(oMacMatrix.RowCount - 1).Specific.string <> "" Then
            ofrTime = CInt(oMacMatrix.Columns.Item("colfrtim").Cells.Item(oMacMatrix.RowCount - 1).Specific.value)
        End If
        If oMacMatrix.Columns.Item("coltotim").Cells.Item(oMacMatrix.RowCount - 1).Specific.string <> "" Then
            oToTime = CInt(oMacMatrix.Columns.Item("coltotim").Cells.Item(oMacMatrix.RowCount - 1).Specific.value)
        End If

        sQry = "Select * from [@PSSIT_OSFT] where code = '" & oForm.Items.Item("txtscode").Specific.string & "'"
        'sQry = "select a.U_StTime,a.U_EndTime from [@PSSIT_PMWCREASONDTL] a join [@PSSIT_PMWCBREAKHDR] b "
        'sQry += "on a.DocEntry = b.DocEntry where b.U_docdate ='2010-10-18' and b.U_deptcode = 'M1' and U_SCode ='GS1'"
        Rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Rs.DoQuery(sQry)
        If Rs.RecordCount > 0 Then
            QFrtime = Integer.Parse(Rs.Fields.Item("U_SfTime").Value.ToString)
            QToTime = Integer.Parse(Rs.Fields.Item("U_StTime").Value.ToString)
        End If
        If pVal.ColUID = "colfrtim" Then
            If QFrtime < QToTime Then
                If ofrTime < QFrtime Then
                    SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    oMacMatrix.Columns.Item("colfrtim").Cells.Item(oMacMatrix.RowCount).Specific.value = 0
                    Return False
                ElseIf ofrTime > QToTime Then
                    SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    oMacMatrix.Columns.Item("coltotim").Cells.Item(oMacMatrix.RowCount).Specific.value = 0
                    Return False
                End If
            Else
                If QToTime <= ofrTime Then
                    If ofrTime < QFrtime Then
                        SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range  ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oMacMatrix.Columns.Item("colfrtim").Cells.Item(oMacMatrix.RowCount).Specific.value = 0
                        Return False
                    ElseIf ofrTime < QToTime Then
                        SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oMacMatrix.Columns.Item("coltotim").Cells.Item(oMacMatrix.RowCount).Specific.value = 0
                        Return False
                    End If
                End If
            End If
        End If
        If pVal.ColUID = "coltotim" Then
            If QFrtime < QToTime Then
                If oToTime < QFrtime Then
                    SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range  ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    oMacMatrix.Columns.Item("coltotim").Cells.Item(oMacMatrix.RowCount).Specific.value = 0
                    Return False
                ElseIf oToTime > QToTime Then
                    SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    oMacMatrix.Columns.Item("coltotim").Cells.Item(oMacMatrix.RowCount).Specific.value = 0
                    Return False
                End If
            Else
                If ofrTime > oToTime Then
                    If oToTime > QFrtime Then
                        SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oMacMatrix.Columns.Item("coltotim").Cells.Item(oMacMatrix.RowCount).Specific.value = 0
                        Return False
                    ElseIf oToTime > QToTime Then
                        SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oMacMatrix.Columns.Item("coltotim").Cells.Item(oMacMatrix.RowCount).Specific.value = 0
                        Return False
                    End If
                Else
                    If ofrTime < oToTime And QToTime < oToTime And QFrtime > ofrTime Then
                        SBO_Application.SetStatusBarMessage("Time should be between the selected shift time range  ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oMacMatrix.Columns.Item("coltotim").Cells.Item(oMacMatrix.RowCount).Specific.value = 0
                        Return False
                    End If
                End If
            End If
        End If
        Return True
    End Function
    ' kabilahan e
End Class
