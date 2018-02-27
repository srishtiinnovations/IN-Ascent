'''' <summary>
'''' Author                     Created Date
'''' Suresh                      23/12/2008
'''' <remarks> This class is used for viewing Master Reports List.</remarks>
Public Class MasterReports
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
    '**************************Items - Button************************************
    Private BtnShift, BtnWorkCentre, BtnMCGroups, BtnMCMaster, BtnSkillGroups, BtnLabour, BtnTools, BtnStoppage As SAPbouiCOM.Button
    Private WithEvents ShiftReportClass As ShiftReport
    Private WithEvents WCReportClass As WorkCentreReport
    Private WithEvents MGReportClass As MachineGroupsReport
    Private WithEvents MCReporttClass As MachineReport
    Private WithEvents SGReporttClass As SkillGroupReport
    Private WithEvents LBReporttClass As LabourReport
    Private WithEvents TLReporttClass As ToolsReport
    Private WithEvents STReporttClass As StoppageReport
    Private WithEvents OprReporttClass As OperationReport
    Private WithEvents OprRouteReporttClass As OperationRoutingReport
#End Region
    ''' <summary>
    ''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmMasterReports.srf") method is called to load the Master Reports form.
    ''' Drawform() method is called to Initialize the form and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company)
        MyBase.New(aSBO_Application, aCompany)
        SBO_Application = aSBO_Application
        oCompany = aCompany
        SetApplication()
        LoadFromXML("FrmMasterReports.srf")
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
            InitializeFormComponent()
            oForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Configuring the items/controls in the form(.srf) by bounding to the object.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitializeFormComponent()
        Try
            BtnShift = oForm.Items.Item("btnshift").Specific
            BtnWorkCentre = oForm.Items.Item("btnwrkcntr").Specific
            BtnMCGroups = oForm.Items.Item("btnmcgroup").Specific
            BtnMCMaster = oForm.Items.Item("btnmcmastr").Specific
            BtnSkillGroups = oForm.Items.Item("btnskilgrp").Specific
            BtnLabour = oForm.Items.Item("btnlabour").Specific
            BtnTools = oForm.Items.Item("btntools").Specific
            BtnStoppage = oForm.Items.Item("btnstopage").Specific
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
      
        Try
            If pVal.FormUID = "FMR" Then
                '*****************************Releasing the Com Object*******************************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                    SBO_Application = Nothing
                    GC.Collect()
                End If
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) Then
                    '*************Shift Button Press*****************
                    If (pVal.ItemUID = "btnshift") And (pVal.BeforeAction = False) Then
                        Dim StrSql As String
                        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Try
                            'StrSql = "select Code,U_Sdescr,U_Sftime,U_Sttime,U_Sbreak,U_Duratmin,U_Durathrs,U_Adnl1,U_Adnl2,U_Remarks from [@PSSIT_OSFT]"
                            StrSql = "select Code,U_Sdescr,(case when len(U_Sftime)=4 then " _
                            & "((Substring(cast(U_Sftime as varchar(5)),1,2))+ ':' +(Substring(cast(U_Sftime as varchar(5)),3,2)) ) " _
                            & "when len(U_Sftime)=2 then  " _
                            & "('00'+ ':' +(Substring(cast(U_Sftime as varchar(5)),1,2)) )   " _
                            & "when len(U_Sftime)=1 then  " _
                            & "('00'+ ':' +'0'+(Substring(cast(U_Sftime as varchar(5)),1,1)) ) " _
                            & "else  " _
                            & "(Substring(cast(U_Sftime as varchar(5)),1,1))+ ':' +(Substring(cast(U_Sftime as varchar(5)),2,2)) " _
                            & "end) as 'FShift', " _
                            & "(case when len(U_Sttime)=4 then  " _
                            & "(Substring(cast(U_Sttime as varchar(5)),1,2))+ ':' +(Substring(cast(U_Sttime as varchar(5)),3,2)) " _
                            & "when len(U_Sttime)=2 then " _
                            & "('00'+ ':' +(Substring(cast(U_Sttime as varchar(5)),1,2)) )  " _
                            & "when len(U_Sttime)=1 then " _
                            & "('00'+ ':' +'0'+(Substring(cast(U_Sttime as varchar(5)),1,1)) )" _
                            & "else " _
                            & "(Substring(cast(U_Sttime as varchar(5)),1,1))+ ':' +(Substring(cast(U_Sttime as varchar(5)),2,2)) " _
                            & "end) as 'TShift' " _
                            & ",U_Sbreak,U_Duratmin, " _
                            & "(case when len(U_Durathrs)=4 then  " _
                            & "(Substring(cast(U_Durathrs as varchar(5)),1,2))+ ':' +(Substring(cast(U_Durathrs as varchar(5)),3,2)) " _
                            & "when len(U_Durathrs)=2 then " _
                            & "('00'+ ':' +(Substring(cast(U_Durathrs as varchar(5)),1,2)))  " _
                            & "when len(U_Durathrs)=1 then " _
                            & "('00'+ ':' +'0'+(Substring(cast(U_Durathrs as varchar(5)),1,1)))" _
                            & "else " _
                            & "(Substring(cast(U_Durathrs as varchar(5)),1,1))+ ':' +(Substring(cast(U_Durathrs as varchar(5)),2,2)) " _
                            & "end) as 'TShift' " _
                            & ",U_Adnl1,U_Adnl2,U_Remarks from [@PSSIT_OSFT] "
                            ShiftReportClass = New ShiftReport(SBO_Application, oCompany, StrSql)

                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        End Try
                    End If
                    '*************Work Centre Button Press*****************
                    If (pVal.ItemUID = "btnwrkcntr") And (pVal.BeforeAction = False) Then
                        Dim StrSql As String
                        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Try
                            StrSql = "select a.Code,a.U_WCname,a.U_WCtype,a.U_Adnl1,a.U_Adnl2, " _
                                     & "a.U_Remarks,b.U_Fcost,b.U_Currency,b.U_UnitCost,b.U_Absmthd, " _
                                     & "b.U_Accode, b.U_Acname, b.U_Adnl1 from [@PSSIT_OWCR] a, " _
                                     & "[@PSSIT_WCR1] b where a.code=b.code"
                            WCReportClass = New WorkCentreReport(SBO_Application, oCompany, StrSql)

                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        End Try
                    End If
                    '*************Machine Group Button Press*****************
                    If (pVal.ItemUID = "btnmcgroup") And (pVal.BeforeAction = False) Then
                        Dim StrSql As String
                        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Try
                            StrSql = "select Code as 'Machine Group Code',U_MGname as 'Machine Group Name',U_WCcode as 'Work Centre Code',U_WCName as 'Work Centre Name',U_RCurrncy as 'Running Rate Currency',U_Runrate as 'Running Rate/Hour', " _
                                     & "U_RAccode as 'Running Rate Acct Code',U_RAcname as 'Running Rate Acct Name',U_SCurrncy as 'Setup Rate Currency', U_Setrate as 'Setup Rate/Hour',U_SAccode as 'Setup Rate Acct Code',U_SAcname as 'Setup Rate Acct Name', " _
                                     & "U_Adnl1 as 'Info1',U_Adnl2 as 'Info2',U_Remarks as 'Remarks' from [@PSSIT_OMGP]"
                            MGReportClass = New MachineGroupsReport(SBO_Application, oCompany, StrSql)
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        End Try
                    End If
                    '*************Machine Master Button Press*****************
                    If (pVal.ItemUID = "btnmcmastr") And (pVal.BeforeAction = False) Then
                        Dim StrSql As String
                        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Try
                            StrSql = "select U_wcno as 'Machine Code',U_wcname as 'Machine Name',U_mfserial as 'MF.S.No',U_makecode as 'Make Code',U_makedesc as 'Make Name', " _
                                     & "U_modecode as 'Model Code',U_modedesc as 'Mode Name',U_deptcode as 'Work Center Code',U_deptdesc as 'Work Centre Name',U_MGcode as 'Machine Group Code', " _
                                     & "U_MGname as 'Machine Group Name',U_insdate as 'Installation Date',U_bpcode as 'Business Partner Code',U_wardate as 'Warrenty date' from [@PSSIT_PMWCHDR]"
                            MCReporttClass = New MachineReport(SBO_Application, oCompany, StrSql)
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        End Try
                    End If
                    '*************Skill Group Button Press*****************
                    If (pVal.ItemUID = "btnskilgrp") And (pVal.BeforeAction = False) Then
                        Dim StrSql As String
                        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Try
                            StrSql = "select  Code as 'Skill Group',U_LGname  as 'Skill Group Name', " _
& "U_WCcode  as 'Work Centre Code',U_WCName  as 'Work Centre Name', " _
& "U_Currncy  as 'Currency',U_Labrate  as 'Labour Rate/Hour', " _
& "U_Accode  as 'Acct Code',U_Acname  as 'Acct name', " _
& "U_Adnl1  as 'Info1',U_Adnl2  as 'Info2',U_Remarks  as 'Remarks' from [@PSSIT_OLGP]"
                            SGReporttClass = New SkillGroupReport(SBO_Application, oCompany, StrSql)
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        End Try
                    End If
                    '*************Labour Button Press*****************
                    If (pVal.ItemUID = "btnlabour") And (pVal.BeforeAction = False) Then
                        Dim StrSql As String
                        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Try
                            StrSql = "select Code as 'Labour ID',U_Empid as 'Employee ID',U_Empnam as 'Employee Name',U_LGCode as 'Skill Group',U_LGname as 'Skill Group Name',U_Currncy as 'Currency', " _
                                & "U_Labrate as 'Labour Rate/Hour',U_Accode as 'Account Code',U_Acname as 'Account Name',U_Adnl1 as 'Info1',U_Adnl2 as 'Info2',U_Remarks as 'Remarks' from [@PSSIT_OLBR]"
                            LBReporttClass = New LabourReport(SBO_Application, oCompany, StrSql)
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        End Try
                    End If
                    '*************Tools Button Press*****************
                    If (pVal.ItemUID = "btntools") And (pVal.BeforeAction = False) Then
                        Dim StrSql As String
                        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Try
                            StrSql = "select Code as 'Tool Code',U_TLname  as 'Tool Name',U_Itemcode  as 'Item Code',U_Itemname  as 'Item Name',U_WCcode  as 'Work Centre Code',U_WCname  as 'Work Centre Name', " _
                                    & "U_Purdate  as 'Purchase Date',U_Lcost  as 'Landed Cost',U_Enou  as 'Expected Strokes',U_Cnou  as 'Completed Strokes',U_Tstime  as 'Tool Setting Time',U_Cpno  as 'Cost/Stroke',U_Accode  as 'Account Code',U_Acname  as 'Account Name', " _
                                    & "U_Partool  as 'Parent Tool',U_Adnl1  as 'Info1',U_Adnl2  as 'Info2',U_Techspec  as 'Technical Specs',U_Remarks  as 'Remarks' from [@PSSIT_OTLS]"
                            TLReporttClass = New ToolsReport(SBO_Application, oCompany, StrSql)
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        End Try
                    End If
                    '*************Stoppage Button Press*****************
                    If (pVal.ItemUID = "btnstopage") And (pVal.BeforeAction = False) Then
                        Dim StrSql As String
                        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Try
                            StrSql = "SELECT [Code] as 'Stoppage Code',[U_Stopname] as 'Stoppage Name',[U_Catcode] as 'Category Code',[U_Catname] as 'Category  Name',[U_Plantime] as 'Planned Time (Mins)', " _
                                     & "[U_Adnl1] as 'Info1',[U_Adnl2] as 'Info2',[U_Remarks] as 'Remarks' FROM [@PSSIT_OSGE]"
                            STReporttClass = New StoppageReport(SBO_Application, oCompany, StrSql)
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        End Try
                    End If
                    '*************Operations Button Press*****************
                    If (pVal.ItemUID = "btnoprns") And (pVal.BeforeAction = False) Then
                        Dim StrSql As String
                        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Try
                            StrSql = "select Code,U_Oprname,U_Oprtype,U_Rework from [@PSSIT_OPRN] order by docentry"
                            OprReporttClass = New OperationReport(SBO_Application, oCompany, StrSql)
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        End Try
                    End If
                    '*************Operations Routing Button Press*****************
                    If (pVal.ItemUID = "btnoprrout") And (pVal.BeforeAction = False) Then
                        Dim StrSql As String
                        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Try
                            StrSql = "select Code,U_Itemcode,U_Itemname,U_Defrte from [@PSSIT_ORTE]"
                            OprRouteReporttClass = New OperationRoutingReport(SBO_Application, oCompany, StrSql)
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        End Try
                    End If

                End If
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try

    End Sub
End Class
