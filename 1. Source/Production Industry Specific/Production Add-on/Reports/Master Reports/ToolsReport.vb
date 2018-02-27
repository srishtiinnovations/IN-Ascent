'''' <summary>
'''' Author                     Created Date
'''' Suresh                      24/12/2008
'''' <remarks> This class is used for viewing Tools Reports.</remarks>
Public Class ToolsReport
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
    '**************************Items - Grid************************************
    Private oTLGrid As SAPbouiCOM.Grid
    '**************************UserDataSource************************************
    Private UserDS As SAPbouiCOM.UserDataSource
    '**************************Items - Option Button************************************
    Private oNOGroupingOptBtn, oWCOptBtn, oItmOptBtn As SAPbouiCOM.OptionBtn

    Private oStrSql As String
#End Region
    ''' <summary>
    ''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmToolsReport.srf") method is called to load the Work Centre form.
    ''' Drawform() method is called to Initialize the form,Datasources and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company, ByVal aStrSql As String)
        MyBase.New(aSBO_Application, aCompany)
        SBO_Application = aSBO_Application
        oCompany = aCompany
        oStrSql = aStrSql
        SetApplication()
        LoadFromXML("FrmToolsReport.srf")
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
    ''' Configuring the Matrix items/controls in the form(.srf) by bounding to the object and to the DBDatasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitializeFormComponent()

        Try

            UserDS = oForm.DataSources.UserDataSources.Add("OpBtnDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)

            oTLGrid = oForm.Items.Item("grdtools").Specific

            oForm.DataSources.DataTables.Add("MyDataTable")
            oForm.DataSources.DataTables.Item(0).ExecuteQuery(oStrSql)
            oTLGrid.DataTable = oForm.DataSources.DataTables.Item("MyDataTable")

            oNOGroupingOptBtn = oForm.Items.Item("optnogrp").Specific
            oNOGroupingOptBtn.DataBind.SetBound(True, , "OpBtnDS")

            oWCOptBtn = oForm.Items.Item("optwc").Specific
            oWCOptBtn.GroupWith("optnogrp")
            oWCOptBtn.DataBind.SetBound(True, , "OpBtnDS")

            oItmOptBtn = oForm.Items.Item("optitm").Specific
            oItmOptBtn.GroupWith("optwc")
            oItmOptBtn.DataBind.SetBound(True, , "OpBtnDS")

            GridEditable()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Set columns Editable False
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub GridEditable()

        oTLGrid.Columns.Item(0).Editable = False
        oTLGrid.Columns.Item(1).Editable = False
        oTLGrid.Columns.Item(2).Editable = False
        oTLGrid.Columns.Item(3).Editable = False
        oTLGrid.Columns.Item(4).Editable = False
        oTLGrid.Columns.Item(5).Editable = False
        oTLGrid.Columns.Item(6).Editable = False
        oTLGrid.Columns.Item(7).Editable = False
        oTLGrid.Columns.Item(8).Editable = False
        oTLGrid.Columns.Item(9).Editable = False
        oTLGrid.Columns.Item(10).Editable = False
        oTLGrid.Columns.Item(11).Editable = False
        oTLGrid.Columns.Item(12).Editable = False
        oTLGrid.Columns.Item(13).Editable = False
        oTLGrid.Columns.Item(14).Editable = False
        oTLGrid.Columns.Item(15).Editable = False
        oTLGrid.Columns.Item(16).Editable = False
    End Sub

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            If pVal.FormUID = "FTR" Then
                '*****************************Releasing the Com Object*******************************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                    SBO_Application = Nothing
                    GC.Collect()
                End If

                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.BeforeAction = False) Then
                    If pVal.ItemUID = "optnogrp" Or pVal.ItemUID = "optwc" Or pVal.ItemUID = "optitm" Then
                        If (UserDS.Value = 1) Then
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                            oForm.DataSources.DataTables.Item(0).ExecuteQuery(oStrSql)
                            GridEditable()
                            oTLGrid.CollapseLevel = 0
                            oForm.Freeze(False)
                        ElseIf (UserDS.Value = 2) Then
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                            oForm.DataSources.DataTables.Item(0).ExecuteQuery("select U_Itemcode  as 'Item Code',U_Itemname  as 'Item Name',Code as 'Tool Code',U_TLname  as 'Tool Name',U_WCcode  as 'Work Centre Code',U_WCname  as 'Work Centre Name', " _
                                    & "U_Purdate  as 'Purchase Date',U_Lcost  as 'Landed Cost',U_Enou  as 'Expected Strokes',U_Cnou  as 'Completed Strokes',U_Tstime  as 'Tool Setting Time',U_Cpno  as 'Cost/Stroke',U_Accode  as 'Account Code',U_Acname  as 'Account Name', " _
                                    & "U_Partool  as 'Parent Tool',U_Adnl1  as 'Info1',U_Adnl2  as 'Info2',U_Techspec  as 'Technical Specs',U_Remarks  as 'Remarks' from [@PSSIT_OTLS]")
                            GridEditable()
                            oTLGrid.CollapseLevel = 1
                            oForm.Freeze(False)
                        ElseIf (UserDS.Value = 3) Then
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                            oForm.DataSources.DataTables.Item(0).ExecuteQuery("select U_WCcode  as 'Work Centre Code',U_WCname  as 'Work Centre Name',Code as 'Tool Code',U_TLname  as 'Tool Name',U_Itemcode  as 'Item Code',U_Itemname  as 'Item Name', " _
                                    & "U_Purdate  as 'Purchase Date',U_Lcost  as 'Landed Cost',U_Enou  as 'Expected Strokes',U_Cnou  as 'Completed Strokes',U_Tstime  as 'Tool Setting Time',U_Cpno  as 'Cost/Stroke',U_Accode  as 'Account Code',U_Acname  as 'Account Name', " _
                                    & "U_Partool  as 'Parent Tool',U_Adnl1  as 'Info1',U_Adnl2  as 'Info2',U_Techspec  as 'Technical Specs',U_Remarks  as 'Remarks' from [@PSSIT_OTLS]")
                            GridEditable()
                            oTLGrid.CollapseLevel = 1
                            oForm.Freeze(False)
                        End If
                    End If

                    If ((pVal.ItemUID = "btncolapse") Or (pVal.ItemUID = "btnexpand")) Then
                        If (pVal.ItemUID = "btncolapse") Then
                            oTLGrid.Rows.CollapseAll()
                            GridEditable()
                        End If
                        If (pVal.ItemUID = "btnexpand") Then
                            oTLGrid.Rows.ExpandAll()
                            GridEditable()
                        End If
                    End If

                End If

            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
