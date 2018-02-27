'''' <summary>
'''' Author                     Created Date
'''' Suresh                      21/01/2009
'''' <remarks> This class is used for entering the Machine Performance Child Report Details.</remarks>
Public Class MCPerfChildReport
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
    '**************************UserDataSource************************************
    Private UMCNo, UComp, UPlanHrs, UPlanQty, UWrkHrs, UOutQty, UPerf As SAPbouiCOM.UserDataSource
    '**************************Items - Matrix************************************
    Private oMCMatrix As SAPbouiCOM.Matrix
    Private oColumns As SAPbouiCOM.Columns
    Private oRowNoCol, oMCNoCol, oCompCol, oPlanHrsCol, oPlanQtyCol, oWrkHrsCol, oOutQtyCol, oPerfCol As SAPbouiCOM.Column
    Private oStrSql As String
    Private WithEvents MachineMasterClass As MachineMaster
#End Region
    ''' <summary>
    ''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmMCPerfChidReport.srf") method is called to load the Machine Performance Child Report form.
    ''' Drawform() method is called to Initialize the form,Datasources and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company, ByVal aStrSql As String)
        MyBase.New(aSBO_Application, aCompany)
        oStrSql = aStrSql
        SBO_Application = aSBO_Application
        oCompany = aCompany
        SetApplication()
        LoadFromXML("FrmMCPerfChidReport.srf")
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
            ConfigureMatrix()
            LoadData()
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
            UMCNo = oForm.DataSources.UserDataSources.Add("UMCNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
            UComp = oForm.DataSources.UserDataSources.Add("UComp", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            UPlanHrs = oForm.DataSources.UserDataSources.Add("UPlanHrs", SAPbouiCOM.BoDataType.dt_PRICE, 10)
            UPlanQty = oForm.DataSources.UserDataSources.Add("UPlanQty", SAPbouiCOM.BoDataType.dt_QUANTITY, 5)
            UWrkHrs = oForm.DataSources.UserDataSources.Add("UWrkHrs", SAPbouiCOM.BoDataType.dt_PRICE, 10)
            UOutQty = oForm.DataSources.UserDataSources.Add("UOutQty", SAPbouiCOM.BoDataType.dt_QUANTITY, 5)
            UPerf = oForm.DataSources.UserDataSources.Add("UPerf", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 5)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Configuring the Matrix items/controls in the form(.srf) by bounding to the object and to the DBDatasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ConfigureMatrix()
        Dim Fsetting As SAPbouiCOM.FormSettings
        Try
            Fsetting = oForm.Settings
            Fsetting.EnableRowFormat = False

            oMCMatrix = oForm.Items.Item("matmac").Specific
            oMCMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oColumns = oMCMatrix.Columns

            oRowNoCol = oColumns.Item("#")
            oRowNoCol.Editable = False

            oMCNoCol = oColumns.Item("machine")
            oMCNoCol.DataBind.SetBound(True, "", " UMCNo")
            oMCNoCol.Editable = False

            oCompCol = oColumns.Item("comp")
            oCompCol.DataBind.SetBound(True, "", "UComp")
            oCompCol.Editable = False

            oPlanHrsCol = oColumns.Item("planhrs")
            oPlanHrsCol.DataBind.SetBound(True, "", "UPlanHrs")
            oPlanHrsCol.Editable = False

            oPlanQtyCol = oColumns.Item("planqty")
            oPlanQtyCol.DataBind.SetBound(True, "", "UPlanQty")
            oPlanQtyCol.Editable = False

            oWrkHrsCol = oColumns.Item("wrkhrs")
            oWrkHrsCol.DataBind.SetBound(True, "", "UWrkHrs")
            oWrkHrsCol.Editable = False

            oOutQtyCol = oColumns.Item("outqty")
            oOutQtyCol.DataBind.SetBound(True, "", "UOutQty")
            oOutQtyCol.Editable = False

            oPerfCol = oColumns.Item("perfmns")
            oPerfCol.DataBind.SetBound(True, "", "UPerf")
            oPerfCol.Editable = False

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub LoadData()
        Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            oRS.DoQuery(oStrSql)
            If oRS.RecordCount > 0 Then
                oRS.MoveFirst()
                For i As Integer = 0 To oRS.RecordCount - 1
                    UMCNo.Value = oRS.Fields.Item(0).Value
                    UComp.Value = oRS.Fields.Item(1).Value
                    UPlanHrs.Value = oRS.Fields.Item(2).Value
                    UPlanQty.Value = oRS.Fields.Item(3).Value
                    UWrkHrs.Value = oRS.Fields.Item(4).Value
                    UOutQty.Value = oRS.Fields.Item(5).Value
                    UPerf.Value = oRS.Fields.Item(6).Value
                    oMCMatrix.AddRow(1)
                    oRS.MoveNext()
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            If pVal.FormUID = "FMPCR" Then
                '*****************************Releasing the Com Object*******************************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                    SBO_Application = Nothing
                    GC.Collect()
                End If

                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED) Then
                    If (pVal.ItemUID = "matmac") And pVal.BeforeAction = False Then
                        Dim oMCNo As String
                        Dim oMCNoEdit As SAPbouiCOM.EditText
                        oMCNoEdit = oMCNoCol.Cells.Item(pVal.Row).Specific
                        oMCNo = oMCNoEdit.Value
                        MachineMasterClass = New MachineMaster(SBO_Application, oCompany, oMCNo, "MCPerfRpt")
                    End If
                End If

            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
