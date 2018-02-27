'''' <summary>
'''' Author                     Created Date
'''' Suresh                      22/01/2009
'''' <remarks> This class is used for entering the Labour Performance Child Report Details.</remarks>
Public Class LabrPerfChildReport
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
    Private ULbrId, UEmpId, UEmpName, UAvlHrs, UWrkHrs, UPerf As SAPbouiCOM.UserDataSource
    '**************************Items - Matrix************************************
    Private oLbrMatrix As SAPbouiCOM.Matrix
    Private oColumns As SAPbouiCOM.Columns
    Private oRowNoCol, oLbrIdCol, oEmpIdCol, oEmpNameCol, oAvlHrsCol, oWrkHrsCol, oPerfCol As SAPbouiCOM.Column
    Private oStrSql As String
    Private WithEvents LabourClass As Labour
#End Region
    ''' <summary>
    ''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmLbrPerfChidReport.srf") method is called to load the Labour Performance Child Report form.
    ''' Drawform() method is called to Initialize the form,Datasources and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company, ByVal aStrSql As String)
        MyBase.New(aSBO_Application, aCompany)
        oStrSql = aStrSql
        SBO_Application = aSBO_Application
        oCompany = aCompany
        SetApplication()
        LoadFromXML("FrmLbrPerfChidReport.srf")
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
            ULbrId = oForm.DataSources.UserDataSources.Add("ULbrId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
            UEmpId = oForm.DataSources.UserDataSources.Add("UEmpId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            UEmpName = oForm.DataSources.UserDataSources.Add("UEmpName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            UAvlHrs = oForm.DataSources.UserDataSources.Add("UAvlHrs", SAPbouiCOM.BoDataType.dt_PRICE, 10)
            UWrkHrs = oForm.DataSources.UserDataSources.Add("UWrkHrs", SAPbouiCOM.BoDataType.dt_PRICE, 10)
            UPerf = oForm.DataSources.UserDataSources.Add("UPerf", SAPbouiCOM.BoDataType.dt_PRICE, 10)

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

            oLbrMatrix = oForm.Items.Item("matlbr").Specific
            oLbrMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oColumns = oLbrMatrix.Columns

            oRowNoCol = oColumns.Item("#")
            oRowNoCol.Editable = False

            oLbrIdCol = oColumns.Item("lbrid")
            oLbrIdCol.DataBind.SetBound(True, "", " ULbrId")
            oLbrIdCol.Editable = False

            oEmpIdCol = oColumns.Item("empid")
            oEmpIdCol.DataBind.SetBound(True, "", "UEmpId")
            oEmpIdCol.Editable = False

            oEmpNameCol = oColumns.Item("empname")
            oEmpNameCol.DataBind.SetBound(True, "", "UEmpName")
            oEmpNameCol.Editable = False

            oAvlHrsCol = oColumns.Item("avlhrs")
            oAvlHrsCol.DataBind.SetBound(True, "", "UAvlHrs")
            oAvlHrsCol.Editable = False

            oWrkHrsCol = oColumns.Item("wrkhrs")
            oWrkHrsCol.DataBind.SetBound(True, "", "UWrkHrs")
            oWrkHrsCol.Editable = False

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
                    ULbrId.Value = oRS.Fields.Item(0).Value
                    UEmpId.Value = oRS.Fields.Item(1).Value
                    UEmpName.Value = oRS.Fields.Item(2).Value
                    UAvlHrs.Value = oRS.Fields.Item(3).Value
                    UWrkHrs.Value = oRS.Fields.Item(4).Value
                    UPerf.Value = oRS.Fields.Item(5).Value
                    oLbrMatrix.AddRow(1)
                    oRS.MoveNext()
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            If pVal.FormUID = "FLPCR" Then
                '*****************************Releasing the Com Object*******************************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                    SBO_Application = Nothing
                    GC.Collect()
                End If

                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED) Then
                    If (pVal.ItemUID = "matlbr") And pVal.BeforeAction = False Then
                        Dim oLbrId As String
                        Dim oLbrIdEdit As SAPbouiCOM.EditText
                        oLbrIdEdit = oLbrIdCol.Cells.Item(pVal.Row).Specific
                        oLbrId = oLbrIdEdit.Value
                        LabourClass = New Labour(SBO_Application, oCompany, oLbrId, "LbrPerfRpt")
                    End If
                End If

            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
