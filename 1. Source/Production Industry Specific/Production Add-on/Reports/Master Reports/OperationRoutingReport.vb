'''' <summary>
'''' Author                     Created Date
'''' Suresh                      26/12/2008
'''' <remarks> This class is used for entering the Operation Routing Report Details.</remarks>
Public Class OperationRoutingReport
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
    Private URouteCode, UItemCode, UItemName, UDefltRout As SAPbouiCOM.UserDataSource
    '**************************Items - Matrix************************************
    Private oRoutMatrix As SAPbouiCOM.Matrix
    Private oColumns As SAPbouiCOM.Columns
    Private oRowNoCol, oRoutCodeCol, oItemCodeCol, oItemNameCol, oDefltRoutCol As SAPbouiCOM.Column
    Private oStrSql As String
    Private WithEvents OprRouteClass As OperationsRouting
#End Region
    ''' <summary>
    ''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmOperationReport.srf") method is called to load the Operation Report form.
    ''' Drawform() method is called to Initialize the form,Datasources and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company, ByVal aStrSql As String)
        MyBase.New(aSBO_Application, aCompany)
        oStrSql = aStrSql
        SBO_Application = aSBO_Application
        oCompany = aCompany
        SetApplication()
        LoadFromXML("FrmRouteCardReport.srf")
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
            URouteCode = oForm.DataSources.UserDataSources.Add("URouteCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            UItemCode = oForm.DataSources.UserDataSources.Add("UItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            UItemName = oForm.DataSources.UserDataSources.Add("UItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            UDefltRout = oForm.DataSources.UserDataSources.Add("UDefltRout", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
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

            oRoutMatrix = oForm.Items.Item("matoprmc").Specific
            oRoutMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oColumns = oRoutMatrix.Columns

            oRowNoCol = oColumns.Item("#")
            oRowNoCol.Editable = False

            oRoutCodeCol = oColumns.Item("colroute")
            oRoutCodeCol.DataBind.SetBound(True, "", "URouteCode")
            oRoutCodeCol.Editable = False

            oItemCodeCol = oColumns.Item("colitmcode")
            oItemCodeCol.DataBind.SetBound(True, "", "UItemCode")
            oItemCodeCol.Editable = False

            oItemNameCol = oColumns.Item("colitmdesc")
            oItemNameCol.DataBind.SetBound(True, "", "UItemName")
            oItemNameCol.Editable = False

            oDefltRoutCol = oColumns.Item("coldefltrt")
            oDefltRoutCol.DataBind.SetBound(True, "", "UDefltRout")
            oDefltRoutCol.Editable = False

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub LoadData()
        Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            oForm.DataSources.DataTables.Add("DTRout")
            oRS.DoQuery(oStrSql)
            If oRS.RecordCount > 0 Then
                oRS.MoveFirst()
                For i As Integer = 0 To oRS.RecordCount - 1
                    URouteCode.Value = oRS.Fields.Item(0).Value
                    UItemCode.Value = oRS.Fields.Item(1).Value
                    UItemName.Value = oRS.Fields.Item(2).Value
                    UDefltRout.Value = oRS.Fields.Item(3).Value

                    oRoutMatrix.AddRow(1)
                    oRS.MoveNext()
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            If pVal.FormUID = "FRCR" Then
                '*****************************Releasing the Com Object*******************************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                    SBO_Application = Nothing
                    GC.Collect()
                End If

                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED) Then
                    If (pVal.ItemUID = "matoprmc") And pVal.BeforeAction = False Then
                        Dim oRouteCode As String
                        Dim oRouteCodeEdit As SAPbouiCOM.EditText
                        oRouteCodeEdit = oRoutCodeCol.Cells.Item(pVal.Row).Specific
                        oRouteCode = oRouteCodeEdit.Value
                        OprRouteClass = New OperationsRouting(SBO_Application, oCompany, oRouteCode, "OprRouting")
                    End If
                End If
               
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
