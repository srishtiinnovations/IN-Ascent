'''' <summary>
'''' Author                     Created Date
'''' Suresh                      26/12/2008
'''' <remarks> This class is used for entering the Operation Details.</remarks>
Public Class OperationReport
    Inherits GeneralLib
    ''' <summary>
    ''' Variable Declaration
    ''' </summary>
    ''' <remarks></remarks>
#Region "Variable Declaration"
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private SboGuiApi As SAPbouiCOM.SboGuiApi
    Private oCompany As SAPbobsCOM.Company
    '**************************Form************************************
    Private oForm As SAPbouiCOM.Form
    '**************************UserDataSource************************************
    Private UOprCode, UOprName, UOprType, URework As SAPbouiCOM.UserDataSource
    '**************************Items - Matrix************************************
    Private oOprMatrix As SAPbouiCOM.Matrix
    Private oColumns As SAPbouiCOM.Columns
    Private oRowNoCol, oOprCodeCol, oOprNameCol, oOprTypeCol, oReworkCol As SAPbouiCOM.Column
    Private oStrSql As String
    Private WithEvents OprMCReporttClass As OperationMachineReport
#End Region
    ''' <summary>
    ''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmOperationReport.srf") method is called to load the Operation Report form.
    ''' Drawform() method is called to Initialize the form,Datasources and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company, ByVal aStrSql As String)
        MyBase.New(aSBO_Application, aCompany)
        SBO_Application = aSBO_Application
        oCompany = aCompany
        oStrSql = aStrSql
        SetApplication()
        LoadFromXML("FrmOperationReport.srf")
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
            UOprCode = oForm.DataSources.UserDataSources.Add("UOprCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            UOprName = oForm.DataSources.UserDataSources.Add("UOprName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            UOprType = oForm.DataSources.UserDataSources.Add("UOprType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            URework = oForm.DataSources.UserDataSources.Add("URework", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
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

            oOprMatrix = oForm.Items.Item("matopr").Specific
            oOprMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oColumns = oOprMatrix.Columns

            oRowNoCol = oColumns.Item("#")
            oRowNoCol.Editable = False

            oOprCodeCol = oColumns.Item("coloprid")
            oOprCodeCol.DataBind.SetBound(True, "", "UOprCode")
            oOprCodeCol.Editable = False

            oOprNameCol = oColumns.Item("coloprnam")
            oOprNameCol.DataBind.SetBound(True, "", "UOprName")
            oOprNameCol.Editable = False

            oOprTypeCol = oColumns.Item("coloprtyp")
            oOprTypeCol.DataBind.SetBound(True, "", "UOprType")
            oOprTypeCol.Editable = False

            oReworkCol = oColumns.Item("colrewrk")
            oReworkCol.DataBind.SetBound(True, "", "URework")
            oReworkCol.Editable = False

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub LoadData()
        Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            oForm.DataSources.DataTables.Add("DTOpr")
            oRS.DoQuery(oStrSql)
            If oRS.RecordCount > 0 Then
                oRS.MoveFirst()
                For i As Integer = 0 To oRS.RecordCount - 1
                    UOprCode.Value = oRS.Fields.Item(0).Value
                    UOprName.Value = oRS.Fields.Item(1).Value
                    UOprType.Value = oRS.Fields.Item(2).Value
                    URework.Value = oRS.Fields.Item(3).Value
                    
                    oOprMatrix.AddRow(1)
                    oRS.MoveNext()
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            If pVal.FormUID = "FOPR" Then
                '*****************************Releasing the Com Object*******************************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                    SBO_Application = Nothing
                    GC.Collect()
                End If

                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK) Then
                    If (pVal.ItemUID = "matopr") And pVal.BeforeAction = False Then
                        Dim oOprID, oOprName As SAPbouiCOM.EditText
                        Dim oCurrentRow As Integer
                        oCurrentRow = pVal.Row

                        Dim oRs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim StrSql As String


                        '******** Adding Empty Row to the  Items Bom Matrix *********
                        oOprMatrix.SelectRow(pVal.Row, True, False)
                        If oOprMatrix.IsRowSelected(pVal.Row) = True Then
                            oOprMatrix.GetLineData(oOprMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                            oOprID = oOprCodeCol.Cells.Item(oCurrentRow).Specific
                            oOprName = oOprNameCol.Cells.Item(oCurrentRow).Specific

                            StrSql = "select U_Wcno,U_wcname,U_MGname from [@PSSIT_PRN1] where Code='" & oOprID.Value & "'"

                            OprMCReporttClass = New OperationMachineReport(SBO_Application, oCompany, StrSql)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
