'''' <summary>
'''' Author                     Created Date
'''' Suresh                      26/12/2008
'''' <remarks> This class is used for viewing Operations Machine Reports.</remarks>
Public Class OperationMachineReport
    Inherits GeneralLib
    ''' <summary>
    ''' Variable Declaration
    ''' </summary>
    ''' <remarks></remarks>
#Region "Variable Declaration"
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private SboGuiApi As SAPbouiCOM.SboGuiApi
    Private oCompany As SAPbobsCOM.Company
    '**************************UserDataSource************************************
    Private UMCCode, UMCName, UMCGroup As SAPbouiCOM.UserDataSource
    '**************************Form************************************
    Private oForm As SAPbouiCOM.Form
    '**************************Items - Matrix************************************
    Private oOprMCMatrix As SAPbouiCOM.Matrix
    Private oColumns As SAPbouiCOM.Columns
    Private oRowNoCol, oMCCodeCol, oMCNameCol, oMCGroupCol As SAPbouiCOM.Column
    Private oStrSql As String

#End Region
    ''' <summary>
    ''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmOprMachineReport.srf") method is called to load the Operations Machine form.
    ''' Drawform() method is called to Initialize the form,Datasources and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company, ByVal aStrSql As String)
        MyBase.New(aSBO_Application, aCompany)
        SBO_Application = aSBO_Application
        oCompany = aCompany
        oStrSql = aStrSql
        SetApplication()
        LoadFromXML("FrmOprMachineReport.srf")
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
            UMCCode = oForm.DataSources.UserDataSources.Add("UMCCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            UMCName = oForm.DataSources.UserDataSources.Add("UMCName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            UMCGroup = oForm.DataSources.UserDataSources.Add("UMCGroup", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Configuring the Matrix items/controls in the form(.srf) by bounding to the object and to the DBDatasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ConfigureMatrix()

        Try
            oOprMCMatrix = oForm.Items.Item("matmac").Specific
            oOprMCMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oColumns = oOprMCMatrix.Columns

            oRowNoCol = oColumns.Item("#")
            oRowNoCol.Editable = False

            oMCCodeCol = oColumns.Item("colmccode")
            oMCCodeCol.DataBind.SetBound(True, "", "UMCCode")
            oMCCodeCol.Editable = False

            oMCNameCol = oColumns.Item("colmcname")
            oMCNameCol.DataBind.SetBound(True, "", "UMCName")
            oMCNameCol.Editable = False

            oMCGroupCol = oColumns.Item("colmcgrp")
            oMCGroupCol.DataBind.SetBound(True, "", "UMCGroup")
            oMCGroupCol.Editable = False
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub LoadData()
        Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            oForm.DataSources.DataTables.Add("DTOprMC")
            oRS.DoQuery(oStrSql)
            If oRS.RecordCount > 0 Then
                oRS.MoveFirst()
                For i As Integer = 0 To oRS.RecordCount - 1
                    UMCCode.Value = oRS.Fields.Item(0).Value
                    UMCName.Value = oRS.Fields.Item(1).Value
                    UMCGroup.Value = oRS.Fields.Item(2).Value
                    oOprMCMatrix.AddRow(1)
                    oRS.MoveNext()
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        If pVal.FormUID = "FOPMR" Then
            '*****************************Releasing the Com Object*******************************
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                SBO_Application = Nothing
                GC.Collect()
            End If
        End If
    End Sub
End Class
