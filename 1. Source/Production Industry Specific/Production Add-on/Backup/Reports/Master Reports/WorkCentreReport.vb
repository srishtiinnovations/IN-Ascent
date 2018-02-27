'''' <summary>
'''' Author                     Created Date
'''' Suresh                      23/12/2008
'''' <remarks> This class is used for viewing Work Centre Reports.</remarks>
Public Class WorkCentreReport
    Inherits GeneralLib
    ''' <summary>
    ''' Variable Declaration
    ''' </summary>
    ''' <remarks></remarks>
#Region "Variable Declaration"
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private SboGuiApi As New SAPbouiCOM.SboGuiApi
    Private oCompany As SAPbobsCOM.Company
    '**************************UserDataSource************************************
    Private UWCCode, UWCName, UWCType, UInfo1, UInfo2, URemarks, UFxdCost, UCurrency, UUnitCost, UAbsMethod, UACCode, UACName, UOthrInfo As SAPbouiCOM.UserDataSource
    '**************************Form************************************
    Private oForm As SAPbouiCOM.Form
    '**************************Items - Matrix************************************
    Private oWCMatrix As SAPbouiCOM.Matrix
    Private oColumns As SAPbouiCOM.Columns
    Private oRowNoCol, oWCCodeCol, oWCNametCol, oWCTypeCol, oInfo1Col, oInfo2Col, oRemarksCol, oFxdCostCol, oCurrencyCol, oUnitCostCol, oAbsMethodCol, oACCodeCol, oACNameCol, oOthrInfoCol As SAPbouiCOM.Column
    Private oStrSql As String

#End Region
    ''' <summary>
    ''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmWorkCentreReport.srf") method is called to load the Work Centre form.
    ''' Drawform() method is called to Initialize the form,Datasources and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company, ByVal aStrSql As String)
        MyBase.New(aSBO_Application, aCompany)
        SBO_Application = aSBO_Application
        oCompany = aCompany
        oStrSql = aStrSql
        SetApplication()
        LoadFromXML("FrmWorkCentreReport.srf")
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
            UWCCode = oForm.DataSources.UserDataSources.Add("UWCCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
            UWCName = oForm.DataSources.UserDataSources.Add("UWCName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50)
            UWCType = oForm.DataSources.UserDataSources.Add("UWCType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50)
            UInfo1 = oForm.DataSources.UserDataSources.Add("UInfo1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
            UInfo2 = oForm.DataSources.UserDataSources.Add("UInfo2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
            URemarks = oForm.DataSources.UserDataSources.Add("URemarks", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            UFxdCost = oForm.DataSources.UserDataSources.Add("UFxdCost", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
            UCurrency = oForm.DataSources.UserDataSources.Add("UCurrency", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 25)
            UUnitCost = oForm.DataSources.UserDataSources.Add("UUnitCost", SAPbouiCOM.BoDataType.dt_PRICE, 19.6)
            UAbsMethod = oForm.DataSources.UserDataSources.Add("UAbsMethod", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 60)
            UACCode = oForm.DataSources.UserDataSources.Add("UACCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
            UACName = oForm.DataSources.UserDataSources.Add("UACName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            UOthrInfo = oForm.DataSources.UserDataSources.Add("UOthrInfo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50)
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
            oWCMatrix = oForm.Items.Item("matwc").Specific
            oWCMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oColumns = oWCMatrix.Columns

            oRowNoCol = oColumns.Item("#")
            oRowNoCol.Editable = False

            oWCCodeCol = oColumns.Item("colwccode")
            oWCCodeCol.DataBind.SetBound(True, "", "UWCCode")
            oWCCodeCol.Editable = False

            oWCNametCol = oColumns.Item("colwcname")
            oWCNametCol.DataBind.SetBound(True, "", "UWCName")
            oWCNametCol.Editable = False

            oWCTypeCol = oColumns.Item("colwctype")
            oWCTypeCol.DataBind.SetBound(True, "", " UWCType")
            oWCTypeCol.Editable = False


            oInfo1Col = oColumns.Item("coladnl1")
            oInfo1Col.DataBind.SetBound(True, "", "UInfo1")
            oInfo1Col.Editable = False

            oInfo2Col = oColumns.Item("coladnl2")
            oInfo2Col.DataBind.SetBound(True, "", "UInfo2")
            oInfo2Col.Editable = False

            oRemarksCol = oColumns.Item("colremarks")
            oRemarksCol.DataBind.SetBound(True, "", "URemarks")
            oRemarksCol.Editable = False

            oFxdCostCol = oColumns.Item("colfxdcost")
            oFxdCostCol.DataBind.SetBound(True, "", "UFxdCost")
            oFxdCostCol.Editable = False

            oCurrencyCol = oColumns.Item("colcurr")
            oCurrencyCol.DataBind.SetBound(True, "", "UCurrency")
            oCurrencyCol.Editable = False

            oUnitCostCol = oColumns.Item("coluntcost")
            oUnitCostCol.DataBind.SetBound(True, "", " UUnitCost")
            oUnitCostCol.Editable = False

            oAbsMethodCol = oColumns.Item("colabsmthd")
            oAbsMethodCol.DataBind.SetBound(True, "", "UAbsMethod")
            oAbsMethodCol.Editable = False
            oAbsMethodCol.Visible = False

            oACCodeCol = oColumns.Item("colaccode")
            oACCodeCol.DataBind.SetBound(True, "", "UACCode")
            oACCodeCol.Editable = False

            oACNameCol = oColumns.Item("colacname")
            oACNameCol.DataBind.SetBound(True, "", " UACName")
            oACNameCol.Editable = False

            oOthrInfoCol = oColumns.Item("colothinfo")
            oOthrInfoCol.DataBind.SetBound(True, "", " UOthrInfo")
            oOthrInfoCol.Editable = False


        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub LoadData()
        Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            oForm.DataSources.DataTables.Add("DTShift")
            oRS.DoQuery(oStrSql)
            If oRS.RecordCount > 0 Then
                oRS.MoveFirst()
                For i As Integer = 0 To oRS.RecordCount - 1
                    UWCCode.Value = oRS.Fields.Item(0).Value
                    UWCName.Value = oRS.Fields.Item(1).Value
                    UWCType.Value = oRS.Fields.Item(2).Value
                    UInfo1.Value = oRS.Fields.Item(3).Value
                    UInfo2.Value = oRS.Fields.Item(4).Value
                    URemarks.Value = oRS.Fields.Item(5).Value
                    UFxdCost.Value = oRS.Fields.Item(6).Value
                    UCurrency.Value = oRS.Fields.Item(7).Value
                    UUnitCost.Value = oRS.Fields.Item(8).Value
                    UAbsMethod.Value = oRS.Fields.Item(9).Value
                    UACCode.Value = oRS.Fields.Item(10).Value
                    UACName.Value = oRS.Fields.Item(11).Value
                    UOthrInfo.Value = oRS.Fields.Item(12).Value
                    oWCMatrix.AddRow(1)
                    oRS.MoveNext()
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        If pVal.FormUID = "FWCR" Then
            '*****************************Releasing the Com Object*******************************
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                SBO_Application = Nothing
                GC.Collect()
            End If
        End If
    End Sub
End Class
