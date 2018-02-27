'''' <summary>
'''' Author                     Created Date
'''' Suresh                      23/12/2008
'''' <remarks> This class is used for viewing Shift Master Reports.</remarks>
Public Class ShiftReport
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
    Private UShiftCode, UShiftName, UFromTime, UToTime, UBreak, UDurMins, UDurHrs, UInfo1, UInfo2, URemarks As SAPbouiCOM.UserDataSource
    '**************************Form************************************
    Private oForm As SAPbouiCOM.Form
    '**************************Items - Matrix************************************
    Private oShiftMatrix As SAPbouiCOM.Matrix
    Private oColumns As SAPbouiCOM.Columns
    Private oRowNoCol, oShiftCodeCol, oShiftNametCol, oFromTimeCol, oToTimeCol, oBreakCol, oDurMinsCol, oDurhrsCol, oInfo1Col, oInfo2Col, oRemarksCol As SAPbouiCOM.Column
    Private oStrSql As String

#End Region
    ''' <summary>
    ''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmShiftReport.srf") method is called to load the Work Centre form.
    ''' Drawform() method is called to Initialize the form,Datasources and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company, ByVal aStrSql As String)
        MyBase.New(aSBO_Application, aCompany)
        SBO_Application = aSBO_Application
        oCompany = aCompany
        oStrSql = aStrSql
        SetApplication()
        LoadFromXML("FrmShiftReport.srf")
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
            UShiftCode = oForm.DataSources.UserDataSources.Add("UShiftCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            UShiftName = oForm.DataSources.UserDataSources.Add("UShiftName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            UFromTime = oForm.DataSources.UserDataSources.Add("UFromTime", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            UToTime = oForm.DataSources.UserDataSources.Add("UToTime", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            UBreak = oForm.DataSources.UserDataSources.Add("UBreak", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            UDurMins = oForm.DataSources.UserDataSources.Add("UDurMins", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            UDurHrs = oForm.DataSources.UserDataSources.Add("UDurHrs", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            UInfo1 = oForm.DataSources.UserDataSources.Add("UInfo1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 40)
            UInfo2 = oForm.DataSources.UserDataSources.Add("UInfo2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 40)
            URemarks = oForm.DataSources.UserDataSources.Add("URemarks", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)

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
            oShiftMatrix = oForm.Items.Item("matshift").Specific
            oShiftMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oColumns = oShiftMatrix.Columns

            oRowNoCol = oColumns.Item("#")
            oRowNoCol.Editable = False

            oShiftCodeCol = oColumns.Item("colshift")
            oShiftCodeCol.DataBind.SetBound(True, "", "UShiftCode")
            oShiftCodeCol.Editable = False

            oShiftNametCol = oColumns.Item("colsftname")
            oShiftNametCol.DataBind.SetBound(True, "", "UShiftName")
            oShiftNametCol.Editable = False

            oFromTimeCol = oColumns.Item("colftime")
            oFromTimeCol.DataBind.SetBound(True, "", "UFromTime")
            oFromTimeCol.Editable = False

            oToTimeCol = oColumns.Item("coltotime")
            oToTimeCol.DataBind.SetBound(True, "", "UToTime")
            oToTimeCol.Editable = False
            
            oBreakCol = oColumns.Item("colbreak")
            oBreakCol.DataBind.SetBound(True, "", "UBreak")
            oBreakCol.Editable = False

            oDurMinsCol = oColumns.Item("coldurmins")
            oDurMinsCol.DataBind.SetBound(True, "", " UDurMins")
            oDurMinsCol.Editable = False

            oDurhrsCol = oColumns.Item("coldurhrs")
            oDurhrsCol.DataBind.SetBound(True, "", "UDurHrs")
            oDurhrsCol.Editable = False

            oInfo1Col = oColumns.Item("coladnl1")
            oInfo1Col.DataBind.SetBound(True, "", "UInfo1")
            oInfo1Col.Editable = False

            oInfo2Col = oColumns.Item("coladnl2")
            oInfo2Col.DataBind.SetBound(True, "", "UInfo2")
            oInfo2Col.Editable = False

            oRemarksCol = oColumns.Item("colremarks")
            oRemarksCol.DataBind.SetBound(True, "", "URemarks")
            oRemarksCol.Editable = False
         

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
                    UShiftCode.Value = oRS.Fields.Item(0).Value
                    UShiftName.Value = oRS.Fields.Item(1).Value
                    UFromTime.Value = oRS.Fields.Item(2).Value
                    UToTime.Value = oRS.Fields.Item(3).Value
                    UBreak.Value = oRS.Fields.Item(4).Value
                    UDurMins.Value = oRS.Fields.Item(5).Value
                    UDurHrs.Value = oRS.Fields.Item(6).Value
                    UInfo1.Value = oRS.Fields.Item(7).Value
                    UInfo2.Value = oRS.Fields.Item(8).Value
                    URemarks.Value = oRS.Fields.Item(9).Value
                    oShiftMatrix.AddRow(1)
                    oRS.MoveNext()
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        If pVal.FormUID = "FMSR" Then
            '*****************************Releasing the Com Object*******************************
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                SBO_Application = Nothing
                GC.Collect()
            End If
        End If
    End Sub
End Class
