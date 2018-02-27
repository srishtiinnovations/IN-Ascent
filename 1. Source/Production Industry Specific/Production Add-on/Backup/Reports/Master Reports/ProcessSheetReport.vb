'''' <summary>
'''' Author                     Created Date
'''' Suresh                      06/01/2009
'''' <remarks> This class is used for viewing Process Sheet Reports List.</remarks>
Public Class ProcessSheetReport
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
    Private UProdNo As SAPbouiCOM.UserDataSource
    '**************************Items - EditText************************************
    Private oPordNoTxt As SAPbouiCOM.EditText
    '**************************Items - Button************************************
    Private BtnPrint As SAPbouiCOM.Button
    Private oPordNo As String
#End Region
    ''' <summary>
    ''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmProcessSheetReport.srf") method is called to load the Process Sheet Reports form.
    ''' Drawform() method is called to Initialize the form and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company, ByVal aPordNo As String)
        MyBase.New(aSBO_Application, aCompany)
        SBO_Application = aSBO_Application
        oCompany = aCompany
        oPordNo = aPordNo
        SetApplication()
        LoadFromXML("FrmProcessSheetReport.srf")
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
            InitializeFormComponent()
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
            UProdNo = oForm.DataSources.UserDataSources.Add("UProdNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)

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
            oPordNoTxt = oForm.Items.Item("txtprodno").Specific
            oPordNoTxt.DataBind.SetBound(True, "", "UProdNo")

            BtnPrint = oForm.Items.Item("btnprint").Specific
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub LoadData()
        Try
            UProdNo.Value = oPordNo
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent

    End Sub
End Class
