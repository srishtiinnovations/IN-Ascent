''' <summary>
''' Author                     Created Date
''' Suresh                      03/12/2008
''' </summary>
''' <remarks> This class is a Common library for all the class which is inheritable.</remarks>
Public Class ConnectionLib
    ''' <summary>
    ''' Variable Declaration
    ''' </summary>
    ''' <remarks></remarks>
#Region "Variable Declaration"
    Public WithEvents SBO_Application As SAPbouiCOM.Application
    Private oForm As SAPbouiCOM.Form
    Public oCompany As New SAPbobsCOM.Company
    Private SboGuiApi As SAPbouiCOM.SboGuiApi
    Private sConnectionString As String
#End Region
    ''' <summary>
    ''' SetApplication() and Initialize() methods are called.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        MyBase.New()
        SetApplication()
        Initialize()
    End Sub
    ''' <summary>
    ''' Connecting the application through connection string.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetApplication()
        SboGuiApi = New SAPbouiCOM.SboGuiApi
        sConnectionString = Environment.GetCommandLineArgs.GetValue(1) '"0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"
        SboGuiApi.Connect(sConnectionString)
        SboGuiApi.AddonIdentifier = "5645523035446576656C6F706D656E743A453038373933323333343581F0D8D8C45495472FC628EF425AD5AC2AEDC411"
        SBO_Application = SboGuiApi.GetApplication()
    End Sub
    ''' <summary>
    ''' Connecting  to the Company database.
    ''' </summary>
    ''' <remarks></remarks>
#Region "Connect to Company"
    Private Sub Initialize()
        If Not SetConnectionContext() = 0 Then
            SBO_Application.MessageBox("Failed setting a connection to DI API")
            End
        End If
        If Not ConnectToCompany() = 0 Then
            SBO_Application.MessageBox("Failed connecting to the oCompany's Data Base")
            End
        End If
    End Sub
    ''' <summary>
    ''' Getting the Connection Context.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SetConnectionContext() As Integer
        Dim sCookie As String
        Dim sConnectionContext As String
        ' Dim lRetCode As Integer
        Try
            oCompany = New SAPbobsCOM.Company
            sCookie = oCompany.GetContextCookie
            sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie)
            If oCompany.Connected = True Then
                oCompany.Disconnect()
            End If
            SetConnectionContext = oCompany.SetSboLoginContext(sConnectionContext)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
    ''' <summary>
    ''' Connecting to the Company.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ConnectToCompany() As Integer
        Try
            ConnectToCompany = oCompany.Connect
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
#End Region
    ''' <summary>
    ''' Destructing the memory object.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="EventType"></param>
    ''' <remarks></remarks>
    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBO_Application.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown
                SBO_Application.SetStatusBarMessage("A Shut Down Event has been caught" & _
                Environment.NewLine() & "Terminating Add On...", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                System.Windows.Forms.Application.Exit()
            Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                SBO_Application.SetStatusBarMessage("A Company Change Event has been caught" & _
                Environment.NewLine() & "Terminating Add On...", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                System.Windows.Forms.Application.Exit()
            Case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
                SBO_Application.SetStatusBarMessage("A Server Terminition Event has been caught" & _
                Environment.NewLine() & "Terminating Add On...", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                System.Windows.Forms.Application.Exit()
        End Select
    End Sub
End Class
