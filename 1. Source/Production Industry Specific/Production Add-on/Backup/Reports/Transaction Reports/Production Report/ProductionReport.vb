'''' <summary>
'''' Author                     Created Date
'''' Suresh                      15/01/2009
'''' <remarks> This class is used for entering the Parameters for the Production Report.</remarks>
Public Class ProductionReport
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
    Private UFPordDate, UTPordDate As SAPbouiCOM.UserDataSource
    '**************************Items - EditText************************************
    Private oFromPordDateTxt, oToPordDateTxt As SAPbouiCOM.EditText
    '**************************Items - Button************************************
    Private BtnPrint As SAPbouiCOM.Button
    Private sPath As String = IO.Directory.GetParent(Application.ExecutablePath).ToString
    Private TestForm As FrmTest
    Private oThread As System.Threading.Thread
#End Region
    ''' <summary>
    ''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmProductionReport.srf") method is called to load the Production Report form.
    ''' Drawform() method is called to Initialize the form,Datasources and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company)
        MyBase.New(aSBO_Application, aCompany)
        SBO_Application = aSBO_Application
        oCompany = aCompany
        LoadFromXML("FrmProductionReport.srf")
        DrawForm()
    End Sub
    ''' <summary>
    ''' Connecting the application through connection string.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetApplication()
        'Dim sConnectionString As String
        'SboGuiApi = New SAPbouiCOM.SboGuiApi
        'sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"
        'SboGuiApi.Connect(sConnectionString)
        'SboGuiApi.AddonIdentifier = "5645523035446576656C6F706D656E743A453038373933323333343581F0D8D8C45495472FC628EF425AD5AC2AEDC411"
        'SBO_Application = SboGuiApi.GetApplication()
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
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
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
            UFPordDate = oForm.DataSources.UserDataSources.Add("UFPordDate", SAPbouiCOM.BoDataType.dt_DATE, 10)
            UTPordDate = oForm.DataSources.UserDataSources.Add("UTPordDate", SAPbouiCOM.BoDataType.dt_DATE, 10)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Configuring the items/controls in the form(.srf) by bounding to the object and to the DBDatasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitializeFormComponent()
        Try
            oFromPordDateTxt = oForm.Items.Item("txtfdate").Specific
            oFromPordDateTxt.DataBind.SetBound(True, "", "UFPordDate")

            oToPordDateTxt = oForm.Items.Item("txttodate").Specific
            oToPordDateTxt.DataBind.SetBound(True, "", "UTPordDate")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            If pVal.FormUID = "FPR" Then
                '*****************************Releasing the Com Object*******************************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                    SBO_Application = Nothing
                    GC.Collect()
                End If
                '**************Item Pressed**************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.ItemUID = "btnprint" Then
                        '***** Validation() method is called for validating the values in the edit text *****
                        If (pVal.BeforeAction = True) Then
                            Try
                                Validation()
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                            End Try
                        End If
                        If (pVal.BeforeAction = False) Then
                            'Modified by Manimaran---------s
                            TestForm = New FrmTest(SBO_Application, oCompany, "ProdReport", oFromPordDateTxt.Value, oToPordDateTxt.Value)
                            'TestForm = New FrmTest(SBO_Application, oCompany, "ProdReport", oFromPordDateTxt.String, oToPordDateTxt.String)
                            'Modified by Manimaran---------e
                            oThread = New Threading.Thread(AddressOf TestForm.StartThread)
                            oThread.Start()
                        End If
                    End If
                End If

            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    Function String2Date(ByVal S As String, _
                            ByVal Fmt As String) As Object
        Select Case Fmt
            Case "MMDDYY", "MMDDYYYY"      '052793   05271993
                String2Date = CDate(Left(S, 2) & "/" & Mid(S, 3, 2) & "/" & _
                                    Mid(S, 5))
            Case "DDMMYY", "DDMMYYYY"      '270593   27051993
                String2Date = CDate(Mid(S, 3, 2) & "/" & Left(S, 2) & "/" & _
                                    Mid(S, 5))
            Case "YYMMDD"                  '930527
                String2Date = CDate(Mid(S, 3, 2) & "/" & Right(S, 2) & "/" & _
                                    Left(S, 2))
            Case "YYYYMMDD"                '19930527
                String2Date = CDate(Mid(S, 5, 2) & "/" & Right(S, 2) & "/" & _
                                    Left(S, 4))
            Case "MM/DD/YY", "MM/DD/YYYY", "M/D/Y", "M/D/YY", "M/D/YYYY", _
                 "DD-MMM-YY", "DD-MMM-YYYY"
                String2Date = CDate(S)
            Case "DD/MM/YY", "DD/MM/YYYY"  '27/05/93   27/05/1993
                String2Date = CDate(Mid(S, 4, 3) & Left(S, 3) & Mid(S, 7))
            Case "YY/MM/DD"                '93/05/27
                String2Date = CDate(Mid(S, 4, 3) & Right(S, 2) & _
                                    "/" & Left(S, 2))
            Case "YYYY/MM/DD"              '1993/05/27
                String2Date = CDate(Mid(S, 6, 3) & Right(S, 2) & _
                                    "/" & Left(S, 4))
            Case Else
                String2Date = Nothing
        End Select
    End Function
    ''' <summary>
    ''' This method is used for validating the values in the EditText.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Validation()

        Try
            If oFromPordDateTxt.Value.Length = 0 Then
                oFromPordDateTxt.Active = True
                Throw New Exception("From Date should not be Empty")
            End If
            If oToPordDateTxt.Value.Length = 0 Then
                oToPordDateTxt.Active = True
                Throw New Exception("To Date should not be Empty")
            End If
           
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
