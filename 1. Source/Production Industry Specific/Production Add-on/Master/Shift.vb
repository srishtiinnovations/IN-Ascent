'''' <summary>
'''' Author                     Created Date
'''' Suresh                      15/12/2008
'''' <remarks> This class is used for entering the shift details.</remarks>

Public Class Shift
    Inherits GeneralLib
    ''' <summary>
    ''' Variable Declaration
    ''' </summary>
    ''' <remarks></remarks>
#Region "Variable Declaration"
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private SboGuiApi As New SAPbouiCOM.SboGuiApi
    Private oCompany As SAPbobsCOM.Company
    '**************************DataSource************************************
    Private oParentDB As SAPbouiCOM.DBDataSource
    '**************************Form************************************
    Private oForm As SAPbouiCOM.Form
    '**************************Items - EditText************************************
    Private oCodeTxt, oNameTxt, oSCodeTxt, oDescTxt, oFromTimeTxt, oToTimeTxt, oDurationTxt, oBreakTimeTxt, oDurationMinTxt, oDurationHrsTxt, oInfo1Txt, oInfo2Txt, oRemarksTxt As SAPbouiCOM.EditText
    '**************************Items - CheckBox************************************
    Private oActiveCheck As SAPbouiCOM.CheckBox
    Private oShiftCode, oFormName As String
#End Region
    ''' <summary>
    ''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmShift.srf") method is called to load the shift form.
    ''' Drawform() method is called to Initialize the form,Datasources and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company, Optional ByVal aShiftCode As String = Nothing, Optional ByVal aFormName As String = Nothing)
        MyBase.New(aSBO_Application, aCompany)
        SBO_Application = aSBO_Application
        oCompany = aCompany
        oShiftCode = aShiftCode
        oFormName = aFormName
        LoadFromXML("FrmShift.srf")
        DrawForm()
        If oFormName = "Production Entry" Or oFormName = "Machine" Or oFormName = "DownTime" Then
            oParentDB.SetValue("Code", oParentDB.Offset, oShiftCode)
            oSCodeTxt.Value = oShiftCode
            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        End If
        oForm.DataBrowser.BrowseBy = "txtscode"
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
            oParentDB = oForm.DataSources.DBDataSources.Item("@PSSIT_OSFT")
            oForm.Freeze(True)
            InitializeFormComponent()
            oForm.Freeze(False)
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
            oSCodeTxt = oForm.Items.Item("txtscode").Specific
            oSCodeTxt.DataBind.SetBound(True, "@PSSIT_OSFT", "Code")

            oDescTxt = oForm.Items.Item("txtsdescr").Specific
            oDescTxt.DataBind.SetBound(True, "@PSSIT_OSFT", "U_Sdescr")

            oFromTimeTxt = oForm.Items.Item("txtsftime").Specific
            oFromTimeTxt.DataBind.SetBound(True, "@PSSIT_OSFT", "U_Sftime")
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                oFromTimeTxt.String = FormatDateTime(Now(), DateFormat.ShortTime)
            End If

            oToTimeTxt = oForm.Items.Item("txtsttime").Specific
            oToTimeTxt.DataBind.SetBound(True, "@PSSIT_OSFT", "U_Sttime")
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                oToTimeTxt.String = FormatDateTime(Now(), DateFormat.ShortTime)
            End If

            oBreakTimeTxt = oForm.Items.Item("txtsbreak").Specific
            oBreakTimeTxt.DataBind.SetBound(True, "@PSSIT_OSFT", "U_Sbreak")
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                oBreakTimeTxt.Value = 0
            End If

            oDurationMinTxt = oForm.Items.Item("txtsdurmin").Specific
            oDurationMinTxt.DataBind.SetBound(True, "@PSSIT_OSFT", "U_Duratmin")
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                oDurationMinTxt.String = DurationMinsCalculation()
            End If

            oDurationHrsTxt = oForm.Items.Item("txtsdurhrs").Specific
            oDurationHrsTxt.DataBind.SetBound(True, "@PSSIT_OSFT", "U_Durathrs")
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                oDurationHrsTxt.String = DurationHrsCalculation()
            End If


            oInfo1Txt = oForm.Items.Item("txtinfo1").Specific
            oForm.Items.Item("lbloi1").Visible = False
            oForm.Items.Item("txtinfo1").Visible = False
            oInfo1Txt.DataBind.SetBound(True, "@PSSIT_OSFT", "U_Adnl1")

            oInfo2Txt = oForm.Items.Item("txtinfo2").Specific
            oForm.Items.Item("lbloi2").Visible = False
            oForm.Items.Item("txtinfo2").Visible = False
            oInfo2Txt.DataBind.SetBound(True, "@PSSIT_OSFT", "U_Adnl2")

            oRemarksTxt = oForm.Items.Item("txtremark").Specific
            oRemarksTxt.DataBind.SetBound(True, "@PSSIT_OSFT", "U_Remarks")

            oActiveCheck = oForm.Items.Item("chkactive").Specific
            oActiveCheck.DataBind.SetBound(True, "@PSSIT_OSFT", "U_Active")
            oActiveCheck.Checked = True

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Handles all the SBO_Application event and executes as per the the event fired.
    ''' </summary>
    ''' <param name="FormUID"></param>
    ''' <param name="pVal"></param>
    ''' <param name="BubbleEvent"></param>
    ''' <remarks></remarks>
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            If pVal.FormUID = "FSFT" Then
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE) Then
                    SBO_Application = Nothing
                    GC.Collect()
                End If
                '******** Validation() method is called for validating the values in the edit text **********
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.ItemUID = "1" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = True Then
                        Try
                            Validation()
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End Try
                    End If
                End If
                '******** Duration is calculated using the DurationCalculation() function **********
                If pVal.CharPressed = Keys.Tab And pVal.BeforeAction = False Then
                    Dim oDuration As String
                    Dim oDurationHrs, BreakTime As String

                    If pVal.ItemUID = "txtsftime" Then
                        Try
                            If (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                'If oFromTimeTxt.String = Convert.ToDateTime(Date.Parse(oFromTimeTxt.String)) Then
                                BreakTime = BreakValidation()

                                If CInt(oBreakTimeTxt.Value) >= BreakTime Then
                                    SBO_Application.SetStatusBarMessage("Break in Mins should be Lesser", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    BubbleEvent = False
                                ElseIf CInt(oBreakTimeTxt.Value) <= BreakTime And oFromTimeTxt.Value.Length > 0 And oToTimeTxt.Value.Length > 0 Then
                                    oDuration = DurationMinsCalculation()
                                    oParentDB.SetValue("U_Duratmin", oParentDB.Offset, oDuration)
                                    oDurationHrs = DurationHrsCalculation()
                                    oDurationHrsTxt.Value = oDurationHrs
                                End If

                            End If
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End Try
                    End If
                    If pVal.ItemUID = "txtsttime" Then
                        Try
                            If (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                BreakTime = BreakValidation()

                                If CInt(oBreakTimeTxt.Value) >= BreakTime Then
                                    SBO_Application.SetStatusBarMessage("Break in Mins should be Lesser", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    BubbleEvent = False
                                ElseIf oBreakTimeTxt.Value <= BreakTime And oToTimeTxt.Value.Length > 0 And oFromTimeTxt.Value.Length > 0 Then
                                    oDuration = DurationMinsCalculation()
                                    oParentDB.SetValue("U_Duratmin", oParentDB.Offset, oDuration)
                                    oDurationHrs = DurationHrsCalculation()
                                    oDurationHrsTxt.Value = oDurationHrs
                                End If
                            End If
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End Try
                    End If
                    If pVal.ItemUID = "txtsbreak" Then
                        Try
                            If (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                BreakTime = BreakValidation()
                                If CInt(oBreakTimeTxt.Value) < 0 Then
                                    ' SBO_Application.SetStatusBarMessage("Break in Mins should not be '-ve' value", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    'BubbleEvent = False
                                    oBreakTimeTxt.Active = True
                                    Throw New Exception("Break in Mins should not be '-ve' value")

                                End If
                                If CInt(oBreakTimeTxt.Value) >= BreakTime Then
                                    SBO_Application.SetStatusBarMessage("Break in Mins should be Lesser", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    BubbleEvent = False
                                ElseIf CInt(oBreakTimeTxt.Value) <= CInt(BreakTime) And oFromTimeTxt.Value.Length > 0 And oToTimeTxt.Value.Length > 0 Then
                                    oDuration = DurationMinsCalculation()
                                    oParentDB.SetValue("U_Duratmin", oParentDB.Offset, oDuration)
                                    'Modified by Manimaran-----------s
                                    oDurationHrs = DurationHrsCalculation()
                                    oDurationHrsTxt.Value = oDurationHrs
                                    'oDurationHrsTxt.Value = oDuration
                                    'Modified by Manimaran-----------e
                                End If
                                'End If
                            End If
                        Catch ex As Exception
                            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End Try
                    End If
                End If

                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    '********** Add Button Press ***********
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            oForm.Items.Item("txtscode").Enabled = False
                            oForm.Items.Item("txtsdescr").Enabled = True
                        End If
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Try
                                oForm.Refresh()
                                oForm.Freeze(True)
                                oActiveCheck.Checked = True
                                oFromTimeTxt.String = FormatDateTime(Now(), DateFormat.ShortTime)
                                oToTimeTxt.String = FormatDateTime(Now(), DateFormat.ShortTime)
                                oBreakTimeTxt.Value = 0
                                oForm.Freeze(False)
                                oSCodeTxt.Active = True
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End Try
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    ''' <summary>
    ''' BreakTime is set to 0 (By default) and Duration is calculated.
    ''' Setting the focus to the ShiftCode Edittext.
    ''' </summary>
    ''' <param name="pVal"></param>
    ''' <param name="BubbleEvent"></param>
    ''' <remarks></remarks>
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Dim FormID As String
        Try
            FormID = SBO_Application.Forms.ActiveForm.UniqueID
            If pVal.MenuUID = "1281" And FormID = "FSFT" Then
                If pVal.BeforeAction = True Then
                    oForm.Items.Item("txtscode").Enabled = True
                    oForm.Items.Item("txtsdescr").Enabled = True
                End If
                If pVal.BeforeAction = False Then
                    oSCodeTxt.Active = True
                End If
            End If
            If pVal.MenuUID = "1282" And pVal.BeforeAction = False And FormID = "FSFT" Then
                oForm.Freeze(True)
                oParentDB.Offset = oParentDB.Size - 1
                oForm.Items.Item("txtscode").Enabled = True
                oForm.Items.Item("txtsdescr").Enabled = True
                oActiveCheck.Checked = True
                oBreakTimeTxt.Value = 0
                oFromTimeTxt.String = FormatDateTime(Now(), DateFormat.ShortTime)
                oToTimeTxt.String = FormatDateTime(Now(), DateFormat.ShortTime)
                'oDuration = DurationMinsCalculation()
                'oParentDB.SetValue("U_Duratmin", oParentDB.Offset, oDuration)
                'oDurationHrs = DurationHrsCalculation()
                'oDurationHrsTxt.Value = oDurationHrs
                oForm.Freeze(False)
                oSCodeTxt.Active = True
            End If
            '*****************************Navigation*******************************
            If (pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") And pVal.BeforeAction = False And FormID = "FSFT" Then

                oForm.Items.Item("txtscode").Enabled = False
                oForm.Items.Item("txtsdescr").Enabled = True
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    ''' <summary>
    ''' This function is for Break Time Validation
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function BreakValidation() As String
        Dim oFromTime, oToTime As DateTime
        Dim oBreakDiff As String
        Try
            If oFromTimeTxt.String <> "" Then
                oFromTime = Convert.ToDateTime(Date.Parse(oFromTimeTxt.String))
                oToTime = Convert.ToDateTime(Date.Parse(oToTimeTxt.String))
                Dim runLength As System.TimeSpan = oToTime.Subtract(oFromTime.ToShortTimeString)
                Dim secs As Integer = runLength.Seconds
                Dim minutes As Integer = runLength.Minutes
                Dim hours As Integer = runLength.Hours
                oBreakDiff = runLength.Hours * 60 + runLength.Minutes
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return oBreakDiff
    End Function
    ''' <summary>
    ''' This function is for calculating the duration in Mins.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function DurationMinsCalculation() As String
        Dim oFromTime, oToTime As DateTime
        Dim oBreakTime As Integer
        Dim oDuration As String
        Dim odurationhrs As String

        Try
            oFromTime = Convert.ToDateTime(Date.Parse(oFromTimeTxt.String))
            oToTime = Convert.ToDateTime(Date.Parse(oToTimeTxt.String))
            'oBreakTime = CInt(oBreakTimeTxt.Value)
            'Added by Manimaran-----------S
            Dim splitMin, splitHr As String
            splitMin = Strings.Right(CStr(oBreakTimeTxt.Value), 2)
            splitHr = Strings.Left(CStr(oBreakTimeTxt.Value), 2)
            oBreakTime = CDbl(splitHr) * 60 + CDbl(splitMin)
            oToTime = oToTime.AddMinutes(-oBreakTime)
            'Added by Manimaran-----------E
            Dim runLength As System.TimeSpan = oToTime.Subtract(oFromTime.ToShortTimeString)
            Dim secs As Integer = runLength.Seconds
            Dim minutes As Integer = runLength.Minutes
            Dim hours As Integer = runLength.Hours
            oDuration = runLength.Hours * 60 + runLength.Minutes
            oDurationHrs = runLength.Hours.ToString("00") + ":" + runLength.Minutes.ToString("00")
        Catch ex As Exception
            Throw ex
        End Try
        'Modified by Manimaran----s
        Return oDuration
        'Return oDurationHrs
        'Modified by Manimaran----e
    End Function
    ''' <summary>
    ''' This function is for calculating the duration in Hrs.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function DurationHrsCalculation() As String
        Dim oFromTime, oToTime As DateTime
        Dim oBreakTime As Integer
        Dim oDurationHrs As String
        Try
            oFromTime = Convert.ToDateTime(Date.Parse(oFromTimeTxt.String))
            oToTime = Convert.ToDateTime(Date.Parse(oToTimeTxt.String))
            'oBreakTime = CInt(oBreakTimeTxt.Value)
            'Added by Manimaran-----------S
            Dim splitMin, splitHr As String
            splitMin = Strings.Right(CStr(oBreakTimeTxt.Value), 2)
            splitHr = Strings.Left(CStr(oBreakTimeTxt.Value), 2)
            oBreakTime = CDbl(splitHr) * 60 + CDbl(splitMin)
            oToTime = oToTime.AddMinutes(-oBreakTime)
            'Added by Manimaran-----------E
            Dim runLength As System.TimeSpan = oToTime.Subtract(oFromTime.ToShortTimeString)
            Dim secs As Integer = runLength.Seconds
            Dim minutes As Integer = runLength.Minutes
            Dim hours As Integer = runLength.Hours
            oDurationHrs = runLength.Hours * 60 + runLength.Minutes
            oDurationHrs = runLength.Hours.ToString("00") + ":" + runLength.Minutes.ToString("00")
        Catch ex As Exception
            Throw ex
        End Try
        Return oDurationHrs
    End Function
    ''' <summary>
    ''' This method is used for validating the values in the EditText.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Validation()
        Dim BreakTime As String
        Dim oRs, oRs1 As SAPbobsCOM.Recordset
        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            If oSCodeTxt.Value.Length = 0 Then
                oSCodeTxt.Active = True
                Throw New Exception("Shift Code should not be Empty")
            End If
            If oDescTxt.Value.Length = 0 Then
                oDescTxt.Active = True
                Throw New Exception("Shift Name should not be Empty")
            End If
            If oFromTimeTxt.String = "00:00" Or oFromTimeTxt.String = "" Then
                oFromTimeTxt.Active = True
                Throw New Exception("From Time should not be Empty")
            End If
            If oToTimeTxt.String = "00:00" Or oToTimeTxt.String = "" Then
                oToTimeTxt.Active = True
                Throw New Exception("To Time should not be Empty")
            End If

            BreakTime = BreakValidation()
            If CInt(oBreakTimeTxt.Value) >= BreakTime Then
                oBreakTimeTxt.Active = True
                Throw New Exception("Break in Mins should be Lesser")
            End If
            'Added by Manimaran------s
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                oRs.DoQuery("select Code from [@PSSIT_OSFT]  where Code= '" & oSCodeTxt.Value & "' ")
                If oRs.RecordCount > 0 Then
                    oSCodeTxt.Active = True
                    Throw New Exception("Shift Code Already Exist")
                End If


                oRs1.DoQuery("select U_Sdescr from [@PSSIT_OSFT]  where U_Sdescr= '" & oDescTxt.Value & "' ")
                If oRs1.RecordCount > 0 Then
                    oDescTxt.Active = True
                    Throw New Exception("Shift Name Already Exist")
                End If
            End If

            If oForm.Items.Item("txtsdurhrs").Specific.string = "" Then
                oParentDB.SetValue("U_Duratmin", oParentDB.Offset, DurationMinsCalculation())
                oDurationHrsTxt.Value = DurationHrsCalculation()
            End If
            'Added by Manimaran------e
        Catch ex As Exception
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
End Class
