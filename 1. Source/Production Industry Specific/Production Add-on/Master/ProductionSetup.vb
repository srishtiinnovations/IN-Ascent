Imports System.Data.OleDb

Public Class ProductionSetup
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
    Private oParentDB, oActiveDB As SAPbouiCOM.DBDataSource
    '**************************UserTable***************************************
    Private PSSIT_OCON As SAPbobsCOM.UserTable
    '**************************Form************************************
    Private oForm As SAPbouiCOM.Form
    '**************************Items - EditText************************************
    Private oCodeTxt As SAPbouiCOM.EditText
    '**************************Items - CheckBox************************************
    Private oAccCheck, oFixedCostCheck, oLabCheck As SAPbouiCOM.CheckBox
    Private oSOHD, oPOHD, OSQLSERVEr, OSQLUID, OSQLPWD As SAPbouiCOM.EditText
    Private sPath As String = IO.Directory.GetParent(Application.ExecutablePath).ToString
#End Region
    ''' <summary>
    ''' SetApplication() method is called to connect the application through the connection string.
    ''' LoadFromXML("FrmLabour.srf") method is called to load the Labour form.
    ''' Drawform() method is called to Initialize the form,Datasources and Items in the form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company)
        MyBase.New(aSBO_Application, aCompany)
        SBO_Application = aSBO_Application
        oCompany = aCompany
        LoadFromXML("Frmsetup.srf")
        DrawForm()
        'oForm.DataBrowser.BrowseBy = "11"
    End Sub
    ''' <summary>
    ''' Initializing the instance of the active form to the form object.
    ''' Initializing the Datasources.
    ''' InitializeFormComponent() method is called to initialize the items.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub DrawForm()
        Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim StrSql As String
        Try
            oForm = SBO_Application.Forms.Item(SBO_Application.Forms.ActiveForm.UniqueID)
            oParentDB = oForm.DataSources.DBDataSources.Add("@PSSIT_OCON")
            PSSIT_OCON = UserTables.Item("PSSIT_OCON")
            oForm.Freeze(True)
            InitializeFormComponent()

            StrSql = "select * from [@PSSIT_OCON]"
            oRS.DoQuery(StrSql)
            If oRS.RecordCount > 0 Then
                oForm.EnableMenu("1288", False)
                oForm.EnableMenu("1289", False)
                oForm.EnableMenu("1290", False)
                oForm.EnableMenu("1291", False)
                oForm.EnableMenu("1282", False)
                oRS.MoveFirst()
                oParentDB.SetValue("Code", oParentDB.Offset, oRS.Fields.Item("Code").Value)
                If oRS.Fields.Item("U_Acckey").Value = "N" Then
                    oParentDB.SetValue("U_Acckey", oParentDB.Offset, "N")
                Else
                    oParentDB.SetValue("U_Acckey", oParentDB.Offset, "Y")
                End If
                If oRS.Fields.Item("U_Fcman").Value = "N" Then
                    oParentDB.SetValue("U_Fcman", oParentDB.Offset, "N")
                Else
                    oParentDB.SetValue("U_Fcman", oParentDB.Offset, "Y")
                End If
                If oRS.Fields.Item("U_Labman").Value = "N" Then
                    oParentDB.SetValue("U_Labman", oParentDB.Offset, "N")
                Else
                    oParentDB.SetValue("U_Labman", oParentDB.Offset, "Y")
                End If
                'Added by Manimaran------s
                oParentDB.SetValue("U_SOHDPer", oParentDB.Offset, oRS.Fields.Item("U_SOHDPer").Value)
                oParentDB.SetValue("U_POHDPer", oParentDB.Offset, oRS.Fields.Item("U_POHDPer").Value)
                'Added by Manimaran------e

                oParentDB.SetValue("U_SqlSer", oParentDB.Offset, oRS.Fields.Item("U_SqlSer").Value)
                oParentDB.SetValue("U_SqlUID", oParentDB.Offset, oRS.Fields.Item("U_SqlUID").Value)
                oParentDB.SetValue("U_SqlPwd", oParentDB.Offset, oRS.Fields.Item("U_SqlPwd").Value)


                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
            ElseIf oRS.RecordCount = 0 Then
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                oForm.Freeze(True)
                oCodeTxt.Value = GenerateSerialNo("PSSIT_OCON")
                oForm.Freeze(False)
            End If
            oForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        Finally
            oRS = Nothing
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' Configuring the items/controls in the form(.srf) by bounding to the object and to the DBDatasources.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitializeFormComponent()
        Try
            oCodeTxt = oForm.Items.Item("txtcode").Specific
            oForm.Items.Item("txtcode").Visible = False
            oForm.Items.Item("lblcode").Visible = False
            oCodeTxt.DataBind.SetBound(True, "@PSSIT_OCON", "Code")

            oAccCheck = oForm.Items.Item("chkacpost").Specific
            oAccCheck.DataBind.SetBound(True, "@PSSIT_OCON", "U_Acckey")
            oAccCheck.Checked = False

            oFixedCostCheck = oForm.Items.Item("chkfixcost").Specific
            oFixedCostCheck.DataBind.SetBound(True, "@PSSIT_OCON", "U_Fcman")
            oFixedCostCheck.Checked = False

            oLabCheck = oForm.Items.Item("chklbentry").Specific
            oLabCheck.DataBind.SetBound(True, "@PSSIT_OCON", "U_Labman")
            oLabCheck.Checked = False
            'Added by Manimaran------s
            oSOHD = oForm.Items.Item("11").Specific
            oSOHD.DataBind.SetBound(True, "@PSSIT_OCON", "U_SOHDPer")

            oPOHD = oForm.Items.Item("13").Specific
            oPOHD.DataBind.SetBound(True, "@PSSIT_OCON", "U_POHDPer")

            OSQLSERVEr = oForm.Items.Item("16").Specific
            OSQLSERVEr.DataBind.SetBound(True, "@PSSIT_OCON", "U_SqlSer")
            OSQLSERVEr.String = oCompany.Server

            OSQLUID = oForm.Items.Item("18").Specific
            OSQLUID.DataBind.SetBound(True, "@PSSIT_OCON", "U_SqlUID")

            OSQLPWD = oForm.Items.Item("20").Specific
            OSQLPWD.DataBind.SetBound(True, "@PSSIT_OCON", "U_SqlPwd")

            'Added by Manimaran------e
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#Region "Validate SQL Login Details"
    Private Function ValidateSQLLoginDetails() As Boolean
        Dim oConnection As OleDb.OleDbConnection
        Dim oCONNECTION_STRING, sqlUID, sqlPWd As String
        sqlUID = OSQLUID.String
        sqlPWd = OSQLPWD.String
        oCONNECTION_STRING = "Provider=SQLOLEDB;Server=" & oCompany.Server & ";Database=" & oCompany.CompanyDB & ";User ID=" & sqlUID & ";Password=" & sqlPWd
        oConnection = New OleDbConnection(oCONNECTION_STRING)
        Try
            oConnection.Open()
        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function
#End Region

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            If pVal.FormUID = "FPS" Then
                '*****************Item Pressed*****************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.BeforeAction = True Then
                        If pVal.ItemUID = "1" Then

                            Try
                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                    '****************Adding the child data to the database table***********
                                    Dim SetupTable As Integer
                                    Try
                                        If Not oCompany.InTransaction Then
                                            oCompany.StartTransaction()
                                        End If


                                        '************** Records Added in to the table**********
                                        If PSSIT_OCON.GetByKey(oCodeTxt.Value) = True Then
                                            PSSIT_OCON.Code = oCodeTxt.Value
                                            PSSIT_OCON.Name = oCodeTxt.Value
                                            PSSIT_OCON.UserFields.Fields.Item("U_Acckey").Value = oParentDB.GetValue("U_Acckey", oParentDB.Offset).Trim()
                                            PSSIT_OCON.UserFields.Fields.Item("U_Fcman").Value = oParentDB.GetValue("U_Fcman", oParentDB.Offset).Trim()
                                            PSSIT_OCON.UserFields.Fields.Item("U_Labman").Value = oParentDB.GetValue("U_Labman", oParentDB.Offset).Trim()
                                            'Added by Manimaran----s
                                            PSSIT_OCON.UserFields.Fields.Item("U_SOHDPer").Value = oParentDB.GetValue("U_SOHDPer", oParentDB.Offset).Trim()
                                            PSSIT_OCON.UserFields.Fields.Item("U_POHDPer").Value = oParentDB.GetValue("U_POHDPer", oParentDB.Offset).Trim()

                                            PSSIT_OCON.UserFields.Fields.Item("U_SqlSer").Value = oParentDB.GetValue("U_SqlSer", oParentDB.Offset).Trim()
                                            PSSIT_OCON.UserFields.Fields.Item("U_SqlUID").Value = oParentDB.GetValue("U_SqlUID", oParentDB.Offset).Trim()
                                            PSSIT_OCON.UserFields.Fields.Item("U_SqlPwd").Value = oParentDB.GetValue("U_SqlPwd", oParentDB.Offset).Trim()
                                            'Added by Manimaran----e
                                            SetupTable = PSSIT_OCON.Update()
                                        Else
                                            PSSIT_OCON.Code = oCodeTxt.Value
                                            PSSIT_OCON.Name = oCodeTxt.Value
                                            PSSIT_OCON.UserFields.Fields.Item("U_Acckey").Value = oParentDB.GetValue("U_Acckey", oParentDB.Offset).Trim()
                                            PSSIT_OCON.UserFields.Fields.Item("U_Fcman").Value = oParentDB.GetValue("U_Fcman", oParentDB.Offset).Trim()
                                            PSSIT_OCON.UserFields.Fields.Item("U_Labman").Value = oParentDB.GetValue("U_Labman", oParentDB.Offset).Trim()
                                            'Added by Manimaran----s
                                            PSSIT_OCON.UserFields.Fields.Item("U_SOHDPer").Value = oParentDB.GetValue("U_SOHDPer", oParentDB.Offset).Trim()
                                            PSSIT_OCON.UserFields.Fields.Item("U_POHDPer").Value = oParentDB.GetValue("U_POHDPer", oParentDB.Offset).Trim()
                                            'Added by Manimaran----e

                                            PSSIT_OCON.UserFields.Fields.Item("U_SqlSer").Value = oParentDB.GetValue("U_SqlSer", oParentDB.Offset).Trim()
                                            PSSIT_OCON.UserFields.Fields.Item("U_SqlUID").Value = oParentDB.GetValue("U_SqlUID", oParentDB.Offset).Trim()
                                            PSSIT_OCON.UserFields.Fields.Item("U_SqlPwd").Value = oParentDB.GetValue("U_SqlPwd", oParentDB.Offset).Trim()

                                            SetupTable = PSSIT_OCON.Add()
                                        End If
                                        If SetupTable = 0 Then
                                            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                            SBO_Application.StatusBar.SetText("Operation Completed Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            BubbleEvent = False
                                        Else
                                            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                        End If
                                    Catch ex As Exception
                                        SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        BubbleEvent = False
                                    End Try
                                End If
                            Catch ex As Exception
                                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                            End Try
                        ElseIf pVal.ItemUID = "21" Then
                            If ValidateSQLLoginDetails() = False Then
                                SBO_Application.SetStatusBarMessage("Sql Login failed. Check the Login Details", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                                Exit Sub
                            Else
                                SBO_Application.SetStatusBarMessage("Login success", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                            End If
                        End If
                    End If
                End If

                If pVal.BeforeAction = False Then
                    '*****************************Refreshing the form to initiate default values*******************************
                    If pVal.ItemUID = "1" Then
                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            oCodeTxt.Value = GenerateSerialNo("PSSIT_OCON")
                            oForm.Refresh()
                        End If
                    End If
                End If
                '*****************************Releasing the Com Object*******************************
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                    SBO_Application = Nothing
                    GC.Collect()
                End If
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub

    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Dim FormID As String
        Try
            FormID = SBO_Application.Forms.ActiveForm.UniqueID
            If pVal.MenuUID = "1282" And pVal.BeforeAction = False And FormID = "FPS" Then
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    oForm.Freeze(True)
                    oCodeTxt.Value = GenerateSerialNo("PSSIT_OCON")
                    oForm.Freeze(False)
                End If
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
End Class
