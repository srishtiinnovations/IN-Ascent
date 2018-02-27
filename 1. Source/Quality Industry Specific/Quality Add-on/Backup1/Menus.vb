Public Class Menus
    Public WithEvents SBO_Application As SAPbouiCOM.Application
    Public oCompany As SAPbobsCOM.Company
    Public objUIXml As SST.UIXML
    Public objGenFunc As SST.GeneralFunctions
    Private sPath As String = IO.Directory.GetParent(Application.ExecutablePath).ToString
    Public ReaMst As ClsReaMst
    Public ReaCat As clsReaCat
    Public ParaCat As clsParaCat
    Public ParaMst As clsParaMst
    Public ItmPI As clsItemParaInward
    Public PrdItm As clsProditem
    Public PrdSmpPl As clsProdSampPlan
    Public SamPlan As clsSamPl
    Public ProdCons As clsProdCons
    Public InwardCons As clsInwardCons
    Public InwardInsp As clsInwInsp
    Public ProdInsp As clsProdIns
    Public UOM As clsUOM

    'Public cat As clsCategoryMaster

    Public GE As clsGateEntry
    Public SCGE As clsSCGateEntry

    Public STUP As clsSetUp
    Public GRPO As clsGRPO

    Public PROD As clsproduction

    Public CFL As clsUserCFL
    Public CFL1 As clsUserCFL1
 
    Public AQL As clsAccpLmt

    Public Sampling As clsSamplingLevel

#Region "Menus"

    Public Sub Intialize()
        Dim objSBOConnector As New SST.SBOConnector
        SBO_Application = objSBOConnector.GetApplication(System.Environment.GetCommandLineArgs.GetValue(1))
        SBO_Application = objSBOConnector.GetApplication(System.Environment.GetCommandLineArgs.GetValue(1))
        oCompany = objSBOConnector.GetCompany(SBO_Application)
        createObjects()
        'EventFilters()
        LoadMenus("MainMenu.xml")
        SBO_Application.SetStatusBarMessage("QC Add-On Connected Succesfully!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
    End Sub

    Private Sub EventFilters()
        Dim oFilters As SAPbouiCOM.EventFilters
        Dim oFilter As SAPbouiCOM.EventFilter
        oFilters = New SAPbouiCOM.EventFilters

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
        oFilter.AddEx("Frm_PrdItm")
        oFilter.AddEx("Frm_PMst")
        oFilter.AddEx("Frm_ItmPm")
        oFilter.AddEx("Frm_GE")
        oFilter.AddEx("Frm_Reas")
        oFilter.AddEx("Frm_SamPl")
        oFilter.AddEx("Frm_PrdPl")
        oFilter.AddEx("Frm_InwCons")
        oFilter.AddEx("Frm_InwInsp")
        oFilter.AddEx("Frm_PrdCons")
        oFilter.AddEx("Frm_PrdInsp")
        oFilter.AddEx("Frm_STP")

        oFilter.AddEx("Frm_SmplLvl")
        oFilter.AddEx("Frm_AccpLmt")

        oFilter.AddEx("143")
        oFilter.AddEx("Frm_RFIInsp")
        'oFilter.AddEx("Frm_Login")

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
        oFilter.AddEx("Frm_PrdItm")
        oFilter.AddEx("Frm_PMst")
        oFilter.AddEx("Frm_ItmPm")
        oFilter.AddEx("Frm_GE")
        oFilter.AddEx("Frm_Reas")
        oFilter.AddEx("Frm_SamPl")
        oFilter.AddEx("Frm_PrdPl")
        oFilter.AddEx("Frm_InwCons")
        oFilter.AddEx("Frm_InwInsp")
        oFilter.AddEx("Frm_PrdCons")
        oFilter.AddEx("Frm_PrdInsp")
        oFilter.AddEx("Frm_STP")

        oFilter.AddEx("Frm_SmplLvl")
        oFilter.AddEx("Frm_AccpLmt")

        oFilter.AddEx("143")
        oFilter.AddEx("Frm_RFIInsp")

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK)
        oFilter.AddEx("Frm_PrdItm")
        oFilter.AddEx("Frm_PMst")
        oFilter.AddEx("Frm_ItmPm")
        oFilter.AddEx("Frm_GE")
        oFilter.AddEx("Frm_Reas")
        oFilter.AddEx("Frm_SamPl")
        oFilter.AddEx("Frm_PrdPl")
        oFilter.AddEx("143")
        oFilter.AddEx("Frm_RFIInsp")
        oFilter.AddEx("Frm_SmplLvl")
        oFilter.AddEx("Frm_AccpLmt")

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
        oFilter.AddEx("Frm_PrdItm")
        oFilter.AddEx("Frm_PMst")
        oFilter.AddEx("Frm_ItmPm")
        oFilter.AddEx("Frm_GE")
        oFilter.AddEx("Frm_Reas")
        oFilter.AddEx("Frm_SamPl")
        oFilter.AddEx("Frm_PrdPl")
        oFilter.AddEx("Frm_InwCons")
        oFilter.AddEx("Frm_InwInsp")
        oFilter.AddEx("Frm_PrdCons")
        oFilter.AddEx("Frm_PrdInsp")
        oFilter.AddEx("Frm_STP")
        oFilter.AddEx("Frm_SmplLvl")
        oFilter.AddEx("Frm_AccpLmt")
        oFilter.AddEx("143")
        oFilter.AddEx("Frm_CFL")
        oFilter.AddEx("Frm_RFIInsp")
        'oFilter.AddEx("Frm_Login")

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
        oFilter.AddEx("Frm_PrdItm")
        oFilter.AddEx("Frm_PMst")
        oFilter.AddEx("Frm_ItmPm")
        oFilter.AddEx("Frm_GE")
        oFilter.AddEx("Frm_Reas")
        oFilter.AddEx("Frm_SamPl")
        oFilter.AddEx("Frm_PrdPl")
        oFilter.AddEx("Frm_PrdCons")
        oFilter.AddEx("Frm_InwInsp")
        oFilter.AddEx("143")
        oFilter.AddEx("Frm_PrdInsp")
        oFilter.AddEx("Frm_InwCons")
        oFilter.AddEx("Frm_RFIInsp")
        'oFilter.AddEx("Frm_Login")
        oFilter.AddEx("Frm_SmplLvl")
        oFilter.AddEx("Frm_AccpLmt")

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_CLOSE)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK)

        SBO_Application.SetFilter(oFilters)
    End Sub

    Private Sub createObjects()
        objUIXml = New SST.UIXML(SBO_Application)
        objGenFunc = New SST.GeneralFunctions(oCompany)
        ReaMst = New ClsReaMst
        ReaCat = New clsReaCat
        ParaCat = New clsParaCat
        ParaMst = New clsParaMst
        ItmPI = New clsItemParaInward
        PrdItm = New clsProditem
        PrdSmpPl = New clsProdSampPlan
        SamPlan = New clsSamPl
        AQL = New clsAccpLmt
        Sampling = New clsSamplingLevel
        ProdCons = New clsProdCons
        InwardCons = New clsInwardCons
        InwardInsp = New clsInwInsp
        ProdInsp = New clsProdIns
        UOM = New clsUOM
        GE = New clsGateEntry
        SCGE = New clsSCGateEntry
        STUP = New clsSetUp
        GRPO = New clsGRPO
        PROD = New clsproduction
        CFL = New clsUserCFL

        CFL1 = New clsUserCFL1
    End Sub

    Public Sub LoadMenus(ByVal XMLFile As String)
        Dim oXML As New System.Xml.XmlDocument
        Dim strXML As String
        Dim strResource As String
        Dim oIcon As SAPbouiCOM.MenuItem
        Try
            strResource = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name & "." & XMLFile
            oXML.Load(System.Reflection.Assembly.GetExecutingAssembly.GetManifestResourceStream(strResource))
            strXML = oXML.InnerXml()
            SBO_Application.LoadBatchActions(strXML)
            oIcon = objAddOn.SBO_Application.Menus.Item("QCT")
            oIcon.Image = Application.StartupPath & "\QC.bmp"
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

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

    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Dim FormType As String
        FormType = SBO_Application.Forms.ActiveForm.Type
        Try
            If pVal.BeforeAction = False Then
                If pVal.MenuUID = "REAMST" Then
                    ReaMst.LoadScreen()
                ElseIf pVal.MenuUID = "REACAT" Then
                    ReaCat.LoadScreen()
                ElseIf pVal.MenuUID = "PACAT" Then
                    ParaCat.LoadScreen()
                ElseIf pVal.MenuUID = "PAMST" Then
                    ParaMst.LoadScreen()
                ElseIf pVal.MenuUID = "ITMP" Then
                    ItmPI.LoadScreen()
                ElseIf pVal.MenuUID = "ITMPPD" Then
                    PrdItm.LoadScreen()
                ElseIf pVal.MenuUID = "SMPLPD" Then
                    PrdSmpPl.LoadScreen()
                ElseIf pVal.MenuUID = "SMPL" Then
                    SamPlan.LoadScreen()
                ElseIf pVal.MenuUID = "AccpLmt" Then
                    AQL.LoadScreen()
                ElseIf pVal.MenuUID = "Sampling" Then
                    Sampling.LoadScreen()
                ElseIf pVal.MenuUID = "PDIPCN" Then
                    ProdCons.LoadScreen()
                ElseIf pVal.MenuUID = "CNENIW" Then
                    InwardCons.LoadScreen()
                ElseIf pVal.MenuUID = "IWIP" Then
                    InwardInsp.LoadScreen()
                ElseIf pVal.MenuUID = "IPPD" Then
                    ProdInsp.LoadScreen()
                ElseIf pVal.MenuUID = "UOM" Then
                    UOM.LoadScreen()
                ElseIf pVal.MenuUID = "GE" Then
                    GE.LoadScreen()
                ElseIf pVal.MenuUID = "SCGE" Then
                    SCGE.LoadScreen()
                ElseIf pVal.MenuUID = "STP" Then
                    STUP.LoadScreen()
                End If


                Select Case objAddOn.SBO_Application.Forms.ActiveForm.TypeEx
                    Case clsGateEntry.formtype
                        GE.MenuEvent(pVal, BubbleEvent)

                    Case clsSCGateEntry.formtype
                        SCGE.MenuEvent(pVal, BubbleEvent)

                    Case clsInwInsp.formtype
                        InwardInsp.MenuEvent(pVal, BubbleEvent)
                    Case clsInwardCons.formtype
                        InwardCons.MenuEvent(pVal, BubbleEvent)
                    Case clsProdIns.formtype
                        ProdInsp.MenuEvent(pVal, BubbleEvent)
                    Case clsProdCons.formtype
                        ProdCons.MenuEvent(pVal, BubbleEvent)
                    Case clsSamPl.formtype
                        SamPlan.MenuEvent(pVal, BubbleEvent)
                    Case clsAccpLmt.formtype
                        AQL.MenuEvent(pVal, BubbleEvent)
                    Case clsSamplingLevel.formtype
                        Sampling.MenuEvent(pVal, BubbleEvent)
                    Case clsItemParaInward.formtype
                        ItmPI.MenuEvent(pVal, BubbleEvent)
                    Case clsProditem.formtype
                        PrdItm.MenuEvent(pVal, BubbleEvent)
                    Case clsProdSampPlan.formtype
                        PrdSmpPl.MenuEvent(pVal, BubbleEvent)
                        
                End Select

            Else
                If pVal.MenuUID = "DelRow" Then
                    Select Case objAddOn.SBO_Application.Forms.ActiveForm.TypeEx
                        Case clsSamplingLevel.formtype
                            Sampling.MenuEvent(pVal, BubbleEvent)
                        Case clsAccpLmt.formtype
                            AQL.MenuEvent(pVal, BubbleEvent)
                    End Select
                End If


                Select Case objAddOn.SBO_Application.Forms.ActiveForm.TypeEx

                    Case clsGateEntry.formtype
                        If pVal.MenuUID = "1283" Then
                            SBO_Application.SetStatusBarMessage("Transactions cannot be deleted", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If

                    Case clsSCGateEntry.formtype
                        If pVal.MenuUID = "1283" Then
                            SBO_Application.SetStatusBarMessage("Transactions cannot be deleted", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                      
                    Case clsInwInsp.formtype
                        If pVal.MenuUID = "1283" Then
                            SBO_Application.SetStatusBarMessage("Transactions cannot be deleted", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                    Case clsInwardCons.formtype
                        If pVal.MenuUID = "1283" Then
                            SBO_Application.SetStatusBarMessage("Transactions cannot be deleted", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                    Case clsProdIns.formtype
                        If pVal.MenuUID = "1283" Then
                            SBO_Application.SetStatusBarMessage("Transactions cannot be deleted", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                    Case clsProdCons.formtype
                        If pVal.MenuUID = "1283" Then
                            SBO_Application.SetStatusBarMessage("Transactions cannot be deleted", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                    Case clsItemParaInward.formtype
                        If pVal.MenuUID = "1283" Then
                            SBO_Application.SetStatusBarMessage("Masters cannot be deleted", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                    Case clsParaMst.formtype
                        If pVal.MenuUID = "1283" Then
                            SBO_Application.SetStatusBarMessage("Masters cannot be deleted", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                    Case clsProditem.formtype
                        If pVal.MenuUID = "1283" Then
                            SBO_Application.SetStatusBarMessage("Masters cannot be deleted", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                    Case clsProdSampPlan.formtype
                        If pVal.MenuUID = "1283" Then
                            SBO_Application.SetStatusBarMessage("Masters cannot be deleted", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                    Case ClsReaMst.formtype
                        If pVal.MenuUID = "1283" Then
                            SBO_Application.SetStatusBarMessage("Masters cannot be deleted", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                    Case clsSamPl.formtype
                        If pVal.MenuUID = "1283" Then
                            SBO_Application.SetStatusBarMessage("Masters cannot be deleted", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                        '**************

                    Case clsSamplingLevel.formtype
                        If pVal.MenuUID = "1283" Then
                            SBO_Application.SetStatusBarMessage("Masters cannot be deleted", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If

                    Case clsAccpLmt.formtype
                        If pVal.MenuUID = "1283" Then
                            SBO_Application.SetStatusBarMessage("Masters cannot be deleted", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                        '****************
                    Case clsSetUp.formtype
                        If pVal.MenuUID = "1283" Then
                            SBO_Application.SetStatusBarMessage("Masters cannot be deleted", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                    Case clsUOM.formtype
                        If pVal.MenuUID = "1283" Then
                            SBO_Application.SetStatusBarMessage("Masters cannot be deleted", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                     
                    Case clsReaCat.formtype
                        If pVal.MenuUID = "1283" Then
                            SBO_Application.SetStatusBarMessage("Masters cannot be deleted", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                    Case clsParaCat.formtype
                        If pVal.MenuUID = "1283" Then
                            SBO_Application.SetStatusBarMessage("Masters cannot be deleted", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                End Select
                If objAddOn.SBO_Application.Forms.ActiveForm.TypeEx = "SST_MOD" Then
                    If pVal.MenuUID = "1283" Then
                        SBO_Application.SetStatusBarMessage("Masters cannot be deleted", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        BubbleEvent = False
                    End If
                End If
                If objAddOn.SBO_Application.Forms.ActiveForm.TypeEx = "SST_STG" Then
                    If pVal.MenuUID = "1283" Then
                        SBO_Application.SetStatusBarMessage("Masters cannot be deleted", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        BubbleEvent = False
                    End If
                End If
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
        End Try
    End Sub

#End Region
#Region ""
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try

            Select Case pVal.FormTypeEx
                Case ClsReaMst.formtype
                    ReaMst.ItemEvent(FormUID, pVal, BubbleEvent)
                Case clsParaMst.formtype
                    ParaMst.ItemEvent(FormUID, pVal, BubbleEvent)
                Case clsProdIns.formtype
                    ProdInsp.ItemEvent(FormUID, pVal, BubbleEvent)
                Case clsGateEntry.formtype
                    GE.ItemEvent(FormUID, pVal, BubbleEvent)

                Case clsSCGateEntry.formtype
                    SCGE.ItemEvent(FormUID, pVal, BubbleEvent)

                Case clsSamPl.formtype
                    SamPlan.ItemEvent(FormUID, pVal, BubbleEvent)

                    '**********
                    'Case clsSmplLvl.formtype
                    '    SLM.ItemEvent(FormUID, pVal, BubbleEvent)
                Case clsAccpLmt.formtype
                    AQL.ItemEvent(FormUID, pVal, BubbleEvent)

                Case clsSamplingLevel.formtype
                    Sampling.ItemEvent(FormUID, pVal, BubbleEvent)
                    '**********

                Case clsItemParaInward.formtype
                    ItmPI.ItemEvent(FormUID, pVal, BubbleEvent)
                Case clsProdSampPlan.formtype
                    PrdSmpPl.ItemEvent(FormUID, pVal, BubbleEvent)
                Case clsProditem.formtype
                    PrdItm.ItemEvent(FormUID, pVal, BubbleEvent)
                Case clsInwInsp.formtype
                    InwardInsp.ItemEvent(FormUID, pVal, BubbleEvent)
                Case clsInwardCons.formtype
                    InwardCons.ItemEvent(FormUID, pVal, BubbleEvent)
                Case clsSetUp.formtype
                    STUP.ItemEvent(FormUID, pVal, BubbleEvent)
                Case clsProdCons.formtype
                    ProdCons.ItemEvent(FormUID, pVal, BubbleEvent)
                Case clsGRPO.formtype
                    GRPO.ItemEvent(FormUID, pVal, BubbleEvent)

                Case clsproduction.formtype
                    PROD.ItemEvent(FormUID, pVal, BubbleEvent)

                Case clsUserCFL.formtype
                    CFL.ItemEvent(FormUID, pVal, BubbleEvent)

                Case clsUserCFL1.formtype
                    CFL1.ItemEvent(FormUID, pVal, BubbleEvent)

            End Select

        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub

#End Region

    Private Sub SBO_Application_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.RightClickEvent
        Try
            If eventInfo.BeforeAction Then
                If eventInfo.FormUID.Contains(Left(clsSamPl.formtype, 8)) Then
                    SamPlan.RightClickEvent(eventInfo, BubbleEvent)
                End If
                If eventInfo.FormUID.Contains(Left(clsSamplingLevel.formtype, 8)) Then
                    Sampling.RightClickEvent(eventInfo, BubbleEvent)
                End If

                If eventInfo.FormUID.Contains(Left(clsAccpLmt.formtype, 8)) Then
                    AQL.RightClickEvent(eventInfo, BubbleEvent)
                End If

                If eventInfo.FormUID.Contains(Left(clsItemParaInward.formtype, 8)) Then
                    ItmPI.RightClickEvent(eventInfo, BubbleEvent)
                End If
                If eventInfo.FormUID.Contains(Left(clsProdSampPlan.formtype, 8)) Then
                    PrdSmpPl.RightClickEvent(eventInfo, BubbleEvent)
                End If
                If eventInfo.FormUID.Contains(Left(clsProditem.formtype, 8)) Then
                    PrdItm.RightClickEvent(eventInfo, BubbleEvent)
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class

