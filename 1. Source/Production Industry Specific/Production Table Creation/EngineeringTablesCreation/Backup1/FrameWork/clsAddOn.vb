
Public Class clsAddOn
    'DECLARE SBO OBJECT
    Public WithEvents objApplication As SAPbouiCOM.Application
    Public objCompany As SAPbobsCOM.Company
    Public objCompanyInfo As SAPbobsCOM.CompanyInfo
    Public oProgBar As SAPbouiCOM.ProgressBar
    Private objForm As SAPbouiCOM.Form

    'DECLARE LIBRARY OBJECTS
    Public objGenFunc As Altrion.SBOLib.GeneralFunctions
    Public objUIXml As Altrion.SBOLib.UIXML

    'DECLARE BUSINESS OBJECTS HERE


    Public Sub New()

    End Sub

    Public Sub Intialize()
        Dim objSBOConnector As New Altrion.SBOLib.SBOConnector
        objApplication = objSBOConnector.GetApplication(System.Environment.GetCommandLineArgs.GetValue(1))
        objCompany = objSBOConnector.GetCompany(objApplication)
        Try
            createTables()
            CreateUDOS()
        Catch ex As Exception
            'MsgBox(ex.ToString)
            ' MsgBox(objAddOn.objCompany.GetLastErrorDescription())
            End
        End Try
        objApplication.SetStatusBarMessage("Engineering Addon Tables Created successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
    End Sub
    Private Sub createObjects()
        'Library Object Initilisation
        objGenFunc = New Altrion.SBOLib.GeneralFunctions(objCompany)
        objUIXml = New Altrion.SBOLib.UIXML(objApplication)
        'Business Object Initialisation
       

    End Sub
    Private Sub objApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles objApplication.AppEvent
        If EventType = SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown Then
            Try
                objUIXml.LoadMenuXML("RemoveMenu.xml", Altrion.SBOLib.UIXML.enuResourceType.Embeded)
                If objCompany.Connected Then objCompany.Disconnect()
                objCompany = Nothing
                objApplication = Nothing
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objCompany)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objApplication)
            Catch ex As Exception
            End Try
            End
        End If
    End Sub
    Private Sub objApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles objApplication.ItemEvent
        Try
            '  objApplication.SetStatusBarMessage("", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            Select Case pVal.FormTypeEx
              
            End Select
        Catch ex As Exception
            '  MsgBox(ex.ToString)

        End Try
    End Sub
    Private Sub objApplication_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles objApplication.MenuEvent
        Try
            If pVal.BeforeAction Then
                ' Dim objParentForm As SAPbouiCOM.Form
                'objParentForm = objAddOn.objInvoice.objReturnForm
                'If pVal.MenuUID = "519" Then

                'End If
                ' End If
            Else
               
               
            End If

        Catch ex As Exception
            ' MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub objApplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles objApplication.RightClickEvent

        If eventInfo.BeforeAction Then
            objForm = objApplication.Forms.Item(eventInfo.FormUID)

           
            'If eventInfo.FormUID.Contains("BusinessInfo") And eventInfo.ItemUID = "68" Then
            'objFactor.RightClickEvent(eventInfo, BubbleEvent)
            'End If

        End If
    End Sub
    

    Private Sub loadMenu()
        'If objApplication.Menus.Exists("USHA") Then
        '    objUIXml.LoadMenuXML("RemoveMenu.xml", Altrion.SBOLib.UIXML.enuResourceType.Embeded)
        'End If
        'objUIXml.LoadMenuXML("AddMenu.xml", Altrion.SBOLib.UIXML.enuResourceType.Embeded)
    End Sub

    Private Sub createTables()

        Dim objUDFEngine As New Altrion.SBOLib.UDFEngine(objCompany)

        objApplication.SetStatusBarMessage("Tables Creating Please Wait........!", SAPbouiCOM.BoMessageTime.bmt_Long, False)
        objUDFEngine.CreateTable("PSSIT_PMWCBREAKHDR", "WCBreakDown", SAPbobsCOM.BoUTBTableType.bott_Document)
        objUDFEngine.AddDateField("@PSSIT_PMWCBREAKHDR", "docdate", "DocumentDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCBREAKHDR", "deptcode", "DepartmentCode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCBREAKHDR", "deptdesc", "DepartmentDesc", 60)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCBREAKHDR", "wccode", "WorkCenterCode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCBREAKHDR", "wcname", "WorkCenterName", 60)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCBREAKHDR", "wcno", "WorkCenterNo", 20)
        objUDFEngine.AddDateField("@PSSIT_PMWCBREAKHDR", "indate", "IntimationDate", SAPbobsCOM.BoFldSubTypes.st_None)
        'Commented by Manimaran-----s
        'objUDFEngine.AddDateField("@PSSIT_PMWCBREAKHDR", "intime", "IntimationTime", SAPbobsCOM.BoFldSubTypes.st_Time)
        'Commented by Manimaran-----e
        objUDFEngine.AddAlphaMemoField("@PSSIT_PMWCBREAKHDR", "natuwork", "NatureofWork", 64000)
        'Commented by Manimaran-----s
        'objUDFEngine.AddDateField("@PSSIT_PMWCBREAKHDR", "attdate", "AttendanceDate", SAPbobsCOM.BoFldSubTypes.st_None)
        'objUDFEngine.AddDateField("@PSSIT_PMWCBREAKHDR", "atttime", "AttendanceTime", SAPbobsCOM.BoFldSubTypes.st_Time)
        'objUDFEngine.AddDateField("@PSSIT_PMWCBREAKHDR", "compdate", "CompletionDate", SAPbobsCOM.BoFldSubTypes.st_None)
        'objUDFEngine.AddDateField("@PSSIT_PMWCBREAKHDR", "comptime", "CompletionTime", SAPbobsCOM.BoFldSubTypes.st_Time)
        'objUDFEngine.AddNumericField("@PSSIT_PMWCBREAKHDR", "totmts", "TotalMinitues", 10)
        'Commented by Manimaran-----e
        objUDFEngine.AddAlphaField("@PSSIT_PMWCBREAKHDR", "empid", "EmployeeID", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCBREAKHDR", "empname", "EmployeeName", 40)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCBREAKHDR", "SCode", "SftCode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCBREAKHDR", "SDesc", "SftDesc", 60)
        'Added by Manimaran---------s
        objUDFEngine.AddAlphaField("@PSSIT_PMWCBREAKHDR", "PENum", "Production Entry number", 20)
        'Added by Manimaran---------e

        objUDFEngine.CreateTable("PSSIT_PMWCREMEDTL", "WCBreakDownRemid", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCREMEDTL", "remecode", "Remedicode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCREMEDTL", "remedesc", "RemediDesc", 60)


        objUDFEngine.CreateTable("PSSIT_PMWCITEMSDTL", "WCBreakDownItem", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCITEMSDTL", "itemcode", "ItemCode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCITEMSDTL", "itemdesc", "ItemDescription", 60)
        objUDFEngine.AddFloatField("@PSSIT_PMWCITEMSDTL", "itemqty", "ItemQuantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCITEMSDTL", "itemuom", "ItemUOM", 10)
        objUDFEngine.AddFloatField("@PSSIT_PMWCITEMSDTL", "itemrate", "ItemRate", SAPbobsCOM.BoFldSubTypes.st_Rate)
        objUDFEngine.AddFloatField("@PSSIT_PMWCITEMSDTL", "itemval", "ItemValue", SAPbobsCOM.BoFldSubTypes.st_Price)


        objUDFEngine.CreateTable("PSSIT_PMWCREASONDTL", "WCBreakDownReason", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCREASONDTL", "reascode", "Reasoncode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCREASONDTL", "reasdesc", "ReasonDesc", 60)
        'Added by Manimaran----------s
        objUDFEngine.AddDateField("@PSSIT_PMWCREASONDTL", "StTime", "Start Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objUDFEngine.AddDateField("@PSSIT_PMWCREASONDTL", "EndTime", "End Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objUDFEngine.AddNumericField("@PSSIT_PMWCREASONDTL", "StpgTime", "Stoppage Time", 10)
        'Added by Manimaran----------e

        objUDFEngine.CreateTable("PSSIT_OLBR", "Labour", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objUDFEngine.AddAlphaField("@PSSIT_OLBR", "Empid", "EmployeeNo.", 20)
        objUDFEngine.AddAlphaField("@PSSIT_OLBR", "Empnam", "EmployeeName", 60)
        objUDFEngine.AddAlphaField("@PSSIT_OLBR", "LGCode", "SkillGroup", 25)
        objUDFEngine.AddAlphaField("@PSSIT_OLBR", "LGname", "SkillGroupName", 50)
        objUDFEngine.AddAlphaField("@PSSIT_OLBR", "Currncy", "Currency", 25)
        objUDFEngine.AddFloatField("@PSSIT_OLBR", "Labrate", "LabourRate/Hour", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddAlphaField("@PSSIT_OLBR", "Accode", "AccountCode", 30)
        objUDFEngine.AddAlphaField("@PSSIT_OLBR", "Acname", "AccountName", 100)
        objUDFEngine.AddAlphaField("@PSSIT_OLBR", "Adnl1", "OtherInfo1", 50)
        objUDFEngine.AddAlphaField("@PSSIT_OLBR", "Adnl2", "OtherInfo2", 50)
        objUDFEngine.AddAlphaField("@PSSIT_OLBR", "Remarks", "Remarks", 100)
        objUDFEngine.addField("@PSSIT_OLBR", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")
        objUDFEngine.AddAlphaField("@PSSIT_OLBR", "ActAcCode", "ActualAccountCode", 30)

        objUDFEngine.CreateTable("PSSIT_OMGP", "MachineGroup", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objUDFEngine.AddAlphaField("@PSSIT_OMGP", "Mgname", "MachineGroupName", 50)
        objUDFEngine.AddAlphaField("@PSSIT_OMGP", "WCcode", "WorkCentre", 25)
        objUDFEngine.AddAlphaField("@PSSIT_OMGP", "WCName", "WorkCentreName", 50)
        objUDFEngine.AddAlphaField("@PSSIT_OMGP", "RCurrncy", "RunningCurrency", 25)
        objUDFEngine.AddFloatField("@PSSIT_OMGP", "Runrate", "RunningRate/Hour", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddAlphaField("@PSSIT_OMGP", "RAccode", "AccountCode", 30)
        objUDFEngine.AddAlphaField("@PSSIT_OMGP", "RAcname", "AccountName", 100)
        objUDFEngine.AddAlphaField("@PSSIT_OMGP", "SCurrncy", "SetupCurrency", 25)
        objUDFEngine.AddFloatField("@PSSIT_OMGP", "Setrate", "SetupRate/Hour", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddAlphaField("@PSSIT_OMGP", "SAccode", "AccountCode", 30)
        objUDFEngine.AddAlphaField("@PSSIT_OMGP", "SAcname", "AccountName", 100)
        objUDFEngine.AddAlphaField("@PSSIT_OMGP", "Adnl1", "OtherInfo1", 50)
        objUDFEngine.AddAlphaField("@PSSIT_OMGP", "Adnl2", "OtherInfo2", 50)
        objUDFEngine.AddAlphaField("@PSSIT_OMGP", "Remarks", "Remarks", 100)
        objUDFEngine.addField("@PSSIT_OMGP", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")
        objUDFEngine.AddAlphaField("@PSSIT_OMGP", "RActAcCode", "RunningAccountCode", 30)
        objUDFEngine.AddAlphaField("@PSSIT_OMGP", "SActAcCode", "SetupAccountCode", 30)


        objUDFEngine.CreateTable("PSSIT_PMWCHDR", "MachineMaster", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "wcno", "WorkCenterNo", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "wcname", "WorkCenterName", 40)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "wcshname", "ShortName", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "typecode", "TypeCode", 60)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "typedesc", "TypeName", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "modecode", "ModelCode", 60)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "modedesc", "ModelName", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "makecode", "MakeCode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "makedesc", "MakeName", 60)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "mfserial", "ManufacturerSerialNo", 20)
        objUDFEngine.AddNumericField("@PSSIT_PMWCHDR", "yearmake", "YearofMake", 8)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "deptcode", "DepartmentCode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "deptdesc", "DepartmentName", 60)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "isgrp", "IsGroup", 3)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "undergrp", "UnderGroup", 40)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "MGcode", "MachineGroupCode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "MGname", "MachineGroupName", 50)
        objUDFEngine.AddNumericField("@PSSIT_PMWCHDR", "ohrsday", "OperatingHours/Days", 4)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "uomcode", "MeasurementUnit", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "uomdesc", "UomName", 60)
        objUDFEngine.AddNumericField("@PSSIT_PMWCHDR", "wccapa", "WorkCenterCapacity", 10)
        objUDFEngine.AddNumericField("@PSSIT_PMWCHDR", "spacelen", "SpaceLength", 10)
        objUDFEngine.AddNumericField("@PSSIT_PMWCHDR", "spacebre", "SpaceBreath", 10)
        objUDFEngine.AddDateField("@PSSIT_PMWCHDR", "insdate", "InstallationDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddNumericField("@PSSIT_PMWCHDR", "inskw", "InstallationKW", 10)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "bpcode", "BusinessPartnerCode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "pono", "PurchaseOrderNumber", 20)
        objUDFEngine.AddDateField("@PSSIT_PMWCHDR", "podate", "Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddDateField("@PSSIT_PMWCHDR", "wardate", "WarrentyExpires", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "status", "Status", 10)
        objUDFEngine.AddFloatField("@PSSIT_PMWCHDR", "opercost", "OperationCost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_PMWCHDR", "powecost", "PowerCost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_PMWCHDR", "Setupcost", "SetUpCost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_PMWCHDR", "cost1", "Cost1", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_PMWCHDR", "cost2", "Cost2", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "Accode", "AccountCode", 30)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "Acname", "AccountName", 100)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "SAccode", "SetupAccountCode", 30)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "SAcname", "SetupAccountName", 100)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "Adnl1", "OtherInfo1", 30)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "Adnl2", "OtherInfo2", 30)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "Remarks", "Remarks", 100)
        objUDFEngine.addField("@PSSIT_PMWCHDR", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "ActAcCode", "ActualAccountCode", 30)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCHDR", "SActAcCode", "SetupAccountCode", 30)

        objUDFEngine.CreateTable("PSSIT_PMWCPARA", "MachineMaster-ProdParameter", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCPARA", "paracode", "ParameterCode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCPARA", "Paradesc", "ParameterName", 60)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCPARA", "paraval", "Parametervalue", 60)

        objUDFEngine.CreateTable("PSSIT_PMWCITEM", "MachineMaster-CriticalItems", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCITEM", "itemcode", "ItemCode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCITEM", "itemdesc", "ItemDescription", 60)
        objUDFEngine.AddDateField("@PSSIT_PMWCITEM", "insdate", "InstallationDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddNumericField("@PSSIT_PMWCITEM", "lifeday", "LifeSpaninDays", 10)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCITEM", "Units", "UomUnits", 20)
        objUDFEngine.AddFloatField("@PSSIT_PMWCITEM", "qty", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCITEM", "Adnl1", "OtherInfo1", 50)

        objUDFEngine.CreateTable("PSSIT_PMWCSFT", "MachineMaster-ShiftDetails", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCSFT", "SCode", "ShiftCode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCSFT", "Sdescr", "ShiftDescription", 30)
        objUDFEngine.AddNumericField("@PSSIT_PMWCSFT", "Duratmin", "Duration(Mins)", 10)

        objUDFEngine.CreateTable("PSSIT_PMWCSPEC", "MachineMaster-Rows", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCSPEC", "speccode", "SpecificationCode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCSPEC", "specval", "SpecificationValue", 20)

        objUDFEngine.CreateTable("PSSIT_ORTE", "OperationRouting", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objUDFEngine.addField("@PSSIT_ORTE", "Defrte", "DefaultRoute", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")
        objUDFEngine.AddAlphaField("@PSSIT_ORTE", "Itemcode", "ItemCode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_ORTE", "Itemname", "ItemDescription", 100)
        objUDFEngine.AddAlphaField("@PSSIT_ORTE", "drgno", "DrawingNo.", 50)
        objUDFEngine.AddAlphaField("@PSSIT_ORTE", "Revno", "RevisionNo.", 50)
        objUDFEngine.AddDateField("@PSSIT_ORTE", "Revdt", "RevisionDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddAlphaField("@PSSIT_ORTE", "Adnl1", "OtherInfo1", 50)
        objUDFEngine.AddAlphaField("@PSSIT_ORTE", "Adnl2", "OtherInfo2", 50)
        objUDFEngine.AddAlphaField("@PSSIT_ORTE", "Remarks", "Remarks", 50)
        objUDFEngine.addField("@PSSIT_ORTE", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")


        objUDFEngine.CreateTable("PSSIT_RTE1", "OperationsRouting-Machines", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        objUDFEngine.AddAlphaField("@PSSIT_RTE1", "wcno", "Machine", 20)
        objUDFEngine.AddAlphaField("@PSSIT_RTE1", "wcname", "MachineName", 40)
        objUDFEngine.AddAlphaField("@PSSIT_RTE1", "MGname", "GroupName", 50)
        objUDFEngine.AddNumericField("@PSSIT_RTE1", "Setime", "SetupTime(Mins)", 5)
        objUDFEngine.AddNumericField("@PSSIT_RTE1", "Opertime", "OperationTime(Mins)", 5)
        objUDFEngine.AddNumericField("@PSSIT_RTE1", "Othetime1", "OtherTime1", 5)
        objUDFEngine.AddNumericField("@PSSIT_RTE1", "Othetime2", "OtherTime2", 5)
        objUDFEngine.AddFloatField("@PSSIT_RTE1", "perqty", "PerQty.", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddAlphaField("@PSSIT_RTE1", "Adnl1", "OtherInfo1", 50)
        objUDFEngine.AddAlphaField("@PSSIT_RTE1", "Adnl2", "OtherInfo2", 50)
        objUDFEngine.AddAlphaField("@PSSIT_RTE1", "Adnl3", "OtherInfo3", 10)
        objUDFEngine.AddAlphaField("@PSSIT_RTE1", "Adnl4", "OtherInfo4", 50)
        objUDFEngine.AddAlphaField("@PSSIT_RTE1", "Rteid", "RouteID", 20)
        objUDFEngine.AddNumericField("@PSSIT_RTE1", "Seqnce", "Sequence", 10)
        objUDFEngine.AddAlphaField("@PSSIT_RTE1", "OprCode", "OperationCode", 20)


        objUDFEngine.CreateTable("PSSIT_RTE2", "OperationRouting-Labour", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        objUDFEngine.AddAlphaField("@PSSIT_RTE2", "Skilgrp", "SkillGroup", 20)
        objUDFEngine.AddAlphaField("@PSSIT_RTE2", "LGname", "Description", 50)
        objUDFEngine.AddNumericField("@PSSIT_RTE2", "Reqtime", "RequiredTime", 5)
        objUDFEngine.AddNumericField("@PSSIT_RTE2", "Reqno", "RequiredNo.", 3)
        objUDFEngine.AddAlphaField("@PSSIT_RTE2", "Adnl1", "OtherInfo1", 50)
        objUDFEngine.AddNumericField("@PSSIT_RTE2", "Adnl2", "OtherInfo2", 10)
        objUDFEngine.AddAlphaField("@PSSIT_RTE2", "Rteid", "RouteID", 20)
        objUDFEngine.AddNumericField("@PSSIT_RTE2", "Seqnce", "Sequence", 10)
        objUDFEngine.AddAlphaField("@PSSIT_RTE2", "wcno", "Machine", 20)
        objUDFEngine.AddAlphaField("@PSSIT_RTE2", "OprCode", "OperationCode", 20)


        objUDFEngine.CreateTable("PSSIT_RTE3", "OperationRouting-Tools", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        objUDFEngine.AddAlphaField("@PSSIT_RTE3", "Toolcode", "ToolCode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_RTE3", "TLname", "ToolDescription", 25)
        objUDFEngine.AddNumericField("@PSSIT_RTE3", "Strokes", "No.ofStrokes", 10)
        objUDFEngine.AddAlphaField("@PSSIT_RTE3", "Adnl1", "OtherInfo1", 50)
        objUDFEngine.AddNumericField("@PSSIT_RTE3", "Adnl2", "OtherInfo2", 10)
        objUDFEngine.AddAlphaField("@PSSIT_RTE3", "Rteid", "RouteID", 20)
        objUDFEngine.AddNumericField("@PSSIT_RTE3", "Seqnce", "Sequence", 10)
        objUDFEngine.AddAlphaField("@PSSIT_RTE3", "wcno", "Machine", 20)
        objUDFEngine.AddAlphaField("@PSSIT_RTE3", "OprCode", "OperationCode", 20)

        objUDFEngine.CreateTable("PSSIT_OPRN", "Operations", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objUDFEngine.AddAlphaField("@PSSIT_OPRN", "Oprname", "OperationName", 100)
        objUDFEngine.AddAlphaField("@PSSIT_OPRN", "Oprtype", "OperationType", 50)
        objUDFEngine.addField("@PSSIT_OPRN", "Rework", "Rework", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")
        objUDFEngine.AddAlphaField("@PSSIT_OPRN", "Accode", "ReworkAccount", 30)
        objUDFEngine.AddAlphaField("@PSSIT_OPRN", "Acname", "ReworkAccountDescription", 100)
        objUDFEngine.AddAlphaField("@PSSIT_OPRN", "Adnl1", "OtherInfo1", 50)
        objUDFEngine.AddAlphaField("@PSSIT_OPRN", "Adnl2", "OtherInfo2", 50)
        objUDFEngine.AddAlphaField("@PSSIT_OPRN", "Remarks", "Remarks", 50)
        objUDFEngine.addField("@PSSIT_OPRN", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")
        objUDFEngine.AddAlphaField("@PSSIT_OPRN", "ActAcCode", "ActualAccountCode", 30)

        objUDFEngine.CreateTable("PSSIT_PRN2", "Operations-SkillGroup-Rows", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objUDFEngine.AddAlphaField("@PSSIT_PRN2", "Skilgrp", "SkillGroup", 50)
        objUDFEngine.AddAlphaField("@PSSIT_PRN2", "LGname", "Description", 50)
        objUDFEngine.AddNumericField("@PSSIT_PRN2", "Reqno", "RequiredNo.", 3)


        objUDFEngine.CreateTable("PSSIT_PRN3", "Operations-Tools", SAPbobsCOM.BoUTBTableType.bott_NoObject)

        objUDFEngine.AddAlphaField("@PSSIT_PRN3", "Toolcode", "ToolCode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PRN3", "TLname", "ToolDescription", 25)
        objUDFEngine.AddAlphaField("@PSSIT_PRN3", "Oprcode", "Operation", 50)
        objUDFEngine.AddAlphaField("@PSSIT_PRN3", "wcno", "Machine", 20)


        objUDFEngine.CreateTable("PSSIT_PRN1", "Operations-SkillGroup-Rows", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)

        objUDFEngine.AddAlphaField("@PSSIT_PRN1", "wcno", "MachineCode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PRN1", "wcname", "MachineName", 40)
        objUDFEngine.AddAlphaField("@PSSIT_PRN1", "MGname", "GroupName", 50)


        objUDFEngine.CreateTable("PSSIT_OPEY", "ProductionEntry", SAPbobsCOM.BoUTBTableType.bott_Document)
        objUDFEngine.AddAlphaField("@PSSIT_OPEY", "Pnordno", "ProductionOrderNo", 20)
        objUDFEngine.AddNumericField("@PSSIT_OPEY", "WORNo", "WorkOrderNo", 10)
        objUDFEngine.AddNumericField("@PSSIT_OPEY", "Pnordser", "ProductionOrderSeries", 6)
        objUDFEngine.AddDateField("@PSSIT_OPEY", "Pordt", "ProductionOrderDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddDateField("@PSSIT_OPEY", "Docdt", "DocumentDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddAlphaField("@PSSIT_OPEY", "Scode", "ShiftCode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_OPEY", "Sdesc", "ShiftDescription", 30)
        objUDFEngine.AddAlphaField("@PSSIT_OPEY", "Itemcode", "ProductCode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_OPEY", "ItemName", "Description", 100)
        objUDFEngine.AddAlphaField("@PSSIT_OPEY", "GLMethod", "GLMethod", 1)
        objUDFEngine.AddAlphaField("@PSSIT_OPEY", "Whscode", "Warehouse", 8)
        objUDFEngine.AddAlphaField("@PSSIT_OPEY", "Whsname", "WarehouseDescription", 100)
        objUDFEngine.AddFloatField("@PSSIT_OPEY", "Planqty", "PlannedQty.", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddFloatField("@PSSIT_OPEY", "Comqty", "CompletedQty.", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objUDFEngine.AddFloatField("@PSSIT_OPEY", "Rewqty", "ReworkQty.", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddFloatField("@PSSIT_OPEY", "Scpqty", "ScrapQty.", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objUDFEngine.AddFloatField("@PSSIT_OPEY", "RejQty", "RejectedQty.", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddAlphaField("@PSSIT_OPEY", "Rteid", "RouteID", 20)
        objUDFEngine.AddNumericField("@PSSIT_OPEY", "Oplnid", "Processlineno", 2)
        objUDFEngine.AddAlphaField("@PSSIT_OPEY", "Oprcode", "Operation", 20)
        objUDFEngine.AddAlphaField("@PSSIT_OPEY", "Oprname", "OperationName", 100)
        objUDFEngine.AddFloatField("@PSSIT_OPEY", "ProdQty", "ProducedQty.", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddFloatField("@PSSIT_OPEY", "Passqty", "PassedQty.", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        'Commented by Manimaran----s
        'objUDFEngine.AddFloatField("@PSSIT_OPEY", "Rewrkqty", "ReworkQty.", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        'objUDFEngine.AddFloatField("@PSSIT_OPEY", "scrapqty", "ScrapQty.", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        'Commented by Manimaran----e
        objUDFEngine.addField("@PSSIT_OPEY", "Rework", "Rework", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")
        'Commented by Manimaran----s
        'objUDFEngine.AddAlphaField("@PSSIT_OPEY", "Rerewk", "ReasonforRe-Work", 50)
        'objUDFEngine.AddAlphaField("@PSSIT_OPEY", "Rescrp", "ReasonforScrap", 50)
        objUDFEngine.addField("@PSSIT_OPEY", "InTime", "Interval time", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")
        'Commented by Manimaran----e
        objUDFEngine.addField("@PSSIT_OPEY", "Acckey", "AccountKey", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")
        objUDFEngine.AddAlphaField("@PSSIT_OPEY", "Raccode", "ReworkAccountCode", 30)
        objUDFEngine.AddAlphaField("@PSSIT_OPEY", "Racname", "ReworkAccountName", 100)
        objUDFEngine.AddAlphaField("@PSSIT_OPEY", "Saccode", "ScrapAccountCode", 30)
        objUDFEngine.AddAlphaField("@PSSIT_OPEY", "Sacname", "ScrapAccountName", 100)
        objUDFEngine.AddFloatField("@PSSIT_OPEY", "Totmcst", "TotalMachineCost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_OPEY", "Tottcst", "TotalToolCost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_OPEY", "Totlcst", "TotalLabourCost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddNumericField("@PSSIT_OPEY", "Jvno", "JournalNo.", 10)
        objUDFEngine.AddAlphaField("@PSSIT_OPEY", "Adnl1", "OtherInfo1", 50)
        objUDFEngine.AddAlphaField("@PSSIT_OPEY", "Adnl2", "OtherInfo2", 50)
        objUDFEngine.AddAlphaField("@PSSIT_OPEY", "Remarks", "Remarks", 50)
        objUDFEngine.addField("@PSSIT_OPEY", "Closekey", "OperationClosedKey", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")
        objUDFEngine.AddFloatField("@PSSIT_OPEY", "AccPassQty", "AccumulatedPassQty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddFloatField("@PSSIT_OPEY", "AccProdQty", "AccumulatedProducedQty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddFloatField("@PSSIT_OPEY", "AccRewQty", "AccumulatedReworkQty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddFloatField("@PSSIT_OPEY", "AccScrapQty", "AccumulatedScrapQty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        'Added by Manimaran-----s
        objUDFEngine.AddAlphaField("@PSSIT_OPEY", "StpRea", "Stoppage Reson", 250)
        'Added by Manimaran-----e


        objUDFEngine.CreateTable("PSSIT_PEY2", "ProductionEntryRows-Labour", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUDFEngine.AddAlphaField("@PSSIT_PEY2", "Lrcode", "LabourCode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PEY2", "LGCode", "SkillGroup", 25)
        objUDFEngine.AddAlphaField("@PSSIT_PEY2", "LGname", "SkillGroupName", 50)
        objUDFEngine.AddNumericField("@PSSIT_PEY2", "Reqno", "RequiredNo.", 3)
        objUDFEngine.addField("@PSSIT_PEY2", "Labkey", "LabourKey", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")
        objUDFEngine.AddAlphaField("@PSSIT_PEY2", "Parallel", "Parallel", 20)
        objUDFEngine.AddDateField("@PSSIT_PEY2", "Frtime", "FromTime", SAPbobsCOM.BoFldSubTypes.st_Time)
        objUDFEngine.AddDateField("@PSSIT_PEY2", "Totime", "ToTime", SAPbobsCOM.BoFldSubTypes.st_Time)
        'Added by Manimaran-----s
        objUDFEngine.AddNumericField("@PSSIT_PEY2", "OTtime", "OT Time", 5)
        objUDFEngine.AddNumericField("@PSSIT_PEY2", "Nop", "No of persons", 5)
        'Added by MAnimaran-----e
        objUDFEngine.AddNumericField("@PSSIT_PEY2", "Wrktime", "WorkedTime", 5)
        objUDFEngine.AddFloatField("@PSSIT_PEY2", "Qty", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.addField("@PSSIT_PEY2", "Acckey", "AccountKey", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")
        objUDFEngine.AddAlphaField("@PSSIT_PEY2", "Accode", "AccountCode", 30)
        objUDFEngine.AddAlphaField("@PSSIT_PEY2", "Acname", "AccountName", 100)
        objUDFEngine.AddAlphaField("@PSSIT_PEY2", "CAccode", "ContraAccountCode", 30)
        objUDFEngine.AddAlphaField("@PSSIT_PEY2", "CAcname", "ContraAccountName", 100)
        objUDFEngine.AddFloatField("@PSSIT_PEY2", "Lrtph", "LabourRate/Hour", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_PEY2", "Totcost", "TotalRunningLabourCost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddAlphaField("@PSSIT_PEY2", "Adnl1", "OtherInfo1", 50)
        objUDFEngine.AddAlphaField("@PSSIT_PEY2", "Adnl2", "OtherInfo2", 50)
        objUDFEngine.AddAlphaField("@PSSIT_PEY2", "Prdentno", "ProductionEntryNo.", 30)
        objUDFEngine.AddNumericField("@PSSIT_PEY2", "Maclid", "MachineLineId", 10)
        objUDFEngine.AddNumericField("@PSSIT_PEY2", "Madcey", "Machinedocentry", 10)
        objUDFEngine.AddAlphaField("@PSSIT_PEY2", "wcno", "Machine", 20)

        objUDFEngine.CreateTable("PSSIT_PEY3", "ProductionEntryRows-Tools", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUDFEngine.AddAlphaField("@PSSIT_PEY3", "Toolcode", "ToolCode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PEY3", "TLname", "ToolDescription", 25)
        objUDFEngine.AddFloatField("@PSSIT_PEY3", "Qty", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.addField("@PSSIT_PEY3", "Acckey", "AccountKey", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")
        objUDFEngine.AddAlphaField("@PSSIT_PEY3", "Accode", "AccountCode", 30)
        objUDFEngine.AddAlphaField("@PSSIT_PEY3", "Acname", "AccountName", 100)
        objUDFEngine.AddAlphaField("@PSSIT_PEY3", "CAccode", "ContraAccountCode", 30)
        objUDFEngine.AddAlphaField("@PSSIT_PEY3", "CAcname", "ContraAccountName", 100)
        objUDFEngine.AddFloatField("@PSSIT_PEY3", "Tlctppie", "ToolCostperpiece", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_PEY3", "Totcost", "TotalRunningToolCost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddAlphaField("@PSSIT_PEY3", "Adnl1", "OtherInfo1", 50)
        objUDFEngine.AddAlphaField("@PSSIT_PEY3", "Adnl2", "OtherInfo2", 50)
        objUDFEngine.AddAlphaField("@PSSIT_PEY3", "Prdentno", "ProductionEntryNo.", 30)
        objUDFEngine.AddNumericField("@PSSIT_PEY3", "Maclid", "MachineLineId", 10)
        objUDFEngine.AddNumericField("@PSSIT_PEY3", "Madcey", "Machinedocentry", 10)
        objUDFEngine.AddAlphaField("@PSSIT_PEY3", "wcno", "Machine", 20)


        objUDFEngine.CreateTable("PSSIT_PEY4", "ProductionEntry-FixedCost", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUDFEngine.AddAlphaField("@PSSIT_PEY4", "Fcost", "FixedCost", 30)
        objUDFEngine.AddFloatField("@PSSIT_PEY4", "UnitCost", "UnitCost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddAlphaField("@PSSIT_PEY4", "Absmthd", "AbsorptionMethod", 60)
        objUDFEngine.AddAlphaField("@PSSIT_PEY4", "Accode", "AccountCode", 30)
        objUDFEngine.AddAlphaField("@PSSIT_PEY4", "Acname", "AccountName", 100)
        objUDFEngine.AddFloatField("@PSSIT_PEY4", "Totfcst", "TotalFixedCost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddNumericField("@PSSIT_PEY4", "Pordser", "ProductionOrderSeries", 6)
        objUDFEngine.AddNumericField("@PSSIT_PEY4", "Pordno", "ProductionOrderNo.", 10)
        objUDFEngine.AddAlphaField("@PSSIT_PEY4", "wcno", "Machine", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PEY4", "Wrkno", "WorkCentre", 20)
        objUDFEngine.AddNumericField("@PSSIT_PEY4", "Prdentno", "ProductionEntryNo.", 10)


        objUDFEngine.CreateTable("PSSIT_PEY1", "ProductionEntry-Rows", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUDFEngine.AddAlphaField("@PSSIT_PEY1", "wcno", "Machine", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PEY1", "wcname", "MachineName", 40)
        objUDFEngine.AddAlphaField("@PSSIT_PEY1", "Wrkno", "WorkCentre", 20)
        objUDFEngine.AddAlphaField("@PSSIT_PEY1", "Wrkname", "WorkCentreName", 50)
        objUDFEngine.AddAlphaField("@PSSIT_PEY1", "type", "Type", 50)
        'Added by Manimaran------s
        objUDFEngine.AddAlphaField("@PSSIT_PEY1", "sotime", "Setup/Operation Time", 10)

        'Added by Manimaran------e
        objUDFEngine.AddDateField("@PSSIT_PEY1", "Frtime", "FromTime", SAPbobsCOM.BoFldSubTypes.st_Time)
        objUDFEngine.AddDateField("@PSSIT_PEY1", "Totime", "ToTime", SAPbobsCOM.BoFldSubTypes.st_Time)
        objUDFEngine.AddNumericField("@PSSIT_PEY1", "Rntime", "RunTime", 5)
        objUDFEngine.AddFloatField("@PSSIT_PEY1", "Qty", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.addField("@PSSIT_PEY1", "Acckey", "AccountKey", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")
        objUDFEngine.AddAlphaField("@PSSIT_PEY1", "Accode", "AccountCode", 30)
        objUDFEngine.AddAlphaField("@PSSIT_PEY1", "Acname", "AccountName", 100)
        objUDFEngine.AddAlphaField("@PSSIT_PEY1", "CAccode", "ContraAccountCode", 30)
        objUDFEngine.AddAlphaField("@PSSIT_PEY1", "CAcname", "ContraAccountName", 100)
        objUDFEngine.AddFloatField("@PSSIT_PEY1", "Mopcph", "MachineOperationCost/Hour", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_PEY1", "Mprcph", "MachinePowerCost/Hour", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_PEY1", "Mohcph1", "MachineOthercost1/Hour", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_PEY1", "Mohcph2", "MachineOthercost2/Hour", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_PEY1", "RMopcph", "RunningMac.Oper.Cost/Hour", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_PEY1", "RMprcph", "RunningMac.PowerCost/Hour", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_PEY1", "RMohcph1", "RunningMac.Othcost1/Hour", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_PEY1", "RMohcph2", "RunningMac.Othcost2/Hour", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_PEY1", "Totcost", "TotalRunningMac.Cost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddAlphaField("@PSSIT_PEY1", "Adnl1", "OtherInfo1", 50)
        objUDFEngine.AddAlphaField("@PSSIT_PEY1", "Adnl2", "OtherInfo2", 50)

        'Added by Manimaran----s
        objUDFEngine.CreateTable("PSSIT_PEY5", "ProductionEntry-Rework", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUDFEngine.AddFloatField("@PSSIT_PEY5", "Rewrkqty", "ReworkQty.", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddAlphaField("@PSSIT_PEY5", "Rerewk", "ReasonforRe-Work", 50)

        objUDFEngine.CreateTable("PSSIT_PEY6", "ProductionEntry-Scrap", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUDFEngine.AddFloatField("@PSSIT_PEY6", "scrapqty", "ScrapQty.", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddAlphaField("@PSSIT_PEY6", "Rescrp", "ReasonforScrap", 50)
        objUDFEngine.AddAlphaField("@PSSIT_PEY6", "MacCode", "Machine Code", 50)
        objUDFEngine.AddAlphaField("@PSSIT_PEY6", "LabCode", "Labour Code", 50)
        'Added by Manimaran----s

        objUDFEngine.CreateTable("PSSIT_OCON", "ProductionSetup", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        objUDFEngine.addField("@PSSIT_OCON", "Acckey", "AccountPosting", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "N")
        objUDFEngine.addField("@PSSIT_OCON", "Fcman", "FixedCostMandatory", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "N")
        objUDFEngine.addField("@PSSIT_OCON", "Labman", "LabourMandatory", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "N")
        'Added by Manimaran----------s
        objUDFEngine.AddNumericField("@PSSIT_OCON", "SOHDPer", "Sourcing Percentage", 10)
        objUDFEngine.AddNumericField("@PSSIT_OCON", "POHDPer", "Production Percentage", 10)
        'Added by Manimaran----------e

        objUDFEngine.CreateTable("PSSIT_ORES", "Reason", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objUDFEngine.AddAlphaField("@PSSIT_ORES", "wcno", "Machine", 100)

        objUDFEngine.CreateTable("PSSIT_PMREMEDIES", "WCRemedies", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objUDFEngine.AddAlphaField("@PSSIT_PMREMEDIES", "remedesc", "RemediesDesc", 60)

        objUDFEngine.CreateTable("PSSIT_OSFT", "Shift", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objUDFEngine.AddAlphaField("@PSSIT_OSFT", "Sdescr", "ShiftDescription", 30)
        objUDFEngine.AddDateField("@PSSIT_OSFT", "Sftime", "FromTime", SAPbobsCOM.BoFldSubTypes.st_Time)
        objUDFEngine.AddDateField("@PSSIT_OSFT", "Sttime", "ToTime", SAPbobsCOM.BoFldSubTypes.st_Time)
        '' ''objUDFEngine.AddNumericField("@PSSIT_OSFT", "Sbreak", "Break(Mins)", 3)
        '' ''objUDFEngine.AddNumericField("@PSSIT_OSFT", "Duratmin", "Duration(Mins)", 5)
        '' ''objUDFEngine.AddNumericField("@PSSIT_OSFT", "Durathrs", "Duration(Hrs)", 5)
        objUDFEngine.AddDateField("@PSSIT_OSFT", "Sbreak", "Break(Mins)", SAPbobsCOM.BoFldSubTypes.st_Time)
        objUDFEngine.AddDateField("@PSSIT_OSFT", "Duratmin", "Duration(Mins)", SAPbobsCOM.BoFldSubTypes.st_Time)
        objUDFEngine.AddDateField("@PSSIT_OSFT", "Durathrs", "Duration(Hrs)", SAPbobsCOM.BoFldSubTypes.st_Time)

        objUDFEngine.AddAlphaField("@PSSIT_OSFT", "Adnl1", "OtherInfo1", 30)
        objUDFEngine.AddAlphaField("@PSSIT_OSFT", "Adnl2", "OtherInfo2", 30)
        objUDFEngine.AddAlphaField("@PSSIT_OSFT", "Remarks", "Remarks", 200)
        objUDFEngine.addField("@PSSIT_OSFT", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")


        objUDFEngine.CreateTable("PSSIT_OLGP", "SkillGroups", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objUDFEngine.AddAlphaField("@PSSIT_OLGP", "LGname", "Description", 50)
        objUDFEngine.AddAlphaField("@PSSIT_OLGP", "WCcode", "WorkCentre", 25)
        objUDFEngine.AddAlphaField("@PSSIT_OLGP", "WCName", "WorkCentreName", 50)
        objUDFEngine.AddAlphaField("@PSSIT_OLGP", "Currncy", "Currency", 25)
        objUDFEngine.AddFloatField("@PSSIT_OLGP", "Labrate", "LabourRate/Hour", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddAlphaField("@PSSIT_OLGP", "Accode", "AccountCode", 30)
        objUDFEngine.AddAlphaField("@PSSIT_OLGP", "Acname", "AccountDescription", 100)
        objUDFEngine.AddAlphaField("@PSSIT_OLGP", "Adnl1", "OtherInfo1", 50)
        objUDFEngine.AddAlphaField("@PSSIT_OLGP", "Adnl2", "OtherInfo2", 50)
        objUDFEngine.AddAlphaField("@PSSIT_OLGP", "Remarks", "Remarks", 100)
        objUDFEngine.addField("@PSSIT_OLGP", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")
        objUDFEngine.AddAlphaField("@PSSIT_OLGP", "ActAcCode", "ActualAccountCode", 30)


        objUDFEngine.CreateTable("PSSIT_OSCY", "Category", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objUDFEngine.AddAlphaField("@PSSIT_OSCY", "Catname", "CategoryDescription", 30)
        objUDFEngine.AddAlphaField("@PSSIT_OSCY", "Remarks", "Remarks", 50)



        objUDFEngine.CreateTable("PSSIT_OSGE", "Stoppage", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objUDFEngine.AddAlphaField("@PSSIT_OSGE", "Stopname", "StoppageDescription", 30)
        objUDFEngine.AddAlphaField("@PSSIT_OSGE", "Catcode", "CategoryCode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_OSGE", "Catname", "CategoryDescription", 30)
        objUDFEngine.AddNumericField("@PSSIT_OSGE", "Plantime", "PlannedTime(Mins)", 3)
        objUDFEngine.AddAlphaField("@PSSIT_OSGE", "Adnl1", "OtherInfo1", 50)
        objUDFEngine.AddAlphaField("@PSSIT_OSGE", "Adnl2", "OtherInfo2", 50)
        objUDFEngine.AddAlphaField("@PSSIT_OSGE", "Remarks", "Remarks", 100)
        objUDFEngine.addField("@PSSIT_OSGE", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")



        objUDFEngine.CreateTable("PSSIT_OTLS", "Tools", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objUDFEngine.AddAlphaField("@PSSIT_OTLS", "TLname", "ToolDescription", 60)
        objUDFEngine.AddAlphaField("@PSSIT_OTLS", "Itemcode", "ItemCode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_OTLS", "Itemname", "ItemDescription", 100)
        objUDFEngine.AddAlphaField("@PSSIT_OTLS", "WCcode", "WorkCentre", 20)
        objUDFEngine.AddAlphaField("@PSSIT_OTLS", "WCname", "WorkCentreName", 50)
        objUDFEngine.AddDateField("@PSSIT_OTLS", "Purdate", "DateofPurchase", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddFloatField("@PSSIT_OTLS", "Lcost", "LandedCost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddNumericField("@PSSIT_OTLS", "Enou", "Expectedstrokes", 10)
        objUDFEngine.AddNumericField("@PSSIT_OTLS", "Cnou", "Completedstrokes", 10)
        objUDFEngine.AddNumericField("@PSSIT_OTLS", "Tstime", "ToolSettingTime", 3)
        objUDFEngine.AddFloatField("@PSSIT_OTLS", "Cpno", "Cost/Stroke", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddAlphaField("@PSSIT_OTLS", "Accode", "AccountCode", 20)
        objUDFEngine.AddAlphaField("@PSSIT_OTLS", "Acname", "AccountDescription", 100)
        objUDFEngine.addField("@PSSIT_OTLS", "Recond", "Reconditioned", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")
        objUDFEngine.AddAlphaField("@PSSIT_OTLS", "Partool", "ParentTool", 50)
        objUDFEngine.AddAlphaField("@PSSIT_OTLS", "Adnl1", "OtherInfo1", 50)
        objUDFEngine.AddAlphaField("@PSSIT_OTLS", "Adnl2", "OtherInfo2", 50)
        objUDFEngine.AddAlphaField("@PSSIT_OTLS", "Techspec", "TechnicalSpecifications", 100)
        objUDFEngine.AddAlphaField("@PSSIT_OTLS", "Remarks", "Remarks", 100)
        objUDFEngine.addField("@PSSIT_OTLS", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")
        objUDFEngine.AddAlphaField("@PSSIT_OTLS", "ActAcCode", "ActualAccountCode", 30)
        'Added by Manimaran-----s
        objUDFEngine.AddNumericField("@PSSIT_OTLS", "TypOfItm", "Types of items", 3)
        'Added by Manimaran-----e


        objUDFEngine.CreateTable("PSSIT_PMWCUOM", "WorkCentreUOM", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCUOM", "uomdesc", "UOMName", 60)



        objUDFEngine.CreateTable("PSSIT_OCST", "WorkCentreUOM", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objUDFEngine.AddAlphaField("@PSSIT_OCST", "Wcstyp", "Description", 50)


        objUDFEngine.CreateTable("PSSIT_OWCR", "WorkCentreUOM", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objUDFEngine.AddAlphaField("@PSSIT_OWCR", "WCname", "Description", 50)
        objUDFEngine.AddAlphaField("@PSSIT_OWCR", "WCtype", "WorkCentreType", 50)
        objUDFEngine.AddAlphaField("@PSSIT_OWCR", "Adnl1", "OtherInfo1", 30)
        objUDFEngine.AddAlphaField("@PSSIT_OWCR", "Adnl2", "OtherInfo2", 30)
        objUDFEngine.AddAlphaField("@PSSIT_OWCR", "Remarks", "Remarks", 100)
        objUDFEngine.addField("@PSSIT_OWCR", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")


        objUDFEngine.CreateTable("PSSIT_PMWCMAKE", "WorkCentreMake", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCMAKE", "makedesc", "MakeName", 60)


        objUDFEngine.CreateTable("PSSIT_PMWCMODEL", "WorkCentreModel", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objUDFEngine.AddAlphaField("@PSSIT_PMWCMODEL", "modedesc", "ModelName", 60)



        objUDFEngine.CreateTable("PSSIT_WCR1", "WorkCentreRows", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objUDFEngine.AddAlphaField("@PSSIT_WCR1", "Fcost", "FixedCost", 30)
        objUDFEngine.AddAlphaField("@PSSIT_WCR1", "Currency", "Currency", 25)
        objUDFEngine.AddFloatField("@PSSIT_WCR1", "UnitCost", "UnitCost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddAlphaField("@PSSIT_WCR1", "Absmthd", "AbsorptionMethod", 60)
        objUDFEngine.AddAlphaField("@PSSIT_WCR1", "Accode", "AccountCode", 30)
        objUDFEngine.AddAlphaField("@PSSIT_WCR1", "Acname", "AccountName", 100)
        objUDFEngine.AddAlphaField("@PSSIT_WCR1", "Adnl1", "OtherInfo", 50)
        objUDFEngine.AddAlphaField("@PSSIT_WCR1", "ActAcCode", "ActualAccountCode", 30)


        objUDFEngine.CreateTable("PSSIT_OTYP", "Work Centre Type", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objUDFEngine.AddAlphaField("@PSSIT_OTYP", "Wctypnam", "Description", 50)


        objUDFEngine.CreateTable("PSSIT_WOR3", "WorkOrder-Cost", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        objUDFEngine.AddNumericField("@PSSIT_WOR3", "Pordser", "ProductionOrderSeries", 6)
        objUDFEngine.AddNumericField("@PSSIT_WOR3", "Pordno", "ProductionOrderNo.", 10)
        objUDFEngine.AddFloatField("@PSSIT_WOR3", "Totcmpcst", "ComponentCost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_WOR3", "Totlbrcst", "TotalLabourCost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_WOR3", "Totmccst", "TotalMachineCost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_WOR3", "Tottoolcst", "TotalToolCost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_WOR3", "Totsubcst", "TotalS.C.Cost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_WOR3", "Totcst", "TotalCost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_WOR3", "Adnl1", "OtherCost1", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_WOR3", "Adnl2", "OtherCost2", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_WOR3", "Adnl3", "OtherCost3", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_WOR3", "Adnl4", "OtherCost4", SAPbobsCOM.BoFldSubTypes.st_Price)


        objUDFEngine.CreateTable("PSSIT_WOR4", "WorkOrder-CostDetails", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        objUDFEngine.AddNumericField("@PSSIT_WOR4", "Pordser", "ProductionOrderSeries", 6)
        objUDFEngine.AddNumericField("@PSSIT_WOR4", "Pordno", "ProductionOrderNo.", 10)
        objUDFEngine.AddNumericField("@PSSIT_WOR4", "DocEntry", "DocEntry", 10)
        objUDFEngine.AddNumericField("@PSSIT_WOR4", "Lineid", "LineId", 10)
        objUDFEngine.AddAlphaField("@PSSIT_WOR4", "Fcost", "FixedCost", 30)
        objUDFEngine.AddFloatField("@PSSIT_WOR4", "UnitCost", "UnitCost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddAlphaField("@PSSIT_WOR4", "Absmthd", "AbsorptionMethod", 60)
        objUDFEngine.AddAlphaField("@PSSIT_WOR4", "Accode", "AccountCode", 30)
        objUDFEngine.AddAlphaField("@PSSIT_WOR4", "Acname", "AccountName", 100)
        objUDFEngine.AddFloatField("@PSSIT_WOR4", "Totfcst", "TotalFixedCost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_WOR4", "Adnl1", "OtherCost1", SAPbobsCOM.BoFldSubTypes.st_Price)


        objUDFEngine.CreateTable("PSSIT_WOR2", "WorkOrder-RouteDetails", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        objUDFEngine.AddNumericField("@PSSIT_WOR2", "Pordser", "ProductionOrderSeries", 6)
        objUDFEngine.AddNumericField("@PSSIT_WOR2", "Pordno", "ProductionOrderNo.", 10)
        objUDFEngine.AddNumericField("@PSSIT_WOR2", "Baslino", "BaseLineNo.", 2)
        objUDFEngine.AddNumericField("@PSSIT_WOR2", "Seqnce", "OperationSequence", 10)
        objUDFEngine.AddNumericField("@PSSIT_WOR2", "Parid", "ParentId", 10)
        objUDFEngine.AddAlphaField("@PSSIT_WOR2", "Oprcode", "Operation", 20)
        objUDFEngine.AddAlphaField("@PSSIT_WOR2", "Oprname", "OperationName", 100)
        objUDFEngine.addField("@PSSIT_WOR2", "Rework", "Rework", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")
        objUDFEngine.AddAlphaField("@PSSIT_WOR2", "Rteid", "RouteID", 20)
        objUDFEngine.AddNumericField("@PSSIT_WOR2", "Seqbaslino", "SequenceBaseLineID", 2)
        objUDFEngine.AddFloatField("@PSSIT_WOR2", "ProdQty", "ProducedQty.", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddFloatField("@PSSIT_WOR2", "Passqty", "PassedQty.", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddFloatField("@PSSIT_WOR2", "Rewrkqty", "Rewrkqty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddFloatField("@PSSIT_WOR2", "PenRewQty", "PendingReworkQty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddFloatField("@PSSIT_WOR2", "scrapqty", "ScrapQty.", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddFloatField("@PSSIT_WOR2", "Lbrcst", "LabourCost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_WOR2", "Mccst", "MachineCost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_WOR2", "Toolcst", "ToolCost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_WOR2", "Subcst", "SubContractingCost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_WOR2", "Scrapcst", "ScrapCost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddNumericField("@PSSIT_WOR2", "Wodoc", "WorkOrderDoc.Entry", 10)
        objUDFEngine.AddFloatField("@PSSIT_WOR2", "Adnl1", "Qty1", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddFloatField("@PSSIT_WOR2", "Adnl2", "Qty2", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddFloatField("@PSSIT_WOR2", "Adnl3", "OtherCost1", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_WOR2", "Adnl4", "OtherCost2", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.addField("@PSSIT_WOR2", "Closekey", "OperationClosedKey", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "N")

        objUDFEngine.CreateTable("PSSIT_RTE4", "Operation Routing - Rows", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objUDFEngine.AddNumericField("@PSSIT_RTE4", "Seqnce", "Sequence", 10)
        objUDFEngine.AddNumericField("@PSSIT_RTE4", "Parid", "Parent Id", 10)
        objUDFEngine.AddAlphaField("@PSSIT_RTE4", "Oprcode", "Operation", 20)
        objUDFEngine.AddAlphaField("@PSSIT_RTE4", "Oprname", "Operation Name", 100)
        objUDFEngine.addField("@PSSIT_RTE4", "Milestne", "Mile Stone", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")
        objUDFEngine.AddAlphaField("@PSSIT_RTE4", "Oprtype", "Operation Type", 50)
        objUDFEngine.AddFloatField("@PSSIT_RTE4", "ScRate", "SubContracting Rate", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@PSSIT_RTE4", "Qty", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddAlphaField("@PSSIT_RTE4", "Adnl1", "Other Info 1", 50)
        objUDFEngine.AddAlphaField("@PSSIT_RTE4", "Adnl2", "Other Info 2", 10)
        objUDFEngine.AddNumericField("@PSSIT_RTE4", "Bselino", "#", 2)
        objUDFEngine.addField("@PSSIT_RTE4", "Rework", "Rework", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")


        objApplication.SetStatusBarMessage("Tables Created Sucessfully!", SAPbouiCOM.BoMessageTime.bmt_Short, False)









    End Sub


    Private Sub CreateUDOS()


        objApplication.SetStatusBarMessage("UDO is Regestering Please Wait.......!", SAPbouiCOM.BoMessageTime.bmt_Long, False)
        Dim Ct(0) As String
        Ct(0) = ""
        createUDO("PSSIT_OSFT", "PSSIT_SFT", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)

        Ct(0) = ""
        createUDO("PSSIT_OTYP", "PSSIT_OTYP", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, True, True)
        Ct(0) = ""
        createUDO("PSSIT_OMGP", "PSSIT_MGP", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)


        Ct(0) = ""
        createUDO("PSSIT_OLGP", "PSSIT_LGP", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)
        Ct(0) = ""
        createUDO("PSSIT_OLBR", "PSSIT_LBR", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)
        Ct(0) = ""
        createUDO("PSSIT_OTLS", "PSSIT_TLS", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)

        Ct(0) = ""
        createUDO("PSSIT_OSGE", "PSSIT_SGE", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData)

        Ct(0) = ""
        createUDO("PSSIT_OCST", "PSSIT_CST", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, True, True)

        Ct(0) = ""
        createUDO("PSSIT_ORES", "PSSIT_RES", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, True)

        Ct(0) = ""
        createUDO("PSSIT_OSCY", "PSSIT_SCY", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, True)

        Ct(0) = ""
        createUDO("PSSIT_PMREMEDIES", "PSSIT_REMEDIES", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, True, True)

        Ct(0) = ""
        createUDO("PSSIT_PMWCMAKE", "PSSIT_MAK", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, True, True)

        Ct(0) = ""
        createUDO("PSSIT_PMWCMODEL", "PSSIT_MOD", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, True, True)

        Ct(0) = ""
        createUDO("PSSIT_PMWCUOM", "PSSIT_OUOM", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, True)


        ReDim Ct(1)
        Ct(0) = "PSSIT_WCR1"
        createUDO("PSSIT_OWCR", "PSSIT_WCR", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)

        Ct(0) = "PSSIT_RTE4"
        createUDO("PSSIT_ORTE", "PSSIT_RTE", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)
        'Modified by Manimaran----s
        'Ct(0) = "PSSIT_PEY1"
        'createUDO("PSSIT_OPEY", "PSSIT_PEY", Ct, SAPbobsCOM.BoUDOObjType.boud_Document)
        ReDim Ct(3)
        Ct(0) = "PSSIT_PEY1"
        Ct(1) = "PSSIT_PEY5"
        Ct(2) = "PSSIT_PEY6"
        createUDO("PSSIT_OPEY", "PSSIT_PEY", Ct, SAPbobsCOM.BoUDOObjType.boud_Document)
        'Modified by Manimaran----e

        ReDim Ct(2)
        Ct(0) = "PSSIT_PRN1"
        Ct(1) = "PSSIT_PRN2"
        'Ct(2) = "PSSIT_PRN3"

        createUDO("PSSIT_OPRN", "PSSIT_PRN", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)


        ReDim Ct(3)
        Ct(0) = "PSSIT_PMWCREASONDTL"
        Ct(1) = "PSSIT_PMWCREMEDTL"
        Ct(2) = "PSSIT_PMWCITEMSDTL"
        createUDO("PSSIT_PMWCBREAKHDR", "PSSIT_WCBREAK", Ct, SAPbobsCOM.BoUDOObjType.boud_Document)



        ReDim Ct(4)
        Ct(0) = "PSSIT_PMWCSPEC"
        Ct(1) = "PSSIT_PMWCPARA"
        Ct(2) = "PSSIT_PMWCITEM"
        Ct(3) = "PSSIT_PMWCSFT"
        createUDO("PSSIT_PMWCHDR", "PSSIT_WCHDR", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)




    End Sub
    'Code has been modified by Shankar 09-Nov-2009.
    'Function has been created by shankar

    Private Sub createUDO(ByVal tblname As String, ByVal udoname As String, ByVal childTable() As String, ByVal type As SAPbobsCOM.BoUDOObjType, Optional ByVal DfltForm As Boolean = False, Optional ByVal FindForm As Boolean = False)
        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Dim creationPackage As SAPbouiCOM.FormCreationParams
        Dim objform As SAPbouiCOM.Form
        Dim i As Integer
        'Dim c_Yes As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tYES
        Dim lRetCode As Long
        oUserObjectMD = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
        If Not oUserObjectMD.GetByKey(udoname) Then
            oUserObjectMD.Code = udoname
            oUserObjectMD.Name = udoname
            oUserObjectMD.ObjectType = type
            oUserObjectMD.TableName = tblname
            oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            If DfltForm = True Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                'oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES

                oUserObjectMD.FormColumns.FormColumnAlias = "Code"
                oUserObjectMD.FormColumns.FormColumnDescription = "Code"
                oUserObjectMD.FormColumns.Add()
                Select Case udoname
                    Case "PSSIT_OTYP"
                        oUserObjectMD.FormColumns.FormColumnAlias = "U_Wctypnam"
                        oUserObjectMD.FormColumns.FormColumnDescription = "Description"
                        oUserObjectMD.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                        oUserObjectMD.FormColumns.Add()
                    Case "PSSIT_WCT"
                        oUserObjectMD.FormColumns.FormColumnAlias = "U_Wctypnam"
                        oUserObjectMD.FormColumns.FormColumnDescription = "Description"
                        oUserObjectMD.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                        oUserObjectMD.FormColumns.Add()
                    Case "PSSIT_CST"
                        oUserObjectMD.FormColumns.FormColumnAlias = "U_Wcstyp"
                        oUserObjectMD.FormColumns.FormColumnDescription = "Description"
                        oUserObjectMD.FormColumns.Add()
                    Case "PSSIT_RES"
                        'oUserObjectMD.FormColumns.FormColumnAlias = "U_wcno"
                        'oUserObjectMD.FormColumns.FormColumnDescription = "Machine"
                        'oUserObjectMD.FormColumns.Add()
                        oUserObjectMD.FormColumns.FormColumnAlias = "Name"
                        oUserObjectMD.FormColumns.FormColumnDescription = "Name"
                        oUserObjectMD.FormColumns.Add()
                    Case "PSSIT_SCY"
                        oUserObjectMD.FormColumns.FormColumnAlias = "U_Catname"
                        oUserObjectMD.FormColumns.FormColumnDescription = "CategoryDescription"
                        oUserObjectMD.FormColumns.Add()
                        oUserObjectMD.FormColumns.FormColumnAlias = "U_Remarks"
                        oUserObjectMD.FormColumns.FormColumnDescription = "Remarks"
                        oUserObjectMD.FormColumns.Add()

                    Case "PSSIT_REMEDIES"
                        oUserObjectMD.FormColumns.FormColumnAlias = "U_remedesc"
                        oUserObjectMD.FormColumns.FormColumnDescription = "RemediesDesc"
                        oUserObjectMD.FormColumns.Add()

                    Case "PSSIT_MAK"
                        oUserObjectMD.FormColumns.FormColumnAlias = "U_makedesc"
                        oUserObjectMD.FormColumns.FormColumnDescription = "Make Name"
                        oUserObjectMD.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                        oUserObjectMD.FormColumns.Add()

                    Case "PSSIT_MOD"
                        oUserObjectMD.FormColumns.FormColumnAlias = "U_modedesc"
                        oUserObjectMD.FormColumns.FormColumnDescription = "ModelName"
                        oUserObjectMD.FormColumns.Add()

                    Case "PSSIT_OUOM"
                        oUserObjectMD.FormColumns.FormColumnAlias = "U_uomdesc"
                        oUserObjectMD.FormColumns.FormColumnDescription = "UOMName"
                        oUserObjectMD.FormColumns.Add()


                End Select


            End If
            If FindForm = True Then
                If type = SAPbobsCOM.BoUDOObjType.boud_MasterData Then
                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                    oUserObjectMD.FindColumns.ColumnAlias = "Code"
                    oUserObjectMD.FindColumns.ColumnDescription = "Code"
                    oUserObjectMD.FindColumns.Add()
                    Select Case udoname
                        Case "PSSIT_WCT"
                            oUserObjectMD.FindColumns.ColumnAlias = "U_Wctypnam"
                            oUserObjectMD.FindColumns.ColumnDescription = "Description"
                            oUserObjectMD.FindColumns.Add()
                        Case "PSSIT_CST"
                            oUserObjectMD.FindColumns.ColumnAlias = "U_Wcstyp"
                            oUserObjectMD.FindColumns.ColumnDescription = "Description"
                            oUserObjectMD.FindColumns.Add()
                        Case "PSSIT_MOD"
                            oUserObjectMD.FindColumns.ColumnAlias = "U_modedesc"
                            oUserObjectMD.FindColumns.ColumnDescription = "ModelName"
                            oUserObjectMD.FindColumns.Add()
                        Case "PSSIT_MAK"
                            oUserObjectMD.FindColumns.ColumnAlias = "U_makedesc"
                            oUserObjectMD.FindColumns.ColumnDescription = "MakeName"
                            oUserObjectMD.FindColumns.Add()
                        Case "PSSIT_LGP"
                            oUserObjectMD.FindColumns.ColumnAlias = "U_LGname"
                            oUserObjectMD.FindColumns.ColumnDescription = "Description"
                            oUserObjectMD.FindColumns.Add()
                        Case "PSSIT_PRN"
                            oUserObjectMD.FindColumns.ColumnAlias = "U_Oprname"
                            oUserObjectMD.FindColumns.ColumnDescription = "OperationName"
                            oUserObjectMD.FindColumns.Add()
                        Case "PSSIT_RTE"
                            oUserObjectMD.FindColumns.ColumnAlias = "Name"
                            oUserObjectMD.FindColumns.ColumnDescription = "Name"
                            oUserObjectMD.FindColumns.Add()
                        Case "PSSIT_REMEDIES"
                            oUserObjectMD.FindColumns.ColumnAlias = "U_remedesc"
                            oUserObjectMD.FindColumns.ColumnDescription = "RemediesDesc"
                            oUserObjectMD.FindColumns.Add()
                        Case "PSSIT_WCR"
                            oUserObjectMD.FindColumns.ColumnAlias = "U_WCname"
                            oUserObjectMD.FindColumns.ColumnDescription = "Description"
                            oUserObjectMD.FindColumns.Add()
                        Case "PSSIT_MGP"
                            oUserObjectMD.FindColumns.ColumnAlias = "U_Mgname"
                            oUserObjectMD.FindColumns.ColumnDescription = "MachineGroupName"
                            oUserObjectMD.FindColumns.Add()
                        Case "PSSIT_SFT"
                            oUserObjectMD.FindColumns.ColumnAlias = "U_Sdescr"
                            oUserObjectMD.FindColumns.ColumnDescription = "ShiftDescription"
                            oUserObjectMD.FindColumns.Add()
                        Case "PSSIT_WCHDR"
                            oUserObjectMD.FindColumns.ColumnAlias = "U_wcno"
                            oUserObjectMD.FindColumns.ColumnDescription = "WorkCenterNo"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_wcname"
                            oUserObjectMD.FindColumns.ColumnDescription = "WorkCenterName"
                            oUserObjectMD.FindColumns.Add()
                        Case "PSSIT_TLS"
                            oUserObjectMD.FindColumns.ColumnAlias = "U_TLname"
                            oUserObjectMD.FindColumns.ColumnDescription = "ToolDescription"
                            oUserObjectMD.FindColumns.Add()
                        Case "PSSIT_LBR"
                            oUserObjectMD.FindColumns.ColumnAlias = "U_Empid"
                            oUserObjectMD.FindColumns.ColumnDescription = "EmployeeNo."
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_Empnam"
                            oUserObjectMD.FindColumns.ColumnDescription = "EmployeeName"
                            oUserObjectMD.FindColumns.Add()
                    End Select

                Else

                End If
            End If
            If childTable.Length > 0 Then
                For i = 0 To childTable.Length - 1
                    If Trim(childTable(i)) <> "" Then
                        oUserObjectMD.ChildTables.TableName = childTable(i)
                        oUserObjectMD.ChildTables.Add()
                    End If
                Next
            End If
            lRetCode = oUserObjectMD.Add()
            If lRetCode <> 0 Then
                MsgBox("error" + CStr(lRetCode))
                MsgBox(objCompany.GetLastErrorDescription)
            Else

            End If
            If DfltForm = True Then
                creationPackage = objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
                ' Need to set the parameter with the object unique ID
                creationPackage.ObjectType = "1"
                creationPackage.UniqueID = udoname
                creationPackage.FormType = udoname
                creationPackage.BorderStyle = SAPbouiCOM.BoFormTypes.ft_Fixed
                objform = objApplication.Forms.AddEx(creationPackage)
            End If
        End If

    End Sub



    Private Sub applyFilter()

    End Sub
    
End Class