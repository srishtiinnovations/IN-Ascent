Imports System.Data.OleDb

Public Class GeneralLib
#Region "Variable Declaration"
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oCompany As SAPbobsCOM.Company
#End Region
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company)
        MyBase.New()
        SBO_Application = aSBO_Application
        oCompany = aCompany
    End Sub
#Region "Loading the Form"
    Public Function LoadForm(ByRef oCompany As SAPbobsCOM.Company, ByRef oApplication As SAPbouiCOM.Application, ByVal FileName As String) As SAPbouiCOM.Form
        Dim sPath As String
        sPath = IO.Directory.GetParent(Application.ExecutablePath).ToString
        LoadForm = Nothing
        Try
            Dim FormUID As String = LoadUniqueFormXML(oCompany, oApplication, sPath & "\Forms\" & FileName)
            LoadForm = oApplication.Forms.Item(FormUID)
        Catch ex As Exception
            oApplication.MessageBox("LoadForm(" & FileName & "): " & oCompany.GetLastErrorCode & ", " & ex.Message)
        End Try
    End Function
    Private Function LoadUniqueFormXML(ByRef oCompany As SAPbobsCOM.Company, ByRef oApplication As SAPbouiCOM.Application, ByVal FileName As String) As String
        Dim xDoc As System.Xml.XmlDocument = New Xml.XmlDocument
        LoadUniqueFormXML = ""
        Try
            xDoc.Load(FileName)
            LoadUniqueFormXML = xDoc.SelectSingleNode("Application/forms/action/form/@FormType").Value & "_" & MaximoTipoForm(oCompany, oApplication, xDoc.SelectSingleNode("Application/forms/action/form/@FormType").Value).ToString
            xDoc.SelectSingleNode("Application/forms/action/form/@uid").Value = LoadUniqueFormXML
            oApplication.LoadBatchActions(xDoc.InnerXml)
        Catch ex As Exception
            oApplication.MessageBox(ex.Message)
        End Try
    End Function
    Private Function MaximoTipoForm(ByRef oCompany As SAPbobsCOM.Company, ByRef oApplication As SAPbouiCOM.Application, ByRef Tipo As String) As Long
        MaximoTipoForm = 0

        Try
            For Each iform As SAPbouiCOM.Form In oApplication.Forms
                If iform.TypeEx = Tipo Then
                    If iform.TypeCount > MaximoTipoForm Then
                        MaximoTipoForm = iform.TypeCount
                    End If
                End If
            Next
            MaximoTipoForm = MaximoTipoForm + 1
        Catch ex As Exception
            oApplication.MessageBox(ex.Message)
        End Try
    End Function
#End Region
    ''' <summary>
    ''' Loading the form.
    ''' </summary>
    ''' <param name="FileName"></param>
    ''' <remarks></remarks>
    Public Sub LoadFromXML(ByRef FileName As String)
        Dim oXmlDoc As Xml.XmlDocument
        oXmlDoc = New Xml.XmlDocument
        Dim sPath As String
        sPath = IO.Directory.GetParent(Application.ExecutablePath).ToString
        oXmlDoc.Load(sPath & "\Forms\" & FileName)
        SBO_Application.LoadBatchActions(oXmlDoc.InnerXml)
    End Sub
#Region "Formatted Search ..."
    Public Sub AssignUserQueries(ByVal aQryName As String, ByVal aCategory As String, ByVal aFormType As String, ByVal aItem As String, ByVal aCol As String)
        Dim oRs As SAPbobsCOM.Recordset
        Dim oTransaction As Boolean
        Dim CategoryID = "", QueryID = "", IndexID As String = "1"
        Try
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery("SELECT CategoryId FROM OQCN WHERE CatName='" & aCategory & "'")
            If oRs.RecordCount > 0 Then
                CategoryID = oRs.Fields.Item("CategoryID").Value
                oRs.DoQuery("SELECT IntrnalKey FROM OUQR WHERE QCategory=" & CategoryID & " AND QName='" & aQryName & "'")
                If oRs.RecordCount > 0 Then
                    QueryID = oRs.Fields.Item("IntrnalKey").Value
                    oRs.DoQuery("SELECT QueryId FROM CSHS WHERE FormID='" & aFormType & "' AND ItemID='" & aItem & "' AND ColID='" & aCol & "'")
                    If oRs.RecordCount = 0 Then
                        oTransaction = True
                        oCompany.StartTransaction()
                        oRs.DoQuery("SELECT TOP 1 IndexID FROM CSHS ORDER BY IndexID DESC")
                        If oRs.RecordCount > 0 Then
                            IndexID = oRs.Fields.Item("IndexID").Value + 1
                        Else
                            IndexID = "1"
                        End If
                        Dim StrSql As String = "INSERT INTO CSHS (FormID, ItemID, ColID, ActionT, QueryId, IndexID, Refresh, FieldID,FrceRfrsh, ByField) " _
                        & " VALUES ('" & aFormType & "','" & aItem & "','" & aCol & "',2," & QueryID & "," & IndexID & ",'N',null,'N','N')"
                        oRs.DoQuery(StrSql)
                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        oTransaction = False
                    End If
                Else
                    Throw New Exception("Query Not Found")
                End If
            Else
                Throw New Exception("Category Not Found")
            End If
        Catch ex As Exception
            If oTransaction = True Then
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            Throw ex
        Finally
            oRs = Nothing
            GC.Collect()
        End Try
    End Sub
#End Region
    ''' <summary>
    ''' Generating the Serial Number.
    ''' </summary> 
    ''' <param name="aTableName"></param>
    ''' <param name="aCriteriaSqlStr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
#Region "Generate Serial No"
    Public Function GenerateSerialNo(ByVal aTableName As String, Optional ByVal aCriteriaSqlStr As String = "") As Integer
        Dim StrSql As String
        Dim oRs As SAPbobsCOM.Recordset
        Dim oCode As Integer
        Try

            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If aCriteriaSqlStr.Length = 0 Then
                StrSql = "Select IsNull(Max(Convert(Float,Code)),0) as Code From [@" & aTableName & "]"
            Else
                StrSql = aCriteriaSqlStr
            End If
            oRs.DoQuery(StrSql)
            If oRs.RecordCount > 0 Then
                oCode = oRs.Fields.Item("Code").Value + 1
            Else
                oCode = 1
            End If
            Return oCode
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region
#Region "Generate Docnum"
    Public Function GenerateDocNo(ByVal aTableName As String, Optional ByVal aCriteriaSqlStr As String = "") As Integer
        Dim StrSql As String
        Dim oRs As SAPbobsCOM.Recordset
        Dim oDocNum As Integer
        Try

            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If aCriteriaSqlStr.Length = 0 Then
                StrSql = "Select Count(*) From [@" & aTableName & "]"
            Else
                StrSql = aCriteriaSqlStr
            End If
            oRs.DoQuery(StrSql)
            If oRs.RecordCount > 0 Then
                oDocNum = oRs.Fields.Item(0).Value + 1
            Else
                oDocNum = 1
            End If
            Return oDocNum
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region
#Region "Loading Default form"
    Public Sub LoadDefaultForm(ByVal sFormUID As String)
        Dim i As Integer
        ' Link to the Default Forms menu
        Dim sboMenu As SAPbouiCOM.MenuItem = SBO_Application.Menus.Item("47616")
        Try
            ' Iterate through the submenus to find the correct UDO
            If sboMenu.SubMenus.Count > 0 Then
                For i = 0 To sboMenu.SubMenus.Count - 1
                    If sboMenu.SubMenus.Item(i).String.Contains(sFormUID) Then
                        sboMenu.SubMenus.Item(i).Activate()
                    End If
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#Region "Loading NoObject Forms"
    Public Sub LoadNoObjectTable(ByVal sTableName As String)
        Dim i As Integer
        ' Link to the User Table menu
        Dim sboMenu As SAPbouiCOM.MenuItem = SBO_Application.Menus.Item("51200")
        Try
            ' Iterate through the submenus to find the correct UDO
            If sboMenu.SubMenus.Count > 0 Then
                For i = 0 To sboMenu.SubMenus.Count - 1
                    If sboMenu.SubMenus.Item(i).String.Contains(sTableName) Then
                        sboMenu.SubMenus.Item(i).Activate()
                    End If
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#Region "Loading Format Account Code"
    Public Function FormatAccountCode(ByVal oFormatCode As String) As String

        Dim StrSql As String
        Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oFCode As String = ""
        Try
            StrSql = "Select Tbl.AcctCode,Tbl.FormatCode, " _
            & "Case when Tbl.Account is null then Tbl.FormatCode Else Tbl.Account End as Account " _
            & "From(select AcctCode,FormatCode, " _
            & "Segment_0 + isnull( '-' + Segment_1, '') + isnull('-' + Segment_2, '') " _
            & "+ isnull( '-' + Segment_3, '') + isnull('-' + Segment_4, '')  " _
            & "+ isnull( '-' + Segment_5, '') + isnull('-' + Segment_6, '') " _
            & "+ isnull( '-' + Segment_7, '') + isnull('-' + Segment_8, '') " _
            & "+ isnull( '-' + Segment_9, '') 	as Account from " _
            & "OACT Where FormatCode='" & oFormatCode & "')Tbl"
            oRS.DoQuery(StrSql)
            If oRS.RecordCount > 0 Then
                oRS.MoveFirst()
                oFCode = oRS.Fields.Item(2).Value
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return oFCode
    End Function
    Private Function AcctCode(ByVal aAcctCode As String) As String
        Dim sStr As String
        Dim vRs As SAPbobsCOM.Recordset
        Dim vBOB As SAPbobsCOM.SBObob
        Dim vCH As SAPbobsCOM.ChartOfAccounts
        vCH = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oChartOfAccounts)
        vBOB = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        vRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        ' When working with segmentation use this function
        ' to find the account key in the ChartOfAccount object

        vRs = vBOB.GetObjectKeyBySingleValue(SAPbobsCOM.BoObjectTypes.oChartOfAccounts, "FormatCode", aAcctCode, SAPbobsCOM.BoQueryConditions.bqc_Equal)
        sStr = vRs.Fields.Item(0).Value

        'The Recordset retrieves the value of the key (for example, sStr = _SYS00000000010).

        Return sStr
    End Function
#End Region
#Region "Crystal Reports"
    Public Shared Sub SetCrystalLogin(ByVal sUser As String, ByVal sPassword As String, ByVal sServer As String, ByVal sCompanyDB As String, _
              ByRef oRpt As CrystalDecisions.CrystalReports.Engine.ReportDocument)

        Dim oDB As CrystalDecisions.CrystalReports.Engine.Database = oRpt.Database
        Dim oTables As CrystalDecisions.CrystalReports.Engine.Tables = oDB.Tables
        Dim oLogonInfo As CrystalDecisions.Shared.TableLogOnInfo
        Dim oConnectInfo As CrystalDecisions.Shared.ConnectionInfo = New CrystalDecisions.Shared.ConnectionInfo()
        oConnectInfo.DatabaseName = sCompanyDB
        oConnectInfo.ServerName = sServer
        oConnectInfo.UserID = sUser
        oConnectInfo.Password = sPassword
        ' Set the logon credentials for all tables
        For Each oTable As CrystalDecisions.CrystalReports.Engine.Table In oTables
            oLogonInfo = oTable.LogOnInfo
            oLogonInfo.ConnectionInfo = oConnectInfo
            oTable.ApplyLogOnInfo(oLogonInfo)
        Next
        ' Check for subreports
        Dim oSections As CrystalDecisions.CrystalReports.Engine.Sections
        Dim oSection As CrystalDecisions.CrystalReports.Engine.Section
        Dim oRptObjs As CrystalDecisions.CrystalReports.Engine.ReportObjects
        Dim oRptObj As CrystalDecisions.CrystalReports.Engine.ReportObject
        Dim oSubRptObj As CrystalDecisions.CrystalReports.Engine.SubreportObject
        Dim oSubRpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        oSections = oRpt.ReportDefinition.Sections
        For Each oSection In oSections
            oRptObjs = oSection.ReportObjects
            For Each oRptObj In oRptObjs

                If oRptObj.Kind = CrystalDecisions.Shared.ReportObjectKind.SubreportObject Then

                    ' This is a subreport so set the logon credentials for this report's tables
                    oSubRptObj = CType(oRptObj, CrystalDecisions.CrystalReports.Engine.SubreportObject)
                    ' Open the subreport
                    oSubRpt = oSubRptObj.OpenSubreport(oSubRptObj.SubreportName)

                    oDB = oSubRpt.Database
                    oTables = oDB.Tables

                    For Each oTable As CrystalDecisions.CrystalReports.Engine.Table In oTables
                        oLogonInfo = oTable.LogOnInfo
                        oLogonInfo.ConnectionInfo = oConnectInfo
                        oTable.ApplyLogOnInfo(oLogonInfo)
                    Next

                End If

            Next
        Next

    End Sub
#End Region
    ''' <summary>
    ''' Properties for Usertables.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
#Region "Properties"
    Public ReadOnly Property UserTables() As SAPbobsCOM.UserTables
        Get
            Return oCompany.UserTables
        End Get
    End Property
#End Region

    Private Sub SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent

    End Sub

#Region "SQL server Login Valiation"
    Public Function LoginValiation() As String
        Dim oTempRs As SAPbobsCOM.Recordset
        Dim sqlserver, sqluid, sqlpwd, oCONNECTION_STRING As String
        Dim oConnection As OleDbConnection
        oTempRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRs.DoQuery("Select * from [@PSSIT_OCON]")
        If oTempRs.RecordCount > 0 Then
            sqlserver = oTempRs.Fields.Item("U_SqlSer").Value
            sqluid = oTempRs.Fields.Item("U_SqlUID").Value
            sqlpwd = oTempRs.Fields.Item("U_SqlPwd").Value
            oCONNECTION_STRING = "Provider=SQLOLEDB;Server=" & oCompany.Server & ";Database=" & oCompany.CompanyDB & ";User ID=" & sqluid & ";Password=" & sqlpwd
            oConnection = New OleDbConnection(oCONNECTION_STRING)
            Try
                oConnection.Open()
            Catch ex As Exception
                SBO_Application.SetStatusBarMessage("Sql server login failed. Check the login details in Production Setup", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Return ""
            End Try
        Else
            SBO_Application.SetStatusBarMessage("Sql login details are missing. Enter the details in the production setup", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return ""
            Exit Function
        End If
        oCONNECTION_STRING = "Provider=SQLOLEDB;Server=" & oCompany.Server & ";Database=" & oCompany.CompanyDB & ";User ID=" & sqluid & ";Password=" & sqlpwd
        Return oCONNECTION_STRING
    End Function
#End Region

End Class
