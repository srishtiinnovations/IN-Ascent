Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Data
Imports System.Data.OleDb
Public Class FrmProcessSheetRptVwr
    Inherits System.Windows.Forms.Form
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private SboGuiApi As New SAPbouiCOM.SboGuiApi
    Private oCompany As SAPbobsCOM.Company
    Private ProcessSheetReport As ReportDocument
    Private oConnection As OleDb.OleDbConnection
    Private oDataAdapter As OleDb.OleDbDataAdapter
    Private oCONNECTION_STRING As String
    Private oQUERY_STRING As String
    Private oDataSet As ProcessSheetDataSet
    Private sPath As String = IO.Directory.GetParent(Application.ExecutablePath).ToString
    Private oPordNo As String
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company, ByVal aPordNo As String)
        MyBase.New()
        Dim StrConnString As String = ""
        Dim SqlStr As String = ""

        SBO_Application = aSBO_Application
        oCompany = aCompany
        oPordNo = aPordNo

        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ConfigureCrystalReports()

    End Sub
    Private Sub ConfigureCrystalReports()
        Dim reportPath As String = IO.Directory.GetParent(Application.ExecutablePath).ToString
        Try
            Dim oTempRs As SAPbobsCOM.Recordset
            Dim sqlserver, sqluid, sqlpwd As String
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
                    Exit Sub
                End Try

            Else
                SBO_Application.SetStatusBarMessage("Sql login details are missing. Enter the details in the production setup", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Exit Sub
            End If
            FetchData()
            ProcessSheetReport = New ReportDocument
            ProcessSheetReport.Load(sPath & "\Reports\ProcessSheetReport1.rpt")
            ProcessSheetReport.SetDataSource(oDataSet)
            'ProcSheetRptVwr.DisplayGroupTree = False
            ProcSheetRptVwr.ReportSource = ProcessSheetReport
            'ProcSheetRptVwr.Refresh()
            ProcSheetRptVwr.Zoom(1)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub FetchData()
        Dim PSSIT_ProcSheetTable, PSSIT_ProcSheetMachTable As DataTable
        Try
            Dim oTempRs As SAPbobsCOM.Recordset
            Dim sqlserver, sqluid, sqlpwd As String
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
                    Exit Sub
                End Try
            Else
                SBO_Application.SetStatusBarMessage("Sql login details are missing. Enter the details in the production setup", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Exit Sub
            End If
            oCONNECTION_STRING = "Provider=SQLOLEDB;Server=" & oCompany.Server & ";Database=" & oCompany.CompanyDB & ";User ID=" & sqluid & ";Password=" & sqlpwd
            'oCONNECTION_STRING = "Provider=SQLOLEDB;Server=" & oCompany.Server & ";Database=" & oCompany.CompanyDB & ";User ID=" & oCompany.DbUserName & ";Password=sa2005"
            oConnection = New OleDbConnection(oCONNECTION_STRING)
            oConnection.Open()
            oQUERY_STRING = "ProcessSheet_Report '" & oPordNo & "'"
            oDataAdapter = New OleDbDataAdapter(oQUERY_STRING, oConnection)
            oDataSet = New ProcessSheetDataSet()

            PSSIT_ProcSheetTable = New DataTable("ProcessSheet_Report")
            oDataAdapter.Fill(oDataSet, "ProcessSheet_Report")

            oQUERY_STRING = "ProcessSheet_Mach_Report '" & oPordNo & "'"
            oDataAdapter = New OleDbDataAdapter(oQUERY_STRING, oConnection)
            PSSIT_ProcSheetMachTable = New DataTable("ProcessSheet_Mach_Report")
            oDataAdapter.Fill(oDataSet, "ProcessSheet_Mach_Report")
        Catch ex As Exception
            Throw ex
        Finally
            oConnection.Close()
        End Try
    End Sub
End Class