Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Data
Imports System.Data.OleDb
Imports System
Imports System.IO
Imports System.Xml
Imports System.Text
Imports System.Security.Cryptography
Public Class FrmProdCostCrystalRptVwr
    Inherits System.Windows.Forms.Form

    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private SboGuiApi As New SAPbouiCOM.SboGuiApi
    Private oCompany As SAPbobsCOM.Company
    Private ProdCostReport As ReportDocument
    Private oConnection As OleDb.OleDbConnection
    Private oDataAdapter As OleDb.OleDbDataAdapter
    Private oCONNECTION_STRING As String
    Private oQUERY_STRING As String
    Private oDataSet As ProdCostDataSet
    Private oFPordNo, oTPordNo, oStatus As String
    Private sPath As String = IO.Directory.GetParent(Application.ExecutablePath).ToString
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company, ByVal aFPordNo As String, ByVal aTPordNo As String, ByVal aStatus As String)
        MyBase.New()
        Dim StrConnString As String = ""
        Dim SqlStr As String = ""
        Dim SqlStr1 As String = ""


        SBO_Application = aSBO_Application
        oCompany = aCompany
        oFPordNo = aFPordNo
        oTPordNo = aTPordNo
        oStatus = aStatus

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
            ProdCostReport = New ReportDocument
            ProdCostReport.Load(sPath & "\Reports\PordCost2N.rpt")
            ProdCostReport.SetDataSource(oDataSet)
            'PordCostCrystalRptVwr.DisplayGroupTree = False
            PordCostCrystalRptVwr.ReportSource = ProdCostReport
            'ProcSheetRptVwr.Refresh()
            PordCostCrystalRptVwr.Zoom(1)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub FetchData()
        Dim PSSIT_ProdCostTable, PSSIT_ProdCostMatrlTable, PSSIT_ProdCostMachTable, PSSIT_ProdCostManPowTable, PSSIT_ProdCostToolTable, PSSIT_ProdCostFxdCstTable As DataTable
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
            'oCONNECTION_STRING = "Provider=SQLOLEDB;Server=" & oCompany.Server & ";Database=" & oCompany.CompanyDB & ";User ID=" & oCompany.DbUserName & ";Password= sa2008"

            oConnection = New OleDbConnection(oCONNECTION_STRING)
            oConnection.Open()
            oQUERY_STRING = "Production_Cost_Report '" & oFPordNo & "','" & oTPordNo & "','" & oStatus & "'"
            oDataAdapter = New OleDbDataAdapter(oQUERY_STRING, oConnection)
            oDataSet = New ProdCostDataSet()

            PSSIT_ProdCostTable = New DataTable("Production_Cost_Report")
            oDataAdapter.Fill(oDataSet, "Production_Cost_Report")

            oQUERY_STRING = "Production_Cost_Matrl_Report '" & oFPordNo & "','" & oTPordNo & "','" & oStatus & "'"
            oDataAdapter = New OleDbDataAdapter(oQUERY_STRING, oConnection)
            PSSIT_ProdCostMatrlTable = New DataTable("Production_Cost_Matrl_Report")
            oDataAdapter.Fill(oDataSet, "Production_Cost_Matrl_Report")

            oQUERY_STRING = "Production_Cost_Mach_Report '" & oFPordNo & "','" & oTPordNo & "','" & oStatus & "'"
            oDataAdapter = New OleDbDataAdapter(oQUERY_STRING, oConnection)
            PSSIT_ProdCostMachTable = New DataTable("Production_Cost_Mach_Report")
            oDataAdapter.Fill(oDataSet, "Production_Cost_Mach_Report")

            oQUERY_STRING = "Production_Cost_ManPow_Report '" & oFPordNo & "','" & oTPordNo & "','" & oStatus & "'"
            oDataAdapter = New OleDbDataAdapter(oQUERY_STRING, oConnection)
            PSSIT_ProdCostManPowTable = New DataTable("Production_Cost_ManPow_Report")
            oDataAdapter.Fill(oDataSet, "Production_Cost_ManPow_Report")

            oQUERY_STRING = "Production_Cost_Tool_Report '" & oFPordNo & "','" & oTPordNo & "','" & oStatus & "'"
            oDataAdapter = New OleDbDataAdapter(oQUERY_STRING, oConnection)
            PSSIT_ProdCostToolTable = New DataTable("Production_Cost_Tool_Report")
            oDataAdapter.Fill(oDataSet, "Production_Cost_Tool_Report")

            oQUERY_STRING = "Production_Cost_FxdCst_Report '" & oFPordNo & "','" & oTPordNo & "','" & oStatus & "'"
            oDataAdapter = New OleDbDataAdapter(oQUERY_STRING, oConnection)
            PSSIT_ProdCostFxdCstTable = New DataTable("Production_Cost_FxdCst_Report")
            oDataAdapter.Fill(oDataSet, "Production_Cost_FxdCst_Report")
        Catch ex As Exception
            Throw ex
        Finally
            oConnection.Close()
        End Try
    End Sub
End Class