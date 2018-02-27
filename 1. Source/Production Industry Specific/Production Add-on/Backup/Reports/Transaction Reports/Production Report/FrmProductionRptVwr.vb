Public Class FrmProductionRptVwr
    Inherits System.Windows.Forms.Form
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private SboGuiApi As New SAPbouiCOM.SboGuiApi
    Private oCompany As SAPbobsCOM.Company
    Private CrReportDocument As ProductionReport2N
    Private adoOledbConnection As OleDb.OleDbConnection
    Private adoOledbAdapter As OleDb.OleDbDataAdapter
    Private Dataset As DataSet
    Private sPath As String = IO.Directory.GetParent(Application.ExecutablePath).ToString
    Private oFDate, oTDate As String
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company, ByVal aFDate As String, ByVal aTDate As String)
        MyBase.New()
        Dim StrConnString As String = ""
        Dim SqlStr As String = ""


        SBO_Application = aSBO_Application
        oCompany = aCompany
        oFDate = aFDate
        oTDate = aTDate

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Dim oTempRs As SAPbobsCOM.Recordset
        Dim sqlserver, sqluid, sqlpwd As String
        oTempRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRs.DoQuery("Select * from [@PSSIT_OCON]")
        If oTempRs.RecordCount > 0 Then
            sqlserver = oTempRs.Fields.Item("U_SqlSer").Value
            sqluid = oTempRs.Fields.Item("U_SqlUID").Value
            sqlpwd = oTempRs.Fields.Item("U_SqlPwd").Value
            StrConnString = "Provider=SQLOLEDB;Server=" & oCompany.Server & ";Database=" & oCompany.CompanyDB & ";User ID=" & sqluid & ";Password=" & sqlpwd
        Else
            SBO_Application.SetStatusBarMessage("Sql login details are missing. Enter the details in the production setup", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Exit Sub
        End If
        StrConnString = "Provider=SQLOLEDB;Server=" & oCompany.Server & ";Database=" & oCompany.CompanyDB & ";User ID=" & sqluid & ";Password=" & sqlpwd


        '  StrConnString = "Provider=SQLOLEDB;Server=" & oCompany.Server & ";Database=" & oCompany.CompanyDB & ";User ID=" & oCompany.DbUserName & ";Password=sa2008"
       
        adoOledbConnection = New OleDb.OleDbConnection(StrConnString)
        '--------------
        'Dim objscrap As SAPbobsCOM.Recordset
        'Dim objrs, strSQL As String
        'objscrap = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'objRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'strSQL = "select * from Production_Report where "
        'objscrap.DoQuery(strSQL)
        'If Not objscrap.EoF Then
        '    Dim u_scrapqty As Decimal
        '    u_scrapqty = objscrap.Fields.Item("Incentive").Value
        'End If
        'Dim objincentive As SAPbobsCOM.Recordset
        'objincentive = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'strSQL = " select sum(doctotal) as Incentive from oinv "
        'strSQL = strSQL & " where u_category = 'incentive' and docstatus = 'O' and project = 'staff' and  docdate <= '" & objAddOn.objGenFunc.GetDateTimeValue(objForm.Items.Item("8").Specific.String).ToString("yyyy-MM-dd") & " ' "
        'objincentive.DoQuery(strSQL)
        'If Not objincentive.EoF Then
        '    Incentive = objincentive.Fields.Item("Incentive").Value
        'End If
        '------------
        If oFDate.Length > 0 And oTDate.Length > 0 Then
            SqlStr = "Production_Report '" & oFDate & "','" & oTDate & "'"
        End If


        adoOledbAdapter = New OleDb.OleDbDataAdapter(SqlStr, adoOledbConnection)

        Dataset = New DataSet
        adoOledbAdapter.Fill(Dataset, "Production_Report")
        CrReportDocument = New ProductionReport2N
        CrReportDocument.SetDataSource(Dataset)
        ProductionRptVwr.ReportSource = CrReportDocument
        ProductionRptVwr.Zoom(1)
        Me.WindowState = FormWindowState.Maximized

    End Sub
End Class