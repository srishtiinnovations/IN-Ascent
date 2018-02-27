Public Class FrmTest
    Inherits System.Windows.Forms.Form

#Region "Variable Declaration"
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private SboGuiApi As New SAPbouiCOM.SboGuiApi
    Private oCompany As SAPbobsCOM.Company
    Private oFormName As String
    Private FrmProdCostCrystalRptVwr As FrmProdCostCrystalRptVwr
    Private FrmProductionRptVwr As FrmProductionRptVwr
    Private FrmProcessSheetRptVwr As FrmProcessSheetRptVwr
    Private oFPordNo, oTPordNo, oStatus, oFDate, oTDate As String
#End Region
    Public Sub New(ByVal aSBO_Application As SAPbouiCOM.Application, ByVal aCompany As SAPbobsCOM.Company, ByVal FormName As String, Optional ByVal aFPordNo As String = "", Optional ByVal aTPordNo As String = "", Optional ByVal aStatus As String = "", Optional ByVal aFDate As String = "", Optional ByVal aTDate As String = "")
        MyBase.New()
        SBO_Application = aSBO_Application
        oCompany = aCompany
        oFormName = FormName
        oFPordNo = aFPordNo
        oTPordNo = aTPordNo
        oStatus = aStatus
        '  oOthers = aOthers
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
    End Sub
    Public Sub StartThread()
        Try
            Dim run As Boolean
            run = True
            Me.Show()
            While (run)
                Application.DoEvents()
                Threading.Thread.Sleep(1)
            End While
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub FrmTest_Load_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim MyProcs As Process()
        MyProcs = Process.GetProcessesByName("SAP Business One")
        Me.Close()
        If MyProcs.Length <> 0 Then
            For i As Integer = 0 To MyProcs.Length - 1
                Dim MyWindow As New WindowWrapper(MyProcs(i).MainWindowHandle)
                Select Case oFormName
                    Case "ProcessSheetReport"
                        FrmProcessSheetRptVwr = New FrmProcessSheetRptVwr(SBO_Application, oCompany, oFPordNo)
                        FrmProcessSheetRptVwr.TopMost = True
                        FrmProcessSheetRptVwr.Show(MyWindow)
                    Case "ProdCostCrystalRpt"
                        FrmProdCostCrystalRptVwr = New FrmProdCostCrystalRptVwr(SBO_Application, oCompany, oFPordNo, oTPordNo, oStatus)
                        FrmProdCostCrystalRptVwr.TopMost = True
                        FrmProdCostCrystalRptVwr.Show(MyWindow)
                    Case "ProdReport"
                        FrmProductionRptVwr = New FrmProductionRptVwr(SBO_Application, oCompany, oFPordNo, oTPordNo)
                        FrmProductionRptVwr.TopMost = True
                        FrmProductionRptVwr.Show(MyWindow)
                End Select
            Next
        End If

    End Sub

End Class