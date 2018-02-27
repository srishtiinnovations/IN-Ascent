<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmProdCostRptVwr
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.ProdCostRptVwr = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SuspendLayout()
        '
        'ProdCostRptVwr
        '
        Me.ProdCostRptVwr.ActiveViewIndex = -1
        Me.ProdCostRptVwr.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ProdCostRptVwr.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ProdCostRptVwr.Location = New System.Drawing.Point(0, 0)
        Me.ProdCostRptVwr.Name = "ProdCostRptVwr"
        Me.ProdCostRptVwr.SelectionFormula = ""
        Me.ProdCostRptVwr.Size = New System.Drawing.Size(667, 266)
        Me.ProdCostRptVwr.TabIndex = 0
        Me.ProdCostRptVwr.ViewTimeSelectionFormula = ""
        '
        'FrmProdCostRptVwr
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(667, 266)
        Me.Controls.Add(Me.ProdCostRptVwr)
        Me.Name = "FrmProdCostRptVwr"
        Me.Text = "FrmProdCostRptVwr"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ProdCostRptVwr As CrystalDecisions.Windows.Forms.CrystalReportViewer
End Class
