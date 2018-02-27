<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmProdCostCrystalRptVwr
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
        Me.PordCostCrystalRptVwr = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SuspendLayout()
        '
        'PordCostCrystalRptVwr
        '
        Me.PordCostCrystalRptVwr.ActiveViewIndex = -1
        Me.PordCostCrystalRptVwr.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PordCostCrystalRptVwr.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PordCostCrystalRptVwr.Location = New System.Drawing.Point(0, 0)
        Me.PordCostCrystalRptVwr.Name = "PordCostCrystalRptVwr"
        Me.PordCostCrystalRptVwr.Size = New System.Drawing.Size(292, 266)
        Me.PordCostCrystalRptVwr.TabIndex = 0
        '
        'FrmProdCostCrystalRptVwr
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.PordCostCrystalRptVwr)
        Me.Name = "FrmProdCostCrystalRptVwr"
        Me.Text = "FrmProdCostCrystalRptVwr"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents PordCostCrystalRptVwr As CrystalDecisions.Windows.Forms.CrystalReportViewer
End Class
