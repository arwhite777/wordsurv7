<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SurveySelectForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnOK = New System.Windows.Forms.Button
        Me.stsStatusBar = New System.Windows.Forms.StatusStrip
        Me.stsLabel1 = New System.Windows.Forms.ToolStripStatusLabel
        Me.lblPrompt = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.cboSurveySelection = New System.Windows.Forms.ComboBox
        Me.stsStatusBar.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(195, 89)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(67, 23)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "OK"
        '
        'stsStatusBar
        '
        Me.stsStatusBar.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.stsLabel1})
        Me.stsStatusBar.Location = New System.Drawing.Point(0, 115)
        Me.stsStatusBar.Name = "stsStatusBar"
        Me.stsStatusBar.Size = New System.Drawing.Size(347, 22)
        Me.stsStatusBar.SizingGrip = False
        Me.stsStatusBar.TabIndex = 4
        Me.stsStatusBar.Text = "StatusStrip1"
        '
        'stsLabel1
        '
        Me.stsLabel1.Name = "stsLabel1"
        Me.stsLabel1.Size = New System.Drawing.Size(0, 17)
        '
        'lblPrompt
        '
        Me.lblPrompt.Location = New System.Drawing.Point(12, 9)
        Me.lblPrompt.Name = "lblPrompt"
        Me.lblPrompt.Size = New System.Drawing.Size(323, 49)
        Me.lblPrompt.TabIndex = 0
        Me.lblPrompt.Text = "Prompt."
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(268, 89)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(67, 23)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        '
        'cboSurveySelection
        '
        Me.cboSurveySelection.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSurveySelection.FormattingEnabled = True
        Me.cboSurveySelection.Location = New System.Drawing.Point(12, 61)
        Me.cboSurveySelection.Name = "cboSurveySelection"
        Me.cboSurveySelection.Size = New System.Drawing.Size(323, 21)
        Me.cboSurveySelection.TabIndex = 5
        '
        'SurveySelectForm
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(347, 137)
        Me.ControlBox = False
        Me.Controls.Add(Me.cboSurveySelection)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.lblPrompt)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.stsStatusBar)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SurveySelectForm"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Input Title"
        Me.stsStatusBar.ResumeLayout(False)
        Me.stsStatusBar.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents stsStatusBar As System.Windows.Forms.StatusStrip
    Friend WithEvents stsLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblPrompt As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents cboSurveySelection As System.Windows.Forms.ComboBox

End Class
