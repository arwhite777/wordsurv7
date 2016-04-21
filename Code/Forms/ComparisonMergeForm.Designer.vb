<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ComparisonMergeForm
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
        Me.cboComparison1 = New System.Windows.Forms.ComboBox
        Me.cboComparison2 = New System.Windows.Forms.ComboBox
        Me.txtName = New System.Windows.Forms.TextBox
        Me.lstPreviouslyMerged = New System.Windows.Forms.ListBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.stsStatusBar.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(195, 237)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(67, 23)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "Merge"
        '
        'stsStatusBar
        '
        Me.stsStatusBar.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.stsLabel1})
        Me.stsStatusBar.Location = New System.Drawing.Point(0, 263)
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
        Me.lblPrompt.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblPrompt.Location = New System.Drawing.Point(12, 9)
        Me.lblPrompt.Name = "lblPrompt"
        Me.lblPrompt.Size = New System.Drawing.Size(323, 25)
        Me.lblPrompt.TabIndex = 0
        Me.lblPrompt.Text = "For each Comparison you want to merge, select them from the drop down menus and t" & _
            "ype a name for the merged Comparison in the text box."
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(268, 237)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(67, 23)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Done"
        '
        'cboComparison1
        '
        Me.cboComparison1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboComparison1.FormattingEnabled = True
        Me.cboComparison1.Location = New System.Drawing.Point(12, 37)
        Me.cboComparison1.Name = "cboComparison1"
        Me.cboComparison1.Size = New System.Drawing.Size(156, 21)
        Me.cboComparison1.TabIndex = 5
        '
        'cboComparison2
        '
        Me.cboComparison2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboComparison2.FormattingEnabled = True
        Me.cboComparison2.Location = New System.Drawing.Point(179, 37)
        Me.cboComparison2.Name = "cboComparison2"
        Me.cboComparison2.Size = New System.Drawing.Size(156, 21)
        Me.cboComparison2.TabIndex = 6
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(99, 64)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(152, 20)
        Me.txtName.TabIndex = 7
        '
        'lstPreviouslyMerged
        '
        Me.lstPreviouslyMerged.FormattingEnabled = True
        Me.lstPreviouslyMerged.Location = New System.Drawing.Point(12, 110)
        Me.lstPreviouslyMerged.Name = "lstPreviouslyMerged"
        Me.lstPreviouslyMerged.Size = New System.Drawing.Size(323, 121)
        Me.lstPreviouslyMerged.TabIndex = 8
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 92)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(159, 13)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Previously merged Comparisons:"
        '
        'ComparisonMergeForm
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(347, 285)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lstPreviouslyMerged)
        Me.Controls.Add(Me.txtName)
        Me.Controls.Add(Me.cboComparison2)
        Me.Controls.Add(Me.cboComparison1)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.lblPrompt)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.stsStatusBar)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ComparisonMergeForm"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Merge Surveys"
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
    Friend WithEvents cboComparison1 As System.Windows.Forms.ComboBox
    Friend WithEvents cboComparison2 As System.Windows.Forms.ComboBox
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents lstPreviouslyMerged As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label

End Class
