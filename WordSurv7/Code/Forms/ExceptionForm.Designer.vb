<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ExceptionForm
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
        Me.OK_Button = New System.Windows.Forms.Button
        Me.lblQuote = New System.Windows.Forms.Label
        Me.txtExceptionText = New System.Windows.Forms.TextBox
        Me.btnCopyToClipboard = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.OK_Button.Location = New System.Drawing.Point(510, 569)
        Me.OK_Button.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(89, 28)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "OK"
        '
        'lblQuote
        '
        Me.lblQuote.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblQuote.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuote.Location = New System.Drawing.Point(17, 16)
        Me.lblQuote.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblQuote.Name = "lblQuote"
        Me.lblQuote.Size = New System.Drawing.Size(837, 153)
        Me.lblQuote.TabIndex = 1
        Me.lblQuote.Text = "There is no bug"
        '
        'txtExceptionText
        '
        Me.txtExceptionText.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtExceptionText.Location = New System.Drawing.Point(16, 172)
        Me.txtExceptionText.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtExceptionText.Multiline = True
        Me.txtExceptionText.Name = "txtExceptionText"
        Me.txtExceptionText.Size = New System.Drawing.Size(837, 389)
        Me.txtExceptionText.TabIndex = 2
        '
        'btnCopyToClipboard
        '
        Me.btnCopyToClipboard.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCopyToClipboard.Location = New System.Drawing.Point(607, 569)
        Me.btnCopyToClipboard.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCopyToClipboard.Name = "btnCopyToClipboard"
        Me.btnCopyToClipboard.Size = New System.Drawing.Size(251, 28)
        Me.btnCopyToClipboard.TabIndex = 3
        Me.btnCopyToClipboard.Text = "Copy Error Message to Clipboard"
        '
        'ExceptionForm
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(871, 620)
        Me.Controls.Add(Me.btnCopyToClipboard)
        Me.Controls.Add(Me.OK_Button)
        Me.Controls.Add(Me.txtExceptionText)
        Me.Controls.Add(Me.lblQuote)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ExceptionForm"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Flagrant System Error"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents lblQuote As System.Windows.Forms.Label
    Friend WithEvents txtExceptionText As System.Windows.Forms.TextBox
    Friend WithEvents btnCopyToClipboard As System.Windows.Forms.Button

End Class
