<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EncodingChoice
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
        Me.cboEncodingSelector = New System.Windows.Forms.ComboBox
        Me.txtFilePreview = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnHelp = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'cboEncodingSelector
        '
        Me.cboEncodingSelector.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEncodingSelector.FormattingEnabled = True
        Me.cboEncodingSelector.Items.AddRange(New Object() {"IBM850", "IBM852", "IBM855", "IBM857", "Windows-1250", "Windows-1251", "Windows-1252", "Windows-1253", "Windows-1254", "Windows-1255", "Windows-1256", "Windows-1257", "Windows-1258", "UCS-2", "UTF-7", "UTF-8", "UTF-16", "UTF-32"})
        Me.cboEncodingSelector.Location = New System.Drawing.Point(229, 9)
        Me.cboEncodingSelector.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cboEncodingSelector.Name = "cboEncodingSelector"
        Me.cboEncodingSelector.Size = New System.Drawing.Size(252, 24)
        Me.cboEncodingSelector.TabIndex = 0
        '
        'txtFilePreview
        '
        Me.txtFilePreview.Location = New System.Drawing.Point(12, 55)
        Me.txtFilePreview.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtFilePreview.Multiline = True
        Me.txtFilePreview.Name = "txtFilePreview"
        Me.txtFilePreview.Size = New System.Drawing.Size(471, 509)
        Me.txtFilePreview.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(212, 43)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Select the encoding that makes your file readable."
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(407, 570)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 26)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(327, 570)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 26)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnHelp
        '
        Me.btnHelp.Location = New System.Drawing.Point(12, 570)
        Me.btnHelp.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnHelp.Name = "btnHelp"
        Me.btnHelp.Size = New System.Drawing.Size(75, 26)
        Me.btnHelp.TabIndex = 5
        Me.btnHelp.Text = "Help"
        Me.btnHelp.UseVisualStyleBackColor = True
        '
        'EncodingChoice
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(493, 606)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnHelp)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtFilePreview)
        Me.Controls.Add(Me.cboEncodingSelector)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Name = "EncodingChoice"
        Me.Text = "EncodingChoice"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cboEncodingSelector As System.Windows.Forms.ComboBox
    Friend WithEvents txtFilePreview As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnHelp As System.Windows.Forms.Button
End Class
