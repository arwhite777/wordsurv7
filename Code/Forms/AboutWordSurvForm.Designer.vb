<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AboutWordSurvForm
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
        Me.btnOK = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.btnLore = New System.Windows.Forms.Button()
        Me.btnEasterEgg = New System.Windows.Forms.Button()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(176, 437)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 306)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(209, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "WordSurv Version 7.0 Beta release Luke b"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 338)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(95, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Advisor:  Art White"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 370)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(257, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Programmers: David Colgan, Art White, Keith Bauson"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(12, 402)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(398, 32)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Testers: Eliezer Rodriguez Frias, Nate White, Rachel Bird, Jonathan Schrock, Vale" & _
    "rie Newby"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.WordSurv7.My.Resources.Resources.WordSurvSplash
        Me.PictureBox1.Location = New System.Drawing.Point(12, 12)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(402, 280)
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'btnLore
        '
        Me.btnLore.Location = New System.Drawing.Point(335, 437)
        Me.btnLore.Name = "btnLore"
        Me.btnLore.Size = New System.Drawing.Size(75, 23)
        Me.btnLore.TabIndex = 6
        Me.btnLore.Text = "Lore"
        Me.btnLore.UseVisualStyleBackColor = True
        Me.btnLore.Visible = False
        '
        'btnEasterEgg
        '
        Me.btnEasterEgg.Location = New System.Drawing.Point(15, 437)
        Me.btnEasterEgg.Name = "btnEasterEgg"
        Me.btnEasterEgg.Size = New System.Drawing.Size(75, 23)
        Me.btnEasterEgg.TabIndex = 7
        Me.btnEasterEgg.Text = "Easter Egg"
        Me.btnEasterEgg.UseVisualStyleBackColor = True
        Me.btnEasterEgg.Visible = False
        '
        'AboutWordSurvForm
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(422, 471)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnEasterEgg)
        Me.Controls.Add(Me.btnLore)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.PictureBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "AboutWordSurvForm"
        Me.Padding = New System.Windows.Forms.Padding(9)
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "AboutWordSurvForm"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnLore As System.Windows.Forms.Button
    Friend WithEvents btnEasterEgg As System.Windows.Forms.Button

End Class
