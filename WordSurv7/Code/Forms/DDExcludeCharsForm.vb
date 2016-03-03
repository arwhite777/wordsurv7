Option Compare Text

Imports System.Windows.Forms

'This glorious form can be used any time the user desires to update the name of something or make a new thing that requires a name (dictionary, survey, variety, etc).
Public Class DDExcludeCharsForm
    Public Result As String = ""
    Public Sub New(ByVal fnt As Font, ByVal initialChars As String)
        Me.InitializeComponent()

        Me.txtInput.Font = fnt

        Me.txtInput.Text = initialChars

        Me.txtInput.Focus()
    End Sub
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.Result = Me.txtInput.Text
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
    End Sub
End Class
