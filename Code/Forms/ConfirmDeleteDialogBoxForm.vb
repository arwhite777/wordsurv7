Imports System.Windows.Forms

Public Class ConfirmDeleteDialogBoxForm

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub ConfirmDeleteDialogBoxForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.picExclamation.Image = SystemIcons.Exclamation.ToBitmap
    End Sub
End Class
