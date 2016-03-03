Option Compare Text

Imports System.Windows.Forms

Public Class MultilineInputForm
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
    End Sub

    Private Sub MultilineInputForm_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Me.txtInput.Focus()
        Me.txtInput.DeselectAll()
    End Sub

    Private Sub MultilineInputForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.txtInput.Focus()
        Me.txtInput.DeselectAll()
    End Sub
End Class
