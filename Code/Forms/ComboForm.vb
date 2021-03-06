﻿Option Compare Text

Imports System.Windows.Forms

'This glorious form can be used any time the user desires to update the name of something or make a new thing that requires a name (dictionary, survey, variety, etc).
Public Class ComboForm
    Public Result As String = ""

    Public Sub New(ByVal titleText As String, ByVal promptText As String, ByVal comboBoxContent As String())
        Me.InitializeComponent()
        Me.Text = titleText
        Me.lblPrompt.Text = promptText
        For Each str As String In comboBoxContent
            Me.cboSelector.Items.Add(str)
        Next
        Me.cboSelector.SelectedIndex = 0
        Me.cboSelector.Focus()
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
    End Sub
End Class
