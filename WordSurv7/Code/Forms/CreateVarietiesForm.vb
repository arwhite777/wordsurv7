Imports System.Windows.Forms
Imports System.Text.RegularExpressions

Public Class CreateVarietiesForm
    Public VarietyNames As String = ""

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.VarietyNames = Me.txtNewVarietyNames.Text
        Me.DialogResult = Windows.Forms.DialogResult.OK
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
    End Sub

    Private Sub txtNewVarietyNames_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNewVarietyNames.TextChanged
        Dim isError As Boolean = False
        Dim lines As String() = Split(Me.txtNewVarietyNames.Text, vbCrLf)

        Dim notEmpty As Boolean = False
        For Each line As String In lines
            If line.Trim(" "c).Length <> 0 Then
                notEmpty = True
            End If
        Next
        If Not notEmpty Then
            Me.setStatusWarning("You must include at least one variety.")
            isError = True
        End If

        For i As Integer = 0 To lines.Length - 1
            For j As Integer = i + 1 To lines.Length - 1
                If (lines(i).Trim(" "c) = lines(j).Trim(" "c) And lines(i).Trim(" "c) <> "" And lines(j).Trim(" "c) <> "") Then
                    Me.setStatusWarning("Duplicate Variety names are not allowed in the same Survey.")
                    isError = True
                    GoTo Done
                End If
            Next
        Next
Done:

        If isError Then
            Me.btnOK.Enabled = False
        Else
            Me.btnOK.Enabled = True
            Me.clearStatusBar()
        End If

    End Sub

    Public Sub setStatusWarning(ByVal msg As String)
        Me.stsLabel1.Text = msg
        Me.stsStatusBar.BackColor = INVALID_COLOR
        Beep()
    End Sub
    Public Sub clearStatusBar()
        Me.stsLabel1.Text = ""
        Me.stsStatusBar.BackColor = Color.Empty
    End Sub

    Private Sub CreateVarietiesForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.btnOK.Enabled = False
    End Sub
End Class
