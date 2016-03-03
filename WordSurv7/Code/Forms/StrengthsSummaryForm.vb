Imports System.Windows.Forms

Public Class StrengthsSummaryForm

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Public Sub fillSummaryGrid(ByRef strengthCounts As StrengthCountsSummary)
        Me.grdSummaryDisplay.Rows.Clear()
        Me.grdSummaryDisplay.Columns.Clear()

        If strengthCounts Is Nothing Then Return
        Me.grdSummaryDisplay.Columns.Add("Range", "Range")
        Me.grdSummaryDisplay.Columns.Add("Transcription Pair Counts", "Transcription Pair Counts")
        Dim newRow As DataGridViewRow = Me.grdSummaryDisplay.Rows(Me.grdSummaryDisplay.Rows.Add())
        newRow.Cells("Range").Value = "Strength = 1.00"
        newRow.Cells("Transcription Pair Counts").Value = strengthCounts.Eq1

        newRow = Me.grdSummaryDisplay.Rows(Me.grdSummaryDisplay.Rows.Add())
        newRow.Cells("Range").Value = "0.75 <= Strength < 1.00"
        newRow.Cells("Transcription Pair Counts").Value = strengthCounts.Gte75lt1

        newRow = Me.grdSummaryDisplay.Rows(Me.grdSummaryDisplay.Rows.Add())
        newRow.Cells("Range").Value = "0.50 <= Strength < 0.75"
        newRow.Cells("Transcription Pair Counts").Value = strengthCounts.Gte50lt75

        newRow = Me.grdSummaryDisplay.Rows(Me.grdSummaryDisplay.Rows.Add())
        newRow.Cells("Range").Value = "0.25 <= Strength < 0.50"
        newRow.Cells("Transcription Pair Counts").Value = strengthCounts.Gte25lt50

        newRow = Me.grdSummaryDisplay.Rows(Me.grdSummaryDisplay.Rows.Add())
        newRow.Cells("Range").Value = "0.00 <= Strength < 0.25"
        newRow.Cells("Transcription Pair Counts").Value = strengthCounts.Gte0lt25

        newRow = Me.grdSummaryDisplay.Rows(Me.grdSummaryDisplay.Rows.Add())
        newRow.Cells("Range").Value = "-0.25 <= Strength < 0.00"
        newRow.Cells("Transcription Pair Counts").Value = strengthCounts.Gten25lt0

        newRow = Me.grdSummaryDisplay.Rows(Me.grdSummaryDisplay.Rows.Add())
        newRow.Cells("Range").Value = "-0.50 <= Strength < -0.25"
        newRow.Cells("Transcription Pair Counts").Value = strengthCounts.Gten50ltn25

        newRow = Me.grdSummaryDisplay.Rows(Me.grdSummaryDisplay.Rows.Add())
        newRow.Cells("Range").Value = "-0.75 <= Strength < -0.50"
        newRow.Cells("Transcription Pair Counts").Value = strengthCounts.Gten75ltn50

        newRow = Me.grdSummaryDisplay.Rows(Me.grdSummaryDisplay.Rows.Add())
        newRow.Cells("Range").Value = "-1.00 < Strength < -0.75"
        newRow.Cells("Transcription Pair Counts").Value = strengthCounts.Gtn1ltn75

        newRow = Me.grdSummaryDisplay.Rows(Me.grdSummaryDisplay.Rows.Add())
        newRow.Cells("Range").Value = "Strength = -1.00"
        newRow.Cells("Transcription Pair Counts").Value = strengthCounts.Eqn1
    End Sub

    Private Sub StrengthsSummaryForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyValue = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub StrengthsSummaryForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.grdSummaryDisplay.AutoResizeColumns()
            Me.grdSummaryDisplay.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders)
            Me.grdSummaryDisplay.Columns("Transcription Pair Counts").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Catch ex As Exception
        End Try
    End Sub


End Class
