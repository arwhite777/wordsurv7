Public Class SearchForm
    Private frmWS As WordSurvForm
    Private prefs As Preferences
    Private data As WordSurvData


    Public Sub New(ByRef frmWordSurv As WordSurvForm, ByRef data As WordSurvData, ByRef prefs As Preferences)
        Me.InitializeComponent()
        Me.frmWS = frmWordSurv
        Me.data = data
        Me.prefs = prefs
    End Sub


    Private Function GetSearchGridType() As SearchType
        If Me.GetSearchGrid Is frmWS.grdGlossDictionary Then Return SearchType.DICTIONARY
        If Me.GetSearchGrid Is frmWS.grdVariety Then Return SearchType.SURVEY
        If Me.GetSearchGrid Is frmWS.grdComparisonGloss Then Return SearchType.COMPARISON_GLOSS
        If Me.GetSearchGrid Is frmWS.grdComparison Then Return SearchType.COMPARISON
        If Me.GetSearchGrid Is frmWS.grdCognateStrengths Then Return SearchType.COGNATE_STRENGTHS
        Return SearchType.NONE
    End Function

    Private Sub SetCurrentObject(ByVal objIndex As Integer)
        If Me.GetSearchGrid Is frmWS.grdGlossDictionary Then data.SetCurrentDictionary(objIndex)
        If Me.GetSearchGrid Is frmWS.grdVariety Then data.SetCurrentSurveysCurrentVariety(objIndex)
        If Me.GetSearchGrid Is frmWS.grdComparisonGloss Then Dim donothing As Integer = 0
        If Me.GetSearchGrid Is frmWS.grdComparison Then data.SetCurrentComparisonsCurrentGloss(objIndex)
        If Me.GetSearchGrid Is frmWS.grdCognateStrengths Then Dim donothing As Integer = 0
    End Sub

    Public Function GetCurrentObjIndex() As Integer
        If Me.GetSearchGrid Is frmWS.grdGlossDictionary Then Return frmWS.cboGlossDictionaries.SelectedIndex
        If Me.GetSearchGrid Is frmWS.grdVariety Then Return frmWS.cboVarieties.SelectedIndex
        If Me.GetSearchGrid Is frmWS.grdComparisonGloss Then Return 0
        If Me.GetSearchGrid Is frmWS.grdComparison Then Return frmWS.grdComparisonGloss.CurrentRow.Index
        If Me.GetSearchGrid Is frmWS.grdCognateStrengths Then Return 0
        Return -1
    End Function


    Private Sub btnSearchNext_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNext.Click
        If Me.GetSearchGridType() = SearchType.NONE Then Return
        If Me.GetSearchGrid.RowCount = 0 Or (Me.GetSearchGrid Is frmWS.grdGlossDictionary And Me.GetSearchGrid.RowCount = 1) Then Return
        If Me.GetSearchGrid.CurrentCell Is Nothing Then Return
        If DoLog Then Log.Add("Searched Next")
        Dim endAddress As CellAddress = data.SearchNext(Me.GetSearchGridType(), Me.txtFindWhat.Text, Me.GetCurrentObjIndex(), Me.GetSearchGrid.CurrentCell.RowIndex, Me.GetSearchGrid.CurrentCell.ColumnIndex)
        If endAddress IsNot Nothing Then
            Me.GetSearchGrid.CurrentCell = Me.GetSearchGrid.Rows(endAddress.RowIndex).Cells(endAddress.ColIndex)
            Me.SetCurrentObject(endAddress.ObjIndex)
        End If
        frmWS.RefreshBasedOnCurrentTab()
        Me.Focus()
    End Sub

    Private Sub btnSearchPrevious_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrevious.Click
        If Me.GetSearchGridType() = SearchType.NONE Then Return
        If Me.GetSearchGrid.RowCount = 0 Or (Me.GetSearchGrid Is frmWS.grdGlossDictionary And Me.GetSearchGrid.RowCount = 1) Then Return
        If Me.GetSearchGrid.CurrentCell Is Nothing Then Return
        If DoLog Then Log.Add("Searched Previous")
        Dim endAddress As CellAddress = data.SearchPrevious(Me.GetSearchGridType(), Me.txtFindWhat.Text, Me.GetCurrentObjIndex(), Me.GetSearchGrid.CurrentCell.RowIndex, Me.GetSearchGrid.CurrentCell.ColumnIndex)
        If endAddress IsNot Nothing Then
            Me.GetSearchGrid.CurrentCell = Me.GetSearchGrid.Rows(endAddress.RowIndex).Cells(endAddress.ColIndex)
            Me.SetCurrentObject(endAddress.ObjIndex)
        End If
        frmWS.RefreshBasedOnCurrentTab()
        Me.Focus()
    End Sub
    Private Sub btnReplace_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReplace.Click
        'MsgBox(Me.GetSearchGrid.VirtualMode)
        'If Me.txtReplaceWith.Text = "" Then Return 'AJW*** now allows replacement with nothing
        If Me.GetSearchGridType() = SearchType.NONE Or Me.GetSearchGridType() = SearchType.COMPARISON_GLOSS Or Me.GetSearchGridType() = SearchType.COGNATE_STRENGTHS Then Return
        If Me.GetSearchGrid.RowCount = 0 Or (Me.GetSearchGrid Is frmWS.grdGlossDictionary And Me.GetSearchGrid.RowCount = 1) Then Return
        If Me.GetSearchGrid.CurrentCell Is Nothing Then Return
        If DoLog Then Log.Add("Replaced " & Me.txtFindWhat.Text & " with " & Me.txtReplaceWith.Text)
        Dim endAddress As CellAddress = data.SearchReplace(Me.GetSearchGridType(), Me.txtFindWhat.Text, Me.txtReplaceWith.Text, Me.GetCurrentObjIndex(), Me.GetSearchGrid.CurrentCell.RowIndex, Me.GetSearchGrid.CurrentCell.ColumnIndex)
        'Can make this work for replace, but not replace all since each replace is handled internally rather than as it is here
        'Therefore, will disable the grouping clearing when transcription changes
        'If Me.GetSearchGridType() = SearchType.SURVEY Then
        '    data.UpdateTranscriptionValue(Me.GetSearchGrid.CurrentCell.RowIndex, Me.GetSearchGrid.CurrentCell.ColumnIndex, Me.txtReplaceWith.Text) 'AJW***Allows the system to undo the grouping just like a direct edit (push is not working here fo some reason)
        'End If
        If endAddress IsNot Nothing Then
            Me.GetSearchGrid.CurrentCell = Me.GetSearchGrid.Rows(endAddress.RowIndex).Cells(endAddress.ColIndex)
            'If Me.GetSearchGridType() = SearchType.SURVEY Then
            '    data.UpdateTranscriptionValue(Me.GetSearchGrid.CurrentCell.RowIndex, Me.GetSearchGrid.CurrentCell.ColumnIndex, Me.txtReplaceWith.Text) 'AJW***Allows the system to undo the grouping just like a direct edit (push is not working here fo some reason)
            'End If
            Me.SetCurrentObject(endAddress.ObjIndex)
        End If
        frmWS.RefreshBasedOnCurrentTab()
        StoreForUndo(data, prefs)
        Me.Focus()
    End Sub

    Private Sub btnReplaceAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReplaceAll.Click
        'If Me.txtReplaceWith.Text = "" Then Return 'AJW*** now allows replacement with nothing
        Dim result As MsgBoxResult = MsgBox("Warning:  This action will replace all occurrences located anywhere in this grid.  If you wish to limit the scope of the replacement, use the next and replace buttons to walk through each individual change.", MsgBoxStyle.OkCancel, "Replace ALL Occurrences in This Grid")
        If result = MsgBoxResult.Ok Then
            If Me.GetSearchGridType() = SearchType.NONE Or Me.GetSearchGridType() = SearchType.COMPARISON_GLOSS Or Me.GetSearchGridType() = SearchType.COGNATE_STRENGTHS Then Return
            If Me.GetSearchGrid.RowCount = 0 Or (Me.GetSearchGrid Is frmWS.grdGlossDictionary And Me.GetSearchGrid.RowCount = 1) Then Return
            If Me.GetSearchGrid.CurrentCell Is Nothing Then Return
            If DoLog Then Log.Add("Replaced all")
            frmWS.setStatusWarning("Replaced " & data.SearchReplaceAll(Me.GetSearchGridType(), Me.txtFindWhat.Text, Me.txtReplaceWith.Text, Me.GetCurrentObjIndex(), Me.GetSearchGrid.CurrentCell.RowIndex, Me.GetSearchGrid.CurrentCell.ColumnIndex).ToString & " occurences.", True)
            frmWS.RefreshBasedOnCurrentTab()
            StoreForUndo(data, prefs)
            Me.Focus()
        End If
    End Sub


    Private Function GetSearchGrid() As DataGridView
        If frmWS.LastActivatedGrid IsNot Nothing Then Return frmWS.LastActivatedGrid
        'If there is no grid focused, pick the most likely one on the given tab.
        If frmWS.tabWordSurv.SelectedIndex = 0 Then Return frmWS.grdVariety
        If frmWS.tabWordSurv.SelectedIndex = 1 Then Return frmWS.grdComparison
        If frmWS.tabWordSurv.SelectedIndex = 5 Then Return frmWS.grdCognateStrengths
        Return Nothing
    End Function


    Private Sub SearchForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyValue = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub SearchForm_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        frmWS.DoingSearchAndReplace = False
        If DoLog Then Log.Add("Closed the search box")
    End Sub
End Class