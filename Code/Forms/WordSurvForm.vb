
Public Class WordSurvForm
    'Saves preferences not bound to a specific database.
    Private prefs As Preferences

    Private ActiveGrid As DataGridView

    Public DoingSearchAndReplace As Boolean = False
    Public LastActivatedGrid As DataGridView = Nothing
    Private frmSearch As SearchForm = Nothing

    Private OperationInProgress As Boolean = False

    'Add all grids in the form to this array so that searching, column highlighting, etc, will work for all grids.
    Private grids As DataGridView()

    'Code in here only does things directly associated with the form.  It should NEVER touch the data objects directly, but only access them through the data object's interface.
    Private rowHeaderCellSize As Integer = 0

    Private data As New WordSurvData 'A reference to the ultimate data object ever. 
    Private newMiddleRow As Boolean = False 'A flag to tell us when we have inserted a row using the menu or shortcut.
    Private newBottomRow As Boolean = False 'A flag to tell us when the user inserted a row using the bottom thingy.
    Private copiedIndices As List(Of Integer) = Nothing 'A holder for copy and paste action.
    Private sourceDictIndex As Integer = Nothing

    Public KillPipeKeyFlag As Boolean = False
    'Entry point for this form
    Private Sub WordSurvForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        'Close the program if it is already running.
        If Process.GetProcessesByName(Process.GetCurrentProcess.ProcessName).Length > 1 Then
            MsgBox("WordSurv is already running.  If the other instance is not responding, use the Task Manager to close it.", MsgBoxStyle.Information, "WordSurv Already Running.")
            Me.Close()
        End If

        Dim currentDomain As AppDomain = AppDomain.CurrentDomain
        AddHandler currentDomain.UnhandledException, AddressOf MYExceptionHandler
        AddHandler Application.ThreadException, AddressOf MYThreadHandler

        Dim grids As DataGridView() = {Me.grdGlossDictionary, Me.grdVariety, Me.grdComparisonGloss, Me.grdComparison, Me.grdComparisonAnalysis, Me.grdDegreesOfDifference, Me.grdPhonoStats, Me.grdPhoneCorr, Me.grdCognateStrengths}
        Me.grids = grids
        For Each grd As DataGridView In Me.grids
            AddHandler grd.ColumnHeaderMouseClick, AddressOf ColumnHeaderClicked
            AddHandler grd.RowHeaderMouseClick, AddressOf RowHeaderClicked
            AddHandler grd.GotFocus, AddressOf Me.GridGotFocus
            'AddHandler grd.MouseUp, AddressOf Me.grdMouseUp
            AddHandler grd.CellBeginEdit, AddressOf gridCellBeginEditHandleMenus
            AddHandler grd.CellEndEdit, AddressOf gridCellEndEditHandleMenus
            AddHandler grd.RowPostPaint, AddressOf gridShowRowIndex
        Next

        'We still have to manually add the columns to the grid.  The virtual mode thing isn't that smart.
        Me.grdGlossDictionary.Columns.Add("Name", "Name")
        Me.grdGlossDictionary.Columns.Add("Name2", "Name2")
        Me.grdGlossDictionary.Columns.Add("PartOfSpeech", "POS")
        Me.grdGlossDictionary.Columns.Add("FieldTip", "Field Tip")
        Me.grdGlossDictionary.Columns.Add("Comments", "Comments")
        Me.grdGlossDictionary.RowHeadersVisible = True


        For Each col As DataGridViewColumn In Me.grdGlossDictionary.Columns
            col.SortMode = DataGridViewColumnSortMode.Programmatic
        Next
        Me.grdGlossDictionary.SelectionMode = DataGridViewSelectionMode.ColumnHeaderSelect
        GlossDictionaryGridColCount = Me.grdGlossDictionary.Columns.Count


        Me.grdVariety.Columns.Add("Name", "Gloss")
        Me.grdVariety.Columns("Name").ReadOnly = True
        Me.grdVariety.Columns("Name").DefaultCellStyle.BackColor = NON_EDITABLE_COLOR
        Me.grdVariety.Columns.Add("Transcription", "Transcription")
        Me.grdVariety.Columns.Add("PluralFrame", "Plural/Frame")
        Me.grdVariety.Columns.Add("Notes", "Notes")
        VarietyGridColCount = Me.grdVariety.Columns.Count
        Me.grdVariety.RowHeadersVisible = True


        For Each col As DataGridViewColumn In Me.grdVariety.Columns
            col.SortMode = DataGridViewColumnSortMode.Programmatic
        Next
        Me.grdVariety.SelectionMode = DataGridViewSelectionMode.ColumnHeaderSelect

        Me.grdComparisonGloss.Columns.Add("Name", "Gloss")
        Me.grdComparisonGloss.Columns("Name").ReadOnly = True
        Me.grdComparisonGloss.Columns("Name").DefaultCellStyle.BackColor = NON_EDITABLE_COLOR
        ComparisonGlossGridColCount = Me.grdComparisonGloss.Columns.Count
        Me.grdComparisonGloss.RowHeadersVisible = True


        Me.grdComparison.Columns.Add("Variety", "Variety")
        Me.grdComparison.Columns("Variety").ReadOnly = True
        Me.grdComparison.Columns("Variety").DefaultCellStyle.BackColor = NON_EDITABLE_COLOR
        Me.grdComparison.Columns.Add("Transcription", "Transcription")
        Me.grdComparison.Columns("Transcription").ReadOnly = True
        Me.grdComparison.Columns("Transcription").DefaultCellStyle.BackColor = NON_EDITABLE_COLOR
        Me.grdComparison.Columns.Add("PluralFrame", "Plural/Frame")
        Me.grdComparison.Columns("PluralFrame").ReadOnly = True
        Me.grdComparison.Columns("PluralFrame").DefaultCellStyle.BackColor = NON_EDITABLE_COLOR
        Me.grdComparison.Columns.Add("AlignedRendering", "Aligned")
        Me.grdComparison.Columns.Add("Grouping", "Grouping")
        Me.grdComparison.Columns.Add("Notes", "Notes")
        Me.grdComparison.Columns.Add("Exclude", "Exclude")
        Me.grdComparison.RowHeadersVisible = True

        For Each col As DataGridViewColumn In Me.grdComparison.Columns
            col.SortMode = DataGridViewColumnSortMode.Programmatic
        Next
        Me.grdComparison.SelectionMode = DataGridViewSelectionMode.ColumnHeaderSelect
        ComparisonGridColCount = Me.grdComparison.Columns.Count


        Me.grdComparisonAnalysis.ReadOnly = True
        Me.grdComparisonAnalysis.DefaultCellStyle.BackColor = NON_EDITABLE_COLOR
        Me.grdComparisonAnalysis.RowHeadersVisible = True

        Me.grdCognateStrengths.Columns.Add("Gloss", "Gloss")
        Me.grdCognateStrengths.Columns.Add("Form 1", "Form 1")
        Me.grdCognateStrengths.Columns.Add("Form 2", "Form 2")
        Me.grdCognateStrengths.Columns.Add("Strength", "Strength")
        CognateStrengthsGridColCount = Me.grdCognateStrengths.Columns.Count
        Me.grdCognateStrengths.RowHeadersVisible = True



        Me.prefs = New Preferences

        'Apply preferences.
        Me.grdGlossDictionary.Columns("Name").Width = prefs.GlossDictionaryGridNameWidth
        Me.grdGlossDictionary.Columns("Name2").Width = prefs.GlossDictionaryGridName2Width
        Me.grdGlossDictionary.Columns("PartOfSpeech").Width = prefs.GlossDictionaryGridPartOfSpeechWidth
        Me.grdGlossDictionary.Columns("FieldTip").Width = prefs.GlossDictionaryGridFieldTipWidth
        Me.grdGlossDictionary.Columns("Comments").Width = prefs.GlossDictionaryGridCommentsWidth

        Me.grdVariety.Columns("Name").Width = prefs.VarietyGridNameWidth
        Me.grdVariety.Columns("Transcription").Width = prefs.VarietyGridTranscriptionWidth
        Me.grdVariety.Columns("PluralFrame").Width = prefs.VarietyGridPluralFrameWidth
        Me.grdVariety.Columns("Notes").Width = prefs.VarietyGridNotesWidth

        Me.grdComparisonGloss.Columns("Name").Width = prefs.ComparisonGlossGridNameWidth

        Me.grdComparison.Columns("Variety").Width = prefs.ComparisonGridVarietyWidth
        Me.grdComparison.Columns("Transcription").Width = prefs.ComparisonGridTranscriptionWidth
        Me.grdComparison.Columns("PluralFrame").Width = prefs.ComparisonGridPluralFrameWidth
        Me.grdComparison.Columns("AlignedRendering").Width = prefs.ComparisonGridAlignedRenderingWidth
        Me.grdComparison.Columns("Grouping").Width = prefs.ComparisonGridGroupingWidth
        Me.grdComparison.Columns("Notes").Width = prefs.ComparisonGridNotesWidth
        Me.grdComparison.Columns("Exclude").Width = prefs.ComparisonGridExcludeWidth

        Me.grdCognateStrengths.Columns("Gloss").Width = prefs.CognateStrengthsGridGlossWidth
        Me.grdCognateStrengths.Columns("Form 1").Width = prefs.CognateStrengthsGridForm1Width
        Me.grdCognateStrengths.Columns("Form 2").Width = prefs.CognateStrengthsGridForm2Width
        Me.grdCognateStrengths.Columns("Strength").Width = prefs.CognateStrengthsGridStrengthWidth


        Try
            Me.Width = prefs.ApplicationWidth
            Me.Height = prefs.ApplicationHeight
            Me.Left = prefs.ApplicationX
            Me.Top = prefs.ApplicationY
            If prefs.ApplicationIsMaximized Then Me.WindowState = FormWindowState.Maximized
        Catch ex As Exception
        End Try

        Me.tabWordSurv.Visible = False
        Dim curTabIndex As Integer = prefs.CurrentTab
        For i As Integer = 0 To Me.tabWordSurv.TabCount - 1
            Application.DoEvents()
            Me.tabWordSurv.SelectedIndex = i
            Application.DoEvents()
        Next
        Me.tabWordSurv.SelectedIndex = curTabIndex
        Me.tabWordSurv.Visible = True


        'open file from command line if it is provided
        If My.Application.CommandLineArgs.Count <> 0 Then
            Me.Open(My.Application.CommandLineArgs(0))
        Else
            If prefs.LastOpenedDatabase <> "" Then Me.Open(prefs.LastOpenedDatabase)
        End If
        If data.IsEmpty() Then Me.RefreshMenus()

        InitUndo(data, prefs)

        Me.RefreshFonts()


        Try
            Me.splTab1A.SplitterDistance = prefs.DictionaryPaneWidth
            Me.splTab1B.SplitterDistance = prefs.SurveyPaneWidth
        Catch ex As Exception
        End Try
        Try
            Me.splComparisons.SplitterDistance = prefs.ComparisonPaneWidth
        Catch ex As Exception
        End Try
        Try
            Me.splCOMPASS.SplitterDistance = prefs.COMPASSPaneWidth
        Catch ex As Exception
        End Try

        AddHandler splTab1A.SplitterMoved, AddressOf splTab1A_SplitterMoved
        AddHandler splTab1B.SplitterMoved, AddressOf splTab1B_SplitterMoved
        AddHandler splComparisons.SplitterMoved, AddressOf splComparisons_SplitterMoved
        AddHandler splCOMPASS.SplitterMoved, AddressOf splCOMPASS_SplitterMoved

        AddHandler Me.Move, AddressOf WordSurvForm_Move
        AddHandler Me.SizeChanged, AddressOf WordSurvForm_SizeChanged

        'Hilarious edge case: if someone uses two monitors and puts the wordsurv window on the right monitor and then
        'goes back to one monitor, on startup the wordsurv window will still be off the screen on the second monitor,
        'so it will look like the program isn't starting.  Therefore check to make sure that the window is not off the screen,
        'and if it is move it back.
        'If Me.Left > Screen.PrimaryScreen.Bounds.Width Then Me.Left = 0
        'If Me.Top > Screen.PrimaryScreen.Bounds.Height Then Me.Top = 0
        '^This solution didn't seem to work.  Someone should fix this someday.

        Me.ActiveGrid = Me.GetDefaultGridForCurrentTab()
        Me.ActiveGrid.Focus()

        Log = New List(Of String)
        If DoLog Then Log.Add("Form Loaded Successfully")
        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        'numberRows(grdGlossDictionary)
    End Sub

    Private Sub gridShowRowIndex(sender As Object, e As DataGridViewRowPostPaintEventArgs)
        Dim dg As DataGridView = DirectCast(sender, DataGridView)
        Dim row As DataGridViewRow = dg.Rows(e.RowIndex)
        Dim str As String = (row.Index + 1).ToString
        If Not row.IsNewRow And row.HeaderCell.Value <> str Then
            row.HeaderCell.Value = str
        End If
        Dim size As Integer = (dg.Rows.Count - 1) \ 10
        If size <> Me.rowHeaderCellSize Then
            Me.rowHeaderCellSize = size
            dg.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders)
        End If
    End Sub

#Region "File Management"
    Private Sub mnuNewDatabase_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNewDatabase.Click
        Me.CommitGrids()

        Dim filename As String = GetNewDatabaseName()
        If filename <> "" Then
            data = New WordSurvData
            data.filename = filename
            Me.save()

            prefs.LastOpenedDatabase = data.filename
            Me.BumpUpPreviousDatabases()
            Me.FillRecentDatabasesList()

            Me.RefreshBasedOnCurrentTab()
            Me.RefreshMenus()
            ClearUndo(data, prefs)
            If DoLog Then Log.Add("Created New Database " & data.filename)
        End If
    End Sub
    Private Sub mnuSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSave.Click
        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Me.CommitGrids()
        Me.save()
        If DoLog Then Log.Add("User Saved")
        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Public Sub mnuOpenDatabase_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOpenDatabase.Click
        openDatabaseFile(data, prefs, "")

        'Me.CommitGrids()
        'Dim filename As String = GetDatabaseNameToOpen()
        'If filename <> "" Then
        '    Me.Open(filename)
        'End If
        'ClearUndo(data, prefs)
        'Me.refreshDictionaryPane()
        'Me.refreshSurveyPane()
        'Me.RefreshMenus()
        'If DoLog Then Log.Add("Opened Database " & data.filename)
    End Sub
    Public Sub openDatabaseFile(ByVal data As WordSurvData, ByVal prefs As Preferences, ByVal filename As String)
        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Me.CommitGrids()
        If filename = "" Then
            filename = GetDatabaseNameToOpen()
            If filename <> "" Then
                Me.Open(filename)
            End If
        End If
        ClearUndo(data, prefs)
        Me.refreshDictionaryPane()
        Me.refreshSurveyPane()
        Me.RefreshMenus()
        HasNotSaved = False
        If DoLog Then Log.Add("Opened Database " & data.filename)
        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub mnuImportVersion2_5Database_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuImportVersion2_5Database.Click
        Me.CommitGrids()
        Dim dbFileName As String = GetWordSurv2_5DatabaseToImport()
        If dbFileName <> "" Then
            Try
                Dim dbFileNameBase As String = dbFileName.Substring(0, InStrRev(dbFileName, ".") - 1)
                'If System.IO.File.Exists(dbFileNameBase & ".cat") = True Then
                WordSurvData.ImportWordSurv2_5Database(dbFileNameBase)
                Me.Open(dbFileNameBase & ".wsv")

                Me.RefreshBasedOnCurrentTab()
                Me.RefreshMenus()
                ClearUndo(data, prefs)
                If DoLog Then Log.Add("Imported 2.5 Database " & data.filename)
                'Else
                'MsgBox("Unable to perform WS2.5 conversion to WS7 since no " & dbFileNameBase & ".cat" & " exists!!", MsgBoxStyle.OkOnly, "No .cat file exists for this WS2.5 database")
                'End If
            Catch ex As OperationCanceledException
            Catch ex As Exception
                MsgBox("Import failed: " & ex.Message, MsgBoxStyle.Exclamation, "WordSurv 2.5 Import Failed")
            End Try
        End If
    End Sub
    Private Function GetWordSurv2_5DatabaseToImport() As String

        'Display a dialog and let them choose a database.
        Dim frmFileDialog As New OpenFileDialog

        frmFileDialog.Title = "Import WordSurv 2.5 File"
        frmFileDialog.DefaultExt = "db"
        frmFileDialog.Filter = "WordSurv 2.5 Database File (.db) | *.db"

        If frmFileDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Return frmFileDialog.FileName
        Else
            Return ""
        End If
    End Function
    Private Function GetDatabaseNameToOpen() As String
        'Display a dialog and let them choose a database.
        Dim frmFileDialog As New OpenFileDialog

        frmFileDialog.Title = "Open Existing WordSurv 7.0 File"
        frmFileDialog.DefaultExt = "wsv"
        frmFileDialog.Filter = "WordSurv 7 Database File (.wsv) | *.wsv"

        If frmFileDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Return frmFileDialog.FileName
        Else
            Return ""
        End If
    End Function
    Private Function GetNewDatabaseName() As String
        Dim frmSaveDialog As New SaveFileDialog

        frmSaveDialog.Title = "Enter Name for New WordSurv 7.0 File"
        frmSaveDialog.DefaultExt = "wsv"
        frmSaveDialog.Filter = "WordSurv 7 Database File (.wsv) | *.wsv"

        If frmSaveDialog.ShowDialog = DialogResult.OK Then
            Return frmSaveDialog.FileName
        Else
            Return ""
        End If
    End Function
    Private Function GetMergeDatabaseName() As String
        Dim frmMergeDialog As New OpenFileDialog

        frmMergeDialog.Title = "Select WordSurv 7.0 File to Import"
        frmMergeDialog.DefaultExt = "wsv"
        frmMergeDialog.Filter = "WordSurv 7 Database File (.wsv) | *.wsv"

        If frmMergeDialog.ShowDialog = DialogResult.OK Then
            Return frmMergeDialog.FileName
        Else
            Return ""
        End If
    End Function
    Public Sub Open(ByVal filename As String)
        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        'Any operation that needs to open a file must call this function.
        If (data IsNot Nothing) And HasNotSaved Then
            Dim result As MsgBoxResult = MsgBox("Do you want to save your changes?", MsgBoxStyle.YesNoCancel, "WordSurv")
            If result = MsgBoxResult.Cancel Then
                Return
            ElseIf result = MsgBoxResult.Yes Then
                Me.save()
                HasNotSaved = False
            Else
                If DoLog Then Log.Add("Did not save with unsaved changes before opening new file!")
                HasNotSaved = False
            End If
        End If

        Try
            Me.data = WordSurvData.LoadFile(filename)
            data.filename = filename

            prefs.LastOpenedDatabase = data.filename
            Me.BumpUpPreviousDatabases()
            Me.FillRecentDatabasesList()
            Me.tabWordSurv.SelectTab(prefs.CurrentTab)
            Me.tabWordSurv_SelectedIndexChanged(New Object, New EventArgs) 'Hack, calling SelectTab doesn't trigger this event if the tab doesn't change.
            LoadInterrupted = False
        Catch ex As Exception
            'Me.setStatusWarning("Could not open most recent WordSurv Database: " & filename & ": " & ex.Message, True)
            MsgBox("Cannot open " & filename & ".  Please try to open one of the backup files.  To help improve future versions of WordSurv, please send the developers 1) the most recent backup file that opens and 2) the file you wish to open or the first backup which does not open. (See contact information under ""Technical Support"" in the Help).  When sending files, please also include 3) the following technical details and 4) the information requested in ""Required information for correspondence"". Technical details (press Control+c to copy this message): " & ex.Message)
            Me.data = New WordSurvData
            'AJW***
            Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            HasNotSaved = False
            LoadInterrupted = True
            Me.Text = ""
            'panNoDatabase.Visible = True 'AJW***
            'grdGlossDictionary.Visible = False 'AJW***
            'Me.data.filename = Nothing 'AJW***

            'Me.RefreshMenus() 'AJW***
            Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

    End Sub
    Private Sub BumpUpPreviousDatabases()

        Dim tempList As New List(Of String)
        tempList.Add(prefs.RecentDatabase0)
        tempList.Add(prefs.RecentDatabase1)
        tempList.Add(prefs.RecentDatabase2)
        tempList.Add(prefs.RecentDatabase3)
        tempList.Add(prefs.RecentDatabase4)
        tempList.Add(prefs.RecentDatabase5)
        tempList.Add(prefs.RecentDatabase6)
        tempList.Add(prefs.RecentDatabase7)
        tempList.Add(prefs.RecentDatabase8)
        tempList.Add(prefs.RecentDatabase9)
        tempList.Add(prefs.RecentDatabase10)
        If tempList(0) <> prefs.LastOpenedDatabase Then 'AJW***

            'Bump down all the databases.
            For i As Integer = 9 To 0 Step -1
                tempList(i + 1) = tempList(i)
            Next
            tempList(0) = prefs.LastOpenedDatabase

            'If there are any duplicates that are the same as the last opened database, move up to cover them
            'For i As Integer = 1 To tempList.Count - 1
            '    If tempList(i) = tempList(0) Then
            '        For x As Integer = i To tempList.Count - 2
            '            tempList(x) = tempList(x + 1)
            '        Next
            '        'tempList(i) = ""
            '        'tempList.RemoveAt(tempList.Count-1)
            '    End If
            'Next
            For i As Integer = 0 To tempList.Count - 1
                For j As Integer = 1 To tempList.Count - 1
                    If i <> j Then
                        If tempList(i) = tempList(j) Then
                            tempList(j) = ""
                        End If
                    End If
                Next
            Next

            prefs.RecentDatabase0 = tempList(0)
            prefs.RecentDatabase1 = tempList(1)
            prefs.RecentDatabase2 = tempList(2)
            prefs.RecentDatabase3 = tempList(3)
            prefs.RecentDatabase4 = tempList(4)
            prefs.RecentDatabase5 = tempList(5)
            prefs.RecentDatabase6 = tempList(6)
            prefs.RecentDatabase7 = tempList(7)
            prefs.RecentDatabase8 = tempList(8)
            prefs.RecentDatabase9 = tempList(9)
        End If
    End Sub
    Private Sub save()
        'Any operation that needs to save a file must call this function.

        'Make sure all the grids are not in edit mode, so to force any uncommitted cells to commit.
        For Each grd As DataGridView In Me.grids
            grd.EndEdit()
        Next

        'Try 'If they are making a new file, this operation will fail because the file does not exist to make a copy from.
        data.MakeSaveBackup()
        'Catch ex As Exception
        ' End Try

        data.WriteFile()

        prefs.save()
        HasNotSaved = False

        Me.setStatusNotification("Save Complete", True)
    End Sub
    Private Class CreateTimeComparer
        Implements IComparer
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
            Return String.Compare(DirectCast(x, String), DirectCast(y, String))
        End Function
    End Class
    Private Sub FillRecentDatabasesList()
        Me.mnuRecentDatabases.DropDownItems.Clear()
        If prefs.RecentDatabase0 <> "" And System.IO.File.Exists(prefs.RecentDatabase0) Then Me.mnuRecentDatabases.DropDownItems.Add(prefs.RecentDatabase0)
        If prefs.RecentDatabase1 <> "" And System.IO.File.Exists(prefs.RecentDatabase1) Then Me.mnuRecentDatabases.DropDownItems.Add(prefs.RecentDatabase1)
        If prefs.RecentDatabase2 <> "" And System.IO.File.Exists(prefs.RecentDatabase2) Then Me.mnuRecentDatabases.DropDownItems.Add(prefs.RecentDatabase2)
        If prefs.RecentDatabase3 <> "" And System.IO.File.Exists(prefs.RecentDatabase3) Then Me.mnuRecentDatabases.DropDownItems.Add(prefs.RecentDatabase3)
        If prefs.RecentDatabase4 <> "" And System.IO.File.Exists(prefs.RecentDatabase4) Then Me.mnuRecentDatabases.DropDownItems.Add(prefs.RecentDatabase4)
        If prefs.RecentDatabase5 <> "" And System.IO.File.Exists(prefs.RecentDatabase5) Then Me.mnuRecentDatabases.DropDownItems.Add(prefs.RecentDatabase5)
        If prefs.RecentDatabase6 <> "" And System.IO.File.Exists(prefs.RecentDatabase6) Then Me.mnuRecentDatabases.DropDownItems.Add(prefs.RecentDatabase6)
        If prefs.RecentDatabase7 <> "" And System.IO.File.Exists(prefs.RecentDatabase7) Then Me.mnuRecentDatabases.DropDownItems.Add(prefs.RecentDatabase7)
        If prefs.RecentDatabase8 <> "" And System.IO.File.Exists(prefs.RecentDatabase8) Then Me.mnuRecentDatabases.DropDownItems.Add(prefs.RecentDatabase8)
        If prefs.RecentDatabase9 <> "" And System.IO.File.Exists(prefs.RecentDatabase9) Then Me.mnuRecentDatabases.DropDownItems.Add(prefs.RecentDatabase9)

        For Each entry As ToolStripMenuItem In Me.mnuRecentDatabases.DropDownItems
            RemoveHandler entry.Click, AddressOf OpenRecentDatabase
            AddHandler entry.Click, AddressOf OpenRecentDatabase
        Next
    End Sub
    Private Sub OpenRecentDatabase(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim mnuItem As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        Dim databaseFile As String = mnuItem.Text
        Me.Open(databaseFile)
        ClearUndo(data, prefs)
        Me.refreshDictionaryPane()
        Me.refreshSurveyPane()
        Me.RefreshMenus()
        If DoLog Then Log.Add("Opened Recent database " & mnuItem.Text)
    End Sub
    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        If DoLog Then Log.Add("Clicked exit in the menu")
        Me.CommitGrids()
        Me.Close()
    End Sub
    Private Sub WordSurvForm_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If (data IsNot Nothing) And HasNotSaved Then
            Dim result As MsgBoxResult = MsgBox("Do you want to save your changes?", MsgBoxStyle.YesNoCancel, "WordSurv")
            If result = MsgBoxResult.Cancel Then
                e.Cancel = True
            ElseIf result = MsgBoxResult.Yes Then
                Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                Me.save()
                HasNotSaved = False
                Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            Else
                If DoLog Then Log.Add("Did not save before exiting with unsaved changes!")
                HasNotSaved = False
            End If
        End If
    End Sub
    Private Sub WordSurvForm_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If prefs IsNot Nothing Then prefs.save()
        'Dim loglines As String = ""
        'For Each logline As String In Log
        '    loglines &= logline & vbCrLf
        'Next
        'MsgBox(loglines, MsgBoxStyle.Information)

        Application.Exit() 'When this form closes, get rid of all the others as well.
    End Sub
    Private Sub mnuSaveAs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSaveAs.Click
        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Me.CommitGrids()
        'If HasNotSaved Then
        '    Dim result As MsgBoxResult = MsgBox("Do you want to save your current changes to your current file before creating your new save file?", MsgBoxStyle.YesNoCancel, "WordSurv")
        '    If result = MsgBoxResult.Cancel Then
        '        Return
        '    ElseIf result = MsgBoxResult.Yes Then
        '        Me.save()
        '        HasNotSaved = False
        '    End If
        'End If

        Dim newFile As String = GetNewDatabaseName()
        If newFile = "" Then Return
        Try
            'System.IO.File.Copy(data.filename, newFile, True)
            Dim x As System.IO.FileStream = System.IO.File.Create(newFile) 'AJW***
            x.Close() 'AJW***
            data.filename = newFile 'AJW***
            Me.save() 'AJW***
            prefs.LastOpenedDatabase = data.filename 'AJW***
            Me.BumpUpPreviousDatabases() 'AJW***
            Me.FillRecentDatabasesList() 'AJW***
            Me.tabWordSurv.SelectTab(prefs.CurrentTab) 'AJW***
            Me.tabWordSurv_SelectedIndexChanged(New Object, New EventArgs) 'Hack, calling SelectTab doesn't trigger this event if the tab doesn't change.'AJW***

        Catch ex As Exception
            MsgBox("Could not create a new file: " & ex.Message, MsgBoxStyle.Exclamation, "File Error")
            Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            Return
        End Try

        If DoLog Then Log.Add("Saved current database as " & newFile)
        'Me.Open(newFile)
        'Me.save()
        Me.RefreshMenus()
        Me.refreshDictionaryPane()
        Me.refreshSurveyPane()
        Me.refreshComparisonTabLeftPane()
        Me.refreshComparisonTabRightPane()
        Me.refreshComparisonAnalysisTab()
        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub mnuMergeDatabases_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuImportVersion7Database.Click
        Me.ActiveGrid.EndEdit()
        Dim filename As String = Me.GetMergeDatabaseName()
        If filename <> "" Then
            data.MergeCurrentDatabaseWithThisOne(filename)
            Me.tabWordSurv.SelectedIndex = 0
            Me.refreshDictionaryPane()
            Me.refreshSurveyPane()
            If DoLog Then Log.Add("Imported Database " & filename & " with current Database")
        End If
    End Sub
    Private Sub mnuImportVersion6Database_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuImportVersion6Database.Click
        'AJW***
        ImportWordSurv6(data, prefs, Me)
        If DoLog Then Log.Add("Attempted to import version 6 database")
    End Sub
    Private Sub mnuDictionaryExcelExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDictionaryExcelExport.Click
        If data.ExportCurrentDictionaryToExcel() Then
            Me.setStatusNotification("Export successful.", True)
        Else
            Me.setStatusWarning("Export failed.", True)
        End If
        If DoLog Then Log.Add("Exported Current Dictionary to Excel")
    End Sub
    Private Sub mnuDictionaryCSVExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDictionaryCSVExport.Click
        If data.ExportCurrentDictionaryToCSV() Then
            Me.setStatusNotification("Export successful.", True)
        Else
            Me.setStatusWarning("Export failed.", True)
        End If
        If DoLog Then Log.Add("Exported Current Dictionary to CSV")
    End Sub
    Private Sub mnuSurveyExcelExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSurveyExcelExport.Click
        If data.ExportCurrentSurveyToExcel() Then
            Me.setStatusNotification("Export successful.", True)
        Else
            Me.setStatusWarning("Export failed.", True)
        End If
        If DoLog Then Log.Add("Exported Current Survey to Excel")
    End Sub
    Private Sub mnuSurveyCSVExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSurveyCSVExport.Click
        If data.ExportCurrentSurveyToCSV() Then
            Me.setStatusNotification("Export successful.", True)
        Else
            Me.setStatusWarning("Export failed.", True)
        End If
        If DoLog Then Log.Add("Exported Current Survey to CSV")
    End Sub
    Private Sub mnuComparisonExcelExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuComparisonExcelExport.Click
        If data.ExportCurrentComparisonToExcel() Then
            Me.setStatusNotification("Export successful.", True)
        Else
            Me.setStatusWarning("Export failed.", True)
        End If
        If DoLog Then Log.Add("Exported Current Comparison to Excel")
    End Sub
    Private Sub mnuComparisonCSVExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuComparisonCSVExport.Click
        If data.ExportCurrentComparisonToCSV() Then
            Me.setStatusNotification("Export successful.", True)
        Else
            Me.setStatusWarning("Export failed.", True)
        End If
        If DoLog Then Log.Add("Exported Current Comparison to CSV")
    End Sub
    Private Sub mnuCAExcelExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCAExcelExport.Click
        If data.ExportCurrentComparisonAnalysisToExcel() Then
            Me.setStatusNotification("Export successful.", True)
        Else
            Me.setStatusWarning("Export failed.", True)
        End If
        If DoLog Then Log.Add("Exported Current Comparison Analysis to Excel")
    End Sub
    Private Sub mnuCACSVExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCACSVExport.Click
        If data.ExportCurrentComparisonAnalysisToCSV() Then
            Me.setStatusNotification("Export successful.", True)
        Else
            Me.setStatusWarning("Export failed.", True)
        End If
        If DoLog Then Log.Add("Exported Current Comparison Analysis to CSV")
    End Sub
    Private Sub mnuDDExcelExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDDExcelExport.Click
        If data.ExportCurrentDDToExcel() Then
            Me.setStatusNotification("Export successful.", True)
        Else
            Me.setStatusWarning("Export failed.", True)
        End If
        If DoLog Then Log.Add("Exported Current DD grid to Excel")
    End Sub
    Private Sub mnuDDCSVExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDDCSVExport.Click
        If data.ExportCurrentDDToCSV() Then
            Me.setStatusNotification("Export successful.", True)
        Else
            Me.setStatusWarning("Export failed.", True)
        End If
        If DoLog Then Log.Add("Exported Current DD grid to CSV")
    End Sub
    Private Sub mnuPhonoStatsExcelExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPhonoStatsExcelExport.Click
        If data.ExportCurrentPhonoStatsToExcel() Then
            Me.setStatusNotification("Export successful.", True)
        Else
            Me.setStatusWarning("Export failed.", True)
        End If
        If DoLog Then Log.Add("Exported Current PhonoStats Analysis to Excel")
    End Sub
    Private Sub mnuPhonoStatsCSVExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPhonoStatsCSVExport.Click
        If data.ExportCurrentPhonoStatsToCSV() Then
            Me.setStatusNotification("Export successful.", True)
        Else
            Me.setStatusWarning("Export failed.", True)
        End If
        If DoLog Then Log.Add("Exported Current PhonoStats Analysis to CSV")
    End Sub
    Private Sub mnuCOMPASSExcelExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCOMPASSExcelExport.Click
        If data.ExportCurrentCOMPASSToExcel() Then
            Me.setStatusNotification("Export successful.", True)
        Else
            Me.setStatusWarning("Export failed.", True)
        End If
        If DoLog Then Log.Add("Exported Current COMPASS data to Excel")
    End Sub
    Private Sub mnuCOMPASSCSVExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCOMPASSCSVExport.Click
        If data.ExportCurrentCOMPASSToCSV() Then
            Me.setStatusNotification("Export successful.", True)
        Else
            Me.setStatusWarning("Export failed.", True)
        End If
        If DoLog Then Log.Add("Exported Current COMPASS data to CSV")
    End Sub
#End Region


#Region "Tab Independent Operations"
    Private Sub GridGotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.ActiveGrid = DirectCast(sender, DataGridView)
        Me.LastActivatedGrid = Me.ActiveGrid
    End Sub
    Private Function GetDefaultGridForCurrentTab() As DataGridView
        Select Case Me.tabWordSurv.SelectedTab.Text
            Case "Word List Management"
                Return Me.grdGlossDictionary
            Case "Comparisons"
                Return Me.grdComparison
            Case "Comparison Analysis"
                Return Me.grdComparisonAnalysis
            Case "Degrees of Difference"
                Return Me.grdDegreesOfDifference
            Case "Phonostatistical Analysis"
                Return Me.grdPhonoStats
            Case "Comparativist's Assistant (COMPASS)"
                Return Me.grdPhoneCorr
        End Select
        Return Nothing
    End Function
    Private Sub MYExceptionHandler(ByVal sender As Object, ByVal e As UnhandledExceptionEventArgs)
        Dim ex As Exception = CType(e.ExceptionObject, Exception)
        Dim frmException As New ExceptionForm
        frmException.txtExceptionText.Text = ex.Message & vbCrLf
        Dim stackTrace As String = ex.StackTrace
        Dim lines As String() = Split(stackTrace, vbCrLf)
        For Each line As String In lines
            If line.Contains(":line ") Then
                frmException.txtExceptionText.Text &= line & vbCrLf
            End If
        Next
        For Each logline As String In Log
            frmException.txtExceptionText.Text &= logline & vbCrLf
        Next
        frmException.ShowDialog()
        Application.Exit()
    End Sub
    Private Sub MYThreadHandler(ByVal sender As Object, ByVal e As Threading.ThreadExceptionEventArgs)
        Dim ex As Exception = CType(e.Exception, Exception)
        Dim frmException As New ExceptionForm
        frmException.txtExceptionText.Text = ex.Message & vbCrLf
        Dim stackTrace As String = ex.StackTrace
        Dim lines As String() = Split(stackTrace, vbCrLf)
        For Each line As String In lines
            If line.Contains(":line ") Then
                frmException.txtExceptionText.Text &= line & vbCrLf
            End If
        Next
        frmException.ShowDialog()
        Application.Exit()
    End Sub

    'This code disables the top menu items during edit mode so that the user can't explode things that might explode should the menu be used in edit mode.
    Private Sub gridCellBeginEditHandleMenus(ByVal sender As System.Object, ByVal e As System.EventArgs)
        For Each mnuItem As ToolStripMenuItem In Me.mnuWordSurv.Items
            mnuItem.Enabled = False
        Next

        cmnuDictionary.Enabled = False
        cmnuVariety.Enabled = False
        cmnuComparison.Enabled = False
        cmnuComparisonAnalysis.Enabled = False
        cmnuDegreesOfDifference.Enabled = False
        cmnuPhonoStats.Enabled = False
        cmnuCOMPASS.Enabled = False

        Try
            If Me.frmSearch IsNot Nothing Then Me.frmSearch.Close()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub gridCellEndEditHandleMenus(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.RefreshMenus()
    End Sub

    Private Sub tabWordSurv_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabWordSurv.SelectedIndexChanged
        RefreshBasedOnCurrentTab()
        prefs.CurrentTab = Me.tabWordSurv.SelectedIndex
        Me.ActiveGrid = Me.GetDefaultGridForCurrentTab()
        Me.ActiveGrid.Focus()
        If DoLog Then Log.Add("Changed current tab to " & Me.tabWordSurv.SelectedTab.Text & " (" & Me.tabWordSurv.SelectedIndex.ToString & ")")
    End Sub
    Public Sub CommitGrids() 'AJW***
        'You can use this function to commit grids in a general purpose event handler that doesn't necessarily know what the current grid is specifically
        If Me.ActiveGrid IsNot Nothing Then Me.ActiveGrid.EndEdit()
    End Sub
    Public Sub RefreshBasedOnCurrentTab()
        Select Case Me.tabWordSurv.SelectedTab.Text
            Case "Word List Management"
                Me.refreshDictionaryPane()
                Me.refreshSurveyPane()
            Case "Comparisons"
                Me.refreshComparisonTabLeftPane()
                Me.refreshComparisonTabRightPane()
            Case "Comparison Analysis"
                Me.refreshComparisonAnalysisTab()
            Case "Degrees of Difference"
                Me.refreshDegreesOfDifferenceTab(True)
            Case "Phonostatistical Analysis"
                Me.refreshPhonoStatsTab()
            Case "Comparativist's Assistant (COMPASS)"
                Me.refreshCOMPASSTab()
        End Select
        Me.RefreshMenus()
    End Sub
    Private Sub setSubmenuItemsEnabledState(ByRef toolStripMenuItem As ToolStripMenuItem, ByVal enable As Boolean)
        For Each innerMnuItem As ToolStripItem In toolStripMenuItem.DropDownItems
            If TypeOf innerMnuItem Is ToolStripMenuItem Then
                innerMnuItem.Enabled = enable
            End If
        Next
        toolStripMenuItem.Enabled = enable
    End Sub
    Private Sub RefreshMenus()
        'This could probably be more concise, but it does work.

        If data.filename Is Nothing Then
            Me.mnuSaveAs.Enabled = False
            'Me.mnuRecentDatabases.Enabled = False
            setSubmenuItemsEnabledState(Me.mnuEdit, False)
            setSubmenuItemsEnabledState(Me.mnuTools, False)
            setSubmenuItemsEnabledState(Me.mnuDictionary, False)
            setSubmenuItemsEnabledState(Me.mnuSurvey, False)
            setSubmenuItemsEnabledState(Me.mnuVariety, False)
            setSubmenuItemsEnabledState(Me.mnuComparison, False)
            setSubmenuItemsEnabledState(Me.mnuComparisonAnalysis, False)
            setSubmenuItemsEnabledState(Me.mnuDegreesOfDifference, False)
            setSubmenuItemsEnabledState(Me.mnuCOMPASS, False)
            Me.cmnuDictionary.Enabled = False
            Me.cmnuVariety.Enabled = False
            Me.cmnuComparison.Enabled = False
            Me.cmnuComparisonAnalysis.Enabled = False
            Me.cmnuDegreesOfDifference.Enabled = False

            Me.mnuDictionary.Visible = False
            Me.mnuSurvey.Visible = False
            Me.mnuVariety.Visible = False
            Me.mnuComparison.Visible = False
            Me.mnuComparisonAnalysis.Visible = False
            Me.mnuDegreesOfDifference.Visible = False
            Me.mnuCOMPASS.Visible = False
            Me.txtComparisonDescription.Enabled = False

            Me.panNoDatabase.Visible = True
            Me.panNoDictionaries.Visible = False
            Me.cboGlossDictionaries.Enabled = False
            Me.grdGlossDictionary.Enabled = False
            Me.cboGlossDictionarySort.Enabled = False
            Me.cboSurveys.Enabled = False
            Me.cboVarieties.Enabled = False
            Me.cboVarietySorts.Enabled = False
            Me.txtSurveyDescription.Enabled = False
            Me.txtAssociatedDictionary.Enabled = False
            Me.txtSurveyDescription.Text = ""
            Me.txtVarietyDescription.Text = ""
            Me.txtVarietyDescription.Enabled = False
            Me.cboComparison.Enabled = False
            Me.cboComparisonSorts.Enabled = False
            Me.btnComparisonStatistics.Enabled = False
            Me.grdGlossDictionary.AllowUserToAddRows = False
            Me.mnuSave.Enabled = False
            Me.mnuSaveAs.Enabled = False

        Else
            Me.mnuFile.Enabled = True
            Me.mnuHelp.Enabled = True
            setSubmenuItemsEnabledState(Me.mnuFile, True)
            setSubmenuItemsEnabledState(Me.mnuHelp, True)
            Me.mnuSave.Enabled = True
            Me.mnuSaveAs.Enabled = True
            Me.grdGlossDictionary.AllowUserToAddRows = True
            Me.Text = "WordSurv 7 (" & data.filename & ")"
            Me.FillRecentDatabasesList()

            Me.panNoDatabase.Visible = False
            Me.mnuSaveAs.Enabled = True
            Me.mnuRecentDatabases.Enabled = True
            setSubmenuItemsEnabledState(Me.mnuEdit, True)
            setSubmenuItemsEnabledState(Me.mnuTools, True)


            Select Case Me.tabWordSurv.SelectedTab.Text
                Case "Word List Management"
                    Me.mnuDictionary.Visible = True
                    Me.mnuSurvey.Visible = True
                    Me.mnuVariety.Visible = True
                    Me.mnuUndo.Enabled = True
                    Me.mnuRedo.Enabled = True
                    setSubmenuItemsEnabledState(Me.mnuDictionary, True)
                    setSubmenuItemsEnabledState(Me.mnuSurvey, True)
                    setSubmenuItemsEnabledState(Me.mnuEdit, True)


                    Me.mnuComparison.Visible = False
                    Me.mnuComparisonAnalysis.Visible = False
                    Me.mnuDegreesOfDifference.Visible = False
                    Me.mnuCOMPASS.Visible = False
                    setSubmenuItemsEnabledState(Me.mnuComparison, False)
                    setSubmenuItemsEnabledState(Me.mnuComparisonAnalysis, False)
                    setSubmenuItemsEnabledState(Me.mnuDegreesOfDifference, False)
                    setSubmenuItemsEnabledState(Me.mnuCOMPASS, False)

                    If Me.cboVarieties.Items.Count = 0 Then
                        setSubmenuItemsEnabledState(Me.mnuVariety, False)
                        Me.cmnuVariety.Enabled = False

                        If Me.cboSurveys.Items.Count <> 0 Then
                            Me.mnuVariety.Enabled = True
                            Me.mnuNewVariety.Enabled = True
                        End If

                        Me.panNoVariety1.Visible = True
                        Me.panNoVariety2.Visible = True
                        Me.txtVarietyDescription.Enabled = False
                    Else
                        setSubmenuItemsEnabledState(Me.mnuVariety, True)
                        Me.cmnuVariety.Enabled = True

                        Me.panNoVariety1.Visible = False
                        Me.panNoVariety2.Visible = False
                        Me.txtVarietyDescription.Enabled = True
                    End If

                    If Me.cboSurveys.Items.Count = 0 Then
                        'disable all but new survey menu item
                        setSubmenuItemsEnabledState(Me.mnuSurvey, False)
                        Me.mnuSurvey.Enabled = True
                        Me.mnuNewSurvey.Enabled = True
                        Me.panNoSurvey.Visible = True

                        Me.cboSurveys.Enabled = False
                        Me.cboVarieties.Enabled = False
                        Me.cboVarietySorts.Enabled = False
                        Me.txtSurveyDescription.Enabled = False
                        Me.txtSurveyDescription.Text = ""
                        Me.txtAssociatedDictionary.Enabled = False
                    Else
                        'enable all survey menu items
                        Me.panNoSurvey.Visible = False
                        Me.cboSurveys.Enabled = True
                        Me.cboVarieties.Enabled = True
                        Me.cboVarietySorts.Enabled = True
                        Me.txtAssociatedDictionary.Enabled = True
                        Me.txtSurveyDescription.Enabled = True
                        setSubmenuItemsEnabledState(Me.mnuSurvey, True)
                    End If

                    If Me.cboGlossDictionaries.Items.Count = 0 Then
                        setSubmenuItemsEnabledState(Me.mnuDictionary, False)
                        Me.cmnuDictionary.Enabled = False
                        Me.mnuDictionary.Enabled = True
                        Me.mnuNewDictionary.Enabled = True
                        Me.panNoDictionaries.Visible = True

                        Me.cboGlossDictionaries.Enabled = False
                        Me.grdGlossDictionary.Enabled = False
                        Me.cboGlossDictionarySort.Enabled = False

                        setSubmenuItemsEnabledState(Me.mnuSurvey, False)
                        Me.mnuSurvey.Enabled = False
                        Me.grdGlossDictionary.AllowUserToAddRows = False
                    Else
                        Me.panNoDictionaries.Visible = False
                        Me.grdGlossDictionary.AllowUserToAddRows = True
                        Me.cboGlossDictionaries.Enabled = True
                        Me.grdGlossDictionary.Enabled = True
                        Me.cboGlossDictionarySort.Enabled = True

                        setSubmenuItemsEnabledState(Me.mnuDictionary, True)
                        Me.cmnuDictionary.Enabled = True
                    End If

                    'Disable Merge Surveys if there is only one Survey
                    If data.GetSurveyNames().Length = 1 Then Me.mnuMergeSurveys.Enabled = False

                Case "Comparisons"
                    Me.mnuComparison.Visible = True
                    setSubmenuItemsEnabledState(Me.mnuComparison, True)
                    setSubmenuItemsEnabledState(Me.mnuEdit, True)

                    Me.mnuDictionary.Visible = False
                    Me.mnuSurvey.Visible = False
                    Me.mnuVariety.Visible = False
                    Me.mnuComparisonAnalysis.Visible = False
                    Me.mnuDegreesOfDifference.Visible = False
                    Me.mnuCOMPASS.Visible = False
                    Me.mnuUndo.Enabled = True
                    Me.mnuRedo.Enabled = True
                    setSubmenuItemsEnabledState(Me.mnuDictionary, False)
                    setSubmenuItemsEnabledState(Me.mnuSurvey, False)
                    setSubmenuItemsEnabledState(Me.mnuComparisonAnalysis, False)
                    setSubmenuItemsEnabledState(Me.mnuDegreesOfDifference, False)
                    setSubmenuItemsEnabledState(Me.mnuCOMPASS, False)

                    If Me.cboComparison.Items.Count = 0 Then
                        setSubmenuItemsEnabledState(Me.mnuComparison, False)
                        Me.cmnuComparison.Enabled = False

                        If data.GetSurveyNames().Length > 0 Then
                            Me.mnuComparison.Enabled = True
                            Me.mnuNewComparison.Enabled = True
                        End If
                        Me.cboComparison.Enabled = False
                        Me.cboComparisonSorts.Enabled = False
                        Me.grdComparison.Enabled = False
                        Me.grdComparisonGloss.Enabled = False
                        Me.panNoComparisonsCompTab.Visible = True
                        Me.btnComparisonStatistics.Enabled = False
                        Me.cmnuCutVarieties.Enabled = False
                        Me.cmnuPasteVarieties.Enabled = False
                        Me.txtComparisonDescription.Enabled = False

                        If Me.cboSurveys.Items.Count <> 0 Then
                            Me.mnuComparison.Enabled = True
                            Me.mnuNewComparison.Enabled = True
                        End If
                    Else
                        setSubmenuItemsEnabledState(Me.mnuComparison, True)
                        Me.cmnuComparison.Enabled = True

                        Me.cboComparison.Enabled = True
                        Me.cboComparisonSorts.Enabled = True
                        Me.grdComparison.Enabled = True
                        Me.grdComparisonGloss.Enabled = True
                        Me.grdComparisonAnalysis.Enabled = True
                        Me.rdoTally.Enabled = True
                        Me.rdoTotal.Enabled = True
                        Me.rdoPercent.Enabled = True
                        Me.mnuUndo.Visible = True
                        Me.mnuUndo.Enabled = True
                        Me.mnuRedo.Visible = True
                        Me.mnuRedo.Enabled = True
                        Me.panNoComparisonsCompTab.Visible = False
                        Me.btnComparisonStatistics.Enabled = True
                        Me.txtComparisonDescription.Enabled = True

                        Me.cmnuCutVarieties.Enabled = True
                        Me.cmnuPasteVarieties.Enabled = True
                    End If

                Case "Comparison Analysis"
                    Me.mnuComparisonAnalysis.Visible = True
                    setSubmenuItemsEnabledState(Me.mnuComparisonAnalysis, True)

                    Me.mnuDictionary.Visible = False
                    Me.mnuSurvey.Visible = False
                    Me.mnuVariety.Visible = False
                    Me.mnuComparison.Visible = False
                    Me.mnuDegreesOfDifference.Visible = False
                    Me.mnuCOMPASS.Visible = False
                    Me.mnuCutCells.Enabled = False
                    Me.mnuDelete.Enabled = False
                    Me.mnuPaste.Enabled = False
                    Me.mnuUndo.Enabled = True
                    Me.mnuRedo.Enabled = True
                    setSubmenuItemsEnabledState(Me.mnuDictionary, False)
                    setSubmenuItemsEnabledState(Me.mnuSurvey, False)
                    setSubmenuItemsEnabledState(Me.mnuComparison, False)
                    setSubmenuItemsEnabledState(Me.mnuDegreesOfDifference, False)
                    setSubmenuItemsEnabledState(Me.mnuCOMPASS, False)

                    If Me.cboComparisonAnalysis.Items.Count = 0 Then
                        Me.cmnuComparisonAnalysis.Enabled = False
                        Me.cboComparisonAnalysis.Enabled = False
                        Me.grdComparisonAnalysis.Enabled = False
                        Me.rdoTally.Enabled = False
                        Me.rdoTotal.Enabled = False
                        Me.rdoPercent.Enabled = False
                        setSubmenuItemsEnabledState(Me.mnuComparisonAnalysis, False)
                        Me.panNoComparisonsCompAnalTab.Visible = True
                        Me.cmnuCutCARows.Enabled = False
                        Me.cmnuPasteCARows.Enabled = False
                    Else
                        Me.cmnuComparisonAnalysis.Enabled = True
                        Me.cboComparisonAnalysis.Enabled = True
                        Me.grdComparisonAnalysis.Enabled = True
                        Me.rdoTally.Enabled = True
                        Me.rdoTotal.Enabled = True
                        Me.rdoPercent.Enabled = True
                        setSubmenuItemsEnabledState(Me.mnuComparisonAnalysis, True)
                        Me.panNoComparisonsCompAnalTab.Visible = False
                        Me.cmnuCutCARows.Enabled = True
                        Me.cmnuPasteCARows.Enabled = True
                    End If

                Case "Degrees of Difference"
                    Me.mnuDegreesOfDifference.Visible = True
                    setSubmenuItemsEnabledState(Me.mnuDegreesOfDifference, True)

                    Me.mnuDictionary.Visible = False
                    Me.mnuSurvey.Visible = False
                    Me.mnuVariety.Visible = False
                    Me.mnuComparison.Visible = False
                    Me.mnuComparisonAnalysis.Visible = False
                    Me.mnuCOMPASS.Visible = False
                    Me.mnuCutCells.Enabled = False
                    Me.mnuDelete.Enabled = False
                    Me.mnuPaste.Enabled = False
                    Me.mnuUndo.Enabled = True
                    Me.mnuRedo.Enabled = True
                    setSubmenuItemsEnabledState(Me.mnuDictionary, False)
                    setSubmenuItemsEnabledState(Me.mnuSurvey, False)
                    setSubmenuItemsEnabledState(Me.mnuComparison, False)
                    setSubmenuItemsEnabledState(Me.mnuComparisonAnalysis, False)
                    setSubmenuItemsEnabledState(Me.mnuCOMPASS, False)

                    If Me.cboDegreesOfDifference.Items.Count = 0 Then
                        Me.cmnuDegreesOfDifference.Enabled = False
                        Me.cboDegreesOfDifference.Enabled = False
                        Me.grdDegreesOfDifference.Enabled = False
                        Me.panNoComparisonsDD.Visible = True
                        Me.cboDDPhoneUsing.Enabled = False
                        Me.cmnuCutDDRows.Enabled = False
                        Me.cmnuPasteDDRows.Enabled = False
                    Else
                        Me.cmnuDegreesOfDifference.Enabled = True
                        Me.cboDegreesOfDifference.Enabled = True
                        Me.grdDegreesOfDifference.Enabled = True
                        Me.panNoComparisonsDD.Visible = False
                        Me.cboDDPhoneUsing.Enabled = True
                        Me.cmnuCutDDRows.Enabled = True
                        Me.cmnuPasteDDRows.Enabled = True
                    End If

                Case "Phonostatistical Analysis"
                    Me.mnuDictionary.Visible = False
                    Me.mnuSurvey.Visible = False
                    Me.mnuVariety.Visible = False
                    Me.mnuComparison.Visible = False
                    Me.mnuComparisonAnalysis.Visible = False
                    Me.mnuDegreesOfDifference.Visible = False
                    Me.mnuCOMPASS.Visible = False
                    Me.mnuCutCells.Enabled = False
                    Me.mnuDelete.Enabled = False
                    Me.mnuPaste.Enabled = False
                    Me.mnuUndo.Enabled = False
                    Me.mnuRedo.Enabled = False
                    setSubmenuItemsEnabledState(Me.mnuDictionary, False)
                    setSubmenuItemsEnabledState(Me.mnuSurvey, False)
                    setSubmenuItemsEnabledState(Me.mnuComparison, False)
                    setSubmenuItemsEnabledState(Me.mnuComparisonAnalysis, False)
                    setSubmenuItemsEnabledState(Me.mnuDegreesOfDifference, False)
                    setSubmenuItemsEnabledState(Me.mnuCOMPASS, False)

                    If Me.cboPhonoStats.Items.Count = 0 Then
                        Me.cboPhonoStats.Enabled = False
                        Me.rdoPhonoStats1.Enabled = False
                        Me.rdoPhonoStats2.Enabled = False
                        Me.rdoPhonoStats3.Enabled = False
                        Me.rdoPhonoStats4.Enabled = False
                        Me.grdPhonoStats.Enabled = False
                        Me.panNoComparisonsPhonoStatsTab.Visible = True
                    Else
                        Me.cboPhonoStats.Enabled = True
                        Me.rdoPhonoStats1.Enabled = True
                        Me.rdoPhonoStats2.Enabled = True
                        Me.rdoPhonoStats3.Enabled = True
                        Me.rdoPhonoStats4.Enabled = True
                        Me.grdPhonoStats.Enabled = True
                        Me.panNoComparisonsPhonoStatsTab.Visible = False
                    End If

                Case "Comparativist's Assistant (COMPASS)"
                    Me.mnuCOMPASS.Visible = True
                    setSubmenuItemsEnabledState(Me.mnuCOMPASS, True)

                    Me.mnuDictionary.Visible = False
                    Me.mnuSurvey.Visible = False
                    Me.mnuVariety.Visible = False
                    Me.mnuComparison.Visible = False
                    Me.mnuComparisonAnalysis.Visible = False
                    Me.mnuDegreesOfDifference.Visible = False
                    Me.mnuCutCells.Enabled = False
                    Me.mnuDelete.Enabled = False
                    Me.mnuPaste.Enabled = False
                    Me.mnuUndo.Enabled = False
                    Me.mnuRedo.Enabled = False
                    setSubmenuItemsEnabledState(Me.mnuDictionary, False)
                    setSubmenuItemsEnabledState(Me.mnuSurvey, False)
                    setSubmenuItemsEnabledState(Me.mnuComparison, False)
                    setSubmenuItemsEnabledState(Me.mnuDegreesOfDifference, False)
                    setSubmenuItemsEnabledState(Me.mnuComparisonAnalysis, False)


                    If Me.cboCOMPASS.Items.Count = 0 Then
                        Me.panNoComparisonsCOMPASSTab.Visible = True
                        Me.cboCOMPASS.Enabled = False
                        Me.cboCOMPASSVariety1.Enabled = False
                        Me.cboCOMPASSVariety2.Enabled = False
                        Me.nudCOMPASSBottom.Enabled = False
                        Me.nudCOMPASSLower.Enabled = False
                        Me.nudCOMPASSUpper.Enabled = False
                        Me.btnWordStrengthsSummary.Enabled = False
                        Me.grdPhoneCorr.Enabled = False
                        Me.grdCognateStrengths.Enabled = False
                        Me.mnuCOMPASS.Enabled = False
                        Me.rdoShowCounts.Enabled = False
                        Me.rdoShowStrengths.Enabled = False
                    Else
                        Me.panNoComparisonsCOMPASSTab.Visible = False
                        Me.cboCOMPASS.Enabled = True
                        Me.cboCOMPASSVariety1.Enabled = True
                        Me.cboCOMPASSVariety2.Enabled = True
                        Me.nudCOMPASSBottom.Enabled = True
                        Me.nudCOMPASSLower.Enabled = True
                        Me.nudCOMPASSUpper.Enabled = True
                        Me.btnWordStrengthsSummary.Enabled = True
                        Me.grdPhoneCorr.Enabled = True
                        Me.grdCognateStrengths.Enabled = True
                        Me.mnuCOMPASS.Enabled = True
                        Me.rdoShowCounts.Enabled = True
                        Me.rdoShowStrengths.Enabled = True
                    End If
            End Select
        End If

        'Make sure we can't export things when there are none of those
        Dim areDictionaries As Boolean = data.GetDictionaryNames().Length > 0
        Dim areSurveys As Boolean = data.GetSurveyNames().Length > 0
        Dim areComparisons As Boolean = data.GetComparisonNames().Length > 0
        Dim areCOMPASSes As Boolean = data.COMPASSValuesExist()
        Me.mnuExportDictionary.Enabled = areDictionaries
        Me.mnuExportSurvey.Enabled = areSurveys
        Me.mnuExportComparison.Enabled = areComparisons
        Me.mnuExportComparisonAnalysis.Enabled = areComparisons
        Me.mnuExportDegreesOfDifference.Enabled = areComparisons
        Me.mnuExportPhonoStats.Enabled = areComparisons
        Me.mnuExportCOMPASS.Enabled = areCOMPASSes


        'These are set to true elsewhere if they should be enabled.
        Me.mnuPasteDictionaryRows.Enabled = False
        Me.cmnuPasteGlosses.Enabled = False

        Me.mnuPasteComparisonRows.Enabled = False
        Me.cmnuPasteVarieties.Enabled = False

        Me.mnuPasteCARows.Enabled = False
        Me.cmnuPasteCARows.Enabled = False

        Me.mnuPasteDDRows.Enabled = False
        Me.cmnuPasteDDRows.Enabled = False
    End Sub

    'Any time you need to display something in the bottom status notification area, use these functions
    Public Sub setStatusNotification(ByVal msg As String, ByVal doTimeout As Boolean)
        Me.stsLabel1.Text = msg
        If doTimeout Then
            Me.tmrMessageLifetime.Start()
        End If
        If DoLog Then Log.Add("Status Notification: " & msg)
    End Sub
    Public Sub setStatusWarning(ByVal msg As String, ByVal doTimeout As Boolean)
        Me.stsLabel1.Text = msg
        Me.stsStatusBar.BackColor = INVALID_COLOR
        Beep()
        If doTimeout Then
            Me.tmrMessageLifetime.Start()
        End If
        If DoLog Then Log.Add("Status Warning: " & msg)
    End Sub
    Public Sub setStatusError(ByVal msg As String, ByVal doTimeout As Boolean)
        Me.stsLabel1.Text = msg
        Me.stsStatusBar.BackColor = ERROR_COLOR
        Beep()
        If doTimeout Then
            Me.tmrMessageLifetime.Start()
        End If
        If DoLog Then Log.Add("Status Error: " & msg)
    End Sub
    Private Sub clearStatusMessage()
        Me.stsLabel1.Text = ""
        Me.stsStatusBar.BackColor = Color.Empty
    End Sub
    Private Sub tmrMessageLifetime_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrMessageLifetime.Tick
        Me.clearStatusMessage()
        Me.tmrMessageLifetime.Stop()
    End Sub

    'The shortcuts for cell copy, paste, etc override the standard shortcuts in the textboxes, so we have to disable the menus while in the text boxes
    Private Sub txt_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSurveyDescription.Enter, txtVarietyDescription.Enter, txtComparisonDescription.Enter
        Me.mnuCopy.Enabled = False
        Me.mnuCutCells.Enabled = False
        Me.mnuDelete.Enabled = False
        Me.mnuPaste.Enabled = False
        Me.mnuUndo.Enabled = False
    End Sub
    Private Sub txt_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSurveyDescription.Leave, txtVarietyDescription.Leave, txtComparisonDescription.Leave
        Me.mnuCopy.Enabled = True
        Me.mnuCutCells.Enabled = True
        Me.mnuDelete.Enabled = True
        Me.mnuPaste.Enabled = True
        Me.mnuUndo.Enabled = True
        StoreForUndo(data, prefs)
    End Sub
    Private Sub txtSurveyDescription_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSurveyDescription.Leave
        If DoLog Then Log.Add("Changed Current Survey's Description to " & Me.txtSurveyDescription.Text)
    End Sub
    Private Sub txtVarietyDescription_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVarietyDescription.Leave
        If DoLog Then Log.Add("Changed Current Variety's Description to " & Me.txtVarietyDescription.Text)
    End Sub
    Private Sub txtComparisonDescription_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtComparisonDescription.Leave
        If DoLog Then Log.Add("Changed Current Comparison's Description to " & Me.txtComparisonDescription.Text)
    End Sub

    Private Function safeToString(ByVal obj As Object) As String
        If obj Is Nothing Then
            Return ""
        Else
            Return obj.ToString
        End If
    End Function
    Private Sub FindActiveGridSelectionBoundaries(ByRef startRow As Integer, ByRef endRow As Integer, ByRef startCol As Integer, ByRef endCol As Integer)
        startRow = Me.ActiveGrid.RowCount - 1
        endRow = 0
        startCol = Me.ActiveGrid.ColumnCount - 1
        endCol = 0

        For Each cell As DataGridViewCell In Me.ActiveGrid.SelectedCells
            If cell.RowIndex < startRow Then startRow = cell.RowIndex
            If cell.RowIndex > endRow Then endRow = cell.RowIndex
            If cell.ColumnIndex < startCol Then startCol = cell.ColumnIndex
            If cell.ColumnIndex > endCol Then endCol = cell.ColumnIndex
        Next
    End Sub
    Private Sub mnuCutCells_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCutCells.Click, cmnuCutDictionaryCells.Click, cmnuCutVarietyCells.Click, cmnuCutComparisonCells.Click
        Me.OperationInProgress = True
        Me.doCellCutOrCopy(True)
        Me.OperationInProgress = False
        StoreForUndo(data, prefs)
    End Sub
    Private Sub mnuCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCopy.Click, cmnuCopyDictionaryCells.Click, cmnuCopyVarietyCells.Click, cmnuCopyComparisonCells.Click, cmnuCopyCACells.Click, cmnuCopyDDCells.Click, cmnuCopyPhonoStatsCells.Click, cmnuCopyCOMPASSCells.Click
        Me.OperationInProgress = True
        Me.doCellCutOrCopy(False)
        Me.OperationInProgress = False
        StoreForUndo(data, prefs)
    End Sub
    Private Sub doCellCutOrCopy(ByVal isCut As Boolean)
        If Me.ActiveGrid Is Nothing Then Return
        If Me.ActiveGrid.CurrentCell Is Nothing Then Return
        If Me.ActiveGrid Is Me.grdGlossDictionary AndAlso Me.grdGlossDictionary.RowCount = 1 Then Return

        'If we are in edit mode, this function is called instead of the usual behavior of cutting that normally happens, so we have to reimplement cutting/copying text. O_o
        If Me.ActiveGrid.CurrentCell.IsInEditMode Then
            Try
                Clipboard.SetText(DirectCast(Me.ActiveGrid.EditingControl, DataGridViewTextBoxEditingControl).SelectedText)
                If isCut Then DirectCast(Me.ActiveGrid.EditingControl, DataGridViewTextBoxEditingControl).SelectedText = ""
            Catch ex As Exception
            End Try
            Return
        End If

        'Otherwise, the user wants to do a cell cut outside of edit mode, so we take all of the cells in the selection, delete their text, and put them in the copy buffer.
        Dim startRow, endRow, startCol, endCol As Integer
        Me.FindActiveGridSelectionBoundaries(startRow, endRow, startCol, endCol)
        If Me.ActiveGrid Is Me.grdGlossDictionary AndAlso endRow = Me.grdGlossDictionary.NewRowIndex Then
            endRow -= 1
            If endRow < 0 Then endRow = 0
        End If

        If Me.ActiveGrid Is Me.grdGlossDictionary And startCol = 0 And isCut Then 'ajw*** AND iScUT
            isCut = False 'don't allow cutting out of the gloss column
            Me.setStatusWarning("Gloss cells can only be copied, not cut. To rearrange glosses use ""Cut rows"" (Control-Shift-X) found in the Dictionary menu.", True)
        End If

        'Me.CellCopyBuffer = Me.Make2DStringList(endRow - startRow + 1, endCol - startCol + 1)

        Dim cpStr As String = ""
        Dim arrRowIndex As Integer = 0
        For grdRowIndex As Integer = startRow To endRow
            Dim arrColIndex As Integer = 0
            For grdColIndex As Integer = startCol To endCol
                'Me.CellCopyBuffer(arrRowIndex)(arrColIndex) = Me.ActiveGrid.Rows(grdRowIndex).Cells(grdColIndex).Value.ToString
                cpStr &= Me.ActiveGrid.Rows(grdRowIndex).Cells(grdColIndex).Value.ToString & vbTab

                If isCut Then Me.ActiveGrid.Rows(grdRowIndex).Cells(grdColIndex).Value = ""
                arrColIndex += 1
            Next
            cpStr = cpStr.Remove(cpStr.Length - 1) 'remove trailing tab
            cpStr &= vbCrLf
            arrRowIndex += 1
        Next
        If DoLog Then Log.Add("Cut/copied cells from (" & startCol.ToString & ", " & startRow.ToString & ") to (" & endCol.ToString & ", " & endRow.ToString & ")")
        If cpStr.Length >= 1 Then cpStr = cpStr.Substring(0, cpStr.LastIndexOf(vbCrLf)) 'Remove trailing newline

        Try 'AJW***
            Clipboard.SetText(cpStr) 'AJW***
        Catch 'AJW***
            'Clipboard.SetText("")
            MsgBox("Cutting or copying empty cells results in no change to the Clipboard buffer!", MsgBoxStyle.Exclamation, "Cutting or Copying Empty Cells")
        End Try 'AJW***
        Me.mnuPaste.Enabled = True
        Me.cmnuPasteDictionaryCells.Enabled = True
        Me.cmnuPasteVarietyCells.Enabled = True
        Me.cmnuPasteComparisonCells.Enabled = True
    End Sub
    Public Function Make2DStringList(ByVal rows As Integer, ByVal cols As Integer) As List(Of List(Of String))
        Dim lst As New List(Of List(Of String))
        For row As Integer = 0 To rows - 1
            lst.Add(New List(Of String))
            For col As Integer = 0 To cols - 1
                lst(row).Add("")
            Next
        Next
        Return lst
    End Function
    Public Function badCharClean(ByVal line As String) As String
        If line.Contains(vbCrLf) Then
            line = line.Replace(vbCrLf, " ")
        End If
        If line.Contains(vbCr) Then
            line = line.Replace(vbCr, " ")
        End If
        If line.Contains(vbLf) Then
            line = line.Replace(vbLf, " ")
        End If
        If line.Contains(vbTab) Then
            line = line.Replace(vbTab, " ")
        End If
        line = line.Replace("|", ChrW(448)) 'AJW 2012-01-18
        Return line
    End Function
    Private Sub mnuPaste_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPaste.Click, cmnuPasteDictionaryCells.Click, cmnuPasteVarietyCells.Click, cmnuPasteComparisonCells.Click
        If Me.ActiveGrid Is Nothing OrElse Me.ActiveGrid.CurrentCell Is Nothing OrElse Clipboard.GetText = "" Then Return
        'If Clipboard.GetText.Contains("|") Then 'AJW*** cover the pipe situation by replacing any pipe in copy buffer with unicode character so .wsv file is not destroyed
        '    Dim temp As String = Clipboard.GetText()
        '    temp = temp.Replace("|", ChrW(448))
        '    Clipboard.Clear()
        '    Clipboard.SetText(temp)
        'End If
        '        Dim temp55 As String = Clipboard.GetText
        '        For i As Int16 = 0 To 5
        '        Debug.Print(Asc(temp55(i)))
        '        Next
        If Clipboard.GetText.Contains("|") Then 'Or Clipboard.GetText.Contains(vbCrLf) Or Clipboard.GetText.Contains(vbCr) Or Clipboard.GetText.Contains(vbLf) Or Clipboard.GetText.Contains(vbTab) Then 'AJW*** cover the pipe situation by replacing any pipe in copy buffer with unicode character so .wsv file is not destroyed
            Dim temp As String = Clipboard.GetText()
            'temp = badCharClean(temp)
            temp = temp.Replace("|", ChrW(448)) 'AJW 2012-01-18
            Clipboard.Clear()
            Clipboard.SetText(temp)
        End If
        Dim q As String = Clipboard.GetText()
        q = q.Replace(vbLf, vbCrLf)
        If q.EndsWith(vbCrLf) Then
            q = Mid$(q, 1, q.Length - 1)
            Clipboard.Clear()
            Clipboard.SetText(q)
        End If

        'Reimplement edit mode paste.
        If (Not TypeOf Me.ActiveGrid.CurrentCell Is DataGridViewCheckBoxCell) AndAlso Me.ActiveGrid.CurrentCell.IsInEditMode Then
            DirectCast(Me.ActiveGrid.EditingControl, DataGridViewTextBoxEditingControl).SelectedText = Clipboard.GetText
            Return
        End If

        'Find the paste start cell
        Dim startCell As DataGridViewCell = Me.ActiveGrid.CurrentCell
        For Each cell As DataGridViewCell In Me.ActiveGrid.SelectedCells
            If cell.RowIndex < startCell.RowIndex Then
                startCell = cell
            End If
            If cell.ColumnIndex < startCell.ColumnIndex Then
                startCell = cell
            End If
        Next
        Me.ActiveGrid.CurrentCell = startCell

        'Make sure that this paste, if in the dictionary grid, does not create duplicate glosses, and if so, cancel the paste.
        If Me.ActiveGrid Is Me.grdGlossDictionary AndAlso Me.grdGlossDictionary.CurrentCell IsNot Nothing AndAlso Me.grdGlossDictionary.CurrentCell.ColumnIndex = 0 Then
            For Each pasteRow As String In Split(Clipboard.GetText(), vbCrLf)
                Dim pasteValues As String() = Split(pasteRow, vbTab)
                If data.IsGlossInCurrentDictionary(pasteValues(0), "") Then
                    Me.setStatusWarning("This paste operation would create a duplicate of the Gloss """ & pasteValues(0) & """.  This paste has been canceled.", True)
                    Return
                End If
            Next
        End If

        Me.OperationInProgress = True

        Dim firstRowIndex As Integer = Me.ActiveGrid.FirstDisplayedScrollingRowIndex
        'Dim pasteRows As String() = Split(Clipboard.GetText(), vbCrLf)
        Dim pasteRows As String() = Split(Clipboard.GetText.Replace("|", ChrW(448)), vbCrLf)

        'Insert the values into the grid.
        Dim rowIndex As Integer = Me.ActiveGrid.CurrentCell.RowIndex
        For Each pasteRow As String In pasteRows
            Dim colIndex As Integer = Me.ActiveGrid.CurrentCell.ColumnIndex
            For Each pasteValue As String In Split(pasteRow, vbTab)
                Dim thisCell As DataGridViewCell = Me.ActiveGrid.Rows(rowIndex).Cells(colIndex)
                If thisCell.Style.BackColor <> NON_EDITABLE_COLOR Then
                    If Me.ActiveGrid Is Me.grdGlossDictionary AndAlso Me.ActiveGrid.Rows(rowIndex).IsNewRow Then
                        data.InsertNewGloss(rowIndex, "")
                        data.UpdateGlossValue(rowIndex, colIndex, pasteValue)
                        If Me.ActiveGrid Is Me.grdGlossDictionary AndAlso data.GetGlossValue(rowIndex, 0) = "" Then data.UpdateGlossValue(rowIndex, 0, data.GetDefaultGlossName(Me.grdGlossDictionary.CurrentRow.Index))
                        Me.ActiveGrid.RowCount += 1
                    Else
                        thisCell.Value = pasteValue
                    End If
                End If
                colIndex += 1
                If colIndex >= Me.ActiveGrid.ColumnCount Then Exit For
            Next
            rowIndex += 1
            If rowIndex >= Me.ActiveGrid.RowCount Then Exit For
        Next
        If DoLog Then Log.Add("Pasted cells at (" & Me.ActiveGrid.CurrentCell.ColumnIndex.ToString & ", " & Me.ActiveGrid.CurrentCell.RowIndex.ToString & ")")
        Me.RefreshBasedOnCurrentTab()
        Me.ActiveGrid.FirstDisplayedScrollingRowIndex = firstRowIndex
        Me.OperationInProgress = False
        StoreForUndo(data, prefs)
    End Sub
    Private Sub mnuDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelete.Click, cmnuDeleteDictionaryCells.Click, cmnuDeleteVarietyCells.Click, cmnuDeleteComparisonCells.Click
        If Me.ActiveGrid.IsCurrentCellInEditMode Then
            Dim txtBox As DataGridViewTextBoxEditingControl = CType(Me.ActiveGrid.EditingControl, DataGridViewTextBoxEditingControl)
            Dim oldStart As Integer = txtBox.SelectionStart
            If txtBox.SelectionLength = 0 AndAlso txtBox.SelectionStart < txtBox.Text.Length Then
                txtBox.Text = txtBox.Text.Remove(txtBox.SelectionStart, 1)
            Else
                txtBox.Text = txtBox.Text.Remove(txtBox.SelectionStart, txtBox.SelectionLength)
            End If
            txtBox.SelectionStart = oldStart
            Return
        End If

        Me.OperationInProgress = True
        Dim glossDictionaryIndices As New List(Of Integer)
        Dim cellAddresses As New List(Of String)

        For Each cell As DataGridViewCell In Me.ActiveGrid.SelectedCells
            If (Not (Me.ActiveGrid Is Me.grdGlossDictionary And cell.RowIndex = Me.grdGlossDictionary.NewRowIndex)) Then
                If Me.ActiveGrid Is Me.grdGlossDictionary And cell.OwningColumn.Name = "Name" Then
                    glossDictionaryIndices.Add(cell.RowIndex)
                Else
                    cell.Value = ""
                End If
                cellAddresses.Add("(" & cell.OwningColumn.Index.ToString & ", " & cell.OwningRow.Index.ToString & ")")
                If DoLog Then Log.Add("Hit Delete key on " & Me.ActiveGrid.Name & " for cells " & Join(cellAddresses.ToArray, ","))
            End If
        Next
        If glossDictionaryIndices.Count > 0 Then data.DeleteGlossesFromCurrentDictionary(glossDictionaryIndices)
        Me.OperationInProgress = False
        StoreForUndo(data, prefs)
        Me.RefreshBasedOnCurrentTab()
    End Sub
    Private Sub ColumnHeaderClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs)
        Dim grd As DataGridView = DirectCast(sender, DataGridView)

        'Select all the cells in that column if the header is clicked.
        grd.ClearSelection()
        For Each row As DataGridViewRow In grd.Rows
            row.Cells(e.ColumnIndex).Selected = True
        Next
    End Sub

    Private Sub RowHeaderClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs)
        Dim grd As DataGridView = DirectCast(sender, DataGridView)

        'Select all the cells in that row if the header is clicked.
        grd.ClearSelection()
        For Each column As DataGridViewColumn In grd.Columns
            grd.Rows(e.RowIndex).Cells(column.Index).Selected = True
        Next
    End Sub

    Private Sub splTab1A_SplitterMoved(ByVal sender As Object, ByVal e As System.Windows.Forms.SplitterEventArgs)
        prefs.DictionaryPaneWidth = Me.splTab1A.SplitterDistance
        If DoLog Then Log.Add("Moved left first tab splitter to " & prefs.DictionaryPaneWidth.ToString)
    End Sub
    Private Sub splTab1B_SplitterMoved(ByVal sender As Object, ByVal e As System.Windows.Forms.SplitterEventArgs)
        prefs.SurveyPaneWidth = Me.splTab1B.SplitterDistance
        If DoLog Then Log.Add("Moved right first tab splitter to " & prefs.SurveyPaneWidth.ToString)
    End Sub
    Private Sub splComparisons_SplitterMoved(ByVal sender As Object, ByVal e As System.Windows.Forms.SplitterEventArgs)
        prefs.ComparisonPaneWidth = Me.splComparisons.SplitterDistance
        If DoLog Then Log.Add("Moved Comparison tab splitter to " & prefs.ComparisonPaneWidth.ToString)
    End Sub
    Private Sub splCOMPASS_SplitterMoved(ByVal sender As Object, ByVal e As System.Windows.Forms.SplitterEventArgs)
        prefs.COMPASSPaneWidth = Me.splCOMPASS.SplitterDistance
        If DoLog Then Log.Add("Moved COMPASS tab splitter to " & prefs.ComparisonPaneWidth.ToString)
    End Sub
    Private Sub WordSurvForm_Move(ByVal sender As Object, ByVal e As System.EventArgs)
        If Me.WindowState <> FormWindowState.Maximized Then
            prefs.ApplicationX = Me.Left
            prefs.ApplicationY = Me.Top
            If DoLog Then Log.Add("Moved the main form to (" & Me.Left.ToString & ", " & Me.Top.ToString & ")")
        End If
    End Sub
    Private Sub WordSurvForm_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Me.WindowState = FormWindowState.Maximized Then
            prefs.ApplicationIsMaximized = True
            If DoLog Then Log.Add("Resized the main form to (" & Me.Width.ToString & ", " & Me.Height.ToString & ") and maximized.")
        Else
            prefs.ApplicationIsMaximized = False
            prefs.ApplicationWidth = Me.Width
            prefs.ApplicationHeight = Me.Height
            If DoLog Then Log.Add("Resized the main form to (" & Me.Width.ToString & ", " & Me.Height.ToString & ") and unmaximized.")
        End If

        prefs.DictionaryPaneWidth = Me.splTab1A.SplitterDistance
        prefs.SurveyPaneWidth = Me.splTab1B.SplitterDistance
        prefs.ComparisonPaneWidth = Me.splComparisons.SplitterDistance
        prefs.COMPASSPaneWidth = Me.splCOMPASS.SplitterDistance
    End Sub

    Private Sub mnuSetPrimaryFont_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSetPrimaryFont.Click
        Me.CommitGrids()
        Dim frmFontDialog As New FontDialog
        frmFontDialog.ShowApply = True
        frmFontDialog.ShowColor = False
        frmFontDialog.ShowEffects = False
        frmFontDialog.ShowHelp = False
        frmFontDialog.Font = data.PrimaryFont
        Try
            If frmFontDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Try
                    Dim fontTest As New Font(frmFontDialog.Font.Name, frmFontDialog.Font.Size)
                Catch ex As Exception
                    Me.setStatusWarning("Error changing font: " & ex.Message, True)
                    Return
                End Try

                If DoLog Then Log.Add("Set Primary Font to " & frmFontDialog.Font.Name & ", " & frmFontDialog.Font.Size.ToString)
                data.PrimaryFont = New Font(frmFontDialog.Font.Name, frmFontDialog.Font.Size)
                Me.RefreshFonts()
                Me.RefreshBasedOnCurrentTab()
                StoreForUndo(data, prefs)
            End If
        Catch erro As Exception
            If erro.Message = "Only TrueType fonts are supported. This is not a TrueType font." Then
                MsgBox("Font installation problem.  Close down WordSurv, restart WordSurv, then select the font again.")
            End If
        End Try
    End Sub
    Private Sub mnuSetSecondaryFont_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSetSecondaryFont.Click
        Me.CommitGrids()
        Dim frmFontDialog As New FontDialog
        frmFontDialog.ShowApply = True
        frmFontDialog.ShowColor = False
        frmFontDialog.ShowEffects = False
        frmFontDialog.ShowHelp = False
        frmFontDialog.Font = data.SecondaryFont
        Try
            If frmFontDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Try
                    Dim fontTest As New Font(frmFontDialog.Font.Name, frmFontDialog.Font.Size)
                Catch ex As Exception
                    Me.setStatusWarning("Error changing font: " & ex.Message, True)
                    Return
                End Try

                If DoLog Then Log.Add("Set Secondary Font to " & frmFontDialog.Font.Name & ", " & frmFontDialog.Font.Size.ToString)
                data.SecondaryFont = New Font(frmFontDialog.Font.Name, frmFontDialog.Font.Size)
                Me.RefreshFonts()
                Me.RefreshBasedOnCurrentTab()
                StoreForUndo(data, prefs)
            End If
        Catch erro As Exception
            If erro.Message = "Only TrueType fonts are supported. This is not a TrueType font." Then
                MsgBox("Font installation problem.  Close down WordSurv, restart WordSurv, then select the font again.")
            End If
        End Try
    End Sub
    Private Sub mnuSetTranscriptionFont_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSetTranscriptionFont.Click
        Me.CommitGrids()
        Dim frmFontDialog As New FontDialog
        frmFontDialog.ShowApply = True
        frmFontDialog.ShowColor = False
        frmFontDialog.ShowEffects = False
        frmFontDialog.ShowHelp = False
        frmFontDialog.Font = data.TranscriptionFont
        Try
            If frmFontDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Try
                    Dim fontTest As New Font(frmFontDialog.Font.Name, frmFontDialog.Font.Size)
                Catch ex As Exception
                    Me.setStatusWarning("Error changing font: " & ex.Message, True)
                    Return
                End Try

                If DoLog Then Log.Add("Set Transcription Font to " & frmFontDialog.Font.Name & ", " & frmFontDialog.Font.Size.ToString)
                data.TranscriptionFont = New Font(frmFontDialog.Font.Name, frmFontDialog.Font.Size)
                Me.RefreshFonts()
                Me.RefreshBasedOnCurrentTab()
                StoreForUndo(data, prefs)
            End If
        Catch erro As Exception
            If erro.Message = "Only TrueType fonts are supported. This is not a TrueType font." Then
                MsgBox("Font installation problem.  Close down WordSurv, restart WordSurv, then select the font again.")
            End If
        End Try
    End Sub
    Private Sub RefreshFonts()

        'Save these fonts for later so we can reference them in the rowheightneeded event
        Try
            Me.grdGlossDictionary.Font = data.PrimaryFont
        Catch ex As Exception
        End Try
        Me.grdGlossDictionary.Columns("Name").DefaultCellStyle.Font = data.PrimaryFont
        Me.grdGlossDictionary.Columns("Name2").DefaultCellStyle.Font = data.SecondaryFont
        Me.grdGlossDictionary.Columns("PartOfSpeech").DefaultCellStyle.Font = data.PrimaryFont
        Me.grdGlossDictionary.Columns("FieldTip").DefaultCellStyle.Font = data.PrimaryFont
        Me.grdGlossDictionary.Columns("Comments").DefaultCellStyle.Font = data.PrimaryFont

        Try
            Me.grdVariety.Font = data.PrimaryFont
        Catch ex As Exception
        End Try
        Me.grdVariety.Columns("Name").DefaultCellStyle.Font = data.PrimaryFont
        Me.grdVariety.Columns("Transcription").DefaultCellStyle.Font = data.TranscriptionFont
        Me.grdVariety.Columns("PluralFrame").DefaultCellStyle.Font = data.TranscriptionFont
        Me.grdVariety.Columns("Notes").DefaultCellStyle.Font = data.PrimaryFont

        Try
            Me.grdComparisonGloss.Font = data.PrimaryFont
        Catch ex As Exception
        End Try
        Me.grdComparisonGloss.Columns("Name").DefaultCellStyle.Font = data.PrimaryFont

        Try
            Me.grdComparison.Font = data.PrimaryFont
        Catch ex As Exception
        End Try
        Me.grdComparison.Columns("Variety").DefaultCellStyle.Font = data.PrimaryFont
        Me.grdComparison.Columns("Transcription").DefaultCellStyle.Font = data.TranscriptionFont
        Me.grdComparison.Columns("PluralFrame").DefaultCellStyle.Font = data.TranscriptionFont
        Me.grdComparison.Columns("AlignedRendering").DefaultCellStyle.Font = data.TranscriptionFont
        Me.grdComparison.Columns("Grouping").DefaultCellStyle.Font = data.PrimaryFont
        Me.grdComparison.Columns("Notes").DefaultCellStyle.Font = data.PrimaryFont

        Try
            Me.txtVarietyMagnification.Font = New Font(data.TranscriptionFont.Name, 18.0)
            Me.txtComparisonMagnification.Font = New Font(data.TranscriptionFont.Name, 18.0)
        Catch ex As Exception
        End Try

        Try
            Me.grdComparisonAnalysis.Font = data.PrimaryFont
            Me.grdDegreesOfDifference.Font = data.TranscriptionFont
        Catch ex As Exception
        End Try
        Me.grdDegreesOfDifference.DefaultCellStyle.Font = data.PrimaryFont
        Try
            Me.grdPhonoStats.Font = data.PrimaryFont

            Me.grdPhoneCorr.Font = data.PrimaryFont
        Catch ex As Exception
        End Try

        Try
            Me.grdCognateStrengths.Columns("Gloss").DefaultCellStyle.Font = data.PrimaryFont
            Me.grdCognateStrengths.Columns("Form 1").DefaultCellStyle.Font = data.TranscriptionFont
            Me.grdCognateStrengths.Columns("Form 2").DefaultCellStyle.Font = data.TranscriptionFont
            Me.grdCognateStrengths.Columns("Strength").DefaultCellStyle.Font = data.PrimaryFont
        Catch ex As Exception
        End Try

    End Sub
    Private Function GetTallestFontHeight() As Integer
        If data.PrimaryFont Is Nothing Then Return 19
        If data.PrimaryFont.Height > data.SecondaryFont.Height Then
            If data.PrimaryFont.Height > data.TranscriptionFont.Height Then
                Return data.PrimaryFont.Height
            Else
                Return data.TranscriptionFont.Height
            End If
        Else
            If data.SecondaryFont.Height > data.TranscriptionFont.Height Then
                Return data.SecondaryFont.Height
            Else
                Return data.TranscriptionFont.Height
            End If
        End If
    End Function
    Private Sub grdGlossDictionary_RowHeightInfoNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowHeightInfoNeededEventArgs) Handles grdGlossDictionary.RowHeightInfoNeeded
        e.Height = Me.GetTallestFontHeight() + 8
    End Sub
    Private Sub grdVariety_RowHeightInfoNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowHeightInfoNeededEventArgs) Handles grdVariety.RowHeightInfoNeeded
        e.Height = Me.GetTallestFontHeight() + 8
    End Sub
    Private Sub grdComparisonGloss_RowHeightInfoNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowHeightInfoNeededEventArgs) Handles grdComparisonGloss.RowHeightInfoNeeded
        e.Height = Me.GetTallestFontHeight() + 8
    End Sub
    Private Sub grdComparison_RowHeightInfoNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowHeightInfoNeededEventArgs) Handles grdComparison.RowHeightInfoNeeded
        e.Height = Me.GetTallestFontHeight() + 8
    End Sub
    Private Sub grdComparisonAnalysis_RowHeightInfoNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowHeightInfoNeededEventArgs) Handles grdComparisonAnalysis.RowHeightInfoNeeded
        e.Height = Me.GetTallestFontHeight() + 8
    End Sub
    Private Sub grdDegreesOfDifference_RowHeightInfoNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowHeightInfoNeededEventArgs) Handles grdDegreesOfDifference.RowHeightInfoNeeded
        e.Height = Me.GetTallestFontHeight() + 8
    End Sub
    Private Sub grdPhonoStats_RowHeightInfoNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowHeightInfoNeededEventArgs) Handles grdPhonoStats.RowHeightInfoNeeded
        e.Height = Me.GetTallestFontHeight() + 8
    End Sub
    Private Sub grdPhoneCorr_RowHeightInfoNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowHeightInfoNeededEventArgs) Handles grdPhoneCorr.RowHeightInfoNeeded
        e.Height = Me.GetTallestFontHeight() + 8
    End Sub
    Private Sub grdCognateStrengths_RowHeightInfoNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowHeightInfoNeededEventArgs) Handles grdCognateStrengths.RowHeightInfoNeeded
        e.Height = Me.GetTallestFontHeight() + 8
    End Sub

    Private Sub mnuUndo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuUndo.Click
        Me.CommitGrids()
        Dim newData As WordSurvData = Undo(prefs)
        If newData IsNot Nothing Then
            data = newData
            Me.tabWordSurv.SelectTab(prefs.CurrentTab)
            Me.tabWordSurv_SelectedIndexChanged(New Object, New EventArgs) 'Hack, calling SelectTab doesn't trigger this event if the tab doesn't change.
            'Me.RefreshBasedOnCurrentTab()
            If Me.DoingSearchAndReplace Then
                frmSearch.Close()
                mnuFindInGrid_Click(New Object, New EventArgs)
            End If
            If DoLog Then Log.Add("Undo")
        End If
    End Sub
    Private Sub mnuRedo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRedo.Click
        Me.CommitGrids()
        Dim newData As WordSurvData = Redo(prefs)
        If newData IsNot Nothing Then
            data = newData
            Me.tabWordSurv.SelectTab(prefs.CurrentTab)
            Me.tabWordSurv_SelectedIndexChanged(New Object, New EventArgs) 'Hack, calling SelectTab doesn't trigger this event if the tab doesn't change.
            'Me.RefreshBasedOnCurrentTab()
            If Me.DoingSearchAndReplace Then
                frmSearch.Close()
                mnuFindInGrid_Click(New Object, New EventArgs)
            End If
            If DoLog Then Log.Add("Redo")
        End If
    End Sub
    Private Sub mnuNumUndos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNumUndos.Click
        Dim frmInput As New InputForm("Enter Max Number of Undo Changes", "Enter the maximum number of undo changes for WordSurv to store.  To turn off undo for increased program speed, set the maximum number of undo changes to 0.", ValidationType.ZERO_OR_POSITIVE_INTEGER, data, prefs.MaxUndos.ToString)
        frmInput.ShowDialog()
        If frmInput.DialogResult = Windows.Forms.DialogResult.OK Then
            prefs.MaxUndos = Integer.Parse(frmInput.txtInput.Text)
            If DoLog Then Log.Add("Set number of undos to " & prefs.MaxUndos.ToString)
        End If
        InitUndo(data, prefs)
    End Sub

    Private Sub mnuAssociateFileExtension_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAssociateFileExtension.Click
        My.Computer.Registry.CurrentUser.CreateSubKey("Software\Classes\.wsv").SetValue("", "WordSurv", Microsoft.Win32.RegistryValueKind.String)
        My.Computer.Registry.CurrentUser.CreateSubKey("Software\Classes\WordSurv\shell\open\command").SetValue("", Application.ExecutablePath & " ""%l"" ", Microsoft.Win32.RegistryValueKind.String)
        'My.Computer.Registry.CurrentUser.CreateSubKey(".wsv").SetValue("", "WordSurv", Microsoft.Win32.RegistryValueKind.String)
        'My.Computer.Registry.ClassesRoot.CreateSubKey("WordSurv\shell\open\command").SetValue("", Application.ExecutablePath & " ""%l"" ", Microsoft.Win32.RegistryValueKind.String)
    End Sub

    Private Sub mnuPrimaryLanguage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPrimaryLanguage.Click
        Me.CommitGrids()
        Dim frmInput As New InputForm("Set Primary Language", "Enter your primary language.  This will be displayed in the first column of the Dictionary grid.", ValidationType.NOT_EMPTY, data, "")
        frmInput.txtInput.Text = data.PrimaryLanguage
        If frmInput.ShowDialog = Windows.Forms.DialogResult.OK Then
            data.PrimaryLanguage = frmInput.txtInput.Text
            Me.RefreshBasedOnCurrentTab()
            StoreForUndo(data, prefs)
            If DoLog Then Log.Add("Set Primary Language to " & data.PrimaryLanguage)
        End If
    End Sub
    Private Sub mnuSecondaryLanguage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSecondaryLanguage.Click
        Me.CommitGrids()
        Dim frmInput As New InputForm("Set Secondary Language", "Enter your secondary language.  This will be displayed in the second column of the Dictionary grid.", ValidationType.NOT_EMPTY, data, "")
        frmInput.txtInput.Text = data.SecondaryLanguage
        If frmInput.ShowDialog = Windows.Forms.DialogResult.OK Then
            data.SecondaryLanguage = frmInput.txtInput.Text
            Me.RefreshBasedOnCurrentTab()
            StoreForUndo(data, prefs)
            If DoLog Then Log.Add("Set Secondary Language to " & data.SecondaryLanguage)
        End If
    End Sub
    Private Sub mnuFindInGrid_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFindInGrid.Click
        Me.CommitGrids()
        If Not Me.DoingSearchAndReplace Then
            frmSearch = New SearchForm(Me, data, prefs)
            DoingSearchAndReplace = True
            frmSearch.Show()
        End If
        If DoLog Then Log.Add("Started find")
        frmSearch.Focus()
    End Sub

    Private Sub mnuAboutWordSurv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAboutWordSurv.Click
        Me.CommitGrids()
        Dim frmAbout As New AboutWordSurvForm
        frmAbout.ShowDialog()
        If DoLog Then Log.Add("Showed About box")
    End Sub
    Private Sub mnuHelpHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelpHelp.Click
        Me.CommitGrids()
        If System.IO.File.Exists(Application.StartupPath & "\WordSurv_User_Helps.chm") Then
            Help.ShowHelp(Me, Application.StartupPath & "\WordSurv_User_Helps.chm")
        Else
            MsgBox("Could not find help file 'WordSurv_User_Helps.chm'.  Please copy it into the same directory as the WordSurv program.", MsgBoxStyle.Exclamation)
        End If
        If DoLog Then Log.Add("Brought up WordSurv_User_Helps.chm")
    End Sub
    Private Sub mnuWordSurvAtAGlance_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuWordSurvAtAGlance.Click
        Me.CommitGrids()
        Try
            System.Diagnostics.Process.Start(Application.StartupPath & "\WordSurv 7 At a Glance.pdf")
        Catch ex As Exception
            MsgBox("Could not find training file 'WordSurv 7 At a Glance.pdf'.  Please copy it into the same directory as the WordSurv program.", MsgBoxStyle.Exclamation)
        End Try
        If DoLog Then Log.Add("Brought up WordSurv 7 At a Glance.pdf")
    End Sub
    Private Sub mnuHelpQuickReference_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelpQuickReference.Click
        Me.CommitGrids()
        Try
            System.Diagnostics.Process.Start(Application.StartupPath & "\WordSurv 7 Quick Reference Guide.pdf")
        Catch ex As Exception
            MsgBox("Could not find training file 'WordSurv 7 Quick Reference Guide.pdf'.  Please copy it into the same directory as the WordSurv program.", MsgBoxStyle.Exclamation)
        End Try
        If DoLog Then Log.Add("Brought up WordSurv 7 Quick Reference Guide.pdf")
    End Sub
    Private Sub WordSurvForm_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If KillPipeKeyFlag Then
            KillPipeKeyFlag = False
            'e.KeyChar = ChrW(Keys.None)
            e.KeyChar = ChrW(448) 'replace Pipe with Unicode character so .wsv file is not destroyed (uses pipe delimiters)
        End If
    End Sub 'Tells us which dictionary we were copying from, since we switched dictionaries for a copy.
    Private Sub WordSurvForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.Shift And e.KeyValue = Keys.OemPipe Then
            Beep()
            KillPipeKeyFlag = True
        End If
        If e.Shift And e.KeyValue = Keys.Space Then
            e.Handled = True
        End If
        If DoLog Then Log.Add("Hit key " & e.KeyData.ToString)

        If e.KeyValue = Keys.Escape Then
            Dim grid As DataGridView = Nothing
            Select Case Me.tabWordSurv.SelectedTab.Text
                Case "Word List Management"
                    grid = Me.grdGlossDictionary
                Case "Comparisons"
                    grid = Me.grdComparison
                Case "Comparison Analysis"
                    grid = Me.grdComparisonAnalysis
            End Select

            If grid IsNot Nothing Then
                For Each row As DataGridViewRow In grid.Rows
                    If row.DefaultCellStyle.BackColor = CUT_SELECTION Then
                        row.DefaultCellStyle.BackColor = Color.Empty
                    End If
                Next
            End If
            Me.RefreshBasedOnCurrentTab()
        End If

        'Hit control enter to go to the next thing
        If e.KeyCode = Keys.Enter AndAlso e.Control Then
            If Not e.Shift Then
                If Me.ActiveGrid Is Me.grdComparison Then 'go to next gloss
                    If Me.grdComparisonGloss.CurrentCell.RowIndex < Me.grdComparisonGloss.Rows.Count - 1 Then
                        Me.grdComparisonGloss.CurrentCell = Me.grdComparisonGloss.Rows(Me.grdComparisonGloss.CurrentRow.Index + 1).Cells("Name")
                    Else
                        Me.grdComparisonGloss.CurrentCell = Me.grdComparisonGloss.Rows(0).Cells("Name")
                    End If
                    Me.grdComparison.CurrentCell = Me.grdComparison.Rows(0).Cells(data.GetCurrentComparisonsCurrentVarietyColumnIndex())
                End If

                If Me.ActiveGrid Is Me.grdVariety Then 'go to next variety
                    If Me.cboVarieties.SelectedIndex < Me.cboVarieties.Items.Count - 1 Then
                        Me.cboVarieties.SelectedIndex += 1
                    Else
                        Me.cboVarieties.SelectedIndex = 0
                    End If
                    Me.grdVariety.CurrentCell = Me.grdVariety.Rows(0).Cells(data.GetCurrentSurveysCurrentVarietyEntryColumnIndex())
                End If
            Else
                If Me.ActiveGrid Is Me.grdComparison Then 'go to previous gloss
                    If Me.grdComparisonGloss.CurrentCell.RowIndex > 0 Then
                        Me.grdComparisonGloss.CurrentCell = Me.grdComparisonGloss.Rows(Me.grdComparisonGloss.CurrentRow.Index - 1).Cells("Name")
                    Else
                        Me.grdComparisonGloss.CurrentCell = Me.grdComparisonGloss.Rows(Me.grdComparisonGloss.Rows.Count - 1).Cells("Name")
                    End If
                    Me.grdComparison.CurrentCell = Me.grdComparison.Rows(0).Cells(data.GetCurrentComparisonsCurrentVarietyColumnIndex())
                End If

                If Me.ActiveGrid Is Me.grdVariety Then 'go to previous variety
                    If Me.cboVarieties.SelectedIndex > 0 Then
                        Me.cboVarieties.SelectedIndex -= 1
                    Else
                        Me.cboVarieties.SelectedIndex = Me.cboVarieties.Items.Count - 1
                    End If
                    Me.grdVariety.CurrentCell = Me.grdVariety.Rows(0).Cells(data.GetCurrentSurveysCurrentVarietyEntryColumnIndex())
                End If
            End If
        End If

    End Sub

    Private Sub DrawCrosshairs(ByRef grid As DataGridView)
        'This is slow and could probably be optimized somehow.
        For i As Integer = 0 To grid.Rows.Count - 1
            If grid.Rows(i).HeaderCell.Style.BackColor = CROSSHAIRS_COLOR Then
                grid.Rows(i).HeaderCell.Style.BackColor = Color.Empty
            End If
            If grid.Columns(i).HeaderCell.Style.BackColor = CROSSHAIRS_COLOR Then
                grid.Columns(i).HeaderCell.Style.BackColor = Color.Empty
            End If
        Next
        Try
            grid.CurrentCell.OwningRow.HeaderCell.Style.BackColor = CROSSHAIRS_COLOR
            grid.CurrentCell.OwningColumn.HeaderCell.Style.BackColor = CROSSHAIRS_COLOR
        Catch ex As Exception
        End Try
    End Sub
#End Region



#Region "Wordlist Management Tab"
    Private Sub refreshDictionaryPane()
        DoLog = False
        'Any time we change the underlying data structures, we call this function to force paint the grid.
        'Note that force paint is a much faster than force fill (manually adding rows to the grid).

        'Get rid of these things so they don't keep firing while we try to do stuff to the grid and cause all sorts of nasty exceptions.
        RemoveHandler cboGlossDictionaries.SelectedIndexChanged, AddressOf cboGlossDictionaries_SelectedIndexChanged
        RemoveHandler cboGlossDictionarySort.SelectedIndexChanged, AddressOf cboGlossDictionarySort_SelectedIndexChanged
        RemoveHandler grdGlossDictionary.CurrentCellChanged, AddressOf grdGlossDictionary_CurrentCellChanged

        'We have to tell the grid how many rows it has, and apparently clearing out the rows and then telling it how many it has is faster
        'than just changing the number of rows.  This fact brought to you by Google because it was buried deep in the docs.
        Me.cboGlossDictionaries.Items.Clear()
        Me.cboGlossDictionarySort.Items.Clear()
        Me.grdGlossDictionary.Rows.Clear()

        'We refill the combo boxen from scratch each time because it doesn't take all that long and it's easier than micromanaging
        'the combo box elements.
        Dim dictNames As String() = data.GetDictionaryNames()
        If dictNames.Length > 0 Then
            Me.grdGlossDictionary.AllowUserToAddRows = True

            Me.cboGlossDictionaries.Items.AddRange(dictNames)
            Me.cboGlossDictionarySort.Items.AddRange(data.GetCurrentDictionarysSortNames)

            'Haha, these do not cause the associated events to fire for the duration of this function. No more complicated event tree firing parties.
            Me.cboGlossDictionaries.SelectedIndex = data.GetCurrentDictionaryIndex
            Me.cboGlossDictionarySort.SelectedIndex = data.GetCurrentDictionarySortIndex
            Me.grdGlossDictionary.RowCount = data.GetCurrentDictionaryLength() + 1 'Leave an extra row for the bottom row entry thingy.

            Try 'because the current cell might just be nothing.
                Me.grdGlossDictionary.CurrentCell = Me.grdGlossDictionary.Rows(data.GetCurrentDictionarysCurrentGlossIndex).Cells(data.GetCurrentDictionarysCurrentGlossColumnIndex)
            Catch ex As Exception
            End Try
        Else
            'Get thee rid of the bottom row if there are no dictionaries.  Confounded edge cases!
            Me.grdGlossDictionary.AllowUserToAddRows = False
            Me.grdGlossDictionary.RowCount = 0
        End If

        'We are sort of wrapping these function which make the grid actually repaint itself.
        Me.grdGlossDictionary.Refresh()
        Me.RefreshFonts()

        Me.grpTotalGlosses.Text = "Gloss Grid (Total: " & data.GetCurrentDictionaryLength().ToString & ")"

        Me.grdGlossDictionary.Columns("Name").HeaderText = data.PrimaryLanguage
        Me.grdGlossDictionary.Columns("Name2").HeaderText = data.SecondaryLanguage

        Try
            Me.splTab1A.SplitterDistance = prefs.DictionaryPaneWidth
            Me.splTab1B.SplitterDistance = prefs.SurveyPaneWidth
        Catch ex As Exception
        End Try

        AddHandler cboGlossDictionaries.SelectedIndexChanged, AddressOf cboGlossDictionaries_SelectedIndexChanged
        AddHandler cboGlossDictionarySort.SelectedIndexChanged, AddressOf cboGlossDictionarySort_SelectedIndexChanged
        AddHandler grdGlossDictionary.CurrentCellChanged, AddressOf grdGlossDictionary_CurrentCellChanged
        DoLog = True
    End Sub
    Private Sub refreshSurveyPane()
        DoLog = False
        'See refreshDictionaryPane() for comments.
        RemoveHandler cboSurveys.SelectedIndexChanged, AddressOf cboSurveys_SelectedIndexChanged
        RemoveHandler cboVarieties.SelectedIndexChanged, AddressOf cboVarieties_SelectedIndexChanged
        RemoveHandler cboVarietySorts.SelectedIndexChanged, AddressOf cboVarietySorts_SelectedIndexChanged
        RemoveHandler txtSurveyDescription.TextChanged, AddressOf txtSurveyDescription_TextChanged
        RemoveHandler txtVarietyDescription.TextChanged, AddressOf txtVarietyDescription_TextChanged
        RemoveHandler grdVariety.CurrentCellChanged, AddressOf grdVariety_CurrentCellChanged
        RemoveHandler grdVariety.CellEndEdit, AddressOf grdVariety_CellEndEdit

        Me.cboSurveys.Items.Clear()
        Me.cboVarieties.Items.Clear()
        Me.cboVarietySorts.Items.Clear()
        Me.grdVariety.Rows.Clear()
        Me.txtSurveyDescription.Clear()
        Me.txtVarietyDescription.Clear()

        Dim survNames As String() = data.GetSurveyNames()
        If survNames.Length > 0 Then
            Me.cboSurveys.Items.AddRange(survNames)
            Me.cboSurveys.SelectedIndex = data.GetCurrentSurveyIndex()
            Me.txtSurveyDescription.Text = data.GetCurrentSurveyDescription()

            Dim varNames As String() = data.GetCurrentSurveysVarietyNames()
            If varNames.Length > 0 Then
                Me.cboVarieties.Items.AddRange(varNames)
                Me.cboVarieties.SelectedIndex = data.GetCurrentSurveysCurrentVarietyIndex()
                Me.cboVarietySorts.Items.AddRange(data.GetCurrentSurveysSortNames())
                Me.cboVarietySorts.SelectedIndex = data.GetCurrentSurveysCurrentSortIndex()

                Me.grdVariety.RowCount = data.GetCurrentSurveyLength()

                Me.txtVarietyDescription.Text = data.GetCurrentSurveysCurrentVarietyDescription()

                Try
                    Me.grdVariety.CurrentCell = Me.grdVariety.Rows(data.GetCurrentSurveysCurrentGlossIndex()).Cells(data.GetCurrentSurveysCurrentVarietyEntryColumnIndex())
                Catch ex As Exception
                End Try
            Else
                Me.grdVariety.RowCount = 0
            End If
        Else
            Me.grdVariety.RowCount = 0
        End If

        Me.grdVariety.Refresh()
        Me.RefreshFonts()
        Me.RefreshMenus()

        Me.txtAssociatedDictionary.Text = data.GetCurrentSurveysAssociatedDictionaryName()
        Me.UpdateTotalTranscribedLabel()
        Me.lblCurrentVariety.Text = data.GetCurrentSurveysCurrentVarietyName()
        If Me.lblCurrentVariety.Text = "" Then Me.lblCurrentVariety.Text = "Current Variety"
        Try
            Me.UpdateVarietyMagnificationText()
        Catch ex As Exception
        End Try

        AddHandler cboSurveys.SelectedIndexChanged, AddressOf cboSurveys_SelectedIndexChanged
        AddHandler cboVarieties.SelectedIndexChanged, AddressOf cboVarieties_SelectedIndexChanged
        AddHandler cboVarietySorts.SelectedIndexChanged, AddressOf cboVarietySorts_SelectedIndexChanged
        AddHandler txtSurveyDescription.TextChanged, AddressOf txtSurveyDescription_TextChanged
        AddHandler txtVarietyDescription.TextChanged, AddressOf txtVarietyDescription_TextChanged
        AddHandler grdVariety.CurrentCellChanged, AddressOf grdVariety_CurrentCellChanged
        AddHandler grdVariety.CellEndEdit, AddressOf grdVariety_CellEndEdit
        DoLog = True
    End Sub
    Private Sub UpdateVarietyMagnificationText()
        Me.txtVarietyMagnification.Text = safeToString(Me.grdVariety.Rows(Me.grdVariety.CurrentCell.RowIndex).Cells("Transcription").Value)
    End Sub
    Private Sub UpdateComparisonMagnificationText()
        Me.txtComparisonMagnification.Text = safeToString(Me.grdComparison.Rows(Me.grdComparison.CurrentCell.RowIndex).Cells("AlignedRendering").Value)
    End Sub
    Private Sub UpdateTotalTranscribedLabel()
        Me.grpTotalTranscribed.Text = "Transcription Grid (Entered: " & data.GetCurrentVarietysNumberTranscribed().ToString & "/" & data.GetCurrentVarietysTranscriptionCount() & ")"
    End Sub
    Private Sub mnuNewDictionary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNewDictionary.Click
        Me.CommitGrids()
        Me.grdGlossDictionary.EndEdit()
        Me.grdVariety.EndEdit()

        Dim frmInput As New InputForm("New Gloss Dictionary", "Enter the new Gloss Dictionary's name.  This is the name of your elicitation list.", ValidationType.DICTIONARY_NAME, data, "")
        If frmInput.ShowDialog = Windows.Forms.DialogResult.OK Then
            data.CreateNewDictionary(frmInput.Result)
            Me.refreshDictionaryPane()
            Me.RefreshMenus()
            Me.grdGlossDictionary.Focus()

            StoreForUndo(data, prefs)
            If DoLog Then Log.Add("Created New Dictionary " & data.GetCurrentDictionaryName())
        End If

    End Sub
    Private Sub grdGlossDictionary_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles grdGlossDictionary.CellBeginEdit
        Me.grdVariety.EndEdit() 'Prevents weird menu enabling errors where the File and Help menus get permanently turned off
        Me.grdGlossDictionary.Tag = Me.grdGlossDictionary.CurrentCell.Value 'Put the original value of the cell into the tag so we can access it later in the validation functions
    End Sub

    Private Sub grdGlossDictionary_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) ' Handles grdGlossDictionary.CellValidating
        If Me.grdGlossDictionary.IsCurrentCellInEditMode Then
            If e.ColumnIndex = 0 AndAlso data.IsGlossInCurrentDictionary(Me.grdGlossDictionary.EditingControl.Text, Me.grdGlossDictionary.Tag.ToString) Then
                Me.setStatusWarning("That Gloss is already in this Gloss Dictionary.", False)
                e.Cancel = True
            Else
                Me.clearStatusMessage()
            End If
        End If
    End Sub
    Private Sub glossDictionaryGridCellTextChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        If Me.grdGlossDictionary.CurrentCell.ColumnIndex = 0 AndAlso data.IsGlossInCurrentDictionary(Me.grdGlossDictionary.EditingControl.Text, Me.grdGlossDictionary.Tag.ToString) Then
            Me.setStatusWarning("That Gloss is already in this Gloss Dictionary.", True)
        Else
            Me.clearStatusMessage()
        End If
    End Sub
    Public Sub mnuRenameDictionary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRenameDictionary.Click
        Me.CommitGrids()
        Dim frmInput As New InputForm("Rename Gloss Dictionary", "Enter the Gloss Dictionary's new name.", ValidationType.DICTIONARY_NAME, data, data.GetCurrentDictionaryName())

        If frmInput.ShowDialog = Windows.Forms.DialogResult.OK Then
            data.RenameCurrentDictionary(frmInput.Result)
            Me.refreshDictionaryPane()
            StoreForUndo(data, prefs)
            If DoLog Then Log.Add("Renamed Dictionary to " & data.GetCurrentDictionaryName())
        End If
    End Sub
    Private Sub mnuDuplicateDictionary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDuplicateDictionary.Click
        Me.CommitGrids()
        Dim frmInput As New InputForm("Duplicate Gloss Dictionary", "Enter the duplicated Gloss Dictionary's name.", ValidationType.DICTIONARY_NAME, data, "")
        If frmInput.ShowDialog = Windows.Forms.DialogResult.OK Then
            data.DuplicateCurrentDictionary(frmInput.Result)
            Me.refreshDictionaryPane()
            StoreForUndo(data, prefs)
            If DoLog Then Log.Add("Duplicated Dictionary to " & data.GetCurrentDictionaryName())
        End If
    End Sub
    Private Sub mnuDeleteDictionary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDeleteDictionary.Click
        Me.CommitGrids()

        Dim frmConfirm As New ConfirmDeleteDialogBoxForm
        frmConfirm.lblText.Text = "Are you sure you want to delete the Gloss Dictionary" & vbCrLf & """" & data.GetCurrentDictionaryName() & """?"
        If frmConfirm.ShowDialog = Windows.Forms.DialogResult.OK Then
            If DoLog Then Log.Add("Deleted Dictionary " & data.GetCurrentDictionaryName())
            data.DeleteCurrentDictionary()
            Me.refreshDictionaryPane()
            Me.refreshSurveyPane()
            StoreForUndo(data, prefs)
        End If
    End Sub
    Private Sub mnuNewDictionarySort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNewDictionarySort.Click
        Me.CommitGrids()
        Dim frmInput As New InputForm("New Gloss Dictionary Sort", "Enter the new Sort's name.", ValidationType.DICTIONARY_SORT_NAME, data, "")
        If frmInput.ShowDialog = Windows.Forms.DialogResult.OK Then
            data.CreateNewDictionarySort(frmInput.Result)
            Me.refreshDictionaryPane()

            StoreForUndo(data, prefs)
            If DoLog Then Log.Add("Created New Dictionary Sort " & data.GetCurrentDictionarysCurrentSortName())
        End If
    End Sub
    Private Sub mnuRenameDictionarySort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRenameDictionarySort.Click
        Me.CommitGrids()
        Dim frmInput As New InputForm("Rename Gloss Dictionary Sort", "Enter the Sort's new name.", ValidationType.DICTIONARY_SORT_NAME, data, data.GetCurrentDictionarysCurrentSortName())
        If frmInput.ShowDialog = Windows.Forms.DialogResult.OK Then
            data.RenameCurrentDictionarysCurrentSort(frmInput.Result)
            Me.refreshDictionaryPane()

            StoreForUndo(data, prefs)
            If DoLog Then Log.Add("Renamed Dictionary Sort to " & data.GetCurrentDictionarysCurrentSortName())
        End If
    End Sub
    Private Sub mnuDeleteDictionarySort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDeleteDictionarySort.Click
        Me.CommitGrids()

        If data.GetCurrentDictionarysSortNames.Length > 1 Then
            Dim frmConfirm As New ConfirmDeleteDialogBoxForm
            frmConfirm.lblText.Text = "Are you sure you want to delete the Sort" & vbCrLf & """" & data.GetCurrentDictionarysCurrentSortName() & """?"
            If frmConfirm.ShowDialog = Windows.Forms.DialogResult.OK Then

                If DoLog Then Log.Add("Deleted Dictionary Sort " & data.GetCurrentDictionarysCurrentSortName())
                data.DeleteCurrentDictionarySort()
                Me.refreshDictionaryPane()
                Me.refreshSurveyPane()

                StoreForUndo(data, prefs)
            End If
        Else
            Me.setStatusWarning("A Gloss Dictionary needs at least one Sort.", True)
        End If
    End Sub
    Private Sub mnuCutDictionaryRows_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCutDictionaryRows.Click, cmnuCutGlosses.Click

        'Since cutting and pasting have the same shortcut in all 4 ways of cutting and pasting, we have to do this little trick.
        Select Case Me.tabWordSurv.SelectedTab.Text
            Case "Comparisons"
                mnuCutVarieties_Click(sender, e)
            Case "Comparison Analysis"
                mnuCutCAVarieties_Click(sender, e)
            Case "Degrees of Difference"
                mnuCutDDRows_Click(sender, e)
            Case Else
                If Not Me.grdGlossDictionary.Focused Then Return
                Me.CommitGrids()
                For Each row As DataGridViewRow In Me.grdGlossDictionary.Rows
                    row.DefaultCellStyle.BackColor = Color.Empty
                Next
                For Each cell As DataGridViewCell In Me.grdGlossDictionary.SelectedCells
                    Dim row As DataGridViewRow = cell.OwningRow
                    If row.IsNewRow Then Continue For
                    If row.DefaultCellStyle.BackColor <> CUT_SELECTION Then
                        row.DefaultCellStyle.BackColor = CUT_SELECTION
                        If DoLog Then Log.Add("Set Cut Highlight for Dictionary Row " & row.Index)
                    End If

                    For Each rowCell As DataGridViewCell In row.Cells
                        cell.Selected = False
                    Next
                Next
                Me.mnuPasteDictionaryRows.Enabled = True
                Me.cmnuPasteGlosses.Enabled = True
        End Select
    End Sub
    Private Sub mnuPasteDictionaryRows_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPasteDictionaryRows.Click, cmnuPasteGlosses.Click
        Select Case Me.tabWordSurv.SelectedTab.Text
            Case "Comparisons"
                mnuPasteVarieties_Click(sender, e)
            Case "Comparison Analysis"
                mnuPasteCAVarieties_Click(sender, e)
            Case "Degrees of Difference"
                mnuPasteDDRows_Click(sender, e)
            Case Else
                Me.CommitGrids()

                If Me.grdGlossDictionary.RowCount <= 1 Then Return

                Dim rowsMovedCount As Integer

                Me.grdGlossDictionary.Visible = False
                Dim savedRowIndex As Integer = Me.grdGlossDictionary.FirstDisplayedScrollingRowIndex
                Dim indexesOfGlossesToMove As New List(Of Integer)
                For Each row As DataGridViewRow In Me.grdGlossDictionary.Rows
                    If row.DefaultCellStyle.BackColor = CUT_SELECTION Then
                        indexesOfGlossesToMove.Add(row.Index)
                    End If
                Next
                If DoLog Then Log.Add("Pasted Dictionary Rows at row " & Me.grdGlossDictionary.CurrentRow.Index.ToString)

                data.MoveGlosses(indexesOfGlossesToMove, Me.grdGlossDictionary.CurrentRow.Index)
                rowsMovedCount = indexesOfGlossesToMove.Count

                Me.refreshDictionaryPane()
                Me.refreshSurveyPane()
                Me.grdGlossDictionary.FirstDisplayedScrollingRowIndex = savedRowIndex
                Me.grdGlossDictionary.Visible = True

                If indexesOfGlossesToMove.Count <> Me.grdGlossDictionary.Rows.Count - 1 Then 'Crashes if you try to paste every row
                    Me.HighlightRowsAfterPaste(Me.grdGlossDictionary, Me.grdGlossDictionary.CurrentRow.Index, rowsMovedCount)
                Else
                    For Each row As DataGridViewRow In Me.grdGlossDictionary.Rows
                        For Each cell As DataGridViewCell In row.Cells
                            cell.Selected = True
                        Next
                    Next
                End If
                Me.grdGlossDictionary.Focus()
                StoreForUndo(data, prefs)
        End Select
    End Sub
    Private Sub HighlightRowsAfterPaste(ByRef grid As DataGridView, ByVal startRowIndex As Integer, ByVal rowsMovedCount As Integer)
        'Highlight the rows we pasted so the user can see what they did
        For i As Integer = 0 To rowsMovedCount - 1
            For Each cell As DataGridViewCell In grid.Rows(startRowIndex - i).Cells
                cell.Selected = True
            Next
        Next
    End Sub
    Private Sub mnuSortSelectionAlphabetically_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSortSelectionAlphabetically.Click, cmnuSortSelectionAlphabetically.Click
        Me.CommitGrids()
        If data.GetCurrentDictionaryLength = 0 Then Return

        Dim firstCellIndex As Integer = 0
        Dim lastCellIndex As Integer = 0
        Dim thisColumnIndex As Integer = Me.grdGlossDictionary.CurrentCell.ColumnIndex
        For i As Integer = 0 To Me.grdGlossDictionary.Rows.Count - 2
            If Me.grdGlossDictionary.Rows(i).Cells(thisColumnIndex).Selected Then
                firstCellIndex = i
                Exit For
            End If
        Next
        For i As Integer = Me.grdGlossDictionary.Rows.Count - 2 To 0 Step -1
            If Me.grdGlossDictionary.Rows(i).Cells(thisColumnIndex).Selected Then
                lastCellIndex = i
                Exit For
            End If
        Next
        If DoLog Then Log.Add("Sorted selection from Dictionary row " & firstCellIndex.ToString & " to " & lastCellIndex.ToString & " by column " & thisColumnIndex.ToString)
        data.SortCurrentDictionaryAlphabetically(firstCellIndex, lastCellIndex, thisColumnIndex)
        Me.refreshDictionaryPane()
        Me.refreshSurveyPane()

        StoreForUndo(data, prefs)
    End Sub
    '    ~~!@~!@%$@#$^$*&)%
    '                                           /;    ;\                        
    '                                   __  \\____//                        
    '                                  /{_\_/   `'\____                     
    '                                  \___   (o)  (o  }   
    '       _____________________________/          :--' /                  
    '   ,-,'`@@@@@@@@       @@@@@@         \_    `__\                       
    '  ;:(  @@@@@@@@@        @@@             \___(o'o)                      
    '  :: )  @@@@          @@@@@@        ,'@@(  `===='  Moo.           
    '  :: : @@@@@:          @@@@         `@@@:                              
    '  :: \  @@@@@:       @@@@@@@)    (  '@@@'                              
    '  ;; /\      /`,    @@@@@@@@@\   :@@@@@)                               
    '  ::/  )    {_----------------:  :~`,~~;                               
    ' ;;'`; :   )                  :  / `; ;                                
    ';;;; : :   ;                  :  ;  ; :                                
    '`'`' / :  :                   :  :  : :                                
    '    )_ \__;      ";"          :_ ;  \_\       `,','                    
    '    :__\  \    * `,'*         \  \  :  \   *  8`;'*  *                 
    '        `^'     \ :/           `^'  `-^-'   \v/ :  \/   BA        
    '    ~@!$$%&%^&(^(*#^^
    Private Sub mnuInsertRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuInsertRow.Click, cmnuInsertRow.Click
        If Me.ActiveGrid IsNot Me.grdGlossDictionary OrElse Me.grdGlossDictionary.CurrentCell Is Nothing Then Return
        Me.CommitGrids()

        'Insert a new blank row into the data layer and then the grid will update itself for some serious new row action.
        Me.newMiddleRow = True
        If Me.grdGlossDictionary.CurrentCell IsNot Nothing Then
            data.InsertNewGloss(Me.grdGlossDictionary.CurrentCell.RowIndex, "")
        Else
            data.InsertNewGloss(0, "")
        End If
        If Me.grdGlossDictionary.Rows(Me.grdGlossDictionary.CurrentCell.RowIndex).Cells("Name").Value.ToString = "" Then
            data.UpdateGlossValue(Me.grdGlossDictionary.CurrentCell.RowIndex, 0, data.GetDefaultGlossName(Me.grdGlossDictionary.CurrentCell.RowIndex))
        End If

        If DoLog Then Log.Add("Added New Dictionary Row at " & Me.grdGlossDictionary.CurrentCell.RowIndex.ToString)

        Me.refreshDictionaryPane()
        Me.refreshSurveyPane()
        Me.grdGlossDictionary.BeginEdit(True)
    End Sub
    Private Sub mnuDeleteRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDeleteRow.Click, cmnuDeleteRow.Click
        Me.CommitGrids()

        Dim indexesOfGlossesToDelete As New List(Of Integer)

        For Each cell As DataGridViewCell In Me.grdGlossDictionary.SelectedCells
            If Me.grdGlossDictionary.Rows(cell.RowIndex).IsNewRow Then Continue For
            If Not indexesOfGlossesToDelete.Contains(cell.RowIndex) Then
                If DoLog Then Log.Add("Deleted Dictionary Row " & cell.RowIndex)
                indexesOfGlossesToDelete.Add(cell.RowIndex)
            End If

        Next
        If indexesOfGlossesToDelete.Count > 0 Then data.DeleteGlossesFromCurrentDictionary(indexesOfGlossesToDelete)
        Me.refreshDictionaryPane()
        Me.refreshSurveyPane()

        StoreForUndo(data, prefs)
    End Sub
    Private Sub grdGlossDictionary_ColumnWidthChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles grdGlossDictionary.ColumnWidthChanged
        If e.Column.Name = "Name" Then prefs.GlossDictionaryGridNameWidth = e.Column.Width
        If e.Column.Name = "Name2" Then prefs.GlossDictionaryGridName2Width = e.Column.Width
        If e.Column.Name = "PartOfSpeech" Then prefs.GlossDictionaryGridPartOfSpeechWidth = e.Column.Width
        If e.Column.Name = "FieldTip" Then prefs.GlossDictionaryGridFieldTipWidth = e.Column.Width
        If e.Column.Name = "Comments" Then prefs.GlossDictionaryGridCommentsWidth = e.Column.Width
        If DoLog Then Log.Add("Changed Dictionary column " & e.Column.Name & " width to " & e.Column.Width.ToString)
    End Sub
    Private Sub grdVariety_ColumnWidthChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles grdVariety.ColumnWidthChanged
        If e.Column.Name = "Name" Then prefs.VarietyGridNameWidth = e.Column.Width
        If e.Column.Name = "Transcription" Then prefs.VarietyGridTranscriptionWidth = e.Column.Width
        If e.Column.Name = "PluralFrame" Then prefs.VarietyGridPluralFrameWidth = e.Column.Width
        If e.Column.Name = "Notes" Then prefs.VarietyGridNotesWidth = e.Column.Width
        If DoLog Then Log.Add("Changed Variety column " & e.Column.Name & " width to " & e.Column.Width.ToString)
    End Sub
    Private Sub grdGlossDictionary_UserAddedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles grdGlossDictionary.UserAddedRow
        'Unintuitively enough, this event only fires when the bottom row thingy is used.  Therefore we must have two different ways of inserting new rows.
        'No refresh needed here because the grid takes care of that by itself when using the bottom row thingy.
        Me.newBottomRow = True
        data.InsertNewGloss(Me.grdGlossDictionary.CurrentCell.RowIndex, "")
        If DoLog Then Log.Add("Added Dictionary row using bottom row")
    End Sub
    Private Sub grdGlossDictionary_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles grdGlossDictionary.EditingControlShowing
        RemoveHandler e.Control.PreviewKeyDown, AddressOf PreviewKeyDownEvent
        AddHandler e.Control.PreviewKeyDown, AddressOf PreviewKeyDownEvent

        RemoveHandler e.Control.TextChanged, AddressOf glossDictionaryGridCellTextChanged
        AddHandler e.Control.TextChanged, AddressOf glossDictionaryGridCellTextChanged
    End Sub
    Private Sub PreviewKeyDownEvent(ByVal sender As Object, ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs)
        'This was the only event we could find that actually caught the Escape key.  All the others missed it, perhaps because the grid ate it.
        If e.KeyValue = Keys.Escape AndAlso (Me.newMiddleRow Or Me.newBottomRow) Then
            'If we escape out of a new cell, we want to get rid of it in the data layer, since we added a row when we did an insert.
            'Though note that we don't want to delete the cell if we are editing a cell that was already there and the user hit escape.
            Dim indices As New List(Of Integer)
            indices.Add(Me.grdGlossDictionary.CurrentRow.Index)
            Me.grdGlossDictionary.EndEdit()
            Me.clearStatusMessage()
            Me.grdGlossDictionary.Rows.Clear()
            data.DeleteGlossesFromCurrentDictionary(indices)
            Me.refreshDictionaryPane()
        End If
    End Sub
    Private Sub grdGlossDictionary_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grdGlossDictionary.CellEndEdit
        'Only refresh if we inserted using the menu because we get all sorts of random crashing if we refresh after using the bottom row,
        'such as reentrant cell address call super dumb exception.
        'If Me.newMiddleRow Then
        ' Me.refreshDictionaryPane()
        '  End If

        Me.refreshSurveyPane()
        Me.newMiddleRow = False
        Me.newBottomRow = False
        'Me.grdGlossDictionary.MultiSelect = True
        Me.grpTotalGlosses.Text = "Gloss Grid (Total: " & data.GetCurrentDictionaryLength().ToString & ")"
        If DoLog Then Log.Add("End Dictionary Cell Edit")


    End Sub
    Private Sub grdGlossDictionary_CellValuePushed(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs) Handles grdGlossDictionary.CellValuePushed
        'Some operations like inserting rows happen in multiple steps, and those need to store for undo before any of the steps happen.
        'Other operations like updating a cell are single operations and can be stored here.

        'Any time the user changes a grid cell, this event pushes the change to the data layer.
        'AndAlso Not (Me.grdGlossDictionary.EditingControl Is Nothing) added to protect when copying and pasting from secondary to primary
        If e.ColumnIndex = 0 AndAlso Not (Me.grdGlossDictionary.EditingControl Is Nothing) AndAlso data.IsGlossInCurrentDictionary(Me.grdGlossDictionary.EditingControl.Text, Me.grdGlossDictionary.Tag.ToString) Then
            data.UpdateGlossValue(e.RowIndex, e.ColumnIndex, Me.grdGlossDictionary.Tag.ToString)
        Else
            data.UpdateGlossValue(e.RowIndex, e.ColumnIndex, safeToString(e.Value))
        End If
        If Me.grdGlossDictionary.Rows(e.RowIndex).Cells("Name").Value.ToString = "" Then
            data.UpdateGlossValue(e.RowIndex, 0, data.GetDefaultGlossName(e.RowIndex))
        End If
        If Not OperationInProgress Then StoreForUndo(data, prefs)
        If DoLog Then Log.Add("Updated Dictionary value at " & e.ColumnIndex.ToString & ", " & e.RowIndex.ToString & " to " & safeToString(e.Value))
        'End If

    End Sub
    Private Sub grdGlossDictionary_CellValueNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs) Handles grdGlossDictionary.CellValueNeeded
        'Herein lies the greatest thing ever.  We don't need to manually fill the grid because it will only request the values it needs using this event.
        'Instead of refilling the grid, we just repaint it.
        e.Value = data.GetGlossValue(e.RowIndex, e.ColumnIndex)
    End Sub
    Private Sub grdGlossDictionary_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdGlossDictionary.CurrentCellChanged
        If Me.grdGlossDictionary.CurrentCell IsNot Nothing AndAlso Me.grdGlossDictionary.CurrentCell.RowIndex < Me.grdGlossDictionary.RowCount - 1 Then
            data.SetCurrentDictionarysCurrentGloss(Me.grdGlossDictionary.CurrentCell.RowIndex)
            data.SetCurrentDictionarysCurrentGlossColumnIndex(Me.grdGlossDictionary.CurrentCell.ColumnIndex)
            If DoLog Then Log.Add("Changed Dictionary Cell to " & Me.grdGlossDictionary.CurrentCell.OwningColumn.Index.ToString & ", " & Me.grdGlossDictionary.CurrentCell.OwningRow.Index.ToString)
        End If
    End Sub
    Private Sub cboGlossDictionaries_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboGlossDictionaries.SelectedIndexChanged
        Me.grdGlossDictionary.EndEdit()
        data.SetCurrentDictionary(Me.cboGlossDictionaries.SelectedIndex)
        Me.refreshDictionaryPane()
        If DoLog Then Log.Add("Changed Current Dictionary to " & Me.cboGlossDictionaries.SelectedItem.ToString & " (" & Me.cboGlossDictionaries.SelectedIndex.ToString & ")")
    End Sub
    Private Sub cboGlossDictionarySort_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboGlossDictionarySort.SelectedIndexChanged
        Me.grdGlossDictionary.EndEdit()
        data.SetCurrentDictionarysCurrentSort(Me.cboGlossDictionarySort.SelectedIndex)
        Me.refreshDictionaryPane()
        If DoLog Then Log.Add("Changed Current Dictionary Sort to " & Me.cboGlossDictionarySort.SelectedItem.ToString & " (" & Me.cboGlossDictionarySort.SelectedIndex.ToString & ")")
    End Sub
    Private Sub txtSurveyDescription_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSurveyDescription.TextChanged
        data.SetCurrentSurveyDescription(Me.txtSurveyDescription.Text)
    End Sub
    Private Sub txtVarietyDescription_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVarietyDescription.TextChanged
        data.SetCurrentSurveysCurrentVarietyDescription(Me.txtVarietyDescription.Text)
    End Sub
    Private Sub grdVariety_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdVariety.CurrentCellChanged
        If Me.grdVariety.CurrentCell IsNot Nothing Then
            data.SetCurrentSurveysCurrentGloss(Me.grdVariety.CurrentCell.RowIndex)
            data.SetCurrentSurveysCurrentVarietyEntryColumnIndex(Me.grdVariety.CurrentCell.ColumnIndex)
            Me.UpdateVarietyMagnificationText()
            If DoLog Then Log.Add("Changed Variety Cell to " & Me.grdVariety.CurrentCell.OwningColumn.Index.ToString & ", " & Me.grdVariety.CurrentCell.OwningRow.Index.ToString)
        End If
    End Sub
    Private Sub cboSurveys_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSurveys.SelectedIndexChanged
        Me.grdVariety.EndEdit()
        data.SetCurrentSurvey(Me.cboSurveys.SelectedIndex)
        Me.refreshSurveyPane()
        If DoLog Then Log.Add("Changed Current Survey to " & Me.cboSurveys.SelectedItem.ToString & " (" & Me.cboSurveys.SelectedIndex.ToString & ")")
    End Sub
    Private Sub cboVarieties_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboVarieties.SelectedIndexChanged
        Me.grdVariety.EndEdit()
        data.SetCurrentSurveysCurrentVariety(Me.cboVarieties.SelectedIndex)
        Me.refreshSurveyPane()
        If DoLog Then Log.Add("Changed Current Variety to " & Me.cboVarieties.SelectedItem.ToString & " (" & Me.cboVarieties.SelectedIndex.ToString & ")")
    End Sub
    Private Sub cboVarietySorts_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboVarietySorts.SelectedIndexChanged
        Me.grdVariety.EndEdit()
        data.SetCurrentSurveysCurrentSort(Me.cboVarietySorts.SelectedIndex)
        Me.refreshSurveyPane()
        If DoLog Then Log.Add("Changed Current Variety Sort to " & Me.cboVarietySorts.SelectedItem.ToString & " (" & Me.cboVarietySorts.SelectedIndex.ToString & ")")
    End Sub
    Private Sub grdVariety_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grdVariety.CellEndEdit
        Me.UpdateTotalTranscribedLabel()
        If DoLog Then Log.Add("Ended Variety Cell Edit")
    End Sub
    Private Sub grdVariety_CellValuePushed(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs) Handles grdVariety.CellValuePushed
        'Any time the user changes a grid cell, this event pushes the change to the data layer.
        data.UpdateTranscriptionValue(e.RowIndex, e.ColumnIndex, safeToString(e.Value))

        If Not OperationInProgress Then StoreForUndo(data, prefs)
        If DoLog Then Log.Add("Updated Variety value at " & e.ColumnIndex.ToString & ", " & e.RowIndex.ToString & " to " & safeToString(e.Value))
    End Sub
    Private Sub grdVariety_CellValueNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs) Handles grdVariety.CellValueNeeded
        'Herein lies the greatest thing ever.  We don't need to manually fill the grid because it will only request the values it needs using this event.
        'Instead of refilling the grid, we just repaint it.
        e.Value = data.GetTranscriptionValue(e.RowIndex, e.ColumnIndex)
    End Sub
    Private Sub mnuNewSurvey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNewSurvey.Click
        Me.CommitGrids()

        Dim frmInput As New InputForm("New Survey", "Enter the new Survey's name.", ValidationType.SURVEY_NAME, data, "")
        If frmInput.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim frmInput2 As New ComboForm("Select Dictionary", "Select the Dictionary to associate with this Survey", data.GetDictionaryNames())
            If frmInput2.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim frmCreateVarieties As New CreateVarietiesForm
                frmCreateVarieties.ShowDialog()
                If frmCreateVarieties.DialogResult = Windows.Forms.DialogResult.OK Then
                    data.CreateNewSurvey(frmInput.Result, frmInput2.cboSelector.SelectedIndex)
                    For Each varietyName As String In Split(frmCreateVarieties.VarietyNames, vbCrLf)
                        If varietyName = "" Then Continue For
                        data.CreateNewVariety(varietyName.Trim())
                    Next
                    data.SetCurrentSurveysCurrentVariety(0)
                    data.SetCurrentSurveysCurrentGloss(0)
                    Me.refreshSurveyPane()
                    Me.RefreshMenus()

                    Me.grdVariety.Focus()
                    Try
                        Me.grdVariety.CurrentCell = Me.grdVariety.Rows(0).Cells("Transcription")
                    Catch ex As Exception
                    End Try

                    StoreForUndo(data, prefs)
                    If DoLog Then Log.Add("Created New Survey " & data.GetCurrentVarietyName() & " with Varieties " & Join(data.GetCurrentSurveysVarietyNames(), ","))
                End If
            End If
        End If
    End Sub
    Private Sub mnuRenameSurvey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRenameSurvey.Click
        Me.CommitGrids()
        Dim frmInput As New InputForm("Rename Survey", "Enter the Survey's new name.", ValidationType.SURVEY_NAME, data, data.GetCurrentSurveyName())
        If frmInput.ShowDialog = Windows.Forms.DialogResult.OK Then
            data.RenameCurrentSurvey(frmInput.Result)
            Me.refreshSurveyPane()

            StoreForUndo(data, prefs)
            If DoLog Then Log.Add("Renamed Current Survey to " & data.GetCurrentSurveyName())
        End If
    End Sub
    Private Sub mnuDeleteSurvey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDeleteSurvey.Click
        Me.CommitGrids()

        Dim frmConfirm As New ConfirmDeleteDialogBoxForm
        frmConfirm.lblText.Text = "Are you sure you want to delete the Survey" & vbCrLf & """" & data.GetCurrentSurveyName() & """?"
        If frmConfirm.ShowDialog = Windows.Forms.DialogResult.OK Then
            If DoLog Then Log.Add("Deleted Current Survey " & data.GetCurrentSurveyName())
            data.DeleteCurrentSurvey()
            Me.refreshSurveyPane()

            StoreForUndo(data, prefs)
        End If
    End Sub
    Private Sub mnuNewVariety_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNewVariety.Click
        Me.CommitGrids()
        Dim frmInput As New InputForm("New Variety", "Enter the new Variety's name.", ValidationType.VARIETY_NAME, data, "")
        If frmInput.ShowDialog = Windows.Forms.DialogResult.OK Then
            data.CreateNewVariety(frmInput.Result)
            Me.refreshSurveyPane()

            StoreForUndo(data, prefs)
            If DoLog Then Log.Add("Created New Variety " & data.GetCurrentSurveysCurrentVarietyName())
        End If
    End Sub
    Private Sub mnuDeleteVariety_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDeleteVariety.Click
        Me.CommitGrids()

        If data.GetCurrentSurveysVarietyNames().Length = 1 Then
            setStatusWarning("A Survey must have at least one Variety.", True)
            Return
        End If

        Dim frmConfirm As New ConfirmDeleteDialogBoxForm
        frmConfirm.lblText.Text = "Are you sure you want to delete the Variety" & vbCrLf & """" & data.GetCurrentSurveysCurrentVarietyName() & """?"
        If frmConfirm.ShowDialog = Windows.Forms.DialogResult.OK Then
            If DoLog Then Log.Add("Deleted Variety " & data.GetCurrentSurveysCurrentVarietyName())
            data.DeleteCurrentVariety()
            Me.refreshSurveyPane()
            StoreForUndo(data, prefs)
        End If
    End Sub
    Private Sub mnuRenameVariety_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRenameVariety.Click
        Me.CommitGrids()
        Dim frmInput As New InputForm("Rename Variety", "Enter the Variety's new name.", ValidationType.VARIETY_NAME, data, data.GetCurrentVarietyName())
        If frmInput.ShowDialog = Windows.Forms.DialogResult.OK Then
            data.RenameCurrentVariety(frmInput.Result)
            Me.refreshSurveyPane()

            StoreForUndo(data, prefs)
            If DoLog Then Log.Add("Renamed Current Variety to " & data.GetCurrentSurveysCurrentVarietyName())
        End If
    End Sub
    Private Sub grdVariety_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles grdVariety.EditingControlShowing
        If data.GetCurrentSurveysCurrentVarietyEntryColumnIndex = 1 Then
            RemoveHandler CType(e.Control, TextBox).TextChanged, AddressOf grdVarietyTextBoxCellTextChange
            AddHandler CType(e.Control, TextBox).TextChanged, AddressOf grdVarietyTextBoxCellTextChange
        End If
    End Sub
    Private Sub grdVarietyTextBoxCellTextChange(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.txtVarietyMagnification.Text = Me.grdVariety.EditingControl.Text
    End Sub
    Private Sub mnuMergeSurveys_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMergeSurveys.Click
        If data.MergeCurrentSurvey() Then
            Me.refreshDictionaryPane()
            Me.refreshSurveyPane()
            StoreForUndo(data, prefs)
            If DoLog Then Log.Add("Merged Surveys")
        Else
            Me.setStatusWarning("Merge Canceled or Failed", True)
        End If
    End Sub
#End Region

#Region "Comparisons Tab"
    Private Sub txtComparisonDescription_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtComparisonDescription.TextChanged
        data.SetCurrentComparisonDescription(Me.txtComparisonDescription.Text)
    End Sub
    Private Sub refreshComparisonTabLeftPane()
        DoLog = False
        RemoveHandler cboComparison.SelectedIndexChanged, AddressOf cboComparison_SelectedIndexChanged
        RemoveHandler cboComparisonSorts.SelectedIndexChanged, AddressOf cboComparisonSorts_SelectedIndexChanged
        RemoveHandler grdComparisonGloss.CurrentCellChanged, AddressOf grdComparisonGloss_CurrentCellChanged
        'RemoveHandler grdComparisonGloss.CellValueNeeded, AddressOf grdComparisonGloss_CellValueNeeded
        RemoveHandler txtComparisonDescription.TextChanged, AddressOf txtComparisonDescription_TextChanged

        Me.cboComparison.Items.Clear()
        Me.cboComparisonSorts.Items.Clear()
        Me.grdComparisonGloss.Rows.Clear()
        Me.txtComparisonDescription.Clear()

        Dim compNames As String() = data.GetComparisonNames()
        If compNames.Length > 0 Then

            Me.cboComparison.Items.AddRange(compNames)
            Me.cboComparison.SelectedIndex = data.GetCurrentComparisonIndex()

            Me.cboComparisonSorts.Items.AddRange(data.GetCurrentComparisonsSortNames())
            Me.cboComparisonSorts.SelectedIndex = data.GetCurrentComparisonsCurrentSortIndex()

            Me.grdComparisonGloss.RowCount = data.GetCurrentComparisonsGlossCount()

            Me.txtComparisonDescription.Text = data.GetCurrentComparisonDescription()

            Try
                Me.grdComparisonGloss.CurrentCell = Me.grdComparisonGloss.Rows(data.GetCurrentComparisonsCurrentGlossIndex()).Cells("Name")
            Catch ex As Exception
            End Try
        Else
            Me.grdComparisonGloss.RowCount = 0
        End If

        Me.grpTotalComparisonGlosses.Text = "Glosses (Total: " & data.GetCurrentComparisonDictionaryLength().ToString & ")"

        'AddHandler grdComparisonGloss.CellValueNeeded, AddressOf grdComparisonGloss_CellValueNeeded

        Me.grdComparisonGloss.Refresh()
        Me.RefreshFonts()

        Me.splComparisons.SplitterDistance = prefs.ComparisonPaneWidth

        AddHandler cboComparison.SelectedIndexChanged, AddressOf cboComparison_SelectedIndexChanged
        AddHandler cboComparisonSorts.SelectedIndexChanged, AddressOf cboComparisonSorts_SelectedIndexChanged
        AddHandler grdComparisonGloss.CurrentCellChanged, AddressOf grdComparisonGloss_CurrentCellChanged
        AddHandler txtComparisonDescription.TextChanged, AddressOf txtComparisonDescription_TextChanged
        DoLog = True
    End Sub
    Private Sub refreshComparisonTabRightPane()
        DoLog = False
        RemoveHandler grdComparison.CurrentCellChanged, AddressOf grdComparison_CurrentCellChanged

        Me.grdComparison.Rows.Clear()

        If data.GetComparisonNames().Length > 0 AndAlso data.GetCurrentComparisonsGlossCount() > 0 Then
            Me.grdComparison.RowCount = data.GetCurrentComparisonsVarietyCount()
            Try
                Me.grdComparison.CurrentCell = Me.grdComparison.Rows(data.GetCurrentComparisonsCurrentVarietyIndex()).Cells(data.GetCurrentComparisonsCurrentVarietyColumnIndex())
            Catch ex As Exception
            End Try
        Else
            Me.grdComparison.RowCount = 0
        End If

        Try
            Me.grdComparison.Columns("Transcription").HeaderText = "'" & Me.grdComparisonGloss.CurrentCell.Value.ToString & "'"
        Catch ex As Exception
            Me.grdComparison.Columns("Transcription").HeaderText = "Transcription"
        End Try

        Me.grdComparison.Refresh()
        Me.RefreshFonts()
        Me.RefreshMenus()
        Try
            Me.UpdateComparisonMagnificationText()
        Catch ex As Exception
        End Try

        'add some color
        'For Each row As DataGridViewRow In Me.grdComparison.Rows
        '    If DirectCast(row.Cells("Exclude"), DataGridViewCheckBoxCell).Value.ToString = "True" Then
        '        row.DefaultCellStyle.ForeColor = Color.Gray
        '    End If
        'Next

        AddHandler grdComparison.CurrentCellChanged, AddressOf grdComparison_CurrentCellChanged
        DoLog = True
    End Sub
    Private Sub refreshComparisonAnalysisTab()
        DoLog = False
        RemoveHandler rdoTally.CheckedChanged, AddressOf rdoTally_CheckedChanged
        RemoveHandler rdoTotal.CheckedChanged, AddressOf rdoTotal_CheckedChanged
        RemoveHandler rdoPercent.CheckedChanged, AddressOf rdoPercent_CheckedChanged
        RemoveHandler cboComparisonAnalysis.SelectedIndexChanged, AddressOf cboComparisonAnalysis_SelectedIndexChanged

        'This handler doesn't want to remove itself unless you try really hard and remove it twice.
        'This would be a good bug hunting project for the next maintainer.  :P
        RemoveHandler grdComparisonAnalysis.CurrentCellChanged, AddressOf grdComparisonAnalysis_CurrentCellChanged
        RemoveHandler grdComparisonAnalysis.CurrentCellChanged, AddressOf grdComparisonAnalysis_CurrentCellChanged

        Select Case prefs.ComparisonAnalysisMode
            Case "Tally"
                Me.rdoTally.Checked = True
            Case "Total"
                Me.rdoTotal.Checked = True
            Case "Percent"
                Me.rdoPercent.Checked = True
        End Select

        Me.cboComparisonAnalysis.Items.Clear()
        Me.grdComparisonAnalysis.Columns.Clear()

        Dim compNames As String() = data.GetComparisonNames()
        If compNames.Length > 0 Then

            data.DoCurrentComparisonsComparisonAnalysis()

            Me.cboComparisonAnalysis.Items.AddRange(compNames)
            Me.cboComparisonAnalysis.SelectedIndex = data.GetCurrentComparisonIndex()


            For Each varName As String In data.GetCurrentComparisonAnalysisVarietyNames()
                Me.grdComparisonAnalysis.Columns.Add(varName, varName)
            Next
            Me.grdComparisonAnalysis.RowCount = Me.grdComparisonAnalysis.Columns.Count
            Try
                Me.grdComparisonAnalysis.CurrentCell = Me.grdComparisonAnalysis.Rows(data.GetCurrentComparisonAnalysisVarietyIndex()).Cells(data.GetCurrentComparisonAnalysisVarietyColumnIndex())
            Catch ex As Exception
            End Try
            Dim cnt As Integer = 0
            For Each varName As String In data.GetCurrentComparisonAnalysisVarietyNames()
                Me.grdComparisonAnalysis.Rows(cnt).HeaderCell.Value = varName
                cnt += 1
            Next
            'Resize columns
            For Each col As DataGridViewColumn In Me.grdComparisonAnalysis.Columns
                col.Width = prefs.ComparisonAnalysisColumnWidth
            Next

            Me.grdComparisonAnalysis.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders)
            Me.grdComparisonAnalysis.Refresh()
            Me.RefreshFonts()

            If Me.rdoTotal.Checked Then
                Me.ColorizeCATotalGrid()
            Else
                Me.ColorizeGrid(Me.grdComparisonAnalysis)
            End If
        End If


        AddHandler cboComparisonAnalysis.SelectedIndexChanged, AddressOf cboComparisonAnalysis_SelectedIndexChanged
        AddHandler rdoTally.CheckedChanged, AddressOf rdoTally_CheckedChanged
        AddHandler rdoTotal.CheckedChanged, AddressOf rdoTotal_CheckedChanged
        AddHandler rdoPercent.CheckedChanged, AddressOf rdoPercent_CheckedChanged
        AddHandler grdComparisonAnalysis.CurrentCellChanged, AddressOf grdComparisonAnalysis_CurrentCellChanged
        DoLog = True
    End Sub
    Private Sub grdComparisonGloss_ColumnWidthChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles grdComparisonGloss.ColumnWidthChanged
        If e.Column.Name = "Name" Then prefs.ComparisonGlossGridNameWidth = e.Column.Width
        If DoLog Then Log.Add("Changed Comparison Gloss column " & e.Column.Name & " width to " & e.Column.Width.ToString)
    End Sub

    'Used for the feature wherein the user can arrow up and down through the aligned rendering column without going out of edit mode
    Private AlignedRenderingColumnInEditMode As Boolean = False
    Private AlignedRenderingCharIndex As Integer = 0
    Private Sub handleAlignedRenderingArrowUpOrDown(ByVal sender As Object, ByVal e As PreviewKeyDownEventArgs)
        If e.KeyValue = Keys.Up Or e.KeyValue = Keys.Down Then
            Me.AlignedRenderingColumnInEditMode = True
            Me.AlignedRenderingCharIndex = DirectCast(Me.grdComparison.EditingControl, DataGridViewTextBoxEditingControl).SelectionStart
        Else
            Me.AlignedRenderingColumnInEditMode = False
        End If
    End Sub
    Private Sub grdComparison_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grdComparison.CellEnter
        If Me.grdComparison.CurrentCell.OwningColumn.Name = "AlignedRendering" AndAlso AlignedRenderingColumnInEditMode Then
            Me.grdComparison.BeginEdit(False)
            DirectCast(Me.grdComparison.EditingControl, DataGridViewTextBoxEditingControl).SelectionStart = Me.AlignedRenderingCharIndex
        End If
    End Sub

    Private Sub grdComparison_ColumnWidthChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles grdComparison.ColumnWidthChanged
        If e.Column.Name = "Variety" Then prefs.ComparisonGridVarietyWidth = e.Column.Width
        If e.Column.Name = "Transcription" Then prefs.ComparisonGridTranscriptionWidth = e.Column.Width
        If e.Column.Name = "PluralFrame" Then prefs.ComparisonGridPluralFrameWidth = e.Column.Width
        If e.Column.Name = "AlignedRendering" Then prefs.ComparisonGridAlignedRenderingWidth = e.Column.Width
        If e.Column.Name = "Grouping" Then prefs.ComparisonGridGroupingWidth = e.Column.Width
        If e.Column.Name = "Notes" Then prefs.ComparisonGridNotesWidth = e.Column.Width
        If e.Column.Name = "Exclude" Then prefs.ComparisonGridExcludeWidth = e.Column.Width
        If DoLog Then Log.Add("Changed Comparison column " & e.Column.Name & " width to " & e.Column.Width.ToString)
    End Sub
    Private Sub cboComparison_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboComparison.SelectedIndexChanged
        Me.grdComparison.EndEdit()
        data.SetCurrentComparison(Me.cboComparison.SelectedIndex)
        Me.refreshComparisonTabLeftPane()
        Me.refreshComparisonTabRightPane()
        If DoLog Then Log.Add("Changed Current Comparison to " & Me.cboComparison.SelectedItem.ToString & " (" & Me.cboComparison.SelectedIndex.ToString & ")")
    End Sub
    Private Sub cboComparisonSorts_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboComparisonSorts.SelectedIndexChanged
        Me.grdComparison.EndEdit()
        data.SetCurrentComparisonsCurrentGlossSort(Me.cboComparisonSorts.SelectedIndex)
        Me.refreshComparisonTabLeftPane()
        Me.refreshComparisonTabRightPane()
        If DoLog Then Log.Add("Changed Current Comparison Sort to " & Me.cboComparisonSorts.SelectedItem.ToString & " (" & Me.cboComparisonSorts.SelectedIndex.ToString & ")")
    End Sub
    Private Sub grdComparison_CellValuePushed(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs) Handles grdComparison.CellValuePushed
        'Any time the user changes a grid cell, this event pushes the change to the data layer.
        Dim errorStr As String = data.UpdateComparisonEntryValue(e.RowIndex, e.ColumnIndex, safeToString(e.Value))
        If errorStr = "" Then
            If Not OperationInProgress Then StoreForUndo(data, prefs)
        Else
            Me.setStatusWarning(errorStr, True)
        End If
        If DoLog Then Log.Add("Updated Comparison value at " & e.ColumnIndex.ToString & ", " & e.RowIndex.ToString & " to " & safeToString(e.Value))
    End Sub
    Private Sub grdComparisonGloss_CellValueNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs) Handles grdComparisonGloss.CellValueNeeded
        e.Value = data.GetComparisonGlossValue(e.RowIndex, e.ColumnIndex)
    End Sub
    Private Sub grdComparison_CellValueNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs) Handles grdComparison.CellValueNeeded
        e.Value = data.GetComparisonEntryValue(e.RowIndex, e.ColumnIndex)
    End Sub
    Private Sub mnuNewComparison_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNewComparison.Click
        Me.CommitGrids()
        Dim frmInput As New InputForm("New Comparison", "Enter the new Comparison's name.", ValidationType.COMPARISON_NAME, data, "")
        If frmInput.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim frmInput2 As New ComboForm("Select Survey", "Select the Survey that has Varieties you wish to compare.", data.GetSurveyNames())
            If frmInput2.ShowDialog = Windows.Forms.DialogResult.OK Then

                data.CreateNewComparison(frmInput.Result, frmInput2.cboSelector.SelectedIndex)
                data.SetCurrentComparisonsCurrentGloss(0)
                data.SetCurrentComparisonsCurrentVariety(0)
                Me.refreshComparisonTabLeftPane()
                Me.refreshComparisonTabRightPane()
                Me.RefreshMenus()

                Me.grdComparison.Focus()
                Try
                    Me.grdComparison.CurrentCell = Me.grdComparison.Rows(0).Cells("Grouping")
                Catch ex As Exception
                End Try

                StoreForUndo(data, prefs)
                If DoLog Then Log.Add("Created New Comparison " & data.GetCurrentComparisonName() & " for Survey " & frmInput2.cboSelector.SelectedItem.ToString)
            End If
        End If

    End Sub
    Private Sub mnuRenameComparison_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRenameComparison.Click
        Me.CommitGrids()
        Dim frmInput As New InputForm("Rename Comparison", "Enter the Comparison's new name.", ValidationType.COMPARISON_NAME, data, data.GetCurrentComparisonName())
        If frmInput.ShowDialog = Windows.Forms.DialogResult.OK Then
            StoreForUndo(data, prefs)
            'Me.UpdateAutoSave()
            data.RenameCurrentComparison(frmInput.Result)
            If DoLog Then Log.Add("Renamed Current Comparison to " & data.GetCurrentComparisonName())
        End If
        Me.refreshComparisonTabLeftPane()
        Me.refreshComparisonTabRightPane()
    End Sub
    Private Sub mnuDeleteComparison_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDeleteComparison.Click
        Me.CommitGrids()

        Dim frmConfirm As New ConfirmDeleteDialogBoxForm
        frmConfirm.lblText.Text = "Are you sure you want to delete the Comparison" & vbCrLf & """" & data.GetCurrentComparisonName() & """?"
        If frmConfirm.ShowDialog = Windows.Forms.DialogResult.OK Then
            If DoLog Then Log.Add("Deleted Comparison " & data.GetCurrentComparisonName())
            data.DeleteCurrentComparison()
            Me.refreshComparisonTabLeftPane()
            Me.refreshComparisonTabRightPane()

            StoreForUndo(data, prefs)
        End If
    End Sub
    Private Sub grdComparisonGloss_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdComparisonGloss.CurrentCellChanged
        If Me.grdComparisonGloss.CurrentCell IsNot Nothing Then
            data.SetCurrentComparisonsCurrentGloss(Me.grdComparisonGloss.CurrentCell.RowIndex)
            Me.refreshComparisonTabRightPane()
            'Me.grdComparison.CurrentCell.Style.BackColor = INACTIVE_COLOR
            'Me.grdComparison.CurrentCell.Selected = False
            If DoLog Then Log.Add("Changed Comparison Gloss Cell to " & Me.grdComparisonGloss.CurrentCell.OwningColumn.Index.ToString & ", " & Me.grdComparisonGloss.CurrentCell.OwningRow.Index.ToString)
        End If
    End Sub
    Private Sub grdComparison_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdComparison.CurrentCellChanged
        If Me.grdComparison.CurrentCell IsNot Nothing Then
            data.SetCurrentComparisonsCurrentVariety(Me.grdComparison.CurrentCell.RowIndex)
            data.SetCurrentComparisonsCurrentVarietyColumnIndex(Me.grdComparison.CurrentCell.ColumnIndex)
            data.SetCurrentComparisonsAssociatedSurveysCurrentVariety(Me.grdComparison.CurrentCell.RowIndex)
            Me.UpdateComparisonMagnificationText()
            If DoLog Then Log.Add("Changed Comparison Cell to " & Me.grdComparison.CurrentCell.OwningColumn.Index.ToString & ", " & Me.grdComparison.CurrentCell.OwningRow.Index.ToString)
        End If
    End Sub
    Private Sub btnComparisonStatistics_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnComparisonStatistics.Click
        Dim stats As String = data.GetComparisonStatistics()
        MsgBox(stats, MsgBoxStyle.OkOnly, "Comparison Statistics")
        If DoLog Then Log.Add("Got Comparison Statistics: " & stats)
    End Sub
    Private Sub mnuGoToNextUngrouped_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuGoToNextUngrouped.Click
        Me.CommitGrids()
        If Me.grdComparison.RowCount = 0 Then Return
        Dim glossAndVariety As IntIntComboMenu = data.GetNextUngroupedGlossAndVariety()
        Me.grdComparisonGloss.CurrentCell = Me.grdComparisonGloss.Rows(glossAndVariety.Int1).Cells("Name")
        Me.grdComparison.Focus()
        Me.grdComparison.CurrentCell = Me.grdComparison.Rows(glossAndVariety.Int2).Cells("Grouping")
        If DoLog Then Log.Add("Went to next ungrouped")
    End Sub
    Private Sub mnuExcludeAllVarieties_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExcludeAllVarieties.Click, cmnuExcludeAllVarieties.Click
        Me.CommitGrids()
        If Me.grdComparison.RowCount = 0 Then Return
        data.SetExcludeValueForAllVarietiesForCurrentGloss("x")
        Me.refreshComparisonTabRightPane()
        StoreForUndo(data, prefs)
        If DoLog Then Log.Add("Excluded all varieties")
    End Sub
    Private Sub mnuIncludeAllVarieties_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuIncludeAllVarieties.Click, cmnuIncludeAllVarieties.Click
        Me.CommitGrids()
        If Me.grdComparison.RowCount = 0 Then Return
        data.SetExcludeValueForAllVarietiesForCurrentGloss("")
        Me.refreshComparisonTabRightPane()
        StoreForUndo(data, prefs)
        If DoLog Then Log.Add("Included all varieties")
    End Sub
    Private Sub mnuDuplicateComparison_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDuplicateComparison.Click
        Me.CommitGrids()
        Dim frmInput As New InputForm("Name Duplicated Comparison", "Enter a name for the duplicated Comparison.", ValidationType.COMPARISON_NAME, data, "")
        If frmInput.ShowDialog() = Windows.Forms.DialogResult.OK Then
            data.DuplicateCurrentComparison(frmInput.Result)
            Me.refreshComparisonTabLeftPane()
            Me.refreshComparisonTabRightPane()
            StoreForUndo(data, prefs)
        End If
    End Sub
    Private Sub mnuCutVarieties_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCutComparisonRows.Click, cmnuCutVarieties.Click
        If Not Me.grdComparison.Focused Then Return
        Me.CommitGrids()

        Dim selectedRows As New List(Of Integer)
        For Each cell As DataGridViewCell In Me.grdComparison.SelectedCells
            selectedRows.Add(cell.OwningRow.Index)
        Next

        Me.refreshComparisonTabRightPane()

        For Each rowIndex As Integer In selectedRows
            For Each cell As DataGridViewCell In Me.grdComparison.Rows(rowIndex).Cells
                cell.Style.BackColor = Color.Empty
            Next
            If Me.grdComparison.Rows(rowIndex).DefaultCellStyle.BackColor <> CUT_SELECTION Then
                Me.grdComparison.Rows(rowIndex).DefaultCellStyle.BackColor = CUT_SELECTION
                If DoLog Then Log.Add("Set Cut Highlight for Comparison Row " & rowIndex)
            End If
        Next

        For Each cell As DataGridViewCell In Me.grdComparison.SelectedCells
            cell.Selected = False
        Next

        Me.mnuPasteComparisonRows.Enabled = True
        Me.cmnuPasteVarieties.Enabled = True
    End Sub
    Private Sub mnuPasteVarieties_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPasteComparisonRows.Click, cmnuPasteVarieties.Click
        Me.CommitGrids()

        If Me.grdComparison.RowCount = 0 Then Return

        Me.grdComparison.Visible = False
        Dim savedRowIndex As Integer = Me.grdComparison.FirstDisplayedScrollingRowIndex
        Dim indexesOfGlossesToMove As New List(Of Integer)

        Dim rowIndices As New List(Of Integer)
        For Each row As DataGridViewRow In Me.grdComparison.Rows
            If row.DefaultCellStyle.BackColor = CUT_SELECTION Then
                rowIndices.Add(row.Index)
            End If
        Next
        If DoLog Then Log.Add("Pasted Comparison Rows at row " & Me.grdComparison.CurrentRow.Index.ToString)

        data.MoveComparisonVariety(rowIndices, Me.grdComparison.CurrentRow.Index)

        Me.refreshComparisonTabRightPane()
        Me.grdComparison.CurrentCell.Selected = False

        Me.grdComparison.FirstDisplayedScrollingRowIndex = savedRowIndex
        Me.grdComparison.Visible = True

        If rowIndices.Count <> Me.grdComparison.Rows.Count Then 'Crashes if you try to paste every row
            Me.HighlightRowsAfterPaste(Me.grdComparison, Me.grdComparison.CurrentRow.Index, rowIndices.Count)
        Else
            For Each row As DataGridViewRow In Me.grdComparison.Rows
                For Each cell As DataGridViewCell In row.Cells
                    cell.Selected = True
                Next
            Next
        End If
        Me.grdComparison.Focus()
        StoreForUndo(data, prefs)
    End Sub
    Private Sub mnuSetStandardVarietyOrder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSetStandardVarietyOrder.Click, cmnuSetStandardVarietyOrder.Click
        data.SetCurrentComparisonsStandardVarietyOrder()
        Me.setStatusNotification("Variety Order saved", True)
        If DoLog Then Log.Add("Set the standard variety order")
    End Sub
    Private Sub mnuRevertToStandardVarietyOrder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRevertToStandardVarietyOrder.Click, cmnuRevertToStandardVarietyOrder.Click
        data.RevertToCurrentComparisonsStandardVarietyOrder()
        Me.refreshComparisonTabRightPane()
        StoreForUndo(data, prefs)
        If DoLog Then Log.Add("Reverted to the standard variety order")
    End Sub

    Private Sub mnuSortComparisonSelectionAlphabetically_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSortComparisonSelectionAlphabetically.Click, cmnuSortComparisonSelectionAlphabetically.Click
        Me.CommitGrids()
        If data.GetCurrentDictionaryLength = 0 Then Return

        Dim firstCellIndex As Integer = 0
        Dim lastCellIndex As Integer = 0
        Dim thisColumnIndex As Integer = Me.grdComparison.CurrentCell.ColumnIndex
        For i As Integer = 0 To Me.grdComparison.Rows.Count - 1
            If Me.grdComparison.Rows(i).Cells(thisColumnIndex).Selected Then
                firstCellIndex = i
                Exit For
            End If
        Next
        For i As Integer = Me.grdComparison.Rows.Count - 1 To 0 Step -1
            If Me.grdComparison.Rows(i).Cells(thisColumnIndex).Selected Then
                lastCellIndex = i
                Exit For
            End If
        Next
        If DoLog Then Log.Add("Sorted selection from Comparison row " & firstCellIndex.ToString & " to " & lastCellIndex.ToString & " by column " & thisColumnIndex.ToString)
        data.SortCurrentComparisonAlphabetically(firstCellIndex, lastCellIndex, thisColumnIndex)
        Me.refreshComparisonTabRightPane()

        StoreForUndo(data, prefs)
    End Sub
    Private Sub mnuAdvanceToNextVariety_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAdvanceToNextVariety.Click, cmnuAdvanceToNextVariety.Click
        Me.CommitGrids()
        Select Case Me.tabWordSurv.SelectedTab.Text
            Case "Comparisons"
                mnuAdvanceToNextGloss_Click(sender, e)
            Case Else
                If DoLog Then Log.Add("Advanced to next Variety")
                If Me.cboVarieties.SelectedIndex < Me.cboVarieties.Items.Count - 1 Then
                    Me.cboVarieties.SelectedIndex += 1
                Else
                    Me.cboVarieties.SelectedIndex = 0
                End If
                Try
                    Me.grdVariety.CurrentCell = Me.grdVariety.Rows(data.GetCurrentSurveysCurrentGlossIndex()).Cells(data.GetCurrentSurveysCurrentVarietyEntryColumnIndex())
                Catch ex As Exception
                End Try
        End Select
    End Sub
    Private Sub mnuAdvanceToPreviousVariety_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAdvanceToPreviousVariety.Click, cmnuAdvanceToPreviousVariety.Click
        Me.CommitGrids()
        Select Case Me.tabWordSurv.SelectedTab.Text
            Case "Comparisons"
                mnuAdvanceToPreviousGloss_Click(sender, e)
            Case Else
                If DoLog Then Log.Add("Advanced to previous Variety")
                If Me.cboVarieties.SelectedIndex > 0 Then
                    Me.cboVarieties.SelectedIndex -= 1
                Else
                    Me.cboVarieties.SelectedIndex = Me.cboVarieties.Items.Count - 1
                End If
                Try
                    Me.grdVariety.CurrentCell = Me.grdVariety.Rows(data.GetCurrentSurveysCurrentGlossIndex()).Cells(data.GetCurrentSurveysCurrentVarietyEntryColumnIndex())
                Catch ex As Exception
                End Try
        End Select
    End Sub
    Private Sub mnuAdvanceToNextGloss_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAdvanceToNextGloss.Click, cmnuAdvanceToNextGloss.Click
        If Me.grdComparison.RowCount = 0 Then Return
        Me.CommitGrids()

        If DoLog Then Log.Add("Advanced to Next Comparison Gloss")
        If Me.grdComparisonGloss.CurrentCell.RowIndex < Me.grdComparisonGloss.Rows.Count - 1 Then
            Me.grdComparisonGloss.CurrentCell = Me.grdComparisonGloss.Rows(Me.grdComparisonGloss.CurrentRow.Index + 1).Cells("Name")
        Else
            Me.grdComparisonGloss.CurrentCell = Me.grdComparisonGloss.Rows(0).Cells("Name")
        End If
        Try
            Me.grdComparison.CurrentCell = Me.grdComparison.Rows(data.GetCurrentComparisonsCurrentVarietyIndex()).Cells(data.GetCurrentComparisonsCurrentVarietyColumnIndex())
        Catch ex As Exception
        End Try
    End Sub
    Private Sub mnuAdvanceToPreviousGloss_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAdvanceToPreviousGloss.Click, cmnuAdvanceToPreviousGloss.Click
        If Me.grdComparison.RowCount = 0 Then Return
        Me.CommitGrids()

        If DoLog Then Log.Add("Advanced to previous Comparison Gloss")
        If Me.grdComparisonGloss.CurrentCell.RowIndex > 0 Then
            Me.grdComparisonGloss.CurrentCell = Me.grdComparisonGloss.Rows(Me.grdComparisonGloss.CurrentRow.Index - 1).Cells("Name")
        Else
            Me.grdComparisonGloss.CurrentCell = Me.grdComparisonGloss.Rows(Me.grdComparisonGloss.Rows.Count - 1).Cells("Name")
        End If
        Try
            Me.grdComparison.CurrentCell = Me.grdComparison.Rows(data.GetCurrentComparisonsCurrentGlossIndex()).Cells(data.GetCurrentComparisonsCurrentVarietyColumnIndex())
        Catch ex As Exception
        End Try
    End Sub
    Private Sub grdComparison_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles grdComparison.EditingControlShowing
        If data.GetCurrentComparisonsCurrentVarietyColumnIndex = 3 Then
            RemoveHandler CType(e.Control, TextBox).TextChanged, AddressOf grdComparisonTextBoxCellTextChange
            AddHandler CType(e.Control, TextBox).TextChanged, AddressOf grdComparisonTextBoxCellTextChange

            RemoveHandler CType(e.Control, TextBox).PreviewKeyDown, AddressOf handleAlignedRenderingArrowUpOrDown
            AddHandler CType(e.Control, TextBox).PreviewKeyDown, AddressOf handleAlignedRenderingArrowUpOrDown
        End If
    End Sub
    Private Sub grdComparisonTextBoxCellTextChange(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.txtComparisonMagnification.Text = Me.grdComparison.EditingControl.Text
    End Sub

    Private Sub mnuSwapSynonymOrder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSwapSynonymOrder.Click, cmnuSwapSynonymOrder.Click
        OperationInProgress = True
        If Me.grdComparison.CurrentRow Is Nothing OrElse Me.grdComparison.RowCount = 0 Then Return
        Me.grdComparison.CurrentRow.Cells("AlignedRendering").Value = Me.RotateAroundCommas(Me.grdComparison.CurrentRow.Cells("AlignedRendering").Value.ToString)
        Me.grdComparison.CurrentRow.Cells("Grouping").Value = Me.RotateAroundCommas(Me.grdComparison.CurrentRow.Cells("Grouping").Value.ToString)
        If DoLog Then Log.Add("Swapped Synonym order for row " & Me.grdComparison.CurrentRow.Index.ToString)
        StoreForUndo(data, prefs)
        OperationInProgress = False
    End Sub
    Private Function RotateAroundCommas(ByVal str As String) As String
        Dim syns As String() = Split(str, ",")
        Dim temp As String = syns(0)
        For i As Integer = 0 To syns.Length - 2
            syns(i) = syns(i + 1)
        Next
        syns(syns.Length - 1) = temp
        Return Join(syns, ",")
    End Function
#End Region


#Region "Comparison Analysis Tab"
    Private Sub ColorizeCATotalGrid()
        'The Comparison Analysis Total grid is colored differently such that the highest value is green, and 50% of the highest is red
        Dim maxVal As Integer = 0
        For Each row As DataGridViewRow In Me.grdComparisonAnalysis.Rows
            For Each cell As DataGridViewCell In row.Cells
                If Integer.Parse(cell.Value.ToString) > maxVal Then
                    maxVal = Integer.Parse(cell.Value.ToString)
                End If
            Next
        Next
        Dim fiftyPercent As Integer = CInt(maxVal / 2)
        For Each row As DataGridViewRow In Me.grdComparisonAnalysis.Rows
            For Each cell As DataGridViewCell In row.Cells
                cell.Style.BackColor = Me.GetColorFromIntRange(Integer.Parse(cell.Value.ToString), fiftyPercent, maxVal)
            Next
        Next
    End Sub
    Private Sub ColorizeGrid(ByRef grid As DataGridView)
        'Colorize a grid's cells such that the highest value is green and the lowest is red
        Dim maxVal As Integer = 0
        For Each row As DataGridViewRow In grid.Rows
            For Each cell As DataGridViewCell In row.Cells
                If Integer.Parse(cell.Value.ToString) > maxVal Then
                    maxVal = Integer.Parse(cell.Value.ToString)
                End If
            Next
        Next
        For Each row As DataGridViewRow In grid.Rows
            For Each cell As DataGridViewCell In row.Cells
                cell.Style.BackColor = Me.GetColorFromStrength(Integer.Parse(cell.Value.ToString), maxVal)
            Next
        Next
    End Sub
    Private Sub ReverseColorizeGrid(ByRef grid As DataGridView)
        'In some grids, a low number is good and a high one is bad
        Dim maxVal As Integer = 0
        For Each row As DataGridViewRow In grid.Rows
            For Each cell As DataGridViewCell In row.Cells
                If cell.Value.ToString <> "" AndAlso Integer.Parse(cell.Value.ToString) > maxVal Then
                    maxVal = Integer.Parse(cell.Value.ToString)
                End If
            Next
        Next
        For Each row As DataGridViewRow In grid.Rows
            For Each cell As DataGridViewCell In row.Cells
                If cell.Value.ToString <> "" Then
                    cell.Style.BackColor = Me.GetColorFromStrength(maxVal - Integer.Parse(cell.Value.ToString), maxVal)
                End If
            Next
        Next
    End Sub
    Private Function GetColorFromStrength(ByVal val As Integer, ByVal maxVal As Integer) As Color
        If maxVal = 0 Then maxVal = 1
        Dim strength As Double = (val / maxVal) * 2.0 - 1.0
        If strength < 0.0 Then
            Return Color.FromArgb(255, 255, CType(255 + (105 * strength), Integer), 150)
        ElseIf strength > 0.0 Then
            Return Color.FromArgb(255, CType(255 - (105 * strength), Integer), 255, 150)
        Else
            Return Color.FromArgb(255, 255, 255, 150)
        End If
    End Function

    Private Function GetColorFromIntRange(ByVal val As Integer, ByVal lower As Integer, ByVal upper As Integer) As Color
        'Higher than upper is green, lower than lower is red
        If val <= lower Then Return Color.FromArgb(255, 255, 150, 150)
        If val >= upper Then Return Color.FromArgb(255, 150, 255, 150)

        Dim range As Integer = upper - lower
        Dim strength As Double = (val - lower) / range * 2.0 - 1.0
        If strength < 0.0 Then
            Return Color.FromArgb(255, 255, CType(255 + (105 * strength), Integer), 150)
        ElseIf strength > 0.0 Then
            Return Color.FromArgb(255, CType(255 - (105 * strength), Integer), 255, 150)
        Else
            Return Color.FromArgb(255, 255, 255, 150)
        End If
    End Function
    Private Sub grdComparisonAnalysis_CellValueNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs) Handles grdComparisonAnalysis.CellValueNeeded
        If Me.rdoTally.Checked Then
            e.Value = data.GetComparisonAnalysisTallyValue(e.RowIndex, e.ColumnIndex)
        ElseIf Me.rdoTotal.Checked Then
            e.Value = data.GetComparisonAnalysisTotalValue(e.RowIndex, e.ColumnIndex)
        ElseIf Me.rdoPercent.Checked Then
            e.Value = data.GetComparisonAnalysisPercentValue(e.RowIndex, e.ColumnIndex)
        End If
    End Sub
    Private Sub rdoTally_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoTally.CheckedChanged
        prefs.ComparisonAnalysisMode = "Tally"
        Me.refreshComparisonAnalysisTab()
        If DoLog Then Log.Add("Changed Comparison Analysis mode to Tally")
    End Sub
    Private Sub rdoTotal_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoTotal.CheckedChanged
        prefs.ComparisonAnalysisMode = "Total"
        Me.refreshComparisonAnalysisTab()
        If DoLog Then Log.Add("Changed Comparison Analysis mode to Total")
    End Sub
    Private Sub rdoPercent_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoPercent.CheckedChanged
        prefs.ComparisonAnalysisMode = "Percent"
        Me.refreshComparisonAnalysisTab()
        If DoLog Then Log.Add("Changed Comparison Analysis mode to Percent")
    End Sub
    Private Sub mnuCutCAVarieties_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCutCARows.Click, cmnuCutCARows.Click
        If Not Me.grdComparisonAnalysis.Focused Then Return
        Me.CommitGrids()

        Dim selectedRows As New List(Of Integer)
        For Each cell As DataGridViewCell In Me.grdComparisonAnalysis.SelectedCells
            selectedRows.Add(cell.OwningRow.Index)
        Next

        Me.refreshComparisonAnalysisTab()

        For Each rowIndex As Integer In selectedRows
            For Each cell As DataGridViewCell In Me.grdComparisonAnalysis.Rows(rowIndex).Cells
                cell.Style.BackColor = Color.Empty
            Next
            If Me.grdComparisonAnalysis.Rows(rowIndex).DefaultCellStyle.BackColor <> CUT_SELECTION Then
                Me.grdComparisonAnalysis.Rows(rowIndex).DefaultCellStyle.BackColor = CUT_SELECTION
                If DoLog Then Log.Add("Set Cut Highlight for Comparison Analysis Row " & rowIndex)
            End If
        Next

        For Each cell As DataGridViewCell In Me.grdComparisonAnalysis.SelectedCells
            cell.Selected = False
        Next

        Me.mnuPasteCARows.Enabled = True
        Me.cmnuPasteCARows.Enabled = True
    End Sub
    Private Sub mnuPasteCAVarieties_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPasteCARows.Click, cmnuPasteCARows.Click
        Me.CommitGrids()

        If Me.grdComparisonAnalysis.RowCount = 0 Then Return

        Me.grdComparisonAnalysis.Visible = False
        Dim savedRowIndex As Integer = Me.grdComparisonAnalysis.FirstDisplayedScrollingRowIndex
        Dim indexesOfGlossesToMove As New List(Of Integer)

        Dim rowIndices As New List(Of Integer)
        For Each row As DataGridViewRow In Me.grdComparisonAnalysis.Rows
            If row.DefaultCellStyle.BackColor = CUT_SELECTION Then
                rowIndices.Add(row.Index)
            End If
        Next

        If DoLog Then Log.Add("Pasted Comparison Analysis Rows at row " & Me.grdComparisonAnalysis.CurrentRow.Index.ToString)
        data.MoveComparisonAnalysisVariety(rowIndices, Me.grdComparisonAnalysis.CurrentRow.Index)

        Me.refreshComparisonAnalysisTab()
        Me.grdComparisonAnalysis.CurrentCell.Selected = False

        Me.grdComparisonAnalysis.FirstDisplayedScrollingRowIndex = savedRowIndex
        Me.grdComparisonAnalysis.Visible = True

        If rowIndices.Count <> Me.grdComparisonAnalysis.Rows.Count Then 'Crashes if you try to paste every row
            Me.HighlightRowsAfterPaste(Me.grdComparisonAnalysis, Me.grdComparisonAnalysis.CurrentRow.Index, rowIndices.Count)
        Else
            For Each row As DataGridViewRow In Me.grdComparisonAnalysis.Rows
                For Each cell As DataGridViewCell In row.Cells
                    cell.Selected = True
                Next
            Next
        End If
        Me.grdComparisonAnalysis.Focus()
        StoreForUndo(data, prefs)
    End Sub
    Private Sub mnuCASetStandardVarietyOrder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCASetStandardVarietyOrder.Click, cmnuCASetStandardVarietyOrder.Click
        data.SetCurrentComparisonsStandardVarietyOrder()
        If DoLog Then Log.Add("Set the standard variety order")
    End Sub
    Private Sub mnuCARevertToStandardVarietyOrder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCARevertToStandardVarietyOrder.Click, cmnuCARevertToStandardVarietyOrder.Click
        data.RevertToCurrentComparisonsStandardVarietyOrder()
        Me.refreshComparisonAnalysisTab()
        If DoLog Then Log.Add("Reverted to the standard variety order")
    End Sub

    Private Sub cboComparisonAnalysis_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboComparisonAnalysis.SelectedIndexChanged
        data.SetCurrentComparison(Me.cboComparisonAnalysis.SelectedIndex)
        data.DoCurrentComparisonsComparisonAnalysis()
        Me.refreshComparisonAnalysisTab()
        If DoLog Then Log.Add("Changed Current Comparison to " & Me.cboComparisonAnalysis.SelectedItem.ToString & " (" & Me.cboComparisonAnalysis.SelectedIndex.ToString & ")")
    End Sub
    Private Sub grdComparisonAnalysis_ColumnWidthChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles grdComparisonAnalysis.ColumnWidthChanged
        For Each col As DataGridViewColumn In Me.grdComparisonAnalysis.Columns
            col.Width = e.Column.Width
        Next
        prefs.ComparisonAnalysisColumnWidth = e.Column.Width
        If DoLog Then Log.Add("Changed Comparison Analysis columns' width to " & e.Column.Width.ToString)
    End Sub
    Private Sub grdComparisonAnalysis_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdComparisonAnalysis.CurrentCellChanged
        Me.DrawCrosshairs(Me.grdComparisonAnalysis)

        If Me.grdComparisonAnalysis.CurrentCell IsNot Nothing Then
            data.SetCurrentComparisonAnalysisVarietyIndex(Me.grdComparisonAnalysis.CurrentCell.RowIndex)
            data.SetCurrentComparisonAnalysisVarietyColumnIndex(Me.grdComparisonAnalysis.CurrentCell.ColumnIndex)
            If DoLog Then Log.Add("Changed Comparison Analysis Cell to " & Me.grdComparisonAnalysis.CurrentCell.OwningColumn.Index.ToString & ", " & Me.grdComparisonAnalysis.CurrentCell.OwningRow.Index.ToString)
        End If
    End Sub
#End Region

#Region "Degrees of Difference Tab"
    Private Sub grdDegreesOfDifference_CellValuePushed(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs) Handles grdDegreesOfDifference.CellValuePushed
        data.UpdateDDValue(e.RowIndex, e.ColumnIndex, safeToString(e.Value))
        StoreForUndo(data, prefs)
        If DoLog Then Log.Add("Updated DD value at " & e.ColumnIndex.ToString & ", " & e.RowIndex.ToString & " to " & safeToString(e.Value))
    End Sub
    Private Sub grdDegreesOfDifference_CellValueNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs) Handles grdDegreesOfDifference.CellValueNeeded
        Dim cellVal As String
        'If e.RowIndex < e.ColumnIndex Then
        'cellVal = ""
        'Else
        Dim DDvalue As Integer = data.GetDDValue(e.RowIndex, e.ColumnIndex)
        If DDvalue <> -1 Then
            cellVal = DDvalue.ToString
        Else
            cellVal = ""
        End If
        'End If

        e.Value = cellVal
    End Sub
    Private Sub cboDegreesOfDifference_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDegreesOfDifference.SelectedIndexChanged
        Me.grdDegreesOfDifference.EndEdit()
        data.SetCurrentComparison(Me.cboDegreesOfDifference.SelectedIndex)
        Me.refreshDegreesOfDifferenceTab(True)
        If DoLog Then Log.Add("Changed Current Comparison to " & Me.cboDegreesOfDifference.SelectedItem.ToString & " (" & Me.cboDegreesOfDifference.SelectedIndex.ToString & ")")
    End Sub
    Private Sub grdDegreesOfDifference_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdDegreesOfDifference.CurrentCellChanged
        Me.DrawCrosshairs(Me.grdDegreesOfDifference)
        data.SetCurrentComparisonsDDCurrentRowIndex(Me.grdDegreesOfDifference.CurrentRow.Index)
        If DoLog Then Log.Add("Changed DD cell to " & Me.grdDegreesOfDifference.CurrentCell.OwningColumn.Index.ToString & ", " & Me.grdDegreesOfDifference.CurrentCell.OwningRow.Index.ToString)

        RemoveHandler cboDDPhoneUsing.SelectedIndexChanged, AddressOf cboDDPhoneUsing_SelectedIndexChanged
        Me.cboDDPhoneUsing.Items.Clear()
        If Me.grdDegreesOfDifference.CurrentCell.Value.ToString <> "" Then
            Me.cboDDPhoneUsing.Items.AddRange(data.GetCurrentComparisonGlossesNamesUsingThisDDPair(Me.grdDegreesOfDifference.CurrentCell.OwningColumn.HeaderText, Me.grdDegreesOfDifference.CurrentCell.OwningRow.HeaderCell.Value.ToString))
        End If
        If Me.cboDDPhoneUsing.Items.Count > 0 Then Me.cboDDPhoneUsing.SelectedIndex = 0
        AddHandler cboDDPhoneUsing.SelectedIndexChanged, AddressOf cboDDPhoneUsing_SelectedIndexChanged
    End Sub
    Private Sub cboDDPhoneUsing_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDDPhoneUsing.SelectedIndexChanged
        Me.grdDegreesOfDifference.EndEdit()
        Dim compGlIndex As Integer = data.GetComparisonGlossIndexFromDDUsedPhonePair(Me.grdDegreesOfDifference.CurrentCell.OwningColumn.HeaderText, Me.grdDegreesOfDifference.CurrentCell.OwningRow.HeaderCell.Value.ToString, Me.cboDDPhoneUsing.SelectedIndex)
        Me.tabWordSurv.SelectedIndex = 1
        Me.grdComparisonGloss.CurrentCell = Me.grdComparisonGloss.Rows(compGlIndex).Cells("Name")
    End Sub
    Private Sub refreshDegreesOfDifferenceTab(ByVal recalculate As Boolean)
        DoLog = False
        'If data.CurrentComparison Is Nothing Then Return

        RemoveHandler grdDegreesOfDifference.CurrentCellChanged, AddressOf grdDegreesOfDifference_CurrentCellChanged
        RemoveHandler cboDegreesOfDifference.SelectedIndexChanged, AddressOf cboDegreesOfDifference_SelectedIndexChanged
        RemoveHandler cboDegreesOfDifference.SelectedIndexChanged, AddressOf cboDegreesOfDifference_SelectedIndexChanged
        RemoveHandler cboDDPhoneUsing.SelectedIndexChanged, AddressOf cboDDPhoneUsing_SelectedIndexChanged
        'RemoveHandler grdDegreesOfDifference.CellValueNeeded, AddressOf grdDegreesOfDifference_CellValueNeeded
        'RemoveHandler grdDegreesOfDifference.CellValueNeeded, AddressOf grdDegreesOfDifference_CellValueNeeded 'Eh, not sure?  It doesn't seem to want to remove itself the first time.

        Me.cboDegreesOfDifference.Items.Clear()
        Me.grdDegreesOfDifference.Columns.Clear()


        Dim compNames As String() = data.GetComparisonNames()
        If compNames.Length > 0 Then

            If recalculate Then data.CurrentComparison.AssociatedDegreesOfDifference.CalculateUsedChars()

            Me.cboDegreesOfDifference.Items.AddRange(compNames)
            Me.cboDegreesOfDifference.SelectedIndex = data.GetCurrentComparisonIndex()


            For Each usedChar As String In data.GetCurrentComparisonsUsedCharsForDDGrid()
                Me.grdDegreesOfDifference.Columns.Add(usedChar, usedChar)
            Next
            Me.grdDegreesOfDifference.RowCount = Me.grdDegreesOfDifference.Columns.Count
            Dim cnt As Integer = 0
            For Each usedChar As String In data.GetCurrentComparisonsUsedCharsForDDGrid()
                Me.grdDegreesOfDifference.Rows(cnt).HeaderCell.Value = usedChar
                cnt += 1
            Next

            Me.grdDegreesOfDifference.AutoResizeColumns()
            Me.grdDegreesOfDifference.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders)

            For Each row As DataGridViewRow In Me.grdDegreesOfDifference.Rows
                For Each cell As DataGridViewCell In row.Cells
                    If safeToString(cell.Value) = "" Then
                        cell.ReadOnly = True
                        If cell.Style.BackColor <> CUT_SELECTION Then cell.Style.BackColor = EMPTY_GRID_COLOR
                    End If
                Next
            Next

            Me.cboDDPhoneUsing.Items.Clear()
            Try
                Me.grdDegreesOfDifference.CurrentCell = Me.grdDegreesOfDifference.Rows(data.GetCurrentComparisonsDDCurrentRowIndex()).Cells(0)
                If Me.grdDegreesOfDifference.CurrentCell.Value.ToString <> "" Then
                    Me.cboDDPhoneUsing.Items.AddRange(data.GetCurrentComparisonGlossesNamesUsingThisDDPair(Me.grdDegreesOfDifference.CurrentCell.OwningColumn.HeaderText, Me.grdDegreesOfDifference.CurrentCell.OwningRow.HeaderCell.Value.ToString))
                    If Me.cboDDPhoneUsing.Items.Count > 0 Then Me.cboDDPhoneUsing.SelectedIndex = 0
                End If
            Catch ex As Exception
            End Try

            ''Fill the Glosses Using combo box using the selected DD pairs
            'Try
            '    Me.cboDDPhoneUsing.Items.AddRange(data.GetGlossesUsingThisDDPair(Me.grdDegreesOfDifference.CurrentRow.Index, Me.grdDegreesOfDifference.CurrentCell.ColumnIndex))
            'Catch ex As Exception
            'End Try

            Me.RefreshFonts()
            Me.RefreshMenus()
        End If

        'AddHandler grdDegreesOfDifference.CellValueNeeded, AddressOf grdDegreesOfDifference_CellValueNeeded

        Me.grdDegreesOfDifference.Refresh()

        AddHandler cboDegreesOfDifference.SelectedIndexChanged, AddressOf cboDegreesOfDifference_SelectedIndexChanged
        AddHandler grdDegreesOfDifference.CurrentCellChanged, AddressOf grdDegreesOfDifference_CurrentCellChanged
        AddHandler cboDDPhoneUsing.SelectedIndexChanged, AddressOf cboDDPhoneUsing_SelectedIndexChanged
        DoLog = True
    End Sub
    Private Sub mnuCutDDRows_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCutDDRows.Click, cmnuCutDDRows.Click
        If Not Me.grdDegreesOfDifference.Focused Then Return
        Me.CommitGrids()

        Dim selectedRows As New List(Of Integer)
        For Each cell As DataGridViewCell In Me.grdDegreesOfDifference.SelectedCells
            selectedRows.Add(cell.OwningRow.Index)
        Next

        Me.refreshDegreesOfDifferenceTab(False)

        For Each rowIndex As Integer In selectedRows
            For Each cell As DataGridViewCell In Me.grdDegreesOfDifference.Rows(rowIndex).Cells
                cell.Style.BackColor = Color.Empty
            Next
            If Me.grdDegreesOfDifference.Rows(rowIndex).DefaultCellStyle.BackColor <> CUT_SELECTION Then
                Me.grdDegreesOfDifference.Rows(rowIndex).DefaultCellStyle.BackColor = CUT_SELECTION
                If DoLog Then Log.Add("Set Cut Highlight for DD Row " & rowIndex)
            End If
        Next

        For Each cell As DataGridViewCell In Me.grdDegreesOfDifference.SelectedCells
            cell.Selected = False
        Next

        Me.mnuPasteDDRows.Enabled = True
        Me.cmnuPasteDDRows.Enabled = True
    End Sub
    Private Sub mnuPasteDDRows_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPasteDDRows.Click, cmnuPasteDDRows.Click

        Me.CommitGrids()

        If Me.grdDegreesOfDifference.RowCount = 0 Then Return

        Me.grdDegreesOfDifference.Visible = False
        Dim currentRowIndex As Integer = Me.grdDegreesOfDifference.CurrentRow.Index
        Dim savedRowIndex As Integer = Me.grdDegreesOfDifference.FirstDisplayedScrollingRowIndex
        Dim indexesOfGlossesToMove As New List(Of Integer)

        Dim rowIndices As New List(Of Integer)
        For Each row As DataGridViewRow In Me.grdDegreesOfDifference.Rows
            If row.DefaultCellStyle.BackColor = CUT_SELECTION Then
                rowIndices.Add(row.Index)
            End If
        Next

        If DoLog Then Log.Add("Pasted DD Rows at row " & Me.grdDegreesOfDifference.CurrentRow.Index.ToString)
        data.MoveDDRows(rowIndices, Me.grdDegreesOfDifference.CurrentRow.Index)

        Me.refreshDegreesOfDifferenceTab(False)
        Me.grdDegreesOfDifference.CurrentCell.Selected = False

        Me.grdDegreesOfDifference.FirstDisplayedScrollingRowIndex = savedRowIndex
        Me.grdDegreesOfDifference.Visible = True


        If rowIndices.Count <> Me.grdDegreesOfDifference.Rows.Count Then
            Me.HighlightRowsAfterPaste(Me.grdDegreesOfDifference, Me.grdDegreesOfDifference.CurrentRow.Index, rowIndices.Count)
        Else
            For Each row As DataGridViewRow In Me.grdDegreesOfDifference.Rows
                For Each cell As DataGridViewCell In row.Cells
                    cell.Selected = True
                Next
            Next
        End If
        Me.grdDegreesOfDifference.Focus()

        StoreForUndo(data, prefs)
    End Sub
    Private Sub mnuSetIgnoredCharacters_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSetIgnoredCharactersDD.Click, mnuSetIgnoredCharactersComparison.Click, cmnuSetIgnoredCharactersComparison.Click
        Me.CommitGrids()
        Dim temp As String = data.GetCurrentComparisonsDDExcludedChars()
        Dim sep() As Char = temp.ToCharArray
        temp = ""
        For Each element As Char In sep
            temp += element & vbCrLf
        Next
        '        Dim frmInput As New DDExcludeCharsForm(data.TranscriptionFont, data.GetCurrentComparisonsDDExcludedChars())
        Dim frmInput As New DDExcludeCharsForm(data.TranscriptionFont, temp)
        If frmInput.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim cleanedInput As String = frmInput.txtInput.Text.Replace(vbCrLf, "").Replace(" ", "").Replace(vbTab, "")
            data.SetCurrentComparisonsDDExcludedChars(cleanedInput)
            Me.refreshDegreesOfDifferenceTab(True)
            StoreForUndo(data, prefs)
            If DoLog Then Log.Add("Set the DD ignored characters to " & cleanedInput)
        End If
    End Sub
#End Region

#Region "Phonostatistical Analysis Tab"
    Private Sub grdPhonoStats_CellValueNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs) Handles grdPhonoStats.CellValueNeeded
        Dim dataIndex As Integer = 0
        If Me.rdoPhonoStats1.Checked Then dataIndex = 1
        If Me.rdoPhonoStats2.Checked Then dataIndex = 2
        If Me.rdoPhonoStats3.Checked Then dataIndex = 3
        If Me.rdoPhonoStats4.Checked Then dataIndex = 4

        Dim cellVal As String
        Dim phonVal As Integer = data.GetPhonoStatsValue(e.RowIndex, e.ColumnIndex, dataIndex)
        Dim cell As DataGridViewCell = Me.grdPhonoStats.Rows(e.RowIndex).Cells(e.ColumnIndex)

        If phonVal <> -1 Then
            cellVal = phonVal.ToString
        Else
            cellVal = ""
            cell.ReadOnly = True
            cell.Style.BackColor = EMPTY_GRID_COLOR
        End If
        e.Value = cellVal
    End Sub
    Private Sub cboPhonoStats_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPhonoStats.SelectedIndexChanged
        data.SetCurrentComparison(Me.cboPhonoStats.SelectedIndex)
        Me.refreshPhonoStatsTab()
        If DoLog Then Log.Add("Changed Current Comparison to " & Me.cboPhonoStats.SelectedItem.ToString & " (" & Me.cboPhonoStats.SelectedIndex.ToString & ")")
    End Sub
    Private Sub grdPhonoStats_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdPhonoStats.CurrentCellChanged
        Me.DrawCrosshairs(Me.grdPhonoStats)
        If DoLog Then Log.Add("Changed PhonoStats Cell to " & Me.grdPhonoStats.CurrentCell.OwningColumn.Index.ToString & ", " & Me.grdPhonoStats.CurrentCell.OwningRow.Index.ToString)
    End Sub
    Private Sub rdoPhonoStats1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoPhonoStats1.CheckedChanged
        If Me.rdoPhonoStats1.Checked Then
            prefs.PhonoStatsAnalysisMode = "DDNumberOfCorrespondences"
            Me.refreshPhonoStatsTab()
            If DoLog Then Log.Add("Changed PhonoStats display mode to DD Number of Correspondences")
        End If
    End Sub
    Private Sub rdoPhonoStats2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoPhonoStats2.CheckedChanged
        If Me.rdoPhonoStats2.Checked Then
            prefs.PhonoStatsAnalysisMode = "Ratio"
            Me.refreshPhonoStatsTab()
            If DoLog Then Log.Add("Changed PhonoStats display mode to Ratio")
        End If
    End Sub
    Private Sub rdoPhonoStats3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoPhonoStats3.CheckedChanged
        If Me.rdoPhonoStats3.Checked Then
            prefs.PhonoStatsAnalysisMode = "DDSummation"
            Me.refreshPhonoStatsTab()
            If DoLog Then Log.Add("Changed PhonoStats display mode to DD Summation")
        End If
    End Sub
    Private Sub rdoPhonoStats4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoPhonoStats4.CheckedChanged
        If Me.rdoPhonoStats4.Checked Then
            prefs.PhonoStatsAnalysisMode = "CorrespondenceTotals"
            Me.refreshPhonoStatsTab()
            If DoLog Then Log.Add("Changed PhonoStats display mode to Correspondence Totals")
        End If
    End Sub
    Private Sub refreshPhonoStatsTab()
        DoLog = False
        RemoveHandler rdoPhonoStats1.CheckedChanged, AddressOf rdoPhonoStats1_CheckedChanged
        RemoveHandler rdoPhonoStats2.CheckedChanged, AddressOf rdoPhonoStats2_CheckedChanged
        RemoveHandler rdoPhonoStats3.CheckedChanged, AddressOf rdoPhonoStats3_CheckedChanged
        RemoveHandler rdoPhonoStats4.CheckedChanged, AddressOf rdoPhonoStats4_CheckedChanged

        RemoveHandler cboPhonoStats.SelectedIndexChanged, AddressOf cboPhonoStats_SelectedIndexChanged

        For i As Integer = 0 To 100 'FIXME
            RemoveHandler grdPhonoStats.CellValueNeeded, AddressOf grdPhonoStats_CellValueNeeded
        Next
        RemoveHandler grdPhonoStats.CellValueNeeded, AddressOf grdPhonoStats_CellValueNeeded 'Someone needs to investigate why this has to be here twice, also in DDs, and not in COMPASS

        Select Case prefs.PhonoStatsAnalysisMode
            Case "DDNumberOfCorrespondences"
                Me.rdoPhonoStats1.Checked = True
            Case "Ratio"
                Me.rdoPhonoStats2.Checked = True
            Case "DDSummation"
                Me.rdoPhonoStats3.Checked = True
            Case "CorrespondenceTotals"
                Me.rdoPhonoStats4.Checked = True
        End Select

        Me.cboPhonoStats.Items.Clear()
        Me.grdPhonoStats.Columns.Clear()


        Dim compNames As String() = data.GetComparisonNames()
        If compNames.Length > 0 Then

            data.CurrentComparison.AssociatedDegreesOfDifference.CalculateUsedChars()
            data.CurrentComparison.AssociatedDegreesOfDifference.DoAnalysis()

            Me.cboPhonoStats.Items.AddRange(compNames)
            Me.cboPhonoStats.SelectedIndex = data.GetCurrentComparisonIndex()


            If Me.rdoPhonoStats1.Checked Then 'the columns and rows will be by character

                For Each usedChar As String In data.GetCurrentComparisonsUsedCharsForDDGrid()
                    Me.grdPhonoStats.Columns.Add(usedChar, usedChar)
                Next
                Me.grdPhonoStats.RowCount = Me.grdPhonoStats.Columns.Count
                Dim cnt As Integer = 0
                For Each usedChar As String In data.GetCurrentComparisonsUsedCharsForDDGrid()
                    Me.grdPhonoStats.Rows(cnt).HeaderCell.Value = usedChar
                    'Me.grdPhonoStats.RowHeadersWidth = Me.grdPhonoStats.RowTemplate.Height * 2
                    cnt += 1
                Next
                'For Each col As DataGridViewColumn In Me.grdPhonoStats.Columns
                'col.Width = Me.grdPhonoStats.RowTemplate.Height
                'Next

            Else 'the columns and rows will be by variety

                For Each varName As String In data.GetCurrentComparisonAnalysisVarietyNames()
                    Me.grdPhonoStats.Columns.Add(varName, varName)
                Next
                Me.grdPhonoStats.RowCount = Me.grdPhonoStats.Columns.Count
                Dim cnt As Integer = 0
                For Each varName As String In data.GetCurrentComparisonAnalysisVarietyNames()
                    Me.grdPhonoStats.Rows(cnt).HeaderCell.Value = varName
                    cnt += 1
                Next

                'Resize columns as per the comparison analysis so these grids are similarly spaced
                'For Each col As DataGridViewColumn In Me.grdPhonoStats.Columns
                '    col.Width = prefs.ComparisonAnalysisColumnWidth
                'Next
            End If

            Me.grdPhonoStats.AutoResizeColumns()
            Me.grdPhonoStats.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders)

            Me.RefreshFonts()
            Me.RefreshMenus()
        End If

        AddHandler grdPhonoStats.CellValueNeeded, AddressOf grdPhonoStats_CellValueNeeded

        Me.grdPhonoStats.Refresh()
        Me.ReverseColorizeGrid(Me.grdPhonoStats)

        AddHandler rdoPhonoStats1.CheckedChanged, AddressOf rdoPhonoStats1_CheckedChanged
        AddHandler rdoPhonoStats2.CheckedChanged, AddressOf rdoPhonoStats2_CheckedChanged
        AddHandler rdoPhonoStats3.CheckedChanged, AddressOf rdoPhonoStats3_CheckedChanged
        AddHandler rdoPhonoStats4.CheckedChanged, AddressOf rdoPhonoStats4_CheckedChanged

        AddHandler cboPhonoStats.SelectedIndexChanged, AddressOf cboPhonoStats_SelectedIndexChanged
        'Me.grdPhonoStats.Visible = True
        'Me.ResumeLayout()
        DoLog = True
    End Sub
#End Region



#Region "COMPASS Tab"
    Private Sub grdCognateStrengths_ColumnWidthChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles grdCognateStrengths.ColumnWidthChanged
        If e.Column.Name = "Gloss" Then prefs.CognateStrengthsGridGlossWidth = e.Column.Width
        If e.Column.Name = "Form 1" Then prefs.CognateStrengthsGridForm1Width = e.Column.Width
        If e.Column.Name = "Form 2" Then prefs.CognateStrengthsGridForm2Width = e.Column.Width
        If e.Column.Name = "Strength" Then prefs.CognateStrengthsGridStrengthWidth = e.Column.Width
        If DoLog Then Log.Add("Changed Cognate Strengths column " & e.Column.Name & " width to " & e.Column.Width.ToString)
    End Sub
    Private Sub grdPhoneCorr_CellValueNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs) Handles grdPhoneCorr.CellValueNeeded
        If data.Comparisons.Count = 0 Then Return
        Dim occurs As Integer = data.GetCOMPASSPhoneOccurences(e.RowIndex, e.ColumnIndex)
        Dim strength As Double = data.GetCOMPASSPhoneStrength(e.RowIndex, e.ColumnIndex)
        If occurs = -1 Or strength = Double.NaN Then
            e.Value = ""
            Me.grdPhoneCorr.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = EMPTY_GRID_COLOR
        Else
            If prefs.CorrespondenceDisplayMode = "ShowCounts" Then e.Value = occurs
            If prefs.CorrespondenceDisplayMode = "ShowStrengths" Then e.Value = strength.ToString("F2")
            Me.grdPhoneCorr.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = Me.GetColorFromStrength(strength)
        End If
    End Sub
    Private Sub grdCognateStrengths_CellValueNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs) Handles grdCognateStrengths.CellValueNeeded
        e.Value = data.GetCOMPASSCognateStrengthsValue(e.RowIndex, e.ColumnIndex)
        Me.grdCognateStrengths.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = Me.GetColorFromStrength(data.GetCOMPASSCognateStrengthsAverageStrength(e.RowIndex))
    End Sub
    Private Sub cboCOMPASS_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCOMPASS.SelectedIndexChanged
        data.SetCurrentComparison(Me.cboCOMPASS.SelectedIndex)
        Me.refreshCOMPASSTab()
        If DoLog Then Log.Add("Changed Current Comparison to " & Me.cboCOMPASS.SelectedItem.ToString & " (" & Me.cboCOMPASS.SelectedIndex.ToString & ")")
    End Sub
    Private Sub cboCOMPASSVariety1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCOMPASSVariety1.SelectedIndexChanged
        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        data.SetCurrentComparisonsCurrentCOMPASSVariety1(Me.cboCOMPASSVariety1.SelectedIndex)
        prefs.COMPASSVariety1Index = Me.cboCOMPASSVariety1.SelectedIndex
        Me.refreshCOMPASSTab()
        'Me.grdCognateStrengths.ClearSelection() 'AJW***
        'Me.grdPhoneCorr.ClearSelection() 'AJW***
        'Me.grdPhoneCorr.Rows(0).Cells(0).Selected = True
        'Dim send As Object = Nothing
        'Dim em As System.Windows.Forms.MouseEventArgs = Nothing
        'grdPhoneCorr_MouseUp(send, em)
        'Me.Refresh() 'AJW***
        If DoLog Then Log.Add("Changed COMPASS Variety 1 to " & Me.cboCOMPASSVariety1.SelectedItem.ToString & " (" & Me.cboCOMPASSVariety1.SelectedIndex.ToString & ")")
        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub cboCOMPASSVariety2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCOMPASSVariety2.SelectedIndexChanged
        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        data.SetCurrentComparisonsCurrentCOMPASSVariety2(Me.cboCOMPASSVariety2.SelectedIndex)
        prefs.COMPASSVariety2Index = Me.cboCOMPASSVariety2.SelectedIndex
        Me.refreshCOMPASSTab()
        'Me.grdCognateStrengths.ClearSelection() 'AJW***
        'Me.grdPhoneCorr.ClearSelection() 'AJW***
        'Me.grdPhoneCorr.Rows(0).Cells(0).Selected = True
        'Dim send As Object = Nothing
        'Dim em As System.Windows.Forms.MouseEventArgs = Nothing
        'grdPhoneCorr_MouseUp(send, em)
        'Me.Refresh() 'AJW***
        If DoLog Then Log.Add("Changed COMPASS Variety 2 to " & Me.cboCOMPASSVariety2.SelectedItem.ToString & " (" & Me.cboCOMPASSVariety2.SelectedIndex.ToString & ")")
        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub nudCOMPASSUpper_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        prefs.COMPASSUpper = CType(Me.nudCOMPASSUpper.Value, Integer)
        Me.refreshCOMPASSTab()
        If DoLog Then Log.Add("Changed COMPASS Upper Threshold to " & Me.nudCOMPASSUpper.Value.ToString)
    End Sub
    Private Sub nudCOMPASSLower_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        prefs.COMPASSLower = CType(Me.nudCOMPASSLower.Value, Integer)
        Me.refreshCOMPASSTab()
        If DoLog Then Log.Add("Changed COMPASS Lower Threshold to " & Me.nudCOMPASSLower.Value.ToString)
    End Sub
    Private Sub nudCOMPASSBottom_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        prefs.COMPASSBottom = CType(Me.nudCOMPASSBottom.Value, Integer)
        Me.refreshCOMPASSTab()
        If DoLog Then Log.Add("Changed COMPASS Bottom Threshold to " & Me.nudCOMPASSBottom.Value.ToString)
    End Sub
    Private Sub rdoShowCounts_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoShowCounts.CheckedChanged
        If Me.rdoShowCounts.Checked Then
            prefs.CorrespondenceDisplayMode = "ShowCounts"
            Me.refreshCOMPASSTab()
            If DoLog Then Log.Add("Changed COMPASS display mode to Show Counts")
        End If
    End Sub
    Private Sub rdoShowStrengths_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoShowStrengths.CheckedChanged
        If Me.rdoShowStrengths.Checked Then
            prefs.CorrespondenceDisplayMode = "ShowStrengths"
            Me.refreshCOMPASSTab()
            If DoLog Then Log.Add("Changed COMPASS display mode to Show Strengths")
        End If
    End Sub
    Private Sub refreshCOMPASSTab()

        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        DoLog = False
        'If data.CurrentComparison Is Nothing Then Return

        RemoveHandler cboCOMPASS.SelectedIndexChanged, AddressOf cboCOMPASS_SelectedIndexChanged
        RemoveHandler cboCOMPASSVariety1.SelectedIndexChanged, AddressOf cboCOMPASSVariety1_SelectedIndexChanged
        RemoveHandler cboCOMPASSVariety2.SelectedIndexChanged, AddressOf cboCOMPASSVariety2_SelectedIndexChanged
        RemoveHandler nudCOMPASSUpper.ValueChanged, AddressOf nudCOMPASSUpper_ValueChanged
        RemoveHandler nudCOMPASSLower.ValueChanged, AddressOf nudCOMPASSLower_ValueChanged
        RemoveHandler nudCOMPASSBottom.ValueChanged, AddressOf nudCOMPASSBottom_ValueChanged
        RemoveHandler grdPhoneCorr.CellValueNeeded, AddressOf grdPhoneCorr_CellValueNeeded
        RemoveHandler grdCognateStrengths.CellValueNeeded, AddressOf grdCognateStrengths_CellValueNeeded
        RemoveHandler rdoShowCounts.CheckedChanged, AddressOf rdoShowCounts_CheckedChanged
        RemoveHandler rdoShowStrengths.CheckedChanged, AddressOf rdoShowStrengths_CheckedChanged
        RemoveHandler grdCognateStrengths.CurrentCellChanged, AddressOf grdCognateStrengths_CurrentCellChanged

        Select Case prefs.CorrespondenceDisplayMode
            Case "ShowCounts"
                Me.rdoShowCounts.Checked = True
            Case "ShowStrengths"
                Me.rdoShowStrengths.Checked = True
        End Select

        Me.cboCOMPASS.Items.Clear()
        Me.cboCOMPASSVariety1.Items.Clear()
        Me.cboCOMPASSVariety2.Items.Clear()
        Me.grdPhoneCorr.Columns.Clear()
        Me.grdCognateStrengths.Rows.Clear()

        Me.grdCognateStrengths.Columns("Form 1").HeaderText = ""
        Me.grdCognateStrengths.Columns("Form 2").HeaderText = ""

        Dim compNames As String() = data.GetComparisonNames()
        If compNames.Length > 0 Then

            Me.cboCOMPASS.Items.AddRange(compNames)
            Me.cboCOMPASS.SelectedIndex = data.GetCurrentComparisonIndex()

            Me.cboCOMPASSVariety1.Items.AddRange(data.GetCurrentComparisonsVarietyNames())
            Me.cboCOMPASSVariety2.Items.AddRange(data.GetCurrentComparisonsVarietyNames())

            Try
                Me.cboCOMPASSVariety1.SelectedIndex = data.GetCurrentComparisonsCurrentCOMPASSVariety1Index()
                Me.cboCOMPASSVariety2.SelectedIndex = data.GetCurrentComparisonsCurrentCOMPASSVariety2Index()
            Catch ex As Exception
                Try
                    Me.cboCOMPASSVariety1.SelectedIndex = 0
                    Me.cboCOMPASSVariety2.SelectedIndex = 1
                Catch ex2 As Exception
                End Try
            End Try

            Try
                Me.grdCognateStrengths.Columns("Form 1").HeaderText = Me.cboCOMPASSVariety1.SelectedItem.ToString
                Me.grdCognateStrengths.Columns("Form 2").HeaderText = Me.cboCOMPASSVariety2.SelectedItem.ToString
            Catch ex As Exception
            End Try

            data.CurrentComparison.AssociatedDegreesOfDifference.CalculateUsedChars()
            data.CurrentComparison.AssociatedDegreesOfDifference.DoAnalysis()
            data.CalculateCOMPASSValues(prefs, Me.cboCOMPASSVariety1.SelectedIndex, Me.cboCOMPASSVariety2.SelectedIndex, CType(Me.nudCOMPASSUpper.Value, Integer), CType(Me.nudCOMPASSLower.Value, Integer), CType(Me.nudCOMPASSBottom.Value, Integer))

            Dim usedChars As List(Of String) = data.GetCurrentComparisonsCOMPASSCalculationUsedChars()

            'Add rows and columns for every character seen, in alphabetical order.
            For Each uniqueChar As String In usedChars
                Me.grdPhoneCorr.Columns.Add(uniqueChar, uniqueChar)
            Next
            Me.grdPhoneCorr.RowCount = Me.grdPhoneCorr.Columns.Count
            Dim i As Integer = 0
            For Each uniqueChar As String In usedChars
                Me.grdPhoneCorr.Rows(i).HeaderCell.Value = uniqueChar
                i += 1
            Next


            Dim cellCoords As List(Of CellAddress) = data.GetSelectedCOMPASSPhoneCoordinates
            Try
                For Each addr As CellAddress In cellCoords
                    Me.grdPhoneCorr.Rows(addr.RowIndex).Cells(addr.ColIndex).Selected = True
                Next
            Catch ex As Exception
            End Try
            Try
                data.FilterCOMPASSStrengthsGrid(cellCoords)
            Catch ex As Exception
                data.ClearSelectedCOMPASSPhoneCoordinates()
            End Try
            Me.grdCognateStrengths.RowCount = data.GetCurrentComparisonsCOMPASSGlossComparedCount()
            Try
                Dim curAddr As CellAddress = data.GetCurrentCOMPASSStrengthsSummaryCellAddress()
                Me.grdCognateStrengths.CurrentCell = Me.grdCognateStrengths.Rows(curAddr.RowIndex).Cells(curAddr.ColIndex)
            Catch ex As Exception
            End Try

            Try
                Me.grdPhoneCorr.CurrentCell = Me.grdPhoneCorr.Rows(data.GetCurrentComparisonsCurrentCOMPASSChar1Index()).Cells(data.GetCurrentComparisonsCurrentCOMPASSChar2Index())
            Catch ex As Exception
            End Try



            'Size the cells so they aren't too wide.
            Me.grdPhoneCorr.AutoResizeColumns()
            Me.grdPhoneCorr.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders)

            Me.RefreshFonts()
            Me.RefreshMenus()

            Try
                Me.lblCOMPASSVariety1.Text = Me.cboCOMPASSVariety2.SelectedItem.ToString
                Me.lblCOMPASSVariety2.Text = Me.MakeVertical(Me.cboCOMPASSVariety1.SelectedItem.ToString)
            Catch ex As Exception
            End Try
        End If

        AddHandler grdCognateStrengths.CellValueNeeded, AddressOf grdCognateStrengths_CellValueNeeded
        AddHandler grdPhoneCorr.CellValueNeeded, AddressOf grdPhoneCorr_CellValueNeeded

        Me.grdPhoneCorr.Refresh()
        Me.grdCognateStrengths.Refresh()


        Me.splCOMPASS.SplitterDistance = prefs.COMPASSPaneWidth

        AddHandler cboCOMPASS.SelectedIndexChanged, AddressOf cboCOMPASS_SelectedIndexChanged
        AddHandler cboCOMPASSVariety1.SelectedIndexChanged, AddressOf cboCOMPASSVariety1_SelectedIndexChanged
        AddHandler cboCOMPASSVariety2.SelectedIndexChanged, AddressOf cboCOMPASSVariety2_SelectedIndexChanged
        AddHandler nudCOMPASSUpper.ValueChanged, AddressOf nudCOMPASSUpper_ValueChanged
        AddHandler nudCOMPASSLower.ValueChanged, AddressOf nudCOMPASSLower_ValueChanged
        AddHandler nudCOMPASSBottom.ValueChanged, AddressOf nudCOMPASSBottom_ValueChanged
        AddHandler rdoShowCounts.CheckedChanged, AddressOf rdoShowCounts_CheckedChanged
        AddHandler rdoShowStrengths.CheckedChanged, AddressOf rdoShowStrengths_CheckedChanged
        AddHandler grdCognateStrengths.CurrentCellChanged, AddressOf grdCognateStrengths_CurrentCellChanged
        DoLog = True

        'Me.grdCognateStrengths.ClearSelection() 'AJW***
        'Me.grdPhoneCorr.ClearSelection() 'AJW***
        'Me.grdCognateStrengths.Rows(0).Cells(0).Selected = True
        'Me.grdCognateStrengths.Rows(0).Cells(0).Selected = True
        Dim send As Object = Nothing
        Dim em As System.EventArgs = Nothing
        grdCognateStrengths_CurrentCellChanged(send, em)
        Me.Refresh() 'AJW***
        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

    End Sub
    Public Function MakeVertical(ByVal str As String) As String
        Dim vertStr As String = ""
        For Each ch As Char In str
            vertStr &= ch & vbCrLf
        Next
        Return vertStr
    End Function
    Private Function getColorFromStrength(ByVal strength As Double) As Color
        'Get a color for the grid with negative numbers being more red, positive numbers being more green, and around zero being yellow.
        'In a 4-tuple representing colors, the first number is the alpha (transparency) value, which we always want to be 255 (solidly opaque).
        'The second number is the red amount, the third number is the green amount, and the fourth number is the blue amount.
        'By keeping the red value the same and changing the green value, we can get a color that goes from red to yellow,
        'and by keeping the green value the same and changing the red value, we can get a color that goes from yellow to green.
        If strength > 3.0 Or strength < -3.0 Then
            Dim x As Integer = 1 + 1
        End If
        If strength < 0.0 Then
            Return Color.FromArgb(255, 255, 255 + CType((105.0 * strength), Integer), 150)
        ElseIf strength > 0.0 Then
            Return Color.FromArgb(255, 255 - CType((105 * strength), Integer), 255, 150)
        Else
            Return Color.FromArgb(255, 255, 255, 150)
        End If
    End Function
    Private Sub btnWordStrengthsSummary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWordStrengthsSummary.Click
        Dim frmStrengthsSummary As New StrengthsSummaryForm
        frmStrengthsSummary.fillSummaryGrid(data.GetCurrentComparisonsCOMPASSStrengthSummary())
        frmStrengthsSummary.ShowDialog()
        If DoLog Then Log.Add("Viewed Word Strengths Summary Popup")
    End Sub
    Private Sub grdPhoneCorr_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdPhoneCorr.CurrentCellChanged
        Me.DrawCrosshairs(Me.grdPhoneCorr)
        If DoLog Then Log.Add("Changed PhoneCorr Cell to " & Me.grdPhoneCorr.CurrentCell.OwningColumn.Index.ToString & ", " & Me.grdPhoneCorr.CurrentCell.OwningRow.Index.ToString)
    End Sub
    Private Sub mnuAdjustThresholds_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAdjustThresholds.Click
        RemoveHandler nudCOMPASSUpper.ValueChanged, AddressOf nudCOMPASSUpper_ValueChanged
        RemoveHandler nudCOMPASSLower.ValueChanged, AddressOf nudCOMPASSLower_ValueChanged
        RemoveHandler nudCOMPASSBottom.ValueChanged, AddressOf nudCOMPASSBottom_ValueChanged
        Try
            Me.nudCOMPASSBottom.Value = CType(Math.Round(1.0 * Math.Log10(CType(data.GetCurrentComparisonsCOMPASSCognateCount(), Double)) - 1.0), Integer)
            Me.nudCOMPASSLower.Value = CType(Math.Round(2.0 * Math.Log10(CType(data.GetCurrentComparisonsCOMPASSCognateCount(), Double)) - 1.0), Integer)
            Me.nudCOMPASSUpper.Value = CType(Math.Round(15.0 * Math.Log10(CType(data.GetCurrentComparisonsCOMPASSCognateCount(), Double)) - 1.0), Integer)
            If DoLog Then Log.Add("Auto-Adjusted COMPASS Thresholds")
        Catch ex As Exception
            MsgBox("Adjustment failed.  You probably do not have enough cognates between these two varieties.  Normally at least 15 are needed for accurate calculations.", MsgBoxStyle.Information)
        End Try
        AddHandler nudCOMPASSUpper.ValueChanged, AddressOf nudCOMPASSUpper_ValueChanged
        AddHandler nudCOMPASSLower.ValueChanged, AddressOf nudCOMPASSLower_ValueChanged
        AddHandler nudCOMPASSBottom.ValueChanged, AddressOf nudCOMPASSBottom_ValueChanged
        Me.refreshCOMPASSTab()
    End Sub

    Private Sub grdPhoneCorr_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles grdPhoneCorr.MouseUp
        If grdPhoneCorr.RowCount = 0 Then Return

        Dim selectedCoords As New List(Of CellAddress)
        data.ClearSelectedCOMPASSPhoneCoordinates()
        For Each cell As DataGridViewCell In Me.grdPhoneCorr.SelectedCells
            If cell.Style.BackColor <> EMPTY_GRID_COLOR Then
                Dim thisAddr As New CellAddress(Nothing, cell.RowIndex, cell.ColumnIndex)
                selectedCoords.Add(thisAddr)
                data.AddSelectedCOMPASSPhoneCoordinate(thisAddr)
            Else
                cell.Selected = False
            End If
        Next
        data.FilterCOMPASSStrengthsGrid(selectedCoords)
        Me.grdCognateStrengths.RowCount = data.GetCurrentComparisonsCOMPASSGlossComparedCount()
        Me.grdCognateStrengths.Refresh()
        If DoLog Then Log.Add("Mouse up on PhoneCorr grid to filter the cognate strengths grid")
    End Sub
    Private Sub grdCognateStrengths_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdCognateStrengths.CurrentCellChanged
        Try
            Dim addr As New CellAddress(0, Me.grdCognateStrengths.CurrentCell.RowIndex, Me.grdCognateStrengths.CurrentCell.ColumnIndex)
            data.SetCurrentCOMPASSStrengthsSummaryCellAddress(addr)
            data.SetCurrentcomparisonsAssociatedDictionaryCurrentGloss(Me.grdCognateStrengths.CurrentCell.RowIndex)
            'MsgBox(Me.grdCognateStrengths.CurrentCell.RowIndex)
            If DoLog Then Log.Add("Changed Cognate Strengths Cell to " & Me.grdCognateStrengths.CurrentCell.OwningColumn.Index.ToString & ", " & Me.grdCognateStrengths.CurrentCell.OwningRow.Index.ToString)
        Catch ex As Exception
        End Try
    End Sub
#End Region
    '                                                  `T",.`-, 
    '                                                     '8, :. 
    '                                              `""`oooob."T,. 
    '                                            ,-`".)O;8:doob.'-. 
    '                                     ,..`'.'' -dP()d8O8Yo8:,..`, 
    '                                   -o8b-     ,..)doOO8:':o; `Y8.`, 
    '                                  ,..bo.,.....)OOO888o' :oO.  ".  `-. 
    '                                , "`"d....88OOOOO8O88o  :O8o;.    ;;,b 
    '                               ,dOOOOO""""""""O88888o:  :O88Oo.;:o888d  
    '                               ""888Ob...,-- :o88O88o:. :o'"""""""Y8OP 
    '                               d8888.....,.. :o8OO888:: ::              
    '                              "" .dOO8bo`'',,;O88O:O8o: ::,              \
    '                                 ,dd8".  ,-)do8O8o:"""; :::               That's a lot of code!
    '                                 ,db(.  T)8P:8o:::::    ::: 
    '                                 -"",`(;O"KdOo::        ::: 
    '                                  ,K,'".doo:::'        :o: 
    '                                    .doo:::"""::  :.    'o: 
    '        ,..            .;ooooooo..o:"""""     ::;. ::;.  'o. 
    '   ,, "'    ` ..   .d;o:"""'                  ::o:;::o::  :; 
    '   d,         , ..ooo::;                      ::oo:;::o"'.:o 
    '  ,d'.       :OOOOO8Oo::" '.. .               ::o8Ooo:;  ;o: 
    '  'P"   ;  ;.OPd8888O8::;. 'oOoo:.;..         ;:O88Ooo:' O"' 
    '  ,8:   o::oO` 88888OOo:::  o8O8Oo:::;;     ,;:oO88OOo;  ' 
    ' ,YP  ,::;:O:  888888o::::  :8888Ooo::::::::::oo888888o;. , 
    ' ',d: :;;O;:   :888888::o;  :8888888Ooooooooooo88888888Oo; , 
    ' dPY:  :o8O     YO8888O:O:;  O8888888888OOOO888"" Y8o:O88o; , 
    ',' O:  'ob`      "8888888Oo;;o8888888888888'"'     `8OO:.`OOb . 
    ''  Y:  ,:o:       `8O88888OOoo"""""""""""'           `OOob`Y8b` 
    '   ::  ';o:        `8O88o:oOoP                        `8Oo `YO. 
    '   `:   Oo:         `888O::oP                          88O  :OY 
    '    :o; 8oP         :888o::P                           do:  8O: 
    '   ,ooO:8O'       ,d8888o:O'                          dOo   ;:. 
    '   ;O8odo'        88888O:o'                          do::  oo.: 
    '  d"`)8O'         "YO88Oo'                          "8O:   o8b' 
    ' ''-'`"            d:O8oK  -hrr-                   dOOo'  :o": 
    '                   O:8o:b.                        :88o:   `8:, 
    '                   `8O:;7b,.                       `"8'     Y: 
    '                    `YO;`8b' 
    '                     `Oo; 8:. 
    '                      `OP"8.` 
    '                       :  Y8P 
    '                       `o  `, 
    '                        Y8bod. 
    '                        `""""' 


    Private Sub mnuEditSurveyMetadata_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEditSurveyMetadata.Click
        Dim frmInput As New MultilineInputForm
        frmInput.txtInput.Text = Me.txtSurveyDescription.Text
        frmInput.ShowDialog()
        If frmInput.DialogResult = Windows.Forms.DialogResult.OK Then
            Me.txtSurveyDescription.Text = frmInput.txtInput.Text
        End If
        StoreForUndo(data, prefs)
    End Sub

    Private Sub mnuEditVarietyMetadata_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEditVarietyMetadata.Click
        Dim frmInput As New MultilineInputForm
        frmInput.txtInput.Text = Me.txtVarietyDescription.Text
        frmInput.ShowDialog()
        If frmInput.DialogResult = Windows.Forms.DialogResult.OK Then
            Me.txtVarietyDescription.Text = frmInput.txtInput.Text
        End If
        StoreForUndo(data, prefs)
    End Sub

    'cmnuDictionary.Opening and mnuDictionary.DropDownOpening are differently typed events and cause runtime errors if handled by the same Sub
    Private Sub cmnuDictionary_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cmnuDictionary.Opening
        For Each cell As DataGridViewTextBoxCell In Me.grdGlossDictionary.SelectedCells
            If grdGlossDictionary.SelectedCells(0).RowIndex <> cell.RowIndex Then
                cmnuSortSelectionAlphabetically.Enabled = True
                Return
            End If
        Next
        cmnuSortSelectionAlphabetically.Enabled = False
    End Sub
    Private Sub mnuDictionary_DropDownOpening(sender As Object, e As System.EventArgs) Handles mnuDictionary.DropDownOpening
        For Each cell As DataGridViewTextBoxCell In Me.grdGlossDictionary.SelectedCells
            If grdGlossDictionary.SelectedCells(0).RowIndex <> cell.RowIndex Then
                mnuSortSelectionAlphabetically.Enabled = True
                Return
            End If
        Next
        mnuSortSelectionAlphabetically.Enabled = False
    End Sub

    Private Sub cmnuVariety_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cmnuVariety.Opening
        For Each cell As DataGridViewTextBoxCell In Me.grdVariety.SelectedCells
            If cell.ColumnIndex = 0 Then
                cmnuCutVarietyCells.Enabled = False
                cmnuPasteVarietyCells.Enabled = False
                cmnuDeleteVarietyCells.Enabled = False
                Return
            End If
        Next
        cmnuCutVarietyCells.Enabled = True
        cmnuPasteVarietyCells.Enabled = True
        cmnuDeleteVarietyCells.Enabled = True
    End Sub
End Class