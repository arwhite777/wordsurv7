Module UndoSystem
    'The undo system works by making a copy of the entire program state after every operation.  This copying operation is very fast
    'and allows us to save and revert back to all aspects of the program state, even things like the row the user was on.
    'Because all of this state is stored in the data structure and all of this data is wrapped up in the WordSurvData object,
    'undoing is as simple as setting the form's 'data' variable to an old state and refreshing the grids.

    'Any time a new feature will change the program's state the programmer must call StoreForUndo() after the
    'operation is complete.  StoreForUndo copies adds a copy of the current state to a linked list of old states.
    'This undo state list can be arbitrarily long, as long as the user doesn't run out of memory.

    'When the undo system is initialized, a copy of the original state is inserted into the empty state list and the original state
    'itself is added to the right side of the list.  Left is earlier states, right is later states.  The list looks like this:

    'curr--\
    '      V
    's     s

    '(For this example, s is a state without data.  Numbers after the s indicate data in the state.)
    'If the user changes the state of the program and adds 1 to it, before StoreForUndo() is called, the state looks like this:

    'curr--\
    '      V
    's     s1

    'After StoreForUndo() is called s1 is copied and the list looks like this:

    'curr--------\
    '            V
    's     s1    s1

    'Notice how there are again two copies of the most recent state.  This allows us to StoreForUndo after the operation instead of before.
    'The user then adds a few more values to the program state and we get this:

    'curr-------------------\
    '                       V
    's     s1    s2    s3   s3

    'Now the user wants to undo.  To maintain the requirement of an additional copy of the current state before the current state, we do
    'the following for an undo operation:

    'Delete the copy of the current state that was before the current state:
    'curr---------------\
    '                   V
    's     s1    s2     s3

    'set the current state to the state before the current state
    'curr--------\
    '            V
    's     s1    s2     s3

    'create a copy of the current state and insert it after the current state
    'curr--------\
    '            V
    's     s1    s2     s2    s3

    'set the current state to be the state after the current state
    'curr---------------\
    '                   V
    's     s1    s2     s2    s3

    'And now the undo tree is stable again.  We return s2 back to the form to use as the current state.

    'If the user makes any new changes to the state after going back in the undo tree, all states after the current one are removed
    'in StoreForUndo().

    'If the user does not change the program state, he can redo the undo which does the following:

    'Remove the copy of the current state that is before the current state:
    'curr---------\
    '             V
    's     s1     s2    s3

    'Set the current state to the state after the current state:
    'curr---------------\
    '                   V
    's     s1    s2     s3

    'Add a copy of the current state after the current state:
    'curr---------------\
    '                   V
    's     s1    s2     s3    s3

    'Set the current state to the state after the current state:
    'curr---------------------\
    '                         V
    's     s1    s2     s3    s3


    'Private ramWatcher As New System.Diagnostics.PerformanceCounter("Memory", "Available MBytes")

    Private UndoBuffer As New Generic.LinkedList(Of WordSurvData)
    Private CurrentState As Generic.LinkedListNode(Of WordSurvData)
    Public Sub InitUndo(ByVal dataObject As WordSurvData, ByVal prefs As Preferences)
        dataObject.CurrentTab = prefs.CurrentTab
        UndoBuffer.AddLast(dataObject.Copy())
        UndoBuffer.AddLast(dataObject)
        CurrentState = UndoBuffer.Last
        CurrentState.Value.CurrentTab = prefs.CurrentTab
    End Sub
    Public Sub StoreForUndo(ByVal dataObject As WordSurvData, ByVal prefs As Preferences)
        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If Now.CompareTo(BackupTimeStamp.AddMinutes(15.0)) > 0 Then
            dataObject.MakeCrashRecoveryBackup()
            BackupTimeStamp = Now
        End If

        Dim dev As New Devices.ComputerInfo()
        If dev.AvailablePhysicalMemory < 50000000 Then
            MsgBox("Your system is very low on RAM.  You may want to turn off Undo.", MsgBoxStyle.Critical)
        End If

        HasNotSaved = True

        'Setting then number of undo states to 0 is equivalent to turning it off
        If prefs.MaxUndos > 0 Then

            'If the current state is in the middle of the list of states and the user performs an operation,
            'chop off all states to the right, add this state, and make it the most recent state.
            While (CurrentState.Next IsNot Nothing)
                UndoBuffer.RemoveLast()
            End While

            If UndoBuffer.Count >= prefs.MaxUndos Then UndoBuffer.RemoveFirst()

            Dim copiedData As WordSurvData = dataObject.Copy()
            copiedData.CurrentTab = prefs.CurrentTab
            dataObject.CurrentTab = prefs.CurrentTab
            UndoBuffer.AddBefore(CurrentState, copiedData)
        Else
            If UndoBuffer.Count > 0 Then UndoBuffer.Clear()
        End If
        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

    End Sub
    Public Function Undo(ByVal prefs As Preferences) As WordSurvData
        If prefs.MaxUndos = 0 OrElse CurrentState.Previous.Previous Is Nothing Then Return Nothing

        Dim destTab As Integer = CurrentState.Previous.Value.CurrentTab
        UndoBuffer.Remove(CurrentState.Previous)
        CurrentState = CurrentState.Previous
        UndoBuffer.AddAfter(CurrentState, CurrentState.Value.Copy())
        CurrentState = CurrentState.Next
        CurrentState.Value.CurrentTab = destTab
        prefs.CurrentTab = destTab

        Return CurrentState.Value
    End Function
    Public Function Redo(ByVal prefs As Preferences) As WordSurvData
        If prefs.maxundos = 0 OrElse CurrentState.Next Is Nothing Then Return Nothing

        Dim prevprevTab As Integer = CurrentState.Previous.Value.CurrentTab
        Dim copysTab As Integer = CurrentState.Value.CurrentTab
        Dim actualsTab As Integer = CurrentState.Next.Value.CurrentTab

        UndoBuffer.Remove(CurrentState.Previous)
        CurrentState = CurrentState.Next
        UndoBuffer.AddAfter(CurrentState, CurrentState.Value.Copy())
        CurrentState = CurrentState.Next

        CurrentState.Previous.Previous.Value.CurrentTab = prevprevTab
        CurrentState.Previous.Value.CurrentTab = copysTab
        CurrentState.Value.CurrentTab = actualsTab
        prefs.CurrentTab = actualsTab

        Return CurrentState.Value
    End Function
    Public Sub ClearUndo(ByVal data As WordSurvData, ByVal prefs As Preferences)
        UndoBuffer = New Generic.LinkedList(Of WordSurvData)
        InitUndo(data, prefs)
    End Sub

    Public Sub DrawUndoBuffer(Optional ByVal msg As String = "")
        Dim pic As String = msg & vbCrLf & "                                   " & vbCrLf
        Dim cnt As Integer = 0

        For Each state As WordSurvData In UndoBuffer
            pic &= state.CurrentTab.ToString & ", " & cnt.ToString
            If state Is CurrentState.Value Then pic &= "<"
            pic &= vbCrLf
            cnt += 1
        Next
        MsgBox(pic)
    End Sub
End Module
