Public Class SplitterTest
    Private Splitter1Distance As Integer = 72
    Private Splitter2Distance As Integer = 72


    Private Sub SplitterTest_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim usedChars As New List(Of String)
        Dim chars As String() = {"a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "`", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "~", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "[", "{", "]", "}", "/", "=", "\", "?", "+", "|", ";", ":", "'", ",", "<", ".", ">"}
        usedChars.AddRange(chars)


        MsgBox("begin char test hash")
        For i As Integer = 0 To 10000
            Dim charsList As New Dictionary(Of String, Integer)
            For Each char1 As String In usedChars
                For Each char2 As String In usedChars
                    If Not charsList.ContainsKey(char1) Then
                        charsList.Add(char1, 0)
                    Else
                        charsList(char1) += 1
                    End If
                    If Not charsList.ContainsKey(char2) Then
                        charsList.Add(char2, 0)
                    Else
                        charsList(char2) += 1
                    End If
                Next
            Next
        Next
        MsgBox("end char test hash")

        MsgBox("begin char test bit")
        For i As Integer = 0 To 10000
            Dim charsList As New Dictionary(Of Integer, Integer)
            For Each char1 As String In usedChars
                For Each char2 As String In usedChars
                    Dim key1 As Integer = AscW(char1)
                    Dim key2 As Integer = AscW(char2)
                    If Not charsList.ContainsKey(key1) Then
                        charsList.Add(key1, 0)
                    Else
                        charsList(key1) += 1
                    End If
                    If Not charsList.ContainsKey(key2) Then
                        charsList.Add(key2, 0)
                    Else
                        charsList(key2) += 1
                    End If
                Next
            Next
        Next
        MsgBox("end char test bit")

        MsgBox("begin 2 level hash")
        For i As Integer = 0 To 10000
            Dim table As New Dictionary(Of String, Dictionary(Of String, Integer))
            For Each char1 As String In usedChars
                For Each char2 As String In usedChars

                    If Not table.ContainsKey(char1) Then
                        table.Add(char1, New Dictionary(Of String, Integer))
                    End If
                    If Not table.ContainsKey(char2) Then
                        table.Add(char2, New Dictionary(Of String, Integer))
                    End If

                    If Not table(char1).ContainsKey(char2) Then
                        If char1 = char2 Then
                            table(char1).Add(char2, 0)
                        Else
                            table(char1).Add(char2, 1)
                        End If
                    Else
                        table(char1)(char2) += 1
                    End If
                    If Not table(char2).ContainsKey(char1) Then
                        If char1 = char2 Then
                            table(char2).Add(char1, 0)
                        Else
                            table(char2).Add(char1, 1)
                        End If
                    Else
                        table(char2)(char1) += 1
                    End If

                Next
            Next
        Next

        MsgBox("end 2 level hash")

        Dim intChars As New List(Of Integer)
        For Each usedchar As String In usedChars
            intChars.Add(AscW(usedchar))
        Next
        MsgBox("begin bitwise")
        For i As Integer = 0 To 10000
            Dim table As New Dictionary(Of Integer, Integer)
            For Each char1 As String In usedChars
                For Each char2 As String In usedChars
                    Dim combination As Int32 = (AscW(char1) << 16) Or AscW(char2)
                    If Not table.ContainsKey(combination) Then
                        table.Add(combination, 1)
                    Else
                        table(combination) += 1
                    End If
                Next
            Next
        Next
        MsgBox("end bitwise")

        'Me.Width = 400
        'Me.Height = 400

        'Me.TabControl1.SelectedIndex = 1
        'Application.DoEvents()
        'Me.TabControl1.SelectedIndex = 0
        'Application.DoEvents()
        'Me.SplitContainer1.SplitterDistance = Splitter1Distance
        'Me.SplitContainer2.SplitterDistance = Splitter2Distance

        'AddHandler SplitContainer1.SplitterMoved, AddressOf SplitContainer1_SplitterMoved
        'AddHandler SplitContainer2.SplitterMoved, AddressOf SplitContainer2_SplitterMoved
    End Sub

    'Private Sub SplitContainer1_SplitterMoved(ByVal sender As System.Object, ByVal e As System.Windows.Forms.SplitterEventArgs)
    '    Splitter1Distance = Me.SplitContainer1.SplitterDistance
    'End Sub

    'Private Sub SplitContainer2_SplitterMoved(ByVal sender As System.Object, ByVal e As System.Windows.Forms.SplitterEventArgs)
    '    Splitter1Distance = Me.SplitContainer2.SplitterDistance
    'End Sub

    'Private Sub SplitContainer1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles SplitContainer1.Resize
    '    Dim x As Integer = 0
    'End Sub
End Class