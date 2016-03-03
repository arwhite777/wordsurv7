Module Module1
    Public Sub truncGroupings(ByRef groupings As String, ByVal synonyms As String)
        Dim groupingsIndex As Integer = 0
        Dim groupingsCommasFound As Integer = 0
        Do
            groupingsIndex = groupings.IndexOf(","c, groupingsIndex)
            If groupingsIndex = -1 Then Exit Do
            groupingsIndex += 1
            groupingsCommasFound += 1
        Loop
        Dim synIndex As Integer = 0
        Dim synCommasFound As Integer = 0
        Do
            synIndex = synonyms.IndexOf(","c, synIndex)
            If synIndex = -1 Then Exit Do
            synIndex += 1
            synCommasFound += 1
        Loop
        If groupingsCommasFound > synCommasFound Then
            Dim index As Integer = 0
            Dim counter As Integer = 0
            Dim place As Integer = 0
            Do While counter < synCommasFound + 1
                index = groupings.IndexOf(",", place)
                place = index + 1
                counter += 1
            Loop
            groupings = groupings.Substring(0, place - 1)
        End If
    End Sub
    Public Sub PadStringsToLongest(ByRef str1 As String, ByRef str2 As String)
        Dim maxLength As Integer
        If str1.Length > str2.Length Then
            maxLength = str1.Length
        Else
            maxLength = str2.Length
        End If
        str1 = str1.PadRight(maxLength)
        str2 = str2.PadRight(maxLength)
    End Sub
End Module
