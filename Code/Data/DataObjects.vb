Imports System.IO
Imports System.Reflection
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions

Public Module DataObjects
    Public Class Gloss
        Public Name As String = ""
        Public Name2 As String = ""
        Public PartOfSpeech As String = ""
        Public FieldTip As String = ""
        Public Comments As String = ""
        Public Shared ColumnCount As Integer = 5

        Public Function GetColumnCount() As Integer
            Return 5
        End Function
        Public Sub New(Optional ByVal name As String = "", Optional ByVal name2 As String = "", Optional ByVal partOfSpeech As String = "", Optional ByVal fieldTip As String = "", Optional ByVal comments As String = "")
            name.Trim(" "c)
            Me.Name = name
            Me.Name2 = name2
            Me.PartOfSpeech = partOfSpeech
            Me.FieldTip = fieldTip
            Me.Comments = comments
        End Sub
        Public Sub SetByIndex(ByVal index As Integer, ByVal val As String)
            Select Case index
                Case 0
                    Me.Name = val
                    HasNotSaved = True 'AJW*** needed now
                Case 1
                    Me.Name2 = val
                    HasNotSaved = True 'AJW*** needed now
                Case 2
                    Me.PartOfSpeech = val
                    HasNotSaved = True 'AJW*** needed now
                Case 3
                    Me.FieldTip = val
                    HasNotSaved = True 'AJW*** needed now
                Case 4
                    Me.Comments = val
                    HasNotSaved = True 'AJW*** needed now
                Case Else
                    MsgBox("Invalid index for gloss information (ill-formed gloss?), gloss " & val & " index ")
                    Throw New IndexOutOfRangeException
            End Select
        End Sub
        Public Function GetByIndex(ByVal index As Integer) As String
            Dim val As String = Nothing
            If index = 0 Then val = Me.Name
            If index = 1 Then val = Me.Name2
            If index = 2 Then val = Me.PartOfSpeech
            If index = 3 Then val = Me.FieldTip
            If index = 4 Then val = Me.Comments
            If val Is Nothing Then
                Return ""
            Else
                Return val
            End If
        End Function
        Public Function Copy() As Gloss
            Return New Gloss(Me.Name, Me.Name2, Me.PartOfSpeech, Me.FieldTip, Me.Comments)
        End Function
    End Class
    Public Class Sort
        Public Name As String
        Public Glosses As New List(Of Gloss)

        Public Sub New(ByVal name As String)
            name.Trim(" "c)
            Me.Name = name
        End Sub
        Public Function CopyGlosses() As List(Of Gloss)
            Dim copiedGlosses As New List(Of Gloss)
            For Each gl As Gloss In Me.Glosses
                copiedGlosses.Add(gl.Copy())
            Next
            Return copiedGlosses
        End Function
        Public Function Copy(ByVal newName As String) As Sort
            Dim newSort As New Sort(newName)
            newSort.Glosses = Me.CopyGlosses()
            Return newSort
        End Function
    End Class
    Public Class Dictionary
        Public Name As String
        Public Sorts As New List(Of Sort)
        Public CurrentSort As Sort = Nothing
        Public CurrentGloss As Gloss = Nothing
        Public CurrentGlossColumnIndex As Integer = 0

        Public Sub New(ByVal name As String)
            name.Trim(" "c)
            Me.Name = name
        End Sub
        Public Function Copy(ByVal newName As String) As Dictionary
            Dim newDictionary As New Dictionary(newName)

            For Each srt As Sort In Me.Sorts
                newDictionary.Sorts.Add(srt.Copy(srt.Name))
            Next

            newDictionary.CurrentSort = newDictionary.Sorts(0)
            If Me.CurrentSort.Glosses.Count > 0 Then newDictionary.CurrentGloss = Me.CurrentSort.Glosses(0)
            Return newDictionary
        End Function
    End Class
    Public Class VarietyEntry
        Public Transcription As String = ""
        Public PluralFrame As String = ""
        Public Notes As String = ""
        Public Sub New(Optional ByVal transcription As String = "", Optional ByVal pluralFrame As String = "", Optional ByVal notes As String = "")
            Me.Transcription = transcription
            Me.PluralFrame = pluralFrame
            Me.Notes = notes
        End Sub
        Public Function Copy() As VarietyEntry
            Return New VarietyEntry(Me.Transcription, Me.PluralFrame, Me.Notes)
        End Function
    End Class
    Public Class Variety
        Public Name As String
        Public VarietyEntries As New Dictionary(Of Gloss, VarietyEntry)
        Public CurrentVarietyEntry As VarietyEntry = Nothing
        Public AssociatedDictionary As Dictionary

        Public Description As String = "Long (variety) name:" & vbCrLf & _
                                        "Variety (name):" & vbCrLf & _
                                        "Abbreviation:" & vbCrLf & _
                                        "ISO 639-3 code:" & vbCrLf & _
                                        "Alternate language names:" & vbCrLf & _
                                        "Genetic classification:" & vbCrLf & _
                                        "Date transcription started:" & vbCrLf & _
                                        "Date transcription finished:" & vbCrLf & _
                                        "Reliability:" & vbCrLf & _
                                        "Country:" & vbCrLf & _
                                        "Province/State:" & vbCrLf & _
                                        "District:" & vbCrLf & _
                                        "Sub-District:" & vbCrLf & _
                                        "Village:" & vbCrLf & _
                                        "Coordinates North Limit:" & vbCrLf & _
                                        "Coordinates South Limit:" & vbCrLf & _
                                        "Coordinates East Limit:" & vbCrLf & _
                                        "Coordinates West Limit:" & vbCrLf & _
                                        "Coordinates(Latitude):" & vbCrLf & _
                                        "Coordinates(Longitude):" & vbCrLf & _
                                        "Speaker:" & vbCrLf & _
                                        "Speaker(Sex):" & vbCrLf & _
                                        "Speaker(Age):" & vbCrLf & _
                                        "Transcriber:" & vbCrLf & _
                                        "Interviewer:" & vbCrLf & _
                                        "Recorder:" & vbCrLf & _
                                        "Unpublished(source):" & vbCrLf & _
                                        "Published(source):" & vbCrLf & _
                                        "Remarks:"






            Public Shared ColumnCount As Integer = 4

        Public Sub New(ByRef dict As Dictionary, ByVal name As String, ByVal fillVariety As Boolean)
            name.Trim(" "c)
            Me.Name = name
            Me.AssociatedDictionary = dict
            If fillVariety Then
                For Each gl As Gloss In Me.AssociatedDictionary.CurrentSort.Glosses
                    Me.VarietyEntries.Add(gl, New VarietyEntry(""))
                Next
                Try
                    Me.CurrentVarietyEntry = Me.VarietyEntries(Me.AssociatedDictionary.CurrentGloss)
                Catch ex As Exception
                    Me.CurrentVarietyEntry = Nothing
                End Try
            End If
        End Sub
        'Public Function Copy(ByRef associatedDictionary As Dictionary, ByVal newName As String) As Survey
        'Dim newVariety As New Variety(associatedDictionary, newName)
        'newVariety.Description = Me.Description

        'For Each var As Variety In Me.Varieties
        '    newSurvey.Varieties.Add(var.Copy(var.Name)))
        'Next

        'newSurvey.CurrentVariety = newSurvey.Varieties(0)
        'Return newVariety
        'End Function
    End Class
    Public Class Survey
        Public Name As String
        Public Varieties As New List(Of Variety)
        Public CurrentVariety As Variety = Nothing
        Public AssociatedDictionary As Dictionary
        Public CurrentVarietyEntryColumnIndex As Integer = 0

        Public Description As String = "Full Title:" & vbCrLf & _
                                        "Description:" & vbCrLf & _
                                        "Remarks:" & vbCrLf & _
                                        "Complier:" & vbCrLf & _
                                        "Consultant:" & vbCrLf & _
                                        "Other Contributor:" & vbCrLf & _
                                        "Publisher:" & vbCrLf & _
                                        "Keywords for Searching:" & vbCrLf & _
                                        "Published Source(s):" & vbCrLf & _
                                        "Geographic Area Covered:" & vbCrLf & _
                                        "Stable Copy Located At:" & vbCrLf & _
                                        "Rights Management:" & vbCrLf & _
                                        "Rights Holder:" & vbCrLf & _
                                        "Year Copyright Asserted :" & vbCrLf & _
                                        "Date Created in WordSurv:" & Date.Now.ToString & vbCrLf & _
                                        "Date Modified in WordSurv:" & Date.Now.ToString & vbCrLf


        Public Sub New(ByRef dict As Dictionary, ByVal name As String)
            name.Trim(" "c)
            Me.AssociatedDictionary = dict
            Me.Name = name
        End Sub

    End Class
    Public Class ComparisonAnalysis
        Public AssociatedComparison As Comparison
        Public TallyMatrix As Dictionary(Of Variety, Dictionary(Of Variety, Integer))
        Public TotalMatrix As Dictionary(Of Variety, Dictionary(Of Variety, Integer))
        Public PercentMatrix As Dictionary(Of Variety, Dictionary(Of Variety, Integer))

        Public CurrentVariety As Variety = Nothing
        Public CurrentVarietyColumnIndex As Integer = 0

        Public Sub New(ByRef assocComp As Comparison, ByRef varieties As List(Of Variety))
            Me.AssociatedComparison = assocComp
            Me.CurrentVariety = assocComp.CurrentVariety
        End Sub

        Public Sub Calculate()
            Me.TallyMatrix = New Dictionary(Of Variety, Dictionary(Of Variety, Integer))
            Me.TotalMatrix = New Dictionary(Of Variety, Dictionary(Of Variety, Integer))
            Me.PercentMatrix = New Dictionary(Of Variety, Dictionary(Of Variety, Integer))

            Dim glosses As List(Of Gloss) = Me.AssociatedComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses
            Dim varieties As List(Of Variety) = Me.AssociatedComparison.AssociatedSurvey.Varieties
            Dim comparisonEntries As Dictionary(Of VarietyEntry, ComparisonEntry) = Me.AssociatedComparison.ComparisonEntries

            'Set up the hash mess.
            For Each varI As Variety In varieties
                Me.TallyMatrix.Add(varI, New Dictionary(Of Variety, Integer))
                Me.TotalMatrix.Add(varI, New Dictionary(Of Variety, Integer))
                Me.PercentMatrix.Add(varI, New Dictionary(Of Variety, Integer))
                For Each varJ As Variety In varieties
                    Me.TallyMatrix(varI).Add(varJ, 0)
                    Me.TotalMatrix(varI).Add(varJ, 0)
                    Me.PercentMatrix(varI).Add(varJ, 0)
                Next
            Next

            For Each gl As Gloss In glosses
                For Each varI As Variety In varieties
                    For Each varJ As Variety In varieties
                        Dim compEntry1 As ComparisonEntry = comparisonEntries(varI.VarietyEntries(gl))
                        Dim compEntry2 As ComparisonEntry = comparisonEntries(varJ.VarietyEntries(gl))

                        If compEntry1.Exclude = "" AndAlso compEntry2.Exclude = "" AndAlso _
                           compEntry1.AlignedRendering <> "" AndAlso compEntry2.AlignedRendering <> "" Then
                            Me.TotalMatrix(varI)(varJ) += 1
                            If compEntry1.Grouping <> "" And compEntry2.Grouping <> "" AndAlso GroupsMatch(compEntry1.Grouping, compEntry2.Grouping) Then
                                Me.TallyMatrix(varI)(varJ) += 1
                            End If
                        End If
                    Next
                Next
            Next

            For Each varI As Variety In varieties
                For Each varJ As Variety In varieties
                    If Me.TotalMatrix(varI)(varJ) <> 0 Then
                        Me.PercentMatrix(varI)(varJ) = CInt(100.0 * Me.TallyMatrix(varI)(varJ) / Me.TotalMatrix(varI)(varJ))
                    Else
                        Me.PercentMatrix(varI)(varJ) = 0
                    End If
                Next
            Next

        End Sub
    End Class
    Public Class ComparisonEntry
        Public AlignedRendering As String = ""
        Public Grouping As String = ""
        Public Notes As String = ""
        Public Exclude As String = ""

        Public Sub New(ByVal alignedRendering As String, Optional ByVal grouping As String = "", Optional ByVal notes As String = "", Optional ByVal exclude As String = "")
            Me.AlignedRendering = alignedRendering
            Me.Grouping = grouping
            Me.Notes = notes
            Me.Exclude = exclude
        End Sub

        Public Function Copy() As ComparisonEntry
            Return New ComparisonEntry(Me.AlignedRendering, Me.Grouping, Me.Notes, Me.Exclude)
        End Function
    End Class
    Public Class Comparison
        Public Name As String
        Public CurrentVarietySort As New List(Of Variety)
        Public DefaultVarietySort As New List(Of Variety)
        Public ComparisonEntries As New Dictionary(Of VarietyEntry, ComparisonEntry)
        Public CurrentVariety As Variety = Nothing
        Public AssociatedSurvey As Survey
        Public AssociatedAnalysis As ComparisonAnalysis
        Public AssociatedDegreesOfDifference As DegreesOfDifferenceGrid
        Public Description As String = ""
        Public StartDate As Date = Date.Now
        Public EndDate As Date = Date.Now
        Public CurrentVarietyColumnIndex As Integer = 0
        Public SelectedPhonePairCoordinates As New List(Of CellAddress)
        Public CurrentCOMPASSStrengthsSummaryCellAddress As CellAddress

        Public Shared ColumnCount As Integer = 7

        Public COMPASSCalculations As COMPASSCalculation = Nothing

        Public Sub New(ByRef surv As Survey, ByVal name As String, ByVal fillEntries As Boolean)
            name.Trim(" "c)
            Me.AssociatedSurvey = surv
            Me.Name = name

            If fillEntries Then
                For Each gl As Gloss In Me.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses
                    For Each var As Variety In Me.AssociatedSurvey.Varieties
                        Me.ComparisonEntries.Add(var.VarietyEntries(gl), New ComparisonEntry(var.VarietyEntries(gl).Transcription)) 'Copy the transcription into the aligned rendering field.
                    Next
                Next
                For Each var As Variety In Me.AssociatedSurvey.Varieties
                    Me.CurrentVarietySort.Add(var)
                    Me.DefaultVarietySort.Add(var)
                Next
                Try
                    Me.CurrentVariety = Me.AssociatedSurvey.Varieties(0)
                Catch ex As Exception
                End Try
            End If
            Me.Description &= "Survey: " & Me.AssociatedSurvey.Name & vbCrLf & "Comparison Created: " & Date.Now.ToString & vbCrLf & "Description:"
            Me.AssociatedAnalysis = New ComparisonAnalysis(Me, Me.AssociatedSurvey.Varieties)

            Me.AssociatedDegreesOfDifference = New DegreesOfDifferenceGrid(Me)
        End Sub
    End Class
    Public Class DegreesOfDifferenceGrid
        Public AssociatedComparison As Comparison

        Public UsedCharsList As New List(Of String)
        Public CharSort As New List(Of Integer)

        Public GlossesUsing As New Dictionary(Of Integer, List(Of Gloss))

        Public ExcludedChars As String = ""

        Public CurrentRowIndex As Integer = 0

        Public DDs As Dictionary(Of Integer, Integer) = Nothing

        Public DDCharCorrespondences As New Dictionary(Of Integer, Integer)
        Public DDMatrixRatio As New Dictionary(Of Variety, Dictionary(Of Variety, Integer))
        Public DDMatrixDegrees As New Dictionary(Of Variety, Dictionary(Of Variety, Integer))
        Public DDMatrixCorrespondences As New Dictionary(Of Variety, Dictionary(Of Variety, Integer))

        Public Sub New(ByRef associatedComparison As Comparison)
            Me.AssociatedComparison = associatedComparison
        End Sub
        Public Sub truncCommas(ByRef s1 As String, ByRef s2 As String)
            Dim index1 As Integer = 0
            Dim found1 As Integer = 0
            Dim index2 As Integer = 0
            Dim found2 As Integer = 0
            Do
                index1 = s1.IndexOf(","c, index1)
                If index1 = -1 Then Exit Do
                index1 += 1
                found1 += 1
            Loop
            Do
                index2 = s2.IndexOf(","c, index2)
                If index2 = -1 Then Exit Do
                index2 += 1
                found2 += 1
            Loop

            If found1 <> found2 Then
                If found1 > found2 Then
                    If found2 > 0 Then
                        s1 = s1.Substring(0, found2 - 1)
                    Else
                        s1 = s1.Substring(0, 1)
                    End If
                Else
                    If found1 > found2 Then
                        s2 = s2.Substring(0, found1 - 1)
                    Else
                        s2 = s2.Substring(0, 1)
                    End If
                End If
                'MsgBox("Truncating Groupings")
            End If

        End Sub

        'If we are assuming the calculations are now much faster, can we combine the calculating of characters and calculating the results into one step? no, different tabs
        Public Sub CalculateUsedChars()
            Dim variety As String = ""
            Dim gloss As String = ""
            'Calc is based on conservation of old data (minus pirs lost) and addition of new pairs
            Dim UsedCharsHash As New Dictionary(Of Integer, Integer) 'simple hash to hold the characters already in the hash table
            Dim GlossesUsingHash As New Dictionary(Of Integer, Dictionary(Of Gloss, Integer))
            Dim newGlossesUsing As New Dictionary(Of String, Dictionary(Of String, List(Of Gloss)))
            Dim newDDs As New Dictionary(Of Integer, Integer) 'to hold new char pairs added due to recalc with changes to the groupings (hence the correspondence pairs included may change)
            Dim resultMSG As MsgBoxResult
            Dim skipMSGFlag As Boolean = False

            'For each gloss in the comparison,
            For Each gl As Gloss In Me.AssociatedComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses
                'Match each variety with every other variety
                For Each var1 As Variety In Me.AssociatedComparison.CurrentVarietySort
                    For Each var2 As Variety In Me.AssociatedComparison.CurrentVarietySort
                        If var1 Is var2 Then Continue For

                        Dim compEntry1 As ComparisonEntry = Me.AssociatedComparison.ComparisonEntries(var1.VarietyEntries(gl))
                        Dim compEntry2 As ComparisonEntry = Me.AssociatedComparison.ComparisonEntries(var2.VarietyEntries(gl))

                        If compEntry1.Exclude <> "" Or compEntry2.Exclude <> "" Then Continue For 'DO NOT INCLUDE THESE CHARACTERS IN DoD!

                        'Truncate the longer set of groupings commas, e.g.  a b,c,d,e,f g,i vs a,b,c,d e would truncate the former to just 5 groups
                        Dim g1 As String = compEntry1.Grouping 'a,b
                        Dim g2 As String = compEntry2.Grouping 'a
                        Dim s1 As String = compEntry1.AlignedRendering
                        Dim s2 As String = compEntry2.AlignedRendering

                        Dim TESTsynonymGroupings1 As String() = Split(g1, ",") 'arm, irm
                        Dim TESTsynonymGroupings2 As String() = Split(g2, ",") 'urm
                        Dim TESTsynonyms1 As String() = Split(s1, ",")
                        Dim TESTsynonyms2 As String() = Split(s2, ",")

                        If TESTsynonymGroupings1.Length <> TESTsynonyms1.Length Then
                            If var1.Name <> variety Or gl.Name <> gloss Then '(To prevent repeat messages for each var x var comparison)
                                If compEntry1.Exclude = "" And compEntry2.Exclude = "" Then
                                    If Not skipMSGFlag Then
                                        'MsgBox("Gloss '" & gl.Name & "' for " & var1.Name & " has a mismatched number of comma separated items in the aligned field versus the groupings field, which may result in incorrectly populated Degrees of Difference, Phonostatistical Analysis, and COMPASS grids!")
                                        resultMSG = MsgBox("Gloss '" & gl.Name & "' for " & var1.Name & " has a mismatched number of comma separated items in the aligned field versus the groupings field, which may result in incorrectly populated Degrees of Difference, Phonostatistical Analysis, and COMPASS grids!" & vbCrLf & vbCrLf & "Would you like to skip the warning message for the rest of the glosses?", MsgBoxStyle.YesNo, "Mismatched number of gloss synonyms and groupings")
                                        If resultMSG = MsgBoxResult.Yes Then
                                            skipMSGFlag = True
                                        End If
                                    End If
                                End If
                            End If
                            variety = var1.Name
                            gloss = gl.Name
                        End If

                        truncGroupings(g1, s1)
                        truncGroupings(g2, s2)
                        Dim synonymGroupings1 As String() = Split(g1, ",") 'arm, irm
                        Dim synonymGroupings2 As String() = Split(g2, ",") 'urm
                        Dim synonyms1 As String() = Split(s1, ",")
                        Dim synonyms2 As String() = Split(s2, ",")



                        'For each synonym grouping that matches, compare the aligned renderings
                        For synIndex1 As Integer = 0 To synonymGroupings1.Length - 1 'synonyms1.Length - 1 'AJW***
                            For synIndex2 As Integer = 0 To synonymGroupings2.Length - 1 'synonyms2.Length - 1 'AJW***
                                If GroupsMatch(synonymGroupings1(synIndex1), synonymGroupings2(synIndex2)) Then

                                    'Take their aligned renderings
                                    Dim word1 As String = synonyms1(synIndex1)
                                    Dim word2 As String = synonyms2(synIndex2)

                                    'Remove all excluded chars
                                    For Each ch As String In Me.ExcludedChars
                                        If word1.Contains(ch) Then word1 = word1.Replace(ch, "")
                                        If word2.Contains(ch) Then word2 = word2.Replace(ch, "")
                                    Next

                                    'Padding the shorter with spaces
                                    'PadStringsToLongest(word1, word2)
                                    Dim maxLength As Integer
                                    If word1.Length > word2.Length Then
                                        maxLength = word1.Length
                                    Else
                                        maxLength = word2.Length
                                    End If
                                    word1 = word1.PadRight(maxLength)
                                    word2 = word2.PadRight(maxLength)


                                    'Now go through both words and add any new letters into DD matrix and the hash of characters used in this comparison
                                    For i As Integer = 0 To maxLength - 1
                                        Dim char1 As Integer = AscW(word1(i))
                                        Dim char2 As Integer = AscW(word2(i))
                                        Dim char1AndChar2 As Integer = (char1 << 16) Or char2

                                        If Not newDDs.ContainsKey(char1AndChar2) Then
                                            If char1 = char2 Then
                                                newDDs.Add(char1AndChar2, 0)
                                            Else
                                                newDDs.Add(char1AndChar2, 1)
                                            End If
                                        End If

                                        If Not UsedCharsHash.ContainsKey(char1) Then
                                            UsedCharsHash.Add(char1, 0)
                                        End If
                                        If Not UsedCharsHash.ContainsKey(char2) Then
                                            UsedCharsHash.Add(char2, 0)
                                        End If

                                        If Not GlossesUsingHash.ContainsKey(char1AndChar2) Then
                                            Dim gls As New Dictionary(Of Gloss, Integer)
                                            gls.Add(gl, 0)
                                            GlossesUsingHash.Add(char1AndChar2, gls)
                                        Else
                                            Dim gls As Dictionary(Of Gloss, Integer) = GlossesUsingHash(char1AndChar2)
                                            If Not gls.ContainsKey(gl) Then
                                                gls.Add(gl, 0)
                                            End If
                                        End If
                                    Next
                                End If
                            Next
                        Next
                    Next
                Next
            Next

            Me.GlossesUsing.Clear()
            For Each kvp As KeyValuePair(Of Integer, Dictionary(Of Gloss, Integer)) In GlossesUsingHash
                Me.GlossesUsing.Add(kvp.Key, New List(Of Gloss))
                For Each gl As Gloss In kvp.Value.Keys
                    Me.GlossesUsing(kvp.Key).Add(gl)
                Next
            Next

            'Fill in the holes in the DD matrix with empty values (-1).  Any place where there is a -1, the grid
            'cell will be grayed out.
            For Each char1 As Integer In UsedCharsHash.Keys
                For Each char2 As Integer In UsedCharsHash.Keys
                    Dim char1AndChar2 As Integer = (char1 << 16) Or char2

                    If Not newDDs.ContainsKey(char1AndChar2) Then
                        newDDs.Add(char1AndChar2, -1)

                    End If
                    If Not newDDs.ContainsKey(char1AndChar2) Then
                        newDDs.Add(char1AndChar2, -1)

                    End If
                Next
            Next


            Dim newChars As New Dictionary(Of Integer, Integer)
            Dim isCharBroughtOver As New Dictionary(Of Integer, Boolean)
            For Each ch As Integer In UsedCharsHash.Keys
                isCharBroughtOver.Add(ch, False)
            Next


            'If they already had a grid in place and did something that requires the grid to be recalculated,
            'like changing a transcription or grouping, copy the old DD values that apply
            If Me.DDs IsNot Nothing Then
                For Each char1 As Integer In UsedCharsHash.Keys
                    For Each char2 As Integer In UsedCharsHash.Keys
                        Dim char1AndChar2 As Integer = (char1 << 16) Or char2
                        Try
                            If newDDs(char1AndChar2) <> -1 AndAlso Me.DDs(char1AndChar2) <> -1 Then 'Only copy a value over if it was present in both the old and new matricies
                                newDDs(char1AndChar2) = Me.DDs(char1AndChar2)
                                isCharBroughtOver(char1) = True
                                isCharBroughtOver(char2) = True
                            End If
                        Catch ex As Exception
                        End Try
                        If Not newChars.ContainsKey(char1) Then newChars.Add(char1, Nothing)
                        If Not newChars.ContainsKey(char2) Then newChars.Add(char2, Nothing)
                    Next
                Next
            End If
            Me.DDs = newDDs

            If Me.UsedCharsList.Count > 0 Then

                For Each ch As Integer In UsedCharsHash.Keys
                    If isCharBroughtOver(ch) Then newChars.Remove(ch)
                Next

                Dim tempList As New List(Of Integer)
                tempList.AddRange(newChars.Keys)
                For Each ch As String In Me.UsedCharsList
                    If isCharBroughtOver.ContainsKey(AscW(ch)) AndAlso isCharBroughtOver(AscW(ch)) Then
                        tempList.Add(AscW(ch))
                    End If
                Next
                Me.UsedCharsList.Clear()
                For Each ch As Integer In tempList
                    Me.UsedCharsList.Add(ChrW(ch).ToString)
                Next
            Else
                Dim keysArray(UsedCharsHash.Count - 1) As Integer
                UsedCharsHash.Keys.CopyTo(keysArray, 0)
                Array.Sort(keysArray)
                For Each ch As Integer In keysArray
                    Me.UsedCharsList.Add(ChrW(ch).ToString)
                Next
            End If

        End Sub
        Public Sub DoAnalysis()
            Dim UsedCharsHash As New Dictionary(Of String, Integer)
            'Initialize the hash tables.
            Me.DDCharCorrespondences.Clear()
            Me.DDMatrixDegrees.Clear()
            Me.DDMatrixCorrespondences.Clear()
            Me.DDMatrixRatio.Clear()
            For Each usedChar1 As String In Me.UsedCharsList
                For Each usedChar2 As String In Me.UsedCharsList
                    Dim char1AndChar2 As Integer = (AscW(usedChar1) << 16) Or AscW(usedChar2)
                    If Me.DDs(char1AndChar2) <> -1 Then
                        Me.DDCharCorrespondences.Add(char1AndChar2, 0)
                    Else
                        Me.DDCharCorrespondences.Add(char1AndChar2, -1)
                    End If
                Next
            Next
            For Each var1 As Variety In Me.AssociatedComparison.CurrentVarietySort
                Me.DDMatrixDegrees.Add(var1, New Dictionary(Of Variety, Integer))
                Me.DDMatrixCorrespondences.Add(var1, New Dictionary(Of Variety, Integer))
                Me.DDMatrixRatio.Add(var1, New Dictionary(Of Variety, Integer))
                For Each var2 As Variety In Me.AssociatedComparison.CurrentVarietySort
                    Me.DDMatrixDegrees(var1).Add(var2, 0)
                    Me.DDMatrixCorrespondences(var1).Add(var2, 0)
                    Me.DDMatrixRatio(var1).Add(var2, 0)
                Next
            Next

            'For each gloss in the comparison,
            For Each gl As Gloss In Me.AssociatedComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses
                'Match each variety with every other variety
                For Each var1 As Variety In Me.AssociatedComparison.CurrentVarietySort
                    For Each var2 As Variety In Me.AssociatedComparison.CurrentVarietySort
                        If var1 Is var2 Then Continue For

                        'Find the comparison entry for this gloss,variety,variety combination
                        Dim compEntry1 As ComparisonEntry = Me.AssociatedComparison.ComparisonEntries(var1.VarietyEntries(gl))
                        Dim compEntry2 As ComparisonEntry = Me.AssociatedComparison.ComparisonEntries(var2.VarietyEntries(gl))

                        If compEntry1.Exclude <> "" Or compEntry2.Exclude <> "" Then Continue For 'DO NOT INCLUDE THESE CHARACTERS IN DoD!


                        'THIS WHOLE NEXT SECTION ADDED BY AJW***
                        'Truncate the longer set of groupings commas, e.g.  a b,c,d,e,f g,i vs a,b,c,d e would truncate the former to just 5 groups
                        Dim g1 As String = compEntry1.Grouping 'a,b
                        Dim g2 As String = compEntry2.Grouping 'a
                        Dim s1 As String = compEntry1.AlignedRendering
                        Dim s2 As String = compEntry2.AlignedRendering
                        Dim TESTsynonymGroupings1 As String() = Split(g1, ",") 'arm, irm
                        Dim TESTsynonymGroupings2 As String() = Split(g2, ",") 'urm
                        Dim TESTsynonyms1 As String() = Split(s1, ",")
                        Dim TESTsynonyms2 As String() = Split(s2, ",")
                        If TESTsynonymGroupings1.Length <> TESTsynonyms1.Length Then
                            'MsgBox("Gloss '" & gl.Name & "' for " & var1.Name & " has a mismatched number of comma separated items in the aligned field versus the groupings field, which may result in an incorrectly populated Degrees of Difference grid!")
                        End If
                        truncGroupings(g1, s1)
                        truncGroupings(g2, s2)



                        Dim synonymGroupings1 As String() = Split(compEntry1.Grouping, ",")
                        Dim synonymGroupings2 As String() = Split(compEntry2.Grouping, ",")
                        Dim synonyms1 As String() = Split(compEntry1.AlignedRendering, ",")
                        Dim synonyms2 As String() = Split(compEntry2.AlignedRendering, ",")

                        'If synonymGroupings1.Length = synonyms1.Length Then
                        'If synonymGroupings2.Length = synonyms2.Length Then



                        'aoeu,oeui  a,bc
                        'aoue,snth  a,c

                        'For each synonym grouping that matches, compare the aligned renderings
                        For synIndex1 As Integer = 0 To synonymGroupings1.Length - 1
                            For synIndex2 As Integer = 0 To synonymGroupings2.Length - 1


                                'Only if the groups are the same
                                If GroupsMatch(synonymGroupings1(synIndex1), synonymGroupings2(synIndex2)) Then

                                    'Take their aligned renderings
                                    Dim word1 As String = synonyms1(synIndex1)
                                    Dim word2 As String = synonyms2(synIndex2)

                                    'Remove all excluded chars
                                    For Each ch As String In Me.ExcludedChars
                                        If word1.Contains(ch) Then word1 = word1.Replace(ch, "")
                                        If word2.Contains(ch) Then word2 = word2.Replace(ch, "")
                                    Next

                                    'Padding the shorter with spaces
                                    Dim maxLength As Integer
                                    If word1.Length > word2.Length Then
                                        maxLength = word1.Length
                                    Else
                                        maxLength = word2.Length
                                    End If
                                    word1 = word1.PadRight(maxLength)
                                    word2 = word2.PadRight(maxLength)

                                    For i As Integer = 0 To maxLength - 1
                                        Dim char1 As Integer = AscW(word1(i))
                                        Dim char2 As Integer = AscW(word2(i))
                                        Dim char1AndChar2 As Integer = (char1 << 16) Or char2

                                        'If key1 > key2 Then
                                        '    Dim temp As String = key1
                                        '    key1 = key2
                                        '    key2 = temp
                                        'End If
                                        'If Me.DDs(key1)(key2) = -1 Then
                                        '    Dim aoaoe As Int16 = 0
                                        'End If
                                        Me.DDMatrixDegrees(var1)(var2) += Me.DDs(char1AndChar2)
                                        Me.DDCharCorrespondences(char1AndChar2) += 1
                                        Me.DDMatrixCorrespondences(var1)(var2) += 1
                                    Next
                                End If
                            Next
                        Next

                        'Else
                        'MsgBox("Gloss '" & gl.Name & "' for " & var2.Name & " has a mismatched number of aligned renderings and groupings which will result in an incorrectly populated Degrees of Difference grid!")
                        'Exit Sub
                        'End If
                        'Else
                        'MsgBox("Gloss '" & gl.Name & "' for " & var1.Name & " has a mismatched number of aligned renderings and groupings which will result in an incorrectly populated Degrees of Difference grid!")
                        'Exit Sub
                        'End If


                    Next
                Next
            Next
            'The diagonal gets twice as many counts because there is only one place on the grid for a letter to itself
            For Each ch As String In Me.UsedCharsList
                Dim char1AndChar2 As Integer = (AscW(ch) << 16) Or AscW(ch)
                If Me.DDCharCorrespondences(char1AndChar2) <> -1 Then Me.DDCharCorrespondences(char1AndChar2) \= 2
            Next
            For Each var1 As Variety In Me.AssociatedComparison.CurrentVarietySort
                For Each var2 As Variety In Me.AssociatedComparison.CurrentVarietySort
                    If var1 Is var2 Then Continue For
                    Try
                        Me.DDMatrixRatio(var1)(var2) = CType(100.0 * CType(Me.DDMatrixDegrees(var1)(var2), Double) / CType(Me.DDMatrixCorrespondences(var1)(var2), Double), Integer)
                    Catch ex As Exception
                        Me.DDMatrixRatio(var1)(var2) = 0
                    End Try

                Next
            Next
        End Sub

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
    End Class
    Public Class COMPASSCalcEntry
        Public Strength As Double = 0.0
        Public Occurences As New List(Of Gloss)
    End Class
    Public Class COMPASSGlossEntry
        Public gl As Gloss
        Public Form As String = ""
        Public PaddedForm1 As String = ""
        Public PaddedForm2 As String = ""
        Public AverageStrength As Double = 0.0
    End Class
    Public Class COMPASSCalculation
        Public UsedChars As New List(Of String)
        Public CharPairRecords As New Dictionary(Of String, COMPASSCalcEntry)
        Public GlossValues As New List(Of COMPASSGlossEntry)
        Public DisplayedGlosses As New List(Of COMPASSGlossEntry)
        Public strengthCounts As New StrengthCountsSummary()
        Public CurrentVarietyIndex1 As Integer = 0
        Public CurrentVarietyIndex2 As Integer = 1
        Public UsedGlosses As New Dictionary(Of String, Dictionary(Of Gloss, Integer))
        Public CurrentChar1Index As Integer = 0
        Public CurrentChar2Index As Integer = 0
    End Class
    Public Class StrengthCountsSummary
        Public Eq1 As Integer = 0
        Public Gte75lt1 As Integer = 0
        Public Gte50lt75 As Integer = 0
        Public Gte25lt50 As Integer = 0
        Public Gte0lt25 As Integer = 0
        Public Gten25lt0 As Integer = 0
        Public Gten50ltn25 As Integer = 0
        Public Gten75ltn50 As Integer = 0
        Public Gtn1ltn75 As Integer = 0
        Public Eqn1 As Integer = 0
    End Class

    Public Class WordSurvData
        'This class wraps all of the data objects and serves as the data layer's interface.
        'Any code that uses the data layer MUST use this object's methods ONLY.
        Public filename As String
        Public Dictionaries As New List(Of Dictionary)
        Public CurrentDictionary As Dictionary = Nothing
        Public Surveys As New List(Of Survey)
        Public CurrentSurvey As Survey = Nothing
        Public Comparisons As New List(Of Comparison)
        Public CurrentComparison As Comparison = Nothing
        Public CurrentTab As Integer = 0

        Public PrimaryLanguage As String = "Primary Gloss"
        Public SecondaryLanguage As String = "Secondary Gloss"

        Public PrimaryFont As New Font("Microsoft Sans Serif", 8)
        Public SecondaryFont As New Font("Microsoft Sans Serif", 8)
        Public TranscriptionFont As New Font("Microsoft Sans Serif", 8)

        Public Function GetCurrentComparisonsDDCurrentRowIndex() As Integer
            Return Me.CurrentComparison.AssociatedDegreesOfDifference.CurrentRowIndex
        End Function
        Public Function GetCurrentCOMPASSStrengthsSummaryCellAddress() As CellAddress
            Return Me.CurrentComparison.CurrentCOMPASSStrengthsSummaryCellAddress
        End Function
        Public Function GetDefaultGlossName(ByVal startingNum As Integer) As String
Restart:
            For Each gl As Gloss In Me.CurrentDictionary.CurrentSort.Glosses
                If gl.Name = "Gloss " & startingNum.ToString Then
                    startingNum += 1
                    GoTo Restart
                End If
            Next
            Return "Gloss " & startingNum
        End Function
        Public Function GetComparisonGlossValue(ByVal rowIndex As Integer, ByVal colIndex As Integer) As String
            Try
                Return Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses(rowIndex).Name
            Catch ex As Exception
                Return ""
            End Try
        End Function
        Public Function GetComparisonEntryValue(ByVal rowIndex As Integer, ByVal colIndex As Integer) As String
            Try
                Dim theGloss As Gloss = Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentGloss
                Dim theVariety As Variety = Me.CurrentComparison.CurrentVarietySort(rowIndex)
                Dim theVarEntry As VarietyEntry = theVariety.VarietyEntries(theGloss)

                Select Case colIndex
                    Case 0
                        Return theVariety.Name
                    Case 1
                        Return theVarEntry.Transcription
                    Case 2
                        Return theVarEntry.PluralFrame
                    Case 3
                        Return Me.CurrentComparison.ComparisonEntries(theVarEntry).AlignedRendering
                    Case 4
                        Return Me.CurrentComparison.ComparisonEntries(theVarEntry).Grouping
                    Case 5
                        Return Me.CurrentComparison.ComparisonEntries(theVarEntry).Notes
                    Case 6
                        Return Me.CurrentComparison.ComparisonEntries(theVarEntry).Exclude.ToString
                    Case Else
                        Throw New AccessViolationException
                End Select
            Catch ex As Exception
                Return ""
            End Try
        End Function
        Public Function GetComparisonNames() As String()
            Dim compNames As New List(Of String)
            For Each comp As Comparison In Me.Comparisons
                compNames.Add(comp.Name)
            Next
            Return compNames.ToArray
        End Function
        Public Function COMPASSValuesExist() As Boolean
            If Me.CurrentComparison IsNot Nothing Then
                Return Me.CurrentComparison.COMPASSCalculations IsNot Nothing
            Else
                Return False
            End If
        End Function
        Public Function GetCurrentComparisonIndex() As Integer
            Try
                Return Me.Comparisons.IndexOf(Me.CurrentComparison)
            Catch ex As Exception
            End Try
        End Function
        Public Function GetCurrentComparisonsVarietyNames() As String()
            Dim varNames As New List(Of String)
            For Each var As Variety In Me.CurrentComparison.CurrentVarietySort
                varNames.Add(var.Name)
            Next
            Return varNames.ToArray
        End Function
        Public Function GetCurrentComparisonDescription() As String
            Try
                Return Me.CurrentComparison.Description
            Catch ex As Exception
                Return ""
            End Try
        End Function
        Public Function GetCurrentComparisonAnalysisVarietyNames() As String()
            Dim varNames As New List(Of String)
            For Each var As Variety In Me.CurrentComparison.CurrentVarietySort
                varNames.Add(var.Name)
            Next
            Return varNames.ToArray
        End Function
        Public Function GetCurrentComparisonsSortNames() As String()
            Dim srtNames As New List(Of String)
            For Each srt As Sort In Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.Sorts  'AJW why not set?
                srtNames.Add(srt.Name)
            Next
            Return srtNames.ToArray
        End Function
        Public Function GetCurrentComparisonsCurrentSortIndex() As Integer
            Try
                Return Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.Sorts.IndexOf(Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort)
            Catch ex As Exception
                Return -1
            End Try
        End Function
        Public Function GetCurrentComparisonsCurrentGlossIndex() As Integer
            Try
                Return Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses.IndexOf(Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentGloss)
            Catch ex As Exception
                Return -1
            End Try
        End Function
        Public Function GetCurrentComparisonsCurrentVarietyIndex() As Integer
            Try
                Return Me.CurrentComparison.CurrentVarietySort.IndexOf(Me.CurrentComparison.CurrentVariety)
            Catch ex As Exception
                Return -1
            End Try
        End Function
        Public Function GetCurrentComparisonAnalysisVarietyIndex() As Integer
            Try
                Return Me.CurrentComparison.CurrentVarietySort.IndexOf(Me.CurrentComparison.AssociatedAnalysis.CurrentVariety)
            Catch ex As Exception
                Return -1
            End Try
        End Function
        Public Function GetCurrentComparisonAnalysisVarietyColumnIndex() As Integer
            Try
                Return Me.CurrentComparison.AssociatedAnalysis.CurrentVarietyColumnIndex
            Catch ex As Exception
                Return -1
            End Try
        End Function
        Public Function GetCurrentComparisonName() As String
            Try
                Return Me.CurrentComparison.Name
            Catch ex As Exception
                Return ""
            End Try
        End Function
        Public Function GetCurrentComparisonsCurrentVarietyColumnIndex() As Integer
            Try
                Return Me.CurrentComparison.CurrentVarietyColumnIndex
            Catch ex As Exception
                Return -1
            End Try
        End Function
        Public Function GetCurrentComparisonsGlossCount() As Integer
            Try
                Return Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses.Count
            Catch ex As Exception
                Return 0
            End Try
        End Function
        Public Function GetCurrentComparisonsVarietyCount() As Integer
            Try
                Return Me.CurrentComparison.AssociatedSurvey.Varieties.Count
            Catch ex As Exception
                Return 0
            End Try
        End Function
        Public Function GetComparisonAnalysisTallyValue(ByVal rowIndex As Integer, ByVal colIndex As Integer) As Integer
            Try
                Dim compAnal As ComparisonAnalysis = Me.CurrentComparison.AssociatedAnalysis
                Return compAnal.TallyMatrix(compAnal.AssociatedComparison.CurrentVarietySort(rowIndex))(compAnal.AssociatedComparison.CurrentVarietySort(colIndex))
            Catch ex As Exception
                Return 0
            End Try
        End Function
        Public Function GetComparisonAnalysisTotalValue(ByVal rowIndex As Integer, ByVal colIndex As Integer) As Integer
            Try
                Dim compAnal As ComparisonAnalysis = Me.CurrentComparison.AssociatedAnalysis
                Return compAnal.TotalMatrix(compAnal.AssociatedComparison.CurrentVarietySort(rowIndex))(compAnal.AssociatedComparison.CurrentVarietySort(colIndex))
            Catch ex As Exception
                Return 0
            End Try
        End Function
        Public Function GetComparisonAnalysisPercentValue(ByVal rowIndex As Integer, ByVal colIndex As Integer) As Integer
            Try
                Dim compAnal As ComparisonAnalysis = Me.CurrentComparison.AssociatedAnalysis
                Return compAnal.PercentMatrix(compAnal.AssociatedComparison.CurrentVarietySort(rowIndex))(compAnal.AssociatedComparison.CurrentVarietySort(colIndex))
            Catch ex As Exception
                Return 0
            End Try
        End Function
        Public Function GetComparisonStatistics() As String
            Dim countWithoutGroupings As Integer = 0
            Dim countExcluded As Integer = 0
            Dim percentExcluded As Double = 0
            Dim glossesWithExcludedItems As Integer = 0
            Dim glossesWithAllExcludedItems As Integer = 0
            Dim glossesWithSynonyms As Integer = 0
            Dim glossesWithMissingGroupings As Integer = 0
            Dim totalTranscriptions As Integer = 0
            Dim glosses As List(Of Gloss) = Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses
            Dim varieties As List(Of Variety) = Me.CurrentComparison.AssociatedSurvey.Varieties

            For Each gl As Gloss In glosses
                Dim thisGlossHasExclusions As Boolean = False
                Dim thisGlossHasAllExcluded As Boolean = True
                Dim thisGlossIsMissingGroupings As Boolean = False
                Dim thisGlossHasSynonyms As Boolean = False
                For Each var As Variety In Me.CurrentComparison.AssociatedSurvey.Varieties
                    Dim varEntry As VarietyEntry = var.VarietyEntries(gl)
                    Dim compEntry As ComparisonEntry = Me.CurrentComparison.ComparisonEntries(varEntry)
                    If compEntry.Exclude = "" Then totalTranscriptions += 1
                    If compEntry.Grouping = "" And compEntry.Exclude = "" Then
                        countWithoutGroupings += 1
                        thisGlossIsMissingGroupings = True
                    End If
                    If compEntry.Exclude <> "" Then
                        countExcluded += 1
                        thisGlossHasExclusions = True
                    Else
                        thisGlossHasAllExcluded = False
                    End If
                    If varEntry.Transcription.Contains(",") Then thisGlossHasSynonyms = True
                Next
                If thisGlossHasExclusions Then glossesWithExcludedItems += 1
                If thisGlossHasAllExcluded Then glossesWithAllExcludedItems += 1
                If thisGlossIsMissingGroupings Then glossesWithMissingGroupings += 1
                If thisGlossHasSynonyms Then glossesWithSynonyms += 1
            Next

            Dim transWithoutGroupings As Integer
            If totalTranscriptions = 0 Then
                transWithoutGroupings = 0
            Else
                transWithoutGroupings = CInt(countWithoutGroupings / totalTranscriptions * 100.0)
            End If
            percentExcluded = (countExcluded * 100.0) / Double.Parse(IIf(Me.CurrentComparison.ComparisonEntries.Count = 0, 1, Me.CurrentComparison.ComparisonEntries.Count).ToString)
            Return "Transcriptions Without Groupings (" & countWithoutGroupings.ToString() & ")" & vbCrLf & _
                   "Percent Transcriptions Without Groupings (" & transWithoutGroupings.ToString() & "%)" & vbCrLf & _
                   "Transcriptions Excluded (" & countExcluded.ToString() & ")" & vbCrLf & _
                   "Percent Transcriptions Excluded (" & CType(percentExcluded, Integer).ToString & "%)" & vbCrLf & _
                   "Glosses with at Least One Exclusion (" & glossesWithExcludedItems & ")" & vbCrLf & _
                   "Glosses with all Varieties Excluded (" & glossesWithAllExcludedItems & ")" & vbCrLf & _
                   "Glosses with at Least One Missing Grouping (" & glossesWithMissingGroupings & ")" & vbCrLf & _
                   "Glosses with at Least One Synonym (" & glossesWithSynonyms & ")"
        End Function
        Public Function GetSurveyNames() As String()
            Dim survNames As New List(Of String)
            For Each surv As Survey In Me.Surveys
                survNames.Add(surv.Name)
            Next
            Return survNames.ToArray
        End Function
        Public Function GetCurrentSurveyName() As String
            Try
                Return Me.CurrentSurvey.Name
            Catch ex As Exception
                Return ""
            End Try
        End Function
        Public Function GetCurrentVarietyName() As String
            Try
                Return Me.CurrentSurvey.CurrentVariety.Name
            Catch ex As Exception
                Return ""
            End Try
        End Function
        Public Function GetCurrentSurveysCurrentVarietyName() As String
            Try
                Return Me.CurrentSurvey.CurrentVariety.Name
            Catch ex As Exception
                Return ""
            End Try
        End Function
        Public Function GetCurrentSurveysVarietyNames() As String()
            Dim varNames As New List(Of String)
            For Each var As Variety In Me.CurrentSurvey.Varieties
                varNames.Add(var.Name)
            Next
            Return varNames.ToArray
        End Function
        Public Function GetCurrentSurveysSortNames() As String()
            Dim srtNames As New List(Of String)
            For Each srt As Sort In Me.CurrentSurvey.AssociatedDictionary.Sorts
                srtNames.Add(srt.Name)
            Next
            Return srtNames.ToArray
        End Function
        Public Function GetCurrentSurveyIndex() As Integer
            Try
                Return Me.Surveys.IndexOf(Me.CurrentSurvey)
            Catch ex As Exception
                Return -1
            End Try
        End Function
        Public Function GetCurrentSurveysCurrentVarietyIndex() As Integer
            Try
                Return Me.CurrentSurvey.Varieties.IndexOf(Me.CurrentSurvey.CurrentVariety)
            Catch ex As Exception
                Return -1
            End Try
        End Function
        Public Function GetCurrentSurveyLength() As Integer
            Try
                Return Me.CurrentSurvey.AssociatedDictionary.CurrentSort.Glosses.Count
            Catch ex As Exception
                Return 0
            End Try
        End Function
        Public Function GetCurrentSurveysCurrentGlossIndex() As Integer
            Try
                Return Me.CurrentSurvey.AssociatedDictionary.CurrentSort.Glosses.IndexOf(Me.CurrentSurvey.AssociatedDictionary.CurrentGloss)
            Catch ex As Exception
                Return -1
            End Try
        End Function
        Public Function GetCurrentSurveysCurrentVarietyEntryColumnIndex() As Integer
            Try
                Return Me.CurrentSurvey.CurrentVarietyEntryColumnIndex
            Catch ex As Exception
                Return -1
            End Try
        End Function
        Public Function GetCurrentComparisonsDDExcludedChars() As String
            Return Me.CurrentComparison.AssociatedDegreesOfDifference.ExcludedChars
        End Function
        Public Function GetTranscriptionValue(ByVal transRow As Integer, ByVal transCol As Integer) As String
            Try
                Select Case transCol
                    Case 0
                        Return Me.CurrentSurvey.AssociatedDictionary.CurrentSort.Glosses(transRow).Name
                    Case 1
                        Return Me.CurrentSurvey.CurrentVariety.VarietyEntries(Me.CurrentSurvey.AssociatedDictionary.CurrentSort.Glosses(transRow)).Transcription
                    Case 2
                        Return Me.CurrentSurvey.CurrentVariety.VarietyEntries(Me.CurrentSurvey.AssociatedDictionary.CurrentSort.Glosses(transRow)).PluralFrame
                    Case 3
                        Return Me.CurrentSurvey.CurrentVariety.VarietyEntries(Me.CurrentSurvey.AssociatedDictionary.CurrentSort.Glosses(transRow)).Notes
                    Case Else
                        Throw New IndexOutOfRangeException
                End Select
            Catch ex As IndexOutOfRangeException
                Throw ex
            Catch ex As Exception
                Return ""
            End Try
        End Function
        Public Function MergeCurrentSurvey() As Boolean
            'Back in the bad old days of WordSurv 6, the dictionaries and surveys and comparisons were all completely independent.
            'The user created one or more dictionaries, and added any glosses desired into the survey from any dictionary.  
            'A comparison could use some or all of the varieties from as many surveys as desired.  This is useful to the surveyors,
            'but it required separate dictionary entries, variety entries, and comparison entries.  In the current design, 
            'we have required each comparison to have exactly one survey and each survey to have exactly one dictionary.

            'We provide this merge survey functionality to serve as a substitute for the restriction of one dictionary per
            'survey and one survey per comparison.

            'To save on the number of choices and combo boxes the user has to wade through, we assume the current survey is one of the
            'two they desire to merge.

            'Get the other survey to merge with.
            Dim frmSurveySelect As New SurveySelectForm()
            frmSurveySelect.Text = "Merge Surveys"
            frmSurveySelect.lblPrompt.Text = "Select the Survey to combine with the current Survey " & Me.GetCurrentSurveyName() & "."
            For Each surv As Survey In Me.Surveys
                If surv.Name <> Me.GetCurrentSurveyName() Then frmSurveySelect.cboSurveySelection.Items.Add(surv.Name)
            Next
            frmSurveySelect.cboSurveySelection.SelectedIndex = 0
            If Not frmSurveySelect.ShowDialog = DialogResult.OK Then Return False
            Dim surv1 As Survey = Me.CurrentSurvey
            Dim surv2 As Survey = Me.Surveys(frmSurveySelect.cboSurveySelection.SelectedIndex)


            'If there are no glosses in either of the dictionary, there is no point of doing the merge.
            If surv1.AssociatedDictionary.CurrentSort.Glosses.Count = 0 Then Return False
            If surv2.AssociatedDictionary.CurrentSort.Glosses.Count = 0 Then Return False

            'We first create a new dictionary that contains the glosses of both surveys.  We assume most of the glosses
            'between these surveys are the same, since otherwise it would not be all that useful to merge them.
            'However, there are bound to be some that are in the first and not the second, and vice versa.
            'When a gloss is in both surveys, we put only one copy of the gloss in the new dictionary and combine
            'the field tip and comments.  We define two glosses as being the same if the primary language form is exactly
            'the same for both.

            'In the merging process, we traverse the linked structure of the data objects of the two original surveys as we
            'build the new merged one.  The linked structure is very good for accessing, but it is slighly more difficult to copy.
            'We build several temporary helper data structures to allow us to fit the new objects in the right places.

            'This list contains all of the original glosses from both surveys' dictionaries.  Remember that each dictionary
            'has its own copy of each gloss, even if two dictionaries share glosses with the same forms.
            Dim allOldGlossesWithDuplicates As New List(Of Gloss)
            allOldGlossesWithDuplicates.AddRange(surv1.AssociatedDictionary.CurrentSort.Glosses)
            allOldGlossesWithDuplicates.AddRange(surv2.AssociatedDictionary.CurrentSort.Glosses)

            'We then build a list that contains one of each unique gloss between the two dictionaries.  These are the
            'gloss objects that will go into the new dictionary.  Note that this is a hash from a gloss's primary
            'name to the actual gloss object.  This mapping is useful later on, allowing us to get a new gloss
            'object given the gloss with the same name in the old dictionaries.

            'For each of these temporary data structures, one with 'old' in the name contains objects from
            'one of the two original merged data sets, and one with 'new' in the name contains the pieces
            'of the new merged data structure we are building.
            Dim allNewGlossesNoDuplicates As New Dictionary(Of String, Gloss)
            For Each gl As Gloss In allOldGlossesWithDuplicates 'Go through the one with duplicates
                If Not allNewGlossesNoDuplicates.ContainsKey(gl.Name) Then 'If we haven't seen this one before, stick in it
                    allNewGlossesNoDuplicates.Add(gl.Name, gl.Copy())
                Else                                                       'If we have, copy over the nonsimilar parts

                    If allNewGlossesNoDuplicates(gl.Name).FieldTip <> "" Then allNewGlossesNoDuplicates(gl.Name).FieldTip &= " / "
                    If allNewGlossesNoDuplicates(gl.Name).Comments <> "" Then allNewGlossesNoDuplicates(gl.Name).Comments &= " / "

                    allNewGlossesNoDuplicates(gl.Name).FieldTip &= gl.FieldTip
                    allNewGlossesNoDuplicates(gl.Name).Comments &= gl.Comments
                End If
            Next

            'These two structures allow us to distinguish which original survey a gloss belongs to.
            Dim oldSurv1Glosses As New Dictionary(Of String, Gloss)
            For Each gl As Gloss In surv1.AssociatedDictionary.CurrentSort.Glosses
                oldSurv1Glosses.Add(gl.Name, gl)
            Next
            Dim oldSurv2Glosses As New Dictionary(Of String, Gloss)
            For Each gl As Gloss In surv2.AssociatedDictionary.CurrentSort.Glosses
                oldSurv2Glosses.Add(gl.Name, gl)
            Next

            'Here we make use of the gloss name to gloss object mapping we made above.  To make a new merged sort,
            Dim newSorts As New List(Of Sort)
            For Each srt As Sort In surv1.AssociatedDictionary.Sorts                      'For each old sort in the first survey's dictionary,
                Dim newSort As New Sort(srt.Name & " from " & surv1.Name)                 '  Create a new empty sort object
                For Each gl As Gloss In srt.Glosses                                       '  For each old gloss in the old sort
                    newSort.Glosses.Add(allNewGlossesNoDuplicates(gl.Name))    '    Use the old gloss's primary language form to find the new gloss
                Next                                                                      '    and add it to the sort

                For Each glName As String In allNewGlossesNoDuplicates.Keys               'Also, loop over all the new glosses again,
                    If Not oldSurv1Glosses.ContainsKey(glName) Then                  '  If there are any glosses that are not in the first survey's dictionary,
                        newSort.Glosses.Add(allNewGlossesNoDuplicates(glName))            '  add them to the end of this sort, again making use of the mapping.
                    End If
                Next
                newSorts.Add(newSort)
            Next
            'Do it again for the sorts in the second survey's dictionary.
            For Each srt As Sort In surv2.AssociatedDictionary.Sorts
                Dim newSort As New Sort(srt.Name & " from " & surv2.Name)
                For Each gl As Gloss In srt.Glosses
                    newSort.Glosses.Add(allNewGlossesNoDuplicates(gl.Name))
                Next
                For Each glName As String In allNewGlossesNoDuplicates.Keys
                    If Not oldSurv2Glosses.ContainsKey(glName) Then
                        newSort.Glosses.Add(allNewGlossesNoDuplicates(glName))
                    End If
                Next
                newSorts.Add(newSort)
            Next

            'Now make a new dictionary with these sorts.
            Me.CurrentDictionary = New Dictionary(surv1.Name & " and " & surv2.Name & " Dict") 'Provide a default name
            Me.Dictionaries.Add(Me.CurrentDictionary)
            Me.CurrentDictionary.Sorts.AddRange(newSorts)
            Me.CurrentDictionary.CurrentSort = Me.CurrentDictionary.Sorts(0)
            Me.CurrentDictionary.CurrentGloss = Me.CurrentDictionary.CurrentSort.Glosses(0)

            'Allow the user to rename the new dictionary.
            Dim frmInput As New InputForm("Name Dictionary", "Enter the merged Dictionary's name.", ValidationType.DICTIONARY_NAME, Me, Me.GetCurrentDictionaryName())

            If Not frmInput.ShowDialog = Windows.Forms.DialogResult.OK Then Return False
            Me.RenameCurrentDictionary(frmInput.txtInput.Text)

            'This structure maps a new variety back to the old one, which is necessary for merging comparisons later.
            Dim newToOldVars As New Dictionary(Of Variety, Variety)

            'Now that the dictionary is made, we merge the two surveys.
            Dim newVars As New List(Of Variety)
            For Each var As Variety In surv1.Varieties                                   'For each old variety in the first survey
                Dim newVar As New Variety(Me.CurrentDictionary, var.Name, False)         '  Create a new survey
                newVar.Description = var.Description                                     '  Copy its simple attributes
                newVar.AssociatedDictionary = Me.CurrentDictionary
                For Each glName As String In allNewGlossesNoDuplicates.Keys              '  For each gloss in our new merged dictionary
                    If oldSurv1Glosses.ContainsKey(glName) Then
                        'If the gloss was in the old first survey, that means a VarietyEntry exists for this gloss in the old first survey, so just copy it over
                        newVar.VarietyEntries.Add( _
                                    allNewGlossesNoDuplicates(glName), _
                                    var.VarietyEntries(oldSurv1Glosses(glName)).Copy())
                    Else
                        'Otherwise this is a new gloss/variety combination that was introduced when we added more glosses to the dictionary,
                        'so make a new blank VarietyEntry.
                        newVar.VarietyEntries.Add( _
                                                  allNewGlossesNoDuplicates(glName), _
                                                  New VarietyEntry())
                    End If
                Next
                newVar.CurrentVarietyEntry = newVar.VarietyEntries(newVar.AssociatedDictionary.CurrentGloss)
                newToOldVars.Add(newVar, var)
                newVars.Add(newVar)
            Next
            'Do the same thing for the second survey
            For Each var As Variety In surv2.Varieties
                Dim newVar As New Variety(Me.CurrentDictionary, var.Name, False)
                newVar.Description = var.Description
                newVar.AssociatedDictionary = Me.CurrentDictionary
                For Each glName As String In allNewGlossesNoDuplicates.Keys
                    If oldSurv2Glosses.ContainsKey(glName) Then
                        newVar.VarietyEntries.Add(allNewGlossesNoDuplicates(glName), var.VarietyEntries(oldSurv2Glosses(glName)).Copy())
                    Else
                        newVar.VarietyEntries.Add(allNewGlossesNoDuplicates(glName), New VarietyEntry())
                    End If
                Next
                newVar.CurrentVarietyEntry = newVar.VarietyEntries(newVar.AssociatedDictionary.CurrentGloss)
                newToOldVars.Add(newVar, var)
                newVars.Add(newVar)
            Next

            'Now make the new survey 
            Dim newSurv As New Survey(Me.CurrentDictionary, surv1.Name & " and " & surv2.Name & " combined")
            newSurv.Description = "Merger of """ & surv1.Description & """ and """ & surv2.Description
            newSurv.Varieties.AddRange(newVars)
            newSurv.CurrentVariety = newSurv.Varieties(0)
            Me.Surveys.Add(newSurv)
            Me.CurrentSurvey = newSurv

            For Each var As Variety In newSurv.Varieties
                For Each otherVar As Variety In newSurv.Varieties
                    If var Is otherVar Then Continue For
                    If var.Name = otherVar.Name Then
                        otherVar.Name &= " from merge"
                    End If
                Next
            Next

            'Allow the user to rename the merged survey also
            Dim frmInput2 As New InputForm("Name Survey", "Enter the merged Survey's name.", ValidationType.SURVEY_NAME, Me, Me.GetCurrentSurveyName())
            If Not frmInput2.ShowDialog = Windows.Forms.DialogResult.OK Then Return False
            Me.RenameCurrentSurvey(frmInput2.txtInput.Text)


            'Now begin merging comparisons.  We can't just blindly merge comparisons, because there is no way to tell which comparisons
            'the user wants to merge.  We don't want to cross every comparison from the first survey with every comparison from the second,
            'since that would quickly lead to a huge number of comparisons generated.  Therefore we bring up a form and allow the
            'user to select which comparisons to merge.

            'Make a list of the comparisons from each survey while also adding them to the combo boxes on the form.

            Dim frmComparisonMerge As New ComparisonMergeForm
            Dim oldSurv1Comps As New List(Of Comparison)
            Dim oldSurv2Comps As New List(Of Comparison)
            For Each comp As Comparison In Me.Comparisons
                If comp.AssociatedSurvey Is surv1 Then
                    oldSurv1Comps.Add(comp)
                    frmComparisonMerge.cboComparison1.Items.Add(comp.Name)
                End If
                If comp.AssociatedSurvey Is surv2 Then
                    oldSurv2Comps.Add(comp)
                    frmComparisonMerge.cboComparison2.Items.Add(comp.Name)
                End If
            Next

            'If there are no comparisons that use this survey, don't try merging the comparisons.
            If frmComparisonMerge.cboComparison1.Items.Count <> 0 And frmComparisonMerge.cboComparison2.Items.Count <> 0 Then

                frmComparisonMerge.cboComparison1.SelectedIndex = 0
                frmComparisonMerge.cboComparison2.SelectedIndex = 0

                Dim newComps As New List(Of Comparison)
                'We keep showing the user the selection form until they want to stop.  The button marked 'Merge' is the form's OK button, and the button 'Stop' is cancel.
                While frmComparisonMerge.ShowDialog() <> DialogResult.Cancel
                    'Get the old comparisons
                    Dim oldComp1 As Comparison = oldSurv1Comps(frmComparisonMerge.cboComparison1.SelectedIndex)
                    Dim oldComp2 As Comparison = oldSurv2Comps(frmComparisonMerge.cboComparison2.SelectedIndex)

                    'Create a new comparison and copy simple attributes
                    Dim newComp As New Comparison(newSurv, frmComparisonMerge.txtName.Text, False)
                    If Not (oldComp1.Description = "" And oldComp2.Description = "") Then
                        newComp.Description = oldComp1.Description & " / " & oldComp2.Description 'If both descriptions have text, put a / between them, otherwise leave it out.
                    Else
                        newComp.Description = oldComp1.Description & oldComp2.Description         'This prevents a / from being out there for no reason.
                    End If

                    'For any gloss/variety combination that was not existant in either of the old comparisons, we need to create
                    'new comparison entries to fill the space.
                    For Each gl As Gloss In Me.CurrentDictionary.CurrentSort.Glosses     'For each new gloss
                        For Each var As Variety In newSurv.Varieties                     '  For each new variety
                            Dim oldVar As Variety = newToOldVars(var)                    '    Find the old gloss and variety that correspond with these new ones

                            'Create a new comparison entry for this gloss/variety combination.
                            Dim newCompEntry As New ComparisonEntry("")

                            'If this gloss is in the old first survey, that means there is a comparison entry for this gloss/variety pair,
                            'so copy over the old comparison entry data.
                            If oldSurv1Glosses.ContainsKey(gl.Name) Then
                                Dim oldGloss1 As Gloss = oldSurv1Glosses(gl.Name)
                                If oldVar.VarietyEntries.ContainsKey(oldGloss1) Then
                                    Dim oldCompEntry1 As ComparisonEntry = oldComp1.ComparisonEntries(oldVar.VarietyEntries(oldGloss1))
                                    newCompEntry.AlignedRendering = oldCompEntry1.AlignedRendering
                                    newCompEntry.Grouping = oldCompEntry1.Grouping
                                    newCompEntry.Notes = oldCompEntry1.Notes
                                    newCompEntry.Exclude = oldCompEntry1.Exclude
                                End If
                            End If

                            'Otherwise if this gloss is in the old second survey, copy the old data from it.  Only one of these if clauses should
                            'be true at a time.
                            If oldSurv2Glosses.ContainsKey(gl.Name) Then
                                Dim oldGloss2 As Gloss = oldSurv2Glosses(gl.Name)
                                If oldVar.VarietyEntries.ContainsKey(oldGloss2) Then
                                    Dim oldCompEntry2 As ComparisonEntry = oldComp2.ComparisonEntries(oldVar.VarietyEntries(oldGloss2))
                                    newCompEntry.AlignedRendering = oldCompEntry2.AlignedRendering
                                    newCompEntry.Grouping = oldCompEntry2.Grouping
                                    newCompEntry.Notes = oldCompEntry2.Notes
                                    newCompEntry.Exclude = oldCompEntry2.Exclude
                                End If
                            End If

                            newComp.ComparisonEntries.Add(var.VarietyEntries(gl), newCompEntry)
                        Next
                    Next
                    'There is not a good way to copy over the variety sorts, so we simply reset it.
                    newComp.CurrentVarietySort.AddRange(newSurv.Varieties)
                    newComp.DefaultVarietySort.AddRange(newSurv.Varieties)
                    newComp.CurrentVariety = newComp.CurrentVarietySort(0)


                    'The last thing to merge is the DD grids.  Nothing else in the program currently is saved to the file, so we end here.  All of the analyses
                    'are recalculted on going to their respective tabs.
                    newComp.AssociatedDegreesOfDifference = New DegreesOfDifferenceGrid(newComp)
                    newComp.AssociatedDegreesOfDifference.CalculateUsedChars()
                    oldComp1.AssociatedDegreesOfDifference.CalculateUsedChars()
                    oldComp2.AssociatedDegreesOfDifference.CalculateUsedChars()

                    'For each character crossed with every other character,
                    For Each usedChar1 As String In newComp.AssociatedDegreesOfDifference.UsedCharsList
                        For Each usedChar2 As String In newComp.AssociatedDegreesOfDifference.UsedCharsList
                            Dim char1AndChar2 As Integer = (AscW(usedChar1) << 16) Or AscW(usedChar2)
                            'If that pair was in the old second table, bring the value over.
                            If oldComp2.AssociatedDegreesOfDifference.DDs.ContainsKey(char1AndChar2) Then
                                newComp.AssociatedDegreesOfDifference.DDs(char1AndChar2) = oldComp2.AssociatedDegreesOfDifference.DDs(char1AndChar2)
                            End If
                            'If that pair was in the old first table, bring it over.  When both old DD grids had a value, we use the one from the first grid.
                            'Averaging the DDs or some other means of combination would probably not be what the user intended.  We assume that most of the DD
                            'values will be the same if they are merging similar surveys.
                            If oldComp1.AssociatedDegreesOfDifference.DDs.ContainsKey(char1AndChar2) Then
                                newComp.AssociatedDegreesOfDifference.DDs(char1AndChar2) = oldComp1.AssociatedDegreesOfDifference.DDs(char1AndChar2)
                            End If

                        Next
                    Next
                    newComps.Add(newComp)
                    'So that the user has some feedback, we display the comparison pairs that they have already combined on the comparison selection form.
                    frmComparisonMerge.lstPreviouslyMerged.Items.Add(oldComp1.Name & " and " & oldComp2.Name)
                End While
                Me.Comparisons.AddRange(newComps)
                Try
                    Me.CurrentComparison = newComps(0)
                Catch ex As Exception
                End Try
            End If

            Return True
        End Function
        Public Function GetCurrentSurveyDescription() As String
            Try
                Return Me.CurrentSurvey.Description
            Catch ex As Exception
                Return ""
            End Try
        End Function
        Public Function GetCurrentSurveysCurrentVarietyDescription() As String
            Try
                Return Me.CurrentSurvey.CurrentVariety.Description
            Catch ex As Exception
                Return ""
            End Try
        End Function
        Public Function GetCurrentSurveysCurrentSortIndex() As Integer
            Try
                Return Me.CurrentSurvey.AssociatedDictionary.Sorts.IndexOf(Me.CurrentSurvey.AssociatedDictionary.CurrentSort)
            Catch ex As Exception
            End Try
        End Function
        Public Function GetCurrentSurveysAssociatedDictionaryName() As String
            Try
                Return Me.CurrentSurvey.AssociatedDictionary.Name
            Catch ex As Exception
                Return ""
            End Try
        End Function
        Public Function GetDictionaryNames() As String()
            Dim dictNames As New List(Of String)
            For Each dict As Dictionary In Me.Dictionaries
                dictNames.Add(dict.Name)
            Next
            Return dictNames.ToArray
        End Function
        Public Function GetCurrentDictionaryName() As String
            Try
                Return Me.CurrentDictionary.Name
            Catch ex As Exception
                Return ""
            End Try
        End Function
        Public Function GetCurrentDictionarysCurrentSortName() As String
            Try
                Return Me.CurrentDictionary.CurrentSort.Name
            Catch ex As Exception
                Return ""
            End Try
        End Function
        Public Function GetCurrentDictionaryLength() As Integer
            'All the sorts have the same glosses in them, so pick an arbitrary one (say 0) and get its length.
            Try
                Return Me.CurrentDictionary.Sorts(0).Glosses.Count
            Catch ex As Exception
                Return 0
            End Try
        End Function
        Public Function GetCurrentComparisonDictionaryLength() As Integer
            Try
                Return Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.Sorts(0).Glosses.Count
            Catch ex As Exception
                Return 0
            End Try
        End Function
        Public Function GetCurrentVarietysTranscriptionCount() As Integer
            Try
                Return (Me.CurrentSurvey.AssociatedDictionary.Sorts(0).Glosses.Count)
            Catch ex As Exception
                Return 0
            End Try
        End Function
        Public Function GetCurrentVarietysNumberTranscribed() As Integer
            Try
                Dim count As Integer = 0
                For Each varEntry As VarietyEntry In Me.CurrentSurvey.CurrentVariety.VarietyEntries.Values
                    If varEntry.Transcription <> "" Then count += 1
                Next
                Return count
            Catch ex As Exception
                Return 0
            End Try
        End Function
        Public Function GetCurrentDictionarysCurrentGlossColumnIndex() As Integer
            Try
                Return Me.CurrentDictionary.CurrentGlossColumnIndex
            Catch ex As Exception
                Return -1
            End Try
        End Function
        Public Function GetCurrentDictionaryIndex() As Integer
            Try
                Return Me.Dictionaries.IndexOf(Me.CurrentDictionary)
            Catch ex As Exception
                Return -1
            End Try
        End Function
        Public Function GetCurrentDictionarySortIndex() As Integer
            Try
                Return Me.CurrentDictionary.Sorts.IndexOf(Me.CurrentDictionary.CurrentSort)
            Catch ex As Exception
            End Try
        End Function
        Public Function GetCurrentComparisonsCurrentCOMPASSChar1Index() As Integer
            Return Me.CurrentComparison.COMPASSCalculations.CurrentChar1Index
        End Function
        Public Function GetCurrentComparisonsCurrentCOMPASSChar2Index() As Integer
            Return Me.CurrentComparison.COMPASSCalculations.CurrentChar2Index
        End Function
        Private Function GetSurveysThatUseThisDictionary(ByRef dict As Dictionary) As List(Of Survey)
            Dim surveys As New List(Of Survey)
            For Each surv As Survey In Me.Surveys
                If surv.AssociatedDictionary.Equals(dict) Then
                    surveys.Add(surv)
                End If
            Next
            Return surveys
        End Function
        Private Function GetComparisonsThatUseThisDictionary(ByRef dict As Dictionary) As List(Of Comparison)
            Dim comparisons As New List(Of Comparison)
            For Each comp As Comparison In Me.Comparisons
                If comp.AssociatedSurvey.AssociatedDictionary.Equals(dict) Then
                    comparisons.Add(comp)
                End If
            Next
            Return comparisons
        End Function
        Public Function GetSelectedCOMPASSPhoneCoordinates() As List(Of CellAddress)
            Return Me.CurrentComparison.SelectedPhonePairCoordinates
        End Function
        Public Function GetCurrentDictionarysSortNames() As String()
            Dim sortNames As New List(Of String)
            For Each srt As Sort In Me.CurrentDictionary.Sorts
                sortNames.Add(srt.Name)
            Next
            Return sortNames.ToArray
        End Function
        Public Function GetCurrentDictionarysCurrentGlossIndex() As Integer
            Try
                Return Me.CurrentDictionary.CurrentSort.Glosses.IndexOf(Me.CurrentDictionary.CurrentGloss)
            Catch ex As Exception
                Return -1
            End Try
        End Function
        Public Function GetGlossValue(ByVal glossRow As Integer, ByVal glossCol As Integer) As String
            Try
                Return Me.CurrentDictionary.CurrentSort.Glosses(glossRow).GetByIndex(glossCol)
            Catch ex As Exception
                Return ""
            End Try
        End Function
        Public Function GetCurrentComparisonGlossesNamesUsingThisDDPair(ByVal char1 As String, ByVal char2 As String) As String()
            Dim char1AndChar2 As Integer = (AscW(char1) << 16) Or AscW(char2)
            Dim glossNames As New List(Of String)
            For Each gl As Gloss In Me.CurrentComparison.AssociatedDegreesOfDifference.GlossesUsing(char1AndChar2)
                glossNames.Add(gl.Name)
            Next
            Return glossNames.ToArray
        End Function
        Public Function GetComparisonGlossIndexFromDDUsedPhonePair(ByVal char1 As String, ByVal char2 As String, ByVal index As Integer) As Integer
            Dim glSort As List(Of Gloss) = Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses
            Dim targetGl As Gloss = Me.CurrentComparison.AssociatedDegreesOfDifference.GlossesUsing((AscW(char1) << 16) Or AscW(char2))(index)
            Return glSort.IndexOf(targetGl)
        End Function
        Public Function GetNextUngroupedGlossAndVariety() As IntIntComboMenu
            Dim currentVarietyIndex As Integer = Me.CurrentComparison.CurrentVarietySort.IndexOf(Me.CurrentComparison.CurrentVariety)
            Dim currentGlossIndex As Integer = Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses.IndexOf(Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentGloss)
            Dim numGlosses As Integer = Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses.Count
            Dim numVarieties As Integer = Me.CurrentComparison.CurrentVarietySort.Count
            Dim varietyList As List(Of Variety) = Me.CurrentComparison.CurrentVarietySort

            'Advance one variety so we don't find the variety we are on and exit
            Dim glI As Integer = currentGlossIndex
            Dim varI As Integer = currentVarietyIndex + 1
            If varI >= numVarieties Then
                varI = 0
                glI += 1
                If glI >= numGlosses Then glI = 0
            End If

            While True
                Dim thisGloss As Gloss = Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses(glI)

                'If the grouping is empty or we are back where we started, stop and return this current gloss|variety combo menu
                While True
                    Dim thisVarietyEntry As VarietyEntry = varietyList(varI).VarietyEntries(thisGloss)
                    Dim thisCompEntry As ComparisonEntry = Me.CurrentComparison.ComparisonEntries(thisVarietyEntry)
                    'AJW*** allow goto next ungrouped to include synonyms without groups
                    If (thisVarietyEntry.Transcription <> "" And thisCompEntry.Exclude = "" And ((thisCompEntry.Grouping = "") Or Not (HaveSameNumberOfCommas(thisCompEntry.AlignedRendering, thisCompEntry.Grouping)))) OrElse _
                       (glI = currentGlossIndex And varI = currentVarietyIndex) Then
                        Return New IntIntComboMenu(glI, varI)
                    End If

                    varI += 1
                    If varI >= numVarieties Then
                        varI = 0
                        Exit While
                    End If
                End While

                glI += 1
                If glI >= numGlosses Then glI = 0
            End While

            Return Nothing
        End Function
        Public Function GetCurrentComparisonsUsedCharsForDDGrid() As List(Of String)
            Return Me.CurrentComparison.AssociatedDegreesOfDifference.UsedCharsList
        End Function
        Public Function GetDDValue(ByVal rowIndex As Integer, ByVal colIndex As Integer) As Integer
            If Me.CurrentComparison Is Nothing Then Return -1
            Dim usedChars As List(Of String) = Me.CurrentComparison.AssociatedDegreesOfDifference.UsedCharsList
            If usedChars.Count = 0 Then Return -1
            Return Me.CurrentComparison.AssociatedDegreesOfDifference.DDs((AscW(usedChars(rowIndex)) << 16) Or AscW(usedChars(colIndex)))
        End Function
        Public Function GetPhonoStatsValue(ByVal rowIndex As Integer, ByVal colIndex As Integer, ByVal dataIndex As Integer) As Integer
            If Me.CurrentComparison Is Nothing OrElse Me.CurrentComparison.AssociatedDegreesOfDifference Is Nothing OrElse Me.CurrentComparison.AssociatedDegreesOfDifference.UsedCharsList Is Nothing Then Return -1

            Dim char1 As String = ""
            Dim char2 As String = ""
            Dim var1 As Variety = Nothing
            Dim var2 As Variety = Nothing
            If dataIndex = 1 Then
                char1 = Me.CurrentComparison.AssociatedDegreesOfDifference.UsedCharsList(rowIndex)
                char2 = Me.CurrentComparison.AssociatedDegreesOfDifference.UsedCharsList(colIndex)
            Else
                If rowIndex > Me.CurrentComparison.CurrentVarietySort.Count - 1 Then Return -1
                If colIndex > Me.CurrentComparison.CurrentVarietySort.Count - 1 Then Return -1
                var1 = Me.CurrentComparison.CurrentVarietySort(rowIndex)
                var2 = Me.CurrentComparison.CurrentVarietySort(colIndex)
            End If

            Select Case dataIndex
                Case 1 : Return Me.CurrentComparison.AssociatedDegreesOfDifference.DDCharCorrespondences((AscW(char1) << 16) Or AscW(char2))
                Case 2 : Return Me.CurrentComparison.AssociatedDegreesOfDifference.DDMatrixRatio(var1)(var2)
                Case 3 : Return Me.CurrentComparison.AssociatedDegreesOfDifference.DDMatrixDegrees(var1)(var2)
                Case 4 : Return Me.CurrentComparison.AssociatedDegreesOfDifference.DDMatrixCorrespondences(var1)(var2)
                Case Else : Throw New AccessViolationException
            End Select
        End Function
        Public Function GetCurrentComparisonsCOMPASSCalculationUsedChars() As List(Of String)
            Dim usedChars As List(Of String)
            Try
                usedChars = Me.CurrentComparison.COMPASSCalculations.UsedChars
            Catch ex As Exception
                usedChars = New List(Of String)
            End Try
            Return usedChars
        End Function
        Public Function GetCOMPASSPhoneOccurences(ByVal rowIndex As Integer, ByVal colIndex As Integer) As Integer
            Dim calc As COMPASSCalculation = Me.CurrentComparison.COMPASSCalculations

            Try
                Dim ch1 As String = calc.UsedChars(rowIndex)
                Dim ch2 As String = calc.UsedChars(colIndex)
                If calc.CharPairRecords.ContainsKey(ch1 & ch2) Then
                    Return calc.CharPairRecords(calc.UsedChars(rowIndex) & calc.UsedChars(colIndex)).Occurences.Count
                Else
                    Return -1
                End If
            Catch ex As Exception
                Return -1
            End Try
        End Function
        Public Function GetCOMPASSPhoneStrength(ByVal rowIndex As Integer, ByVal colIndex As Integer) As Double
            Dim calc As COMPASSCalculation = Me.CurrentComparison.COMPASSCalculations

            Try
                Dim ch1 As String = calc.UsedChars(rowIndex)
                Dim ch2 As String = calc.UsedChars(colIndex)
                If calc.CharPairRecords.ContainsKey(ch1 & ch2) Then
                    Return calc.CharPairRecords(calc.UsedChars(rowIndex) & calc.UsedChars(colIndex)).Strength
                Else
                    Return Double.NaN
                End If
            Catch ex As Exception
                Return Double.NaN
            End Try
        End Function
        Public Function GetCurrentComparisonsCurrentCOMPASSVariety1Index() As Integer
            Try
                Return Me.CurrentComparison.COMPASSCalculations.CurrentVarietyIndex1
            Catch ex As Exception
                Return 0
            End Try
        End Function
        Public Function GetCurrentComparisonsCurrentCOMPASSVariety2Index() As Integer
            Try
                Return Me.CurrentComparison.COMPASSCalculations.CurrentVarietyIndex2
            Catch ex As Exception
                Return 1 'The default value for this combo box
            End Try
        End Function
        Public Function GetCurrentComparisonsCOMPASSGlossComparedCount() As Integer
            Try
                Return Me.CurrentComparison.COMPASSCalculations.DisplayedGlosses.Count
            Catch ex As Exception
                Return 0
            End Try
        End Function
        Public Function GetCurrentComparisonsCOMPASSCognateCount() As Integer
            Return Me.CurrentComparison.COMPASSCalculations.GlossValues.Count
        End Function
        Public Function GetCOMPASSCognateStrengthsValue(ByVal rowIndex As Integer, ByVal colIndex As Integer) As String
            If Me.CurrentComparison Is Nothing Then Return ""

            Try
                Dim glVal As COMPASSGlossEntry = Me.CurrentComparison.COMPASSCalculations.DisplayedGlosses(rowIndex)
                Select Case colIndex
                    Case 0 : Return glVal.Form
                    Case 1 : Return glVal.PaddedForm1
                    Case 2 : Return glVal.PaddedForm2
                    Case 3 : Return glVal.AverageStrength.ToString("F2")
                    Case Else : Return ""
                End Select
            Catch ex As Exception
                Return ""
            End Try
        End Function
        Public Function GetCOMPASSCognateStrengthsAverageStrength(ByVal rowIndex As Integer) As Double
            Try
                Return Me.CurrentComparison.COMPASSCalculations.DisplayedGlosses(rowIndex).AverageStrength
            Catch ex As Exception
                Return -2.0
            End Try
        End Function
        Public Function GetCurrentComparisonsCOMPASSStrengthSummary() As StrengthCountsSummary
            Try
                Return Me.CurrentComparison.COMPASSCalculations.strengthCounts
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        Public Sub SetCurrentCOMPASSStrengthsSummaryCellAddress(ByVal addr As CellAddress)
            Me.CurrentComparison.CurrentCOMPASSStrengthsSummaryCellAddress = addr
        End Sub
        Public Sub SetCurrentComparison(ByVal index As Integer)
            Try
                Me.CurrentComparison = Me.Comparisons(index)
            Catch ex As Exception
            End Try
        End Sub
        Public Sub SetCurrentComparisonsCurrentGlossSort(ByVal index As Integer)
            Try
                Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort = Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.Sorts(index)
            Catch ex As Exception
            End Try
        End Sub
        Public Sub SetCurrentcomparisonsAssociatedDictionaryCurrentGloss(ByVal index As Integer)
            Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentGloss = Me.CurrentComparison.COMPASSCalculations.DisplayedGlosses(index).gl
        End Sub
        Public Sub SetCurrentComparisonDescription(ByVal val As String)
            Try
                Me.CurrentComparison.Description = val
            Catch ex As Exception
            End Try
        End Sub
        Public Sub SetCurrentComparisonsStandardVarietyOrder()
            Dim copyOfSort As New List(Of Variety)
            For Each var As Variety In Me.CurrentComparison.CurrentVarietySort
                copyOfSort.Add(var)
            Next
            Me.CurrentComparison.DefaultVarietySort = copyOfSort
        End Sub
        Public Sub SetCurrentComparisonAnalysisVarietyColumnIndex(ByVal index As Integer)
            Try
                Me.CurrentComparison.AssociatedAnalysis.CurrentVarietyColumnIndex = index
            Catch ex As Exception
            End Try
        End Sub
        Public Sub SetCurrentComparisonAnalysisVarietyIndex(ByVal index As Integer)
            Try
                Me.CurrentComparison.AssociatedAnalysis.CurrentVariety = Me.CurrentComparison.CurrentVarietySort(index)
            Catch ex As Exception
            End Try
        End Sub
        Public Sub SetCurrentComparisonsCurrentGloss(ByVal index As Integer)
            Try
                Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentGloss = Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses(index)
            Catch ex As Exception
            End Try
        End Sub
        Public Sub SetCurrentComparisonsCurrentVariety(ByVal index As Integer)
            Try
                Me.CurrentComparison.CurrentVariety = Me.CurrentComparison.CurrentVarietySort(index)
            Catch ex As Exception
            End Try
        End Sub
        Public Sub SetCurrentComparisonsAssociatedSurveysCurrentVariety(ByVal index As Integer)
            Try
                Me.CurrentComparison.AssociatedSurvey.CurrentVariety = Me.CurrentComparison.CurrentVarietySort(index)
            Catch ex As Exception
            End Try
        End Sub
        Public Sub SetCurrentComparisonsCurrentVarietyColumnIndex(ByVal index As Integer)
            Try
                Me.CurrentComparison.CurrentVarietyColumnIndex = index
            Catch ex As Exception
            End Try
        End Sub
        Public Sub SetCurrentSurvey(ByVal index As Integer)
            Try
                Me.CurrentSurvey = Me.Surveys(index)
            Catch ex As Exception
            End Try
        End Sub
        Public Sub SetCurrentSurveysCurrentVariety(ByVal index As Integer)
            Try
                Me.CurrentSurvey.CurrentVariety = Me.CurrentSurvey.Varieties(index)
            Catch ex As Exception
            End Try
        End Sub
        Public Sub SetCurrentSurveysCurrentVarietyEntryColumnIndex(ByVal index As Integer)
            Try
                Me.CurrentSurvey.CurrentVarietyEntryColumnIndex = index
            Catch ex As Exception
            End Try
        End Sub
        Public Sub SetCurrentComparisonsDDExcludedChars(ByVal chars As String)
            'We cannot allow spaces in the excluded chars because spaces are used to pad the words, and without padding everything would break when words are of different lengths.
            Me.CurrentComparison.AssociatedDegreesOfDifference.ExcludedChars = chars.Replace(" ", "")
        End Sub
        Public Sub SetCurrentSurveysCurrentSort(ByVal index As Integer)
            Try
                Me.CurrentSurvey.AssociatedDictionary.CurrentSort = Me.CurrentSurvey.AssociatedDictionary.Sorts(index)
            Catch ex As Exception
            End Try
        End Sub
        Public Sub SetCurrentSurveyDescription(ByVal val As String)
            Try
                Me.CurrentSurvey.Description = val
            Catch ex As Exception
            End Try
        End Sub
        Public Sub SetCurrentSurveysCurrentVarietyDescription(ByVal val As String)
            Try
                Me.CurrentSurvey.CurrentVariety.Description = val
            Catch ex As Exception
            End Try
        End Sub
        Public Sub SetCurrentSurveysCurrentGloss(ByVal rowIndex As Integer)
            Try
                Me.CurrentSurvey.AssociatedDictionary.CurrentGloss = Me.CurrentSurvey.AssociatedDictionary.CurrentSort.Glosses(rowIndex)
            Catch ex As Exception
            End Try
        End Sub
        Public Sub SetCurrentDictionarysCurrentGlossColumnIndex(ByVal index As Integer)
            Try
                Me.CurrentDictionary.CurrentGlossColumnIndex = index
            Catch ex As Exception
            End Try
        End Sub
        Public Sub SetExcludeValueForAllVarietiesForCurrentGloss(ByVal val As String)
            Dim currentGloss As Gloss = Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentGloss
            For Each var As Variety In Me.CurrentComparison.CurrentVarietySort
                Me.CurrentComparison.ComparisonEntries(var.VarietyEntries(currentGloss)).Exclude = val
            Next
        End Sub
        Public Sub SetCurrentComparisonsCurrentCOMPASSChar1Index(ByVal index As Integer)
            Me.CurrentComparison.COMPASSCalculations.CurrentChar1Index = index
        End Sub
        Public Sub SetCurrentComparisonsCurrentCOMPASSChar2Index(ByVal index As Integer)
            Me.CurrentComparison.COMPASSCalculations.CurrentChar2Index = index
        End Sub
        Public Sub SetCurrentDictionary(ByVal dictIndex As Integer)
            Try
                Me.CurrentDictionary = Me.Dictionaries(dictIndex)
            Catch ex As Exception
            End Try
        End Sub
        Public Sub SetCurrentDictionarysCurrentSort(ByVal sortIndex As Integer)
            Try
                Me.CurrentDictionary.CurrentSort = Me.CurrentDictionary.Sorts(sortIndex)
            Catch ex As Exception
            End Try
        End Sub
        Public Sub SetCurrentDictionarysCurrentGloss(ByVal glossIndex As Integer)
            Try
                Me.CurrentDictionary.CurrentGloss = Me.CurrentDictionary.CurrentSort.Glosses(glossIndex)
            Catch ex As Exception
            End Try
        End Sub
        Public Sub SetCurrentComparisonsCurrentCOMPASSVariety1(ByVal index As Integer)
            Try
                Me.CurrentComparison.COMPASSCalculations.CurrentVarietyIndex1 = index
            Catch ex As Exception
            End Try
        End Sub
        Public Sub SetCurrentComparisonsCurrentCOMPASSVariety2(ByVal index As Integer)
            Try
                Me.CurrentComparison.COMPASSCalculations.CurrentVarietyIndex2 = index
            Catch ex As Exception
            End Try
        End Sub


        Public Sub UpdateDDValue(ByVal rowIndex As Integer, ByVal colIndex As Integer, ByVal val As String)
            Dim UsedChars As List(Of String) = Me.CurrentComparison.AssociatedDegreesOfDifference.UsedCharsList
            Try
                Dim intVal As Integer = Integer.Parse(val)
                If intVal < 0 Then Return
            Catch ex As Exception
            End Try
            Try
                Me.CurrentComparison.AssociatedDegreesOfDifference.DDs((AscW(UsedChars(rowIndex)) << 16) Or AscW(UsedChars(colIndex))) = Integer.Parse(val)
                Me.CurrentComparison.AssociatedDegreesOfDifference.DDs((AscW(UsedChars(colIndex)) << 16) Or AscW(UsedChars(rowIndex))) = Integer.Parse(val)
            Catch ex As Exception
            End Try
        End Sub
        Public Sub CalculateCOMPASSValues(ByRef prefs As Preferences, ByVal varietyIndex1 As Integer, ByVal varietyIndex2 As Integer, ByVal upper As Integer, ByVal lower As Integer, ByVal bottom As Integer)
            Dim variety As String = ""
            Dim gloss As String = ""
            Dim var1 As Variety
            Dim var2 As Variety
            'Dim resultMSG As MsgBoxResult
            Dim skipMSGFlag As Boolean = False
            Try
                var1 = Me.CurrentComparison.CurrentVarietySort(varietyIndex1)
                var2 = Me.CurrentComparison.CurrentVarietySort(varietyIndex2)
            Catch ex As Exception
                Return
            End Try

            Dim calc As New COMPASSCalculation
            calc.CurrentVarietyIndex1 = varietyIndex1
            calc.CurrentVarietyIndex2 = varietyIndex2

            'For each gloss in the comparison
            '   If the two varieties have the same group and both are not excluded
            '       For each character in the Aligned Renderings
            '           If either character is not in our list of used characters, add it
            '           Add the glossID to the hash of occurences
            For Each gl As Gloss In Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses

                Dim compEntry1 As ComparisonEntry = Me.CurrentComparison.ComparisonEntries(var1.VarietyEntries(gl))
                Dim compEntry2 As ComparisonEntry = Me.CurrentComparison.ComparisonEntries(var2.VarietyEntries(gl))

                'If Not GroupsMatch(compEntry1.Grouping, compEntry2.Grouping) Then Continue For
                If compEntry1.Exclude <> "" Or compEntry2.Exclude <> "" Then Continue For 'DO NOT INCLUDE THESE CHARACTERS IN DoD!



                'ajw
                'Truncate the longer set of groupings commas, e.g.  a b,c,d,e,f g,i vs a,b,c,d e would truncate the former to just 5 groups
                Dim g1 As String = compEntry1.Grouping 'a,b
                Dim g2 As String = compEntry2.Grouping 'a
                Dim s1 As String = compEntry1.AlignedRendering
                Dim s2 As String = compEntry2.AlignedRendering

                Dim TESTsynonymGroupings1 As String() = Split(g1, ",") 'arm, irm
                Dim TESTsynonymGroupings2 As String() = Split(g2, ",") 'urm
                Dim TESTsynonyms1 As String() = Split(s1, ",")
                Dim TESTsynonyms2 As String() = Split(s2, ",")

                'If TESTsynonymGroupings1.Length <> TESTsynonyms1.Length Then
                '    If var1.Name <> Variety Or gl.Name <> Gloss Then '(To prevent repeat messages for each var x var comparison)
                '        If compEntry1.Exclude = "" And compEntry2.Exclude = "" Then
                '            If Not skipMSGFlag Then
                '                'MsgBox("Gloss '" & gl.Name & "' for " & var1.Name & " has a mismatched number of comma separated items in the aligned field versus the groupings field, which may result in incorrectly populated Degrees of Difference, Phonostatistical Analysis, and COMPASS grids!")
                '                resultMSG = MsgBox("Gloss '" & gl.Name & "' for " & var1.Name & " has a mismatched number of comma separated items in the aligned field versus the groupings field, which may result in incorrectly populated Degrees of Difference, Phonostatistical Analysis, and COMPASS grids!" & vbCrLf & vbCrLf & "Would you like to skip the warning message for the rest of the glosses?", MsgBoxStyle.YesNo, "Mismatched number of gloss synonyms and groupings")
                '                If resultMSG = MsgBoxResult.Yes Then
                '                    skipMSGFlag = True
                '                End If
                '            End If
                '        End If
                '    End If
                variety = var1.Name
                gloss = gl.Name
                'End If

                truncGroupings(g1, s1)
                truncGroupings(g2, s2)
                Dim synonymGroupings1 As String() = Split(g1, ",") 'arm, irm
                Dim synonymGroupings2 As String() = Split(g2, ",") 'urm
                Dim synonyms1 As String() = Split(s1, ",")
                Dim synonyms2 As String() = Split(s2, ",")
                'AJW



                For synIndex1 As Integer = 0 To synonymGroupings1.Length - 1 'synonyms1.Length - 1 'AJW***
                    For synIndex2 As Integer = 0 To synonymGroupings2.Length - 1 'synonyms2.Length - 1 'AJW***
                        If GroupsMatch(synonymGroupings1(synIndex1), synonymGroupings2(synIndex2)) Then
                            'Pad the words with spaces so they are the same length.
                            'Dim word1 As String = compEntry1.AlignedRendering
                            'Dim word2 As String = compEntry2.AlignedRendering
                            Dim word1 As String = synonyms1(synIndex1)
                            Dim word2 As String = synonyms2(synIndex2)

                            For Each ch As String In Me.CurrentComparison.AssociatedDegreesOfDifference.ExcludedChars
                                If word1.Contains(ch) Then word1 = word1.Replace(ch, "")
                                If word2.Contains(ch) Then word2 = word2.Replace(ch, "")
                            Next
                            PadStringsToLongest(word1, word2)

                            'Store these values for the other calculations.
                            Dim calcdValues As New COMPASSGlossEntry
                            calcdValues.gl = gl
                            calcdValues.Form = gl.Name
                            calcdValues.PaddedForm1 = word1
                            calcdValues.PaddedForm2 = word2
                            calcdValues.AverageStrength = 0.0
                            calc.GlossValues.Add(calcdValues)

                            'Go character by character.
                            For i As Integer = 0 To word1.Length - 1

                                'Make a list of the unique characters, storing strings of length one for simplicity.
                                'If Not calc.UsedChars.Contains(word1(i).ToString) Then calc.UsedChars.Add(word1(i).ToString.ToLower())
                                'If Not calc.UsedChars.Contains(word2(i).ToString) Then calc.UsedChars.Add(word2(i).ToString.ToLower())
                                If Not calc.UsedChars.Contains(word1(i).ToString) Then calc.UsedChars.Add(word1(i).ToString.ToLower())
                                If Not calc.UsedChars.Contains(word2(i).ToString) Then calc.UsedChars.Add(word2(i).ToString.ToLower())

                                Dim hashKey As String = word1(i) & word2(i) 'The key is the letter pair.

                                'If we haven't seen this pair before, add a new entry in the hash table.
                                If Not calc.CharPairRecords.ContainsKey(hashKey) Then calc.CharPairRecords.Add(hashKey, New COMPASSCalcEntry)

                                'Make a list of all the glosses that contain this character pair.
                                calc.CharPairRecords(hashKey).Occurences.Add(gl)
                                If Not calc.UsedGlosses.ContainsKey(hashKey) Then
                                    calc.UsedGlosses.Add(hashKey, New Dictionary(Of Gloss, Integer))
                                End If
                                If Not calc.UsedGlosses(hashKey).ContainsKey(gl) Then
                                    calc.UsedGlosses(hashKey).Add(gl, 1)
                                Else
                                    calc.UsedGlosses(hashKey)(gl) += 1
                                End If
                            Next
                        End If
                    Next
                Next
            Next


            'AJW*** Here is where the repetitive code for each synonym ends (new addition, previously was treated as one long string)



            calc.UsedChars.Sort() 'Alphabetically

            'Calculate the strengths.
            For Each rowChar As String In calc.UsedChars
                For Each colChar As String In calc.UsedChars
                    'If this letter pair is in our table,
                    If calc.CharPairRecords.ContainsKey(rowChar & colChar) Then
                        'Calculate the strength and store it in the hash for future use.
                        Dim strength As Double = CalculateCharPairStrength(calc.CharPairRecords(rowChar & colChar).Occurences.Count, upper, lower, bottom)
                        calc.CharPairRecords(rowChar & colChar).Strength = strength
                    End If
                Next
            Next

            'Now do the calculations for the word average strengths.  This is easy since we already calculated most of the values.
            For Each glVal As COMPASSGlossEntry In calc.GlossValues
                Dim sumOfStrengths As Double = 0.0

                'Go character by character, adding up the strengths for this word and find the average.
                For i As Integer = 0 To glVal.PaddedForm1.Length - 1
                    sumOfStrengths += calc.CharPairRecords(glVal.PaddedForm1(i) & glVal.PaddedForm2(i)).Strength
                Next
                If glVal.PaddedForm1.Length > 0 Then
                    glVal.AverageStrength = sumOfStrengths / glVal.PaddedForm1.Length
                Else
                    glVal.AverageStrength = 0
                End If

                AddValueToStrengthCounts(calc.strengthCounts, glVal.AverageStrength)
            Next
            calc.GlossValues.Sort(New GlossValueSorter)

            If Me.CurrentComparison.COMPASSCalculations Is Nothing Then
                calc.CurrentVarietyIndex1 = prefs.COMPASSVariety1Index
                calc.CurrentVarietyIndex2 = prefs.COMPASSVariety2Index
            End If
            Me.CurrentComparison.COMPASSCalculations = calc
        End Sub
        Public Class GlossValueSorter
            Implements IComparer(Of COMPASSGlossEntry)

            Public Function Compare(ByVal x As COMPASSGlossEntry, ByVal y As COMPASSGlossEntry) As Integer Implements System.Collections.Generic.IComparer(Of WordSurv7.DataObjects.COMPASSGlossEntry).Compare
                Dim diff As Double = y.AverageStrength - x.AverageStrength
                If diff > 0.0 Then Return 1
                If diff < 0.0 Then Return -1
                Return 0
            End Function
        End Class
        Private Sub AddValueToStrengthCounts(ByRef strengthCounts As StrengthCountsSummary, ByVal strength As Double)
            If strength = 1.0 Then
                strengthCounts.Eq1 += 1
            ElseIf strength >= 0.75 Then
                strengthCounts.Gte75lt1 += 1
            ElseIf strength >= 0.5 Then
                strengthCounts.Gte50lt75 += 1
            ElseIf strength >= 0.25 Then
                strengthCounts.Gte25lt50 += 1
            ElseIf strength >= 0.0 Then
                strengthCounts.Gte0lt25 += 1
            ElseIf strength >= -0.25 Then
                strengthCounts.Gten25lt0 += 1
            ElseIf strength >= -0.5 Then
                strengthCounts.Gten50ltn25 += 1
            ElseIf strength >= -0.75 Then
                strengthCounts.Gten75ltn50 += 1
            ElseIf strength > -1.0 Then
                strengthCounts.Gtn1ltn75 += 1
            Else
                strengthCounts.Eqn1 += 1
            End If
        End Sub
        Private Function CalculateCharPairStrength(ByVal amt As Integer, ByVal upper As Integer, ByVal lower As Integer, ByVal bottom As Integer) As Double
            'Calculation taken from the WordSurv 2.5 manual.
            'n > t       -> 1.0
            't >= n > l  -> n / t
            'l >= n > b  -> -0.5
            'b >= n      -> -1.0
            'n = number of occurences, t = upper threshold, l = lower threshold, b = bottom threshold

            If amt > upper Then Return 1.0
            If upper >= amt And amt > lower Then Return (amt * 1.0) / upper
            If lower >= amt And amt > bottom Then Return -0.5
            If bottom >= amt Then Return -1.0
        End Function
        Public Sub FilterCOMPASSStrengthsGrid(ByVal selectedCoords As List(Of CellAddress))
            Dim usedChars As List(Of String) = Me.CurrentComparison.COMPASSCalculations.UsedChars
            Me.CurrentComparison.COMPASSCalculations.DisplayedGlosses.Clear()
            Dim setOfEntries As New Dictionary(Of COMPASSGlossEntry, Integer)

            'Iterate over each cell in the selection and add the glosses for that phone pair to the UsedGlosses list
            For Each addr As CellAddress In selectedCoords

                'The list of all the glosses between the two COMPASS varieties
                Dim usedGlosses As List(Of COMPASSGlossEntry) = Me.CurrentComparison.COMPASSCalculations.GlossValues

                'A set of all the gloss objects that use this phone pair
                Dim glossesUsingThisPhonePair As Dictionary(Of Gloss, Integer) = Me.CurrentComparison.COMPASSCalculations.UsedGlosses(usedChars(addr.RowIndex) & usedChars(addr.ColIndex))
                'A set of the names in string form or those glosses.
                Dim glossNamesUsingThisPhonePair As New Dictionary(Of String, Integer)
                For Each gl As Gloss In glossesUsingThisPhonePair.Keys
                    glossNamesUsingThisPhonePair.Add(gl.Name, 1)
                Next

                'Iterate over the list of all glosses used between the two varieties
                For Each glEntry As COMPASSGlossEntry In usedGlosses
                    'Add any unseen glosses into the list which holds the glosses that are displayed in the right pane.
                    If glossNamesUsingThisPhonePair.ContainsKey(glEntry.Form) Then
                        If Not setOfEntries.ContainsKey(glEntry) Then
                            setOfEntries.Add(glEntry, Nothing)
                        End If
                    End If
                Next
            Next

            'If there was nothing selected, display everything.
            If setOfEntries.Keys.Count = 0 Then
                For Each glEntry As COMPASSGlossEntry In Me.CurrentComparison.COMPASSCalculations.GlossValues
                    setOfEntries.Add(glEntry, Nothing)
                Next
            End If

            Dim sortedEntries As New List(Of COMPASSGlossEntry)
            For Each k As COMPASSGlossEntry In setOfEntries.Keys
                sortedEntries.Add(k)
            Next
            sortedEntries.Sort(New GlossValueSorter)
            Me.CurrentComparison.COMPASSCalculations.DisplayedGlosses.AddRange(sortedEntries)
        End Sub
        Public Function Copy() As WordSurvData

            'Every time the user performs an operation, this function copies the entire data pile and puts that saved state into the undo buffer.
            Dim wsCopy As New WordSurvData
            wsCopy.filename = Me.filename
            wsCopy.CurrentTab = Me.CurrentTab
            wsCopy.PrimaryFont = Me.PrimaryFont
            wsCopy.SecondaryFont = Me.SecondaryFont
            wsCopy.TranscriptionFont = Me.TranscriptionFont
            wsCopy.PrimaryLanguage = Me.PrimaryLanguage
            wsCopy.SecondaryLanguage = Me.SecondaryLanguage

            'This function copies the linked structure of the WordSurvData object.  Because of this we need temporary mappings from the original
            'data structure's objects to the new data structure's objects.
            Dim oldToNewDict As New Dictionary(Of Dictionary, Dictionary) 'Given a dictionary from the old structure, return a new dictionary
            Dim oldToNewSurv As New Dictionary(Of Survey, Survey)         'Given a survey from the old structure, return a new survey

            Dim oldGlossLists As New Dictionary(Of Dictionary, List(Of Gloss)) 'Given a dictionary, give a list of glosses belonging to it
            Dim newGlossLists As New Dictionary(Of Dictionary, List(Of Gloss))

            Dim oldSurvVarTransLists As New List(Of List(Of List(Of VarietyEntry))) 'Survey, Variety, Transcription
            Dim newSurvVarTransLists As New List(Of List(Of List(Of VarietyEntry)))

            For Each dict As Dictionary In Me.Dictionaries
                Dim oldToNewGlossMapping As New Dictionary(Of Gloss, Gloss) 'Used for sorts
                Dim oldGlossList As New List(Of Gloss)                      'Used elsewhere
                Dim newGlossList As New List(Of Gloss)

                For Each gl As Gloss In dict.Sorts(0).Glosses
                    Dim newGloss As New Gloss(gl.Name)
                    newGloss.Name2 = gl.Name2
                    newGloss.PartOfSpeech = gl.PartOfSpeech
                    newGloss.FieldTip = gl.FieldTip
                    newGloss.Comments = gl.Comments
                    oldToNewGlossMapping.Add(gl, newGloss)
                    oldGlossList.Add(gl)
                    newGlossList.Add(newGloss)
                Next
                oldGlossLists.Add(dict, oldGlossList)

                Dim newDictionary As New Dictionary(dict.Name)
                For Each srt As Sort In dict.Sorts

                    Dim newSort As New Sort(srt.Name)
                    For Each gl As Gloss In srt.Glosses
                        '                 __
                        '                / *_) ?
                        '     _.----. _ /../
                        '    /............/
                        ' __/...(..|.(../
                        '/__.-|_|--|_|

                        newSort.Glosses.Add(oldToNewGlossMapping(gl))
                    Next
                    newDictionary.Sorts.Add(newSort)
                Next
                newGlossLists.Add(newDictionary, newGlossList)

                Try
                    newDictionary.CurrentSort = newDictionary.Sorts(dict.Sorts.IndexOf(dict.CurrentSort))
                    newDictionary.CurrentGloss = newDictionary.CurrentSort.Glosses(dict.CurrentSort.Glosses.IndexOf(dict.CurrentGloss))
                    newDictionary.CurrentGlossColumnIndex = dict.CurrentGlossColumnIndex
                Catch ex As Exception
                End Try

                wsCopy.Dictionaries.Add(newDictionary)
                oldToNewDict.Add(dict, newDictionary)
            Next
            Try
                wsCopy.CurrentDictionary = wsCopy.Dictionaries(Me.GetCurrentDictionaryIndex())
            Catch ex As Exception
            End Try






            For Each surv As Survey In Me.Surveys
                Dim oldVarList As New List(Of List(Of VarietyEntry))
                Dim newVarList As New List(Of List(Of VarietyEntry))
                Dim newSurvey As New Survey(oldToNewDict(surv.AssociatedDictionary), surv.Name)
                newSurvey.Description = surv.Description
                newSurvey.CurrentVarietyEntryColumnIndex = newSurvey.CurrentVarietyEntryColumnIndex
                For Each var As Variety In surv.Varieties
                    'Massage the transcriptions into a usable form (make a list of transcriptions in gloss order)
                    Dim oldTransList As New List(Of VarietyEntry)
                    Dim newTransList As New List(Of VarietyEntry)
                    Dim oldGlossList As List(Of Gloss) = oldGlossLists(var.AssociatedDictionary)
                    Dim newGlossList As List(Of Gloss) = newGlossLists(newSurvey.AssociatedDictionary)

                    For Each gl As Gloss In oldGlossList
                        Dim oldTrans As VarietyEntry = var.VarietyEntries(gl)
                        oldTransList.Add(oldTrans)
                        newTransList.Add(oldTrans.Copy())
                    Next
                    oldVarList.Add(oldTransList)
                    newVarList.Add(newTransList)

                    'Using the newly massaged data, copy the transcriptions into the new variety
                    Dim newVar As New Variety(oldToNewDict(var.AssociatedDictionary), var.Name, False)
                    newVar.Description = var.Description
                    For i As Integer = 0 To oldTransList.Count - 1
                        newVar.VarietyEntries.Add(newGlossList(i), newTransList(i))
                    Next
                    newSurvey.Varieties.Add(newVar)


                    'Reverse engineer the current transcription
                    For i As Integer = 0 To oldTransList.Count - 1
                        If oldTransList(i) Is var.CurrentVarietyEntry Then
                            newVar.CurrentVarietyEntry = newTransList(i)
                        End If
                    Next

                Next

                Try
                    newSurvey.CurrentVariety = newSurvey.Varieties(surv.Varieties.IndexOf(surv.CurrentVariety))
                Catch ex As Exception
                End Try

                wsCopy.Surveys.Add(newSurvey)
                oldSurvVarTransLists.Add(oldVarList)
                newSurvVarTransLists.Add(newVarList)
                oldToNewSurv.Add(surv, newSurvey)
            Next
            Try
                wsCopy.CurrentSurvey = wsCopy.Surveys(Me.GetCurrentSurveyIndex())
            Catch ex As Exception
            End Try


            For Each comp As Comparison In Me.Comparisons
                Dim newComp As New Comparison(oldToNewSurv(comp.AssociatedSurvey), comp.Name, False)
                newComp.Description = comp.Description
                newComp.CurrentVarietyColumnIndex = comp.CurrentVarietyColumnIndex
                Dim survIndex As Integer = Me.Surveys.IndexOf(comp.AssociatedSurvey)
                For varIndex As Integer = 0 To newSurvVarTransLists(survIndex).Count - 1
                    For transIndex As Integer = 0 To newSurvVarTransLists(survIndex)(varIndex).Count - 1
                        newComp.ComparisonEntries.Add(newSurvVarTransLists(survIndex)(varIndex)(transIndex), comp.ComparisonEntries(oldSurvVarTransLists(survIndex)(varIndex)(transIndex)).Copy())
                    Next
                Next

                newComp.AssociatedSurvey = wsCopy.Surveys(survIndex)

                For Each var As Variety In comp.CurrentVarietySort
                    newComp.CurrentVarietySort.Add(newComp.AssociatedSurvey.Varieties(comp.AssociatedSurvey.Varieties.IndexOf(var)))
                Next
                For Each var As Variety In comp.DefaultVarietySort
                    newComp.DefaultVarietySort.Add(newComp.AssociatedSurvey.Varieties(comp.AssociatedSurvey.Varieties.IndexOf(var)))
                Next

                wsCopy.Comparisons.Add(newComp)

                Try
                    newComp.CurrentVariety = wsCopy.Surveys(survIndex).Varieties(Me.Surveys(survIndex).Varieties.IndexOf(comp.CurrentVariety))
                Catch ex As Exception
                End Try

                newComp.AssociatedAnalysis = New ComparisonAnalysis(newComp, wsCopy.Surveys(survIndex).Varieties)

                newComp.AssociatedDegreesOfDifference = New DegreesOfDifferenceGrid(newComp)

                If Not comp.AssociatedDegreesOfDifference.DDs Is Nothing Then
                    newComp.AssociatedDegreesOfDifference.DDs = New Dictionary(Of Integer, Integer)

                    'Copy the DD values and the used chars list
                    For Each kvp As KeyValuePair(Of Integer, Integer) In comp.AssociatedDegreesOfDifference.DDs
                        newComp.AssociatedDegreesOfDifference.DDs.Add(kvp.Key, kvp.Value)
                    Next
                    newComp.AssociatedDegreesOfDifference.UsedCharsList = New List(Of String)(comp.AssociatedDegreesOfDifference.UsedCharsList)

                    newComp.AssociatedDegreesOfDifference.ExcludedChars = comp.AssociatedDegreesOfDifference.ExcludedChars
                End If
            Next
            Try
                wsCopy.CurrentComparison = wsCopy.Comparisons(Me.GetCurrentComparisonIndex())
            Catch ex As Exception
            End Try


            Return wsCopy
        End Function
        Public Shared Function LoadFile(ByVal filename As String) As WordSurvData
            LoadInterrupted = True
            Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            Dim reader As IO.StreamReader = Nothing
            Dim lineCount As Integer = 0
            Dim tempLine As String
            Dim newData As WordSurvData
            newData = New WordSurvData
            Dim transMaps As New Dictionary(Of Variety, Dictionary(Of Integer, VarietyEntry))
            reader = New IO.StreamReader(filename, System.Text.Encoding.UTF8)
            While reader.Peek <> -1
                Dim line As String = reader.ReadLine().Trim()
                lineCount += 1
                Dim parts As String() = line.Split(" "c)
                'If this is the start of a block
                Select Case line
                    Case "Start Dictionary"
                        Dim newDict As New Dictionary(Split(reader.ReadLine().Trim().Replace("♠"c, "|"c), "=", 2)(1))
                        lineCount += 1
                        Dim glosses As New List(Of Gloss)

                        'Create Gloss objects for every gloss in the Dictionary
                        reader.ReadLine() 'Consume Start Glosses
                        lineCount += 1
                        line = reader.ReadLine().Trim()
                        lineCount += 1
                        While line <> "End Glosses"
                            line = line.Replace("♠"c, "|"c)
                            parts = Split(line, "|")
                            Dim newGloss As New Gloss
                            For i As Integer = 0 To parts.Length - 1 'Read in the glosses's columns
                                newGloss.SetByIndex(i, parts(i).Replace("♠"c, "|"c))
                            Next
                            glosses.Add(newGloss)
                            line = reader.ReadLine().Trim()
                            lineCount += 1
                        End While

                        'Create Sort objects for every sort in the Dictionary
                        reader.ReadLine() 'Consume Start Sorts
                        lineCount += 1
                        line = reader.ReadLine().Trim()
                        lineCount += 1
                        While line <> "End Sorts"
                            line = line.Replace("♠"c, "|"c)
                            parts = Split(line, "|")
                            Dim newSort As New Sort(parts(0).Replace("♠"c, "|"c))
                            For i As Integer = 1 To parts.Length - 1 'Read in the Gloss id's for this Sort
                                newSort.Glosses.Add(glosses(Integer.Parse(parts(i))))
                            Next
                            newDict.Sorts.Add(newSort)
                            line = reader.ReadLine().Trim()
                            lineCount += 1
                        End While

                        'Set the current sort and gloss of the dictionary
                        Try
                            newDict.CurrentSort = newDict.Sorts(Integer.Parse(Split(reader.ReadLine().Trim(), "=")(1)))
                            lineCount += 1
                            tempLine = reader.ReadLine().Trim()
                            lineCount += 1
                            newDict.CurrentGloss = newDict.CurrentSort.Glosses(Integer.Parse(Split(tempLine, "=")(1)))
                            tempLine = reader.ReadLine().Trim()
                            lineCount += 1
                            newDict.CurrentGlossColumnIndex = Integer.Parse(Split(tempLine, "=")(1))
                        Catch ex As Exception
                            MsgBox("Dictionary creation partial failure!!!!")
                        End Try

                        newData.Dictionaries.Add(newDict)

                    Case "Start Survey"
                        Dim newSurv As New Survey(newData.Dictionaries(Integer.Parse(Split(reader.ReadLine().Trim(), "=", 2)(1))), _
                                                  Split(reader.ReadLine().Trim().Replace("♠"c, "|"c), "=", 2)(1))
                        lineCount += 1
                        lineCount += 1

                        'Create Varitey objects for every variety in the Survey
                        reader.ReadLine() 'Consume Start Varieties
                        lineCount += 1
                        line = reader.ReadLine().Trim() 'Consume Start Variety
                        lineCount += 1
                        Dim varCountTemp As Integer = 0
                        While line <> "End Varieties"
                            Dim assocDictID As Integer = Integer.Parse(Split(reader.ReadLine().Trim(), "=", 2)(1))
                            lineCount += 1
                            Dim newVar As New Variety(newData.Dictionaries(assocDictID), Split(reader.ReadLine().Trim().Replace("♠"c, "|"c), "=", 2)(1), True)
                            lineCount += 1
                            Dim transMap As New Dictionary(Of Integer, VarietyEntry)

                            'Create Transcription objects for every transcription in the Variety
                            reader.ReadLine() 'Consume Start Transcriptions
                            lineCount += 1
                            line = reader.ReadLine().Trim()
                            lineCount += 1
                            While line <> "End Transcriptions"
                                line = line.Replace("♠"c, "|"c) 'AJW***
                                parts = Split(line, "|"c) 'AJW***
                                'Debug.Print(line)
                                'Debug.Print(parts(0))
                                Dim varEntry As VarietyEntry = newVar.VarietyEntries(newVar.AssociatedDictionary.CurrentSort.Glosses(Integer.Parse(parts(0))))
                                transMap.Add(Integer.Parse(parts(1)), varEntry)
                                varEntry.Transcription = parts(2).Replace("♠"c, "|"c)
                                varEntry.PluralFrame = parts(3).Replace("♠"c, "|"c)
                                varEntry.Notes = parts(4).Replace("♠"c, "|"c)
                                line = reader.ReadLine().Trim()
                                lineCount += 1
                            End While
                            varCountTemp += 1
                            'If varCountTemp = 13 Then MsgBox(varCountTemp)
                            'Debug.Print("Finished Variety " & varCountTemp.ToString)
                            Try
                                newVar.CurrentVarietyEntry = newVar.VarietyEntries(newVar.AssociatedDictionary.CurrentSort.Glosses(Integer.Parse(Split(reader.ReadLine().Trim(), "=")(1))))
                                lineCount += 1
                            Catch ex As Exception
                                newVar.CurrentVarietyEntry = Nothing
                            End Try
                            newVar.Description = Split(reader.ReadLine().Trim().Replace("\", vbCrLf).Replace("♠"c, "|"c), "=")(1)
                            lineCount += 1

                            newSurv.Varieties.Add(newVar)
                            transMaps.Add(newVar, transMap)

                            reader.ReadLine().Trim() 'Consume End Variety
                            lineCount += 1
                            line = reader.ReadLine().Trim()
                            lineCount += 1
                        End While

                        Try
                            newSurv.CurrentVariety = newSurv.Varieties(Integer.Parse(Split(reader.ReadLine().Trim(), "=")(1)))
                            lineCount += 1
                        Catch ex As Exception
                        End Try
                        newSurv.Description = Split(reader.ReadLine().Trim().Replace("\", vbCrLf).Replace("♠"c, "|"c), "=")(1)
                        lineCount += 1
                        newSurv.CurrentVarietyEntryColumnIndex = Integer.Parse(Split(reader.ReadLine().Trim().Replace("\", vbCrLf), "=")(1))
                        lineCount += 1
                        newData.Surveys.Add(newSurv)
                        reader.ReadLine() 'Consume the End Survey
                        lineCount += 1
                    Case "Start Comparison"
                        Dim newComp As New Comparison(newData.Surveys(Integer.Parse(Split(reader.ReadLine().Trim(), "=", 2)(1))), _
                                                      Split(reader.ReadLine().Trim().Replace("♠"c, "|"c), "=", 2)(1), False)
                        lineCount += 1
                        lineCount += 1

                        Dim varParts As String() = Split(reader.ReadLine().Trim(), "=")
                        lineCount += 1
                        Dim varIDs As String() = Split(varParts(1), "|")
                        newComp.CurrentVarietySort = New List(Of Variety)
                        Try
                            For Each varID As Integer In varIDs 'AJW
                                newComp.CurrentVarietySort.Add(newComp.AssociatedSurvey.Varieties(varID))
                                newComp.DefaultVarietySort.Add(newComp.AssociatedSurvey.Varieties(varID))
                            Next
                        Catch ex As Exception
                        End Try

                        'Create Comparison Entry objects for every Comparison Entry in the Dictionary
                        Dim numVarieties As Integer = newComp.AssociatedSurvey.Varieties.Count
                        Dim numGlosses As Integer = newComp.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses.Count
                        Dim curVariety As Integer = 0
                        Dim curGloss As Integer = 0

                        reader.ReadLine() 'Consume Start Comparison Entries
                        lineCount += 1
                        line = reader.ReadLine().Trim()
                        lineCount += 1
                        Try 'AJW
                            While line <> "End Comparison Entries"
                                'Debug.Print(line)
                                line = line.Replace("♠"c, "|"c)
                                parts = Split(line, "|")
                                Dim newCompEntry As New ComparisonEntry(parts(1).Replace("♠"c, "|"c))
                                newCompEntry.Grouping = parts(2).Replace("♠"c, "|"c)
                                newCompEntry.Notes = parts(3).Replace("♠"c, "|"c)
                                newCompEntry.Exclude = parts(4).Replace("♠"c, "|"c)
                                Try
                                    newComp.ComparisonEntries.Add(transMaps(newComp.AssociatedSurvey.Varieties(curVariety))(curGloss), newCompEntry)
                                Catch ex As Exception
                                    MsgBox("Ill-formed .wsv file, probably excess comparison entries (> dictionary length), line 2414 in code, gloss count " & curGloss.ToString & " in variety count " & curVariety.ToString)
                                    MsgBox(ex.Message & vbCrLf & ex.Data.ToString & vbCrLf & ex.GetBaseException.ToString & vbCrLf & ex.GetType.ToString & vbCrLf & ex.HelpLink.ToString & vbCrLf & ex.Source.ToString)
                                End Try
                                line = reader.ReadLine().Trim()
                                lineCount += 1
                                curGloss += 1
                                If curGloss > numGlosses - 1 Then
                                    curGloss = 0
                                    curVariety += 1
                                End If
                            End While
                        Catch bup As Exception 'AJW
                            MsgBox(bup.Message & "   -   " & bup.Source)
                        End Try 'AJW

                        Try
                            newComp.CurrentVariety = newComp.AssociatedSurvey.Varieties(Integer.Parse(Split(reader.ReadLine().Trim(), "=")(1)))
                            lineCount += 1
                        Catch ex As Exception
                        End Try
                        newComp.Description = Split(reader.ReadLine().Trim().Replace("\", vbCrLf).Replace("♠"c, "|"c), "=")(1)
                        lineCount += 1
                        Try
                            newComp.StartDate = Date.Parse(Split(reader.ReadLine().Trim(), "=")(1))
                            lineCount += 1
                        Catch ex As Exception
                        End Try
                        Try
                            newComp.EndDate = Date.Parse(Split(reader.ReadLine().Trim(), "=")(1))
                            lineCount += 1
                        Catch ex As Exception
                        End Try
                        newComp.CurrentVarietyColumnIndex = Integer.Parse(Split(reader.ReadLine().Trim(), "=")(1))
                        lineCount += 1

                        'Make a new DD grid for this comparison
                        newComp.AssociatedDegreesOfDifference = New DegreesOfDifferenceGrid(newComp)
                        'Read in the used characters
                        newComp.AssociatedDegreesOfDifference.UsedCharsList = New List(Of String)
                        Dim xline As String = reader.ReadLine()
                        lineCount += 1
                        xline = xline.Replace("♠"c, "|"c)
                        'THIS IS THE DDUSEDCHARS LINE WHERE AN = SYMBOL CAUSES PROBLEMS 'AJW 2012-12-15
                        Dim chrs As String() = Split(xline.Trim(vbCrLf(0)), "=", 2)(1).Split("|"c)
                        For i As Integer = 0 To chrs.Length - 1
                            chrs(i) = chrs(i).Replace("♠"c, "|"c)
                        Next

                        reader.ReadLine() 'Consume Start DD Values
                        lineCount += 1

                        If chrs(0) <> "" Then 'If there are used chars and therefore a DD grid
                            newComp.AssociatedDegreesOfDifference.UsedCharsList.AddRange(chrs)

                            newComp.AssociatedDegreesOfDifference.DDs = New Dictionary(Of Integer, Integer)

                            For Each usedChar As String In newComp.AssociatedDegreesOfDifference.UsedCharsList
                                line = reader.ReadLine().Trim()
                                lineCount += 1
                                line = line.Replace("♠"c, "|"c)
                                Dim ddVals As String() = Split(line, "|")
                                Dim cnt As Integer = 0
                                For Each usedChar2 As String In newComp.AssociatedDegreesOfDifference.UsedCharsList
                                    Dim char1AndChar2 As Integer = (AscW(usedChar) << 16) Or AscW(usedChar2)
                                    newComp.AssociatedDegreesOfDifference.DDs.Add(char1AndChar2, Integer.Parse(ddVals(cnt)))
                                    cnt += 1
                                Next
                            Next
                        End If
                        Dim junkline As String = reader.ReadLine() 'Consume the End DD Values
                        lineCount += 1
                        'Read in the excluded characters
                        junkline = reader.ReadLine()
                        lineCount += 1
                        newComp.AssociatedDegreesOfDifference.ExcludedChars = Split(junkline.Trim().Replace("♠"c, "|"c), "=")(1)

                        reader.ReadLine() 'Consume the End Comparison
                        lineCount += 1
                        newData.Comparisons.Add(newComp)

                    Case "Start WordSurv Data"
                        'Read in the current dictionary, survey, and comparison
                        Try
                            newData.CurrentDictionary = newData.Dictionaries(Integer.Parse(Split(reader.ReadLine().Trim(), "=")(1)))
                            lineCount += 1
                        Catch ex As Exception
                            newData.CurrentDictionary = Nothing
                        End Try
                        Try
                            newData.CurrentSurvey = newData.Surveys(Integer.Parse(Split(reader.ReadLine().Trim(), "=")(1)))
                            lineCount += 1
                        Catch ex As Exception
                            newData.CurrentSurvey = Nothing
                        End Try
                        Try
                            newData.CurrentComparison = newData.Comparisons(Integer.Parse(Split(reader.ReadLine().Trim(), "=")(1)))
                            lineCount += 1
                        Catch ex As Exception
                            newData.CurrentComparison = Nothing
                        End Try

                        Try
                            newData.PrimaryLanguage = Split(reader.ReadLine().Trim().Replace("♠"c, "|"c), "=")(1)
                            lineCount += 1
                        Catch ex As Exception
                            newData.PrimaryLanguage = "Primary Gloss"
                        End Try
                        Try
                            newData.SecondaryLanguage = Split(reader.ReadLine().Trim().Replace("♠"c, "|"c), "=")(1)
                            lineCount += 1
                        Catch ex As Exception
                            newData.SecondaryLanguage = "Secondary Gloss"
                        End Try

                        Try
                            Dim primaryFontData As String() = Split(Split(reader.ReadLine().Trim(), "=")(1), ",")
                            lineCount += 1
                            newData.PrimaryFont = New Font(primaryFontData(0), Single.Parse(primaryFontData(1)))
                        Catch ex As Exception
                            newData.PrimaryFont = New Font("Microsoft Sans Serif", 8)
                        End Try
                        Try
                            Dim secondaryFontData As String() = Split(Split(reader.ReadLine().Trim(), "=")(1), ",")
                            lineCount += 1
                            newData.SecondaryFont = New Font(secondaryFontData(0), Single.Parse(secondaryFontData(1)))
                        Catch ex As Exception
                            newData.SecondaryFont = New Font("Microsoft Sans Serif", 8)
                        End Try
                        Try
                            Dim transcriptionFontData As String() = Split(Split(reader.ReadLine().Trim(), "=")(1), ",")
                            lineCount += 1
                            newData.TranscriptionFont = New Font(transcriptionFontData(0), Single.Parse(transcriptionFontData(1)))
                        Catch ex As Exception
                            newData.TranscriptionFont = New Font("Microsoft Sans Serif", 8)
                        End Try
                    Case Else 'AJW*** to accomodate first two lines of .wsv file containing 'WordSurv version X' and 'Beta release Kemuel'
                        If line.Contains("WordSurv version ") Then
                            Dim junklineD As String = reader.ReadLine() 'consume the detailed release line (2nd line, e.g. Beta release Kemuel)
                            lineCount += 1
                        End If
                End Select
            End While
            If reader IsNot Nothing Then reader.Close()
            Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            LoadInterrupted = False
            Return newData
        End Function
        Public Sub MergeCurrentDatabaseWithThisOne(ByVal filename As String)
            Dim otherData As WordSurvData = WordSurvData.LoadFile(filename)

            For Each mergeDict As Dictionary In otherData.Dictionaries
                For Each oldDict As Dictionary In Me.Dictionaries
                    If mergeDict.Name = oldDict.Name Then
                        mergeDict.Name &= " from Import"
                    End If
                Next
            Next
            For Each mergeSurv As Survey In otherData.Surveys
                For Each oldSurv As Survey In Me.Surveys
                    If mergeSurv.Name = oldSurv.Name Then
                        mergeSurv.Name &= " from Import"
                    End If
                Next
            Next
            For Each mergeComp As Comparison In otherData.Comparisons
                For Each oldComp As Comparison In Me.Comparisons
                    If mergeComp.Name = oldComp.Name Then
                        mergeComp.Name &= " from Import"
                    End If
                Next
            Next

            Me.Dictionaries.AddRange(otherData.Dictionaries)
            Me.Surveys.AddRange(otherData.Surveys)
            Me.Comparisons.AddRange(otherData.Comparisons)
        End Sub
        Public Shared Function badCharClean(ByVal line As String) As String
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
            Return line
        End Function
        Public Shared Sub ImportWordSurv2_5Database(ByVal dbFileNameBase As String)
            Dim inString As String
            Dim langSymbol As String
            Dim VarietyName As String
            Dim counter As Integer
            Dim comparisonGlossList As New List(Of String)
            Dim cognateSymbol As String = ""
            Dim varietyCount As Integer

            Dim frmEncodingChoice As New EncodingChoice(dbFileNameBase)
            If frmEncodingChoice.ShowDialog() = DialogResult.Cancel Then
                Throw New System.OperationCanceledException
            End If
            Dim encoding2_5 As String = frmEncodingChoice.cboEncodingSelector.SelectedItem.ToString

            Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            Dim dbLinesTEST As New List(Of String)
            Dim inputDBFileTEST As New IO.StreamReader(dbFileNameBase & ".db", System.Text.Encoding.GetEncoding(encoding2_5))
            While inputDBFileTEST.Peek <> -1
                inString = inputDBFileTEST.ReadLine
                'WideCharToMultiByte(1252, 0, inString, inString.Length, inString, 0, "", 0)
                dbLinesTEST.Add(inString)
            End While
            inputDBFileTEST.Close()

            If dbLinesTEST(5).StartsWith("\cm") Then 'This file is old format and must be converted first
                'convert files
                MsgBox("This conversion may fail because the .db file is a version prior to WS2.5.  Wordsurv will attempt to start the wsconv25.exe utility program to convert the file (you will see a console window appear), then will attempt the import again.  If this process fails, you may manually run the wsconv25.exe program in order to convert your .db file to 2.5 format for input." & vbCrLf & vbCrLf & "Usage from the command line is (assuming wsconv25.exe and your .db file are in the same directory):" & vbCrLf & "wsconv25 mywsfile.db" & vbCrLf & vbCrLf & "Your WS2.5 file will be converted and replace the original .db file (a backup of the original is also made  with a .xdb extension)", MsgBoxStyle.OkOnly, "File not WS2.5 format")
                'Throw New System.OperationCanceledException
                Dim process As System.Diagnostics.Process = New Process()
                process.StartInfo.FileName = AppDomain.CurrentDomain.BaseDirectory & "wsconv25.exe"
                process.StartInfo.Arguments = Chr(Asc("""")) & dbFileNameBase & ".db" & Chr(Asc(""""))
                'MsgBox(process.StartInfo.Arguments)
                process.Start()
                process.WaitForExit(15000)
                'If thepath contains the .xdb backup then
                MsgBox("Conversion was successful!!!  Loading converted file . . . ", MsgBoxStyle.OkOnly, "Conversion Successful")
            End If

            'open files
            Dim dbLines As New List(Of String)
            Dim inputDBFile As New IO.StreamReader(dbFileNameBase & ".db", System.Text.Encoding.GetEncoding(encoding2_5))

            While inputDBFile.Peek <> -1
                inString = inputDBFile.ReadLine
                inString = badCharClean(inString)
                'WideCharToMultiByte(1252, 0, inString, inString.Length, inString, 0, "", 0)
                dbLines.Add(inString)
            End While
            inputDBFile.Close()

            Dim catLines As New List(Of String)
            If System.IO.File.Exists(dbFileNameBase & ".cat") = True Then 'AJW***
                Dim inputCATFile As New IO.StreamReader(dbFileNameBase & ".cat", System.Text.Encoding.GetEncoding(encoding2_5))
                While inputCATFile.Peek <> -1
                    inString = inputCATFile.ReadLine
                    catLines.Add(inString)
                End While
                inputCATFile.Close()
            Else
                'Dim CATmaker As New IO.StreamWriter(dbFileNameBase & ".cat")
                Dim strLine As String
                'Dim symbols As New List(Of String)
                For Each strLine In dbLines
                    If Mid(strLine, 1, 2) = "\l" Then
                        For i As Integer = 5 To strLine.Length()
                            If Not catLines.Contains("\symbol " & Mid(strLine, i, 1)) Then
                                If Not Mid(strLine, i, 1) = "," Then
                                    catLines.Add("\symbol " & Mid(strLine, i, 1))
                                    catLines.Add("\title  " & Mid(strLine, i, 1))
                                    catLines.Add("")
                                End If
                            End If
                        Next
                    End If
                Next
                'Dim symbol As String
                'For Each symbol In symbols
                '    CATmaker.WriteLine("\symbol " & symbol)
                '    CATmaker.WriteLine("\title  " & symbol)
                '    CATmaker.WriteLine("")
                'Next
                'CATmaker.Close()
                'create from .db file
            End If

            Dim outputFile As New IO.StreamWriter(dbFileNameBase & ".wsv", False, System.Text.Encoding.UTF8)

            'parse data and produce structure
            'DICTIONARY
            outputFile.WriteLine("Start Dictionary")
            outputFile.WriteLine(vbTab & "Name=New Dictionary 1")

            'GLOSSES
            outputFile.WriteLine(vbTab & "Start Glosses")
            'loop through entire list looking for all \record labels
            counter = 0
            For i As Integer = 0 To dbLines.Count - 1
                For j As Integer = 0 To i - 1
                    If dbLines(j).Contains("\record") AndAlso dbLines(j) = dbLines(i) Then
                        dbLines(i) &= i.ToString
                    End If
                Next
                If dbLines(i).Contains("\record") Then
                    outputFile.WriteLine(vbTab & vbTab & dbLines(i).Substring(8) & "||||")
                    counter += 1
                End If
            Next
            outputFile.WriteLine(vbTab & "End Glosses")

            'SORT
            outputFile.WriteLine(vbTab & "Start Sorts")
            outputFile.Write(vbTab & vbTab & "Import Sort Order")
            ''loop through entire file looking for all \al records (ACCORDING TO INTERNAL SORT VALUE)
            'For i As Int16 = 0 To dbLines.Count - 1
            '    If dbLines(i).Contains("\al") Then
            '        outputFile.Write("|" & Trim(dbLines(i).Substring(4)))
            '    End If
            'Next
            'Just straight count as entered
            For i As Integer = 0 To counter - 1
                outputFile.Write("|" & i.ToString)
            Next
            outputFile.WriteLine()
            outputFile.WriteLine(vbTab & "End Sorts")

            'END OF DICTIONARY INFORMATION
            outputFile.WriteLine(vbTab & "Current Sort=0")
            outputFile.WriteLine(vbTab & "Current Gloss=0")
            outputFile.WriteLine(vbTab & "Current Gloss Column Index=0")
            outputFile.WriteLine("End Dictionary")

            'SURVEY
            Dim varietyIndex As Integer = 0
            outputFile.WriteLine("Start Survey")
            outputFile.WriteLine(vbTab & "Associated Dictionary=0")
            outputFile.WriteLine(vbTab & "Name=Survey1")
            outputFile.WriteLine(vbTab & "Start Varieties")
            varietyCount = 0
            While (getNextVariety(catLines, varietyIndex))
                varietyCount += 1
                langSymbol = catLines(varietyIndex).Substring(8)
                outputFile.WriteLine(vbTab & vbTab & "Start Variety")
                outputFile.WriteLine(vbTab & vbTab & vbTab & "Associated Dictionary=0")

                VarietyName = catLines(varietyIndex + 1).Substring(8)
                For j As Integer = 0 To varietyIndex
                    If catLines(j).Contains("\title") AndAlso catLines(j) = catLines(varietyIndex + 1) Then
                        VarietyName &= varietyIndex.ToString
                    End If
                Next

                outputFile.WriteLine(vbTab & vbTab & vbTab & "Name = " & VarietyName)
                outputFile.WriteLine(vbTab & vbTab & vbTab & "Start Transcriptions")
                Dim dbLineIndex As Integer = 0
                Dim currentGlossIndex As Integer = -1
                Dim buildTranscription As String
                Dim excludeString As String = ""
                While (dbLineIndex < dbLines.Count)
                    If dbLines(dbLineIndex).Contains("\record") Then
                        currentGlossIndex += 1
                        buildTranscription = ""
                        excludeString = ""


                        While Not (dbLines(dbLineIndex).Contains("\end"))
                            While ((dbLineIndex < dbLines.Count) AndAlso (Not (dbLines(dbLineIndex).Contains("\l")) AndAlso (Not (dbLines(dbLineIndex).Contains("\end")))))
                                dbLineIndex += 1
                            End While
                            'Found an \l, or \end, or reached end of file
                            If (dbLineIndex < dbLines.Count) Then 'be sure still a valid line in array
                                If dbLines(dbLineIndex).Contains("\l  ") Then
                                    If (dbLines(dbLineIndex).Contains(langSymbol)) Then
                                        'write out gloss form transcription
                                        cognateSymbol = ""
                                        buildTranscription = buildTranscription & Trim(dbLines(dbLineIndex - 2).Substring(2)) & ","
                                        cognateSymbol = Trim(dbLines(dbLineIndex - 3).Substring(2))
                                        If cognateSymbol = "0" Then
                                            excludeString = "x"
                                            cognateSymbol = ""
                                        End If
                                    End If
                                    dbLineIndex += 1 'AJW***Moved from below the next end if upt 1 line to here so only if a /l line is found
                                End If
                            End If
                            If dbLineIndex = dbLines.Count Then '***Exceeding array count, trips error on while for bad index
                                Exit While
                            End If
                        End While 'FOUND an \end

                        buildTranscription = MapKeymanStrToUnicodeStr(buildTranscription)

                        If buildTranscription = "" Then
                            outputFile.WriteLine(vbTab & vbTab & vbTab & vbTab & currentGlossIndex & "|" & currentGlossIndex & "|||")
                            comparisonGlossList.Add(currentGlossIndex & "||" & cognateSymbol & "||" & excludeString)
                        Else
                            outputFile.WriteLine(vbTab & vbTab & vbTab & vbTab & currentGlossIndex & "|" & currentGlossIndex & "|" & Mid(buildTranscription, 1, (Len(buildTranscription) - 1)) & "||")
                            comparisonGlossList.Add(currentGlossIndex & "|" & Mid(buildTranscription, 1, (Len(buildTranscription) - 1)) & "|" & cognateSymbol & "||" & excludeString)
                        End If
                        dbLineIndex += 1
                    End If
                    dbLineIndex += 1
                End While
                outputFile.WriteLine(vbTab & vbTab & vbTab & "End Transcriptions")
                outputFile.WriteLine(vbTab & vbTab & vbTab & "Current Transcription=0")
                outputFile.WriteLine(vbTab & vbTab & vbTab & "Description=This is a note about variety " & VarietyName)
                outputFile.WriteLine(vbTab & vbTab & "End Variety")
                varietyIndex += 1
            End While
            outputFile.WriteLine(vbTab & "End Varieties")
            outputFile.WriteLine(vbTab & "Current Variety=0")
            outputFile.WriteLine(vbTab & "Description=This is a note about the Survey")
            outputFile.WriteLine(vbTab & "Current VarietyEntry Column Index=0")
            outputFile.WriteLine("End Survey")

            outputFile.WriteLine()
            outputFile.WriteLine("Start Comparison")
            outputFile.WriteLine(vbTab & "Associated Survey=0")
            outputFile.WriteLine(vbTab & "Name=Comparison1")
            outputFile.Write(vbTab & "Variety Sort=")
            For i As Integer = 0 To varietyCount - 2
                outputFile.Write(i.ToString & "|")
            Next
            outputFile.Write((varietyCount - 1).ToString)

            outputFile.WriteLine()
            outputFile.WriteLine(vbTab & "Start Comparison Entries")

            'MONGO BUILD COMP ENTRIES CODE
            For Each comval As String In comparisonGlossList
                outputFile.WriteLine(vbTab & vbTab & comval)
            Next

            outputFile.WriteLine(vbTab & "End Comparison Entries")
            outputFile.WriteLine(vbTab & "Current Variety=0")
            outputFile.WriteLine(vbTab & "Description=")
            outputFile.WriteLine(vbTab & "Start Date=")
            outputFile.WriteLine(vbTab & "End Date=")
            outputFile.WriteLine(vbTab & "Current Variety Column Index=0")
            outputFile.WriteLine(vbTab & "DD Used Chars=")
            outputFile.WriteLine(vbTab & "Start DD Values")
            outputFile.WriteLine(vbTab & "End DD Values")
            outputFile.WriteLine(vbTab & "Excluded DD Characters=")
            outputFile.WriteLine(vbTab & "End Comparison")
            outputFile.WriteLine()

            outputFile.WriteLine("Start WordSurv Data")
            outputFile.WriteLine(vbTab & "Current Dictionary=0")
            outputFile.WriteLine(vbTab & "Current Survey=0")
            outputFile.WriteLine(vbTab & "Current Comparison=0")
            outputFile.WriteLine(vbTab & "Primary Language=Primary Gloss")
            outputFile.WriteLine(vbTab & "Secondary Language=Secondary Gloss")
            outputFile.WriteLine(vbTab & "Primary Font=Microsoft Sans Serif,8")
            outputFile.WriteLine(vbTab & "Secondary Font=Microsoft Sans Serif,8")
            outputFile.WriteLine(vbTab & "Transcription Font=Microsoft Sans Serif,8")
            outputFile.WriteLine("End WordSurv Data")

            outputFile.Close()
            Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

        End Sub


        Private Shared Function getNextVariety(ByVal catLines As List(Of String), ByRef varietyIndex As Integer) As Boolean
            While (varietyIndex < catLines.Count) AndAlso Not (catLines(varietyIndex).Contains("\symbol"))
                varietyIndex += 1
            End While
            If varietyIndex = catLines.Count Then
                Return False
            Else
                Return True
            End If
        End Function
        Private Shared Function getNextTranscription(ByVal dbLines As List(Of String), ByVal symbol As String, ByRef dbLineIndex As Integer) As Boolean
            While (dbLineIndex < dbLines.Count) AndAlso Not (dbLines(dbLineIndex).Contains("\l  " & symbol))
                dbLineIndex += 1
            End While
            If dbLineIndex = dbLines.Count Then
                Return False
            Else
                Return True
            End If
        End Function
        Public Sub WriteFile()
            Try
                Dim writer As New StreamWriter(Me.filename, False, System.Text.Encoding.UTF8)
                Me.WriteFileUsingWriter2(writer)
                writer.Close()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical)
            End Try
        End Sub
        Public Sub WriteFileUsingWriter(ByVal writer As StreamWriter)

            Dim glossMappers As New Dictionary(Of Dictionary, Dictionary(Of Gloss, Integer))
            Dim theFileStr As String = ""
            'Write the header
            theFileStr &= "WordSurv version 7" & vbCrLf & "Beta release Kemuel" & vbCrLf
            'Write Dictionaries
            For Each dict As Dictionary In Me.Dictionaries
                Dim glossMapper As New Dictionary(Of Gloss, Integer)
                theFileStr &= "Start Dictionary" & vbCrLf
                theFileStr &= vbTab & "Name=" & dict.Name & vbCrLf
                theFileStr &= vbTab & "Start Glosses" & vbCrLf
                Dim index As Integer = 0
                For Each gl As Gloss In dict.CurrentSort.Glosses
                    glossMapper.Add(gl, index) 'Give the glosses an id relative to the current sort order
                    index += 1
                    theFileStr &= vbTab & vbTab & gl.Name.Replace("|"c, "♠"c) & "|" & _
                                             gl.Name2.Replace("|"c, "♠"c) & "|" & _
                                             gl.PartOfSpeech.Replace("|"c, "♠"c) & "|" & _
                                             gl.FieldTip.Replace("|"c, "♠"c) & "|" & _
                                             gl.Comments.Replace("|"c, "♠"c) & vbCrLf
                Next
                theFileStr &= vbTab & "End Glosses" & vbCrLf
                theFileStr &= vbTab & "Start Sorts" & vbCrLf
                For Each srt As Sort In dict.Sorts
                    Dim sortStr As String = srt.Name.Replace("|"c, "♠"c)
                    For Each gl As Gloss In srt.Glosses
                        sortStr &= "|" & glossMapper(gl).ToString.Replace("|"c, "♠"c)
                    Next
                    theFileStr &= vbTab & vbTab & sortStr & vbCrLf
                Next
                theFileStr &= vbTab & "End Sorts" & vbCrLf
                theFileStr &= vbTab & "Current Sort=" & dict.Sorts.IndexOf(dict.CurrentSort).ToString & vbCrLf
                theFileStr &= vbTab & "Current Gloss=" & dict.CurrentSort.Glosses.IndexOf(dict.CurrentGloss).ToString & vbCrLf
                theFileStr &= vbTab & "Current Gloss Column Index=" & dict.CurrentGlossColumnIndex.ToString & vbCrLf

                theFileStr &= "End Dictionary" & vbCrLf & vbCrLf
                glossMappers.Add(dict, glossMapper) 'Store these for later use
            Next

            'Write Surveys
            Dim transMappers As New Dictionary(Of Variety, Dictionary(Of VarietyEntry, Integer))
            For Each surv As Survey In Me.Surveys
                theFileStr &= "Start Survey" & vbCrLf
                theFileStr &= vbTab & "Associated Dictionary=" & Me.Dictionaries.IndexOf(surv.AssociatedDictionary).ToString & vbCrLf
                theFileStr &= vbTab & "Name=" & surv.Name.Replace("|"c, "♠"c) & vbCrLf
                theFileStr &= vbTab & "Start Varieties" & vbCrLf
                For Each var As Variety In surv.Varieties
                    Dim transMapper As New Dictionary(Of VarietyEntry, Integer)

                    theFileStr &= vbTab & vbTab & "Start Variety" & vbCrLf
                    theFileStr &= vbTab & vbTab & vbTab & "Associated Dictionary=" & Me.Dictionaries.IndexOf(var.AssociatedDictionary).ToString & vbCrLf
                    theFileStr &= vbTab & vbTab & vbTab & "Name=" & var.Name.Replace("|"c, "♠"c) & vbCrLf
                    theFileStr &= vbTab & vbTab & vbTab & "Start Transcriptions" & vbCrLf

                    Dim glossMapper As Dictionary(Of Gloss, Integer) = glossMappers(var.AssociatedDictionary)

                    Dim index As Integer = 0
                    For Each gl As Gloss In glossMapper.Keys
                        Dim varEntry As VarietyEntry = var.VarietyEntries(gl)
                        transMapper.Add(varEntry, index)
                        theFileStr &= vbTab & vbTab & vbTab & vbTab & index.ToString & "|" & glossMapper(gl).ToString.Replace("|"c, "♠"c) & "|" & varEntry.Transcription.Replace("|"c, "♠"c) & "|" & _
                                                                                                          varEntry.PluralFrame.Replace("|"c, "♠"c) & "|" & _
                                                                                                          varEntry.Notes.Replace("|"c, "♠"c) & vbCrLf
                        index += 1
                    Next
                    theFileStr &= vbTab & vbTab & vbTab & "End Transcriptions" & vbCrLf
                    Dim currentVarietyEntry As String
                    Try
                        currentVarietyEntry = transMapper(var.CurrentVarietyEntry).ToString
                    Catch ex As Exception
                        currentVarietyEntry = "-1"
                    End Try
                    theFileStr &= vbTab & vbTab & vbTab & "Current VarietyEntry=" & currentVarietyEntry & vbCrLf
                    theFileStr &= vbTab & vbTab & vbTab & "Description=" & var.Description.Replace(vbCrLf, "\").Replace("|"c, "♠"c) & vbCrLf

                    theFileStr &= vbTab & vbTab & "End Variety" & vbCrLf
                    transMappers.Add(var, transMapper)
                Next
                theFileStr &= vbTab & "End Varieties" & vbCrLf
                theFileStr &= vbTab & "Current Variety=" & surv.Varieties.IndexOf(surv.CurrentVariety).ToString & vbCrLf
                theFileStr &= vbTab & "Description=" & surv.Description.Replace(vbCrLf, "\").Replace("|"c, "♠"c) & vbCrLf
                theFileStr &= vbTab & "Current VarietyEntry Column Index=" & surv.CurrentVarietyEntryColumnIndex.ToString & vbCrLf
                theFileStr &= "End Survey" & vbCrLf & vbCrLf
            Next

            'Write Comparisons
            For Each comp As Comparison In Me.Comparisons
                theFileStr &= "Start Comparison" & vbCrLf
                theFileStr &= vbTab & "Associated Survey=" & Me.Surveys.IndexOf(comp.AssociatedSurvey).ToString & vbCrLf
                theFileStr &= vbTab & "Name=" & comp.Name.Replace("|"c, "♠"c) & vbCrLf
                Dim varSortStr As String = "Variety Sort="
                For Each var As Variety In comp.CurrentVarietySort
                    varSortStr &= comp.AssociatedSurvey.Varieties.IndexOf(var).ToString & "|"
                Next
                varSortStr = varSortStr.TrimEnd("|"c)
                theFileStr &= vbTab & varSortStr & vbCrLf

                theFileStr &= vbTab & "Start Comparison Entries" & vbCrLf
                For Each var As Variety In comp.AssociatedSurvey.Varieties
                    Dim transMapper As Dictionary(Of VarietyEntry, Integer) = transMappers(var)
                    For Each tr As VarietyEntry In transMappers(var).Keys
                        Dim compEntry As ComparisonEntry = comp.ComparisonEntries(tr)
                        theFileStr &= vbTab & vbTab & transMapper(tr).ToString & "|" & compEntry.AlignedRendering.Replace("|"c, "♠"c) & "|" & _
                                                                                          compEntry.Grouping.Replace("|"c, "♠"c) & "|" & _
                                                                                          compEntry.Notes.Replace("|"c, "♠"c) & "|" & _
                                                                                          compEntry.Exclude.Replace("|"c, "♠"c) & vbCrLf
                    Next
                Next
                theFileStr &= vbTab & "End Comparison Entries" & vbCrLf
                theFileStr &= vbTab & "Current Variety=" & comp.AssociatedSurvey.Varieties.IndexOf(comp.CurrentVariety).ToString & vbCrLf
                theFileStr &= vbTab & "Description=" & comp.Description.Replace(vbCrLf, "\").Replace("|"c, "♠"c) & vbCrLf
                theFileStr &= vbTab & "Start Date=" & vbCrLf
                theFileStr &= vbTab & "End Date=" & vbCrLf
                theFileStr &= vbTab & "Current Variety Column Index=" & comp.CurrentVarietyColumnIndex.ToString & vbCrLf

                'Write out DDs.  Even though the user only sees half of the grid, the full grid is still stored to make it easier to deal with.
                Dim DDs As Dictionary(Of Integer, Integer) = comp.AssociatedDegreesOfDifference.DDs
                Dim usedChars As List(Of String) = comp.AssociatedDegreesOfDifference.UsedCharsList

                'Collect the characters used by the DD grid and write them out on one line
                Dim usedCharsStr As String = ""
                For Each usedChar As String In usedChars
                    usedCharsStr &= usedChar.Replace("|"c, "♠"c) & "|"
                Next
                usedCharsStr = usedCharsStr.TrimEnd("|"c)
                theFileStr &= vbTab & "DD Used Chars=" & usedCharsStr & vbCrLf
                theFileStr &= vbTab & "Start DD Values" & vbCrLf

                'Now collect all of the DD values and write them out in a big matrix
                For Each usedChar As String In usedChars
                    Dim ddRowStr As String = ""
                    For Each usedChar2 As String In usedChars
                        Dim char1AndChar2 As Integer = (AscW(usedChar) << 16) Or AscW(usedChar2)
                        ddRowStr &= DDs(char1AndChar2).ToString & "|"
                    Next
                    ddRowStr = ddRowStr.TrimEnd("|"c)
                    theFileStr &= vbTab & vbTab & ddRowStr & vbCrLf
                Next
                theFileStr &= vbTab & "End DD Values" & vbCrLf

                theFileStr &= vbTab & "Excluded DD Chars=" & comp.AssociatedDegreesOfDifference.ExcludedChars & vbCrLf
                theFileStr &= "End Comparison" & vbCrLf & vbCrLf
            Next

            theFileStr &= "Start WordSurv Data" & vbCrLf
            theFileStr &= vbTab & "Current Dictionary=" & Me.Dictionaries.IndexOf(Me.CurrentDictionary).ToString & vbCrLf
            theFileStr &= vbTab & "Current Survey=" & Me.Surveys.IndexOf(Me.CurrentSurvey).ToString & vbCrLf
            theFileStr &= vbTab & "Current Comparison=" & Me.Comparisons.IndexOf(Me.CurrentComparison).ToString & vbCrLf
            theFileStr &= vbTab & "Primary Language=" & Me.PrimaryLanguage.Replace("|"c, "♠"c) & vbCrLf
            theFileStr &= vbTab & "Secondary Language=" & Me.SecondaryLanguage.Replace("|"c, "♠"c) & vbCrLf
            theFileStr &= vbTab & "Primary Font=" & Me.PrimaryFont.Name & "," & Me.PrimaryFont.Size.ToString & vbCrLf
            theFileStr &= vbTab & "Secondary Font=" & Me.SecondaryFont.Name & "," & Me.SecondaryFont.Size.ToString & vbCrLf
            theFileStr &= vbTab & "Transcription Font=" & Me.TranscriptionFont.Name & "," & Me.TranscriptionFont.Size.ToString & vbCrLf
            theFileStr &= "End WordSurv Data" & vbCrLf

            'write code to remove all lone newlines 
            'Dim reg As Regex = New Regex("[^\r]\r-[^\r]")
            'AJW
            'reg.IsMatch(theFileStr)

            writer.Write(theFileStr)
        End Sub
        Public Sub WriteFileUsingWriter2(ByVal writer As StreamWriter)
            Dim temp As String

            Dim glossMappers As New Dictionary(Of Dictionary, Dictionary(Of Gloss, Integer))
            Dim theFileStr As String = ""
            'Write the header
            theFileStr &= "WordSurv version 7" & vbCrLf & "Beta release Kemuel" & vbCrLf
            'Write Dictionaries
            For Each dict As Dictionary In Me.Dictionaries
                Dim glossMapper As New Dictionary(Of Gloss, Integer)
                theFileStr &= "Start Dictionary" & vbCrLf
                theFileStr &= vbTab & "Name=" & dict.Name & vbCrLf
                theFileStr &= vbTab & "Start Glosses" & vbCrLf
                Dim index As Integer = 0
                For Each gl As Gloss In dict.CurrentSort.Glosses
                    glossMapper.Add(gl, index) 'Give the glosses an id relative to the current sort order
                    index += 1
                    temp = gl.Name.Replace("|", ChrW(448)) & "|" & gl.Name2.Replace("|", ChrW(448)) & "|" & gl.PartOfSpeech.Replace("|", ChrW(448)) & "|" & gl.FieldTip.Replace("|", ChrW(448)) & "|" & gl.Comments.Replace("|", ChrW(448))
                    temp = temp.Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ")
                    theFileStr &= vbTab & vbTab & temp & vbCrLf
                Next
                theFileStr &= vbTab & "End Glosses" & vbCrLf
                theFileStr &= vbTab & "Start Sorts" & vbCrLf
                For Each srt As Sort In dict.Sorts
                    Dim sortStr As String = srt.Name.Replace("|", ChrW(448)).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ")
                    For Each gl As Gloss In srt.Glosses
                        sortStr &= "|" & glossMapper(gl).ToString.Replace("|", ChrW(448)).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ")
                    Next
                    theFileStr &= vbTab & vbTab & sortStr & vbCrLf
                Next
                theFileStr &= vbTab & "End Sorts" & vbCrLf
                theFileStr &= vbTab & "Current Sort=" & dict.Sorts.IndexOf(dict.CurrentSort).ToString & vbCrLf
                theFileStr &= vbTab & "Current Gloss=" & dict.CurrentSort.Glosses.IndexOf(dict.CurrentGloss).ToString & vbCrLf
                theFileStr &= vbTab & "Current Gloss Column Index=" & dict.CurrentGlossColumnIndex.ToString & vbCrLf

                theFileStr &= "End Dictionary" & vbCrLf & vbCrLf
                glossMappers.Add(dict, glossMapper) 'Store these for later use
            Next

            'Write Surveys
            Dim transMappers As New Dictionary(Of Variety, Dictionary(Of VarietyEntry, Integer))
            For Each surv As Survey In Me.Surveys
                theFileStr &= "Start Survey" & vbCrLf
                theFileStr &= vbTab & "Associated Dictionary=" & Me.Dictionaries.IndexOf(surv.AssociatedDictionary).ToString & vbCrLf
                theFileStr &= vbTab & "Name=" & surv.Name.Replace("|", ChrW(448)).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ") & vbCrLf
                theFileStr &= vbTab & "Start Varieties" & vbCrLf
                For Each var As Variety In surv.Varieties
                    Dim transMapper As New Dictionary(Of VarietyEntry, Integer)

                    theFileStr &= vbTab & vbTab & "Start Variety" & vbCrLf
                    theFileStr &= vbTab & vbTab & vbTab & "Associated Dictionary=" & Me.Dictionaries.IndexOf(var.AssociatedDictionary).ToString.Replace("|", ChrW(448)).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ") & vbCrLf
                    theFileStr &= vbTab & vbTab & vbTab & "Name=" & var.Name.Replace("|", ChrW(448)).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ") & vbCrLf
                    theFileStr &= vbTab & vbTab & vbTab & "Start Transcriptions" & vbCrLf

                    Dim glossMapper As Dictionary(Of Gloss, Integer) = glossMappers(var.AssociatedDictionary)

                    Dim index As Integer = 0
                    For Each gl As Gloss In glossMapper.Keys
                        Dim varEntry As VarietyEntry = var.VarietyEntries(gl)
                        transMapper.Add(varEntry, index)
                        temp = glossMapper(gl).ToString.Replace("|", ChrW(448)).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ") & "|" & _
                            varEntry.Transcription.Replace("|", ChrW(448)).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ") & "|" & _
                            varEntry.PluralFrame.Replace("|", ChrW(448)).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ") & "|" & _
                            varEntry.Notes.Replace("|", ChrW(448)).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ")
                        temp = temp.Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ")
                        theFileStr &= vbTab & vbTab & vbTab & vbTab & index.ToString & "|" & temp & vbCrLf
                        index += 1
                    Next
                    theFileStr &= vbTab & vbTab & vbTab & "End Transcriptions" & vbCrLf
                    Dim currentVarietyEntry As String
                    Try
                        currentVarietyEntry = transMapper(var.CurrentVarietyEntry).ToString
                    Catch ex As Exception
                        currentVarietyEntry = "-1"
                    End Try
                    theFileStr &= vbTab & vbTab & vbTab & "Current VarietyEntry=" & currentVarietyEntry & vbCrLf
                    theFileStr &= vbTab & vbTab & vbTab & "Description=" & var.Description.Replace("|", ChrW(448)).Replace(vbCrLf, "\").Replace(vbCr, "\").Replace(vbLf, "\") & vbCrLf
                    theFileStr &= vbTab & vbTab & "End Variety" & vbCrLf
                    transMappers.Add(var, transMapper)
                Next
                theFileStr &= vbTab & "End Varieties" & vbCrLf
                theFileStr &= vbTab & "Current Variety=" & surv.Varieties.IndexOf(surv.CurrentVariety).ToString & vbCrLf
                theFileStr &= vbTab & "Description=" & surv.Description.Replace("|", ChrW(448)).Replace(vbCrLf, "\").Replace(vbCr, "\").Replace(vbLf, "\") & vbCrLf
                theFileStr &= vbTab & "Current VarietyEntry Column Index=" & surv.CurrentVarietyEntryColumnIndex.ToString & vbCrLf
                theFileStr &= "End Survey" & vbCrLf & vbCrLf
            Next
            'TO HERE
            'Write Comparisons
            For Each comp As Comparison In Me.Comparisons
                theFileStr &= "Start Comparison" & vbCrLf
                theFileStr &= vbTab & "Associated Survey=" & Me.Surveys.IndexOf(comp.AssociatedSurvey).ToString.Replace("|", ChrW(448)).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ") & vbCrLf
                theFileStr &= vbTab & "Name=" & comp.Name.Replace("|", ChrW(448)).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ") & vbCrLf
                Dim varSortStr As String = "Variety Sort="
                For Each var As Variety In comp.CurrentVarietySort
                    varSortStr &= comp.AssociatedSurvey.Varieties.IndexOf(var).ToString.Replace("|", ChrW(448)).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ") & "|"
                Next
                varSortStr = varSortStr.TrimEnd("|"c)
                theFileStr &= vbTab & varSortStr & vbCrLf

                theFileStr &= vbTab & "Start Comparison Entries" & vbCrLf
                For Each var As Variety In comp.AssociatedSurvey.Varieties
                    Dim transMapper As Dictionary(Of VarietyEntry, Integer) = transMappers(var)
                    For Each tr As VarietyEntry In transMappers(var).Keys
                        Dim compEntry As ComparisonEntry = comp.ComparisonEntries(tr)
                        theFileStr &= vbTab & vbTab & transMapper(tr).ToString.Replace("|", ChrW(448)).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ") & "|" & compEntry.AlignedRendering.Replace("|", ChrW(448)).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ") & "|" & _
                                                                                          compEntry.Grouping.Replace("|", ChrW(448)).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ") & "|" & _
                                                                                          compEntry.Notes.Replace("|", ChrW(448)).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ") & "|" & _
                                                                                          compEntry.Exclude.Replace("|", ChrW(448)).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ") & vbCrLf
                    Next
                Next
                theFileStr &= vbTab & "End Comparison Entries" & vbCrLf
                theFileStr &= vbTab & "Current Variety=" & comp.AssociatedSurvey.Varieties.IndexOf(comp.CurrentVariety).ToString.Replace("|", ChrW(448)).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ") & vbCrLf
                theFileStr &= vbTab & "Description=" & comp.Description.Replace("|", ChrW(448)).Replace(vbCrLf, "\").Replace(vbCr, "\").Replace(vbLf, "\") & vbCrLf
                theFileStr &= vbTab & "Start Date=" & vbCrLf
                theFileStr &= vbTab & "End Date=" & vbCrLf
                theFileStr &= vbTab & "Current Variety Column Index=" & comp.CurrentVarietyColumnIndex.ToString.Replace("|", ChrW(448)).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ") & vbCrLf

                'Write out DDs.  Even though the user only sees half of the grid, the full grid is still stored to make it easier to deal with.
                Dim DDs As Dictionary(Of Integer, Integer) = comp.AssociatedDegreesOfDifference.DDs
                Dim usedChars As List(Of String) = comp.AssociatedDegreesOfDifference.UsedCharsList

                'Collect the characters used by the DD grid and write them out on one line
                Dim usedCharsStr As String = ""
                For Each usedChar As String In usedChars
                    usedCharsStr &= usedChar.Replace("|", ChrW(448)).Replace(vbCrLf, " ") & "|"   '.Replace(vbCr, " ").Replace(vbLf, " ") & "|"
                Next
                usedCharsStr = usedCharsStr.TrimEnd("|"c)
                theFileStr &= vbTab & "DD Used Chars=" & usedCharsStr & vbCrLf
                theFileStr &= vbTab & "Start DD Values" & vbCrLf

                'Now collect all of the DD values and write them out in a big matrix
                For Each usedChar As String In usedChars
                    Dim ddRowStr As String = ""
                    For Each usedChar2 As String In usedChars
                        Dim char1AndChar2 As Integer = (AscW(usedChar) << 16) Or AscW(usedChar2)
                        ddRowStr &= DDs(char1AndChar2).ToString & "|"
                        'ddRowStr &= DDs(char1AndChar2).ToString.Replace("|", ChrW(448)).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ") & "|"
                    Next
                    ddRowStr = ddRowStr.TrimEnd("|"c)
                    theFileStr &= vbTab & vbTab & ddRowStr & vbCrLf
                Next
                theFileStr &= vbTab & "End DD Values" & vbCrLf

                theFileStr &= vbTab & "Excluded DD Chars=" & comp.AssociatedDegreesOfDifference.ExcludedChars & vbCrLf
                theFileStr &= "End Comparison" & vbCrLf & vbCrLf
            Next

            theFileStr &= "Start WordSurv Data" & vbCrLf
            theFileStr &= vbTab & "Current Dictionary=" & Me.Dictionaries.IndexOf(Me.CurrentDictionary).ToString & vbCrLf
            theFileStr &= vbTab & "Current Survey=" & Me.Surveys.IndexOf(Me.CurrentSurvey).ToString & vbCrLf
            theFileStr &= vbTab & "Current Comparison=" & Me.Comparisons.IndexOf(Me.CurrentComparison).ToString & vbCrLf
            theFileStr &= vbTab & "Primary Language=" & Me.PrimaryLanguage.Replace("|", ChrW(448)) & vbCrLf
            theFileStr &= vbTab & "Secondary Language=" & Me.SecondaryLanguage.Replace("|", ChrW(448)) & vbCrLf
            theFileStr &= vbTab & "Primary Font=" & Me.PrimaryFont.Name & "," & Me.PrimaryFont.Size.ToString & vbCrLf
            theFileStr &= vbTab & "Secondary Font=" & Me.SecondaryFont.Name & "," & Me.SecondaryFont.Size.ToString & vbCrLf
            theFileStr &= vbTab & "Transcription Font=" & Me.TranscriptionFont.Name & "," & Me.TranscriptionFont.Size.ToString & vbCrLf
            theFileStr &= "End WordSurv Data" & vbCrLf

            'write code to remove all lone newlines 
            'Dim reg As Regex = New Regex("[^\r]\r-[^\r]")
            'AJW
            'reg.IsMatch(theFileStr)

            writer.Write(theFileStr)
        End Sub
        Public Sub SetCurrentComparisonsDDCurrentRowIndex(ByVal index As Integer)
            Me.CurrentComparison.AssociatedDegreesOfDifference.CurrentRowIndex = index
        End Sub
        Public Sub DuplicateCurrentComparison(ByVal newName As String)
            Dim newComparison As New Comparison(Me.CurrentComparison.AssociatedSurvey, newName, False)

            newComparison.AssociatedSurvey = Me.CurrentComparison.AssociatedSurvey
            newComparison.CurrentVariety = Me.CurrentComparison.CurrentVariety
            newComparison.Description = Me.CurrentComparison.Description
            newComparison.EndDate = Me.CurrentComparison.EndDate
            newComparison.StartDate = Me.CurrentComparison.StartDate
            newComparison.CurrentVarietySort = Me.CurrentComparison.CurrentVarietySort

            For Each gl As Gloss In Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses
                For Each var As Variety In Me.CurrentComparison.AssociatedSurvey.Varieties
                    Dim thisVarEntry As VarietyEntry = var.VarietyEntries(gl)
                    newComparison.ComparisonEntries.Add(thisVarEntry, Me.CurrentComparison.ComparisonEntries(thisVarEntry).Copy())
                Next
            Next

            Me.Comparisons.Add(newComparison)
            Me.CurrentComparison = newComparison
        End Sub
        Public Sub SortCurrentDictionaryAlphabetically(ByVal firstCellIndex As Integer, ByVal lastCellIndex As Integer, ByVal colIndex As Integer)
            dontUseThisGlobalUnlessYouAreTheGlossComparer = colIndex
            Me.CurrentDictionary.CurrentSort.Glosses.Sort(firstCellIndex, lastCellIndex - firstCellIndex + 1, New GlossComparer)
        End Sub
        Public Sub SortCurrentComparisonAlphabetically(ByVal firstCellIndex As Integer, ByVal lastCellIndex As Integer, ByVal colindex As Integer)
            Dim sorter As New ComparisonComparer
            sorter.ComparisonComparerColIndex = colindex
            sorter.ComparisonComparerGloss = Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentGloss
            sorter.ComparisonEntries = Me.CurrentComparison.ComparisonEntries
            Me.CurrentComparison.CurrentVarietySort.Sort(firstCellIndex, lastCellIndex - firstCellIndex + 1, sorter)
        End Sub
        Private Class ComparisonComparer
            Implements IComparer(Of Variety)

            Public ComparisonComparerColIndex As Integer
            Public ComparisonComparerGloss As Gloss
            Public ComparisonEntries As Dictionary(Of VarietyEntry, ComparisonEntry)

            Public Function Compare(ByVal x As Variety, ByVal y As Variety) As Integer Implements System.Collections.Generic.IComparer(Of WordSurv7.DataObjects.Variety).Compare
                Select Case Me.ComparisonComparerColIndex
                    Case 0 : Return String.Compare(x.Name, y.Name)
                    Case 1 : Return String.Compare(x.VarietyEntries(Me.ComparisonComparerGloss).Transcription, y.VarietyEntries(Me.ComparisonComparerGloss).Transcription)
                    Case 2 : Return String.Compare(x.VarietyEntries(Me.ComparisonComparerGloss).PluralFrame, y.VarietyEntries(Me.ComparisonComparerGloss).PluralFrame)
                    Case 3 : Return String.Compare(Me.ComparisonEntries(x.VarietyEntries(Me.ComparisonComparerGloss)).AlignedRendering, Me.ComparisonEntries(y.VarietyEntries(Me.ComparisonComparerGloss)).AlignedRendering)
                    Case 4 : Return String.Compare(Me.ComparisonEntries(x.VarietyEntries(Me.ComparisonComparerGloss)).Grouping, Me.ComparisonEntries(y.VarietyEntries(Me.ComparisonComparerGloss)).Grouping)
                    Case 5 : Return String.Compare(Me.ComparisonEntries(x.VarietyEntries(Me.ComparisonComparerGloss)).Notes, Me.ComparisonEntries(y.VarietyEntries(Me.ComparisonComparerGloss)).Notes)
                    Case 6 : Return String.Compare(Me.ComparisonEntries(x.VarietyEntries(Me.ComparisonComparerGloss)).Exclude, Me.ComparisonEntries(y.VarietyEntries(Me.ComparisonComparerGloss)).Exclude)
                End Select
                Throw New ArgumentException
            End Function
        End Class
        Public Sub PropogateNewGlossAfterDictionary(ByRef gl As Gloss)
            'Put the gloss into all the varieties in all the surveys that use this dictionary,
            For Each surv As Survey In GetSurveysThatUseThisDictionary(Me.CurrentDictionary)
                For Each var As Variety In surv.Varieties
                    Dim newVarEntry As New VarietyEntry
                    var.VarietyEntries.Add(gl, newVarEntry)
                    'and all the comparisons that use all the surveys that use this dictionary.
                    For Each comp As Comparison In GetComparisonsThatUseThisDictionary(Me.CurrentDictionary)
                        comp.ComparisonEntries.Add(newVarEntry, New ComparisonEntry(newVarEntry.Transcription))
                    Next
                Next
            Next
        End Sub
        Public Sub InsertNewGloss(ByVal glossRow As Integer, ByVal glossName As String)
            Dim newGloss As New Gloss(glossName)

            'Since each sort is just a list of the same glosses, intert this gloss into all of them.
            Me.CurrentDictionary.CurrentSort.Glosses.Insert(glossRow, newGloss)
            For Each srt As Sort In Me.CurrentDictionary.Sorts
                If srt Is Me.CurrentDictionary.CurrentSort Then Continue For
                srt.Glosses.Add(newGloss)
            Next
            Me.CurrentDictionary.CurrentGloss = newGloss

            Me.PropogateNewGlossAfterDictionary(newGloss)
        End Sub
        Public Sub UpdateGlossValue(ByVal glossRow As Integer, ByVal glossCol As Integer, ByVal val As String)
            If val = "=rand()" Then
                For Each gl As Gloss In Me.CurrentDictionary.CurrentSort.Glosses
                    gl.Name = "ALL WORK AND NO FRUIT MAKES GEL-ARSHIE A DULL BOY"
                    gl.Name2 = "ALL WORK AND NO FRUIT MAKES GEL-ARSHIE A DULL BOY"
                    gl.PartOfSpeech = "ALL WORK AND NO FRUIT MAKES GEL-ARSHIE A DULL BOY"
                    gl.FieldTip = "ALL WORK AND NO FRUIT MAKES GEL-ARSHIE A DULL BOY"
                    gl.Comments = "ALL WORK AND NO FRUIT MAKES GEL-ARSHIE A DULL BOY"
                Next
            Else
                Me.CurrentDictionary.CurrentSort.Glosses(glossRow).SetByIndex(glossCol, val)
            End If
        End Sub

        Public Function IsEmpty() As Boolean
            Return Me.filename Is Nothing
        End Function
        Public Function IsUniqueDictionaryName(ByVal name As String) As Boolean
            For Each dict As Dictionary In Me.Dictionaries
                If dict.Name = name Then Return False
            Next
            Return True
        End Function
        Public Function IsUniqueDictionarySortName(ByVal name As String) As Boolean
            For Each srt As Sort In Me.CurrentDictionary.Sorts
                If srt.Name = name Then Return False
            Next
            Return True
        End Function
        Public Function IsUniqueSurveyName(ByVal name As String) As Boolean
            For Each surv As Survey In Me.Surveys
                If surv.Name = name Then Return False
            Next
            Return True
        End Function
        Public Function IsUniqueVarietyName(ByVal name As String) As Boolean
            For Each var As Variety In Me.CurrentSurvey.Varieties
                If var.Name = name Then Return False
            Next
            Return True
        End Function
        Public Function IsUniqueComparisonName(ByVal name As String) As Boolean
            For Each comp As Comparison In Me.Comparisons
                If comp.Name = name Then Return False
            Next
            Return True
        End Function
        Public Function IsGlossInCurrentDictionary(ByVal glossName As String, ByVal excludeText As String) As Boolean
            For Each gl As Gloss In Me.CurrentDictionary.CurrentSort.Glosses
                If gl.Name = glossName And Not gl.Name = excludeText Then Return True
            Next
            Return False
        End Function


        Public Sub AddSelectedCOMPASSPhoneCoordinate(ByVal addr As CellAddress)
            Me.CurrentComparison.SelectedPhonePairCoordinates.Add(addr)
        End Sub
        Public Sub ClearSelectedCOMPASSPhoneCoordinates()
            Me.CurrentComparison.SelectedPhonePairCoordinates.Clear()
        End Sub
        Public Sub MoveGlosses(ByVal cutIndices As List(Of Integer), ByVal pasteIndex As Integer)
            If cutIndices.Count = 0 Then Return
            'Stupid edge case - if they select all rows and try to paste, it will crash
            If cutIndices.Count = Me.CurrentDictionary.CurrentSort.Glosses.Count Then Return

            Dim pasteGloss As Gloss
            Try
                pasteGloss = Me.CurrentDictionary.CurrentSort.Glosses(pasteIndex)
            Catch ex As Exception
                pasteGloss = Nothing
            End Try

            Dim movedGlosses As New List(Of Gloss)
            cutIndices.Sort()
            cutIndices.Reverse()

            For Each cutIndex As Integer In cutIndices
                movedGlosses.Add(Me.CurrentDictionary.CurrentSort.Glosses(cutIndex))
                Me.CurrentDictionary.CurrentSort.Glosses.RemoveAt(cutIndex)
                If pasteIndex > cutIndex + 1 Then pasteIndex -= 1
            Next

            'If they paste onto the last row, then the paste index will be one larger than the size of the collection of glosses.
            'In this case, dump them at the end, otherwise at the index needed.
            If pasteGloss IsNot Nothing Then
                For Each movedGloss As Gloss In movedGlosses
                    Me.CurrentDictionary.CurrentSort.Glosses.Insert(pasteIndex, movedGloss)
                Next
                Me.CurrentDictionary.CurrentGloss = movedGlosses(0)
            Else
                movedGlosses.Reverse()
                For Each movedGloss As Gloss In movedGlosses
                    Me.CurrentDictionary.CurrentSort.Glosses.Add(movedGloss)
                Next
                Me.CurrentDictionary.CurrentGloss = movedGlosses(movedGlosses.Count - 1)
            End If
        End Sub
        Public Sub DeleteCurrentDictionary()
            Me.Dictionaries.Remove(Me.CurrentDictionary)

            Dim surveysToRemove As New List(Of Survey)
            'Delete all other data objects that use this dictionary (deleting the other objects happen inside DeleteThisSurvey()).
            For Each surv As Survey In Me.Surveys
                If surv.AssociatedDictionary.Equals(Me.CurrentDictionary) Then
                    surveysToRemove.Add(surv)
                End If
            Next
            For Each surv As Survey In surveysToRemove 'Do this because you cannot modify a collection you are iterating over.
                Me.DeleteThisSurvey(surv)
            Next

            Try
                Me.CurrentDictionary = Me.Dictionaries(0)
            Catch ex As Exception 'If they deleted the last dictionary,
                Me.CurrentDictionary = Nothing
            End Try
        End Sub
        Public Sub RenameCurrentDictionary(ByVal newName As String)
            Me.CurrentDictionary.Name = newName
        End Sub
        Public Sub RenameCurrentDictionarysCurrentSort(ByVal newName As String)
            Me.CurrentDictionary.CurrentSort.Name = newName
        End Sub

        Public Sub DeleteGlossesFromCurrentDictionary(ByVal indicesOfGlossesToDelete As List(Of Integer))

            'Convert the indices to gloss objects.
            Dim glossesToDelete As New List(Of Gloss)
            For Each index As Integer In indicesOfGlossesToDelete
                If index < Me.CurrentDictionary.CurrentSort.Glosses.Count Then glossesToDelete.Add(Me.CurrentDictionary.CurrentSort.Glosses(index))
            Next

            'Remove entries from the affected dictionary.
            For Each srt As Sort In Me.CurrentDictionary.Sorts
                For Each gl As Gloss In glossesToDelete
                    srt.Glosses.Remove(gl)
                Next
            Next

            'Prepare lists of items to be removed to aid in looping over them.
            Dim surveysAffected As New List(Of Survey)
            For Each surv As Survey In Me.Surveys
                If surv.AssociatedDictionary.Equals(Me.CurrentDictionary) Then
                    surveysAffected.Add(surv)
                End If
            Next
            Dim comparisonsAffected As New List(Of Comparison)
            For Each comp As Comparison In Me.Comparisons
                For Each surv As Survey In surveysAffected
                    If comp.AssociatedSurvey.Equals(surv) Then
                        comparisonsAffected.Add(comp)
                        Exit For
                    End If
                Next
            Next

            For Each surv As Survey In surveysAffected
                'Remove entries from any comparisons that use this survey.
                For Each comp As Comparison In comparisonsAffected
                    For Each var As Variety In surv.Varieties
                        For Each gl As Gloss In glossesToDelete
                            Dim varEntry As VarietyEntry = var.VarietyEntries(gl)
                            comp.ComparisonEntries.Remove(varEntry)

                        Next
                    Next
                Next
                'Remove each gloss from any varieties in surveys that use this dictionary.
                For Each var As Variety In surv.Varieties
                    For Each gl As Gloss In glossesToDelete
                        var.VarietyEntries.Remove(gl)
                    Next
                Next
            Next

            'Put the current gloss to be the gloss nearest in sort order to the ones we deleted.
            If indicesOfGlossesToDelete(indicesOfGlossesToDelete.Count - 1) < Me.CurrentDictionary.CurrentSort.Glosses.Count Then
                Me.CurrentDictionary.CurrentGloss = Me.CurrentDictionary.CurrentSort.Glosses(indicesOfGlossesToDelete(indicesOfGlossesToDelete.Count - 1))
            ElseIf Me.CurrentDictionary.CurrentSort.Glosses.Count > 0 Then
                Me.CurrentDictionary.CurrentGloss = Me.CurrentDictionary.CurrentSort.Glosses(Me.CurrentDictionary.CurrentSort.Glosses.Count - 1)
            Else
                Me.CurrentDictionary.CurrentGloss = Nothing
            End If
        End Sub
        Public Function UpdateComparisonEntryValue(ByVal rowIndex As Integer, ByVal colIndex As Integer, ByVal val As String) As String
            Dim theGloss As Gloss = Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentGloss
            Dim theVariety As Variety = Me.CurrentComparison.CurrentVarietySort(rowIndex)
            Dim theVarEntry As VarietyEntry = theVariety.VarietyEntries(theGloss)

            Select Case colIndex
                Case 3
                    Me.CurrentComparison.ComparisonEntries(theVarEntry).AlignedRendering = val
                    HasNotSaved = True 'AJW*** was not saving auto on changes to comp grid
                    If Not (HaveSameNumberOfCommas(Me.CurrentComparison.ComparisonEntries(theVarEntry).Grouping, val)) Then
                        Return "The number of comma-separated items in the aligned rendering and groupings fields do not match!"
                    End If
                Case 4
                    'Validate the grouping to make sure it has the right number of commas.
                    'If HaveSameNumberOfCommas(theVarEntry.Transcription, val) Or val = "" Then 'AJW*** changed to check aligned rendering, not transcription
                    If HaveSameNumberOfCommas(Me.CurrentComparison.ComparisonEntries(theVarEntry).AlignedRendering, val) Or val = "" Then
                        Me.CurrentComparison.ComparisonEntries(theVarEntry).Grouping = val
                        HasNotSaved = True 'AJW*** was not saving auto on changes to comp grid
                    Else
                        Me.CurrentComparison.ComparisonEntries(theVarEntry).Grouping = val
                        HasNotSaved = True 'AJW*** was not saving auto on changes to comp grid
                        Return "The number of comma-separated items in the aligned rendering and groupings fields do not match!" 'AJW***
                    End If
                Case 5
                    Me.CurrentComparison.ComparisonEntries(theVarEntry).Notes = val
                    HasNotSaved = True 'AJW*** was not saving auto on changes to comp grid
                Case 6
                    Me.CurrentComparison.ComparisonEntries(theVarEntry).Exclude = val
                    HasNotSaved = True 'AJW*** was not saving auto on changes to comp grid
            End Select
            Return ""
        End Function
        Public Function HaveSameNumberOfCommas(ByVal str1 As String, ByVal str2 As String) As Boolean
            Dim cnt1 As Integer = 0
            For Each ch As Char In str1
                If ch = ","c Then
                    cnt1 += 1
                End If
            Next
            Dim cnt2 As Integer = 0
            For Each ch As Char In str2
                If ch = ","c Then
                    cnt2 += 1
                End If
            Next
            Return cnt1 = cnt2
        End Function
        Public Sub RevertToCurrentComparisonsStandardVarietyOrder()
            Dim copyOfSort As New List(Of Variety)
            For Each var As Variety In Me.CurrentComparison.DefaultVarietySort
                copyOfSort.Add(var)
            Next
            Me.CurrentComparison.CurrentVarietySort = copyOfSort
        End Sub
        Public Sub CreateNewComparison(ByVal name As String, ByVal survIndex As Integer)
            Me.CurrentComparison = New Comparison(Me.Surveys(survIndex), name, True)
            Me.Comparisons.Add(Me.CurrentComparison)
        End Sub
        Public Sub RenameCurrentComparison(ByVal newName As String)
            Me.CurrentComparison.Name = newName
        End Sub
        Public Sub DeleteCurrentComparison()
            Me.Comparisons.Remove(Me.CurrentComparison)
            Try
                Me.CurrentComparison = Me.Comparisons(0)
            Catch ex As Exception 'If they deleted the last comparison,
                Me.CurrentComparison = Nothing
            End Try
        End Sub




        Public Sub DoCurrentComparisonsComparisonAnalysis()
            Try
                Me.CurrentComparison.AssociatedAnalysis.Calculate()
            Catch ex As Exception
            End Try
        End Sub

        Public Sub MoveDDRows(ByVal cutIndices As List(Of Integer), ByVal pasteIndex As Integer)
            If cutIndices.Count = 0 Then Return
            'Stupid edge case - if they select all rows and try to paste, it will crash
            If cutIndices.Count = Me.CurrentComparison.AssociatedDegreesOfDifference.UsedCharsList.Count Then Return

            Dim usedChars As List(Of String) = Me.CurrentComparison.AssociatedDegreesOfDifference.UsedCharsList

            Dim movedChars As New List(Of String)
            cutIndices.Sort()
            cutIndices.Reverse()

            For Each cutIndex As Integer In cutIndices
                movedChars.Add(usedChars(cutIndex))
                usedChars.RemoveAt(cutIndex)
                If pasteIndex > cutIndex + 1 Then pasteIndex -= 1
            Next

            For Each movedChar As String In movedChars
                usedChars.Insert(pasteIndex, movedChar)
            Next

            Me.CurrentComparison.AssociatedDegreesOfDifference.CurrentRowIndex = usedChars.IndexOf(movedChars(0))
        End Sub
        Public Sub MoveComparisonAnalysisVariety(ByVal cutIndices As List(Of Integer), ByVal pasteIndex As Integer)
            If cutIndices.Count = 0 Then Return
            'Stupid edge case - if they select all rows and try to paste, it will crash
            If cutIndices.Count = Me.CurrentComparison.CurrentVarietySort.Count Then Return

            Dim movedVarieties As New List(Of Variety)
            cutIndices.Sort()
            cutIndices.Reverse()

            For Each cutIndex As Integer In cutIndices
                movedVarieties.Add(Me.CurrentComparison.CurrentVarietySort(cutIndex))
                Me.CurrentComparison.CurrentVarietySort.RemoveAt(cutIndex)
                If pasteIndex > cutIndex + 1 Then pasteIndex -= 1
            Next

            For Each movedVar As Variety In movedVarieties
                Me.CurrentComparison.CurrentVarietySort.Insert(pasteIndex, movedVar)
            Next
            'movedVarieties.Reverse()
            Me.CurrentComparison.AssociatedAnalysis.CurrentVariety = movedVarieties(0)
        End Sub
        Public Sub MoveComparisonVariety(ByVal cutIndices As List(Of Integer), ByVal pasteIndex As Integer)
            If cutIndices.Count = 0 Then Return
            'Stupid edge case - if they select all rows and try to paste, it will crash
            If cutIndices.Count = Me.CurrentComparison.CurrentVarietySort.Count Then Return

            Dim movedVarieties As New List(Of Variety)
            cutIndices.Sort()
            cutIndices.Reverse()

            For Each cutIndex As Integer In cutIndices
                movedVarieties.Add(Me.CurrentComparison.CurrentVarietySort(cutIndex))
                Me.CurrentComparison.CurrentVarietySort.RemoveAt(cutIndex)
                If pasteIndex > cutIndex + 1 Then pasteIndex -= 1
            Next

            For Each movedVar As Variety In movedVarieties
                Me.CurrentComparison.CurrentVarietySort.Insert(pasteIndex, movedVar)
            Next

            Me.CurrentComparison.CurrentVariety = movedVarieties(0)
        End Sub
        Public Sub CreateNewSurvey(ByVal name As String, ByVal dictIndex As Integer)
            Me.CurrentSurvey = New Survey(Me.Dictionaries(dictIndex), name)
            Me.Surveys.Add(Me.CurrentSurvey)
        End Sub
        Public Sub CreateNewVariety(ByVal name As String)
            If name = "" Then Return
            Me.CurrentSurvey.CurrentVariety = New Variety(Me.CurrentSurvey.AssociatedDictionary, name, True)
            Me.CurrentSurvey.Varieties.Add(Me.CurrentSurvey.CurrentVariety)

            'Add variety to any comparisons using this survey.
            For Each comp As Comparison In Me.Comparisons
                If comp.AssociatedSurvey.Equals(Me.CurrentSurvey) Then
                    comp.CurrentVarietySort.Add(Me.CurrentSurvey.CurrentVariety)
                    comp.DefaultVarietySort.Add(Me.CurrentSurvey.CurrentVariety)
                    For Each gl As Gloss In comp.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses
                        comp.ComparisonEntries.Add(Me.CurrentSurvey.CurrentVariety.VarietyEntries(gl), New ComparisonEntry(Me.CurrentSurvey.CurrentVariety.VarietyEntries(gl).Transcription)) 'Copy the transcription into the aligned rendering field.
                    Next
                End If
            Next
        End Sub
        Public Sub RenameCurrentSurvey(ByVal newName As String)
            Me.CurrentSurvey.Name = newName
        End Sub
        Public Sub RenameCurrentVariety(ByVal newName As String)
            Me.CurrentSurvey.CurrentVariety.Name = newName
        End Sub
        Public Sub DeleteCurrentSurvey()
            Me.DeleteThisSurvey(Me.CurrentSurvey)
        End Sub
        Public Sub DeleteThisSurvey(ByRef surv As Survey)
            Me.Surveys.Remove(Me.CurrentSurvey)

            'Also delete all comparisons that use this survey.
            Dim comparisonsAffected As New List(Of Comparison)
            For Each comp As Comparison In Me.Comparisons
                If comp.AssociatedSurvey.Equals(Me.CurrentSurvey) Then
                    comparisonsAffected.Add(comp)
                End If
            Next
            For Each comp As Comparison In comparisonsAffected
                Me.Comparisons.Remove(comp)
                Try
                    Me.CurrentComparison = Me.Comparisons(0)
                Catch ex As Exception
                    Me.CurrentComparison = Nothing
                End Try
            Next
            Try
                Me.CurrentSurvey = Me.Surveys(0)
            Catch ex As Exception 'If they deleted the last survey,
                Me.CurrentSurvey = Nothing
            End Try
        End Sub
        Public Sub DeleteCurrentVariety()
            'IS THIS COMPLETE?

            'Delete it from all comparisons
            For Each comp As Comparison In Me.Comparisons
                If comp.AssociatedSurvey.Equals(Me.CurrentSurvey) Then
                    comp.CurrentVarietySort.Remove(Me.CurrentSurvey.CurrentVariety)
                    If comp.CurrentVariety.Equals(Me.CurrentSurvey.CurrentVariety) Then comp.CurrentVariety = comp.CurrentVarietySort(0)
                End If
            Next

            'Delete it from the default and current sort of all comparisons that use this survey
            For Each comp As Comparison In Me.Comparisons
                If comp.AssociatedSurvey.Equals(Me.CurrentSurvey) Then

                    Try
                        comp.DefaultVarietySort.Remove(Me.CurrentSurvey.CurrentVariety)
                    Catch ex As Exception 'Only remove the variety if we find it in the sort
                    End Try
                    Try
                        comp.CurrentVarietySort.Remove(Me.CurrentSurvey.CurrentVariety)
                    Catch ex As Exception
                    End Try
                End If
            Next

            'Delete it from the current survey
            Me.CurrentSurvey.Varieties.Remove(Me.CurrentSurvey.CurrentVariety)

            Try
                Me.CurrentSurvey.CurrentVariety = Me.CurrentSurvey.Varieties(0)
            Catch ex As Exception 'If they deleted the last variety,
                Me.CurrentSurvey.CurrentVariety = Nothing
            End Try
        End Sub
        Public Sub UpdateTranscriptionValue(ByVal transRow As Integer, ByVal transCol As Integer, ByVal val As String)
            Select Case transCol
                Case 1
                    'Strip spaces between the synonyms around the commas; this preserves the forms when we rotate them later in the comparison grid.
                    Dim synonyms As String() = Split(val, ",")
                    Dim cleanedTrans As String = ""
                    For i As Integer = 0 To synonyms.Length - 2
                        cleanedTrans &= synonyms(i).Trim() & ","
                    Next
                    cleanedTrans &= synonyms(synonyms.Length() - 1).Trim()

                    Me.CurrentSurvey.CurrentVariety.VarietyEntries(Me.CurrentSurvey.AssociatedDictionary.CurrentSort.Glosses(transRow)).Transcription = cleanedTrans
                    For Each comp As Comparison In Me.Comparisons
                        Try
                            comp.ComparisonEntries(Me.CurrentSurvey.CurrentVariety.VarietyEntries(Me.CurrentSurvey.AssociatedDictionary.CurrentSort.Glosses(transRow))).AlignedRendering = cleanedTrans
                            'Beep()
                            'Beep()
                            'comp.ComparisonEntries(Me.CurrentSurvey.CurrentVariety.VarietyEntries(Me.CurrentSurvey.AssociatedDictionary.CurrentSort.Glosses(transRow))).Grouping = ""
                        Catch ex As Exception
                        End Try
                    Next
                Case 2
                    Me.CurrentSurvey.CurrentVariety.VarietyEntries(Me.CurrentSurvey.AssociatedDictionary.CurrentSort.Glosses(transRow)).PluralFrame = val
                Case 3
                    Me.CurrentSurvey.CurrentVariety.VarietyEntries(Me.CurrentSurvey.AssociatedDictionary.CurrentSort.Glosses(transRow)).Notes = val
            End Select
        End Sub
        Public Sub MakeCrashRecoveryBackup()
            Try
                Me.MakeBackup("!Crash Recovery -- " & Mid(Me.filename, Me.filename.LastIndexOf("\"c) + 2))
            Catch
            End Try
        End Sub

        Public Sub MakeSaveBackup()
            Dim currentFileName As String = Mid(Me.filename, Me.filename.LastIndexOf("\"c) + 2)

            Me.MakeBackup(Mid(currentFileName, 1, currentFileName.Length - 4) & " -- " & Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".wsv")
        End Sub

        Public Sub MakeBackup(ByVal backupFileName As String)
            Dim writer As StreamWriter = Nothing
            Try
                'Copy the current database into the backups folder before writing over it.
                Dim backupPath As String = Mid(Me.filename, 1, Me.filename.LastIndexOf("\"c) + 1) & "Backups\"

                If Not IO.Directory.Exists(backupPath) Then
                    IO.Directory.CreateDirectory(backupPath)
                End If
                '                                           V--Find the filename without the path------------V
                writer = New StreamWriter(backupPath & backupFileName, False, System.Text.Encoding.UTF8)
                Me.WriteFileUsingWriter2(writer)
                writer.Close()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical)
                If writer IsNot Nothing Then writer.Close()
            End Try
        End Sub



        Public Sub CreateNewDictionary(ByVal name As String)
            Me.CurrentDictionary = New Dictionary(name)
            Me.Dictionaries.Add(Me.CurrentDictionary)
            Me.CurrentDictionary.CurrentSort = New Sort("Elicitation Order")
            Me.CurrentDictionary.Sorts.Add(Me.CurrentDictionary.CurrentSort)
        End Sub
        Public Sub DuplicateCurrentDictionary(ByVal newName As String)
            Me.CurrentDictionary = Me.CurrentDictionary.Copy(newName)
            Me.Dictionaries.Add(Me.CurrentDictionary)
        End Sub
        Public Sub CreateNewDictionarySort(ByVal name As String)
            Dim sourceSortGlosses As New List(Of Gloss)(Me.CurrentDictionary.CurrentSort.Glosses)
            Me.CurrentDictionary.CurrentSort = New Sort(name)
            Me.CurrentDictionary.Sorts.Add(Me.CurrentDictionary.CurrentSort)
            Me.CurrentDictionary.CurrentSort.Glosses = sourceSortGlosses
        End Sub
        Public Sub DeleteCurrentDictionarySort()
            Me.CurrentDictionary.Sorts.Remove(Me.CurrentDictionary.CurrentSort)
            Me.CurrentDictionary.CurrentSort = Me.CurrentDictionary.Sorts(0)
        End Sub

#Region "Search"
        Public Function GetNextPositionSub(ByVal objIndex As Integer, ByVal rowIndex As Integer, ByVal colIndex As Integer, ByVal objCount As Integer, ByVal rowCount As Integer, ByVal colCount As Integer, ByVal advanceObject As Boolean) As CellAddress
            Dim ca As New CellAddress(objIndex, rowIndex, colIndex)
            ca.ColIndex += 1
            If ca.ColIndex > colCount - 1 Then
                ca.ColIndex = 0
                ca.RowIndex += 1
                If ca.RowIndex > rowCount - 1 Then
                    ca.RowIndex = 0

                    If advanceObject Then
                        ca.ObjIndex += 1
                        If ca.ObjIndex > objCount - 1 Then
                            ca.ObjIndex = 0
                        End If
                    End If
                End If
            End If
            Return ca
        End Function
        Public Function GetPreviousPositionSub(ByVal objIndex As Integer, ByVal rowIndex As Integer, ByVal colIndex As Integer, ByVal objCount As Integer, ByVal rowCount As Integer, ByVal colCount As Integer, ByVal advanceObject As Boolean) As CellAddress
            Dim ca As New CellAddress(objIndex, rowIndex, colIndex)
            ca.ColIndex -= 1
            If ca.ColIndex < 0 Then
                ca.ColIndex = colCount - 1
                ca.RowIndex -= 1
                If ca.RowIndex < 0 Then
                    ca.RowIndex = rowCount - 1

                    If advanceObject Then
                        ca.ObjIndex -= 1
                        If ca.ObjIndex < 0 Then
                            ca.ObjIndex = objCount - 1
                        End If
                    End If
                End If
            End If
            Return ca
        End Function
        Public Function GetNextPosition(ByVal kind As SearchType, ByVal objIndex As Integer, ByVal rowIndex As Integer, ByVal colIndex As Integer) As CellAddress
            Select Case kind
                Case SearchType.DICTIONARY : Return GetNextPositionSub(objIndex, rowIndex, colIndex, 1, Me.CurrentDictionary.CurrentSort.Glosses.Count, GlossDictionaryGridColCount, False)
                Case SearchType.SURVEY : Return GetNextPositionSub(objIndex, rowIndex, colIndex, Me.CurrentSurvey.Varieties.Count, Me.CurrentSurvey.AssociatedDictionary.CurrentSort.Glosses.Count, VarietyGridColCount, True)
                Case SearchType.COMPARISON_GLOSS : Return GetNextPositionSub(objIndex, rowIndex, colIndex, Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses.Count, Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses.Count, ComparisonGlossGridColCount, False)
                Case SearchType.COMPARISON : Return GetNextPositionSub(objIndex, rowIndex, colIndex, Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses.Count, Me.CurrentComparison.AssociatedSurvey.Varieties.Count, ComparisonGridColCount, True)
                Case SearchType.COGNATE_STRENGTHS : Return GetNextPositionSub(objIndex, rowIndex, colIndex, 1, Me.CurrentComparison.COMPASSCalculations.DisplayedGlosses.Count, CognateStrengthsGridColCount, False)
                Case Else : Throw New ArgumentOutOfRangeException
            End Select
        End Function
        Public Function GetPreviousPosition(ByVal kind As SearchType, ByVal objIndex As Integer, ByVal rowIndex As Integer, ByVal colIndex As Integer) As CellAddress
            Select Case kind
                Case SearchType.DICTIONARY : Return GetPreviousPositionSub(objIndex, rowIndex, colIndex, 1, Me.CurrentDictionary.CurrentSort.Glosses.Count, GlossDictionaryGridColCount, False)
                Case SearchType.SURVEY : Return GetPreviousPositionSub(objIndex, rowIndex, colIndex, Me.CurrentSurvey.Varieties.Count, Me.CurrentSurvey.AssociatedDictionary.CurrentSort.Glosses.Count, VarietyGridColCount, True)
                Case SearchType.COMPARISON_GLOSS : Return GetPreviousPositionSub(objIndex, rowIndex, colIndex, Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses.Count, Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses.Count, ComparisonGlossGridColCount, False)
                Case SearchType.COMPARISON : Return GetPreviousPositionSub(objIndex, rowIndex, colIndex, Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses.Count, Me.CurrentComparison.AssociatedSurvey.Varieties.Count, ComparisonGridColCount, True)
                Case SearchType.COGNATE_STRENGTHS : Return GetPreviousPositionSub(objIndex, rowIndex, colIndex, 1, Me.CurrentComparison.COMPASSCalculations.DisplayedGlosses.Count, CognateStrengthsGridColCount, False)
                Case Else : Throw New ArgumentOutOfRangeException
            End Select
        End Function
        Public Function SearchNext(ByVal kind As SearchType, ByVal searchText As String, ByVal startObjIndex As Integer, ByVal startRowIndex As Integer, ByVal startColIndex As Integer) As CellAddress
            Dim timesSeenSecondToLastRow As Integer = 0
            Dim secondToLastAddress As New CellAddress(startObjIndex, startRowIndex - 1, startColIndex)
            If secondToLastAddress.RowIndex < 0 Then secondToLastAddress.RowIndex = 0

            'Return the location of the next match, or do nothing if there is no next match or the only match is the current location.
            Dim curAddress As New CellAddress(startObjIndex, startRowIndex, startColIndex)
            Do
                curAddress = GetNextPosition(kind, curAddress.ObjIndex, curAddress.RowIndex, curAddress.ColIndex)
                If IsMatch(kind, searchText, curAddress.ObjIndex, curAddress.RowIndex, curAddress.ColIndex) Then
                    Return curAddress
                End If
                If curAddress.ObjIndex = secondToLastAddress.ObjIndex And curAddress.RowIndex = secondToLastAddress.RowIndex And curAddress.ColIndex = secondToLastAddress.ColIndex Then
                    timesSeenSecondToLastRow += 1
                End If
            Loop While (Not (curAddress.ObjIndex = startObjIndex And curAddress.RowIndex = startRowIndex And curAddress.ColIndex = startColIndex)) And timesSeenSecondToLastRow < 2
            Return Nothing
        End Function
        Public Function SearchPrevious(ByVal kind As SearchType, ByVal searchText As String, ByVal startObjIndex As Integer, ByVal startRowIndex As Integer, ByVal startColIndex As Integer) As CellAddress
            Dim timesSeenSecondToLastRow As Integer = 0
            Dim secondToLastAddress As New CellAddress(startObjIndex, startRowIndex - 1, startColIndex)
            If secondToLastAddress.RowIndex < 0 Then secondToLastAddress.RowIndex = 0

            Dim curAddress As New CellAddress(startObjIndex, startRowIndex, startColIndex)
            Do
                curAddress = GetPreviousPosition(kind, curAddress.ObjIndex, curAddress.RowIndex, curAddress.ColIndex)
                If IsMatch(kind, searchText, curAddress.ObjIndex, curAddress.RowIndex, curAddress.ColIndex) Then Return curAddress
                If curAddress.ObjIndex = secondToLastAddress.ObjIndex And curAddress.RowIndex = secondToLastAddress.RowIndex And curAddress.ColIndex = secondToLastAddress.ColIndex Then
                    timesSeenSecondToLastRow += 1
                End If
            Loop While (Not (curAddress.ObjIndex = startObjIndex And curAddress.RowIndex = startRowIndex And curAddress.ColIndex = startColIndex)) And timesSeenSecondToLastRow < 2
            Return Nothing
        End Function
        Public Function SearchReplace(ByVal kind As SearchType, ByVal findText As String, ByVal replaceText As String, ByVal startObjIndex As Integer, ByVal startRowIndex As Integer, ByVal startColIndex As Integer) As CellAddress
            Dim curAddress As New CellAddress(startObjIndex, startRowIndex, startColIndex)
            'If the current location is a match, do a replace and advance to the next match if it exists.
            If IsMatch(kind, findText, curAddress.ObjIndex, curAddress.RowIndex, curAddress.ColIndex) Then
                If IsReplaceableCol(kind, curAddress.ColIndex) Then DoReplace(kind, findText, replaceText, startObjIndex, startRowIndex, startColIndex)

                'If the search after the replace comes to a nonwriteable cell, keep searching until we find a writable one or it gets back to where it was

                'First search for the next cell
                Dim nextAddress As CellAddress = SearchNext(kind, findText, startObjIndex, startRowIndex, startColIndex)

                'If we haven't found something, we are done
                If nextAddress Is Nothing Then Return Nothing
                'Otherwise make that found thing the starting point and search for the next thing
                Dim startingAddress As New CellAddress(nextAddress.ObjIndex, nextAddress.RowIndex, nextAddress.ColIndex)
                'Look for a writable cell unless we wrap around
                Do
                    If IsReplaceableCol(kind, nextAddress.ColIndex) Then Return nextAddress
                    nextAddress = SearchNext(kind, findText, nextAddress.ObjIndex, nextAddress.RowIndex, nextAddress.ColIndex)
                    If nextAddress.ObjIndex = startingAddress.ObjIndex And nextAddress.RowIndex = startingAddress.RowIndex And nextAddress.ColIndex = startingAddress.ColIndex Then Return Nothing
                Loop
            End If
            Return Nothing
        End Function
        Public Function SearchReplaceAll(ByVal kind As SearchType, ByVal findText As String, ByVal replaceText As String, ByVal startObjIndex As Integer, ByVal startRowIndex As Integer, ByVal startColIndex As Integer) As Integer 'Returns the number of replacements
            Dim startingAddress As CellAddress = SearchNext(kind, findText, startObjIndex, startRowIndex, startColIndex)
            If startingAddress Is Nothing Then Return 0

            Dim firstStartingAddress As New CellAddress(startingAddress.ObjIndex, startingAddress.RowIndex, startingAddress.ColIndex)

            While (Not IsReplaceableCol(kind, startingAddress.ColIndex))
                startingAddress = SearchNext(kind, findText, startingAddress.ObjIndex, startingAddress.RowIndex, startingAddress.ColIndex)
                If startingAddress.ObjIndex = firstStartingAddress.ObjIndex And startingAddress.RowIndex = firstStartingAddress.RowIndex And startingAddress.ColIndex = firstStartingAddress.ColIndex Then
                    Return 0
                End If
            End While

            Dim curAddress As CellAddress = SearchNext(kind, findText, startingAddress.ObjIndex, startingAddress.RowIndex, startingAddress.ColIndex)
            Dim numReplacements As Integer = 0
            While curAddress IsNot Nothing AndAlso (Not (curAddress.ObjIndex = startingAddress.ObjIndex And curAddress.RowIndex = startingAddress.RowIndex And curAddress.ColIndex = startingAddress.ColIndex))
                numReplacements += 1
                curAddress = SearchReplace(kind, findText, replaceText, curAddress.ObjIndex, curAddress.RowIndex, curAddress.ColIndex)
            End While
            If curAddress IsNot Nothing Then
                SearchReplace(kind, findText, replaceText, curAddress.ObjIndex, curAddress.RowIndex, curAddress.ColIndex)
                numReplacements += 1
            End If
            Return numReplacements
        End Function
        Public Sub DoReplace(ByVal kind As SearchType, ByVal findText As String, ByVal replaceText As String, ByVal objIndex As Integer, ByVal rowIndex As Integer, ByVal colIndex As Integer)
            Select Case kind
                Case SearchType.DICTIONARY : DoDictionaryReplace(findText, replaceText, objIndex, rowIndex, colIndex)
                Case SearchType.SURVEY : DoSurveyReplace(findText, replaceText, objIndex, rowIndex, colIndex)
                Case SearchType.COMPARISON_GLOSS : Dim donothing As Integer = 0
                Case SearchType.COMPARISON : DoComparisonReplace(findText, replaceText, objIndex, rowIndex, colIndex)
                Case SearchType.COGNATE_STRENGTHS : Dim donothing As Integer = 0
                Case Else : Throw New ArgumentOutOfRangeException
            End Select
        End Sub
        Public Function IsReplaceableCol(ByVal kind As SearchType, ByVal colIndex As Integer) As Boolean
            Select Case kind
                Case SearchType.DICTIONARY : Return True
                Case SearchType.SURVEY : Return colIndex > 0
                Case SearchType.COMPARISON_GLOSS : Return False
                Case SearchType.COMPARISON : Return colIndex > 2
                Case SearchType.COGNATE_STRENGTHS : Return False
            End Select
        End Function
        Public Sub DoDictionaryReplace(ByVal findText As String, ByVal replaceText As String, ByVal objIndex As Integer, ByVal rowIndex As Integer, ByVal colIndex As Integer)
            Select Case colIndex
                Case 0 : Me.CurrentDictionary.CurrentSort.Glosses(rowIndex).Name = Replace(Me.CurrentDictionary.CurrentSort.Glosses(rowIndex).Name, findText, replaceText)
                Case 1 : Me.CurrentDictionary.CurrentSort.Glosses(rowIndex).Name2 = Replace(Me.CurrentDictionary.CurrentSort.Glosses(rowIndex).Name2, findText, replaceText)
                Case 2 : Me.CurrentDictionary.CurrentSort.Glosses(rowIndex).PartOfSpeech = Replace(Me.CurrentDictionary.CurrentSort.Glosses(rowIndex).PartOfSpeech, findText, replaceText)
                Case 3 : Me.CurrentDictionary.CurrentSort.Glosses(rowIndex).FieldTip = Replace(Me.CurrentDictionary.CurrentSort.Glosses(rowIndex).FieldTip, findText, replaceText)
                Case 4 : Me.CurrentDictionary.CurrentSort.Glosses(rowIndex).Comments = Replace(Me.CurrentDictionary.CurrentSort.Glosses(rowIndex).Comments, findText, replaceText)
            End Select
        End Sub
        Public Sub DoSurveyReplace(ByVal findText As String, ByVal replaceText As String, ByVal objIndex As Integer, ByVal rowIndex As Integer, ByVal colIndex As Integer)
            Dim glosses As List(Of Gloss) = Me.CurrentSurvey.AssociatedDictionary.CurrentSort.Glosses
            Dim thisVar As Variety = Me.CurrentSurvey.Varieties(objIndex)

            Select Case colIndex
                Case 1 : thisVar.VarietyEntries(glosses(rowIndex)).Transcription = Replace(thisVar.VarietyEntries(glosses(rowIndex)).Transcription, findText, replaceText)
                Case 2 : thisVar.VarietyEntries(glosses(rowIndex)).PluralFrame = Replace(thisVar.VarietyEntries(glosses(rowIndex)).PluralFrame, findText, replaceText)
                Case 3 : thisVar.VarietyEntries(glosses(rowIndex)).Notes = Replace(thisVar.VarietyEntries(glosses(rowIndex)).Notes, findText, replaceText)
                Case Else : Throw New ArgumentOutOfRangeException
            End Select
        End Sub
        Public Sub DoComparisonReplace(ByVal findText As String, ByVal replaceText As String, ByVal objIndex As Integer, ByVal rowIndex As Integer, ByVal colIndex As Integer)
            Dim theGloss As Gloss = Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses(objIndex)
            Dim theVariety As Variety = Me.CurrentComparison.CurrentVarietySort(rowIndex)
            Dim theVarEntry As VarietyEntry = theVariety.VarietyEntries(theGloss)

            Select Case colIndex
                Case 3 : Me.CurrentComparison.ComparisonEntries(theVarEntry).AlignedRendering = Replace(Me.CurrentComparison.ComparisonEntries(theVarEntry).AlignedRendering, findText, replaceText)
                Case 4 : Me.CurrentComparison.ComparisonEntries(theVarEntry).Grouping = Replace(Me.CurrentComparison.ComparisonEntries(theVarEntry).Grouping, findText, replaceText)
                Case 5 : Me.CurrentComparison.ComparisonEntries(theVarEntry).Notes = Replace(Me.CurrentComparison.ComparisonEntries(theVarEntry).Notes, findText, replaceText)
                Case 6 : Me.CurrentComparison.ComparisonEntries(theVarEntry).Exclude = Replace(Me.CurrentComparison.ComparisonEntries(theVarEntry).Exclude, findText, replaceText)
                Case Else : Throw New AccessViolationException
            End Select
        End Sub
        Public Function IsMatch(ByVal kind As SearchType, ByVal searchText As String, ByVal objIndex As Integer, ByVal rowIndex As Integer, ByVal colIndex As Integer) As Boolean
            Select Case kind
                Case SearchType.DICTIONARY : Return IsDictionaryMatch(searchText, objIndex, rowIndex, colIndex)
                Case SearchType.SURVEY : Return IsSurveyMatch(searchText, objIndex, rowIndex, colIndex)
                Case SearchType.COMPARISON_GLOSS : Return IsComparisonGlossMatch(searchText, objIndex, rowIndex, colIndex)
                Case SearchType.COMPARISON : Return IsComparisonMatch(searchText, objIndex, rowIndex, colIndex)
                Case SearchType.COGNATE_STRENGTHS : Return IsCognateStrengthsMatch(searchText, objIndex, rowIndex, colIndex)
                Case Else : Throw New ArgumentOutOfRangeException
            End Select
        End Function
        Public Function IsDictionaryMatch(ByVal searchText As String, ByVal objIndex As Integer, ByVal rowIndex As Integer, ByVal colIndex As Integer) As Boolean
            If rowIndex < Me.CurrentDictionary.CurrentSort.Glosses.Count Then 'Then we aren't in the last empty row, which would cause an exception.
                Dim glosses As List(Of Gloss) = Me.CurrentDictionary.CurrentSort.Glosses
                Return glosses(rowIndex).GetByIndex(colIndex).ToLower.Contains(searchText.ToLower)
            Else
                Return False
            End If
        End Function
        Public Function IsSurveyMatch(ByVal searchText As String, ByVal objIndex As Integer, ByVal rowIndex As Integer, ByVal colIndex As Integer) As Boolean
            Dim glosses As List(Of Gloss) = Me.CurrentSurvey.AssociatedDictionary.CurrentSort.Glosses
            Dim thisVar As Variety = Me.CurrentSurvey.Varieties(objIndex)
            Dim thisText As String

            Select Case colIndex
                Case 0 : thisText = glosses(rowIndex).Name
                Case 1 : thisText = thisVar.VarietyEntries(glosses(rowIndex)).Transcription
                Case 2 : thisText = thisVar.VarietyEntries(glosses(rowIndex)).PluralFrame
                Case 3 : thisText = thisVar.VarietyEntries(glosses(rowIndex)).Notes
                Case Else : Throw New ArgumentOutOfRangeException
            End Select
            If thisText Is Nothing Then thisText = ""
            Return thisText.ToLower.Contains(searchText.ToLower)
        End Function
        Public Function IsComparisonGlossMatch(ByVal searchText As String, ByVal objIndex As Integer, ByVal rowIndex As Integer, ByVal colIndex As Integer) As Boolean
            Dim glosses As List(Of Gloss) = Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses
            Return glosses(rowIndex).Name.ToLower.Contains(searchText.ToLower)
        End Function
        Public Function IsComparisonMatch(ByVal searchText As String, ByVal objIndex As Integer, ByVal rowIndex As Integer, ByVal colIndex As Integer) As Boolean
            Dim theGloss As Gloss = Me.CurrentComparison.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses(objIndex)
            Dim theVariety As Variety = Me.CurrentComparison.CurrentVarietySort(rowIndex)
            Dim theVarEntry As VarietyEntry = theVariety.VarietyEntries(theGloss)

            Dim thisText As String

            Select Case colIndex
                Case 0 : thisText = theVariety.Name
                Case 1 : thisText = theVarEntry.Transcription
                Case 2 : thisText = theVarEntry.PluralFrame
                Case 3 : thisText = Me.CurrentComparison.ComparisonEntries(theVarEntry).AlignedRendering
                Case 4 : thisText = Me.CurrentComparison.ComparisonEntries(theVarEntry).Grouping
                Case 5 : thisText = Me.CurrentComparison.ComparisonEntries(theVarEntry).Notes
                Case 6 : thisText = Me.CurrentComparison.ComparisonEntries(theVarEntry).Exclude.ToString
                Case Else : Throw New AccessViolationException
            End Select

            If thisText Is Nothing Then thisText = ""
            Return thisText.ToLower.Contains(searchText.ToLower)
        End Function
        Public Function IsCognateStrengthsMatch(ByVal searchText As String, ByVal objIndex As Integer, ByVal rowIndex As Integer, ByVal colIndex As Integer) As Boolean
            Dim entry As COMPASSGlossEntry = Me.CurrentComparison.COMPASSCalculations.DisplayedGlosses(rowIndex)
            Dim thisText As String

            Select Case colIndex
                Case 0 : thisText = entry.Form
                Case 1 : thisText = entry.PaddedForm1
                Case 2 : thisText = entry.PaddedForm2
                Case 3 : thisText = entry.AverageStrength.ToString
                Case Else : Throw New ArgumentOutOfRangeException
            End Select

            Return thisText.ToLower.Contains(searchText.ToLower)
        End Function
#End Region

#Region "Exporting"
        Public Function ExportCurrentDictionaryToExcel() As Boolean
            Return Me.ExportDataTableToExcel(Me.WriteToDataSetForDictionaryExport())
        End Function
        Public Function ExportCurrentDictionaryToCSV() As Boolean
            Return Me.ExportDataTableToCSV(Me.WriteToDataSetForDictionaryExport())
        End Function
        Public Function ExportCurrentSurveyToExcel() As Boolean
            Return Me.ExportDataTableToExcel(Me.WriteToDataSetForSurveyExport())
        End Function
        Public Function ExportCurrentSurveyToCSV() As Boolean
            Return Me.ExportDataTableToCSV(Me.WriteToDataSetForSurveyExport())
        End Function
        Public Function ExportCurrentComparisonToExcel() As Boolean
            Return Me.ExportDataTableToExcel(Me.WriteToDataSetForComparisonExport())
        End Function
        Public Function ExportCurrentComparisonToCSV() As Boolean
            Return Me.ExportDataTableToCSV(Me.WriteToDataSetForComparisonExport())
        End Function
        Public Function ExportCurrentComparisonAnalysisToExcel() As Boolean
            Return Me.ExportDataTableToExcel(Me.WriteToDataSetForComparisonAnalysisExport())
        End Function
        Public Function ExportCurrentComparisonAnalysisToCSV() As Boolean
            Return Me.ExportDataTableToCSV(Me.WriteToDataSetForComparisonAnalysisExport())
        End Function
        Public Function ExportCurrentDDToExcel() As Boolean
            Return Me.ExportDataTableToExcel(Me.WriteToDataSetForDDExport())
        End Function
        Public Function ExportCurrentDDToCSV() As Boolean
            Return Me.ExportDataTableToCSV(Me.WriteToDataSetForDDExport())
        End Function
        Public Function ExportCurrentPhonoStatsToExcel() As Boolean
            Return Me.ExportDataTableToExcel(Me.WriteToDataSetForPhonoStatsExport())
        End Function
        Public Function ExportCurrentPhonoStatsToCSV() As Boolean
            Return Me.ExportDataTableToCSV(Me.WriteToDataSetForPhonoStatsExport())
        End Function
        Public Function ExportCurrentCOMPASSToExcel() As Boolean
            Return Me.ExportDataTableToExcel(Me.WriteToDataSetForCOMPASSExport())
        End Function
        Public Function ExportCurrentCOMPASSToCSV() As Boolean
            Return Me.ExportDataTableToCSV(Me.WriteToDataSetForCOMPASSExport())
        End Function
        Public Function ExportDataTableToExcel(ByVal dt As DataTable) As Boolean


            Dim frmSaveDialog As New SaveFileDialog

            frmSaveDialog.Title = "Enter Name for New Excel Export File"
            frmSaveDialog.DefaultExt = "xls"
            frmSaveDialog.Filter = "Excel Spreadsheet File (.xls) | *.xls"

            If Not frmSaveDialog.ShowDialog = DialogResult.OK Then Return False

            'Begin Excel export
            Dim excelApp As Excel.Application = Nothing
            Dim excelWB As Excel.Workbook = Nothing
            Dim ws As Excel.Worksheet = Nothing
            Try
                excelApp = New Excel.Application
                excelApp.EnableEvents = False
                excelWB = excelApp.Workbooks.Add()
                While excelWB.Sheets.Count > 1
                    CType(excelWB.Sheets(1), Excel.Worksheet).Delete()
                End While
                'CType(excelWB.Sheets(1), Excel.Worksheet).Delete()
                'CType(excelWB.Sheets(1), Excel.Worksheet).Delete()

                ws = CType(excelWB.Sheets(1), Excel.Worksheet)

                ws.Name = dt.TableName
                Dim data(dt.Rows.Count, dt.Columns.Count) As String
                For rowIndex As Integer = 0 To dt.Rows.Count - 1
                    For colIndex As Integer = 0 To dt.Columns.Count - 1
                        data(rowIndex, colIndex) = dt.Rows(rowIndex)(colIndex).ToString
                    Next
                Next

                ws.Range("A1").Resize(dt.Rows.Count, dt.Columns.Count).Value = data

                excelWB.SaveAs(frmSaveDialog.FileName,Excel.XlFileFormat.xlExcel8)

            Catch ex As Exception
                'Make sure we can find the Excel DLLs and if not, give a nice error message instead of crashing incomprehensibly.
                If _
                     Not System.IO.File.Exists(System.Windows.Forms.Application.StartupPath & "\Microsoft.Office.Interop.Excel.dll") Or _
                     Not System.IO.File.Exists(System.Windows.Forms.Application.StartupPath & "\Microsoft.Vbe.Interop.dll") Or _
                     Not System.IO.File.Exists(System.Windows.Forms.Application.StartupPath & "\Office.dll") Or _
                     Not System.IO.File.Exists(System.Windows.Forms.Application.StartupPath & "\stdole.dll") Then
                    MsgBox("Cannot find Excel DLLs.  Please find them and put them in the WordSurv executable path or reinstall WordSurv.")
                    'Return False
                    '***AJW 07/20/2011 this code is preventing the system from running on a machine where Excel is installed - what gives?
                End If
                MessageBox.Show("Could not export to Excel: " & ex.Message)
            Finally
                If Not excelApp Is Nothing Then
                    GC.Collect()
                    GC.WaitForPendingFinalizers()
                    GC.Collect()
                    GC.WaitForPendingFinalizers()

                    Marshal.FinalReleaseComObject(ws)

                    If Not excelWB Is Nothing Then
                        excelWB.Close(SaveChanges:=False)
                        Marshal.FinalReleaseComObject(excelWB)
                    End If

                    excelApp.Quit()
                    Marshal.FinalReleaseComObject(excelApp)
                End If
            End Try

            Return True
        End Function
        Public Function ExportDataTableToCSV(ByVal dt As DataTable) As Boolean
            Dim frmSaveDialog As New SaveFileDialog

            frmSaveDialog.Title = "Enter Name for New CSV Export File"
            frmSaveDialog.DefaultExt = "csv"
            frmSaveDialog.Filter = "Comma Separated Values File (.csv) | *.csv"

            If Not frmSaveDialog.ShowDialog = DialogResult.OK Then Return False

            Dim writer As New StreamWriter(frmSaveDialog.FileName, False, System.Text.Encoding.UTF8)

            'First surround everything in quote marks so we don't not escape things.
            'For rowIndex As Integer = 0 To dt.Rows.Count - 1
            'Dim line As String = ""
            'For colIndex As Integer = 0 To dt.Columns.Count - 1
            'dt.Rows(rowIndex)(colIndex) = """" & dt.Rows(rowIndex)(colIndex).ToString & """"
            'Next
            'Next

            'Then actually write it.
            writer.WriteLine(dt.TableName)
            For rowIndex As Integer = 0 To dt.Rows.Count - 1
                Dim line As String = ""
                For colIndex As Integer = 0 To dt.Columns.Count - 1
                    line &= """" & dt.Rows(rowIndex)(colIndex).ToString & ""","
                Next
                writer.WriteLine(line.TrimEnd(","c))
            Next
            writer.WriteLine()

            writer.Close()

            Return True
        End Function
        Public Function WriteToDataSetForDictionaryExport() As DataTable
            Dim row As DataRow

            Dim dictTable As New DataTable("Gloss Dictionary")
            Dim dictCols As DataColumn() = {New DataColumn, New DataColumn, New DataColumn, New DataColumn, New DataColumn}
            dictTable.Columns.AddRange(dictCols)
            Dim dict As Dictionary = Me.CurrentDictionary
            row = dictTable.NewRow
            row(0) = dict.Name
            dictTable.Rows.Add(row)
            row = dictTable.NewRow
            row(0) = "Primary Language" : row(1) = "Secondary Language" : row(2) = "Part of Speech" : row(3) = "Field Tip" : row(4) = "Comments"
            dictTable.Rows.Add(row)
            For Each gl As Gloss In dict.CurrentSort.Glosses
                row = dictTable.NewRow
                For i As Integer = 0 To Gloss.ColumnCount - 1
                    row(i) = gl.GetByIndex(i)
                Next
                dictTable.Rows.Add(row)
            Next
            dictTable.Rows.Add(dictTable.NewRow)

            Return dictTable
        End Function
        Public Function old_WriteToDataSetForSurveyExport() As DataTable
            Dim row As DataRow

            Dim surv As Survey = Me.CurrentSurvey
            Dim survTable As New DataTable("Survey")
            Dim survCols As DataColumn() = {New DataColumn, New DataColumn, New DataColumn, New DataColumn}
            survTable.Columns.AddRange(survCols)

            row = survTable.NewRow
            row(0) = surv.Name
            survTable.Rows.Add(row)

            row = survTable.NewRow
            row(0) = "Using " & surv.AssociatedDictionary.Name
            survTable.Rows.Add(row)

            row = survTable.NewRow
            row(0) = surv.Description
            survTable.Rows.Add(row)
            For Each var As Variety In surv.Varieties
                survTable.Rows.Add(survTable.NewRow)

                row = survTable.NewRow
                row(0) = var.Name
                survTable.Rows.Add(row)

                row = survTable.NewRow
                row(0) = var.Description
                survTable.Rows.Add(row)

                row = survTable.NewRow
                row(0) = "Gloss" : row(1) = "Transcription" : row(2) = "Plural/Frame" : row(3) = "Notes"
                survTable.Rows.Add(row)

                For Each gl As Gloss In surv.AssociatedDictionary.CurrentSort.Glosses
                    Dim tr As VarietyEntry = var.VarietyEntries(gl)
                    row = survTable.NewRow
                    Dim synonyms As String() = Split(tr.Transcription, ",")
                    row(0) = gl.Name : row(1) = synonyms(0) : row(2) = tr.PluralFrame : row(3) = tr.Notes
                    survTable.Rows.Add(row)
                    For i As Integer = 1 To synonyms.Length() - 1
                        row = survTable.NewRow
                        row(0) = "" : row(1) = synonyms(i) : row(2) = "" : row(3) = ""
                        survTable.Rows.Add(row)
                    Next
                Next
            Next
            survTable.Rows.Add(survTable.NewRow)

            Return survTable
        End Function
        Public Function WriteToDataSetForSurveyExport() As DataTable
            Dim row As DataRow

            Dim surv As Survey = Me.CurrentSurvey
            Dim survTable As New DataTable("Survey")
            survTable.Columns.AddRange({New DataColumn, New DataColumn, New DataColumn, New DataColumn})

            row = survTable.NewRow
            row(0) = surv.Name
            survTable.Rows.Add(row)

            row = survTable.NewRow
            row(0) = "Using " & surv.AssociatedDictionary.Name
            survTable.Rows.Add(row)

            row = survTable.NewRow
            row(0) = surv.Description
            survTable.Rows.Add(row)
            Dim heading1 As DataRow = survTable.NewRow
            survTable.Rows.Add(heading1)
            Dim heading2 As DataRow = survTable.NewRow
            heading2(0) = "Gloss"
            survTable.Rows.Add(heading2)
            For Each gl As Gloss In surv.AssociatedDictionary.CurrentSort.Glosses
                row = survTable.NewRow
                Dim col As Integer = 1
                row(0) = gl.Name
                For Each var As Variety In surv.Varieties
                    heading1(col) = var.Name : heading1(col + 1) = var.Description
                    heading2(col) = "Transcription" : heading2(col + 1) = "Plural/Frame" : heading2(col + 2) = "Notes"
                    survTable.Columns.AddRange({New DataColumn, New DataColumn, New DataColumn, New DataColumn})
                    Dim tr As VarietyEntry = var.VarietyEntries(gl)
                    'Dim synonyms As String() = Split(tr.Transcription, ",")
                    row(col) = tr.Transcription
                    row(col + 1) = tr.PluralFrame
                    row(col + 2) = tr.Notes
                    col = col + 3
                    Debug.Print(col)
                Next
                survTable.Rows.Add(row)
            Next
            'For Each var As Variety In surv.Varieties
            '    survTable.Rows.Add(survTable.NewRow)

            '    row = survTable.NewRow
            '    row(0) = var.Name
            '    survTable.Rows.Add(row)

            '    row = survTable.NewRow
            '    row(0) = var.Description
            '    survTable.Rows.Add(row)

            '    row = survTable.NewRow
            '    row(0) = "Gloss" : row(1) = "Transcription" : row(2) = "Plural/Frame" : row(3) = "Notes"
            '    survTable.Rows.Add(row)

            '    For Each gl As Gloss In surv.AssociatedDictionary.CurrentSort.Glosses
            '        Dim tr As VarietyEntry = var.VarietyEntries(gl)
            '        row = survTable.NewRow
            '        Dim synonyms As String() = Split(tr.Transcription, ",")
            '        row(0) = gl.Name : row(1) = synonyms(0) : row(2) = tr.PluralFrame : row(3) = tr.Notes
            '        survTable.Rows.Add(row)
            '        For i As Integer = 1 To synonyms.Length() - 1
            '            row = survTable.NewRow
            '            row(0) = "" : row(1) = synonyms(i) : row(2) = "" : row(3) = ""
            '            survTable.Rows.Add(row)
            '        Next
            '    Next
            'Next
            survTable.Rows.Add(survTable.NewRow)

            Return survTable
        End Function
        Public Function WriteToDataSetForComparisonExport() As DataTable
            Dim row As DataRow

            Dim comp As Comparison = Me.CurrentComparison
            Dim compTable As New DataTable("Comparison")
            Dim compCols As DataColumn() = {New DataColumn, New DataColumn, New DataColumn, New DataColumn, New DataColumn, New DataColumn, New DataColumn}
            compTable.Columns.AddRange(compCols)

            row = compTable.NewRow
            row(0) = comp.Name
            compTable.Rows.Add(row)

            row = compTable.NewRow
            row(0) = "Using " & comp.AssociatedSurvey.Name
            compTable.Rows.Add(row)

            row = compTable.NewRow
            row(0) = comp.Description
            compTable.Rows.Add(row)

            For Each gl As Gloss In comp.AssociatedSurvey.AssociatedDictionary.CurrentSort.Glosses
                compTable.Rows.Add(compTable.NewRow)

                row = compTable.NewRow
                row(0) = gl.Name
                compTable.Rows.Add(row)

                row = compTable.NewRow
                row(0) = "Variety" : row(1) = "Transcription" : row(2) = "Plural/Frame" : row(3) = "AlignedRendering" : row(4) = "Grouping" : row(5) = "Notes" : row(6) = "Exclude"
                compTable.Rows.Add(row)

                For Each var As Variety In comp.CurrentVarietySort
                    Dim tr As VarietyEntry = var.VarietyEntries(gl)
                    Dim compEntry As ComparisonEntry = comp.ComparisonEntries(tr)

                    row = compTable.NewRow
                    Dim transcriptions As String() = Split(tr.Transcription, ",")
                    Dim alignedRenderings As String() = Split(compEntry.AlignedRendering, ",")
                    Dim groupings As String() = Split(compEntry.Grouping, ",")

                    row(0) = var.Name : row(1) = transcriptions(0) : row(2) = tr.PluralFrame : row(3) = alignedRenderings(0) : row(4) = groupings(0) : row(5) = compEntry.Notes : row(6) = compEntry.Exclude
                    compTable.Rows.Add(row)

                    For i As Integer = 1 To transcriptions.Length() - 1
                        Dim group As String
                        Try
                            group = groupings(i)
                        Catch ex As Exception
                            group = ""
                        End Try
                        row = compTable.NewRow
                        row(0) = "" : row(1) = transcriptions(i) : row(2) = "" : row(3) = alignedRenderings(i - 1) : row(4) = group : row(5) = "" : row(6) = "" 'AJW777 made it i-1 instead of i
                        compTable.Rows.Add(row)
                    Next
                Next
            Next

            Return compTable
        End Function
        Public Function WriteToDataSetForComparisonAnalysisExport() As DataTable
            Dim row As DataRow
            Dim comp As Comparison = Me.CurrentComparison

            comp.AssociatedAnalysis.Calculate()
            Dim compAnalysisTable As New DataTable("Comparison Analysis")
            For colIndex As Integer = 0 To comp.CurrentVarietySort.Count 'One more than the number of columns
                compAnalysisTable.Columns.Add(New DataColumn)
            Next
            row = compAnalysisTable.NewRow
            row(0) = comp.Name & " Analysis"
            compAnalysisTable.Rows.Add(row)

            compAnalysisTable.Rows.Add(compAnalysisTable.NewRow)

            row = compAnalysisTable.NewRow
            row(0) = "Total"
            compAnalysisTable.Rows.Add(row)
            row = compAnalysisTable.NewRow
            For colIndex As Integer = 1 To comp.CurrentVarietySort.Count
                row(colIndex) = comp.CurrentVarietySort(colIndex - 1).Name
            Next
            compAnalysisTable.Rows.Add(row)
            For Each var1 As Variety In comp.CurrentVarietySort
                row = compAnalysisTable.NewRow
                row(0) = var1.Name
                Dim colIndex As Integer = 1
                For Each var2 As Variety In comp.CurrentVarietySort
                    row(colIndex) = comp.AssociatedAnalysis.TallyMatrix(var1)(var2).ToString
                    colIndex += 1
                Next
                compAnalysisTable.Rows.Add(row)
            Next

            compAnalysisTable.Rows.Add(compAnalysisTable.NewRow)

            row = compAnalysisTable.NewRow
            row(0) = "Total"
            compAnalysisTable.Rows.Add(row)
            row = compAnalysisTable.NewRow
            For colIndex As Integer = 1 To comp.CurrentVarietySort.Count
                row(colIndex) = comp.CurrentVarietySort(colIndex - 1).Name
            Next
            compAnalysisTable.Rows.Add(row)
            For Each var1 As Variety In comp.CurrentVarietySort
                row = compAnalysisTable.NewRow
                row(0) = var1.Name
                Dim colIndex As Integer = 1
                For Each var2 As Variety In comp.CurrentVarietySort
                    row(colIndex) = comp.AssociatedAnalysis.TotalMatrix(var1)(var2).ToString
                    colIndex += 1
                Next
                compAnalysisTable.Rows.Add(row)
            Next

            compAnalysisTable.Rows.Add(compAnalysisTable.NewRow)

            row = compAnalysisTable.NewRow
            row(0) = "Percent"
            compAnalysisTable.Rows.Add(row)
            row = compAnalysisTable.NewRow
            For colIndex As Integer = 1 To comp.CurrentVarietySort.Count
                row(colIndex) = comp.CurrentVarietySort(colIndex - 1).Name
            Next
            compAnalysisTable.Rows.Add(row)
            For Each var1 As Variety In comp.CurrentVarietySort
                row = compAnalysisTable.NewRow
                row(0) = var1.Name
                Dim colIndex As Integer = 1
                For Each var2 As Variety In comp.CurrentVarietySort
                    row(colIndex) = comp.AssociatedAnalysis.PercentMatrix(var1)(var2).ToString
                    colIndex += 1
                Next
                compAnalysisTable.Rows.Add(row)
            Next

            Return compAnalysisTable
        End Function
        Public Function WriteToDataSetForDDExport() As DataTable
            Dim row As DataRow
            Dim comp As Comparison = Me.CurrentComparison

            comp.AssociatedDegreesOfDifference.CalculateUsedChars()
            comp.AssociatedDegreesOfDifference.DoAnalysis()
            Dim ddTable As New DataTable("Degrees of Difference")
            For colIndex As Integer = 0 To comp.AssociatedDegreesOfDifference.UsedCharsList.Count
                ddTable.Columns.Add(New DataColumn)
            Next
            row = ddTable.NewRow
            row(0) = comp.Name & " Degrees of Difference Grid"
            ddTable.Rows.Add(row)

            ddTable.Rows.Add(ddTable.NewRow)

            row = ddTable.NewRow
            For colIndex As Integer = 1 To comp.AssociatedDegreesOfDifference.UsedCharsList.Count - 1
                row(colIndex) = comp.AssociatedDegreesOfDifference.UsedCharsList(colIndex - 1)
            Next
            ddTable.Rows.Add(row)

            For Each usedChar1 As String In comp.AssociatedDegreesOfDifference.UsedCharsList
                row = ddTable.NewRow
                row(0) = usedChar1
                Dim ddCol As Integer = 1
                For Each usedChar2 As String In comp.AssociatedDegreesOfDifference.UsedCharsList
                    Dim char1AndChar2 As Integer = (AscW(usedChar1) << 16) Or AscW(usedChar2)
                    row(ddCol) = comp.AssociatedDegreesOfDifference.DDs(char1AndChar2).ToString
                    If row(ddCol).ToString = "-1" Then row(ddCol) = ""
                    ddCol += 1
                Next
                ddTable.Rows.Add(row)
            Next
            Return ddTable
        End Function
        Public Function WriteToDataSetForPhonoStatsExport() As DataTable
            Dim row As DataRow
            Dim comp As Comparison = Me.CurrentComparison

            comp.AssociatedDegreesOfDifference.CalculateUsedChars()
            comp.AssociatedDegreesOfDifference.DoAnalysis()
            Dim phonoStatsTable As New DataTable("Phonostatistical Analysis")

            Dim rowCount As Integer
            If comp.AssociatedDegreesOfDifference.UsedCharsList.Count > comp.CurrentVarietySort.Count Then
                rowCount = comp.AssociatedDegreesOfDifference.UsedCharsList.Count
            Else
                rowCount = comp.CurrentVarietySort.Count
            End If
            For colIndex As Integer = 0 To rowCount 'One more than the number of columns
                phonoStatsTable.Columns.Add(New DataColumn)
            Next

            row = phonoStatsTable.NewRow
            row(0) = "DD Number of Correspondences"
            phonoStatsTable.Rows.Add(row)

            row = phonoStatsTable.NewRow
            For colIndex As Integer = 1 To comp.AssociatedDegreesOfDifference.UsedCharsList.Count - 1
                row(colIndex) = comp.AssociatedDegreesOfDifference.UsedCharsList(colIndex - 1)
            Next
            phonoStatsTable.Rows.Add(row)
            For Each usedChar1 As String In comp.AssociatedDegreesOfDifference.UsedCharsList
                row = phonoStatsTable.NewRow
                row(0) = usedChar1
                Dim phCol As Integer = 1
                For Each usedChar2 As String In comp.AssociatedDegreesOfDifference.UsedCharsList
                    Dim char1AndChar2 As Integer = (AscW(usedChar1) << 16) Or AscW(usedChar2)
                    row(phCol) = comp.AssociatedDegreesOfDifference.DDCharCorrespondences(char1AndChar2).ToString
                    If row(phCol).ToString = "-1" Then row(phCol) = ""
                    phCol += 1
                Next
                phonoStatsTable.Rows.Add(row)
            Next


            phonoStatsTable.Rows.Add(phonoStatsTable.NewRow)


            row = phonoStatsTable.NewRow
            row(0) = "DD Summation"
            phonoStatsTable.Rows.Add(row)

            row = phonoStatsTable.NewRow
            For colIndex As Integer = 1 To comp.CurrentVarietySort.Count
                row(colIndex) = comp.CurrentVarietySort(colIndex - 1).Name
            Next
            phonoStatsTable.Rows.Add(row)
            For Each var1 As Variety In comp.CurrentVarietySort
                row = phonoStatsTable.NewRow
                row(0) = var1.Name
                Dim colIndex As Integer = 1
                For Each var2 As Variety In comp.CurrentVarietySort
                    row(colIndex) = comp.AssociatedDegreesOfDifference.DDMatrixDegrees(var1)(var2).ToString
                    colIndex += 1
                Next
                phonoStatsTable.Rows.Add(row)
            Next


            phonoStatsTable.Rows.Add(phonoStatsTable.NewRow)


            row = phonoStatsTable.NewRow
            row(0) = "Correspondence Totals"
            phonoStatsTable.Rows.Add(row)

            row = phonoStatsTable.NewRow
            For colIndex As Integer = 1 To comp.CurrentVarietySort.Count
                row(colIndex) = comp.CurrentVarietySort(colIndex - 1).Name
            Next
            phonoStatsTable.Rows.Add(row)
            For Each var1 As Variety In comp.CurrentVarietySort
                row = phonoStatsTable.NewRow
                row(0) = var1.Name
                Dim colIndex As Integer = 1
                For Each var2 As Variety In comp.CurrentVarietySort
                    row(colIndex) = comp.AssociatedDegreesOfDifference.DDMatrixCorrespondences(var1)(var2).ToString
                    colIndex += 1
                Next
                phonoStatsTable.Rows.Add(row)
            Next


            phonoStatsTable.Rows.Add(phonoStatsTable.NewRow)


            row = phonoStatsTable.NewRow
            row(0) = "Ratio"
            phonoStatsTable.Rows.Add(row)

            row = phonoStatsTable.NewRow
            For colIndex As Integer = 1 To comp.CurrentVarietySort.Count
                row(colIndex) = comp.CurrentVarietySort(colIndex - 1).Name
            Next
            phonoStatsTable.Rows.Add(row)
            For Each var1 As Variety In comp.CurrentVarietySort
                row = phonoStatsTable.NewRow
                row(0) = var1.Name
                Dim colIndex As Integer = 1
                For Each var2 As Variety In comp.CurrentVarietySort
                    row(colIndex) = comp.AssociatedDegreesOfDifference.DDMatrixRatio(var1)(var2).ToString
                    colIndex += 1
                Next
                phonoStatsTable.Rows.Add(row)
            Next

            Return phonoStatsTable
        End Function
        Public Function WriteToDataSetForCOMPASSExport() As DataTable
            Dim row As DataRow
            Dim COMPASSTable As New DataTable("COMPASS")

            Dim comp As Comparison = Me.CurrentComparison

            Dim colCount As Integer
            If comp.COMPASSCalculations.UsedChars.Count > 5 Then
                colCount = comp.COMPASSCalculations.UsedChars.Count
            Else
                colCount = 5
            End If
            For colIndex As Integer = 0 To colCount 'One more than the number of columns
                COMPASSTable.Columns.Add(New DataColumn)
            Next

            row = COMPASSTable.NewRow
            row(0) = "COMPASS"
            COMPASSTable.Rows.Add(row)

            COMPASSTable.Rows.Add(COMPASSTable.NewRow)

            row = COMPASSTable.NewRow
            row(0) = comp.CurrentVarietySort(comp.COMPASSCalculations.CurrentVarietyIndex1)
            COMPASSTable.Rows.Add(row)
            row = COMPASSTable.NewRow
            row(0) = comp.CurrentVarietySort(comp.COMPASSCalculations.CurrentVarietyIndex2)
            COMPASSTable.Rows.Add(row)

            COMPASSTable.Rows.Add(COMPASSTable.NewRow)

            row = COMPASSTable.NewRow
            row(0) = "Counts"
            COMPASSTable.Rows.Add(row)

            row = COMPASSTable.NewRow
            For colIndex As Integer = 1 To comp.COMPASSCalculations.UsedChars.Count - 1
                row(colIndex) = comp.COMPASSCalculations.UsedChars(colIndex - 1)
            Next
            COMPASSTable.Rows.Add(row)
            For Each usedChar1 As String In comp.COMPASSCalculations.UsedChars
                row = COMPASSTable.NewRow
                row(0) = usedChar1
                Dim comCol As Integer = 1
                For Each usedChar2 As String In comp.COMPASSCalculations.UsedChars
                    If comp.COMPASSCalculations.CharPairRecords.ContainsKey(usedChar1 & usedChar2) Then
                        row(comCol) = comp.COMPASSCalculations.CharPairRecords(usedChar1 & usedChar2).Occurences.Count
                    Else
                        row(comCol) = ""
                    End If
                    comCol += 1
                Next
                COMPASSTable.Rows.Add(row)
            Next

            COMPASSTable.Rows.Add(COMPASSTable.NewRow)

            row = COMPASSTable.NewRow
            row(0) = "Strengths"
            COMPASSTable.Rows.Add(row)

            row = COMPASSTable.NewRow
            For colIndex As Integer = 1 To comp.COMPASSCalculations.UsedChars.Count - 1
                row(colIndex) = comp.COMPASSCalculations.UsedChars(colIndex - 1)
            Next
            COMPASSTable.Rows.Add(row)
            For Each usedChar1 As String In comp.COMPASSCalculations.UsedChars
                row = COMPASSTable.NewRow
                row(0) = usedChar1
                Dim comCol As Integer = 1
                For Each usedChar2 As String In comp.COMPASSCalculations.UsedChars
                    If comp.COMPASSCalculations.CharPairRecords.ContainsKey(usedChar1 & usedChar2) Then
                        row(comCol) = comp.COMPASSCalculations.CharPairRecords(usedChar1 & usedChar2).Strength.ToString("F2")
                    Else
                        row(comCol) = ""
                    End If
                    comCol += 1
                Next
                COMPASSTable.Rows.Add(row)
            Next

            COMPASSTable.Rows.Add(COMPASSTable.NewRow)

            COMPASSTable.Rows.Add(COMPASSTable.NewRow)

            row = COMPASSTable.NewRow
            row(0) = "Strengths Summary"
            COMPASSTable.Rows.Add(row)

            row = COMPASSTable.NewRow
            row(0) = "Gloss" : row(1) = comp.CurrentVarietySort(comp.COMPASSCalculations.CurrentVarietyIndex1) : row(2) = comp.CurrentVarietySort(comp.COMPASSCalculations.CurrentVarietyIndex2) : row(3) = "Strength"
            COMPASSTable.Rows.Add(row)
            For Each strEntry As COMPASSGlossEntry In comp.COMPASSCalculations.GlossValues
                row = COMPASSTable.NewRow
                row(0) = strEntry.Form
                row(1) = strEntry.PaddedForm1
                row(2) = strEntry.PaddedForm2
                row(3) = strEntry.AverageStrength.ToString("F2")
                COMPASSTable.Rows.Add(row)
            Next
            COMPASSTable.Rows.Add(COMPASSTable.NewRow)

            Return COMPASSTable
        End Function
#End Region

    End Class

    Private dontUseThisGlobalUnlessYouAreTheGlossComparer As Integer
    Private Class GlossComparer
        Implements IComparer(Of Gloss)

        Public Function Compare(ByVal x As Gloss, ByVal y As Gloss) As Integer Implements System.Collections.Generic.IComparer(Of WordSurv7.DataObjects.Gloss).Compare
            Return String.Compare(x.GetByIndex(dontUseThisGlobalUnlessYouAreTheGlossComparer), y.GetByIndex(dontUseThisGlobalUnlessYouAreTheGlossComparer))
        End Function
    End Class

    Public Function GroupsMatch(ByVal group1 As String, ByVal group2 As String) As Boolean
        'If any of the letters match, the groups match.  How convenient!
        For Each char1 As Char In group1
            If char1 = ","c Or char1 = " "c Then Continue For
            For Each char2 As Char In group2
                If char2 = ","c Or char1 = " "c Then Continue For
                If char1 = char2 Then Return True
            Next
        Next
        Return False
    End Function

End Module
