Imports System.IO
Imports System.Reflection

Public Class Preferences
    Public GlossDictionaryGridNameWidth As Integer = 100
    Public GlossDictionaryGridName2Width As Integer = 100
    Public GlossDictionaryGridPartOfSpeechWidth As Integer = 100
    Public GlossDictionaryGridFieldTipWidth As Integer = 100
    Public GlossDictionaryGridCommentsWidth As Integer = 100

    Public VarietyGridNameWidth As Integer = 100
    Public VarietyGridTranscriptionWidth As Integer = 100
    Public VarietyGridPluralFrameWidth As Integer = 100
    Public VarietyGridNotesWidth As Integer = 100

    Public ComparisonGlossGridNameWidth As Integer = 100

    Public ComparisonGridVarietyWidth As Integer = 100
    Public ComparisonGridTranscriptionWidth As Integer = 100
    Public ComparisonGridPluralFrameWidth As Integer = 100
    Public ComparisonGridAlignedRenderingWidth As Integer = 100
    Public ComparisonGridGroupingWidth As Integer = 100
    Public ComparisonGridNotesWidth As Integer = 100
    Public ComparisonGridExcludeWidth As Integer = 100

    Public CognateStrengthsGridGlossWidth As Integer = 100
    Public CognateStrengthsGridForm1Width As Integer = 100
    Public CognateStrengthsGridForm2Width As Integer = 100
    Public CognateStrengthsGridStrengthWidth As Integer = 100

    Public LastOpenedDatabase As String = ""
    Public CurrentTab As Integer = 0

    Public ApplicationWidth As Integer = 800
    Public ApplicationHeight As Integer = 600
    Public ApplicationX As Integer = 0
    Public ApplicationY As Integer = 0
    Public ApplicationIsMaximized As Boolean = False

    Public DictionaryPaneWidth As Integer = 190
    Public SurveyPaneWidth As Integer = 200
    Public ComparisonPaneWidth As Integer = 145
    Public COMPASSPaneWidth As Integer = 495

    Public RecentDatabase0 As String = ""
    Public RecentDatabase1 As String = ""
    Public RecentDatabase2 As String = ""
    Public RecentDatabase3 As String = ""
    Public RecentDatabase4 As String = ""
    Public RecentDatabase5 As String = ""
    Public RecentDatabase6 As String = ""
    Public RecentDatabase7 As String = ""
    Public RecentDatabase8 As String = ""
    Public RecentDatabase9 As String = ""
    Public RecentDatabase10 As String = ""

    'Public PrimaryFontFace As String = "Microsoft Sans Serif"
    'Public PrimaryFontSize As Single = 8
    'Public SecondaryFontFace As String = "Microsoft Sans Serif"
    'Public SecondaryFontSize As Single = 8
    'Public TranscriptionFontFace As String = "Microsoft Sans Serif"
    'Public TranscriptionFontSize As Single = 8
    Public ComparisonAnalysisColumnWidth As Integer = 80

    Public ComparisonAnalysisMode As String = "Tally"
    Public PhonoStatsAnalysisMode As String = "DDNumberOfCorrespondences"

    'Public PrimaryLanguage As String = "Primary Lang"
    'Public SecondaryLanguage As String = "Secondary Lang"

    Public NumberOfBackups As Integer = 10
    Public MaxUndos As Integer = 30

    Public COMPASSUpper As Integer = 15
    Public COMPASSLower As Integer = 2
    Public COMPASSBottom As Integer = 1

    Public COMPASSVariety1Index As Integer = 0
    Public COMPASSVariety2Index As Integer = 1

    Public CorrespondenceDisplayMode As String = "ShowCounts"

    Public Sub New()
        'If the settings.dat file exists, load in its values.
        If File.Exists(AppDomain.CurrentDomain.BaseDirectory & "settings.dat") Then
            Dim reader As StreamReader = File.OpenText(AppDomain.CurrentDomain.BaseDirectory & "settings.dat")
            Dim line As String
            Dim parts(1) As String

            While reader.Peek <> -1 'Read until the end of the file.
                line = reader.ReadLine()
                parts = line.Split(","c) : parts(0).Trim() : parts(1).Trim()
                Try
                    CallByName(Me, parts(0), CallType.Set, parts(1))
                Catch ex As Exception
                End Try

            End While
            reader.Close()
        End If
        If Me.ApplicationX < 0 Or Me.ApplicationX > 3000 Then
            Me.ApplicationX = 0
        End If
        If Me.ApplicationY < 0 Or Me.ApplicationY > 3000 Then
            Me.ApplicationY = 0
        End If
    End Sub

    Public Sub save() 'Dump everything out to a file.
        Dim writer As StreamWriter = File.CreateText(AppDomain.CurrentDomain.BaseDirectory & "settings.dat")

        For Each prop As FieldInfo In Me.GetType.GetFields()
            Dim propName As String = prop.Name
            Dim propVal As Object = CallByName(Me, propName, CallType.Get)
            'If propName = "ApplicationX" Or propName = "ApplicationY" Then
            '    Dim value As Integer = Integer.Parse(propVal.ToString)
            '    If value < 0 Then
            '        writer.WriteLine(propName & "," & 0)
            '    End If
            'Else
            writer.WriteLine(propName & "," & propVal.ToString)
            'End If
        Next
        writer.Close()
    End Sub

End Class
