Imports System.Data.OleDb
Module WordSurv6Import
    Public DBConnectionString As String
    Public DBConnection As New OleDbConnection
    Public DBConnection2 As New OleDbConnection
    Public DBConnection3 As New OleDbConnection
    Public DBConnection4 As New OleDbConnection
    Public DBConnection5 As New OleDbConnection
    Public DBConnection6 As New OleDbConnection
    Public DBConnection7 As New OleDbConnection
    Public DBCommand As New OleDbCommand
    Public DBCommand2 As New OleDbCommand
    Public DBCommand3 As New OleDbCommand
    Public DBCommand4 As New OleDbCommand
    Public DBCommand5 As New OleDbCommand
    Public DBCommand6 As New OleDbCommand
    Public DBCommand7 As New OleDbCommand
    Public SQL As String
    Public Class TransEntry
        Public Transcription As String
        Public Notes As String
        Public GlossID As Integer
        Public GlossName As String 'AJW***
    End Class
    Public Class CompEntry
        Public Gloss As String
        Public AlignedRendering As String
        Public Grouping As String
        Public Notes As String
        Public Exclude As Boolean
        Public GlossID As Integer
        Public Transcription As String
    End Class

    Declare Sub Sleep Lib "kernel32.dll" (ByVal Milliseconds As Integer)

    Public Sub ConvertWS6ToWS7(fromMDB As String, toWSV As String)
        Dim cursurv As Integer = -1
        Dim curcomp As Integer = -1

        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        DBConnectionString = "Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Registry Path=;Jet OLEDB:Database Locking Mode=0;Data Source=" & fromMDB & ";Jet OLEDB:Engine Type=5;Provider=""Microsoft.Jet.OLEDB.4.0"";Jet OLEDB:System database=;Jet OLEDB:SFP=False;persist security info=False;Extended Properties=;Mode=Share Deny None;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Create System Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;User ID=Admin;Jet OLEDB:Global Bulk Transactions=1"
        Try
            DBConnection = New OleDbConnection(DBConnectionString)
            DBConnection.Open()
            DBCommand.Connection = DBConnection

            DBConnection2 = New OleDbConnection(DBConnectionString)
            DBConnection2.Open()
            DBCommand2.Connection = DBConnection2

            DBConnection3 = New OleDbConnection(DBConnectionString)
            DBConnection3.Open()
            DBCommand3.Connection = DBConnection3

            DBConnection4 = New OleDbConnection(DBConnectionString)
            DBConnection4.Open()
            DBCommand4.Connection = DBConnection4

            DBConnection5 = New OleDbConnection(DBConnectionString)
            DBConnection5.Open()
            DBCommand5.Connection = DBConnection5

            DBConnection6 = New OleDbConnection(DBConnectionString)
            DBConnection6.Open()
            DBCommand6.Connection = DBConnection6

            DBConnection7 = New OleDbConnection(DBConnectionString)
            DBConnection7.Open()
            DBCommand7.Connection = DBConnection7
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try






        Dim outputFile As New IO.StreamWriter(toWSV, False, System.Text.Encoding.UTF8)
        Dim counter As Integer = 0
        Dim temp As String
        Dim temp2 As String
        Dim VarCount As Integer = 0
        'Embedded ' in text can cause a problem - replace in all tables with a null string


        'START OF DICTIONARIES
        'First, remove any glosses from GLOSS that do not appear in any gloss dictionary
        'Try
        '    DBCommand2.CommandText = "DROP TABLE tempGlossesNotUsed;"
        '    Dim dropWordlistResult As Integer = DBCommand2.ExecuteNonQuery()
        '    Application.DoEvents()
        '    Application.DoEvents()
        'Catch ex As Exception
        '    MsgBox("Failed deleting tempGlossesNotUsed with error - " & ex.Message)
        'End Try

        'SQL = "SELECT Gloss.GlossID, Gloss.Name INTO tempGlossesNotUsed FROM Gloss LEFT JOIN GlossDictionary ON Gloss.GlossID = GlossDictionary.GlossID WHERE (((GlossDictionary.GlossID) Is Null));"
        'DBCommand6.CommandText = SQL
        'Try
        '    DBCommand6.ExecuteNonQuery()
        'Catch ex As Exception
        '    MsgBox("Failed creating tempGlossesNotUsed with error - " & ex.Message)
        'End Try

        'Application.DoEvents()
        'Application.DoEvents()
        'SQL = "DELETE Gloss.GlossID, * FROM(Gloss) WHERE (((Gloss.GlossID) In (SELECT GlossID FROM tempGlossesNotUsed)));"
        'DBCommand7.CommandText = SQL
        'Try
        '    DBCommand7.ExecuteNonQuery()
        'Catch ex As Exception
        '    MsgBox("Failed deleting from gloss with error - " & ex.Message)
        'End Try



        'MUST HANDLE DUPLICATE ENTRIES (GLOSSES THAT HAVE THE SAME NAME BUT DIFFERENT ID, e.g. fly(noun) and fly(verb)) #AJW2013-02-10
        'DROP tempDictionary table if it exists
        Try
            DBCommand2.CommandText = "DROP TABLE tempglossDupes;"
            Dim dropWordlistResult As Integer = DBCommand2.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        SQL = "SELECT Gloss.[Name] INTO tempglossDupes FROM Gloss GROUP BY Gloss.[Name] HAVING(((Count(Gloss.[Name])) > 1)) ORDER BY Count(Gloss.[Name]) DESC;"
        'SQL = "SELECT GlossDictionary.GlossDictionaryID, Gloss.Name INTO tempglossDupes FROM Gloss INNER JOIN GlossDictionary ON Gloss.GlossID = GlossDictionary.GlossID " & _
        '    "GROUP BY GlossDictionary.GlossDictionaryID, Gloss.Name HAVING(((Count(Gloss.Name)) > 1)) ORDER BY Gloss.Name;"
        DBCommand2.CommandText = SQL
        Try
            DBCommand2.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Failed creating tempGlossDupes with error - " & ex.Message)
        End Try

        SQL = "UPDATE tempGlossDupes INNER JOIN Gloss ON tempGlossDupes.[Name] = Gloss.[Name] SET Gloss.[Name] = [Gloss]![Name] & ':-:-:' & [Gloss]![GlossID];"
        'SQL = "UPDATE (tempGlossDupes INNER JOIN Gloss ON tempGlossDupes.Name = Gloss.Name) INNER JOIN GlossDictionary ON (tempGlossDupes.GlossDictionaryID = GlossDictionary.GlossDictionaryID) AND (Gloss.GlossID = GlossDictionary.GlossID) SET Gloss.Name = [Gloss]![Name] & ':-:-:' & [Gloss]![GlossID];"
        DBCommand2.CommandText = SQL
        Try
            DBCommand2.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Failed gloss to contain dupes names with appended :-:-:GlossID with error - " & ex.Message)
        End Try

        'DICTIONARY PER SURVEY RATHER THAN FROM THE DATABASE
        'because wordlists may have words from multiple dictionaries - not compatible with WS7 1 dictionary per survey
        'so we create a new one from the glosses utilized by the survey, challenge is those unused by others
        'DROP tempALLDictionaries if it exists
        Try
            DBCommand2.CommandText = "DROP TABLE tempALLDictionaries;"
            Dim dropDictionaryResult As Integer = DBCommand2.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        'Create a selection of all surveys
        SQL = "SELECT Survey.[SurveyID], Survey.[Name], Survey.[Description], Survey.[Location], Survey.[CreationDate] FROM Survey ORDER BY Survey.[Name];"
        DBCommand.CommandText = SQL
        Dim SReader As OleDbDataReader = DBCommand.ExecuteReader()
        Dim SurveyList As New List(Of String)
        Dim SurveyCount As Integer = 0
        While SReader.Read 'FOR EACH SURVEY . . . . CREATE A DICTIONARY
            SurveyList.Add(SReader.GetValue(0).ToString.Replace(",", "-") & "," & SReader.GetValue(1).ToString.Replace(",", "-") & "," & SReader.GetValue(2).ToString.Replace(",", "-") & "," & SReader.GetValue(3).ToString.Replace(",", "-") & "," & SReader.GetValue(4).ToString.Replace(",", "-"))
            outputFile.WriteLine("Start Dictionary")
            temp = vbTab & "Name=Dictionary Imported For " & SReader.GetValue(1).ToString
            outputFile.WriteLine(temp)
            'GLOSSES
            outputFile.WriteLine(vbTab & "Start Glosses")

            'Create the temp dictionary table to use below to fill the dictionary AND for later use to create the variety word lists with all possible glosses used by any of the varieties in the survey!!!

            If Not (DoesTableExist("tempALLDictionaries", DBConnectionString)) Then ', Gloss.[Name] & ':-:-:' & Gloss.[GlossID] AS ExtendedGlossName
                SQL = "SELECT DISTINCT Survey.[SurveyID], Survey.[Name], Survey.[Description], Survey.[Location], Survey.[CreationDate], Gloss.[Name], Gloss.[PartOfSpeech], Gloss.[Definition], " & SurveyCount & " AS DictionaryID, Gloss.[GlossID] INTO tempALLDictionaries " & _
                        "FROM Gloss INNER JOIN ((Survey INNER JOIN WordListInfo ON Survey.[SurveyID] = WordListInfo.[SurveyID]) INNER JOIN WordList ON WordListInfo.[WordListID] = WordList.[WordListID]) ON Gloss.[GlossID] = WordList.[GlossID] " & _
                        "WHERE(((Survey.SurveyID) = " & SReader.GetValue(0).ToString & "))" & _
                        "ORDER BY Survey.[Name], Gloss.[Name]; "
            Else ', Gloss.[Name] & ':-:-:' & Gloss.[GlossID] AS ExtendedGlossName
                SQL = "INSERT INTO tempALLDictionaries (SurveyID, Survey_Name, Description, Location, CreationDate, Gloss_Name, PartOfSpeech, Definition, DictionaryID, GlossID) " & _
                        "SELECT DISTINCT Survey.[SurveyID], Survey.[Name], Survey.[Description], Survey.[Location], Survey.[CreationDate], Gloss.[Name], Gloss.[PartOfSpeech], Gloss.[Definition], " & SurveyCount & " AS DictionaryID, Gloss.[GlossID] " & _
                        "FROM Gloss INNER JOIN ((Survey INNER JOIN WordListInfo ON Survey.[SurveyID] = WordListInfo.[SurveyID]) INNER JOIN WordList ON WordListInfo.[WordListID] = WordList.[WordListID]) ON Gloss.[GlossID] = WordList.[GlossID] " & _
                        "WHERE(((Survey.SurveyID) = " & SReader.GetValue(0).ToString & ")) " & _
                        "ORDER BY Survey.[Name], Gloss.[Name]; "
            End If
            'Either create or append to tempAllDictionaries
            DBCommand2.CommandText = SQL
            DBCommand2.ExecuteNonQuery()

            'Now retrieve just the data for this current survey
            SQL = "SELECT * FROM tempALLDictionaries WHERE (tempALLDictionaries.[DictionaryID] = " & SurveyCount & ") ORDER BY tempALLDictionaries.[Survey_Name], tempALLDictionaries.[Gloss_Name];"
            DBCommand2.CommandText = SQL
            Dim newGlossAlphaD As New System.Collections.Generic.Dictionary(Of String, Integer)
            Try
                Dim GlossReader As OleDbDataReader = DBCommand2.ExecuteReader()
                counter = 0
                While GlossReader.Read
                    temp = vbTab & vbTab & GlossReader.GetValue(5).ToString & "||" & GlossReader.GetValue(6).ToString & "|" & GlossReader.GetValue(7).ToString & "|"
                    If Not (newGlossAlphaD.ContainsKey(GlossReader.GetValue(5).ToString)) Then
                        newGlossAlphaD.Add(GlossReader.GetValue(5).ToString, counter)
                    Else

                    End If

                    outputFile.WriteLine(temp)
                    counter = counter + 1
                End While
                outputFile.WriteLine(vbTab & "End Glosses")
                GlossReader.Close()
            Catch ex As Exception
                MsgBox("Failed creating gloss dictionaries with error - " & ex.Message)
            End Try

            'SORTS CREATION (ONLY ONE CREATED - Alphabetic ORDER) PER DICTIONARY
            outputFile.WriteLine(vbTab & "Start Sorts")
            outputFile.Write(vbTab & vbTab & "Alphabetic Order")
            temp = ""
            For i As Integer = 0 To counter - 1
                temp += "|" & i.ToString
            Next
            outputFile.WriteLine(temp)

            ''query out the creation date (really the original gloss id)
            ''lookup new id in local dictionary
            ''output as second collation
            Dim theGloss As String
            Dim foundGloss As Boolean
            Dim entryPosition As Integer
            SQL = "SELECT * FROM tempALLDictionaries WHERE (tempALLDictionaries.[DictionaryID] = " & SurveyCount & ") ORDER BY tempALLDictionaries.[GlossID];"
            DBCommand4.CommandText = SQL
            Dim GlossReader4 As OleDbDataReader = DBCommand4.ExecuteReader()
            counter = 0
            temp = vbTab & vbTab & "Entry Order"
            While GlossReader4.Read
                theGloss = GlossReader4.GetValue(5).ToString
                foundGloss = newGlossAlphaD.TryGetValue(theGloss, entryPosition)
                If foundGloss Then
                    temp += "|" & entryPosition.ToString
                End If
            End While

            outputFile.WriteLine(temp)
            GlossReader4.Close()

            outputFile.WriteLine(vbTab & "End Sorts")

            'END OF DICTIONARY INFORMATION
            outputFile.WriteLine(vbTab & "Current Sort=0")
            outputFile.WriteLine(vbTab & "Current Gloss=0")
            outputFile.WriteLine(vbTab & "Current Gloss Column Index=0")
            outputFile.WriteLine("End Dictionary" & vbCrLf)
            SurveyCount += 1
        End While 'Creation of a dictionary per each survey
        SReader.Close()
        'END OF DICTIONARIES



        'START OF SURVEYS
        counter = 0
        Dim fields() As String
        Dim surveyName As String
        If SurveyCount > 0 Then cursurv = 0
        While counter < SurveyCount
            fields = Split(SurveyList(counter), ",")
            outputFile.WriteLine("Start Survey")
            outputFile.WriteLine(vbTab & "Associated Dictionary=" & counter)
            outputFile.WriteLine(vbTab & "Name=" & fields(1))
            surveyName = fields(1)
            outputFile.WriteLine(vbTab & "Start Varieties")
            'Get all the varieties for this survey
            SQL = "SELECT WordListInfo.[SurveyID], WordListInfo.[WordListID], WordListInfo.[Name], WordListInfo.[Description], WordListInfo.[StartDate], WordListInfo.[EndDate], WordListInfo.[Surveyors], WordListInfo.[Consultants], WordListInfo.[LanguageHelper], WordListInfo.[LanguageHelperAge], WordListInfo.[LanguageHelperGender], WordListInfo.[Reliability], WordListInfo.[CreationDate], WordListInfo.[Language], WordListInfo.[Village], WordListInfo.[ProvinceState], WordListInfo.[SubDistrict], WordListInfo.[District], WordListInfo.[Country], WordListInfo.[Coordinates], WordListInfo.[PalmsurvTemplate], WordListInfo.[CurrentCollationID] FROM(WordListInfo) WHERE (((WordListInfo.[SurveyID])=" & fields(0) & ")) ORDER BY WordListInfo.[Name];"
            DBCommand.CommandText = SQL
            Try
                temp2 = ""
                VarCount = 0
                Dim VarietyReader As OleDbDataReader = DBCommand.ExecuteReader()
                'FOR EACH VARIETY IN THE SURVEY
                While VarietyReader.Read
                    outputFile.WriteLine(vbTab & vbTab & "Start Variety")
                    outputFile.WriteLine(vbTab & vbTab & vbTab & "Associated Dictionary=" & counter)
                    outputFile.WriteLine(vbTab & vbTab & vbTab & "Name=" & VarietyReader.GetValue(2).ToString)
                    outputFile.WriteLine(vbTab & vbTab & vbTab & "Start Transcriptions")

                    'DROP tempDictionary table if it exists
                    Try
                        DBCommand2.CommandText = "DROP TABLE tempDictionary;"
                        Dim dropWordlistResult As Integer = DBCommand2.ExecuteNonQuery()
                    Catch ex As Exception
                    End Try
                    'Create the tempDictionary for this survey
                    SQL = "SELECT tempALLDictionaries.* into tempDictionary FROM tempALLDictionaries  WHERE (tempALLDictionaries.Survey_Name = '" & surveyName.Replace("'", "''") & "') ORDER BY Gloss_Name;"
                    DBCommand2.CommandText = SQL
                    DBCommand2.ExecuteNonQuery()

                    'DROP tempVarietyWordlist table if it exists
                    Try
                        DBCommand2.CommandText = "DROP TABLE tempVarietyWordlist;"
                        Dim dropWordlistResult As Integer = DBCommand2.ExecuteNonQuery()
                    Catch ex As Exception
                    End Try
                    'Create the tempVarietyWordlist table so a later query involving the dictionary will include null glosses (in the dictionary but not in the variety wordlist)
                    SQL = "SELECT DISTINCT WordListInfo.SurveyID, WordList.WordListID, WordList.Transcription, WordList.Notes, Gloss.Name INTO tempVarietyWordlist " & _
                            "FROM WordListInfo INNER JOIN (Gloss INNER JOIN WordList ON Gloss.GlossID = WordList.GlossID) ON WordListInfo.WordListID = WordList.WordListID " & _
                            "WHERE (((WordListInfo.SurveyID)=" & fields(0) & ") AND ((WordList.WordListID)=" & VarietyReader.GetValue(1).ToString & "));"
                    DBCommand2.CommandText = SQL
                    Dim createWordListTableResult As Integer = DBCommand2.ExecuteNonQuery()
                    'Now create the recordset for the variety containing ALL of the glosses (even if not in the original variety)
                    SQL = "SELECT tempVarietyWordlist.SurveyID, tempVarietyWordlist.WordListID, tempVarietyWordlist.Transcription, tempVarietyWordlist.Notes, tempDictionary.Gloss_Name " & _
                            "FROM tempDictionary LEFT JOIN tempVarietyWordlist ON tempDictionary.Gloss_Name = tempVarietyWordlist.Name;"
                    DBCommand2.CommandText = SQL
                    Dim TransReader As OleDbDataReader = DBCommand2.ExecuteReader()
                    Dim TransList As New List(Of TransEntry)
                    Dim TransCount As Integer = 0
                    While TransReader.Read
                        'transcription, plural/frame, notes
                        Dim trans As New TransEntry
                        trans.Transcription = IfNullStr(TransReader.GetValue(2))
                        trans.Notes = IfNullStr(TransReader.GetValue(3))
                        'trans.GlossID = TransReader.GetValue(4)
                        trans.GlossName = TransReader.GetValue(4).ToString 'AJW***
                        TransList.Add(trans)
                        TransCount += 1
                    End While
                    TransReader.Close()


                    'Combine synonyms
                    Dim i As Integer = 0
                    Dim nextGlossID As Integer = -1
                    Dim synonyms As String = ""
                    Dim glossCnt As Integer = 0
                    While i < TransCount
                        'If i < TransCount - 1 AndAlso TransList(i).GlossID = TransList(i + 1).GlossID Then
                        If i < TransCount - 1 AndAlso TransList(i).GlossName = TransList(i + 1).GlossName Then 'AJW***
                            synonyms &= TransList(i).Transcription & ","
                        Else
                            synonyms &= TransList(i).Transcription
                            temp = glossCnt & "|" & glossCnt & "|" & synonyms & "||" & TransList(i).Notes
                            outputFile.WriteLine(vbTab & vbTab & vbTab & vbTab & temp)
                            glossCnt += 1
                            synonyms = ""
                        End If

                        i += 1
                    End While

                    outputFile.WriteLine(vbTab & vbTab & vbTab & "End Transcriptions")
                    outputFile.WriteLine(vbTab & vbTab & vbTab & "Current VarietyEntry=0")
                    outputFile.WriteLine(vbTab & vbTab & vbTab & "Description=" & "Notes:" & IfNullStr(VarietyReader.GetValue(3)) & "\StartDate:" & (VarietyReader.GetValue(4).ToString & "\EndDate:" & VarietyReader.GetValue(5).ToString & "\Surveyors:" & IfNullStr(VarietyReader.GetValue(6)) & "\Consultants:" & IfNullStr(VarietyReader.GetValue(7)) & "\LanguageHelper:" & IfNullStr(VarietyReader.GetValue(8)) & "\Reliability:" & IfNullStr(VarietyReader.GetValue(11)) & "\Language:" & IfNullStr(VarietyReader.GetValue(13)) & "\Country:" & IfNullStr(VarietyReader.GetValue(18)) & "\ProvinceState:" & IfNullStr(VarietyReader.GetValue(15)) & "\District:" & IfNullStr(VarietyReader.GetValue(17)) & "\Subdistrict:" & IfNullStr(VarietyReader.GetValue(16)) & "\Village:" & IfNullStr(VarietyReader.GetValue(14)) & "\Coordinates:" & IfNullStr(VarietyReader.GetValue(19)) & "\CreationDate:" & VarietyReader.GetValue(12).ToString))
                    outputFile.WriteLine(vbTab & vbTab & "End Variety")
                    VarCount += 1
                End While
                VarietyReader.Close()
            Catch varEx As Exception
                MsgBox(varEx.Message)
            End Try
            outputFile.WriteLine(vbTab & "End Varieties")
            outputFile.WriteLine(vbTab & "Current Variety=0")
            outputFile.WriteLine(vbTab & "Description=" & fields(2) & "   -   " & fields(3) & "   -   " & fields(4))
            outputFile.WriteLine(vbTab & "Current VarietyEntry Column Index=0")
            outputFile.WriteLine("End Survey" & vbCrLf)
            SurveyList(counter) = SurveyList(counter) & "," & VarCount.ToString
            counter += 1
        End While
        'END OF SURVEYS

        'START OF COMPARISONS
        'Query out each unique CompID, Survey name combo that exists in WS6
        SQL = "SELECT DISTINCT [ComparisonInfo]![Name] & '-' & [Survey]![Name] AS CompSurveyCombined, ComparisonInfo.[ComparisonID], Survey.[Name], ComparisonInfo.[Description], ComparisonInfo.[CreationDate] FROM (Survey INNER JOIN WordListInfo ON Survey.[SurveyID] = WordListInfo.[SurveyID]) INNER JOIN (WordList INNER JOIN (ComparisonInfo INNER JOIN Comparison ON ComparisonInfo.[ComparisonID] = Comparison.[ComparisonID]) ON WordList.[WordListEntryID] = Comparison.[WordListEntryID]) ON WordListInfo.[WordListID] = WordList.[WordListID] GROUP BY [ComparisonInfo]![Name] & '-' & [Survey]![Name], ComparisonInfo.[ComparisonID], Survey.[Name], ComparisonInfo.[Description], ComparisonInfo.[CreationDate] ORDER BY ComparisonInfo.ComparisonID;"
        DBCommand.CommandText = SQL
        Dim ComparisonID As Integer = -1
        Dim VarietyName As String = ""
        Dim comparisonEntryCount As Integer = 0
        Dim numVarieties As Integer = 0
        Dim VarList As New List(Of String)
        Dim CompReader As OleDbDataReader = DBCommand.ExecuteReader()
        While CompReader.Read 'FOR EACH COMPARISON/SURVEY PAIR
            outputFile.WriteLine("Start Comparison")
            outputFile.WriteLine(vbTab & "Associated Survey=" & GetSurveyIndex(CompReader.GetValue(2).ToString, SurveyList))
            outputFile.WriteLine(vbTab & "Name=" & CompReader.GetValue(0).ToString)
            'Gets variety names for a particular comparison/survey
            'DROP tempVarieties table if it exists
            Try
                DBCommand2.CommandText = "DROP TABLE tempVarieties;"
                Dim dropWordlistResult As Integer = DBCommand2.ExecuteNonQuery()
            Catch ex As Exception
            End Try
            'NEW SQL HERE FOR ALL VARIETIeS IN A SURVEY, NOT THE COMPARISON (BECUASE OF WS7 ASSUMPIONS THAT ALL VARIETIES FROM A SURVEY ARE IN A COMPARISON)
            SQL = "SELECT DISTINCT Survey.Name AS [DUMMY], Survey.Name, WordListInfo.Name into tempVarieties " & _
                    "FROM (Survey INNER JOIN WordListInfo ON Survey.[SurveyID] = WordListInfo.[SurveyID]) INNER JOIN WordList ON WordListInfo.[WordListID] = WordList.[WordListID] " & _
                    "GROUP BY Survey.Name, Survey.Name, WordListInfo.Name " & _
                    "HAVING (((Survey.Name)='" & CompReader.GetValue(2).ToString.Replace("'", "''") & "')) " & _
                    "ORDER BY WordListInfo.Name; "
            DBCommand2.CommandText = SQL
            DBCommand2.ExecuteNonQuery()

            SQL = "SELECT tempVarieties.* FROM tempVarieties ORDER BY WordListInfo_Name;"
            DBCommand2.CommandText = SQL
            Dim VarietiesReader As OleDbDataReader = DBCommand2.ExecuteReader()
            numVarieties = 0
            Dim varSort As String = ""
            VarList.Clear()
            While VarietiesReader.Read
                varSort = varSort & numVarieties.ToString & "|"
                VarList.Add(VarietiesReader.GetValue(2).ToString)
                numVarieties += 1
            End While
            VarietiesReader.Close()
            varSort = varSort.TrimEnd("|"c)
            outputFile.WriteLine(vbTab & "Variety Sort=" & varSort)
            outputFile.WriteLine(vbTab & "Start Comparison Entries")
            ComparisonID = Integer.Parse(CompReader.GetValue(1).ToString)
            surveyName = CompReader.GetValue(2).ToString
            'DROP tempDictionary table if it exists
            Try
                DBCommand2.CommandText = "DROP TABLE tempDictionary;"
                Dim dropWordlistResult As Integer = DBCommand2.ExecuteNonQuery()
            Catch ex As Exception
            End Try
            'Create the tempDictionary for this survey and counts the number of gloss entries
            SQL = "SELECT tempALLDictionaries.* into tempDictionary FROM tempALLDictionaries  WHERE (tempALLDictionaries.Survey_Name = '" & surveyName.Replace("'", "''") & "') ORDER BY Gloss_Name;"
            DBCommand2.CommandText = SQL
            Dim lastIndex As Integer = DBCommand2.ExecuteNonQuery() 'returns the number of records
            lastIndex -= 1
            'DROP tempVarietiesAllGlosses table if it exists
            Try
                DBCommand2.CommandText = "DROP TABLE tempVarietiesAllGlosses;"
                Dim asddd As Integer = DBCommand2.ExecuteNonQuery()
            Catch ex As Exception
            End Try
            'Now create the tempVarietiesAllGlosses table (glosses for all varieties in this Comparison)
            SQL = "SELECT DISTINCT WordListInfo.Name, tempDictionary.Gloss_Name INTO tempVarietiesAllGlosses " & _
                    "FROM tempVarieties INNER JOIN (tempDictionary INNER JOIN WordListInfo ON tempDictionary.SurveyID = WordListInfo.SurveyID) ON tempVarieties.WordListInfo_Name = WordListInfo.Name " & _
                    "ORDER BY WordListInfo.Name, tempDictionary.Gloss_Name; "
            DBCommand2.CommandText = SQL
            DBCommand2.ExecuteNonQuery()
            SQL = "SELECT DISTINCT tempVarietiesAllGlosses.[Name] FROM tempVarietiesAllGlosses " & _
                    "ORDER BY tempVarietiesAllGlosses.[Name];"
            DBCommand3.CommandText = SQL
            'for each variety in tempAllVarietiesInSurvey
            Dim entryCount As Integer = 0
            Dim CompEntryList As New List(Of CompEntry)
            Sleep(5000)
            Dim VarietiesPerCompSurvReader As OleDbDataReader = DBCommand3.ExecuteReader()
            While VarietiesPerCompSurvReader.Read
                VarietyName = VarietiesPerCompSurvReader.GetValue(0).ToString
                'DROP tempXComparisonData table if it exists
                Try
                    DBCommand2.CommandText = "DROP TABLE tempXComparisonData;"
                    Dim asddd As Integer = DBCommand2.ExecuteNonQuery()
                Catch ex As Exception
                End Try
                SQL = "SELECT Comparison.ComparisonID, WordListInfo.Name, Gloss.Name, Comparison.AlignedRendering, Comparison.GroupAssigned, Comparison.Notes, Comparison.Exclude, Gloss.GlossID, WordList.Transcription INTO tempXComparisonData " & _
                        "FROM WordListInfo INNER JOIN ((Gloss INNER JOIN WordList ON Gloss.GlossID = WordList.GlossID) INNER JOIN Comparison ON WordList.WordListEntryID = Comparison.WordListEntryID) ON WordListInfo.WordListID = WordList.WordListID " & _
                        "WHERE(((Comparison.ComparisonID) = " & ComparisonID & ") And ((WordListInfo.Name) = '" & VarietyName.Replace("'", "''") & "')) " & _
                        "ORDER BY WordListInfo.Name, Gloss.Name; "
                DBCommand2.CommandText = SQL
                DBCommand2.ExecuteNonQuery()
                SQL = "SELECT tempVarietiesAllGlosses.Name, tempXComparisonData.WordListInfo_Name, tempVarietiesAllGlosses.Gloss_Name, tempXComparisonData.AlignedRendering, tempXComparisonData.GroupAssigned, tempXComparisonData.Notes, tempXComparisonData.Exclude, tempXComparisonData.GlossID, tempXComparisonData.Transcription " & _
                        "FROM tempVarietiesAllGlosses LEFT JOIN tempXComparisonData ON tempVarietiesAllGlosses.Gloss_Name = tempXComparisonData.Gloss_Name " & _
                        "WHERE (((tempVarietiesAllGlosses.Name)='" & VarietyName.Replace("'", "''") & "'));"
                DBCommand2.CommandText = SQL
                Dim CompEntriesReader As OleDbDataReader = DBCommand2.ExecuteReader()
                While CompEntriesReader.Read
                    Dim compEnt As New CompEntry
                    compEnt.Gloss = CompEntriesReader.GetValue(2).ToString
                    compEnt.AlignedRendering = IfNullStr(CompEntriesReader.GetValue(3))
                    compEnt.Grouping = IfNullStr(CompEntriesReader.GetValue(4))
                    compEnt.Notes = IfNullStr(CompEntriesReader.GetValue(5))
                    compEnt.Exclude = IfNullBool(CompEntriesReader.GetValue(6))
                    compEnt.GlossID = Integer.Parse(IfNullInt(CompEntriesReader.GetValue(7)))
                    compEnt.Transcription = IfNullStr(CompEntriesReader.GetValue(8))
                    CompEntryList.Add(compEnt)
                    entryCount += 1
                End While
                CompEntriesReader.Close()
                CompEntriesReader = Nothing
            End While
            VarietiesPerCompSurvReader.Close()
            VarietiesPerCompSurvReader = Nothing


            Dim i As Integer = 0
            Dim nextGlossID As Integer = -1
            Dim synonyms As String = ""
            Dim groupings As String = ""
            Dim entryIndex As Integer = 0

            While i < entryCount
                If i < entryCount - 1 AndAlso CompEntryList(i).Gloss = CompEntryList(i + 1).Gloss Then 'If there is a synonym or additional grouping because the gloss IDs are the same 'AJW changed from checking ID to checking Gloss Name
                    'If i < entryCount - 1 AndAlso CompEntryList(i).GlossID = CompEntryList(i + 1).GlossID Then 'If there is a synonym or additional grouping because the gloss IDs are the same

                    If CompEntryList(i).Transcription <> CompEntryList(i + 1).Transcription Then 'If it is not just an additional grouping but a full synonym,
                        groupings &= CompEntryList(i).Grouping & ","
                        Dim parts As String() = Split(synonyms, ",")
                        If synonyms = "" OrElse (parts.Length > 0 AndAlso (Not parts(parts.Length - 1) = CompEntryList(i).Transcription)) Then
                            synonyms &= CompEntryList(i).Transcription
                        End If
                        synonyms &= ","
                    Else
                        groupings &= CompEntryList(i).Grouping
                        Dim parts As String() = Split(synonyms, ",")
                        If synonyms = "" OrElse (parts.Length > 0 AndAlso (Not parts(parts.Length - 1) = CompEntryList(i).Transcription)) Then
                            synonyms &= CompEntryList(i).Transcription
                        End If
                    End If
                Else
                    groupings &= CompEntryList(i).Grouping
                    Dim parts As String() = Split(synonyms, ",")
                    If synonyms = "" OrElse (parts.Length > 0 AndAlso (Not parts(parts.Length - 1) = CompEntryList(i).Transcription)) Then
                        synonyms &= CompEntryList(i).Transcription
                    End If


                    temp = entryIndex.ToString & "|" & synonyms & "|" & groupings & "|" & CompEntryList(i).Notes & "|" & convertBoolToString(CompEntryList(i).Exclude)
                    outputFile.WriteLine(vbTab & vbTab & temp)
                    If entryIndex = lastIndex Then
                        entryIndex = 0
                    Else
                        entryIndex += 1
                    End If
                    synonyms = ""
                    groupings = ""
                End If

                i += 1
            End While


            outputFile.WriteLine(vbTab & "End Comparison Entries")
            outputFile.WriteLine(vbTab & "Current Variety=0")
            outputFile.WriteLine(vbTab & "Description=" & IfNullStr(CompReader.GetValue(3)))
            outputFile.WriteLine(vbTab & "Start Date=")
            outputFile.WriteLine(vbTab & "End Date=")
            outputFile.WriteLine(vbTab & "Current Variety Column Index=0")
            outputFile.Write(vbTab & "DD Used Chars=")

            outputFile.WriteLine()
            'Dim usedChars As New List(Of String)
            'Dim DDs As New Dictionary(Of String, Dictionary(Of String, Integer))
            'SQL = "SELECT DISTINCT DegreesOfDifference.ComparisonID, DegreesOfDifference.Character1 FROM(DegreesOfDifference) WHERE(((DegreesOfDifference.ComparisonID) = " & CompReader.GetValue(1) & ")) ORDER BY DegreesOfDifference.Character1;"
            'temp = ""
            'DBCommand2.CommandText = SQL
            'Dim DDCharsReader As OleDbDataReader = DBCommand2.ExecuteReader()
            'Dim DDcharCount As Integer = 0
            'While DDCharsReader.Read
            '    temp = temp & DDCharsReader.GetValue(1) & "|"
            '    usedChars.Add(DDCharsReader.GetValue(1))
            '    DDcharCount += 1
            'End While
            'DDCharsReader.Close()
            ''DBCommand2.Dispose()
            'outputFile.WriteLine(temp.TrimEnd("|"))
            outputFile.WriteLine(vbTab & "Start DD Values")

            ''fill the matrix with -1's
            'For Each usedChar1 As String In usedChars
            '    DDs.Add(usedChar1, New Dictionary(Of String, Integer))
            '    For Each usedChar2 As String In usedChars
            '        DDs(usedChar1).Add(usedChar2, -1)
            '    Next
            'Next

            ''replace with matching pair values
            'SQL = "SELECT DegreesOfDifference.ComparisonID, DegreesOfDifference.Character1, DegreesOfDifference.Character2, DegreesOfDifference.DD FROM(DegreesOfDifference) WHERE(((DegreesOfDifference.ComparisonID) =  " & CompReader.GetValue(1) & ")) ORDER BY DegreesOfDifference.Character1, DegreesOfDifference.Character2;"
            'DBCommand2.CommandText = SQL
            'Dim DDValuesReader As OleDbDataReader = DBCommand2.ExecuteReader()
            'While DDValuesReader.Read
            '    DDs(DDValuesReader.GetValue(1).ToString)(DDValuesReader.GetValue(2)) = DDValuesReader.GetValue(3)
            'End While
            'DDValuesReader.Close()

            'For Each usedChar1 As String In usedChars
            '    Dim ddRowStr As String = ""
            '    For Each usedChar2 As String In usedChars
            '        ddRowStr &= DDs(usedChar1)(usedChar2).ToString & "|"
            '    Next
            '    ddRowStr = ddRowStr.TrimEnd("|"c)
            '    outputFile.WriteLine(vbTab & vbTab & ddRowStr)
            'Next

            outputFile.WriteLine(vbTab & "End DD Values")
            outputFile.WriteLine(vbTab & "Excluded DD Chars=")
            outputFile.WriteLine("End Comparison" & vbCrLf)
            comparisonEntryCount += 1
            curcomp = 0
        End While
        'END OF COMPARISONS

        'START OF WORDSURV DATA
        outputFile.WriteLine("Start WordSurv Data")
        outputFile.WriteLine(vbTab & "Current Dictionary=0")
        outputFile.WriteLine(vbTab & "Current Survey=" & cursurv.ToString)
        outputFile.WriteLine(vbTab & "Current Comparison=" & curcomp.ToString)
        outputFile.WriteLine(vbTab & "Primary Language=English")
        outputFile.WriteLine(vbTab & "Secondary Language=English")
        outputFile.WriteLine(vbTab & "Primary Font=Microsoft Sans Serif,8")
        outputFile.WriteLine(vbTab & "Secondary Font=Microsoft Sans Serif,8")
        outputFile.WriteLine(vbTab & "Transcription Font=Microsoft Sans Serif,8")
        outputFile.WriteLine("End WordSurv Data")
        outputFile.WriteLine("WordSurv 7.0 Beta release Kemuel")
        'END OF WORDSURV DATA

        outputFile.Close()

        closeImportWS6()
        MsgBox("Completed conversion.  " & toWSV & " is ready for you to open!  You will be prompted to save any changes to the current .wsv file if you wish before the new imported database is opened.")
        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

        'WordSurvForm.openDatabaseFile(data, prefs, filedialog.FileName.Substring(0, filedialog.FileName.LastIndexOf(".")) & ".wsv")
    End Sub

    Public Sub ImportWordSurv6(ByVal data As WordSurvData, ByVal prefs As Preferences, wsform As WordSurvForm)
        Dim mdbdialog As New OpenFileDialog
        'mdbdialog.InitialDirectory = Application.StartupPath
        mdbdialog.Title = "Open WordSurv 6 File for Conversion to WordSurv 7 Format (.mdb to .wsv)"
        mdbdialog.Filter = "mdbfiles|*.mdb"
        mdbdialog.FilterIndex = 0
        'mdbdialog.ShowDialog() = System.Windows.Forms.DialogResult.OK

        'if file exists, pop up a message to ask if they want to overwrite it
        Dim wsvdialog As New SaveFileDialog
        wsvdialog.Title = "Save .wsv File as..."
        wsvdialog.Filter = "wsvfiles|*.wsv"
        wsvdialog.FilterIndex = 0

        If mdbdialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim basename As String = mdbdialog.FileName
            basename = basename.Substring(0, basename.LastIndexOf("."))
            wsvdialog.FileName = basename.Substring(basename.LastIndexOf("\") + 1) & ".wsv"
            If wsvdialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                Try
                    ConvertWS6ToWS7(mdbdialog.FileName, wsvdialog.FileName)
                Catch ex As Exception
                    MsgBox("Cannot import " & wsvdialog.FileName & ".  To help improve future versions of WordSurv, please send the file you wish to import to the developers (See contact information under ""Technical Support"" in the Help).  When sending files, please also include 1) the following technical details 2) how many other files you have in a similar format that need to be imported, 3) how important importing this data is.  Technical details (press Control+c to copy this message): " & ex.Message)
                End Try
                ' open the converted file
                wsform.Open(wsvdialog.FileName)
            Else
                MsgBox("WordSurv 6 to WordSurv 7 Database Conversion Cancelled!")
            End If
        Else
            MsgBox("WordSurv 6 to WordSurv 7 Database Conversion Cancelled!")
        End If
    End Sub

    Public Function GetSurveyIndex(ByVal surveyName As String, ByVal SurveyList As List(Of String)) As Integer
        'given a survey name, returns index that was included in the survey string in the storage array
        Dim fields() As String
        For i As Integer = 0 To SurveyList.Count - 1
            fields = Split(SurveyList(i), ",")
            If fields(1) = surveyName Then Return i
        Next
        Return -1
    End Function
    Public Function convertBoolToString(ByVal boolVal As Boolean) As String
        If boolVal = False Then Return ""
        If boolVal = True Then Return "x"
        Return "Invalid Boolean value in converter"
    End Function
    Public Function IfNullStr(ByVal val As Object) As String
        If IsDBNull(val) Or val Is Nothing Then
            Return ""
        Else
            Dim temp As String = val.ToString
            If temp.Contains(vbCrLf) Then
                temp = temp.Replace(vbCrLf, " ")
            End If
            If temp.Contains(vbCr) Then
                temp = temp.Replace(vbCr, " ")
            End If
            If temp.Contains(vbLf) Then
                temp = temp.Replace(vbLf, " ")
            End If
            If temp.Contains(vbTab) Then
                val = temp.Replace(vbTab, " ")
            End If
            Return temp
        End If
    End Function
    Public Function IfNullBool(ByVal val As Object) As Boolean
        If IsDBNull(val) OrElse val Is Nothing Or val.ToString = "False" Then
            Return False
        Else
            Return True
        End If
    End Function
    Public Function IfNullInt(ByVal val As Object) As String
        If IsDBNull(val) Or val Is Nothing Then
            Return "-1"
        Else
            Return val.ToString
        End If
    End Function
    Public Function DoesTableExist(ByVal tblName As String, ByVal cnnStr As String) As Boolean
        ' Open connection to the database
        Dim dbConn As New OleDbConnection(cnnStr)
        dbConn.Open()

        Dim restrictions(3) As String
        restrictions(2) = tblName
        Dim dbTbl As DataTable = dbConn.GetSchema("Tables", restrictions)

        If dbTbl.Rows.Count = 0 Then
            'Table does not exist
            DoesTableExist = False
        Else
            'Table exists
            DoesTableExist = True
        End If

        dbTbl.Dispose()
        dbConn.Close()
        dbConn.Dispose()
    End Function
    Public Sub closeImportWS6()
        DBCommand.Dispose()
        DBCommand2.Dispose()
        DBCommand3.Dispose()
        DBConnection.Close()
        DBConnection.Dispose()
        DBConnection2.Close()
        DBConnection2.Dispose()
        DBConnection3.Close()
        DBConnection3.Dispose()
    End Sub
End Module
