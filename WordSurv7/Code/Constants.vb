Public Module Constants
    Public INVALID_COLOR As Color = Color.FromArgb(255, 255, 153)
    Public ERROR_COLOR As Color = Color.Red
    Public NUMBER_OF_RECENT_DATABASES As Integer = 10
    Public NON_EDITABLE_COLOR As Color = Color.LightGray
    Public EMPTY_GRID_COLOR As Color = Color.FromArgb(175, 175, 175)
    Public SEARCH_COLOR As Color = Color.CornflowerBlue
    Public CUT_SELECTION As Color = Color.CornflowerBlue
    Public CROSSHAIRS_COLOR As Color = Color.LightBlue
    Public INACTIVE_COLOR As Color = Color.FromArgb(170, 170, 170)
    Public HasNotSaved As Boolean = False
    Public BackupTimeStamp As DateTime = Now
    Public DoLog As Boolean = True
    Public Log As New List(Of String)
    Public Enum SearchType
        DICTIONARY
        SURVEY
        COMPARISON_GLOSS
        COMPARISON
        COGNATE_STRENGTHS
        NONE
    End Enum
    Public Enum ValidationType
        DICTIONARY_NAME
        DICTIONARY_SORT_NAME
        SURVEY_NAME
        VARIETY_NAME
        COMPARISON_NAME
        POSITIVE_INTEGER
        ZERO_OR_POSITIVE_INTEGER
        NOT_EMPTY
        NONE
    End Enum
    'These all get set at the beginning of the program's execution and are used by other parts of the program like searching
    'so that we don't have to hardcode how many columns each grid has
    Public GlossDictionaryGridColCount As Integer
    Public VarietyGridColCount As Integer
    Public ComparisonGlossGridColCount As Integer
    Public ComparisonGridColCount As Integer
    Public CognateStrengthsGridColCount As Integer
    Public LoadInterrupted As Boolean = False
    Public Class CellAddress
        Public ObjIndex As Integer
        Public RowIndex As Integer
        Public ColIndex As Integer
        Public Sub New(ByVal objIndex As Integer, ByVal rowIndex As Integer, ByVal colIndex As Integer)
            Me.ObjIndex = objIndex
            Me.RowIndex = rowIndex
            Me.ColIndex = colIndex
        End Sub
    End Class
    Public Class IntIntComboMenu
        Public Int1 As Integer
        Public Int2 As Integer
        Public Sub New(ByVal int1 As Integer, ByVal int2 As Integer)
            Me.Int1 = int1
            Me.Int2 = int2
        End Sub
    End Class
    Public Sub Main()
        Dim frmWS As New WordSurvForm
        Application.Run(frmWS)
        'SwapFile.Close()
        'System.IO.File.Delete(SwapFileName)
    End Sub
End Module
