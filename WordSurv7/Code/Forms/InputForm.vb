Option Compare Text

Imports System.Windows.Forms

'This glorious form can be used any time the user desires to update the name of something or make a new thing that requires a name (dictionary, survey, variety, etc).
Public Class InputForm
    Public Result As String = ""
    Private data As WordSurvData
    Private InitialValue As String = ""
    Public Sub New(ByVal titleText As String, ByVal promptText As String, ByVal kind As ValidationType, ByRef data As WordSurvData, ByVal initialValue As String, Optional ByVal fnt As Font = Nothing)
        Me.InitializeComponent()
        Me.Text = titleText
        Me.lblPrompt.Text = promptText
        Me.data = data
        Me.InitialValue = initialValue

        If fnt IsNot Nothing Then Me.txtInput.Font = fnt

        Select Case kind
            Case ValidationType.DICTIONARY_NAME
                AddHandler txtInput.TextChanged, AddressOf inputTextChangedValidateDictionaryName
            Case ValidationType.DICTIONARY_SORT_NAME
                AddHandler txtInput.TextChanged, AddressOf inputTextChangedValidateDictionarySortName
            Case ValidationType.SURVEY_NAME
                AddHandler txtInput.TextChanged, AddressOf inputTextChangedValidateSurveyName
            Case ValidationType.VARIETY_NAME
                AddHandler txtInput.TextChanged, AddressOf inputTextChangedValidateVarietyName
            Case ValidationType.COMPARISON_NAME
                AddHandler txtInput.TextChanged, AddressOf inputTextChangedValidateComparisonName
            Case ValidationType.POSITIVE_INTEGER
                AddHandler txtInput.TextChanged, AddressOf inputTextChangedValidatePositiveInteger
            Case ValidationType.ZERO_OR_POSITIVE_INTEGER
                AddHandler txtInput.TextChanged, AddressOf inputTextChangedValidateZeroOrPositiveInteger
            Case ValidationType.NOT_EMPTY
                AddHandler txtInput.TextChanged, AddressOf inputTextChangedValidateNotEmpty
            Case ValidationType.NONE
                'do nothing
        End Select

        Me.txtInput.Text = initialValue

        If Me.txtInput.Text = "" And kind <> ValidationType.NONE Then
            Me.btnOK.Enabled = False
        End If
    End Sub


    Public Sub inputTextChangedValidatePositiveInteger(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim val As Integer
        Dim parsed As Boolean = Integer.TryParse(Me.txtInput.Text, val)

        If Not parsed OrElse val <= 0 OrElse Me.txtInput.Text = "" Then
            Me.setStatusWarning("Please enter an integer value greater than 0.")
            Me.btnOK.Enabled = False
        Else
            Me.clearStatusBar()
            Me.btnOK.Enabled = True
        End If
    End Sub


    Public Sub inputTextChangedValidateZeroOrPositiveInteger(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim val As Integer
        Dim parsed As Boolean = Integer.TryParse(Me.txtInput.Text, val)

        If Not parsed OrElse val < 0 OrElse Me.txtInput.Text = "" Then
            Me.setStatusWarning("Please enter an integer value of 0 or greater.")
            Me.btnOK.Enabled = False
        Else
            Me.clearStatusBar()
            Me.btnOK.Enabled = True
        End If
    End Sub

    Public Sub inputTextChangedValidateDictionaryName(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If (Not Me.txtInput.Text = Me.InitialValue) AndAlso (Not data.IsUniqueDictionaryName(Me.txtInput.Text)) Then
            Me.setStatusWarning("A Dictionary with that name already exists.")
            Me.btnOK.Enabled = False
        ElseIf Me.txtInput.Text = "" Then
            Me.setStatusWarning("A Dictionary's name cannot be blank.")
            Me.btnOK.Enabled = False
        Else
            Me.clearStatusBar()
            Me.btnOK.Enabled = True
        End If
    End Sub

    Public Sub inputTextChangedValidateDictionarySortName(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If (Not Me.txtInput.Text = Me.InitialValue) AndAlso (Not data.IsUniqueDictionarySortName(Me.txtInput.Text)) Then
            Me.setStatusWarning("A Sort with that name already exists.")
            Me.btnOK.Enabled = False
        ElseIf Me.txtInput.Text = "" Then
            Me.setStatusWarning("A Sort's name cannot be blank.")
            Me.btnOK.Enabled = False
        Else
            Me.clearStatusBar()
            Me.btnOK.Enabled = True
        End If
    End Sub

    Public Sub inputTextChangedValidateVarietyName(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If (Not Me.txtInput.Text = Me.InitialValue) AndAlso (Not data.IsUniqueVarietyName(Me.txtInput.Text)) Then
            Me.setStatusWarning("A Variety with that name already exists.")
            Me.btnOK.Enabled = False
        ElseIf Me.txtInput.Text = "" Then
            Me.setStatusWarning("A Variety's name cannot be blank.")
            Me.btnOK.Enabled = False
        Else
            Me.clearStatusBar()
            Me.btnOK.Enabled = True
        End If
    End Sub

    Public Sub inputTextChangedValidateSurveyName(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If (Not Me.txtInput.Text = Me.InitialValue) AndAlso (Not data.IsUniqueSurveyName(Me.txtInput.Text)) Then
            Me.setStatusWarning("A Survey with that name already exists.")
            Me.btnOK.Enabled = False
        ElseIf Me.txtInput.Text = "" Then
            Me.setStatusWarning("A Survey's name cannot be blank.")
            Me.btnOK.Enabled = False
        Else
            Me.clearStatusBar()
            Me.btnOK.Enabled = True
        End If
    End Sub


    Public Sub inputTextChangedValidateComparisonName(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If (Not Me.txtInput.Text = Me.InitialValue) AndAlso (Not data.IsUniqueComparisonName(Me.txtInput.Text)) Then
            Me.setStatusWarning("A Comparison with that name already exists.")
            Me.btnOK.Enabled = False
        ElseIf Me.txtInput.Text = "" Then
            Me.setStatusWarning("A Comparison's name cannot be blank.")
            Me.btnOK.Enabled = False
        Else
            Me.clearStatusBar()
            Me.btnOK.Enabled = True
        End If
    End Sub

    Public Sub inputTextChangedValidateNotEmpty(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Me.txtInput.Text = "" Then
            Me.setStatusWarning("This value cannot be blank.")
            Me.btnOK.Enabled = False
        Else
            Me.clearStatusBar()
            Me.btnOK.Enabled = True
        End If
    End Sub


    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.Result = Me.txtInput.Text
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
    End Sub

    Public Sub setStatusWarning(ByVal msg As String)
        Me.stsLabel1.Text = msg
        Me.stsStatusBar.BackColor = INVALID_COLOR
        Beep()
    End Sub
    Public Sub clearStatusBar()
        Me.stsLabel1.Text = ""
        Me.stsStatusBar.BackColor = Color.Empty
    End Sub

End Class
