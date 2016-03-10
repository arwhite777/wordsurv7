Public Class EncodingChoice
    Private dbFileNameBase As String
    Public Sub New(ByVal dbFileNameBase As String)
        Me.InitializeComponent()
        Me.dbFileNameBase = dbFileNameBase
    End Sub

    Private Sub EncodingChoice_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.cboEncodingSelector.SelectedIndex = 0
        Me.FillTextbox()
    End Sub

    Private Sub cboEncodingSelector_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboEncodingSelector.SelectedIndexChanged
        Me.FillTextbox()
    End Sub
    Private Sub FillTextbox()
        Try
            Dim inputDBFile As New IO.StreamReader(dbFileNameBase & ".db", System.Text.Encoding.GetEncoding(Me.cboEncodingSelector.SelectedItem.ToString))

            Me.txtFilePreview.Text = inputDBFile.ReadToEnd
            inputDBFile.Close()

            Dim inputCATFile As New IO.StreamReader(dbFileNameBase & ".cat", System.Text.Encoding.GetEncoding(Me.cboEncodingSelector.SelectedItem.ToString))
            Me.txtFilePreview.Text = inputCATFile.ReadToEnd
            inputCATFile.Close()
        Catch ex As Exception
            'MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
    End Sub
    Private Sub btnHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHelp.Click
        MsgBox("WordSurv 2.5 uses an older way of storing text.  Select the encoding standard from the combo box that makes the text readable.  The default is the most common.")
    End Sub

End Class