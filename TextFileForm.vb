Public Class TextFileForm
    Private Sub TextFileForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Dim RichTextBox1ReadOnly As Boolean
        Me.Label1.Text = My.Resources.FileMsg & " " & strFileName
        Me.RichTextBox1.Visible = True
        Me.RichTextBox1.Text = strTextFile
    End Sub
End Class