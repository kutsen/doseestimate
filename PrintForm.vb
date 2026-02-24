

Imports System.IO
Public Class PrintForm
    Private FileToPrint As StreamReader
    Private printfont As New Font("System", 10)
    '    Private memoryImage As Bitmap


    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.OpenFileDialog1.ShowDialog()
        Me.TextBox1.Text = Me.OpenFileDialog1.FileName
        Me.PrintDocument1.DocumentName = Me.TextBox1.Text
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If Me.TextBox1.Text.Length <= 1 Then
            MsgBox("Не указано имя файла", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
            Exit Sub
        Else

            PrintingOfTextFile()

        End If
    End Sub
    Private Sub PrintingOfTextFile()
        FileToPrint = New StreamReader(Me.TextBox1.Text, System.Text.Encoding.GetEncoding(1251))
        Try
            Me.PrintDocument1.Print()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, My.Resources.MainTitle)
        End Try
        FileToPrint.Close()
    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim y As Single = e.MarginBounds.Top
        Dim line As String = Nothing
        While y < e.MarginBounds.Bottom
            line = FileToPrint.ReadLine()
            If line Is Nothing Then
                Exit While
            End If
            y += printfont.Height
            e.Graphics.DrawString(line, printfont, Brushes.Black, e.MarginBounds.Left, y)
        End While
        If Not (line Is Nothing) Then
            e.HasMorePages = True
        Else
            e.HasMorePages = False
        End If
    End Sub
End Class