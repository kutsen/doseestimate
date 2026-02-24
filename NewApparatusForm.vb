Public Class NewApparatusForm

    Dim drSelection() As DataDataSet.DevicesListRow
    Dim drInsert As DataRow
    Private Sub DiagnisticRadioBut_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DiagnosticRadioBut.CheckedChanged
        intDevType = 1
        Me.DevicesListBox.DataSource = dt.Select("DeviceType=" & intDevType)
        Me.DevicesListBox.DisplayMember = "DeviceName"
        Me.DevicesListBox.Focus()
    End Sub
    Private Sub MobileRadioBut_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MobileRadioBut.CheckedChanged
        intDevType = 2
        Me.DevicesListBox.DataSource = dt.Select("DeviceType=" & intDevType)
        Me.DevicesListBox.DisplayMember = "DeviceName"
    End Sub
    Private Sub FluoroRadioBut_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FluoroRadioBut.CheckedChanged
        intDevType = 3
        Me.DevicesListBox.DataSource = dt.Select("DeviceType=" & intDevType)
        Me.DevicesListBox.DisplayMember = "DeviceName"
    End Sub
    Private Sub AddButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddButton.Click
        If Me.DevicesListBox.SelectedIndices.Count > 0 Then
            For i = 0 To Me.DevicesListBox.SelectedIndices.Count - 1

                strDeviceName = Me.DevicesListBox.SelectedItems(i)("DeviceName")

                drSelection = dt.Select("DeviceName=" & "'" & strDeviceName & "'") '??? 
                Me.DevicesInUseListBox.Items.Add(Me.DevicesListBox.SelectedItems(i)("DeviceName")) 'errorfull string
                dblFilter = drSelection(0)("Filter")
                dblPower = drSelection(0)("Power")
                dblRip = drSelection(0)("Rip")
                dblYield = drSelection(0)("Yield")

                drInsert = dt2.NewRow
                Dim account1 As Integer = MainForm.DataGridView1.Rows.Count
                drInsert("DeviceNumber") = account1 + i
                drInsert("DeviceName") = strDeviceName
                drInsert("Filter") = dblFilter
                drInsert("Power") = dblPower
                drInsert("Rip") = dblRip
                drInsert("Yield") = dblYield

                dt2.Rows.Add(drInsert)
                da2.Update(drInsert)

                dblFilter = Nothing
                dblPower = Nothing
                dblRip = Nothing
                dblYield = Nothing
                drInsert = Nothing
                drSelection = Nothing
            Next i
            da2.Fill(dt2)
            MainForm.DataGridView1.DataSource = dt2
            Me.DevicesListBox.ClearSelected()
        Else
            MsgBox("Выберите рентгеновский аппарат из списка!", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
        End If
    End Sub
    Private Sub RemoveButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveButton.Click
        'remove selected from ListBox2
        If (Me.DevicesInUseListBox.Items.Count > 0) Then 'check if there anything at all in ListBox2
            'Dim ListBox2Cont As String 'one way is writing all rows to array, eliminate selected and write array to ListBox2
            If (Me.DevicesInUseListBox.SelectedIndices.Count > 0) Then 'check if there anything selected
                'ListBox2Cont = 
                'Me.ListBox2.Items(Me.ListBox2.SelectedIndices.Count)
            End If
            '1
        End If
    End Sub
    Public Sub OkButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKButton.Click
        'writing opts to RentgenCalculator.ini
        Using sw As System.IO.StreamWriter = New System.IO.StreamWriter("RentgenCalculator.ini")
            ' Add text to file
            sw.Write("HospCode=" + OrganizationTextBox.Text)
            sw.Close()
        End Using
        Me.Close()
    End Sub
End Class