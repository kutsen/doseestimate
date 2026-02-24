Imports System.IO
Imports System.Collections
Imports System.Globalization

Public Class CalculationForm
    Dim memoryImage As Bitmap
    'Dim drPatientSelectn() As PatientDataSet.XCalc_graphRow
    Public Sub DataOutput()
        Me.EntranceDoseLabel.Text = checkVal(Math.Round(dblmAs * dblRateDev * 0.87 / 100 * Math.Pow(100 / 85, 2), 1)) '0,87/100 - наверное, перевод из Р в Гр. Math.Pow(100/85) - наверное, пересчет от точки, на которой измерен радиационный выход, к точке, на которой определяется входная доза
        Select Case intSex
            Case 2 'women first
                Me.UterusLabel.Visible = True 'display labels for female organ doses
                Me.UterusDoseLabel.Visible = True
                Me.EffectiveDose.Visible = True
                Me.EffectiveDoseLabel.Visible = True
                Me.EffectiveDoseLabel.Text = checkVal(dbEffDose)
                Me.GonadsLabel.Text = "1 " & My.Resources.Ovaries ' название гонад
                ' doses
                Me.GonadsDoseLabel.Text = checkVal(MainDosesFInterpFinal(0))
                Me.RBMDoseLabel.Text = checkVal(MainDosesFInterpFinal(1))
                Me.LungsDoseLabel.Text = checkVal(MainDosesFInterpFinal(3))
                Me.StomachDoseLabel.Text = checkVal(MainDosesFInterpFinal(4))
                Me.UrinaryBladderDoseLabel.Text = checkVal(MainDosesFInterpFinal(5))
                Me.BreastDoseLabel.Text = checkVal(MainDosesFInterpFinal(11))
                Me.LiverDoseLabel.Text = checkVal(MainDosesFInterpFinal(6))
                Me.OesophagusDoseLabel.Text = checkVal(MainDosesFInterpFinal(7))
                Me.ThyroidDoseLabel.Text = checkVal(MainDosesFInterpFinal(8))
                Me.SkinDoseLabel.Text = checkVal(MainDosesFInterpFinal(9))
                Me.EndosteumDoseLabel.Text = checkVal(MainDosesFInterpFinal(10))
                Me.BrainDoseLabel.Text = checkVal(MainDosesFInterpFinal(12))
                Me.LargeIntesineDoseLabel.Text = checkVal(MainDosesFInterpFinal(2))
                Me.AdrenalsDoseLabel.Text = checkVal(AdditionalDosesFInterpFinal(0))
                Me.GallbladderDoseLabel.Text = checkVal(AdditionalDosesFInterpFinal(9))
                Me.SmallIntestineDoseLabel.Text = checkVal(AdditionalDosesFInterpFinal(2))
                Me.KidneysDoseLabel.Text = checkVal(AdditionalDosesFInterpFinal(3))
                Me.MuscleDoseLabel.Text = checkVal(AdditionalDosesFInterpFinal(7))
                Me.ETDoseLabel.Text = checkVal(AdditionalDosesFInterpFinal(8))
                Me.PancreasDoseLabel.Text = checkVal(AdditionalDosesFInterpFinal(4))
                Me.SpleenDoseLabel.Text = checkVal(AdditionalDosesFInterpFinal(5))
                Me.ThymusDoseLabel.Text = checkVal(AdditionalDosesFInterpFinal(6))
                Me.UterusDoseLabel.Text = checkVal(AdditionalDosesFInterpFinal(12)) 'uterus
                Me.HeartDoseLabel.Text = checkVal(AdditionalDosesFInterpFinal(1))
                Me.LymphNodesDoseLabel.Text = checkVal(AdditionalDosesFInterpFinal(10))
                Me.OralMucosaDoseLabel.Text = checkVal(AdditionalDosesFInterpFinal(11))
                Me.SalivaryGlandsDoseLabel.Text = checkVal(MainDosesFInterpFinal(13))
            Case 1
                Me.EffectiveDose.Visible = True
                Me.EffectiveDoseLabel.Visible = True
                If dbEffDose < 0.001 Then
                    Me.EffectiveDoseLabel.Text = strLess
                Else
                    Me.EffectiveDoseLabel.Text = Math.Round(dbEffDose, 3).ToString(CultureInfo.CurrentUICulture)
                End If


                Me.UterusLabel.Text = "22 " & Global.RentgenCalculator.My.Resources.Resources.Prostate
                Me.UterusDoseLabel.Text = checkVal(AdditionalDosesMInterpFinal(12)) 'prostate
                'male case
                Me.GonadsDoseLabel.Text = checkVal(MainDosesMInterpFinal(0))
                Me.RBMDoseLabel.Text = checkVal(MainDosesMInterpFinal(1))
                Me.LungsDoseLabel.Text = checkVal(MainDosesMInterpFinal(3))
                Me.StomachDoseLabel.Text = checkVal(MainDosesMInterpFinal(4))
                Me.UrinaryBladderDoseLabel.Text = checkVal(MainDosesMInterpFinal(5))
                Me.LiverDoseLabel.Text = checkVal(MainDosesMInterpFinal(6))
                Me.OesophagusDoseLabel.Text = checkVal(MainDosesMInterpFinal(7))
                Me.ThyroidDoseLabel.Text = checkVal(MainDosesMInterpFinal(8))
                Me.SkinDoseLabel.Text = checkVal(MainDosesMInterpFinal(9))
                Me.EndosteumDoseLabel.Text = checkVal(MainDosesMInterpFinal(10))
                Me.BreastDoseLabel.Text = checkVal(MainDosesMInterpFinal(11))
                Me.LargeIntesineDoseLabel.Text = checkVal(MainDosesMInterpFinal(2))
                Me.AdrenalsDoseLabel.Text = checkVal(AdditionalDosesMInterpFinal(0))
                Me.BrainDoseLabel.Text = checkVal(MainDosesMInterpFinal(12))
                Me.HeartDoseLabel.Text = checkVal(AdditionalDosesMInterpFinal(1))
                Me.SmallIntestineDoseLabel.Text = checkVal(AdditionalDosesMInterpFinal(2))
                Me.KidneysDoseLabel.Text = checkVal(AdditionalDosesMInterpFinal(3))
                Me.MuscleDoseLabel.Text = checkVal(AdditionalDosesMInterpFinal(7))
                Me.ETDoseLabel.Text = checkVal(AdditionalDosesMInterpFinal(8))
                Me.PancreasDoseLabel.Text = checkVal(AdditionalDosesMInterpFinal(4))
                Me.SpleenDoseLabel.Text = checkVal(AdditionalDosesMInterpFinal(5))
                Me.ThymusDoseLabel.Text = checkVal(AdditionalDosesMInterpFinal(6))
                Me.GallbladderDoseLabel.Text = checkVal(AdditionalDosesMInterpFinal(9))
                Me.LymphNodesDoseLabel.Text = checkVal(AdditionalDosesMInterpFinal(10))
                Me.OralMucosaDoseLabel.Text = checkVal(AdditionalDosesMInterpFinal(11))
                Me.SalivaryGlandsDoseLabel.Text = checkVal(MainDosesMInterpFinal(13))
        End Select
        RentgenCalculator.MainForm.ProcedureDataGridView.DataSource = Nothing 'стирание данных, загруженных в табличку во вкладке "Аппарат"
        RentgenCalculator.MainForm.ProcTree.CollapseAll()
        'RentgenCalculator.MainForm.MaskedTextBox1.Text = "" 'стирание даты рождения 'убрано 6 июня 2019
    End Sub
    Private Sub CaptureScreen()
        Dim myGraphics As Graphics = Me.CreateGraphics()
        Dim s As Size = Me.Size
        memoryImage = New Bitmap(s.Width, s.Height, myGraphics)
        Dim memoryGraphics As Graphics = Graphics.FromImage(memoryImage)
        memoryGraphics.CopyFromScreen(Me.Location.X, Me.Location.Y, 0, 0, s)
    End Sub
    Private Sub printDocument1_PrintPage(ByVal sender As System.Object, _
          ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles _
          PrintDocument1.PrintPage
        e.Graphics.DrawImage(memoryImage, 0, 0)
    End Sub
    Private Sub PrintToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        CaptureScreen()
        PrintDocument1.Print()

    End Sub

    Private Sub ПечатьToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ПечатьToolStripMenuItem.Click
        CaptureScreen()
        PrintDocument1.Print()
    End Sub

    Private Sub ВФайлеToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ВФайлеToolStripMenuItem.Click
        Dim Gender As String
        Dim NumProc As Byte = 1 'defines the current number of procs for the current patient today
        Dim num As Integer = 0 'loop var for loop in records of Patients db
        Try

            If (intSex.Equals(1)) Or (MainForm.RadioButton1.Checked) Then 'the condition consists of two statements because it is not clear at what moment the variable was changed 
                Gender = "М" 'change encoding to numbers
            Else
                Gender = "Ж"
            End If
            'the text file is for debug only. should be removed for user
            Dim FileName As String = Environment.CurrentDirectory & "\PatientsList.txt"
            Dim strCurrentResult As String = (MainForm.DateTimePicker1.Value.ToString & Chr(9) & Global.RentgenCalculator.MainForm.MaskedTextBox1.Text & Chr(9) & strSurname & Chr(9) & strName & Chr(9) & strPatronymic & Chr(9) & Gender & Chr(9) & strProcedureName & Chr(9) & strSelectedDeviceName & Chr(9) & dblPowerCh & Chr(9) & dblmAs & Chr(9) & strRipCh & Chr(9) & strWidthCh & Chr(9) & dbHeight & Chr(9) & Math.Round(dbEffDose, 4).ToString(CultureInfo.CurrentUICulture))
            Dim strWriter1 As New System.IO.StreamWriter(FileName, True, System.Text.Encoding.GetEncoding(1251))
            strWriter1.WriteLine(strCurrentResult)
            strWriter1.Close()
            'start check whether the patient has been irradiated that day more than once
            If (File.GetAttributes(strPatientsdbname) <> FileAttributes.ReadOnly) Then
                da4.Fill(dt4)
                'da4.Update(drInsert) 'here can be an error if the db file is read-only
            End If
            'просмотр базы данных и проверка, есть ли в ней данная запись
            Do While num < dt4.Rows.Count 'incremetrs NumProc if needed
                'check for the calculation number for today should be done
                If ((strName = dt4.Rows(num).Item(7)) And (strSurname = dt4.Rows(num).Item(8)) And (strPatronymic = dt4.Rows(num).Item(9)) And (dt4.Rows(num).Item(3) = MainForm.DateTimePicker1.Text)) Then 'добавить проверку на РДА
                    NumProc = NumProc + 1
                End If
                num = num + 1
            Loop
            'PREPARE to write to database;
            'CHECK WHETHER THE DB IS NOT OPEN
            drInsert = dt4.NewRow() 'allocate memory for a new row
            Dim Statecheck As Integer = da4.Connection.State
            'end check whether the patient has been irradiated that day more than once
            drInsert("ID") = strHospCode & "_" & strSelectedDeviceName & "_" & strRegnum & "_" & Today.ToString(CultureInfo.CurrentUICulture) & "_" & NumProc 'ID
            drInsert("HOSPITAL") = strHospCode 'Hospital code
            drInsert("MEASDATE") = MainForm.DateTimePicker1.Value.Date 'PRESUMED that the calculation is made at the day of the proc
            drInsert("TYPEPROC") = Split(strProcedureName, "\")(1) & "_" & ExtrctProjection(Split(strProcedureName, "\")(2)) 'taking first part of the splitted line
            drInsert("NUMPROC") = NumProc 'incremented number of examinations a day
            drInsert("REGNUM") = strRegnum 'Must be patient's registration card number-integer. Why not won't we choose it automatically if it is already in db?
            drInsert("NAME") = strName 'Patient's name
            drInsert("FAMNAME") = strSurname 'Patient's family name
            drInsert("PATR") = strPatronymic 'Patient's patronymic
            drInsert("BIRTDATE") = Convert.ToDateTime(MainForm.MaskedTextBox1.Text) 'Patient's birthdate
            drInsert("APPARAT") = strSelectedDeviceName 'Name of the used apparatus
            drInsert("VOLTAGE") = dblPowerCh 'double precised Voltage 'How is it casted?
            drInsert("MAS") = dblmAs 'Exposure of double type
            drInsert("FOCDIST") = strRipCh 'Focal distance of double type
            drInsert("WIDTH") = strWidthCh 'Width of the field-fixed of integer type
            drInsert("HEIGHT") = dbHeight 'Height of the field
            drInsert("EFFDOSE") = dbEffDose 'write eff dose without rounding
            RentgenCalculator.MainForm.MaskedTextBox1.Text = "" 'стирание даты рождения 'добавлено 6 июня 2019
        Catch ex As Exception
            MsgBox("Ошибка при записи информации" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, My.Resources.MainTitle)
            Return

        End Try
        If (File.GetAttributes(strPatientsdbname) = FileAttributes.ReadOnly) Then 'to avoid an error if the db file is read-only
            MsgBox("Файл базы данных доступен только для чтения, невозможно сохранить данные.", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
        Else
            num = 0
            Try
                dt4.Rows.Add(drInsert)
                ' And (Math.Round(dbEffDose, 10) = Math.Round(dt4.Rows(num).Item(17), 10)))'useful for checking duplicates
            Catch CE As System.Data.ConstraintException
                MsgBox("Невозможно записать результат в связи с ограничениями, наложенными на повторяюшиеся записи.", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
            Catch ex As NoNullAllowedException
                MsgBox(ex.Message & vbCrLf & "Возможно, вы только что уже записали это исследование.", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
            End Try
            Try
                da4.Update(drInsert)
            Catch ex As OverflowException
                MsgBox("Переполнение. Не удалось сохранить результаты расчета доз в базе данных.", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
            End Try
            MsgBox("Данные успешно сохранены", MsgBoxStyle.Information, My.Resources.MainTitle)
            'Me.Close() 'Эта строка удалена из таблицы и не содержит данных. BeginEdit() позволит создать в этой строке новые данные.
        End If
    End Sub

End Class