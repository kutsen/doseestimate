'Option Explicit On Option Strict On
Imports System.Collections
Imports System.Globalization

<Assembly: System.Reflection.AssemblyKeyFileAttribute("sgKey.snk")> 
<Assembly: System.CLSCompliant(True)> 

Public Class MainForm
    'where is the definition of a row in my devices table?
    Dim drDeleted As DataRow
    Dim dblFilter2 As Double
    Dim maxAge As Short = 120
    Dim MaxVoltage As Short = 140 ' starting min value of voltage
    Dim MinVoltage As Short = 40 ' starting max value of voltage
    Dim ToolTip1 As New ToolTip

    Private Sub MainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        da2.Fill(dt2) 'load list of used devices if any
        Dim account As Integer = dt2.Rows.Count 'count the number of devices included into the list
        da.Fill(dt) 'load full list of devices
        da1.Fill(dt1) 'load dose data table - DataForCalculation
        da3.Fill(dt3) 'load input table - ProceduresList
        If Not System.IO.File.Exists(strFileName) Then
            Try
                Dim sr0 As System.IO.Stream = System.IO.File.Open(strFileName, IO.FileMode.CreateNew)
                sr0.Close()
            Catch Exception As System.Exception
                System.IO.File.Create(strFileName)
                MsgBox("Файл " & Chr(34) & strFileName & Chr(34) & " отсутствует. Он был создан заново. Убедитесь в правильности задания ОКПО.", MsgBoxStyle.Exclamation, My.Resources.MainTitle) 'или расчет может быть произведен, но не сохраняться в базу
                Return
            End Try
        End If 'check whether the file exists
        Using sr1 As System.IO.StreamReader = System.IO.File.OpenText(strFileName)
            Try
                strHospCode = sr1.ReadLine.Split(New [Char]() {"="c})(1)
            Catch NRE As System.NullReferenceException
                MsgBox("Файл " & Chr(34) & strFileName & Chr(34) & " испорчен. Необходима переустановка программы.", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                NewApparatusForm.Show()
                'здесь бы записать в файл и тут же прочесть из него
            End Try
            sr1.Close()
            NewApparatusForm.OrganizationTextBox.Text = strHospCode
        End Using
        Me.ProcedureDataGridView.DataSource = Nothing
        'account is 0 if no X-ray machines are in use.
        Select Case account
            Case 0
                'show settings Dialog to add a new apparatus
                NewApparatusForm.ShowDialog()
                'wait until Dialog ends
                Me.TabControl1.SelectTab(1)
                Me.DataGridView1.AllowUserToAddRows = True
                Me.XRayUnitListLabel.Text = My.Resources.CreateApListLabel '"Создание списка используемых рентгеновских аппаратов"
            Case Is > 0
                Me.DataGridView1.DataSource = dt2
        End Select
        MaskedTextBox1.ValidatingType = GetType(System.DateTime) '?
    End Sub
    Private Sub ExitButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitButton.Click
        Application.Exit()
    End Sub
    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RegistryNumTextBox.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter, Keys.Tab, Keys.Down
                Me.FamilyNameTextBox.Focus()
        End Select
    End Sub
    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles FamilyNameTextBox.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter, Keys.Tab, Keys.Down
                Me.GivenNameTextBox.Focus()
            Case Keys.Up
                Me.RegistryNumTextBox.Focus()
        End Select
    End Sub
    Private Sub TextBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles GivenNameTextBox.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter, Keys.Tab, Keys.Down
                Me.PatronymicTextBox.Focus()
            Case Keys.Up
                Me.FamilyNameTextBox.Focus()
        End Select
    End Sub
    Private Sub TextBox4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles PatronymicTextBox.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter, Keys.Tab, Keys.Down
                Me.MaskedTextBox1.Focus()
            Case Keys.Up
                Me.GivenNameTextBox.Focus()
        End Select
    End Sub
    Private Sub HeightTextBox_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles HeightTextBox.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter, Keys.Tab, Keys.Down
                Me.PatientWeightTextBox.Focus()
            Case Keys.Up
                Me.MaskedTextBox1.Focus()

        End Select
    End Sub
    Private Sub WeightTextBox_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles PatientWeightTextBox.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter, Keys.Tab, Keys.Down
                Me.RunButton.Focus()
            Case Keys.Up
                Me.HeightTextBox.Focus()
        End Select
    End Sub
    Private Sub RadioButton1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If Me.RadioButton1.Checked Then
            intSex = 1
            GroupBox1.ForeColor = Color.Black 'color change
        End If
        'по-видимому, проверка, какие процедуры загружены во вкладке "Процедуры"
        ' при расчете доз эффективная доза выдается на оба фантома, но органные дозы только на соответствующий
        'If strProjectionCode <> Nothing Then
        ' Me.DataGridView2.DataSource = dt3.Select("ProjectionCode=" & "'" & strProjectionCode & "'") '& " and Group=" & intSex)
        ' End If
    End Sub
    Private Sub RadioButton1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RadioButton1.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter, Keys.Tab, Keys.Down
                Me.RadioButton2.Focus()
            Case Keys.Up
                Me.PatientWeightTextBox.Focus()
        End Select
    End Sub
    Private Sub RadioButton2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If Me.RadioButton2.Checked Then
            intSex = 2
            GroupBox1.ForeColor = Color.Black
        End If
        'If strProjectionCode <> Nothing Then
        'Me.DataGridView2.DataSource = dt3.Select("ProjectionCode=" & "'" & strProjectionCode & "'") ' & " and Group=" & intSex)
        'End If

    End Sub
    Private Sub RadioButton2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RadioButton2.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter, Keys.Tab, Keys.Down
                Me.TabControl1.SelectTab(1)
            Case Keys.Up
                Me.RunButton.Focus()
        End Select
    End Sub
    Private Sub MaskedTextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MaskedTextBox1.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter, Keys.Tab, Keys.Down
                Me.HeightTextBox.Focus()
            Case Keys.Up
                Me.PatronymicTextBox.Focus()
        End Select
    End Sub
    Private Sub MaskedTextBox1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles MaskedTextBox1.Leave
        'convenient variable for handling of data ' 23.5.2019
        Dim separator As String = CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator
        Dim dtBirthDate As Date
        If (Me.MaskedTextBox1.Text <> "  " & separator & "  " & separator) And (Me.MaskedTextBox1.TextLength = 10) And MaskedTextBox1.MaskCompleted Then
            If DateTime.TryParse(MaskedTextBox1.Text, dtBirthDate) Then
                If Convert.ToDateTime(MaskedTextBox1.Text, CultureInfo.CurrentUICulture) < Date.Now Then
                    Dim MaskedTextBox1Split As String() = Me.MaskedTextBox1.Text.Split(separator) 'reading birth date.
                    intAge = DateDiff(DateInterval.Year, Convert.ToDateTime(MaskedTextBox1.Text, CultureInfo.CurrentUICulture), Now()) ' присвоение переменной полных лет
                    MaskedTextBox1.ForeColor = Color.Black
                Else
                    ToolTip1.Show(Global.RentgenCalculator.My.Resources.TipDate, MaskedTextBox1, 2000)
                    MaskedTextBox1.ForeColor = Color.Red
                End If
            Else
                ToolTip1.Show(My.Resources.TipWrongDateFormat, MaskedTextBox1, 2000)
                MaskedTextBox1.ForeColor = Color.Red
                Exit Sub
            End If ' попытка преобразовать число в формат даты
        End If ' проверка, заполнено ли значение.
    End Sub
    Private Sub TextBox2_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles FamilyNameTextBox.Leave
        'здесь может быть масса проблем: в том числе переносы на новую строку.
        If Not FamilyNameTextBox.Text = "" Then
            strSurname = Me.FamilyNameTextBox.Text
        Else
            ToolTip1.Show(Global.RentgenCalculator.My.Resources.Resources.MsgEmpty & Global.RentgenCalculator.My.Resources.Resources.FamilyNameLabel, FamilyNameTextBox, 2000)
        End If
    End Sub
    Private Sub TextBox3_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles GivenNameTextBox.Leave
        If Not GivenNameTextBox.Text = "" Then
            strName = Me.GivenNameTextBox.Text
        Else
            ToolTip1.Show("Введите имя пациента.", GivenNameTextBox, 2000)
        End If
    End Sub
    Private Sub TextBox4_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PatronymicTextBox.Leave
        If Not PatronymicTextBox.Text = "" Then
            strPatronymic = Me.PatronymicTextBox.Text
        Else
            ToolTip1.Show("Введите отчество пациента", GivenNameTextBox, 2000)
        End If
    End Sub
    Private Sub HeightTextBox_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        ' since 4 Feb 2022 is not called; the value is not used;
        If Len(Me.HeightTextBox.Text) > 0 Then
            Try
                If Convert.ToDecimal(HeightTextBox.Text, CultureInfo.CurrentUICulture) > MaxHeight Or Convert.ToDecimal(HeightTextBox.Text, CultureInfo.CurrentUICulture) < MinHeight Then
                    ToolTip1.ToolTipTitle = "Недействительное значение." 'это показывается не сразу
                    ToolTip1.Show(strHeightMsg, HeightTextBox, 2000)
                    HeightTextBox.ForeColor = Color.Red
                Else
                    Try
                        ToolTip1.Hide(Nothing)
                    Catch ANE As ArgumentNullException
                        ToolTip1.Hide(Me)
                    End Try
                    HeightTextBox.ForeColor = Color.Black
                    sngHeight = Convert.ToSingle(Me.HeightTextBox.Text, CultureInfo.CurrentUICulture)
                End If
            Catch SFE As System.FormatException
                ToolTip1.ToolTipTitle = "Недействительное значение." 'это показывается не сразу
                ToolTip1.Show("Неверно введен рост пациента.", HeightTextBox, 2000)
                HeightTextBox.ForeColor = Color.Red
            End Try
        End If
    End Sub
    Private Sub WeightTextBox_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PatientWeightTextBox.Leave
        Dim WeightTextToolTip As New ToolTip
        If Len(Me.PatientWeightTextBox.Text) > 0 Then
            Try
                If Convert.ToDecimal(Me.PatientWeightTextBox.Text, CultureInfo.CurrentUICulture) > 150 Or Convert.ToDecimal(Me.PatientWeightTextBox.Text, CultureInfo.CurrentUICulture) < 3 Then
                    WeightTextToolTip.Show(My.Resources.MsgWrongWeight, PatientWeightTextBox, 5000)
                    'Me.WeightTextBox.Text = Nothing
                    Me.PatientWeightTextBox.ForeColor = Color.Red
                    Me.PatientWeightTextBox.Focus()
                Else
                    sngWeight = Convert.ToDouble(Me.PatientWeightTextBox.Text, CultureInfo.CurrentUICulture)
                    Me.PatientWeightTextBox.ForeColor = Color.Black
                End If
            Catch SFE As System.FormatException
                'MsgBox("Значение веса введено неверно.", MsgBoxStyle.Exclamation, My.Resources.MainTitle) 'SFE.Message
                WeightTextToolTip.ToolTipTitle = "Недействительное значение."
                WeightTextToolTip.Show(My.Resources.MsgWrongWeight, PatientWeightTextBox, 5000)
                PatientWeightTextBox.ForeColor = Color.Red
            End Try
        End If
    End Sub
    Private Sub MaskedTextBox1_MaskInputRejected(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MaskInputRejectedEventArgs) Handles MaskedTextBox1.MaskInputRejected
        Dim tooltip3 As New ToolTip
        tooltip3.ToolTipTitle = "Недействительный ввод"
        tooltip3.Show(My.Resources.MsgDate1, MaskedTextBox1, 5000)
    End Sub
    Private Sub MainForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If Me.ProcedureDataGridView.Focused Then

            Select Case e.KeyCode
                Case Keys.Enter
                    SendKeys.Send("{Tab}")
                    e.Handled = True
            End Select
        End If
    End Sub
    Private Sub RunButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RunButton.Click
        If Not RegistryNumTextBox.Text = "" Then 'проверка присвоено ли и заполнено: добавлено 18.2.2019
            strRegnum = Me.RegistryNumTextBox.Text 'ввод регистрационного номера  'добавлено 18.2.2019
        Else
            MsgBox("Введите регистрационный номер!", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
            TabControl1.SelectTab(0)
            RegistryNumTextBox.Focus()
            Exit Sub
        End If 'поле "Регистрационный номер" не пустое
        If intSex = Nothing Then
            MsgBox("Не задан пол пациента.", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
            Me.TabControl1.SelectTab(0)
            Me.GroupBox1.ForeColor = Color.Red
            Exit Sub
        Else
        End If
        'начало проверки роста
        '4 Feb 2022: is not neccessary for now
        'CheckHeight()
        'конец проверки роста
        'начало проверки фамилии
        If Not FamilyNameTextBox.Text = "" Then
            strSurname = Me.FamilyNameTextBox.Text ' внутренняя переменная значения, введенного в строку TextBox2 (фамилия)
        Else
            MsgBox("Введите фамилию пациента.", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
            Me.TabControl1.SelectTab(0)
            FamilyNameTextBox.Focus()
            Exit Sub
        End If
        'конец проверки фамилии
        'начало проверки имени
        If Not GivenNameTextBox.Text = "" Then
            TextBox3_Leave(Me, Nothing) 'имя
        Else
            MsgBox("Введите имя пациента.", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
            Me.TabControl1.SelectTab(0)
            GivenNameTextBox.Focus()
            Exit Sub
        End If
        'конец проверки имени
        'начало проверки очества
        If Not PatronymicTextBox.Text = "" Then
            TextBox4_Leave(Me, Nothing) 'отчество
        Else
            MsgBox("Введите отчество пациента.", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
            Me.TabControl1.SelectTab(0)
            PatronymicTextBox.Focus()
            Exit Sub
        End If
        'конец проверки отчества
        'начало проверки веса
        If sngWeight = Nothing Then
            If PatientWeightTextBox.Text = "" Then
                MsgBox("Введите вес пациента.", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                TabControl1.SelectTab(0)
                Me.PatientWeightTextBox.Focus()
                Exit Sub
            Else 'строка веса введена
                Try 'попытка ввести вес
                    If Convert.ToDecimal(Me.PatientWeightTextBox.Text, CultureInfo.CurrentUICulture) > 150 Or Convert.ToDecimal(Me.PatientWeightTextBox.Text, CultureInfo.CurrentUICulture) < 3 Then
                        MsgBox(My.Resources.MsgWrongWeight, MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                        Me.PatientWeightTextBox.ForeColor = Color.Red
                        Me.TabControl1.SelectTab(0)
                        Me.PatientWeightTextBox.Focus()
                        Exit Sub
                    Else
                        Me.PatientWeightTextBox.ForeColor = Color.Black
                        sngWeight = Convert.ToDouble(Me.PatientWeightTextBox.Text, CultureInfo.CurrentUICulture)
                    End If
                Catch SFE As System.FormatException
                    MsgBox("Ошибка чтения веса", MsgBoxStyle.Exclamation, My.Resources.MainTitle) 'SFE.Message
                    PatientWeightTextBox.ForeColor = Color.Red
                    PatientWeightTextBox.Focus()
                End Try
            End If ' вес не введен
        End If ' вес не определен
        'patient weight check end
        'ввод РИПа для процедуры
        If Not FIDComboBox.Text Is Nothing Then
            If Not FIDComboBox.Text.Length = 0 Then
                Try
                    strRipCh = Convert.ToDouble(FIDComboBox.Text, CultureInfo.CurrentUICulture)
                Catch SFE As System.FormatException
                    MsgBox("Выберите РИП", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                    TabControl1.SelectTab(1)
                    FIDComboBox.Focus()
                    Exit Sub
                End Try
            Else
                MsgBox("Не введен РИП.", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                TabControl1.SelectTab(1)
                FIDComboBox.Focus()
                Exit Sub
            End If
        End If
        'конец ввода РИПа
        'начало ввода напряжения
        Try
            If VoltageProcTextBox.Text = "" Then
                MsgBox("Введите напряжение на рентгеновской трубке", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                VoltageProcTextBox.BackColor = Color.Red
                TabControl1.SelectTab(1)
                VoltageProcTextBox.Focus()
                Exit Sub
            Else
                dblPowerCh = Convert.ToDouble(VoltageProcTextBox.Text)
                VoltageProcTextBox.BackColor = Color.White
            End If
        Catch SFE As System.FormatException
            MsgBox("Напряжение введено не верно. Должно быть число.", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
            TabControl1.SelectTab(1)
            VoltageProcTextBox.ForeColor = Color.Red
            Exit Sub
        End Try
        'конец ввода напряжения
        ' Работа с базой
        Try
        Catch NRE As System.NullReferenceException
            MsgBox("Не задана процедура обследования.", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
            Exit Sub
        End Try
        'проверка фильтра
        If FilterTextBox.Text = "" Then
            MsgBox("Не введена толщина фильтра рентгеновской трубки.", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
            TabControl1.SelectTab(1)
            FilterTextBox.BackColor = Color.Red
            FilterTextBox.Focus()
            Exit Sub
        Else
            If Double.TryParse(FilterTextBox.Text, dblFilter2) Then
                strFilterCh = Convert.ToDouble(FilterTextBox.Text) 'задание значение фильтра излучения
                If (strFilterCh < 0.1) Or (strFilterCh > 6) Then
                    MsgBox("Значение фильтра должно лежать в пределах от 0.1 до 6 мм.", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                    FilterTextBox.ForeColor = Color.Red
                    TabControl1.SelectTab(1)
                    FilterTextBox.Focus()
                    Exit Sub
                Else
                    FilterTextBox.ForeColor = Color.Black
                    FilterTextBox.BackColor = Color.White
                End If
            Else
                MsgBox("Толщина фильтра введена неверно.", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                FilterTextBox.ForeColor = Color.Red
                Exit Sub
            End If 'попытка перевести введенное значение фильтра в число
        End If ' если фильтр не введен в поле.
        '  выборка радиационного выхода
        If dblmAs = Nothing Then
            If mAsTextBox.Text = "" Then ' проверить поле, если заполнено
                MsgBox("Не задано значение мАс!", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                Exit Sub
            Else
                If Not Double.TryParse(mAsTextBox.Text, dblmAs) Then
                    MsgBox("Неверный формат данных. Введите число.", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                    mAsTextBox.ForeColor = Color.Red
                    Exit Sub
                Else
                    If dblmAs < 0.1 OrElse dblmAs > 1000 Then
                        MsgBox("Неверное значение мАс. Величина мАс не может быть меньше 0,1 или больше 1000.", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                        Exit Sub
                    Else
                        mAsTextBox.ForeColor = Color.Black
                        dblmAs = Convert.ToDouble(mAsTextBox.Text, CultureInfo.CurrentUICulture)
                    End If 'произведение тока на выдержку находится в разрешенном диапазоне
                End If 'неверный формат произведения тока на выдержку
            End If ' произвдение тока на выдержку не введено
        End If ' произвдение тока на выдержку не введено
        'проверка информации о дате рождения
        If MaskedTextBox1.MaskCompleted = True Then
            If Not IsDate(MaskedTextBox1.Text) Then
                MsgBox(My.Resources.TipWrongDateFormat, MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                MaskedTextBox1.ForeColor = Color.Red
                Exit Sub
            Else
                If Convert.ToDateTime(MaskedTextBox1.Text, CultureInfo.CurrentUICulture) > Date.Now Then
                    MsgBox(My.Resources.TipDate, MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                    MaskedTextBox1.ForeColor = Color.Red
                    TabControl1.SelectTab(0)
                    MaskedTextBox1.Focus()
                    Exit Sub
                Else
                    MaskedTextBox1.ForeColor = Color.Black
                    intAge = DateDiff(DateInterval.Year, Convert.ToDateTime(MaskedTextBox1.Text, CultureInfo.CurrentUICulture), Now())
                End If 'дата меньше сегодняшней
            End If 'Дата правильная
        End If ' Поле "дата рождения" заполнена полностью
        'чтение размеров поля
        If FieldSizeComboBox.Text.Length > 0 Then
            Dim FieldSize As String() = FieldSizeComboBox.Text.Split("x")
            If FieldSize.Length = 2 Then
                Try
                    strWidthCh = FieldSize(0)
                    dbHeight = FieldSize(1)
                Catch FE As FormatException
                    MsgBox("Неверный формат поля. Введите поле в формате 00x00. Размеры поля указывайте в сантиметрах", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                    TabControl1.SelectTab(1)
                    FieldSizeComboBox.Focus()
                    Exit Sub
                End Try
            Else
                MsgBox("Поле облучения введено неверно. Введите поле в формате 00x00. Размеры поля указывайте в сантиметрах", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                TabControl1.SelectTab(1)
                FieldSizeComboBox.Focus()
                Exit Sub
            End If
        Else
            MsgBox("Не введено поле облучения. Введите поле.", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
            TabControl1.SelectTab(1)
            FieldSizeComboBox.Focus()
            Exit Sub
        End If ' проверка, введено ли поле
        ' выбранные в ниспадающих списках ширина, высота поля и РИП
        If RippleTextBox.Text = "" Then
            MsgBox("Не указана пульсация напряжения.", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
        Else
            Dim tmpdblRipple As Double
            If Double.TryParse(RippleTextBox.Text, tmpdblRipple) Then
                If Convert.ToDouble(RippleTextBox.Text) < 0 Or Convert.ToDouble(RippleTextBox.Text) > 7 Then
                    MsgBox("Введено неверное значение пульсации напряжения. Значение может быть в пределах от 0 до 7", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                    RippleTextBox.ForeColor = Color.Red
                Else
                    dblRipple = RippleTextBox.Text
                    RippleTextBox.ForeColor = Color.Black
                End If
            Else
                dblRipple = -10000000.0
                MsgBox("Пульсация напряжения введена неверно", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                TabControl1.SelectTab(1)
                RippleTextBox.Focus()
                RippleTextBox.ForeColor = Color.Red
                Exit Sub
            End If 'попытка прочесть значение
        End If 'пульсация напряжения не введена
        'обработка данных, полученных из БД
        If dblYield = Nothing And dblFilter = Nothing And dblPower = Nothing And dblRip = Nothing Then
            MsgBox("Не выбран рентгеновский аппарат из списка!", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
            TabControl1.SelectTab(2)
            Me.DataGridView1.Focus()
            Exit Sub
        Else
            If dblYield = Nothing Then
                MsgBox("Не задано значение радиационного выхода рентгеновского аппарата!", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                Exit Sub
            End If
            If dblPower = Nothing Then
                MsgBox("Не задано значение напряжения на аноде трубки рентгеновского аппарата, " & Chr(13) & "при котором измерялся радиационный выход!", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                Exit Sub
            End If
            If dblRip = Nothing Then
                MsgBox("Не задано значение РИП, при котором измерялся радиационный выход аппарата!", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                Exit Sub
            End If
            If dblFilter = Nothing Then
                MsgBox("Не задана величина фильтра, при котором измерялся радиационный выход!", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                Exit Sub
            End If
        End If
        'dblRateDev – радиационный выход в мР/с
        dblRateDev = transformRate(dblYield, dblRip, strRipCh, dblFilter, strFilterCh, dblPower, dblPowerCh) * Math.Pow(dblRip / 100, 2)
        ' поскольку радиационный выход задается в мР/с, необходимо пересчитать рентгены в Греи
        DoseCoeff = 0.001 * dblRateDev * dblmAs
        Try
            ' Интерполяция по фильтру, напряжению и пульсации напряжения
            AbsorbedDosesInterp()
        Catch noRipple As Exception
            MsgBox("Не указана пульсация напряжения.", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
            TabControl1.SelectTab(1)
            RippleTextBox.Focus()
            Exit Sub
        End Try
        CalculationForm.InfoOutputLabel.Text = "Регистрационный номер пациента: " + strRegnum + Chr(13) + strSurname + " " + strName + " " + strPatronymic + Chr(13) + "Дата обследования " & DateTimePicker1.Value.ToShortDateString & Chr(13) & "аппарат " _
& strSelectedDeviceName & Chr(13) & strProcedureName & Chr(13) & "Параметры: " & dblmAs & " мАс, РИП=" & strRipCh & " см, напряжение=" & dblPowerCh.ToString & " кВ, " & Chr(13) & "поле облучения " & strWidthCh & " см x " & dbHeight.ToString & " см, фильтр=" & strFilterCh & " мм Al" & Chr(13) & "рад. выход=" & Math.Round(dblRateDev, 2).ToString & " мР/(мАс)" & Chr(13)
        CalculationForm.Label_optimal_dose_value.Text = DRLCheck()
        CalculationForm.DataOutput()
        ClearInput()
        FilterTextBox.Text = ""
        mAsTextBox.Text = ""
        VoltageProcTextBox.Text = ""
        RippleTextBox.Text = ""
        intAge = Nothing
        HeightTextBox.Text = ""
        sngHeight = Nothing
        PatientWeightTextBox.Text = ""
        sngWeight = Nothing
        'стереть регистрационный номер, чтобы это не пришлось делать пользователю
        RegistryNumTextBox.Text = ""
        ' Стереть фамилию, чтобы это не пришлось делать пользователю
        FamilyNameTextBox.Text = ""
        ' Стереть имя, чтобы это не пришлось делать пользователю
        GivenNameTextBox.Text = ""
        ' Стереть отчество, чтобы это не пришлось делать пользователю
        PatronymicTextBox.Text = ""
        'это поле должно быть, т.к. при расчете следующей процедуры ввод данных начинается с ФИО пациента
        TabControl1.SelectTab(0)
        CalculationForm.Show()
        CalculationForm.Focus()
    End Sub

    Private Sub ListBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Select Case intDevType
            Case 0
                MsgBox("Не указан тип аппарата", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
        End Select
    End Sub
    Private Sub ToolStripLabel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel1.Click
        Me.DataGridView1.AllowUserToAddRows = True
        NewApparatusForm.ShowDialog()
        Me.Width = 965
    End Sub
    Private Sub ToolStripLabel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel2.Click
        'if there are rows with empty cells this deletes them too
        Me.DataGridView1.AllowUserToDeleteRows = True
        'add procedure that empty dblPower, dblRip, dblYield and dblFilter if 
        Dim intCurrRow As Integer
        intCurrRow = Me.DataGridView1.CurrentRow.Index
        If intCurrRow < 0 Then
            MsgBox("Не выбрана строка таблицы для удаления", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
            Return
        Else
            If MsgBox("Из таблицы будет удалена строка. После нажатия кнопки ОК отмена удаления будет невозможна.", MsgBoxStyle.OkCancel, My.Resources.MainTitle) = MsgBoxResult.Ok Then
                drDeleted = dt2.Rows(intCurrRow)
                drDeleted.Delete()
                da2.Update(drDeleted)
                ds.AcceptChanges()
                da2.Fill(dt2)
                dblYield = Nothing
                dblPower = Nothing
                dblRip = Nothing
                dblFilter = Nothing
                'disagree to delete row
            Else
                ds.RejectChanges()
            End If
        End If
    End Sub
    Private Sub ToolStripLabel3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel3.Click
        'сохранить изменения в списке аппаратов
        Dim changedRecords As DataTable = dt2.GetChanges(Data.DataRowState.Modified)
        If (Not changedRecords Is Nothing) Then
            'Нарушение параллелизма: UpdateCommand затронула 0 из ожидаемых 1 записей
            da2.Update(changedRecords)
            ds.AcceptChanges()
        End If
        Dim addedRecords As DataTable = dt2.GetChanges(Data.DataRowState.Added)
        If Not addedRecords Is Nothing Then
            da2.Update(addedRecords)
        End If
        ds.AcceptChanges()
        changedRecords = Nothing
    End Sub
    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        Dim openFileDialog1 As New OpenFileDialog()
        myStream = Nothing
        openFileDialog1.InitialDirectory = Environment.CurrentDirectory 'the place from where the prog has been launched
        openFileDialog1.Filter = "txt files (*.txt)|*.txt"
        openFileDialog1.FilterIndex = 2
        openFileDialog1.RestoreDirectory = True

        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                myStream = openFileDialog1.OpenFile()
                strFileName = openFileDialog1.FileName
                If (myStream IsNot Nothing) Then
                    Dim streamReader1 As New System.IO.StreamReader(myStream, System.Text.Encoding.GetEncoding(1251), False) 'It has never been here
                    strTextFile = streamReader1.ReadToEnd
                    myStream.Close()
                    TextFileForm.Show()
                End If
            Catch Ex As Exception
                MsgBox(My.Resources.MsgFileOpenError & Ex.Message, MsgBoxStyle.Exclamation, My.Resources.MainTitle)
            Finally
                If (myStream IsNot Nothing) Then
                    myStream.Close()
                End If
            End Try
        End If
    End Sub
    Private Sub ProcTree_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles ProcTree.AfterSelect
        Dim strTempsrc, strTempDest As String ' temp strings for storage of loop values of FID
        Dim strTempsrc2 As String
        ' service flag for check whether the current value of FID is already written in ComboBox
        Dim flag As Boolean = False
        Dim flagFieldSize As Boolean = False '
        strProjectionCode = Me.ProcTree.SelectedNode.Tag ' get the procedure tag
        If strProjectionCode = Nothing Then
            Return
        End If
        strProcedureName = Me.ProcTree.SelectedNode.FullPath.ToString
        ClearInput() 'empty the Comboboxes
        If Not strProjectionCode Is Nothing Then
            If strProjectionCode > 0 Then 'check whether anything is selected
                Me.ProcedureDataGridView.DataSource = dt3.Select("ProjectionCode='" & strProjectionCode & "' and AgeGroup=1")
                'Me.VoltageTextBox.Text
                ' adding info about FIDs of selected procedure to ComboBox
                Me.FIDComboBox.Items.Add(Me.ProcedureDataGridView.Rows(0).Cells("FID").Value)
                'adding field sizes in ComboBox
                Me.FieldSizeComboBox.Items.Add(Me.ProcedureDataGridView.Rows(0).Cells("FieldWidthDataGridViewTextBoxColumn").Value.ToString & "x" & Me.ProcedureDataGridView.Rows(0).Cells("FieldHeightDataGridViewTextBoxColumn").Value)
                'first loop over source FIDs
                For i = 1 To Me.ProcedureDataGridView.RowCount - 1
                    strTempsrc = Math.Round(Me.ProcedureDataGridView.Rows(i).Cells("FID").Value).ToString
                    flag = False
                    ' second loop over the values that are already in ComboBox
                    For j = 0 To Me.FIDComboBox.Items.Count - 1
                        strTempDest = Me.FIDComboBox.Items.Item(j).ToString
                        If (strTempDest = strTempsrc) Then
                            flag = False
                            Exit For
                        Else
                            flag = True
                        End If
                    Next j
                    If flag Then
                        Me.FIDComboBox.Items.Add(Me.ProcedureDataGridView.Rows(i).Cells("FID").Value)
                    End If
                    strTempsrc2 = Me.ProcedureDataGridView.Rows(i).Cells("FieldWidthDataGridViewTextBoxColumn").Value.ToString & "x" & Me.ProcedureDataGridView.Rows(i).Cells("FieldHeightDataGridViewTextBoxColumn").Value.ToString
                    MinVoltage = Me.ProcedureDataGridView.Rows(0).Cells("PowerDataGridViewTextBoxColumn").Value
                    MaxVoltage = Me.ProcedureDataGridView.Rows(0).Cells("PowerDataGridViewTextBoxColumn").Value
                    For j = 0 To Me.FieldSizeComboBox.Items.Count - 1
                        If Me.FieldSizeComboBox.Items.Item(j) = strTempsrc2 Then
                            flagFieldSize = False
                            Exit For
                        Else
                            flagFieldSize = True
                        End If
                    Next
                    If flagFieldSize Then
                        Me.FieldSizeComboBox.Items.Add(strTempsrc2)
                    End If
                    'searching min and voltages
                    If MinVoltage > Me.ProcedureDataGridView.Rows(i).Cells("PowerDataGridViewTextBoxColumn").Value Then
                        MinVoltage = Me.ProcedureDataGridView.Rows(i).Cells("PowerDataGridViewTextBoxColumn").Value
                    End If
                    If MaxVoltage < Me.ProcedureDataGridView.Rows(i).Cells("PowerDataGridViewTextBoxColumn").Value Then
                        MaxVoltage = Me.ProcedureDataGridView.Rows(i).Cells("PowerDataGridViewTextBoxColumn").Value
                    End If
                Next i
                'END loop over records for voltage
                VoltageList(0) = MinVoltage ' filling in the array of voltages needed for interpolation
                VoltageList(1) = MaxVoltage
                Me.FieldSizeComboBox.Sorted = True
                If Me.FieldSizeComboBox.Items.Count = 1 Then
                    Me.FieldSizeComboBox.SelectedText = Me.FieldSizeComboBox.Items(0)
                End If
                Me.FIDComboBox.Sorted = True
                If Me.FIDComboBox.Items.Count = 1 Then
                    Me.FIDComboBox.SelectedText = Me.FIDComboBox.Items(0)
                Else
                    Me.FIDComboBox.Text = "Выберите РИП"
                End If
                'erase FIDComboBox after changing the procedure and/Or after calculation
                Me.VoltageProcLabel.Text = Global.RentgenCalculator.My.Resources.Resources.VoltageLabel + " (" + MinVoltage.ToString + " - " + MaxVoltage.ToString + " кВ)"
            Else
                Me.ProcedureDataGridView.DataSource = Nothing
                Me.VoltageProcLabel.Text = Global.RentgenCalculator.My.Resources.Resources.VoltageLabel + "," + Global.RentgenCalculator.My.Resources.Resources.kiloVoltsshorttext
            End If
        Else
            Me.ProcedureDataGridView.DataSource = Nothing
            Me.VoltageProcLabel.Text = Global.RentgenCalculator.My.Resources.Resources.VoltageLabel + "," + Global.RentgenCalculator.My.Resources.Resources.kiloVoltsshorttext
        End If
        'Me.ActiveControl = ProcedureDataGridView
        'Me.ProcedureDataGridView.Focus()
        Me.ProcedureDataGridView.Rows(0).Cells(0).Selected = True 'ArgumentOutOfRangeException; when age is not selected or the database doesn't have the record' эту строку вообще нужно будет убрать
        Me.ProcedureDataGridView.CurrentCell = Me.ProcedureDataGridView.Rows(0).Cells(0)
    End Sub
    Private Sub ProcTree_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ProcTree.Click
        Me.TabControl1.SelectTab(1)
    End Sub
    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        PrintForm.Show()
    End Sub
    Private Sub DataGridView2_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles ProcedureDataGridView.CellValidating
        Dim newDouble As Double
        Me.ProcedureDataGridView.Rows(e.RowIndex).ErrorText = ""

        If ProcedureDataGridView.Rows(e.RowIndex).IsNewRow Then Return ' чтобы убрать эту строку, нужно отлаживать процесс заполнения таблицы: если она выполняется при 
        If Not Double.TryParse(e.FormattedValue.ToString(), newDouble) Then

            e.Cancel = True
            Me.ProcedureDataGridView.Rows(e.RowIndex).ErrorText = "Неверный формат. Значение может быть только числом. (503)" ' при вводе в mAs 0,1

        End If

        Select Case (e.ColumnIndex)
            Case Me.ProcedureDataGridView.Columns("PowerDataGridViewTextBoxColumn").Index
                If Double.TryParse(e.FormattedValue.ToString(), newDouble) Then
                    If newDouble < 20 OrElse newDouble > 150 Then

                        e.Cancel = True
                        Me.ProcedureDataGridView.Rows(e.RowIndex).ErrorText = "Неверное значение. Величина напряжения не может быть меньше 20 кВ или больше 150 кВ."
                    End If
                End If
            Case Me.ProcedureDataGridView.Columns("FilterDataGridViewTextBoxColumn").Index
                If newDouble < 0 OrElse newDouble > 6 Then

                    e.Cancel = True
                    Me.ProcedureDataGridView.Rows(e.RowIndex).ErrorText = "Неверное значение. Толщина фильтра не может быть меньше 0 мм или больше 6 мм Al."
                End If

            Case Me.ProcedureDataGridView.Columns("FID").Index ' когда переключаешь назад во вкладку Пациент из выбранной процедуры
                If Double.TryParse(e.FormattedValue.ToString(), newDouble) Then
                    If newDouble < 10 OrElse newDouble > 200 Then

                        e.Cancel = True
                        Me.ProcedureDataGridView.Rows(e.RowIndex).ErrorText = "Неверное значение. Величина РИП не может быть меньше 10 см или больше 200 см."
                    End If
                End If
        End Select

    End Sub
    Private Sub DataGridView1_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles DataGridView1.CellValidating
        Me.DataGridView1.Rows(e.RowIndex).ErrorText = ""
        Dim newDouble As Double


        If DataGridView1.Rows(e.RowIndex).IsNewRow Then Return
        If DataGridView1.Columns(e.ColumnIndex).Name = "DeviceNameDataGridViewTextBoxColumn" Then Return
        If DataGridView1.Columns(e.ColumnIndex).Name = "CommentsDataGridViewTextBoxColumn" Then Return

        If Not Double.TryParse(e.FormattedValue.ToString(), newDouble) Then

            e.Cancel = True
            Me.DataGridView1.Rows(e.RowIndex).ErrorText = "Неверный формат. В ячейку можно ввести только числовое значение. (555)"

        End If

        Select Case (e.ColumnIndex)

            Case Me.DataGridView1.Columns("PowerDataGridViewTextBoxColumn1").Index
                If Double.TryParse(e.FormattedValue.ToString(), newDouble) Then
                    If newDouble < 20 OrElse newDouble > 150 Then
                        e.Cancel = True
                        Me.DataGridView1.Rows(e.RowIndex).ErrorText = _
                        "Неверное значение. Величина напряжения не может быть меньше 20 кВ или больше 150 кВ."
                    End If
                End If
            Case Me.DataGridView1.Columns("YieldDataGridViewTextBoxColumn").Index 'error.Tip:use new keyword
                If Not (e.FormattedValue = Nothing) Then
                    'Try
                    If Double.TryParse(e.FormattedValue.ToString(), newDouble) Then
                        If (e.FormattedValue > 1000000.0) Or (e.FormattedValue <= 0) Then
                            e.Cancel = True
                            Me.DataGridView1.Rows(e.RowIndex).ErrorText = "Величина радиационного выхода РДА должна лежать в пределах от 0 до 10 мР."
                        End If
                    Else
                        e.Cancel = True
                        Me.DataGridView1.Rows(e.RowIndex).ErrorText = "Недействительное значение рацидационного выхода"
                        '                        Exit Sub
                    End If 'проверка, введено ли правильное значение
                    'catch InvalidCastException
                Else
                    e.Cancel = True
                    Me.DataGridView1.Rows(e.RowIndex).ErrorText = "Введите значение рацидационного выхода"
                End If
            Case Me.DataGridView1.Columns("FilterDataGridViewTextBoxColumn1").Index
                If newDouble < 1 OrElse newDouble > 6 Then
                    e.Cancel = True
                    Me.DataGridView1.Rows(e.RowIndex).ErrorText = _
                    "Неверное значение. Толщина фильтра не может быть меньше 1 мм или больше 6 мм Al."
                End If
                'Case Me.DataGridView1.Columns("Rip").Index
                'If Double.TryParse(e.FormattedValue.ToString(), newDouble) Then
                '                If newDouble < 10 OrElse newDouble > 200 Then
                '                e.Cancel = True
                'Me.DataGridView1.Rows(e.RowIndex).ErrorText = "Неверное значение. Величина РИП не может быть меньше 10 см или больше 200 см."
                'End If
                'End If
        End Select
    End Sub
    Private Sub DataGridView2_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles ProcedureDataGridView.DataError
        Dim strStringContexts As String = ""
        If (e.Context = DataGridViewDataErrorContexts.Commit) _
            Then
            strStringContexts = "при изменении значения ячейки."
        End If
        If (e.Context = DataGridViewDataErrorContexts _
            .CurrentCellChange) Then
            strStringContexts = "при изменении значения ячейки."
        End If
        If (e.Context = DataGridViewDataErrorContexts.Parsing) _
            Then
            strStringContexts = "при изменении значения ячейки."
        End If
        If (e.Context = _
            DataGridViewDataErrorContexts.LeaveControl) Then
            strStringContexts = "при изменении значения ячейки."
        End If
        If (e.Context = _
                   DataGridViewDataErrorContexts.Formatting) Then
            strStringContexts = "при изменении значения ячейки."
        End If
        MsgBox(My.Resources.MsgDefault & strStringContexts, MsgBoxStyle.Exclamation, My.Resources.MainTitle)

        If (TypeOf (e.Exception) Is ConstraintException) Then
            Dim view As DataGridView = CType(sender, DataGridView)
            view.Rows(e.RowIndex).ErrorText = "Ошибка"
            view.Rows(e.RowIndex).Cells(e.ColumnIndex) _
                .ErrorText = "Ошибка"

            e.ThrowException = False
        End If

    End Sub
    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        Dim strStringContexts As String = ""
        If ((e.Context = DataGridViewDataErrorContexts.Commit) Or _
 (e.Context = DataGridViewDataErrorContexts.CurrentCellChange) Or _
 (e.Context = DataGridViewDataErrorContexts.Parsing) Or _
 (e.Context = DataGridViewDataErrorContexts.LeaveControl) Or _
 (e.Context = DataGridViewDataErrorContexts.Formatting)) Then
            strStringContexts = "при изменении значения ячейки."
        End If
        MsgBox("Произошла ошибка " & strStringContexts, MsgBoxStyle.Exclamation, My.Resources.MainTitle) ' при вводе дробного зачения напряжения

        If (TypeOf (e.Exception) Is ConstraintException) Then
            Dim view As DataGridView = CType(sender, DataGridView)
            view.Rows(e.RowIndex).ErrorText = "Ошибка"
            view.Rows(e.RowIndex).Cells(e.ColumnIndex) _
                .ErrorText = "Ошибка"

            e.ThrowException = False
        End If

    End Sub
    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        'теоретически, эта процедура может сильно подтормаживать программу, если много раз щелкать по таблице
        ' а еще ее нужно переработать, чтобы не надо было все время выбирать прибор, если он всего один
        If Me.DataGridView1.Focused Then '? the opposite situation and this event occurs when the form loads
            If Me.DataGridView1.Rows.Count > 1 Then '?check whether there are any devices in the list of currently used devices
                If Not Me.DataGridView1.CurrentRow.Cells("DeviceNameDataGridViewTextBoxColumn").Value Is Nothing Then
                    If Not Me.DataGridView1.CurrentRow.Cells("DeviceNameDataGridViewTextBoxColumn").Value.ToString.Length = 0 Then

                        strSelectedDeviceName = Me.DataGridView1.CurrentRow.Cells("DeviceNameDataGridViewTextBoxColumn").Value
                    Else
                        strSelectedDeviceName = ""
                    End If
                End If
                If Not Me.DataGridView1.CurrentRow.Cells("PowerDataGridViewTextBoxColumn1").Value Is Nothing Then
                    If Not Me.DataGridView1.CurrentRow.Cells("PowerDataGridViewTextBoxColumn1").Value.ToString.Length = 0 Then
                        dblPower = Me.DataGridView1.CurrentRow.Cells("PowerDataGridViewTextBoxColumn1").Value
                    Else
                        dblPower = Nothing
                    End If
                End If
                If Not Me.DataGridView1.CurrentRow.Cells("YieldDataGridViewTextBoxColumn").Value Is Nothing Then
                    If Not Me.DataGridView1.CurrentRow.Cells("YieldDataGridViewTextBoxColumn").Value.ToString.Length = 0 Then
                        dblYield = Me.DataGridView1.CurrentRow.Cells("YieldDataGridViewTextBoxColumn").Value
                    Else
                        dblYield = Nothing
                    End If
                End If
                If Not Me.DataGridView1.CurrentRow.Cells("Rip").Value Is Nothing Then
                    If Not Me.DataGridView1.CurrentRow.Cells("Rip").Value.ToString.Length = 0 Then
                        dblRip = Me.DataGridView1.CurrentRow.Cells("Rip").Value
                    Else
                        dblRip = Nothing
                    End If
                End If
                If Not Me.DataGridView1.CurrentRow.Cells("FilterDataGridViewTextBoxColumn1").Value Is Nothing Then
                    If Not Me.DataGridView1.CurrentRow.Cells("FilterDataGridViewTextBoxColumn1").Value.ToString.Length = 0 Then
                        dblFilter = Me.DataGridView1.CurrentRow.Cells("FilterDataGridViewTextBoxColumn1").Value
                    End If
                Else
                    dblFilter = Nothing ' это будет означать, что ни один прибор не выбран
                End If
            End If
        End If
    End Sub
    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub
    Private Sub ContentsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ContentsToolStripMenuItem.Click
        Dim helpFileName As String = Environment.CurrentDirectory & "\RentgenCalculator.chm"
        Help.ShowHelp(Me, helpFileName)
    End Sub
    Private Sub HelpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripMenuItem.Click
        Dim helpFileName As String = Environment.CurrentDirectory & "\XRayCalc.chm"
        Help.ShowHelp(Me, helpFileName) 'could be parent or navigator
    End Sub
    Private Sub НастройкиToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SettingsToolStripMenuItem.Click
        NewApparatusForm.ShowDialog()
        'below the file RentgenCalculator.ini should be considered
        'If
    End Sub
    Private Sub RippleTextBox_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RippleTextBox.Leave
        Dim ToolTip1 As New ToolTip
        If RippleTextBox.Text.Length = 0 Then
            ToolTip1.Show(My.Resources.MsgEmpty + My.Resources.Ripple, RippleTextBox, 2000)
        Else
            If Double.TryParse(RippleTextBox.Text, dblRipple) Then
                If Convert.ToDouble(RippleTextBox.Text) < 0 Or Convert.ToDouble(RippleTextBox.Text) > 7 Then

                    'MsgBox("Введено неверное значение пульсации напряжения. Значение может быть в пределах от 0 до 7", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                    RippleTextBox.ForeColor = Color.Red
                Else
                    dblRipple = RippleTextBox.Text
                    RippleTextBox.ForeColor = Color.Black
                End If
            Else
                dblRipple = -1
                ToolTip1.Show(My.Resources.MsgDefault + "при вводе пульсации напряжения.", RippleTextBox, 2000)
                'MsgBox(My.Resources.MsgDefault + " при вводе пульсации напряжения.", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                RippleTextBox.ForeColor = Color.Red
            End If 'try parse
        End If
    End Sub
    Private Sub TextBox1_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegistryNumTextBox.Leave
        If (RegistryNumTextBox.Text = Nothing) Or (RegistryNumTextBox.Text = "") Then
            Dim ToolTip2 As New ToolTip
            ToolTip2.Show("Пожалуйста, введите регистрационный номер", RegistryNumTextBox, 2000)
        End If
    End Sub

    Private Sub VoltageTextBox_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles VoltageProcTextBox.Validating
        If VoltageProcTextBox.Text = "" Then
            ToolTip1.Show("Пожалуйста, введите напряжение", VoltageProcTextBox, 2000)
        Else
            If Double.TryParse(VoltageProcTextBox.Text, dblPower) Then
                If dblPower < MinVoltage Then
                    ToolTip1.Show("Значение напряжения должно быть выше " & MinVoltage, VoltageProcTextBox, 2000)
                    VoltageProcTextBox.ForeColor = Color.Red
                ElseIf dblPower > MaxVoltage Then
                    ToolTip1.Show(My.Resources.WarningHighVoltage & " " & MaxVoltage, VoltageProcTextBox, 2000)
                    VoltageProcTextBox.ForeColor = Color.Red
                Else
                    VoltageProcTextBox.ForeColor = Color.Black
                    VoltageProcTextBox.BackColor = Color.White
                End If
            Else
                ToolTip1.Show("Неверный формат числа. Разделителем целой и дробной части должна быть запятая", VoltageProcTextBox, 2000)
                VoltageProcTextBox.ForeColor = Color.Red
            End If
        End If
    End Sub

    Private Sub FilterTextBox_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FilterTextBox.Leave
        'Dim ToolTip5 As New ToolTip
        If Not (FilterTextBox.Text Is Nothing) Then
            If FilterTextBox.Text.Length = 0 Then
                ToolTip1.Show("Пожалуйста, введите толщину фильтра рентгеновской трубки.", FilterTextBox, 2000)
                TabControl1.SelectTab(1)
            Else
                Try
                    ToolTip1.Hide(Nothing)
                Catch ANE As ArgumentNullException
                    ToolTip1.Hide(Me)
                End Try
                If Not Double.TryParse(FilterTextBox.Text, dblFilter2) Then
                    'Dim ToolTip1 As New ToolTip
                    ToolTip1.Show("Неверный формат данных. Введите число.", FilterTextBox, 2000)
                    FilterTextBox.ForeColor = Color.Red
                Else
                    If (FilterTextBox.Text < 0.1) Or (FilterTextBox.Text > 6) Then
                        ToolTip1.Show("Значение фильтра должно лежать в пределах от 0.1 до 6 мм.", FilterTextBox, 2000)
                        FilterTextBox.ForeColor = Color.Red
                    Else
                        FilterTextBox.ForeColor = Color.Black
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub mAsTextBox_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mAsTextBox.Leave
        Dim mAsTextBoxTooltip As New ToolTip
        mAsTextBoxTooltip.ToolTipTitle = "мАс"
        If mAsTextBox.Text = "" Then
            mAsTextBoxTooltip.Show("Пожалуйста, введите мАс", mAsTextBox, 2000)
        ElseIf Not Double.TryParse(mAsTextBox.Text, dblmAs) Then
            mAsTextBoxTooltip.Show("Неверный формат данных. Введите число.", mAsTextBox, 2000)
            'MsgBox("Неверный формат данных. Введите число.", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
            mAsTextBox.ForeColor = Color.Red
        Else
            mAsTextBox.ForeColor = Color.Black
            If dblmAs < 0.1 OrElse dblmAs > 1000 Then
                mAsTextBoxTooltip.Show("Введено неверное значение мАс. Величина мАс не может быть меньше 0,1 или больше 1000.", mAsTextBox, 2000)
                '    MsgBox("Введено неверное значение мАс. Величина мАс не может быть меньше 0,1 или больше 1000.", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                Exit Sub
            Else
                dblmAs = Convert.ToDouble(mAsTextBox.Text, CultureInfo.CurrentUICulture) 'подпрограмма сделана 18.2.2019
            End If
        End If
    End Sub
End Class
