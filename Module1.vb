Imports System
Imports System.IO
Imports System.Globalization
Module Module1

    Public strRegnum, strSurname, strName, strPatronymic, strDeviceName, strProcedureName, strProjectionCode, strSelectedDeviceName As String 'strCriteria, 
    Public strFileName As String = "RentgenCalculator.ini"
    Public intAge, intSex, intDevType, i As Integer
    ' For Compatibility with possible further addition of pediatric phantoms
    Public intAgeGroup As Integer
    'Rate from list of active devices (MyDevices db)
    Public dblYield As Double
    Public dblmAs, dblRateDev, dblFilter, dblPower, dblRip, DoseCoeff As Double
    ' ripple
    Public dblRipple As Double = -1
    'вес и рост пациента
    Public sngWeight, sngHeight As Single
    Public drInsert As DataRow
    Public ds As New DataDataSet
    Public da As New DataDataSetTableAdapters.DevicesListTableAdapter
    Public da1 As New DataDataSetTableAdapters.DataForCalculationTableAdapter
    Public da2 As New DataDataSetTableAdapters.MyDevicesListTableAdapter
    Public da3 As New DataDataSetTableAdapters.ProceduresListTableAdapter
    Public da4 As New PatientDataSetTableAdapters.XCalc_graphTableAdapter 'adapter for Patient table
    'Public da5 As New doses.
    ' adding da6 as the container for DRL values' they should be modified in the future 7 Dec 2020
    Public dataAdapterDRL As New DataDataSetTableAdapters.DRLTableAdapter
    Public dt As New DataDataSet.DevicesListDataTable
    Public dt1 As New DataDataSet.DataForCalculationDataTable
    Public dt2 As New DataDataSet.MyDevicesListDataTable
    Public dt3 As New DataDataSet.ProceduresListDataTable
    Public dt4 As New PatientDataSet.XCalc_graphDataTable
    Public dt5 As New doses.dosesDataTable
    'Public dt5Sorted As New doses.dosesDataTable
    Public dataTableDRL As New DataDataSet.DRLDataTable
    
    Public dr() As DataDataSet.DevicesListRow
    Public dr1() As DataDataSet.DataForCalculationRow
    Public dr2() As DataDataSet.MyDevicesListRow
    Public dr3() As DataDataSet.ProceduresListRow
    Public dr4() As PatientDataSet.XCalc_graphRow 'variable representing row of the patient DB
    Public dr5() As doses.dosesRow ' united table for doses with input data??? should be DataRow
    ' a variable for operating with rows of DRL table (added on 7 Dec 2020)
    Public datarowDRL() As DataDataSet.DRLRow
    
    Public strHospCode As String
    Public strFilterCh, strRipCh As String 'current values in DataGrid2()
    Public dblPowerCh, strWidthCh, dbHeight As Double 'current value in DataGrid2() - ципа напряжение при проведении процедуры
    Public dbEffDose As Double
    Public lngKeyFields() As Long
    Public j As Integer
    Public k As Byte
    Public NumFilter As Byte
    Public strTextFile As String
    Public drSource() As DataRow
    Public drInterpolation2() As DataRow

    Public myStream As System.IO.Stream
    Public PowCoeff As Double = 2.2 'power coefficient for recalculating dose depending on tube high voltage
    Public VoltageList() As Single = {0, 0}
    Public rippleList() As Single = {0, 5}
    'weighting coefficients for tissues and organs in respective order
    Public MainDoseCoeffM() As Double = {0.08, 0.12, 0.12, 0.12, 0.12, 0.04, 0.04, 0.04, 0.04, 0.01, 0.01, 0.12, 0.01, 0.01}
    Public MainDoseCoeffF() As Double = {0.08, 0.12, 0.12, 0.12, 0.12, 0.04, 0.04, 0.04, 0.04, 0.01, 0.01, 0.12, 0.01, 0.01}
    Public MainDosesM(13) As Double
    Public MainDosesF(13) As Double
    Public AdditionalDosesM(12) As Double
    Public AdditionalDosesF(12) As Double
    Public MainDosesMInterp(13, 2, 1) As Single '14 органов, 2 напряжения+1усредненное
    Public MainDosesFInterp(13, 2, 1) As Single '14 органов, 2 напряжения+1усредненное
    Public MainDosesMInterpFinal() As Single = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0} 'for safety: some organs can have zero doses
    Public MainDosesFInterpFinal() As Single = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0} 'for safety: some organs can have zero doses
    Public AdditionalDosesMInterp(12, 2, 1) As Double
    Public AdditionalDosesFInterp(12, 2, 1) As Double
    Public AdditionalDosesMInterpFinal(12) As Single
    Public AdditionalDosesFInterpFinal(12) As Single
    Public FieldorderMainM() As Byte = {0, 1, 2, 3, 4, 5, 11, 6, 7, 8, 9, 10, 255, 12, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 13}
    'The following is the number of MainDoses which corresponds to current loop Field element
    Public FieldorderMainF() As Byte = {0, 1, 2, 3, 4, 5, 11, 6, 7, 8, 9, 10, 255, 12, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 13}
    Public FieldorderAdditionalM() As Byte = {255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 127, 255, 8, 2, 3, 7, 4, 5, 6, 255, 9, 1, 10, 11, 12}
    Public FieldorderAdditionalF() As Byte = {255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 127, 255, 8, 2, 3, 7, 4, 5, 6, 12, 9, 1, 10, 11, 255}
    Public strText() As Byte
    Public strLess As String = "менее 0,001" 'this is written instead of absorbed dose if it's value is lower than 0.001 mSv
    Public strPatientsdbname = "Patients.mdb"
    'test values
    Public MaxHeight As Single = 250
    Public MinHeight As Single = 50
    'msg
    Public strHeightMsg As String = My.Resources.MsgWrongHeight & MinHeight & My.Resources.andWord & MaxHeight


    Public Function transformRate(ByVal y0 As Double, ByVal r0 As Double, ByVal r1 As Double, ByVal f0 As Double, ByVal f1 As Double, ByVal u0 As Double, ByVal u1 As Double) As Double
        transformRate = Math.Pow(r0 / r1, 2) * Math.Pow(u1 / u0, PowCoeff) * Math.Pow(f0 / f1, 0.755) * y0
    End Function
    Public Function checkVal(ByVal dose As Double) 'the purpose of this function is to make DataOutput more laconic
        If Math.Round(dose, 3) < 0.001 Then
            checkVal = strLess
        Else
            checkVal = Convert.ToString(Math.Round(dose, 3), CultureInfo.CurrentUICulture)
        End If
    End Function
    Public Sub AbsorbedDosesInterp() 'and also effective doses
        Dim iOrg, i, j, iVolt, iRipple As Byte 'loop index
        Dim drInsert2 As DataRow
        'Array for interpolation of doses along filter using Newton method
        Dim delta(,) As Single = {{0, 0, 0, 0, 0}, {0, 0, 0, 0, 0}, {0, 0, 0, 0, 0}, {0, 0, 0, 0, 0}, {0, 0, 0, 0, 0}}
        Dim Minfilteri() As Byte = {0, 1, 2, 3, 4, 5} ' the indicies in order of acsending filter
        Dim tmpMinfilter As Byte
        Dim Sorted As Boolean = False
        Dim statechanged As Boolean = False 'пришлось ввести эту переменную, потому что иначе стирает некоторые значения при дальнейшем проходе списка после изменения мест переменных
        Dim strCriteria2 As String ' request line
        Dim currentsum As Single ' temporary variable for interpolation
        Dim tmpDose As Single ' temporary value of dose
        'MALE CASE
        strCriteria2 = "" ' creating empty request line
        'цикл заполнения doses, имеющих одинаковые индексы ProcedureCode, но разные фильтры
        For iRipple = 0 To 1
            For iVolt = 0 To 1
                drSource = dt3.Select("ProjectionCode='" & strProjectionCode & "' and AgeGroup=1 and ripple=" & rippleList(iRipple) & " and Power=" & VoltageList(iVolt) & " and FID=" & strRipCh & " and FieldHeight=" & dbHeight & " and FieldWidth=" & strWidthCh) 'выборка из таблицы ProceduresList, содержащая ProcedureCode для массива доз
                k = drSource.Length 'число строк, возвращаемых при выполнении предыдущего запроса = 6 для одной комбинации пол/напряжение/рипл
                For i = 0 To k - 1 'цикл для создания строки запроса. только по фильтру.
                    'из всех отобранных записей таблицы ProceduresList извлекаем связанные с ними записи из БД DataForCalculation
                    strCriteria2 = "Gender=1 AND ProcedureCode=" & drSource(i).Item("ProcedureCode")
                    drInsert2 = dt5.NewRow()
                    drInterpolation2 = dt1.Select(strCriteria2)
                    drInsert2("ProcedureCode") = drSource(i).Item("ProcedureCode")
                    drInsert2("AgeGroup") = 1
                    drInsert2("FID") = drSource(i).Item("FID")
                    drInsert2("Filter") = drSource(i).Item("Filter")
                    drInsert2("Gender") = 1 ' male
                    drInsert2("ripple") = drSource(i).Item("ripple")
                    drInsert2("Power") = drSource(i).Item("Power")
                    For iOrg = 2 To 29
                        drInsert2("Field" & iOrg) = drInterpolation2(0).Item("Field" & iOrg)
                    Next
                    dt5.Rows.Add(drInsert2)
                    'что быстрее: 1) читать по строкам, 2) читать сразу всю таблицу, а потом выдергивать из нее по-отдельности
                Next
                'сортировка по аргументу (фильтру)
                Do While Not Sorted
                    For i = 1 To k - 1
                        If Not statechanged Then
                            If (dt5.Rows(Minfilteri(i)).Item("filter") < dt5.Rows(Minfilteri(i - 1)).Item("filter")) Then
                                tmpMinfilter = Minfilteri(i - 1)
                                Minfilteri(i - 1) = Minfilteri(i)
                                Minfilteri(i) = tmpMinfilter
                                statechanged = True
                            Else
                                If i = k - 1 Then
                                    Sorted = True
                                End If
                            End If 'statechanged
                        End If
                    Next 'i
                    statechanged = False
                Loop 'конец сортировки по фильтру
                For iOrg = 2 To 29
                    'Интерполяция по формулам Ньютона. Корн.
                    'Если все значения равны нулю (например, у органов, для которых дозы не насчитаны), то никакую интерполяцию проводить не нужно!
                    For i = 1 To k - 1 'расчет разделенных разностей
                        delta(0, i - 1) = (dt5.Rows(Minfilteri(i)).Item("Field" & iOrg.ToString) - dt5.Rows(Minfilteri(i - 1)).Item("Field" & iOrg.ToString)) / (dt5.Rows(Minfilteri(i)).Item("filter") - dt5.Rows(Minfilteri(i - 1)).Item("filter"))
                    Next
                    For j = 1 To 4
                        For i = j + 1 To k - 1
                            delta(j, i - 1 - j) = (delta(j - 1, i - j) - delta(j - 1, i - 1 - j)) / (dt5.Rows(Minfilteri(i)).Item("filter") - dt5.Rows(Minfilteri(i - 1 - j)).Item("filter"))
                        Next
                    Next 'j last divided difference
                    currentsum = delta(4, 0)
                    For i = 0 To k - 2
                        currentsum = (strFilterCh - dt5.Rows(Minfilteri(i)).Item("filter")) * currentsum
                    Next
                    For j = 1 To 3
                        tmpDose = delta(4 - j, 0)
                        For i = 0 To k - 2 - j
                            tmpDose = (strFilterCh - dt5.Rows(Minfilteri(i)).Item("filter")) * tmpDose
                        Next ' i
                        currentsum = currentsum + tmpDose
                    Next ' j calculation of (y) based on 
                    'setting variables to dose values
                    'two separate variables with indicies of organs: FieldorderMainM for main organs
                    'and FieldorderAdditionalM for additional organs
                    If Not (FieldorderMainM(iOrg - 2) = 255) Then 'filling Main organs first
                        MainDosesMInterp(FieldorderMainM(iOrg - 2), iVolt, iRipple) = dt5.Rows(Minfilteri(0)).Item("Field" & iOrg.ToString) + (strFilterCh - dt5.Rows(Minfilteri(0)).Item("filter")) * delta(0, 0) + currentsum
                    ElseIf FieldorderAdditionalM(iOrg - 2) = 127 Then
                        AdditionalDosesMInterp(0, iVolt, iRipple) = dt5.Rows(Minfilteri(0)).Item("Field" & iOrg.ToString) + (strFilterCh - dt5.Rows(Minfilteri(0)).Item("filter")) * delta(0, 0) + currentsum
                    ElseIf Not (FieldorderAdditionalM(iOrg - 2) = 255) Then
                        AdditionalDosesMInterp(FieldorderAdditionalM(iOrg - 2), iVolt, iRipple) = dt5.Rows(Minfilteri(0)).Item("Field" & iOrg.ToString) + (strFilterCh - dt5.Rows(Minfilteri(0)).Item("filter")) * delta(0, 0) + currentsum
                    End If
                Next iOrg
                dt5.Clear()
            Next iVolt
            'interpolation over voltage' only 2 values -> linear interpolation
            For i = 0 To MainDosesMInterp.Length / 6 - 1
                MainDosesMInterp(i, 2, iRipple) = (MainDosesMInterp(i, 1, iRipple) - MainDosesMInterp(i, 0, iRipple)) * (dblPowerCh - VoltageList(0)) / (VoltageList(1) - VoltageList(0)) + MainDosesMInterp(i, 0, iRipple)
            Next
            For i = 0 To AdditionalDosesMInterp.Length / 6 - 1
                AdditionalDosesMInterp(i, 2, iRipple) = (AdditionalDosesMInterp(i, 1, iRipple) - AdditionalDosesMInterp(i, 0, iRipple)) * (dblPowerCh - VoltageList(0)) / (VoltageList(1) - VoltageList(0)) + AdditionalDosesMInterp(i, 0, iRipple)
            Next
        Next iRipple
        'interpolation over ripple
        If Not (dblRipple = -1) Then
            For i = 0 To MainDosesMInterp.Length / 6 - 1 '3 values of doses for ripple 0 +3 values of doses for voltages for ripple 5
                MainDosesMInterpFinal(i) = ((MainDosesMInterp(i, 2, 1) - MainDosesMInterp(i, 2, 0)) * (dblRipple - rippleList(0)) / (rippleList(1) - rippleList(0)) + MainDosesMInterp(i, 2, 0)) * DoseCoeff 'с учетом значения радиационного выхода
            Next
        Else
            Throw New System.Exception()
            '2) Return RunButton_Click()
        End If
        For i = 0 To AdditionalDosesMInterp.Length / 6 - 1
            AdditionalDosesMInterpFinal(i) = ((AdditionalDosesMInterp(i, 2, 1) - AdditionalDosesMInterp(i, 2, 0)) * (dblRipple - rippleList(0)) / (rippleList(1) - rippleList(0)) + AdditionalDosesMInterp(i, 2, 0)) * DoseCoeff 'с учетом значения радиационного выхода
        Next
        'FEMALE BLOCK
        'цикл заполнения doses, имеющих одинаковые индексы ProcedureCode, но разные фильтры
        For iRipple = 0 To 1
            For iVolt = 0 To 1
                drSource = dt3.Select("ProjectionCode='" & strProjectionCode & "' and AgeGroup=1 and ripple=" & rippleList(iRipple) & " and Power=" & VoltageList(iVolt) & " and FID=" & strRipCh & " and FieldHeight=" & dbHeight & " and FieldWidth=" & strWidthCh) 'выборка из таблицы ProceduresList, содержащая ProcedureCode для массива доз
                ' на данный момент предыдущая строка повторяется ровно дважды. Если будет необходимость, можно будет включить перебор по полам в один цикл
                k = drSource.Length 'число строк, возвращаемых при выполнении предыдущего запроса = 6 для одной комбинации пол/напряжение/рипл
                For i = 0 To k - 1 'цикл для создания строки запроса. пока только по фильтру.
                    'в дальнейшем будут вложенные циклы по напряжению и пульсации
                    'из всех отобранных записей таблицы ProceduresList извлекаем связанные с ними записи из БД DataForCalculation
                    strCriteria2 = "Gender=2 AND ProcedureCode=" & drSource(i).Item("ProcedureCode")
                    drInsert2 = dt5.NewRow()
                    drInterpolation2 = dt1.Select(strCriteria2)
                    drInsert2("ProcedureCode") = drSource(i).Item("ProcedureCode")
                    drInsert2("AgeGroup") = 1
                    drInsert2("FID") = drSource(i).Item("FID")
                    drInsert2("Filter") = drSource(i).Item("Filter")
                    drInsert2("Gender") = 2 ' male
                    drInsert2("ripple") = drSource(i).Item("ripple")
                    drInsert2("Power") = drSource(i).Item("Power")
                    For iOrg = 2 To 29
                        drInsert2("Field" & iOrg) = drInterpolation2(0).Item("Field" & iOrg)
                    Next
                    dt5.Rows.Add(drInsert2)
                    'что быстрее: 1) читать по строкам, 2) читать сразу всю таблицу, а потом выдергивать из нее по-отдельности
                Next
                'сортировка по аргументу (фильтру)
                Do While Not Sorted
                    For i = 1 To k - 1
                        If Not statechanged Then
                            If (dt5.Rows(Minfilteri(i)).Item("filter") < dt5.Rows(Minfilteri(i - 1)).Item("filter")) Then
                                tmpMinfilter = Minfilteri(i - 1)
                                Minfilteri(i - 1) = Minfilteri(i)
                                Minfilteri(i) = tmpMinfilter
                                statechanged = True
                            Else
                                If i = k - 1 Then
                                    Sorted = True
                                End If
                            End If 'statechanged
                        End If
                    Next 'i
                    statechanged = False
                Loop 'конец сортировки по фильтру
                For iOrg = 2 To 29
                    'Интерполяция по формулам Ньютона. Корн.
                    'Если все значения равны нулю, то никакую интерполяцию проводить не нужно
                    For i = 1 To k - 1 'расчет разделенных разностей
                        delta(0, i - 1) = (dt5.Rows(Minfilteri(i)).Item("Field" & iOrg.ToString) - dt5.Rows(Minfilteri(i - 1)).Item("Field" & iOrg.ToString)) / (dt5.Rows(Minfilteri(i)).Item("filter") - dt5.Rows(Minfilteri(i - 1)).Item("filter"))
                    Next
                    For j = 1 To 4
                        For i = j + 1 To k - 1
                            delta(j, i - 1 - j) = (delta(j - 1, i - j) - delta(j - 1, i - 1 - j)) / (dt5.Rows(Minfilteri(i)).Item("filter") - dt5.Rows(Minfilteri(i - 1 - j)).Item("filter"))
                        Next
                    Next
                    currentsum = delta(4, 0)
                    For i = 0 To k - 2
                        currentsum = (strFilterCh - dt5.Rows(Minfilteri(i)).Item("filter")) * currentsum
                    Next
                    For j = 1 To 3
                        tmpDose = delta(4 - j, 0)
                        For i = 0 To k - 2 - j
                            tmpDose = (strFilterCh - dt5.Rows(Minfilteri(i)).Item("filter")) * tmpDose
                        Next ' i
                        currentsum = currentsum + tmpDose
                    Next ' j
                    'setting variables to dose values
                    'if strFilterCh=0 then resulting doses are invalid
                    If Not (FieldorderMainF(iOrg - 2) = 255) Then 'filling Main organs first
                        MainDosesFInterp(FieldorderMainF(iOrg - 2), iVolt, iRipple) = dt5.Rows(Minfilteri(0)).Item("Field" & iOrg.ToString) + (strFilterCh - dt5.Rows(Minfilteri(0)).Item("filter")) * delta(0, 0) + currentsum
                    ElseIf FieldorderAdditionalF(iOrg - 2) = 127 Then
                        AdditionalDosesFInterp(0, iVolt, iRipple) = dt5.Rows(Minfilteri(0)).Item("Field" & iOrg.ToString) + (strFilterCh - dt5.Rows(Minfilteri(0)).Item("filter")) * delta(0, 0) + currentsum
                    ElseIf Not (FieldorderAdditionalF(iOrg - 2) = 255) Then
                        AdditionalDosesFInterp(FieldorderAdditionalF(iOrg - 2), iVolt, iRipple) = dt5.Rows(Minfilteri(0)).Item("Field" & iOrg.ToString) + (strFilterCh - dt5.Rows(Minfilteri(0)).Item("filter")) * delta(0, 0) + currentsum
                    End If
                Next iOrg
                dt5.Clear()
            Next iVolt
            'interpolation over voltage' only 2 values -> linear interpolation
            For i = 0 To MainDosesFInterp.Length / 6 - 1
                MainDosesFInterp(i, 2, iRipple) = (MainDosesFInterp(i, 1, iRipple) - MainDosesFInterp(i, 0, iRipple)) * (dblPowerCh - VoltageList(0)) / (VoltageList(1) - VoltageList(0)) + MainDosesFInterp(i, 0, iRipple)
            Next
            For i = 0 To AdditionalDosesFInterp.Length / 6 - 1
                AdditionalDosesFInterp(i, 2, iRipple) = (AdditionalDosesFInterp(i, 1, iRipple) - AdditionalDosesFInterp(i, 0, iRipple)) * (dblPowerCh - VoltageList(0)) / (VoltageList(1) - VoltageList(0)) + AdditionalDosesFInterp(i, 0, iRipple)
            Next
        Next iRipple
        'interpolation over ripple
        For i = 0 To MainDosesFInterp.Length / 6 - 1 '3 values of doses for ripple 0 +3 values of doses for voltages for ripple 5
            MainDosesFInterpFinal(i) = ((MainDosesFInterp(i, 2, 1) - MainDosesFInterp(i, 2, 0)) * (dblRipple - rippleList(0)) / (rippleList(1) - rippleList(0)) + MainDosesFInterp(i, 2, 0)) * DoseCoeff
        Next
        For i = 0 To AdditionalDosesFInterp.Length / 6 - 1
            AdditionalDosesFInterpFinal(i) = ((AdditionalDosesFInterp(i, 2, 1) - AdditionalDosesFInterp(i, 2, 0)) * (dblRipple - rippleList(0)) / (rippleList(1) - rippleList(0)) + AdditionalDosesFInterp(i, 2, 0)) * DoseCoeff
        Next
        dbEffDose = 0
        For i = 0 To MainDosesMInterpFinal.Length - 1
            dbEffDose = dbEffDose + MainDosesMInterpFinal(i) * MainDoseCoeffM(i)
        Next
        For i = 0 To AdditionalDosesMInterpFinal.Length - 1
            dbEffDose = dbEffDose + 0.12 * AdditionalDosesMInterpFinal(i) / (AdditionalDosesMInterpFinal.Length)
        Next
        For i = 0 To MainDosesFInterpFinal.Length - 1
            dbEffDose = dbEffDose + MainDosesFInterpFinal(i) * MainDoseCoeffF(i)
        Next
        For i = 0 To AdditionalDosesFInterpFinal.Length - 1
            dbEffDose = dbEffDose + 0.12 * AdditionalDosesFInterpFinal(i) / (AdditionalDosesFInterpFinal.Length)
        Next
        dbEffDose = dbEffDose / 2
    End Sub
    Public Function ExtrctProjection(ByVal Proj As String)
        ' used in CalculationForm
        Select Case Proj
            Case "Прямая передняя проекция"
                Return "ПЗ"
            Case "Прямая задняя проекция"
                Return "ЗП"
            Case "Боковая левая проекция"
                Return "БЛ"
            Case "Боковая правая проекция"
                Return "БП"
            Case "Косая задняя проекция (справа)"
                Return "КЗП"
            Case "Косая задняя проекция (слева)"
                Return "КЗЛ"
            Case "Косая передняя проекция (справа)"
                Return "КПП"
            Case "Косая задняя проекция (слева)"
                Return "КПЛ"
            Case Else
                Return ""
        End Select
    End Function
    Public Sub ClearInput()
        RentgenCalculator.MainForm.VoltageProcLabel.Text = RentgenCalculator.My.Resources.Resources.VoltageLabel + "," + RentgenCalculator.My.Resources.Resources.kiloVoltsshorttext
        'erase FIDComboBox after changing the procedure and/Or after calculation
        RentgenCalculator.MainForm.FIDComboBox.Items.Clear()
        If RentgenCalculator.MainForm.FIDComboBox.Text = "" Then
            RentgenCalculator.MainForm.FIDComboBox.SelectedText = ""
        Else
            RentgenCalculator.MainForm.FIDComboBox.Text = ""
        End If
        RentgenCalculator.MainForm.FieldSizeComboBox.Items.Clear()
        RentgenCalculator.MainForm.FieldSizeComboBox.SelectedText = ""
        RentgenCalculator.MainForm.FieldSizeComboBox.Text = ""
    End Sub
    Public Function DRLCheck()
        Try
            dataAdapterDRL.Fill(dataTableDRL)
            'Dim DRLvalue=
            Return "Да"
        Catch error1 As Exception
            'if the database with DRL couldn't be read.
            MsgBox("Couldn't fill in the table")
            Return "Не известно."
        End Try
    End Function
    'Public Function Read
    'Return
    'End Function
    Public Sub CheckHeight()
        'начало проверки роста
        'Function is not doing anything if sngHeight is filled
        If sngHeight = Nothing Then
            If MainForm.HeightTextBox.Text = "" Then
                MsgBox("Не задан рост обследуемого лица!", MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                MainForm.TabControl1.SelectTab(0)
                MainForm.HeightTextBox.Focus()
                Exit Sub
            Else 'не нулевая строка
                Try
                    If Convert.ToDecimal(MainForm.HeightTextBox.Text, CultureInfo.CurrentUICulture) > MaxHeight Or Convert.ToDecimal(MainForm.HeightTextBox.Text, CultureInfo.CurrentUICulture) < MinHeight Then
                        MsgBox(strHeightMsg, MsgBoxStyle.Exclamation, My.Resources.MainTitle)
                        MainForm.TabControl1.SelectTab(0)
                        MainForm.HeightTextBox.ForeColor = Color.Red
                        Exit Sub
                    Else
                        MainForm.HeightTextBox.ForeColor = Color.Black
                        sngHeight = MainForm.HeightTextBox.Text
                    End If 'проверка на значение роста в пределах нормальных значений
                Catch SFE As System.FormatException
                    MsgBox("Неверно введен рост пациента.", MsgBoxStyle.Exclamation, My.Resources.MainTitle) 'SFE.Message 
                    MainForm.HeightTextBox.ForeColor = Color.Red
                    MainForm.TabControl1.SelectTab(0)
                    MainForm.HeightTextBox.Focus()
                    Exit Sub
                End Try 'попытка присвоить переменной Рост значение из HeightTextBox
            End If 'введен рост
        End If 'заполнена переменная роста пациента
    End Sub
End Module